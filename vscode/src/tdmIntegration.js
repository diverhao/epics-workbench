const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const https = require("https");
const childProcess = require("child_process");
const net = require("net");
const vscode = require("vscode");
const { WebSocket, WebSocketServer } = require("ws");
const {
  EXTERNAL_RUNTIME_MODE,
  VENDORED_RUNTIME_MODE,
  getVendoredTdmLaunchInfo,
  getVendoredTdmRoot,
} = require("./tdmVendoredRuntime");

const OPEN_IN_TDM_COMMAND = "vscode-epics.openInTdm";
const TDM_CUSTOM_EDITOR_VIEW_TYPE = "epicsWorkbench.tdmEditor";
const TDM_HTTP_HOST = "127.0.0.1";
const TDM_WEBVIEW_HOST = "localhost";
const TDM_PROXY_IPC_PATH_PREFIX = "/_epics_tdm_ipc/";
const DEFAULT_READY_TIMEOUT_MS = 30000;
const READY_POLL_INTERVAL_MS = 500;
const LOCALHOST_HTTPS_AGENT = new https.Agent({
  rejectUnauthorized: false,
});
const SUPPORTED_TDM_EXTENSIONS = new Set([
  ".tdl",
  ".edl",
  ".bob",
  ".stp",
  ".plt",
]);

function registerTdmIntegration(extensionContext) {
  const controller = new EpicsTdmController(extensionContext);
  const provider = new EpicsTdmCustomEditorProvider(controller);

  extensionContext.subscriptions.push(controller, provider);
  extensionContext.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      TDM_CUSTOM_EDITOR_VIEW_TYPE,
      provider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
        },
        supportsMultipleEditorsPerDocument: true,
      },
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_TDM_COMMAND, async (resourceUri) => {
      await controller.openInTdm(resourceUri);
    }),
  );
}

class EpicsTdmCustomEditorProvider {
  constructor(controller) {
    this.controller = controller;
  }

  async resolveCustomTextEditor(document, webviewPanel) {
    await this.controller.resolveCustomEditor(document, webviewPanel);
  }

  dispose() {}
}

class EpicsTdmController {
  constructor(extensionContext) {
    this.extensionContext = extensionContext;
    this.outputChannel = vscode.window.createOutputChannel("EPICS TDM");
    this.runtime = undefined;
    this.runtimeStartPromise = undefined;
    this.runtimeStartConfigKey = undefined;
    this.openDisplayPanels = new Map();
  }

  dispose() {
    this.runtimeStartPromise = undefined;
    this.runtimeStartConfigKey = undefined;
    this.disposeRuntime();
    this.outputChannel.dispose();
  }

  async openInTdm(resourceUri) {
    const targetUri = resolveTdmTargetUri(resourceUri);
    if (!targetUri) {
      vscode.window.showErrorMessage(
        "Open a supported TDM display file first or invoke the command from the explorer.",
      );
      return;
    }
    if (!isSupportedTdmUri(targetUri)) {
      vscode.window.showErrorMessage(
        "Supported TDM files are .tdl, .edl, .bob, .stp, and .plt.",
      );
      return;
    }

    await vscode.commands.executeCommand(
      "vscode.openWith",
      targetUri,
      TDM_CUSTOM_EDITOR_VIEW_TYPE,
    );
  }

  async resolveCustomEditor(document, webviewPanel) {
    try {
      const display = await this.createDisplayForUri(document.uri);
      webviewPanel.title = path.basename(document.uri.fsPath);
      this.attachTdmPanel(
        webviewPanel,
        display.href,
        path.basename(document.uri.fsPath),
      );
    } catch (error) {
      const message = formatErrorMessage(error);
      this.outputChannel.appendLine(`[error] ${message}`);
      this.outputChannel.show(true);
      webviewPanel.webview.options = {
        enableScripts: false,
      };
      webviewPanel.webview.html = buildTdmErrorHtml(
        webviewPanel.webview,
        message,
      );
    }
  }

  async createDisplayForUri(uri) {
    if (!uri?.fsPath || !isSupportedTdmUri(uri)) {
      throw new Error(
        "TDM integration only supports local .tdl, .edl, .bob, .stp, and .plt files.",
      );
    }

    const runtime = await this.ensureRuntime();
    const profileOptions = loadTdmProfileOptions(runtime.profilesPath, runtime.profileName);
    const requestPayload = {
      tdlFileNames: [uri.fsPath],
      mode: profileOptions.mode,
      editable: profileOptions.editable,
      macros: [],
      replaceMacros: false,
      currentTdlFolder: path.dirname(uri.fsPath),
      windowId: "0",
    };
    const response = await postJsonHttps(
      runtime.tdmPort,
      "/command",
      {
        command: "create-display-window-agent",
        data: JSON.stringify(requestPayload),
      },
    );

    const displayWindowId = response?.data?.displayWindowId;
    if (!displayWindowId) {
      throw new Error("TDM did not return a display window id.");
    }

    return {
      displayWindowId,
      href: buildProxyDisplayHref(runtime.proxyPort, displayWindowId),
    };
  }

  async openDisplayHref(href, title = "TDM Display") {
    const displayWindowId = parseDisplayWindowIdFromHref(href);
    if (displayWindowId) {
      const existingPanel = this.openDisplayPanels.get(displayWindowId);
      if (existingPanel) {
        try {
          existingPanel.reveal(existingPanel.viewColumn || vscode.ViewColumn.Active);
          return existingPanel;
        } catch (error) {
          this.openDisplayPanels.delete(displayWindowId);
        }
      }
    }

    const panel = vscode.window.createWebviewPanel(
      "epicsWorkbench.tdmDisplay",
      title,
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );
    this.attachTdmPanel(panel, href, title);
    return panel;
  }

  attachTdmPanel(panel, href, title) {
    const currentDisplayWindowId = parseDisplayWindowIdFromHref(href);
    if (currentDisplayWindowId) {
      this.openDisplayPanels.set(currentDisplayWindowId, panel);
    }

    panel.webview.options = {
      enableScripts: true,
      portMapping: resolveTdmPortMappings(href),
    };
    panel.webview.html = buildTdmHostWebviewHtml(panel.webview, href, title);

    const messageDisposable = panel.webview.onDidReceiveMessage(async (message) => {
      if (message?.type === "open-window" && typeof message.href === "string") {
        const targetDisplayWindowId = parseDisplayWindowIdFromHref(message.href);
        this.outputChannel.appendLine(
          `[iframe:open-window] ${title}: ${message.href}`,
        );
        if (
          currentDisplayWindowId &&
          targetDisplayWindowId &&
          currentDisplayWindowId === targetDisplayWindowId
        ) {
          this.outputChannel.appendLine(
            `[iframe:open-window] ${title}: ignored self-open for display ${targetDisplayWindowId}`,
          );
          return;
        }
        await this.openDisplayHref(message.href);
        return;
      }

      if (message?.type === "tdm-log") {
        const level = typeof message.level === "string" ? message.level : "info";
        const text = typeof message.text === "string" ? message.text : JSON.stringify(message);
        this.outputChannel.appendLine(`[iframe:${level}] ${title}: ${text}`);
        if (level === "error") {
          this.outputChannel.show(true);
        }
        return;
      }

      if (message?.type === "tdm-websocket") {
        const eventType = typeof message.eventType === "string" ? message.eventType : "event";
        const url = typeof message.url === "string" ? message.url : "<unknown>";
        const extra = typeof message.extra === "string" && message.extra.length > 0
          ? ` ${message.extra}`
          : "";
        this.outputChannel.appendLine(`[iframe:ws] ${title}: ${eventType} ${url}${extra}`);
        if (eventType === "error" || eventType === "close") {
          this.outputChannel.show(true);
        }
      }
    });

    panel.onDidDispose(() => {
      if (currentDisplayWindowId && this.openDisplayPanels.get(currentDisplayWindowId) === panel) {
        this.openDisplayPanels.delete(currentDisplayWindowId);
      }
      messageDisposable.dispose();
    });
  }

  async ensureRuntime() {
    const config = resolveTdmLaunchConfiguration(this.extensionContext);
    const configKey = JSON.stringify(config);

    if (this.runtime && this.runtime.configKey === configKey && this.runtimeIsUsable()) {
      return this.runtime;
    }

    if (this.runtimeStartPromise && this.runtimeStartConfigKey === configKey) {
      return this.runtimeStartPromise;
    }

    if (this.runtime && this.runtime.configKey !== configKey) {
      this.disposeRuntime();
    }

    this.runtimeStartConfigKey = configKey;
    this.runtimeStartPromise = this.startRuntime(config, configKey);
    try {
      const runtime = await this.runtimeStartPromise;
      this.runtime = runtime;
      return runtime;
    } catch (error) {
      this.disposeRuntime();
      throw error;
    } finally {
      this.runtimeStartPromise = undefined;
      this.runtimeStartConfigKey = undefined;
    }
  }

  runtimeIsUsable() {
    return Boolean(
      this.runtime &&
        this.runtime.proxyServer &&
        this.runtime.proxyWsServer &&
        this.runtime.tdmProcess &&
        !this.runtime.tdmProcess.killed &&
        typeof this.runtime.tdmPort === "number" &&
        typeof this.runtime.proxyPort === "number",
    );
  }

  async startRuntime(config, configKey) {
    const launchInfo = resolveTdmLaunchInfo(config, this.extensionContext.extensionPath);
    const requestedTdmPort = parsePortNumber(config.httpServerPort);
    const tdmPort = requestedTdmPort || await findAvailablePort();
    const proxyPort = await findAvailablePort();

    this.outputChannel.appendLine(
      `[info] Starting TDM: ${launchInfo.command} ${launchInfo.args.join(" ")}`,
    );
    this.outputChannel.appendLine(`[info] TDM port ${tdmPort}, proxy port ${proxyPort}`);

    const tdmProcess = childProcess.spawn(launchInfo.command, launchInfo.args.concat([
      "--main-process-mode",
      "web",
      "--http-server-port",
      String(tdmPort),
      "--settings",
      config.profilesPath,
      "--profile",
      config.profileName,
    ]), {
      cwd: launchInfo.cwd,
      env: process.env,
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });

    const runtime = {
      configKey,
      config,
      tdmPort,
      proxyPort,
      profilesPath: config.profilesPath,
      profileName: config.profileName,
      tdmProcess,
      proxyServer: undefined,
      proxyWsServer: undefined,
    };

    const pipeOutput = (stream, prefix) => {
      if (!stream) {
        return;
      }
      stream.on("data", (chunk) => {
        const text = chunk.toString();
        for (const line of text.split(/\r?\n/)) {
          if (line.trim().length > 0) {
            this.outputChannel.appendLine(`[${prefix}] ${line}`);
          }
        }
      });
    };

    pipeOutput(tdmProcess.stdout, "stdout");
    pipeOutput(tdmProcess.stderr, "stderr");

    tdmProcess.on("exit", (code, signal) => {
      this.outputChannel.appendLine(
        `[info] TDM exited with code=${code ?? "null"} signal=${signal ?? "null"}`,
      );
      if (this.runtime?.tdmProcess === tdmProcess) {
        this.disposeRuntime();
      }
    });

    try {
      await Promise.race([
        waitForTdmReady(tdmPort),
        onceProcessExit(tdmProcess).then(({ code, signal }) => {
          throw new Error(
            `TDM exited before the web service became ready (code=${code ?? "null"}, signal=${signal ?? "null"}).`,
          );
        }),
      ]);

      const proxy = await startTdmProxyServer(tdmPort, proxyPort, this.outputChannel);
      runtime.proxyServer = proxy.server;
      runtime.proxyWsServer = proxy.wsServer;

      this.outputChannel.appendLine("[info] TDM runtime is ready.");
      return runtime;
    } catch (error) {
      try {
        runtime.proxyWsServer?.close();
      } catch {}
      try {
        runtime.proxyServer?.close();
      } catch {}
      try {
        if (!tdmProcess.killed) {
          tdmProcess.kill();
        }
      } catch {}
      throw error;
    }
  }

  disposeRuntime() {
    const runtime = this.runtime;
    this.runtime = undefined;

    if (!runtime) {
      return;
    }

    try {
      runtime.proxyWsServer?.close();
    } catch (error) {
      this.outputChannel.appendLine(`[warn] Failed to close TDM proxy ws server: ${formatErrorMessage(error)}`);
    }
    try {
      runtime.proxyServer?.close();
    } catch (error) {
      this.outputChannel.appendLine(`[warn] Failed to close TDM proxy server: ${formatErrorMessage(error)}`);
    }
    try {
      if (runtime.tdmProcess && !runtime.tdmProcess.killed) {
        runtime.tdmProcess.kill();
      }
    } catch (error) {
      this.outputChannel.appendLine(`[warn] Failed to stop TDM: ${formatErrorMessage(error)}`);
    }
  }
}

function resolveTdmTargetUri(resourceUri) {
  if (resourceUri instanceof vscode.Uri) {
    return resourceUri;
  }

  const activeDocument = vscode.window.activeTextEditor?.document;
  if (activeDocument?.uri) {
    return activeDocument.uri;
  }

  return undefined;
}

function isSupportedTdmUri(uri) {
  return uri?.scheme === "file" && SUPPORTED_TDM_EXTENSIONS.has(path.extname(uri.fsPath).toLowerCase());
}

function resolveTdmLaunchConfiguration(extensionContext) {
  const config = vscode.workspace.getConfiguration("epicsWorkbench.tdm");
  const configuredRuntimeMode = normalizeRuntimeMode(config.get("runtimeMode"));
  const rootPathSetting = normalizeOptionalPath(config.get("rootPath"));
  const executablePathSetting = normalizeOptionalPath(config.get("executablePath"));
  const profilesPathSetting = normalizeOptionalPath(config.get("profilesPath"));
  const configuredProfileName = normalizeOptionalString(config.get("profile"));
  const httpServerPort = config.get("httpServerPort");
  const vendoredRoot = getVendoredTdmRoot(extensionContext.extensionPath);
  const runtimeMode =
    configuredRuntimeMode || (vendoredRoot ? VENDORED_RUNTIME_MODE : EXTERNAL_RUNTIME_MODE);

  if (runtimeMode === VENDORED_RUNTIME_MODE) {
    if (!vendoredRoot) {
      throw buildSettingsError(
        "The vendored TDM runtime is missing. Run `npm run sync:tdm-vendor` in the vscode project or switch `epicsWorkbench.tdm.runtimeMode` to `externalBinary`.",
      );
    }

    const profilesPath = profilesPathSetting || path.join(os.homedir(), ".tdm", "profiles.json");
    if (!fs.existsSync(profilesPath)) {
      throw buildSettingsError(
        `Cannot find TDM profiles at ${profilesPath}. Set \`epicsWorkbench.tdm.profilesPath\` to a valid file.`,
      );
    }

    const profileName = configuredProfileName || pickDefaultTdmProfileName(profilesPath);
    if (!profileName) {
      throw buildSettingsError(
        `Cannot determine a TDM profile from ${profilesPath}. Set \`epicsWorkbench.tdm.profile\`.`,
      );
    }

    return {
      runtimeMode,
      rootPath: vendoredRoot,
      executablePath: undefined,
      profilesPath,
      profileName,
      httpServerPort,
    };
  }

  const rootPath =
    rootPathSetting ||
    deriveTdmRootFromExecutable(executablePathSetting) ||
    detectDefaultTdmRoot(extensionContext);
  if (!rootPath) {
    throw buildSettingsError(
      "Cannot locate the TDM checkout or installation. Set `epicsWorkbench.tdm.rootPath` or `epicsWorkbench.tdm.executablePath`.",
    );
  }

  const profilesPath = profilesPathSetting || path.join(os.homedir(), ".tdm", "profiles.json");
  if (!fs.existsSync(profilesPath)) {
    throw buildSettingsError(
      `Cannot find TDM profiles at ${profilesPath}. Set \`epicsWorkbench.tdm.profilesPath\` to a valid file.`,
    );
  }

  const profileName = configuredProfileName || pickDefaultTdmProfileName(profilesPath);
  if (!profileName) {
    throw buildSettingsError(
      `Cannot determine a TDM profile from ${profilesPath}. Set \`epicsWorkbench.tdm.profile\`.`,
    );
  }

  return {
    runtimeMode,
    rootPath,
    executablePath: executablePathSetting,
    profilesPath,
    profileName,
    httpServerPort,
  };
}

function resolveTdmLaunchInfo(config, extensionPath) {
  if (config.runtimeMode === VENDORED_RUNTIME_MODE) {
    const launchInfo = getVendoredTdmLaunchInfo(extensionPath);
    if (!launchInfo) {
      throw buildSettingsError(
        "The vendored TDM runtime could not be launched because its copied runtime files are missing.",
      );
    }
    return launchInfo;
  }

  const executablePath = resolveTdmExecutablePath(config);
  if (!executablePath) {
    throw buildSettingsError(
      "Cannot locate a runnable TDM executable. Set `epicsWorkbench.tdm.executablePath` or point `epicsWorkbench.tdm.rootPath` at a built TDM checkout.",
    );
  }

  const command = executablePath;
  const cwd = config.rootPath;
  const args = [];

  if (isElectronBinary(command) && fs.existsSync(path.join(config.rootPath, "package.json"))) {
    args.push(config.rootPath);
  }

  return {
    command,
    args,
    cwd,
  };
}

function resolveTdmExecutablePath(config) {
  const explicit = resolveMacAppExecutable(config.executablePath);
  if (explicit && fs.existsSync(explicit)) {
    return explicit;
  }

  const rootPath = config.rootPath;
  const candidates = [
    path.join(rootPath, "out", "mac", "TDM.app"),
    path.join(rootPath, "out", "mac", "TDM.app", "Contents", "MacOS", "TDM"),
    path.join(rootPath, "out", "linux-unpacked", "tdm"),
    path.join(rootPath, "out", "linux-arm64-unpacked", "tdm"),
    path.join(rootPath, "out", "win-unpacked", "TDM.exe"),
    path.join(rootPath, "out", "win-arm64-unpacked", "TDM.exe"),
    path.join(rootPath, "node_modules", "electron", "dist", "Electron.app"),
    path.join(rootPath, "node_modules", "electron", "dist", "Electron.app", "Contents", "MacOS", "Electron"),
    path.join(rootPath, "node_modules", "electron", "dist", "electron"),
    path.join(rootPath, "node_modules", "electron", "dist", "electron.exe"),
  ];

  for (const candidate of candidates) {
    const resolvedCandidate = resolveMacAppExecutable(candidate);
    if (resolvedCandidate && fs.existsSync(resolvedCandidate)) {
      return resolvedCandidate;
    }
  }

  return undefined;
}

function deriveTdmRootFromExecutable(executablePath) {
  if (!executablePath) {
    return undefined;
  }

  const normalizedPath = executablePath.replaceAll("\\", "/");
  const knownSuffixes = [
    "/out/mac/TDM.app/Contents/MacOS/TDM",
    "/out/linux-unpacked/tdm",
    "/out/linux-arm64-unpacked/tdm",
    "/out/win-unpacked/TDM.exe",
    "/out/win-arm64-unpacked/TDM.exe",
    "/node_modules/electron/dist/Electron.app/Contents/MacOS/Electron",
    "/node_modules/electron/dist/electron",
    "/node_modules/electron/dist/electron.exe",
  ];

  for (const suffix of knownSuffixes) {
    if (normalizedPath.endsWith(suffix)) {
      return executablePath.slice(0, executablePath.length - suffix.length);
    }
  }

  return path.dirname(executablePath);
}

function isElectronBinary(executablePath) {
  const baseName = path.basename(executablePath).toLowerCase();
  return baseName === "electron" || baseName === "electron.exe";
}

function resolveMacAppExecutable(inputPath) {
  if (!inputPath) {
    return undefined;
  }
  if (!inputPath.endsWith(".app")) {
    return inputPath;
  }

  const appName = path.basename(inputPath, ".app");
  return path.join(inputPath, "Contents", "MacOS", appName);
}

function detectDefaultTdmRoot(extensionContext) {
  const environmentCandidate = normalizeOptionalPath(process.env.TDM_ROOT);
  if (environmentCandidate && fs.existsSync(environmentCandidate)) {
    return environmentCandidate;
  }

  const workspaceCandidate = vscode.workspace.workspaceFolders?.[0]?.uri?.fsPath
    ? path.resolve(vscode.workspace.workspaceFolders[0].uri.fsPath, "../tdm")
    : undefined;
  if (workspaceCandidate && fs.existsSync(workspaceCandidate)) {
    return workspaceCandidate;
  }

  const extensionSiblingCandidate = path.resolve(extensionContext.extensionPath, "../../tdm");
  if (fs.existsSync(extensionSiblingCandidate)) {
    return extensionSiblingCandidate;
  }

  return undefined;
}

function pickDefaultTdmProfileName(profilesPath) {
  const profiles = readJsonFile(profilesPath);
  if (!profiles || typeof profiles !== "object") {
    return undefined;
  }

  for (const profileName of Object.keys(profiles)) {
    if (profileName !== "For All Profiles") {
      return profileName;
    }
  }

  return undefined;
}

function loadTdmProfileOptions(profilesPath, profileName) {
  const profiles = readJsonFile(profilesPath);
  const profile = profiles?.[profileName];
  const env = profile?.["EPICS Custom Environment"] || {};
  const manuallyOpenedMode = env?.["Manually Opened TDL Mode"]?.value;
  const manuallyOpenedEditable = env?.["Manually Opened TDL Editable"]?.value;

  return {
    mode: manuallyOpenedMode === "editing" ? "editing" : "operating",
    editable:
      manuallyOpenedMode === "editing" ||
      String(manuallyOpenedEditable || "").trim().toUpperCase() === "YES",
  };
}

function readJsonFile(filePath) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (error) {
    throw new Error(`Failed to read JSON from ${filePath}: ${formatErrorMessage(error)}`);
  }
}

function parsePortNumber(value) {
  const numericValue =
    typeof value === "string"
      ? Number(value)
      : value;

  if (typeof numericValue === "number" && Number.isInteger(numericValue) && numericValue > 0 && numericValue < 65536) {
    return numericValue;
  }

  return undefined;
}

function normalizeOptionalString(value) {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function normalizeOptionalPath(value) {
  const normalized = normalizeOptionalString(value);
  return normalized ? path.resolve(expandHomeDirectory(normalized)) : undefined;
}

function normalizeRuntimeMode(value) {
  return value === VENDORED_RUNTIME_MODE || value === EXTERNAL_RUNTIME_MODE
    ? value
    : undefined;
}

function expandHomeDirectory(inputPath) {
  if (!inputPath.startsWith("~/")) {
    return inputPath;
  }
  return path.join(os.homedir(), inputPath.slice(2));
}

function buildProxyDisplayHref(proxyPort, displayWindowId) {
  return `http://${TDM_WEBVIEW_HOST}:${proxyPort}/DisplayWindow.html?displayWindowId=${encodeURIComponent(displayWindowId)}`;
}

function parseDisplayWindowIdFromHref(href) {
  try {
    const parsedHref = new URL(href);
    return normalizeOptionalString(parsedHref.searchParams.get("displayWindowId"));
  } catch (error) {
    return undefined;
  }
}

function resolveTdmPortMappings(href) {
  try {
    const parsedHref = new URL(href);
    if (parsedHref.hostname !== TDM_WEBVIEW_HOST) {
      return undefined;
    }

    const port = Number(parsedHref.port);
    if (!Number.isInteger(port) || port <= 0) {
      return undefined;
    }

    return [
      {
        webviewPort: port,
        extensionHostPort: port,
      },
    ];
  } catch (error) {
    return undefined;
  }
}

function waitForTdmReady(port, timeoutMs = DEFAULT_READY_TIMEOUT_MS) {
  const deadline = Date.now() + timeoutMs;

  const tryOnce = async () => {
    await postJsonHttps(port, "/command", { command: "get-ipc-server-port" });
  };

  return new Promise((resolve, reject) => {
    const attempt = async () => {
      try {
        await tryOnce();
        resolve();
      } catch (error) {
        if (Date.now() >= deadline) {
          reject(new Error(`Timed out waiting for TDM web mode on port ${port}: ${formatErrorMessage(error)}`));
          return;
        }

        setTimeout(() => {
          void attempt();
        }, READY_POLL_INTERVAL_MS);
      }
    };

    void attempt();
  });
}

function onceProcessExit(childProcessHandle) {
  return new Promise((resolve) => {
    childProcessHandle.once("exit", (code, signal) => {
      resolve({ code, signal });
    });
  });
}

async function findAvailablePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen(0, TDM_HTTP_HOST, () => {
      const address = server.address();
      const port = typeof address === "object" && address ? address.port : undefined;
      server.close((closeError) => {
        if (closeError) {
          reject(closeError);
          return;
        }
        resolve(port);
      });
    });
  });
}

function postJsonHttps(port, requestPath, body) {
  return requestJson({
    protocolModule: https,
    options: {
      hostname: TDM_HTTP_HOST,
      port,
      path: requestPath,
      method: "POST",
      agent: LOCALHOST_HTTPS_AGENT,
      headers: {
        "content-type": "application/json",
      },
    },
    body: JSON.stringify(body),
  });
}

function requestJson({ protocolModule, options, body }) {
  return new Promise((resolve, reject) => {
    const request = protocolModule.request(options, (response) => {
      const chunks = [];
      response.on("data", (chunk) => {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      });
      response.on("end", () => {
        const responseText = Buffer.concat(chunks).toString("utf8");
        if (response.statusCode && response.statusCode >= 400) {
          reject(new Error(`HTTP ${response.statusCode}: ${responseText || "<empty response>"}`));
          return;
        }
        try {
          resolve(responseText.length > 0 ? JSON.parse(responseText) : {});
        } catch (error) {
          reject(new Error(`Invalid JSON response: ${formatErrorMessage(error)}`));
        }
      });
    });

    request.on("error", reject);
    request.setTimeout(DEFAULT_READY_TIMEOUT_MS, () => {
      request.destroy(new Error("Request timed out."));
    });
    if (body) {
      request.write(body);
    }
    request.end();
  });
}

async function startTdmProxyServer(tdmPort, proxyPort, outputChannel) {
  const server = http.createServer((request, response) => {
    void proxyHttpRequest(tdmPort, request, response, outputChannel);
  });
  const wsServer = new WebSocketServer({ noServer: true });

  server.on("upgrade", (request, socket, head) => {
    wsServer.handleUpgrade(request, socket, head, (clientSocket) => {
      proxyWebSocketUpgrade(tdmPort, request, clientSocket, outputChannel);
    });
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(proxyPort, TDM_HTTP_HOST, resolve);
  });

  return {
    server,
    wsServer,
  };
}

async function proxyHttpRequest(tdmPort, clientRequest, clientResponse, outputChannel) {
  const upstreamRequest = https.request(
    {
      hostname: TDM_HTTP_HOST,
      port: tdmPort,
      path: clientRequest.url,
      method: clientRequest.method,
      headers: stripHopByHopHeaders(clientRequest.headers),
      agent: LOCALHOST_HTTPS_AGENT,
    },
    (upstreamResponse) => {
      const contentType = String(upstreamResponse.headers["content-type"] || "");
      if (shouldRewriteProxyResponse(contentType)) {
        const chunks = [];
        upstreamResponse.on("data", (chunk) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });
        upstreamResponse.on("end", () => {
          const body = Buffer.concat(chunks).toString("utf8");
          const rewrittenBody = rewriteTdmProxyBody(body, contentType);
          const headers = { ...upstreamResponse.headers };
          delete headers["content-length"];
          delete headers["content-security-policy"];
          clientResponse.writeHead(upstreamResponse.statusCode || 200, headers);
          clientResponse.end(rewrittenBody, "utf8");
        });
        return;
      }

      clientResponse.writeHead(upstreamResponse.statusCode || 200, upstreamResponse.headers);
      upstreamResponse.pipe(clientResponse);
    },
  );

  upstreamRequest.on("error", (error) => {
    outputChannel.appendLine(`[proxy] HTTP error: ${formatErrorMessage(error)}`);
    if (!clientResponse.headersSent) {
      clientResponse.writeHead(502, { "content-type": "text/plain; charset=utf-8" });
    }
    clientResponse.end(`TDM proxy error: ${formatErrorMessage(error)}`);
  });

  clientRequest.pipe(upstreamRequest);
}

function proxyWebSocketUpgrade(tdmPort, clientRequest, clientSocket, outputChannel) {
  let upstreamTarget;
  try {
    upstreamTarget = resolveProxyWebSocketTarget(tdmPort, clientRequest.url);
  } catch (error) {
    outputChannel.appendLine(`[proxy] websocket target error: ${formatErrorMessage(error)}`);
    clientSocket.close(1011, "Invalid TDM websocket target.");
    return;
  }
  outputChannel.appendLine(
    `[proxy] websocket ${clientRequest.url || "/"} -> wss://${TDM_HTTP_HOST}:${upstreamTarget.port}${upstreamTarget.path}`,
  );

  const protocolsHeader = clientRequest.headers["sec-websocket-protocol"];
  const protocols = typeof protocolsHeader === "string"
    ? protocolsHeader.split(",").map((value) => value.trim()).filter(Boolean)
    : undefined;
  const upstreamSocket = new WebSocket(
    `wss://${TDM_HTTP_HOST}:${upstreamTarget.port}${upstreamTarget.path}`,
    protocols,
    {
      rejectUnauthorized: false,
      headers: stripHopByHopHeaders(clientRequest.headers),
    },
  );
  const pendingClientMessages = [];

  const closeBoth = () => {
    if (clientSocket.readyState === WebSocket.OPEN || clientSocket.readyState === WebSocket.CONNECTING) {
      clientSocket.close();
    }
    if (upstreamSocket.readyState === WebSocket.OPEN || upstreamSocket.readyState === WebSocket.CONNECTING) {
      upstreamSocket.close();
    }
  };

  clientSocket.on("message", (data, isBinary) => {
    if (upstreamSocket.readyState === WebSocket.OPEN) {
      upstreamSocket.send(data, { binary: isBinary });
      return;
    }

    if (upstreamSocket.readyState === WebSocket.CONNECTING) {
      pendingClientMessages.push({ data, isBinary });
    }
  });
  upstreamSocket.on("message", (data, isBinary) => {
    if (clientSocket.readyState === WebSocket.OPEN) {
      clientSocket.send(data, { binary: isBinary });
    }
  });
  upstreamSocket.on("open", () => {
    outputChannel.appendLine(
      `[proxy] upstream websocket open ${clientRequest.url || "/"} pending=${pendingClientMessages.length}`,
    );
    while (pendingClientMessages.length > 0 && upstreamSocket.readyState === WebSocket.OPEN) {
      const message = pendingClientMessages.shift();
      upstreamSocket.send(message.data, { binary: message.isBinary });
    }
  });

  clientSocket.on("close", () => {
    outputChannel.appendLine(`[proxy] client websocket close ${clientRequest.url || "/"}`);
    closeBoth();
  });
  upstreamSocket.on("close", (code, reason) => {
    outputChannel.appendLine(
      `[proxy] upstream websocket close ${clientRequest.url || "/"} code=${code} reason=${String(reason || "")}`,
    );
    closeBoth();
  });

  clientSocket.on("error", (error) => {
    outputChannel.appendLine(`[proxy] client websocket error: ${formatErrorMessage(error)}`);
    closeBoth();
  });
  upstreamSocket.on("error", (error) => {
    outputChannel.appendLine(`[proxy] upstream websocket error: ${formatErrorMessage(error)}`);
    closeBoth();
  });
}

function resolveProxyWebSocketTarget(tdmPort, requestUrl) {
  const rawUrl = typeof requestUrl === "string" && requestUrl.length > 0 ? requestUrl : "/";
  const parsedUrl = new URL(rawUrl, `http://${TDM_HTTP_HOST}`);

  if (parsedUrl.pathname.startsWith(TDM_PROXY_IPC_PATH_PREFIX)) {
    const encodedPort = parsedUrl.pathname.slice(TDM_PROXY_IPC_PATH_PREFIX.length);
    const ipcPort = parsePortNumber(decodeURIComponent(encodedPort));
    if (!ipcPort) {
      throw new Error(`Invalid TDM IPC proxy port in path ${rawUrl}.`);
    }

    return {
      port: Number(ipcPort),
      path: "/",
    };
  }

  return {
    port: tdmPort,
    path: `${parsedUrl.pathname}${parsedUrl.search}`,
  };
}

function stripHopByHopHeaders(headers) {
  const nextHeaders = { ...headers };
  delete nextHeaders.connection;
  delete nextHeaders["proxy-connection"];
  delete nextHeaders["keep-alive"];
  delete nextHeaders["transfer-encoding"];
  delete nextHeaders.upgrade;
  delete nextHeaders.host;
  return nextHeaders;
}

function shouldRewriteProxyResponse(contentType) {
  return (
    contentType.includes("text/html") ||
    contentType.includes("javascript") ||
    contentType.includes("ecmascript")
  );
}

function rewriteTdmProxyBody(body, contentType) {
  let rewritten = body
    .replaceAll(
      "https://${window.location.host}/",
      "${window.location.protocol}//${window.location.host}/",
    )
    .replaceAll(
      "wss://${host}:${this.getIpcServerPort()}",
      "${window.__EPICS_WORKBENCH_TDM_IPC_URL__(this.getIpcServerPort())}",
    )
    .replaceAll(
      "wss://127.0.0.1:${this.getIpcServerPort()}",
      "${window.__EPICS_WORKBENCH_TDM_IPC_URL__(this.getIpcServerPort())}",
    )
    .replaceAll(
      "if (userAgent.indexOf(\\' electron/\\') > -1) {",
      "if (window.__EPICS_WORKBENCH_FORCE_WEB__ !== true && userAgent.indexOf(\\' electron/\\') > -1) {",
    )
    .replaceAll(
      "if (userAgent.indexOf(' electron/') > -1) {",
      "if (window.__EPICS_WORKBENCH_FORCE_WEB__ !== true && userAgent.indexOf(' electron/') > -1) {",
    )
    .replaceAll("window.open(", "window.__EPICS_WORKBENCH_TDM_OPEN__(");

  if (contentType.includes("text/html")) {
    rewritten = injectTdmProxyHelper(rewritten);
  }

  return rewritten;
}

function injectTdmProxyHelper(html) {
  const helperScript = `
<script>
window.__EPICS_WORKBENCH_TDM_POST__ = function(type, payload) {
  if (window.parent && window.parent !== window) {
    window.parent.postMessage(
      Object.assign(
        {
          source: "epics-workbench-tdm",
          type: type
        },
        payload || {}
      ),
      "*"
    );
  }
};
window.__EPICS_WORKBENCH_TDM_LOG__ = function(level, args) {
  try {
    const text = (args || []).map(function(value) {
      if (typeof value === "string") {
        return value;
      }
      try {
        return JSON.stringify(value);
      } catch (error) {
        return String(value);
      }
    }).join(" ").slice(0, 4000);
    window.__EPICS_WORKBENCH_TDM_POST__("tdm-log", {
      level: level,
      text: text
    });
  } catch (error) {}
};
window.__EPICS_WORKBENCH_FORCE_WEB__ = true;
window.__EPICS_WORKBENCH_TDM_IPC_URL__ = function(ipcPort) {
  const protocol = window.location.protocol === "https:" ? "wss" : "ws";
  return protocol + "://" + window.location.host + "${TDM_PROXY_IPC_PATH_PREFIX}" + encodeURIComponent(String(ipcPort));
};
["warn", "error"].forEach(function(level) {
  const original = console[level];
  console[level] = function() {
    const args = Array.prototype.slice.call(arguments);
    window.__EPICS_WORKBENCH_TDM_LOG__(level, args);
    return original.apply(this, args);
  };
});
window.addEventListener("error", function(event) {
  window.__EPICS_WORKBENCH_TDM_POST__("tdm-log", {
    level: "error",
    text: "window.error " + String(event.message || "") + " at " + String(event.filename || "") + ":" + String(event.lineno || 0) + ":" + String(event.colno || 0)
  });
});
window.addEventListener("unhandledrejection", function(event) {
  const reason = event && event.reason;
  window.__EPICS_WORKBENCH_TDM_POST__("tdm-log", {
    level: "error",
    text: "unhandledrejection " + (typeof reason === "string" ? reason : String(reason))
  });
});
(function() {
  const NativeWebSocket = window.WebSocket;
  if (typeof NativeWebSocket !== "function") {
    return;
  }

  function wrapSocket(socket, url) {
    socket.addEventListener("open", function() {
      window.__EPICS_WORKBENCH_TDM_POST__("tdm-websocket", {
        eventType: "open",
        url: String(url)
      });
    });
    socket.addEventListener("error", function() {
      window.__EPICS_WORKBENCH_TDM_POST__("tdm-websocket", {
        eventType: "error",
        url: String(url)
      });
    });
    socket.addEventListener("close", function(event) {
      window.__EPICS_WORKBENCH_TDM_POST__("tdm-websocket", {
        eventType: "close",
        url: String(url),
        extra: "code=" + String(event.code) + " reason=" + String(event.reason || "") + " clean=" + String(event.wasClean)
      });
    });
    return socket;
  }

  function InstrumentedWebSocket(url, protocols) {
    const socket = protocols === undefined
      ? new NativeWebSocket(url)
      : new NativeWebSocket(url, protocols);
    return wrapSocket(socket, url);
  }

  InstrumentedWebSocket.prototype = NativeWebSocket.prototype;
  InstrumentedWebSocket.CONNECTING = NativeWebSocket.CONNECTING;
  InstrumentedWebSocket.OPEN = NativeWebSocket.OPEN;
  InstrumentedWebSocket.CLOSING = NativeWebSocket.CLOSING;
  InstrumentedWebSocket.CLOSED = NativeWebSocket.CLOSED;
  window.WebSocket = InstrumentedWebSocket;
})();
window.__EPICS_WORKBENCH_TDM_OPEN__ = function(url, target, features) {
  if (window.parent && window.parent !== window) {
    window.__EPICS_WORKBENCH_TDM_POST__("open-window", {
      href: url
    });
    return null;
  }
  return globalThis.open.call(window, url, target, features);
};
</script>`;

  if (html.includes("</head>")) {
    return html.replace("</head>", `${helperScript}\n</head>`);
  }

  return `${helperScript}\n${html}`;
}

function buildTdmHostWebviewHtml(webview, href, title) {
  const nonce = createNonce();
  const escapedHref = escapeHtml(href);
  const escapedTitle = escapeHtml(title);

  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}'; frame-src http://${TDM_HTTP_HOST}:* https://${TDM_HTTP_HOST}:* http://${TDM_WEBVIEW_HOST}:* https://${TDM_WEBVIEW_HOST}:*;"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${escapedTitle}</title>
    <style>
      html, body {
        margin: 0;
        padding: 0;
        width: 100%;
        height: 100%;
        overflow: hidden;
        background: var(--vscode-editor-background);
        color: var(--vscode-editor-foreground);
        font-family: var(--vscode-font-family);
      }
      .shell {
        position: relative;
        width: 100%;
        height: 100%;
      }
      .frame {
        width: 100%;
        height: 100%;
        border: 0;
        background: #fff;
      }
      .loading {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        font-size: 12px;
        color: var(--vscode-descriptionForeground);
        background:
          linear-gradient(180deg, rgba(0, 0, 0, 0.08), transparent 40%),
          var(--vscode-editor-background);
      }
      .loading[data-hidden="true"] {
        display: none;
      }
    </style>
  </head>
  <body>
    <div class="shell">
      <div class="loading" id="loading">Loading TDM display...</div>
      <iframe
        id="tdm-frame"
        class="frame"
        src="${escapedHref}"
        sandbox="allow-same-origin allow-scripts allow-forms allow-downloads"
      ></iframe>
    </div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const frame = document.getElementById("tdm-frame");
      const loading = document.getElementById("loading");

      frame.addEventListener("load", () => {
        loading.dataset.hidden = "true";
      });

      window.addEventListener("message", (event) => {
        const data = event.data;
        if (!data || data.source !== "epics-workbench-tdm") {
          return;
        }
        vscode.postMessage(data);
      });
    </script>
  </body>
</html>`;
}

function buildTdmErrorHtml(webview, message) {
  const escapedMessage = escapeHtml(message);

  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline';"
    />
    <style>
      html, body {
        margin: 0;
        padding: 0;
        width: 100%;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-editor-foreground);
        font-family: var(--vscode-font-family);
      }
      .container {
        box-sizing: border-box;
        max-width: 900px;
        margin: 0 auto;
        padding: 32px 28px;
      }
      .title {
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 16px;
      }
      .message {
        line-height: 1.6;
        white-space: pre-wrap;
        color: var(--vscode-descriptionForeground);
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="title">TDM integration could not start</div>
      <div class="message">${escapedMessage}</div>
    </div>
  </body>
</html>`;
}

function buildSettingsError(message) {
  return new Error(`${message}\nOpen Settings and search for "epicsWorkbench.tdm".`);
}

function formatErrorMessage(error) {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

function createNonce() {
  return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

module.exports = {
  registerTdmIntegration,
  OPEN_IN_TDM_COMMAND,
  TDM_CUSTOM_EDITOR_VIEW_TYPE,
};
