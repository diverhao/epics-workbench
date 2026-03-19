const vscode = require("vscode");

const RUNTIME_VIEW_ID = "epicsRuntimeMonitors";
const ADD_RUNTIME_MONITOR_COMMAND = "vscode-epics.addRuntimeMonitor";
const REMOVE_RUNTIME_MONITOR_COMMAND = "vscode-epics.removeRuntimeMonitor";
const CLEAR_RUNTIME_MONITORS_COMMAND = "vscode-epics.clearRuntimeMonitors";
const RESTART_RUNTIME_CONTEXT_COMMAND = "vscode-epics.restartRuntimeContext";
const STOP_RUNTIME_CONTEXT_COMMAND = "vscode-epics.stopRuntimeContext";
const DEFAULT_PROTOCOL = "ca";
const DEFAULT_CHANNEL_CREATION_TIMEOUT_SECONDS = 3;
const DEFAULT_MONITOR_SUBSCRIBE_TIMEOUT_SECONDS = 3;
const DEFAULT_LOG_LEVEL = "ERROR";

function registerRuntimeMonitor(extensionContext) {
  const controller = new EpicsRuntimeMonitorController();
  const treeView = vscode.window.createTreeView(RUNTIME_VIEW_ID, {
    treeDataProvider: controller,
    showCollapseAll: false,
  });

  extensionContext.subscriptions.push(controller, treeView);
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      ADD_RUNTIME_MONITOR_COMMAND,
      async () => controller.addMonitorInteractive(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      REMOVE_RUNTIME_MONITOR_COMMAND,
      async (entry) => controller.removeMonitor(entry),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      CLEAR_RUNTIME_MONITORS_COMMAND,
      async () => controller.clearMonitors(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      RESTART_RUNTIME_CONTEXT_COMMAND,
      async () => controller.restartContext(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      STOP_RUNTIME_CONTEXT_COMMAND,
      async () => controller.stopContext(),
    ),
  );
}

class EpicsRuntimeMonitorController {
  constructor() {
    this._onDidChangeTreeData = new vscode.EventEmitter();
    this.onDidChangeTreeData = this._onDidChangeTreeData.event;
    this.contextNode = { type: "context" };
    this.contextStatus = "stopped";
    this.contextError = undefined;
    this.runtimeContext = undefined;
    this.runtimeLibrary = undefined;
    this.contextInitializationPromise = undefined;
    this.monitorEntries = [];
  }

  getTreeItem(element) {
    if (element?.type === "context") {
      return this.createContextTreeItem();
    }

    return this.createMonitorTreeItem(element);
  }

  getChildren(element) {
    if (!element) {
      return [this.contextNode, ...this.monitorEntries];
    }

    return [];
  }

  dispose() {
    this.stopContextInternal();
    this._onDidChangeTreeData.dispose();
  }

  async addMonitorInteractive() {
    const protocolItems = [
      {
        label: "Channel Access",
        description: "ca",
        protocol: "ca",
      },
      {
        label: "PV Access",
        description: "pva",
        protocol: "pva",
      },
    ];
    const preferredProtocol = vscode.workspace
      .getConfiguration("epicsWorkbench.runtime")
      .get("defaultProtocol", DEFAULT_PROTOCOL);
    protocolItems.sort((left, right) => {
      if (left.protocol === preferredProtocol && right.protocol !== preferredProtocol) {
        return -1;
      }
      if (right.protocol === preferredProtocol && left.protocol !== preferredProtocol) {
        return 1;
      }
      return 0;
    });

    const protocol = await vscode.window.showQuickPick(
      protocolItems,
      {
        placeHolder: "Select the EPICS protocol to monitor",
      },
    );
    if (!protocol) {
      return;
    }

    const pvName = await vscode.window.showInputBox({
      prompt: "EPICS PV name",
      value: getSuggestedPvName(),
      validateInput: (value) =>
        String(value || "").trim() ? undefined : "PV name is required",
      ignoreFocusOut: true,
    });
    if (pvName === undefined) {
      return;
    }

    let pvRequest = "";
    if (protocol.protocol === "pva") {
      const input = await vscode.window.showInputBox({
        prompt: "PV Access request",
        placeHolder: "Leave empty for the default request",
        ignoreFocusOut: true,
      });
      if (input === undefined) {
        return;
      }
      pvRequest = String(input || "").trim();
    }

    await this.addMonitor({
      pvName: String(pvName).trim(),
      protocol: protocol.protocol,
      pvRequest,
    });
  }

  async addMonitor(definition) {
    const key = createMonitorKey(definition);
    const existingEntry = this.monitorEntries.find((entry) => entry.key === key);
    if (existingEntry) {
      if (existingEntry.monitor) {
        vscode.window.showInformationMessage(
          `${existingEntry.pvName} is already being monitored.`,
        );
        return;
      }

      await this.connectEntry(existingEntry);
      return;
    }

    const entry = {
      type: "monitor",
      key,
      pvName: definition.pvName,
      protocol: definition.protocol,
      pvRequest: definition.pvRequest || "",
      channel: undefined,
      monitor: undefined,
      status: "pending",
      valueText: "",
      lastUpdated: undefined,
      lastError: undefined,
      serverAddress: undefined,
    };
    this.monitorEntries.push(entry);
    this.refresh();
    await this.connectEntry(entry);
  }

  async removeMonitor(entry) {
    const resolvedEntry = await this.resolveMonitorEntry(entry);
    if (!resolvedEntry) {
      return;
    }

    await this.disconnectEntry(resolvedEntry);
    this.monitorEntries = this.monitorEntries.filter(
      (candidate) => candidate.key !== resolvedEntry.key,
    );
    if (!this.monitorEntries.length) {
      this.stopContextInternal();
    }
    this.refresh();
  }

  async clearMonitors() {
    if (!this.monitorEntries.length) {
      return;
    }

    const confirmation = await vscode.window.showWarningMessage(
      "Remove all EPICS runtime monitors?",
      { modal: true },
      "Clear",
    );
    if (confirmation !== "Clear") {
      return;
    }

    this.stopContextInternal();
    this.monitorEntries = [];
    this.refresh();
  }

  async restartContext() {
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: "Restarting EPICS runtime context",
      },
      async () => {
        this.stopContextInternal();
        if (!this.monitorEntries.length) {
          this.contextStatus = "stopped";
          this.refresh();
          return;
        }

        for (const entry of this.monitorEntries) {
          entry.status = "pending";
          entry.lastError = undefined;
          entry.channel = undefined;
          entry.monitor = undefined;
        }
        this.refresh();

        for (const entry of this.monitorEntries) {
          await this.connectEntry(entry);
        }
      },
    );
  }

  async stopContext() {
    this.stopContextInternal();
    this.refresh();
  }

  async connectEntry(entry) {
    entry.status = "connecting";
    entry.lastError = undefined;
    this.refresh(entry);

    try {
      const runtimeContext = await this.ensureRuntimeContext();
      const channel = await runtimeContext.createChannel(
        entry.pvName,
        entry.protocol,
        this.getChannelCreationTimeoutSeconds(),
      );
      if (!channel) {
        throw new Error(
          `Failed to create ${entry.protocol.toUpperCase()} channel for ${entry.pvName}.`,
        );
      }

      channel.setDestroySoftCallback(() => {
        entry.status = "disconnected";
        entry.lastError = "Channel disconnected. Waiting for recovery.";
        this.refresh(entry);
      });
      channel.setDestroyHardCallback(() => {
        entry.status = "destroyed";
        entry.channel = undefined;
        entry.monitor = undefined;
        entry.lastError = "Channel was destroyed.";
        this.refresh(entry);
      });

      entry.channel = channel;
      entry.serverAddress = channel.getServerAddress();

      let monitor;
      if (entry.protocol === "pva") {
        monitor = await channel.createMonitorPva(
          this.getMonitorSubscribeTimeoutSeconds(),
          entry.pvRequest,
          (activeMonitor) => {
            this.handleMonitorUpdate(entry, activeMonitor);
          },
        );
      } else {
        monitor = await channel.createMonitor(
          this.getMonitorSubscribeTimeoutSeconds(),
          (activeMonitor) => {
            this.handleMonitorUpdate(entry, activeMonitor);
          },
        );
      }

      if (!monitor) {
        throw new Error(
          `Failed to subscribe to ${entry.protocol.toUpperCase()} monitor for ${entry.pvName}.`,
        );
      }

      entry.monitor = monitor;
      entry.status = "subscribed";
      this.updateEntryValue(entry, monitor);
      this.refresh(entry);
    } catch (error) {
      entry.status = "error";
      entry.channel = undefined;
      entry.monitor = undefined;
      entry.lastError = getErrorMessage(error);
      this.refresh(entry);
      vscode.window.showErrorMessage(
        `EPICS runtime monitor failed for ${entry.pvName}: ${entry.lastError}`,
      );
    }
  }

  async disconnectEntry(entry) {
    const monitor = entry.monitor;
    const channel = entry.channel;

    entry.monitor = undefined;
    entry.channel = undefined;
    entry.serverAddress = undefined;

    try {
      monitor?.destroyHard();
    } catch (error) {
      // Ignore monitor teardown errors while cleaning up the UI state.
    }

    try {
      await channel?.destroyHard();
    } catch (error) {
      // Ignore channel teardown errors while cleaning up the UI state.
    }

    entry.status = "stopped";
  }

  handleMonitorUpdate(entry, monitor) {
    entry.status = "subscribed";
    entry.lastError = undefined;
    this.updateEntryValue(entry, monitor);
    this.refresh(entry);
  }

  updateEntryValue(entry, monitor) {
    entry.lastUpdated = new Date();
    if (entry.protocol === "pva") {
      entry.valueText = formatRuntimeValue(monitor.getPvaData());
      return;
    }

    entry.valueText = formatRuntimeValue(monitor.getChannel().getDbrData()?.value);
  }

  async ensureRuntimeContext() {
    if (this.runtimeContext) {
      return this.runtimeContext;
    }

    if (this.contextInitializationPromise) {
      return this.contextInitializationPromise;
    }

    this.contextStatus = "connecting";
    this.contextError = undefined;
    this.refresh(this.contextNode);

    this.contextInitializationPromise = (async () => {
      const { Context } = this.requireRuntimeLibrary();
      const runtimeContext = new Context(
        this.getRuntimeEnvironment(),
        this.getRuntimeLogLevel(),
      );
      await runtimeContext.initialize();
      this.runtimeContext = runtimeContext;
      this.contextStatus = "connected";
      this.contextError = undefined;
      this.refresh(this.contextNode);
      return runtimeContext;
    })();

    try {
      return await this.contextInitializationPromise;
    } catch (error) {
      this.runtimeContext = undefined;
      this.contextStatus = "error";
      this.contextError = getErrorMessage(error);
      this.refresh(this.contextNode);
      throw error;
    } finally {
      this.contextInitializationPromise = undefined;
    }
  }

  stopContextInternal() {
    try {
      this.runtimeContext?.destroyHard();
    } catch (error) {
      // Ignore teardown errors when the extension is deactivating.
    }

    this.runtimeContext = undefined;
    this.contextInitializationPromise = undefined;
    this.contextStatus = "stopped";
    this.contextError = undefined;

    for (const entry of this.monitorEntries) {
      entry.channel = undefined;
      entry.monitor = undefined;
      if (entry.status !== "error") {
        entry.status = "stopped";
      }
    }
  }

  requireRuntimeLibrary() {
    if (!this.runtimeLibrary) {
      try {
        this.runtimeLibrary = require("epics-tca");
      } catch (error) {
        throw new Error(
          "The epics-tca dependency is unavailable. Run `npm install` in the extension workspace to use runtime monitoring.",
        );
      }
    }

    return this.runtimeLibrary;
  }

  refresh(element) {
    this._onDidChangeTreeData.fire(element);
  }

  createContextTreeItem() {
    const treeItem = new vscode.TreeItem("EPICS Runtime");
    treeItem.description = this.getContextDescription();
    treeItem.tooltip = this.getContextTooltip();
    treeItem.iconPath = new vscode.ThemeIcon(
      this.contextStatus === "connected"
        ? "plug"
        : this.contextStatus === "connecting"
          ? "sync~spin"
          : this.contextStatus === "error"
            ? "error"
            : "debug-disconnect",
    );
    treeItem.contextValue = "epicsRuntimeContext";
    treeItem.collapsibleState = vscode.TreeItemCollapsibleState.None;
    return treeItem;
  }

  createMonitorTreeItem(entry) {
    const treeItem = new vscode.TreeItem(entry.pvName);
    treeItem.description = entry.valueText || entry.status;
    treeItem.tooltip = this.getMonitorTooltip(entry);
    treeItem.iconPath = new vscode.ThemeIcon(getMonitorIconName(entry.status));
    treeItem.contextValue = "epicsRuntimeMonitor";
    treeItem.collapsibleState = vscode.TreeItemCollapsibleState.None;
    return treeItem;
  }

  getContextDescription() {
    if (this.contextStatus === "connected") {
      return `${this.monitorEntries.filter((entry) => entry.monitor).length} active`;
    }

    if (this.contextStatus === "connecting") {
      return "connecting";
    }

    if (this.contextStatus === "error") {
      return "error";
    }

    return "stopped";
  }

  getContextTooltip() {
    const configuration = vscode.workspace.getConfiguration("epicsWorkbench.runtime");
    const lines = [
      `Status: ${this.contextStatus}`,
      `Default protocol: ${configuration.get("defaultProtocol", DEFAULT_PROTOCOL)}`,
      `Channel timeout: ${String(this.getChannelCreationTimeoutSeconds() ?? "none")}s`,
      `Monitor timeout: ${String(this.getMonitorSubscribeTimeoutSeconds() ?? "none")}s`,
    ];
    if (this.contextError) {
      lines.push("", `Error: ${this.contextError}`);
    }
    return lines.join("\n");
  }

  getMonitorTooltip(entry) {
    const lines = [
      entry.pvName,
      `Protocol: ${entry.protocol.toUpperCase()}`,
      `Status: ${entry.status}`,
    ];

    if (entry.pvRequest) {
      lines.push(`PV Request: ${entry.pvRequest}`);
    }
    if (entry.serverAddress) {
      lines.push(`Server: ${entry.serverAddress}`);
    }
    if (entry.valueText) {
      lines.push(`Value: ${entry.valueText}`);
    }
    if (entry.lastUpdated) {
      lines.push(`Updated: ${entry.lastUpdated.toLocaleTimeString()}`);
    }
    if (entry.lastError) {
      lines.push(`Error: ${entry.lastError}`);
    }

    return lines.join("\n");
  }

  async resolveMonitorEntry(entry) {
    if (entry?.type === "monitor") {
      return entry;
    }

    if (!this.monitorEntries.length) {
      return undefined;
    }

    if (this.monitorEntries.length === 1) {
      return this.monitorEntries[0];
    }

    const selected = await vscode.window.showQuickPick(
      this.monitorEntries.map((candidate) => ({
        label: candidate.pvName,
        description: candidate.protocol.toUpperCase(),
        detail: candidate.valueText || candidate.status,
        entry: candidate,
      })),
      {
        placeHolder: "Select the EPICS monitor to remove",
      },
    );
    return selected?.entry;
  }

  getChannelCreationTimeoutSeconds() {
    return getPositiveNumberSetting(
      vscode.workspace
        .getConfiguration("epicsWorkbench.runtime")
        .get(
          "channelCreationTimeoutSeconds",
          DEFAULT_CHANNEL_CREATION_TIMEOUT_SECONDS,
        ),
    );
  }

  getMonitorSubscribeTimeoutSeconds() {
    return getPositiveNumberSetting(
      vscode.workspace
        .getConfiguration("epicsWorkbench.runtime")
        .get(
          "monitorSubscribeTimeoutSeconds",
          DEFAULT_MONITOR_SUBSCRIBE_TIMEOUT_SECONDS,
        ),
    );
  }

  getRuntimeLogLevel() {
    return vscode.workspace
      .getConfiguration("epicsWorkbench.runtime")
      .get("logLevel", DEFAULT_LOG_LEVEL);
  }

  getRuntimeEnvironment() {
    const raw = vscode.workspace
      .getConfiguration("epicsWorkbench.runtime")
      .get("environment", {});
    const environment = {};

    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return environment;
    }

    for (const [key, value] of Object.entries(raw)) {
      if (!key) {
        continue;
      }

      environment[key] = String(value);
    }

    return environment;
  }
}

function createMonitorKey({ pvName, protocol, pvRequest }) {
  return `${protocol}:${pvName}:${pvRequest || ""}`;
}

function getSuggestedPvName() {
  const editor = vscode.window.activeTextEditor;
  const selection = editor?.selection;
  if (!editor || !selection || selection.isEmpty) {
    return "";
  }

  return editor.document.getText(selection).trim();
}

function formatRuntimeValue(value) {
  if (value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return truncateText(value, 80);
  }

  try {
    return truncateText(JSON.stringify(value), 80);
  } catch (error) {
    return truncateText(String(value), 80);
  }
}

function truncateText(value, maxLength) {
  const text = String(value ?? "");
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength - 3)}...`;
}

function getPositiveNumberSetting(value) {
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue) || numericValue <= 0) {
    return undefined;
  }

  return numericValue;
}

function getErrorMessage(error) {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return String(error || "Unknown error");
}

function getMonitorIconName(status) {
  switch (status) {
    case "subscribed":
      return "radio-tower";

    case "connecting":
      return "sync~spin";

    case "error":
    case "destroyed":
      return "error";

    case "disconnected":
      return "warning";

    default:
      return "circle-large-outline";
  }
}

module.exports = {
  registerRuntimeMonitor,
};
