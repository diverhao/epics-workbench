const fs = require("fs");
const path = require("path");
const vscode = require("vscode");
const { formatMonitorText } = require("./formatters");

const RUNTIME_VIEW_ID = "epicsRuntimeMonitors";
const ADD_RUNTIME_MONITOR_COMMAND = "vscode-epics.addRuntimeMonitor";
const REMOVE_RUNTIME_MONITOR_COMMAND = "vscode-epics.removeRuntimeMonitor";
const CLEAR_RUNTIME_MONITORS_COMMAND = "vscode-epics.clearRuntimeMonitors";
const RESTART_RUNTIME_CONTEXT_COMMAND = "vscode-epics.restartRuntimeContext";
const STOP_RUNTIME_CONTEXT_COMMAND = "vscode-epics.stopRuntimeContext";
const START_ACTIVE_FILE_RUNTIME_COMMAND = "vscode-epics.startActiveFileRuntimeContext";
const STOP_ACTIVE_FILE_RUNTIME_COMMAND = "vscode-epics.stopActiveFileRuntimeContext";
const PUT_RUNTIME_VALUE_COMMAND = "vscode-epics.putRuntimeValue";
const OPEN_PROJECT_RUNTIME_CONFIGURATION_COMMAND =
  "vscode-epics.openProjectRuntimeConfiguration";
const DEFAULT_PROTOCOL = "ca";
const DEFAULT_CHANNEL_CREATION_TIMEOUT_SECONDS = 0;
const DEFAULT_MONITOR_SUBSCRIBE_TIMEOUT_SECONDS = 0;
const DEFAULT_LOG_LEVEL = "ERROR";
const DATABASE_RUNTIME_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const LINE_RUNTIME_EXTENSIONS = new Set([".monitor", ".txt"]);
const STATUS_BAR_PRIORITY = 110;
const MOUSE_DOUBLE_CLICK_INTERVAL_MS = 400;
const PROJECT_RUNTIME_CONFIG_FILE_NAME = ".epics-workbench-config.json";
const PROJECT_RUNTIME_CONFIGURATION_VIEW_TYPE =
  "epicsWorkbench.projectRuntimeConfiguration";
const PROJECT_RUNTIME_CONFIGURATION_PROTOCOL_VALUES = ["ca", "pva"];
const PROJECT_RUNTIME_CONFIGURATION_AUTO_ADDR_LIST_VALUES = ["Yes", "No"];
const PVA_HAS_DATA_WITHOUT_VALUE_TEXT = "Has data, but no value";
const RUNTIME_VALUE_DISPLAY_MAX_LENGTH = 120;
const RUNTIME_HOVER_VALUE_MAX_LENGTH = 240;
const DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH = 18;
const CONTEXT_INITIALIZATION_CANCELLED_MESSAGE =
  "EPICS runtime context initialization was cancelled.";

function registerRuntimeMonitor(extensionContext, databaseHelpers = {}) {
  const controller = new EpicsRuntimeMonitorController(databaseHelpers);
  const treeView = vscode.window.createTreeView(RUNTIME_VIEW_ID, {
    treeDataProvider: controller,
    showCollapseAll: false,
  });
  const diagnostics = vscode.languages.createDiagnosticCollection("vscode-epics-monitor");
  const statusBarItem = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Left,
    STATUS_BAR_PRIORITY,
  );
  const hoverDecorationType = vscode.window.createTextEditorDecorationType({
    after: {
      color: new vscode.ThemeColor("descriptionForeground"),
    },
  });
  const databaseTocValueDecorationType = vscode.window.createTextEditorDecorationType({
    after: {
      color: new vscode.ThemeColor("descriptionForeground"),
    },
  });

  controller.attachStatusBar(statusBarItem);
  controller.attachDiagnosticsCollection(diagnostics);
  controller.attachHoverDecorationType(hoverDecorationType);
  controller.attachDatabaseTocValueDecorationType(databaseTocValueDecorationType);

  extensionContext.subscriptions.push(
    controller,
    treeView,
    diagnostics,
    statusBarItem,
    hoverDecorationType,
    databaseTocValueDecorationType,
  );
  extensionContext.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: "monitor" },
      new EpicsMonitorFormattingProvider(),
    ),
  );
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
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      START_ACTIVE_FILE_RUNTIME_COMMAND,
      async () => controller.startActiveFileRuntimeContext(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      STOP_ACTIVE_FILE_RUNTIME_COMMAND,
      async () => controller.stopActiveFileRuntimeContext(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      PUT_RUNTIME_VALUE_COMMAND,
      async (target) => controller.putRuntimeValue(target),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_PROJECT_RUNTIME_CONFIGURATION_COMMAND,
      async () => controller.openProjectRuntimeConfiguration(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor((editor) => {
      controller.handleActiveEditorChange(editor);
      controller.refreshMonitorHoverDecorationsForEditor(editor);
      controller.refreshDatabaseTocValueDecorationsForEditor(editor);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidChangeVisibleTextEditors(() => {
      controller.refreshVisibleMonitorHoverDecorations();
    }),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidChangeTextEditorSelection((event) => {
      void controller.handleTextEditorSelectionChanged(event);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.workspace.onDidChangeTextDocument((event) => {
      void controller.handleDocumentChanged(event);
      controller.refreshMonitorDiagnosticsForDocument(event.document);
      controller.refreshMonitorHoverDecorationsForDocument(event.document);
      controller.refreshDatabaseTocValueDecorationsForDocument(event.document);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((document) => {
      controller.refreshMonitorDiagnosticsForDocument(document);
      controller.refreshMonitorHoverDecorationsForDocument(document);
      controller.refreshDatabaseTocValueDecorationsForDocument(document);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.workspace.onDidCloseTextDocument((document) => {
      controller.clearMonitorDiagnosticsForDocument(document);
      controller.clearMonitorHoverDecorationsForDocument(document);
      controller.clearDatabaseTocValueDecorationsForDocument(document);
      void controller.handleDocumentClosed(document);
    }),
  );

  const hoverRefreshTimer = setInterval(() => {
    controller.refreshVisibleMonitorHoverDecorations();
  }, 1000);
  controller.setHoverRefreshTimer(hoverRefreshTimer);

  extensionContext.subscriptions.push(
    new vscode.Disposable(() => {
      controller.disposeHoverRefreshTimer();
    }),
  );

  controller.handleActiveEditorChange(vscode.window.activeTextEditor);
  for (const document of vscode.workspace.textDocuments) {
    controller.refreshMonitorDiagnosticsForDocument(document);
  }
  controller.refreshVisibleMonitorHoverDecorations();
}

class EpicsRuntimeMonitorController {
  constructor(databaseHelpers = {}) {
    this._onDidChangeTreeData = new vscode.EventEmitter();
    this.onDidChangeTreeData = this._onDidChangeTreeData.event;
    this.contextNode = { type: "context" };
    this.contextStatus = "stopped";
    this.contextError = undefined;
    this.runtimeContext = undefined;
    this.runtimeLibrary = undefined;
    this.contextInitializationPromise = undefined;
    this.contextGeneration = 0;
    this.monitorEntries = [];
    this.monitorDiagnostics = undefined;
    this.statusBarItem = undefined;
    this.hoverDecorationType = undefined;
    this.databaseTocValueDecorationType = undefined;
    this.hoverRefreshTimer = undefined;
    this.runtimeWorkspaceFolder = undefined;
    this.runtimeConfigurationPanel = undefined;
    this.activePutRequestKeys = new Set();
    this.lastMousePutRequest = undefined;
    this.extractDatabaseTocEntries =
      typeof databaseHelpers.extractDatabaseTocEntries === "function"
        ? databaseHelpers.extractDatabaseTocEntries
        : undefined;
    this.extractDatabaseTocMacroAssignments =
      typeof databaseHelpers.extractDatabaseTocMacroAssignments === "function"
        ? databaseHelpers.extractDatabaseTocMacroAssignments
        : undefined;
    this.extractRecordDeclarations =
      typeof databaseHelpers.extractRecordDeclarations === "function"
        ? databaseHelpers.extractRecordDeclarations
        : undefined;
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
    this.disposeHoverRefreshTimer();
    this.runtimeConfigurationPanel?.dispose();
    this.runtimeConfigurationPanel = undefined;
    this.statusBarItem?.dispose();
    this._onDidChangeTreeData.dispose();
  }

  attachStatusBar(statusBarItem) {
    this.statusBarItem = statusBarItem;
  }

  attachDiagnosticsCollection(diagnostics) {
    this.monitorDiagnostics = diagnostics;
  }

  attachHoverDecorationType(decorationType) {
    this.hoverDecorationType = decorationType;
  }

  attachDatabaseTocValueDecorationType(decorationType) {
    this.databaseTocValueDecorationType = decorationType;
  }

  setHoverRefreshTimer(timer) {
    this.disposeHoverRefreshTimer();
    this.hoverRefreshTimer = timer;
  }

  disposeHoverRefreshTimer() {
    if (this.hoverRefreshTimer) {
      clearInterval(this.hoverRefreshTimer);
      this.hoverRefreshTimer = undefined;
    }
  }

  handleActiveEditorChange(editor) {
    if (!this.statusBarItem) {
      return;
    }

    const document = editor?.document;
    if (!isRuntimeDocument(document)) {
      this.statusBarItem.hide();
      return;
    }

    const documentLabel = getRuntimeDocumentLabel(document);
    const isRunning = this.isDocumentMonitoringRunning(document);
    this.statusBarItem.text = isRunning ? "$(primitive-square) EPICS" : "$(play) EPICS";
    this.statusBarItem.command = isRunning
      ? STOP_ACTIVE_FILE_RUNTIME_COMMAND
      : START_ACTIVE_FILE_RUNTIME_COMMAND;
    this.statusBarItem.tooltip = isRunning
      ? `Stop EPICS runtime monitoring for ${documentLabel}`
      : `Start EPICS runtime monitoring for ${documentLabel}`;
    this.statusBarItem.show();
  }

  refreshMonitorDiagnosticsForDocument(document) {
    if (!this.monitorDiagnostics || !document?.uri) {
      return;
    }

    if (!isStrictMonitorDocument(document)) {
      this.monitorDiagnostics.delete(document.uri);
      return;
    }

    const analysis = analyzeStrictMonitorDocument(
      document,
      this.getDefaultProtocol(),
    );
    const diagnostics = analysis.diagnostics.map((diagnostic) =>
      createMonitorDiagnostic(diagnostic),
    );
    this.monitorDiagnostics.set(document.uri, diagnostics);
  }

  clearMonitorDiagnosticsForDocument(document) {
    if (!this.monitorDiagnostics || !document?.uri) {
      return;
    }

    this.monitorDiagnostics.delete(document.uri);
  }

  refreshVisibleMonitorHoverDecorations() {
    if (!this.hoverDecorationType && !this.databaseTocValueDecorationType) {
      return;
    }

    this.reconcileMonitorStates();

    for (const editor of vscode.window.visibleTextEditors) {
      this.refreshMonitorHoverDecorationsForEditor(editor);
      this.refreshDatabaseTocValueDecorationsForEditor(editor);
    }
  }

  reconcileMonitorStates() {
    for (const entry of this.monitorEntries) {
      const channelState = getRuntimeChannelState(entry.channel);
      const monitorState = getRuntimeMonitorState(entry.monitor);
      let didChange = false;
      let shouldRecover = false;

      if (monitorState === "SUBSCRIBED" && entry.monitor) {
        if (entry.status !== "subscribed" || entry.lastError) {
          entry.status = "subscribed";
          entry.lastError = undefined;
          didChange = true;
        }

        const previousValueText = entry.valueText;
        const previousTimestamp = entry.lastUpdated?.getTime();
        this.updateEntryValue(entry, entry.monitor);
        if (
          entry.valueText !== previousValueText ||
          entry.lastUpdated?.getTime() !== previousTimestamp
        ) {
          didChange = true;
        }
      } else if (
        channelState === "CREATED" &&
        (!entry.monitor || monitorState === "FAILED" || monitorState === "DESTROYED")
      ) {
        if (entry.status !== "connecting" || entry.lastError) {
          entry.status = "connecting";
          entry.lastError = undefined;
          didChange = true;
        }
        shouldRecover = true;
      } else if (
        channelState === "RECREATING" ||
        channelState === "SEARCHING" ||
        channelState === "RESOLVED" ||
        monitorState === "SUBSCRIBING"
      ) {
        if (entry.status !== "connecting" || entry.lastError) {
          entry.status = "connecting";
          entry.lastError = undefined;
          didChange = true;
        }
      } else if (
        channelState === "DISCONNECTED" ||
        monitorState === "FAILED"
      ) {
        const nextError = "Channel disconnected. Waiting for recovery.";
        if (entry.status !== "disconnected" || entry.lastError !== nextError) {
          entry.status = "disconnected";
          entry.lastError = nextError;
          didChange = true;
        }
      } else if (
        channelState === "DESTROYED" ||
        monitorState === "DESTROYED"
      ) {
        const nextError = "Channel was destroyed.";
        if (entry.status !== "destroyed" || entry.lastError !== nextError) {
          entry.status = "destroyed";
          entry.lastError = nextError;
          didChange = true;
        }
        shouldRecover = true;
      } else if (!entry.channel && this.contextStatus === "connected") {
        shouldRecover = true;
      } else if (entry.status === "error" && this.contextStatus === "connected") {
        shouldRecover = true;
      }

      if (didChange) {
        this.refresh(entry);
      }
      if (shouldRecover) {
        void this.tryRecoverEntry(entry);
      }
    }
  }

  async tryRecoverEntry(entry) {
    if (
      !entry ||
      !this.monitorEntries.includes(entry) ||
      entry.recoveryInProgress ||
      this.contextStatus !== "connected"
    ) {
      return;
    }

    const channelState = getRuntimeChannelState(entry.channel);
    if (
      channelState === "SEARCHING" ||
      channelState === "RESOLVED" ||
      channelState === "RECREATING"
    ) {
      return;
    }

    entry.recoveryInProgress = true;
    try {
      if (!entry.channel || channelState === "DESTROYED" || entry.status === "error") {
        const runtimeContext = await this.ensureRuntimeContext();
        await this.connectEntry(entry, runtimeContext);
        return;
      }

      if (channelState !== "CREATED") {
        return;
      }

      const monitorState = getRuntimeMonitorState(entry.monitor);
      if (
        entry.monitor &&
        monitorState === "FAILED" &&
        typeof entry.monitor.resubscribe === "function"
      ) {
        await entry.monitor.resubscribe();
        this.applyMonitorState(entry, entry.monitor);
        return;
      }

      if (!entry.monitor || monitorState === "DESTROYED") {
        const monitor = await this.createMonitorForEntry(entry, entry.channel);
        if (!monitor) {
          return;
        }
        entry.monitor = monitor;
        this.applyMonitorState(entry, monitor);
      }
    } catch (error) {
      entry.status = "error";
      entry.lastError = getErrorMessage(error);
      this.refresh(entry);
    } finally {
      entry.recoveryInProgress = false;
    }
  }

  refreshMonitorHoverDecorationsForDocument(document) {
    if (!isStrictMonitorDocument(document) || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        this.refreshMonitorHoverDecorationsForEditor(editor);
      }
    }
  }

  clearMonitorHoverDecorationsForDocument(document) {
    if (!this.hoverDecorationType || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        editor.setDecorations(this.hoverDecorationType, []);
      }
    }
  }

  refreshDatabaseTocValueDecorationsForDocument(document) {
    if (!isDatabaseRuntimeDocument(document) || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        this.refreshDatabaseTocValueDecorationsForEditor(editor);
      }
    }
  }

  clearDatabaseTocValueDecorationsForDocument(document) {
    if (!this.databaseTocValueDecorationType || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        editor.setDecorations(this.databaseTocValueDecorationType, []);
      }
    }
  }

  refreshMonitorHoverDecorationsForEditor(editor) {
    if (!this.hoverDecorationType || !editor) {
      return;
    }

    const document = editor.document;
    if (!isStrictMonitorDocument(document) || !document?.uri) {
      editor.setDecorations(this.hoverDecorationType, []);
      return;
    }

    const analysis = analyzeStrictMonitorDocument(
      document,
      this.getDefaultProtocol(),
    );
    const sourceUri = document.uri.toString();
    let valueColumn = 0;

    for (const reference of analysis.lineReferences) {
      if (!reference) {
        continue;
      }

      valueColumn = Math.max(valueColumn, reference.endCharacter);
    }

    const decorationOptions = [];

    for (let lineNumber = 0; lineNumber < document.lineCount; lineNumber += 1) {
      const reference = analysis.lineReferences[lineNumber];
      if (!reference) {
        continue;
      }

      const entry = this.findDocumentMonitorEntry(sourceUri, reference);
      if (!entry) {
        continue;
      }

      decorationOptions.push({
        range: new vscode.Range(
          lineNumber,
          reference.startCharacter,
          lineNumber,
          reference.endCharacter,
        ),
        hoverMessage: this.createMonitorHoverMarkdown(entry),
        renderOptions: {
          after: {
            contentText: buildAlignedMonitorInlineText(entry),
            margin: buildAlignedMonitorInlineMargin(
              valueColumn,
              reference.endCharacter,
            ),
          },
        },
      });
    }

    editor.setDecorations(this.hoverDecorationType, decorationOptions);
  }

  refreshDatabaseTocValueDecorationsForEditor(editor) {
    if (!this.databaseTocValueDecorationType || !editor) {
      return;
    }

    const document = editor.document;
    if (!isDatabaseRuntimeDocument(document) || !document?.uri) {
      editor.setDecorations(this.databaseTocValueDecorationType, []);
      return;
    }

    const tocEntries = this.extractDatabaseTocEntries?.(document.getText()) || [];
    if (tocEntries.length === 0) {
      editor.setDecorations(this.databaseTocValueDecorationType, []);
      return;
    }

    const sourceUri = document.uri.toString();
    const macroDefinitions = createDatabaseMonitorMacroDefinitions(
      this.extractDatabaseTocMacroAssignments?.(document.getText()) || new Map(),
    );
    const macroExpansionCache = new Map();
    const decorationOptions = [];
    for (const tocEntry of tocEntries) {
      if (typeof tocEntry.valueStart !== "number") {
        continue;
      }

      const entry = this.findDatabaseTocMonitorEntry(
        sourceUri,
        tocEntry,
        macroDefinitions,
        macroExpansionCache,
      );
      if (!entry) {
        continue;
      }

      const contentText = buildDatabaseTocValueText(entry);
      if (!contentText) {
        continue;
      }

      const startPosition = document.positionAt(tocEntry.valueStart);
      decorationOptions.push({
        range: new vscode.Range(startPosition, startPosition),
        hoverMessage: this.createMonitorHoverMarkdown(entry),
        renderOptions: {
          after: {
            contentText,
            width: "0",
          },
        },
      });
    }

    editor.setDecorations(this.databaseTocValueDecorationType, decorationOptions);
  }

  findDocumentMonitorEntry(sourceUri, reference) {
    if (!sourceUri || !reference) {
      return undefined;
    }

    return this.monitorEntries.find(
      (candidate) =>
        candidate.sourceUri === sourceUri &&
        candidate.protocol === reference.protocol &&
        candidate.pvName === reference.pvName &&
        (candidate.pvRequest || "") === (reference.pvRequest || ""),
    );
  }

  findDatabaseTocMonitorEntry(
    sourceUri,
    tocEntry,
    macroDefinitions = new Map(),
    macroExpansionCache = new Map(),
  ) {
    if (!sourceUri || !tocEntry) {
      return undefined;
    }

    const directMatch = this.monitorEntries.find(
      (candidate) =>
        candidate.sourceUri === sourceUri &&
        candidate.tocRecordName === tocEntry.recordName &&
        candidate.recordType === tocEntry.recordType,
    );
    if (directMatch) {
      return directMatch;
    }

    const expandedPvName = normalizeDatabaseMonitorPvName(
      expandDatabaseMonitorValue(
        tocEntry.recordName,
        macroDefinitions,
        macroExpansionCache,
        [],
      ),
      tocEntry.recordName,
    );
    return this.monitorEntries.find(
      (candidate) =>
        candidate.sourceUri === sourceUri && candidate.pvName === expandedPvName,
    );
  }

  createMonitorHoverMarkdown(entry) {
    const markdown = new vscode.MarkdownString();
    markdown.isTrusted = true;
    markdown.appendMarkdown(`**${escapeMarkdownText(entry.pvName)}**`);
    markdown.appendMarkdown(
      `\n\nProtocol: \`${entry.protocol.toUpperCase()}\``,
    );
    markdown.appendMarkdown(`\n\nStatus: \`${entry.status}\``);

    const hoverValue = getMonitorHoverValue(entry);
    if (hoverValue !== undefined) {
      markdown.appendMarkdown(
        `\n\nValue: \`${escapeInlineCode(formatRuntimeHoverValue(hoverValue))}\``,
      );
    }
    if (entry.lastUpdated) {
      markdown.appendMarkdown(
        `\n\nUpdated: ${escapeMarkdownText(entry.lastUpdated.toLocaleTimeString())}`,
      );
    }
    if (entry.lastError) {
      markdown.appendMarkdown(
        `\n\nError: ${escapeMarkdownText(entry.lastError)}`,
      );
    }
    if (this.canPutRuntimeValue(entry)) {
      markdown.appendMarkdown(
        `\n\n[Put Value](${buildPutRuntimeValueCommandUri(entry)})`,
      );
    }

    return markdown;
  }

  async handleTextEditorSelectionChanged(event) {
    if (
      !event?.textEditor ||
      event.kind !== vscode.TextEditorSelectionChangeKind.Mouse ||
      !Array.isArray(event.selections) ||
      event.selections.length !== 1
    ) {
      return;
    }

    const document = event.textEditor.document;
    const position = event.selections[0].active;
    const targetEntry = this.findRuntimeEntryAtDocumentPosition(
      document,
      position,
      true,
    );
    if (!targetEntry) {
      return;
    }

    const requestKey = `${targetEntry.key}:${position.line}`;
    const now = Date.now();
    const isDoubleClick =
      this.lastMousePutRequest?.key === requestKey &&
      now - this.lastMousePutRequest.time <= MOUSE_DOUBLE_CLICK_INTERVAL_MS;
    this.lastMousePutRequest = {
      key: requestKey,
      time: now,
    };
    if (!isDoubleClick) {
      return;
    }
    this.lastMousePutRequest = undefined;

    await this.putRuntimeValue(targetEntry);
  }

  async handleDocumentClosed(document) {
    if (!document?.uri) {
      return;
    }

    await this.removeEntriesBySourceUri(document.uri.toString());
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async handleDocumentChanged(event) {
    const document = event?.document;
    if (!document?.uri || !event?.contentChanges?.length) {
      return;
    }

    if (!this.hasEntriesForSourceUri(document.uri.toString())) {
      return;
    }

    await this.removeEntriesBySourceUri(document.uri.toString());
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
    const queuedEntry = this.queueMonitorEntry(definition, {
      notifyIfActiveDuplicate: !definition.sourceUri,
    });
    if (!queuedEntry) {
      return;
    }

    await this.connectEntry(queuedEntry);
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
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
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
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
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
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
        try {
          const runtimeContext = await this.ensureRuntimeContext();
          this.handleActiveEditorChange(vscode.window.activeTextEditor);
          void this.connectEntriesInParallel(this.monitorEntries, runtimeContext);
        } catch (error) {
          if (!isContextInitializationCancelledError(error)) {
            throw error;
          }
        }
      },
    );
  }

  async stopContext() {
    this.stopContextInternal();
    this.refresh();
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async startActiveFileRuntimeContext() {
    const editor = vscode.window.activeTextEditor;
    const document = editor?.document;
    if (!isRuntimeDocument(document)) {
      vscode.window.showWarningMessage(
        "Open a database, template, .monitor, or plain text file to start file runtime monitoring.",
      );
      return;
    }

    const workspaceFolder =
      this.getWorkspaceFolderForDocument(document) || this.getDefaultWorkspaceFolder();
    if (workspaceFolder) {
      this.runtimeWorkspaceFolder = workspaceFolder;
    }

    const analysis = analyzeRuntimeDocument(
      document,
      this.getDefaultProtocol(),
      this.getDatabaseRuntimeHelpers(),
    );
    const sourceLabel = getRuntimeDocumentLabel(document);
    if (analysis.diagnostics?.length) {
      vscode.window.showErrorMessage(
        `Cannot start EPICS runtime for ${sourceLabel} until the .monitor file errors are fixed.`,
      );
      return;
    }

    const definitions = analysis.definitions;
    if (!definitions.length) {
      vscode.window.showWarningMessage(
        `No EPICS monitor targets were found in ${sourceLabel}.`,
      );
      return;
    }

    const sourceUri = document.uri.toString();
    await this.removeEntriesBySourceUri(sourceUri);

    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `Starting EPICS runtime for ${sourceLabel}`,
      },
      async () => {
        const queuedEntries = definitions
          .map((definition) =>
            this.queueMonitorEntry(
              {
                ...definition,
                sourceUri,
                sourceLabel,
              },
              {
                notifyIfActiveDuplicate: false,
              },
            ),
          )
          .filter(Boolean);
        try {
          const runtimeContext = await this.ensureRuntimeContext();
          this.handleActiveEditorChange(vscode.window.activeTextEditor);
          void this.connectEntriesInParallel(queuedEntries, runtimeContext);
        } catch (error) {
          if (!isContextInitializationCancelledError(error)) {
            throw error;
          }
        }
      },
    );

    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async stopActiveFileRuntimeContext() {
    const editor = vscode.window.activeTextEditor;
    const document = editor?.document;
    if (!document?.uri) {
      return;
    }

    await this.removeEntriesBySourceUri(document.uri.toString());
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async openProjectRuntimeConfiguration() {
    const workspaceFolder = await this.resolveInteractiveWorkspaceFolder();
    if (!workspaceFolder) {
      vscode.window.showWarningMessage(
        "Open a folder or workspace to configure EPICS runtime environment for this project.",
      );
      return;
    }

    this.runtimeWorkspaceFolder = workspaceFolder;
    const configPath = getProjectRuntimeConfigPath(workspaceFolder);
    const loadedConfig = loadProjectRuntimeConfiguration(configPath);
    if (loadedConfig.error) {
      vscode.window.showWarningMessage(
        `Failed to read ${PROJECT_RUNTIME_CONFIG_FILE_NAME}: ${loadedConfig.error}`,
      );
    }

    this.runtimeConfigurationPanel?.dispose();
    const panel = vscode.window.createWebviewPanel(
      PROJECT_RUNTIME_CONFIGURATION_VIEW_TYPE,
      `EPICS Runtime Config: ${workspaceFolder.name}`,
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );
    this.runtimeConfigurationPanel = panel;
    let currentConfig = loadedConfig.config;
    panel.webview.html = buildProjectRuntimeConfigurationWebviewHtml(
      panel.webview,
      workspaceFolder,
      configPath,
      loadedConfig.config,
    );
    panel.onDidDispose(() => {
      if (this.runtimeConfigurationPanel === panel) {
        this.runtimeConfigurationPanel = undefined;
      }
    });
    panel.webview.onDidReceiveMessage(async (message) => {
      if (message?.type !== "saveProjectRuntimeConfiguration") {
        return;
      }

      const normalizedConfig = normalizeProjectRuntimeConfiguration(message.config);
      const protocolChanged = normalizedConfig.protocol !== currentConfig.protocol;
      try {
        await fs.promises.writeFile(
          configPath,
          serializeProjectRuntimeConfiguration(normalizedConfig),
          "utf8",
        );
      } catch (error) {
        const failureMessage = getErrorMessage(error);
        await panel.webview.postMessage({
          type: "saveProjectRuntimeConfigurationResult",
          success: false,
          message: failureMessage,
        });
        vscode.window.showErrorMessage(
          `Failed to save ${PROJECT_RUNTIME_CONFIG_FILE_NAME}: ${failureMessage}`,
        );
        return;
      }

      await panel.webview.postMessage({
        type: "saveProjectRuntimeConfigurationResult",
        success: true,
        message: `Saved to ${configPath}`,
      });
      currentConfig = normalizedConfig;
      const runtimeIsActive =
        this.contextStatus !== "stopped" || Boolean(this.contextInitializationPromise);
      if (runtimeIsActive) {
        const restartMessage = protocolChanged
          ? `Saved ${PROJECT_RUNTIME_CONFIG_FILE_NAME}. Restart EPICS runtime context to apply environment changes. Protocol changes are used the next time file runtime monitoring starts.`
          : `Saved ${PROJECT_RUNTIME_CONFIG_FILE_NAME}. Restart EPICS runtime context to apply the new environment values?`;
        const restartChoice = await vscode.window.showInformationMessage(
          restartMessage,
          "Restart",
        );
        if (restartChoice === "Restart") {
          await this.restartContext();
        }
      }
    });
  }

  async putRuntimeValue(target) {
    const entry = this.resolveRuntimePutEntry(target);
    if (!entry) {
      return;
    }

    const putSupport = this.getRuntimePutSupport(entry);
    if (!putSupport.canPut) {
      vscode.window.showWarningMessage(putSupport.reason);
      return;
    }

    if (this.activePutRequestKeys.has(entry.key)) {
      return;
    }

    this.activePutRequestKeys.add(entry.key);
    const input = await showRuntimePutInput(entry, putSupport);
    if (input === undefined) {
      this.activePutRequestKeys.delete(entry.key);
      return;
    }

    const parsed = parseRuntimePutInput(input, putSupport);
    if (parsed.error) {
      this.activePutRequestKeys.delete(entry.key);
      vscode.window.showErrorMessage(parsed.error);
      return;
    }

    entry.lastError = undefined;
    this.refresh(entry);
    try {
      const runtimeLibrary = this.requireRuntimeLibrary();
      const result = entry.protocol === "pva"
        ? await entry.channel.putPva(putSupport.putPvRequest || "value", [parsed.value])
        : await entry.channel.put(parsed.value, undefined, true);
      if (!isSuccessfulRuntimePutResult(result, runtimeLibrary, entry.protocol)) {
        throw new Error(
          getRuntimePutFailureMessage(result, runtimeLibrary, entry.protocol),
        );
      }
    } catch (error) {
      const message = getErrorMessage(error);
      entry.lastError = message;
      this.refresh(entry);
      vscode.window.showErrorMessage(`Failed to put ${entry.pvName}: ${message}`);
      return;
    } finally {
      this.activePutRequestKeys.delete(entry.key);
    }

    this.refresh(entry);
  }

  queueMonitorEntry(definition, options = {}) {
    const key = createMonitorKey(definition);
    const existingEntry = this.monitorEntries.find((entry) => entry.key === key);
    if (existingEntry) {
      if (existingEntry.monitor) {
        if (options.notifyIfActiveDuplicate) {
          vscode.window.showInformationMessage(
            `${existingEntry.pvName} is already being monitored.`,
          );
        }
        return undefined;
      }

      return existingEntry;
    }

    const entry = {
      type: "monitor",
      key,
      pvName: definition.pvName,
      protocol: definition.protocol,
      pvRequest: definition.pvRequest || "",
      tocRecordName: definition.tocRecordName,
      recordType: definition.recordType,
      sourceUri: definition.sourceUri,
      sourceLabel: definition.sourceLabel,
      channel: undefined,
      monitor: undefined,
      status: "pending",
      valueText: "",
      lastUpdated: undefined,
      lastError: undefined,
      serverAddress: undefined,
      connectionAttemptId: 0,
      recoveryInProgress: false,
      caEnumChoices: undefined,
      pvaEnumChoices: undefined,
    };
    this.monitorEntries.push(entry);
    this.refresh();
    return entry;
  }

  async connectEntriesInParallel(entries, runtimeContext) {
    await Promise.allSettled(
      (entries || []).map((entry) => this.connectEntry(entry, runtimeContext)),
    );
  }

  async connectEntry(entry, runtimeContext) {
    if (!entry || !this.monitorEntries.includes(entry)) {
      return;
    }

    const attemptId = Number(entry.connectionAttemptId || 0) + 1;
    entry.connectionAttemptId = attemptId;
    entry.status = "connecting";
    entry.lastError = undefined;
    this.refresh(entry);

    let channel;
    let monitor;
    let keepResources = false;

    try {
      const activeRuntimeContext = runtimeContext || await this.ensureRuntimeContext();
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }

      channel = await activeRuntimeContext.createChannel(
        entry.pvName,
        entry.protocol,
        this.getChannelCreationTimeoutSeconds(),
      );
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      if (!channel) {
        throw new Error(
          `Failed to create ${entry.protocol.toUpperCase()} channel for ${entry.pvName}.`,
        );
      }

      this.attachChannelToEntry(entry, channel, attemptId);
      monitor = await this.createMonitorForEntry(entry, channel);

      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      if (!monitor) {
        throw new Error(
          `Failed to subscribe to ${entry.protocol.toUpperCase()} monitor for ${entry.pvName}.`,
        );
      }

      entry.monitor = monitor;
      this.applyMonitorState(entry, monitor);
      keepResources = true;
    } catch (error) {
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      entry.status = "error";
      entry.channel = undefined;
      entry.monitor = undefined;
      entry.lastError = getErrorMessage(error);
      this.refresh(entry);
    } finally {
      if (!keepResources) {
        await this.cleanupConnectionResources(monitor, channel);
      }
      this.handleActiveEditorChange(vscode.window.activeTextEditor);
    }
  }

  attachChannelToEntry(entry, channel, attemptId) {
    channel.setDestroySoftCallback(() => {
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      entry.status = "disconnected";
      entry.lastError = "Channel disconnected. Waiting for recovery.";
      this.refresh(entry);
    });
    channel.setDestroyHardCallback(() => {
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      entry.status = "destroyed";
      entry.channel = undefined;
      entry.monitor = undefined;
      entry.lastError = "Channel was destroyed.";
      this.refresh(entry);
    });

    entry.channel = channel;
    entry.serverAddress = channel.getServerAddress();
  }

  async createMonitorForEntry(entry, channel) {
    if (entry.protocol === "pva") {
      return channel.createMonitorPva(
        this.getMonitorSubscribeTimeoutSeconds(),
        entry.pvRequest,
        (activeMonitor) => {
          this.handleMonitorUpdate(entry, activeMonitor);
        },
      );
    }

    const caMonitorOptions = getCaEnumMonitorOptions(channel);
    if (caMonitorOptions) {
      return channel.createMonitor(
        this.getMonitorSubscribeTimeoutSeconds(),
        (activeMonitor) => {
          this.handleMonitorUpdate(entry, activeMonitor);
        },
        caMonitorOptions.dbrType,
        caMonitorOptions.valueCount,
      );
    }

    return channel.createMonitor(
      this.getMonitorSubscribeTimeoutSeconds(),
      (activeMonitor) => {
        this.handleMonitorUpdate(entry, activeMonitor);
      },
    );
  }

  applyMonitorState(entry, monitor) {
    const monitorState = getRuntimeMonitorState(monitor);
    if (monitorState === "SUBSCRIBED") {
      entry.status = "subscribed";
      entry.lastError = undefined;
      this.updateEntryValue(entry, monitor);
      this.refresh(entry);
      return;
    }

    entry.status = monitorState === "FAILED" ? "disconnected" : "connecting";
    entry.lastError =
      monitorState === "FAILED"
        ? "Monitor subscription is waiting for recovery."
        : monitorState
          ? `Monitor state: ${monitorState.toLowerCase()}.`
          : undefined;
    this.refresh(entry);
  }

  async disconnectEntry(entry) {
    const monitor = entry.monitor;
    const channel = entry.channel;
    entry.connectionAttemptId = Number(entry.connectionAttemptId || 0) + 1;

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

  async cleanupConnectionResources(monitor, channel) {
    try {
      monitor?.destroyHard();
    } catch (error) {
      // Ignore teardown errors for abandoned async connection attempts.
    }

    try {
      await channel?.destroyHard();
    } catch (error) {
      // Ignore teardown errors for abandoned async connection attempts.
    }
  }

  handleMonitorUpdate(entry, monitor) {
    if (entry.monitor && entry.monitor !== monitor) {
      return;
    }

    entry.status = "subscribed";
    entry.lastError = undefined;
    this.updateEntryValue(entry, monitor);
    this.refresh(entry);
  }

  updateEntryValue(entry, monitor) {
    entry.lastUpdated = new Date();
    if (entry.protocol === "pva") {
      updatePvaEnumChoicesCache(entry, monitor.getPvaData());
      entry.valueText = formatRuntimeValue(
        getPvaRuntimeDisplayValue(monitor.getPvaData(), entry.pvaEnumChoices),
      );
      return;
    }

    updateCaEnumChoicesCache(entry, monitor.getChannel().getDbrData?.());
    entry.valueText = formatRuntimeValue(
      getCaRuntimeDisplayValue(entry, monitor.getChannel().getDbrData?.()),
    );
  }

  isCurrentConnectionAttempt(entry, attemptId) {
    return (
      Boolean(entry) &&
      this.monitorEntries.includes(entry) &&
      Number(entry.connectionAttemptId || 0) === Number(attemptId)
    );
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

    const contextGeneration = this.contextGeneration;
    const initializationPromise = (async () => {
      const { Context } = this.requireRuntimeLibrary();
      const runtimeContext = new Context(
        this.getRuntimeEnvironment(),
        this.getRuntimeLogLevel(),
      );
      await runtimeContext.initialize();
      if (contextGeneration !== this.contextGeneration) {
        try {
          runtimeContext.destroyHard();
        } catch (error) {
          // Ignore teardown errors for superseded context initialization.
        }
        throw new Error(CONTEXT_INITIALIZATION_CANCELLED_MESSAGE);
      }
      this.runtimeContext = runtimeContext;
      this.contextStatus = "connected";
      this.contextError = undefined;
      this.refresh(this.contextNode);
      return runtimeContext;
    })();
    this.contextInitializationPromise = initializationPromise;

    try {
      return await initializationPromise;
    } catch (error) {
      if (contextGeneration === this.contextGeneration) {
        this.runtimeContext = undefined;
        this.contextStatus = "error";
        this.contextError = getErrorMessage(error);
        this.refresh(this.contextNode);
      }
      throw error;
    } finally {
      if (this.contextInitializationPromise === initializationPromise) {
        this.contextInitializationPromise = undefined;
      }
    }
  }

  stopContextInternal() {
    this.contextGeneration += 1;
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
      entry.connectionAttemptId = Number(entry.connectionAttemptId || 0) + 1;
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
    treeItem.description = buildMonitorDescription(entry);
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
    const lines = [
      `Status: ${this.contextStatus}`,
      `Default protocol: ${this.getDefaultProtocol()}`,
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

    if (entry.sourceLabel) {
      lines.push(`Source: ${entry.sourceLabel}`);
    }
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
        detail: buildMonitorDescription(candidate),
        entry: candidate,
      })),
      {
        placeHolder: "Select the EPICS monitor to remove",
      },
    );
    return selected?.entry;
  }

  isDocumentMonitoringRunning(document) {
    if (!document?.uri) {
      return false;
    }

    return (
      this.hasEntriesForSourceUri(document.uri.toString()) &&
      this.contextStatus === "connected"
    );
  }

  hasEntriesForSourceUri(sourceUri) {
    if (!sourceUri) {
      return false;
    }

    return this.monitorEntries.some((entry) => entry.sourceUri === sourceUri);
  }

  async removeEntriesBySourceUri(sourceUri) {
    if (!sourceUri) {
      return;
    }

    const entries = this.monitorEntries.filter(
      (entry) => entry.sourceUri === sourceUri,
    );
    if (!entries.length) {
      return;
    }

    for (const entry of entries) {
      await this.disconnectEntry(entry);
    }

    this.monitorEntries = this.monitorEntries.filter(
      (entry) => entry.sourceUri !== sourceUri,
    );
    if (!this.monitorEntries.length) {
      this.stopContextInternal();
    }
    this.refresh();
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
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

  getDefaultProtocol() {
    const runtimeWorkspaceFolder = this.getRuntimeWorkspaceFolder();
    if (runtimeWorkspaceFolder) {
      const loadedProjectConfig = loadProjectRuntimeConfiguration(
        getProjectRuntimeConfigPath(runtimeWorkspaceFolder),
      );
      if (loadedProjectConfig.exists) {
        return normalizeRuntimeProtocol(loadedProjectConfig.config.protocol);
      }
    }

    return normalizeRuntimeProtocol(
      vscode.workspace
        .getConfiguration("epicsWorkbench.runtime")
        .get("defaultProtocol", DEFAULT_PROTOCOL),
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

  getWorkspaceFolderForDocument(document) {
    if (!document?.uri) {
      return undefined;
    }

    return vscode.workspace.getWorkspaceFolder(document.uri);
  }

  getDefaultWorkspaceFolder() {
    const [workspaceFolder] = vscode.workspace.workspaceFolders || [];
    return workspaceFolder;
  }

  getRuntimeWorkspaceFolder() {
    if (this.runtimeWorkspaceFolder) {
      return this.runtimeWorkspaceFolder;
    }

    const activeDocument = vscode.window.activeTextEditor?.document;
    return (
      this.getWorkspaceFolderForDocument(activeDocument) ||
      this.getDefaultWorkspaceFolder()
    );
  }

  async resolveInteractiveWorkspaceFolder() {
    const activeDocument = vscode.window.activeTextEditor?.document;
    const activeWorkspaceFolder = this.getWorkspaceFolderForDocument(activeDocument);
    if (activeWorkspaceFolder) {
      return activeWorkspaceFolder;
    }

    const workspaceFolders = vscode.workspace.workspaceFolders || [];
    if (workspaceFolders.length === 1) {
      return workspaceFolders[0];
    }
    if (workspaceFolders.length === 0) {
      return undefined;
    }

    const selected = await vscode.window.showQuickPick(
      workspaceFolders.map((candidate) => ({
        label: candidate.name,
        description: candidate.uri.fsPath,
        workspaceFolder: candidate,
      })),
      {
        placeHolder: "Select the workspace folder to store EPICS runtime configuration",
      },
    );
    return selected?.workspaceFolder;
  }

  getRuntimeEnvironment() {
    const raw = vscode.workspace
      .getConfiguration("epicsWorkbench.runtime")
      .get("environment", {});
    const environment = {};

    if (raw && typeof raw === "object" && !Array.isArray(raw)) {
      for (const [key, value] of Object.entries(raw)) {
        if (!key) {
          continue;
        }

        environment[key] = String(value);
      }
    }

    const runtimeWorkspaceFolder = this.getRuntimeWorkspaceFolder();
    if (runtimeWorkspaceFolder) {
      const loadedProjectConfig = loadProjectRuntimeConfiguration(
        getProjectRuntimeConfigPath(runtimeWorkspaceFolder),
      );
      if (loadedProjectConfig.exists) {
        Object.assign(
          environment,
          createRuntimeEnvironmentFromProjectConfiguration(
            loadedProjectConfig.config,
          ),
        );
      }
    }

    return environment;
  }

  getDatabaseRuntimeHelpers() {
    return {
      extractDatabaseTocMacroAssignments: this.extractDatabaseTocMacroAssignments,
      extractRecordDeclarations: this.extractRecordDeclarations,
    };
  }

  resolveRuntimePutEntry(target) {
    if (target?.type === "monitor" && this.monitorEntries.includes(target)) {
      return target;
    }

    if (target?.key) {
      return this.monitorEntries.find((entry) => entry.key === target.key);
    }

    const editor = vscode.window.activeTextEditor;
    const selection = editor?.selection;
    if (!editor || !selection?.isEmpty) {
      return undefined;
    }

    return this.findRuntimeEntryAtDocumentPosition(
      editor.document,
      selection.active,
      false,
    );
  }

  findRuntimeEntryAtDocumentPosition(
    document,
    position,
    requireValueSurfaceHit,
  ) {
    if (!document?.uri || !position) {
      return undefined;
    }

    if (isStrictMonitorDocument(document)) {
      const analysis = analyzeStrictMonitorDocument(
        document,
        this.getDefaultProtocol(),
      );
      const reference = analysis.lineReferences[position.line];
      if (!reference) {
        return undefined;
      }

      const isWithinRecordName =
        position.character >= reference.startCharacter &&
        position.character <= reference.endCharacter;
      const isWithinInlineValue = position.character >= reference.endCharacter;
      if (
        (requireValueSurfaceHit && !isWithinInlineValue) ||
        (!requireValueSurfaceHit && !isWithinRecordName && !isWithinInlineValue)
      ) {
        return undefined;
      }

      return this.findDocumentMonitorEntry(document.uri.toString(), reference);
    }

    if (!isDatabaseRuntimeDocument(document)) {
      return undefined;
    }

    const tocEntries = this.extractDatabaseTocEntries?.(document.getText()) || [];
    if (tocEntries.length === 0) {
      return undefined;
    }

    const offset = document.offsetAt(position);
    const macroDefinitions = createDatabaseMonitorMacroDefinitions(
      this.extractDatabaseTocMacroAssignments?.(document.getText()) || new Map(),
    );
    const macroExpansionCache = new Map();
    for (const tocEntry of tocEntries) {
      const isWithinValueCell = isWithinDatabaseTocValueHitArea(tocEntry, offset);
      const isWithinRow =
        typeof tocEntry.linkStart === "number" &&
        typeof tocEntry.linkEnd === "number" &&
        offset >= tocEntry.linkStart &&
        offset <= tocEntry.linkEnd;
      if (
        (requireValueSurfaceHit && !isWithinValueCell) ||
        (!requireValueSurfaceHit && !isWithinValueCell && !isWithinRow)
      ) {
        continue;
      }

      const entry = this.findDatabaseTocMonitorEntry(
        document.uri.toString(),
        tocEntry,
        macroDefinitions,
        macroExpansionCache,
      );
      if (entry) {
        return entry;
      }
    }

    return undefined;
  }

  canPutRuntimeValue(entry) {
    return this.getRuntimePutSupport(entry).canPut;
  }

  getRuntimePutSupport(entry) {
    if (!entry?.channel) {
      return {
        canPut: false,
        reason: "This EPICS channel is not connected.",
      };
    }

    if (entry.status !== "subscribed") {
      return {
        canPut: false,
        reason: `${entry.pvName} is not currently subscribed.`,
      };
    }

    if (entry.protocol === "pva") {
      const pvaData = entry.monitor?.getPvaData?.();
      const currentValue = resolvePvaRuntimeValue(pvaData, entry.pvaEnumChoices);
      if (currentValue === undefined) {
        return {
          canPut: false,
          reason: hasPvaRuntimeDataWithoutValue(pvaData)
            ? `${entry.pvName} has data, but no value field.`
            : `No runtime value is available for ${entry.pvName}.`,
        };
      }

      if (Array.isArray(currentValue)) {
        return {
          canPut: false,
          reason: "Array values cannot be changed from inline runtime value editing.",
        };
      }

      if (isPvaEnumLikeValue(currentValue)) {
        return {
          canPut: true,
          valueKind: "pva-enum",
          typeLabel: "PVA enum",
          initialValue: String(getPvaEnumIndex(currentValue)),
          enumChoices: getPvaEnumChoices(currentValue),
          putPvRequest: "value.index",
        };
      }

      if (typeof currentValue === "string") {
        return {
          canPut: true,
          valueKind: "string",
          typeLabel: "PVA string",
          initialValue: currentValue,
          putPvRequest: "value",
        };
      }

      if (typeof currentValue === "number" && Number.isFinite(currentValue)) {
        return {
          canPut: true,
          valueKind: "number",
          typeLabel: "PVA number",
          initialValue: String(currentValue),
          putPvRequest: "value",
        };
      }

      return {
        canPut: false,
        reason: `${entry.pvName} uses an unsupported PVA value type for inline put.`,
      };
    }

    const runtimeLibrary = this.requireRuntimeLibrary();
    const accessRight = entry.channel.getAccessRight?.();
    if (
      accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NOT_AVAILABLE ||
      accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NO_ACCESS ||
      accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.READ_ONLY
    ) {
      return {
        canPut: false,
        reason: `${entry.pvName} is not writable.`,
      };
    }

    const valueCount = Number(entry.channel.getValueCount?.() || 0);
    if (valueCount > 1) {
      return {
        canPut: false,
        reason: "Array values cannot be changed from inline runtime value editing.",
      };
    }

    const currentValue = getMonitorHoverValue(entry);
    if (currentValue === undefined) {
      return {
        canPut: false,
        reason: `No runtime value is available for ${entry.pvName}.`,
      };
    }

    if (Array.isArray(currentValue)) {
      return {
        canPut: false,
        reason: "Array values cannot be changed from inline runtime value editing.",
      };
    }

    if (isPvaEnumLikeValue(currentValue)) {
      return {
        canPut: true,
        valueKind: "ca-enum",
        typeLabel: "CA enum",
        initialValue: String(getPvaEnumIndex(currentValue)),
        enumChoices: getPvaEnumChoices(currentValue),
      };
    }

    const dbrType = Number(entry.channel.getDbrType?.());
    const dbrTypeKind = getRuntimePutValueKind(dbrType, runtimeLibrary);
    if (!dbrTypeKind) {
      return {
        canPut: false,
        reason: `${entry.pvName} uses an unsupported DBR type for inline put.`,
      };
    }

    return {
      canPut: true,
      valueKind: dbrTypeKind,
      typeLabel: String(entry.channel.getDbrTypeStr?.() || "value"),
      initialValue: currentValue === undefined ? "" : String(currentValue),
    };
  }
}

class EpicsMonitorFormattingProvider {
  provideDocumentFormattingEdits(document) {
    const formattedText = formatMonitorText(document.getText());
    if (formattedText === document.getText()) {
      return [];
    }

    const fullRange = new vscode.Range(
      new vscode.Position(0, 0),
      document.positionAt(document.getText().length),
    );
    return [vscode.TextEdit.replace(fullRange, formattedText)];
  }
}

function createMonitorKey({ pvName, protocol, pvRequest, sourceUri }) {
  return `${sourceUri || "manual"}:${protocol}:${pvName}:${pvRequest || ""}`;
}

function buildPutRuntimeValueCommandUri(entry) {
  const commandArguments = encodeURIComponent(
    JSON.stringify([
      {
        key: entry.key,
      },
    ]),
  );
  return `command:${PUT_RUNTIME_VALUE_COMMAND}?${commandArguments}`;
}

function isWithinDatabaseTocValueHitArea(tocEntry, offset) {
  if (
    typeof tocEntry?.valueStart !== "number" ||
    typeof tocEntry?.hoverStart !== "number"
  ) {
    return false;
  }

  return offset >= tocEntry.valueStart && offset < tocEntry.hoverStart;
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

  return truncateText(
    formatRuntimeDisplayValue(value),
    RUNTIME_VALUE_DISPLAY_MAX_LENGTH,
  );
}

function getMonitorHoverValue(entry) {
  if (!entry) {
    return undefined;
  }

  try {
    if (entry.protocol === "pva") {
      return getPvaRuntimeDisplayValue(
        entry.monitor?.getPvaData(),
        entry.pvaEnumChoices,
      );
    }

    return getCaRuntimeDisplayValue(
      entry,
      entry.monitor?.getChannel().getDbrData?.(),
    );
  } catch (error) {
    return entry.valueText || undefined;
  }
}

function getRuntimeMonitorState(monitor) {
  if (!monitor || typeof monitor.getStateStr !== "function") {
    return undefined;
  }

  try {
    return String(monitor.getStateStr() || "");
  } catch (error) {
    return undefined;
  }
}

function getRuntimeChannelState(channel) {
  if (!channel || typeof channel.getStateStr !== "function") {
    return undefined;
  }

  try {
    return String(channel.getStateStr() || "");
  } catch (error) {
    return undefined;
  }
}

function buildMonitorDescription(entry) {
  const sourceText = entry.sourceLabel ? `${entry.sourceLabel}: ` : "";
  if (entry.valueText) {
    return truncateText(`${sourceText}${entry.valueText}`, 80);
  }

  return truncateText(`${sourceText}${entry.status}`, 80);
}

function buildMonitorInlineText(entry) {
  const hoverValue = getMonitorHoverValue(entry);
  if (hoverValue !== undefined) {
    return formatRuntimeHoverValue(hoverValue, RUNTIME_VALUE_DISPLAY_MAX_LENGTH);
  }

  if (entry.lastError) {
    return truncateText(`(${entry.lastError})`, 80);
  }

  return entry.status ? truncateText(`(${entry.status})`, 80) : "";
}

function buildDatabaseTocValueText(entry) {
  const hoverValue = getMonitorHoverValue(entry);
  if (hoverValue !== undefined) {
    return formatDatabaseTocRuntimeValue(hoverValue);
  }

  if (entry.status === "subscribed" && entry.valueText === "") {
    return '""';
  }

  if (entry.lastError) {
    return truncateText(
      `(${entry.lastError})`,
      DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH,
    );
  }

  return entry.status
    ? truncateText(`(${entry.status})`, DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH)
    : "";
}

function formatDatabaseTocRuntimeValue(value) {
  if (value === "") {
    return '""';
  }

  return truncateText(
    formatRuntimeDisplayValue(value),
    DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH,
  );
}

function getCaEnumMonitorOptions(channel) {
  if (!isCaEnumChannel(channel)) {
    return undefined;
  }

  const dbrType = Number(channel?.getDbrType_GR?.());
  const valueCount = Number(channel?.getValueCount?.() || 0);
  if (!Number.isFinite(dbrType) || dbrType < 0) {
    return undefined;
  }

  return {
    dbrType,
    valueCount: valueCount > 0 ? valueCount : 1,
  };
}

function isCaEnumChannel(channel) {
  return String(channel?.getDbrTypeStr?.() || "") === "DBR_ENUM";
}

function getCaRuntimeDisplayValue(entry, dbrData) {
  return resolveCaRuntimeValue(entry, dbrData);
}

function resolveCaRuntimeValue(entry, dbrData) {
  const value = dbrData?.value;
  if (!isCaEnumChannel(entry?.channel)) {
    return value;
  }

  if (Array.isArray(value) || typeof value !== "number" || !Number.isFinite(value)) {
    return value;
  }

  const choices = extractCaEnumChoices(dbrData, entry?.caEnumChoices);
  if (choices.length === 0) {
    return value;
  }

  return {
    index: value,
    choices,
  };
}

function extractCaEnumChoices(dbrData, fallbackChoices) {
  const rawChoices = Array.isArray(dbrData?.strings)
    ? dbrData.strings.map((choice) => String(choice ?? ""))
    : undefined;
  const rawCount = Number(dbrData?.number_of_string_used);
  const validCount = Number.isFinite(rawCount)
    ? Math.max(0, Math.min(rawChoices?.length || 0, rawCount))
    : 0;
  if (rawChoices && validCount > 0) {
    return rawChoices.slice(0, validCount);
  }

  if (Array.isArray(fallbackChoices)) {
    return fallbackChoices.map((choice) => String(choice ?? ""));
  }

  return [];
}

function updateCaEnumChoicesCache(entry, dbrData) {
  if (!entry) {
    return;
  }

  const choices = extractCaEnumChoices(dbrData);
  if (choices.length > 0) {
    entry.caEnumChoices = choices;
  }
}

function buildAlignedMonitorInlineText(entry) {
  const inlineText = buildMonitorInlineText(entry);
  if (!inlineText) {
    return "";
  }

  return `= ${inlineText}`;
}

function buildAlignedMonitorInlineMargin(valueColumn, currentColumn) {
  const paddingWidth = Math.max(1, Number(valueColumn || 0) - Number(currentColumn || 0) + 1);
  return `0 0 0 ${paddingWidth}ch`;
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

function isContextInitializationCancelledError(error) {
  return getErrorMessage(error) === CONTEXT_INITIALIZATION_CANCELLED_MESSAGE;
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

function isRuntimeDocument(document) {
  if (!document?.uri) {
    return false;
  }

  if (isDatabaseRuntimeDocument(document) || isStrictMonitorDocument(document)) {
    return true;
  }

  if (document.uri.scheme !== "file") {
    return false;
  }

  const extension = path.extname(document.uri.fsPath).toLowerCase();
  return (
    document.languageId === "plaintext" ||
    LINE_RUNTIME_EXTENSIONS.has(extension)
  );
}

function isStrictMonitorDocument(document) {
  if (!document?.uri) {
    return false;
  }

  if (document.languageId === "monitor") {
    return true;
  }

  return (
    document.uri.scheme === "file" &&
    path.extname(document.uri.fsPath).toLowerCase() === ".monitor"
  );
}

function isDatabaseRuntimeDocument(document) {
  if (!document?.uri) {
    return false;
  }

  if (document.languageId === "database") {
    return true;
  }

  return (
    document.uri.scheme === "file" &&
    DATABASE_RUNTIME_EXTENSIONS.has(path.extname(document.uri.fsPath).toLowerCase())
  );
}

function getRuntimeDocumentLabel(document) {
  if (!document) {
    return "Untitled";
  }

  const rawPath = document.fileName || document.uri?.fsPath || document.uri?.path || "";
  const baseName = path.basename(rawPath);
  if (baseName) {
    return baseName;
  }

  if (document.isUntitled) {
    return "Untitled";
  }

  return document.uri?.toString() || "Untitled";
}

function getProjectRuntimeConfigPath(workspaceFolder) {
  return path.join(workspaceFolder.uri.fsPath, PROJECT_RUNTIME_CONFIG_FILE_NAME);
}

function getDefaultProjectRuntimeConfiguration() {
  return {
    protocol: DEFAULT_PROTOCOL,
    EPICS_CA_ADDR_LIST: [],
    EPICS_CA_AUTO_ADDR_LIST: "Yes",
  };
}

function normalizeProjectRuntimeConfiguration(rawConfig) {
  const defaults = getDefaultProjectRuntimeConfiguration();
  const normalized = {
    protocol: defaults.protocol,
    EPICS_CA_ADDR_LIST: [],
    EPICS_CA_AUTO_ADDR_LIST: defaults.EPICS_CA_AUTO_ADDR_LIST,
  };
  const rawProtocol = rawConfig?.protocol ?? rawConfig?.defaultProtocol;
  normalized.protocol = PROJECT_RUNTIME_CONFIGURATION_PROTOCOL_VALUES.includes(
    normalizeRuntimeProtocol(rawProtocol),
  )
    ? normalizeRuntimeProtocol(rawProtocol)
    : defaults.protocol;
  const rawAddressList = rawConfig?.EPICS_CA_ADDR_LIST;
  const addressValues = Array.isArray(rawAddressList)
    ? rawAddressList
    : typeof rawAddressList === "string"
      ? rawAddressList.split(/\r?\n|,/)
      : [];
  normalized.EPICS_CA_ADDR_LIST = addressValues
    .map((value) => String(value ?? "").trim())
    .filter(Boolean);

  const rawAutoAddrList = String(rawConfig?.EPICS_CA_AUTO_ADDR_LIST || "").trim();
  normalized.EPICS_CA_AUTO_ADDR_LIST =
    PROJECT_RUNTIME_CONFIGURATION_AUTO_ADDR_LIST_VALUES.includes(rawAutoAddrList)
      ? rawAutoAddrList
      : defaults.EPICS_CA_AUTO_ADDR_LIST;

  return normalized;
}

function loadProjectRuntimeConfiguration(configPath) {
  const defaultConfig = getDefaultProjectRuntimeConfiguration();
  if (!configPath || !fs.existsSync(configPath)) {
    return {
      exists: false,
      config: defaultConfig,
    };
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(configPath, "utf8"));
    return {
      exists: true,
      config: normalizeProjectRuntimeConfiguration(parsed),
    };
  } catch (error) {
    return {
      exists: false,
      config: defaultConfig,
      error: getErrorMessage(error),
    };
  }
}

function createRuntimeEnvironmentFromProjectConfiguration(config) {
  const normalized = normalizeProjectRuntimeConfiguration(config);
  return {
    EPICS_CA_ADDR_LIST: normalized.EPICS_CA_ADDR_LIST.join(" "),
    EPICS_CA_AUTO_ADDR_LIST: normalized.EPICS_CA_AUTO_ADDR_LIST,
  };
}

function serializeProjectRuntimeConfiguration(config) {
  return `${JSON.stringify(normalizeProjectRuntimeConfiguration(config), null, 2)}\n`;
}

function buildProjectRuntimeConfigurationWebviewHtml(
  webview,
  workspaceFolder,
  configPath,
  initialConfig,
) {
  const nonce = createNonce();
  const initialState = JSON.stringify(
    normalizeProjectRuntimeConfiguration(initialConfig),
  ).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Runtime Config</title>
    <style>
      :root {
        color-scheme: light dark;
      }

      body {
        font-family: var(--vscode-font-family);
        color: var(--vscode-foreground);
        background: var(--vscode-editor-background);
        margin: 0;
        padding: 24px;
      }

      main {
        max-width: 760px;
        margin: 0 auto;
      }

      h1 {
        font-size: 1.4rem;
        margin: 0 0 8px;
      }

      p,
      label,
      button,
      input,
      select,
      code {
        font-size: 0.95rem;
      }

      .meta {
        color: var(--vscode-descriptionForeground);
        margin-bottom: 24px;
      }

      .section {
        border: 1px solid var(--vscode-panel-border);
        border-radius: 10px;
        padding: 18px;
        margin-bottom: 18px;
        background: color-mix(in srgb, var(--vscode-editor-background) 92%, var(--vscode-panel-border));
      }

      .section h2 {
        margin: 0 0 10px;
        font-size: 1.05rem;
      }

      .section p {
        margin: 0 0 14px;
        color: var(--vscode-descriptionForeground);
      }

      .address-list {
        display: grid;
        gap: 10px;
      }

      .address-row {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 10px;
      }

      input,
      select {
        width: 100%;
        box-sizing: border-box;
        padding: 8px 10px;
        border-radius: 6px;
        border: 1px solid var(--vscode-input-border);
        color: var(--vscode-input-foreground);
        background: var(--vscode-input-background);
      }

      .actions {
        display: flex;
        gap: 10px;
        align-items: center;
        margin-top: 16px;
      }

      button {
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 6px;
        padding: 8px 14px;
        cursor: pointer;
      }

      button.primary {
        color: var(--vscode-button-foreground);
        background: var(--vscode-button-background);
      }

      button.secondary {
        color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
        background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
      }

      .status {
        min-height: 1.4em;
        color: var(--vscode-descriptionForeground);
      }

      .status.success {
        color: var(--vscode-testing-iconPassed);
      }

      .status.error {
        color: var(--vscode-errorForeground);
      }
    </style>
  </head>
  <body>
    <main>
      <h1>EPICS Runtime Configuration</h1>
      <div class="meta">
        <div>Workspace: <code>${escapeHtml(workspaceFolder.name)}</code></div>
        <div>Config file: <code>${escapeHtml(configPath)}</code></div>
      </div>

      <section class="section">
        <h2>Protocol</h2>
        <p>Select the default EPICS protocol for runtime monitoring in this project.</p>
        <label for="protocol">Value</label>
        <select id="protocol">
          <option value="ca">Channel Access</option>
          <option value="pva">PV Access</option>
        </select>
      </section>

      <section class="section">
        <h2>EPICS_CA_ADDR_LIST</h2>
        <p>Array of Channel Access search addresses. It is saved as a JSON array and passed to epics-tca as a space-separated environment variable.</p>
        <div id="address-list" class="address-list"></div>
        <div class="actions">
          <button id="add-address" type="button" class="secondary">Add Address</button>
        </div>
      </section>

      <section class="section">
        <h2>EPICS_CA_AUTO_ADDR_LIST</h2>
        <p>Whether EPICS Channel Access should automatically add local broadcast addresses.</p>
        <label for="auto-addr-list">Value</label>
        <select id="auto-addr-list">
          <option value="Yes">Yes</option>
          <option value="No">No</option>
        </select>
      </section>

      <div class="actions">
        <button id="save-config" type="button" class="primary">Save Project Config</button>
        <div id="status" class="status"></div>
      </div>
    </main>

    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const initialConfig = ${initialState};
      const protocolElement = document.getElementById("protocol");
      const addressListElement = document.getElementById("address-list");
      const autoAddrListElement = document.getElementById("auto-addr-list");
      const statusElement = document.getElementById("status");

      function createAddressRow(value) {
        const rowElement = document.createElement("div");
        rowElement.className = "address-row";

        const inputElement = document.createElement("input");
        inputElement.type = "text";
        inputElement.className = "address-input";
        inputElement.placeholder = "192.168.1.10";
        inputElement.value = value || "";

        const removeButton = document.createElement("button");
        removeButton.type = "button";
        removeButton.className = "secondary";
        removeButton.textContent = "Remove";
        removeButton.addEventListener("click", () => {
          rowElement.remove();
          ensureAtLeastOneAddressRow();
        });

        rowElement.appendChild(inputElement);
        rowElement.appendChild(removeButton);
        addressListElement.appendChild(rowElement);
      }

      function ensureAtLeastOneAddressRow() {
        if (addressListElement.children.length === 0) {
          createAddressRow("");
        }
      }

      function getAddressValues() {
        return Array.from(document.querySelectorAll(".address-input"))
          .map((inputElement) => inputElement.value.trim())
          .filter(Boolean);
      }

      function setStatus(message, kind) {
        statusElement.textContent = message || "";
        statusElement.className = kind ? "status " + kind : "status";
      }

      document.getElementById("add-address").addEventListener("click", () => {
        createAddressRow("");
      });
      document.getElementById("save-config").addEventListener("click", () => {
        setStatus("Saving...", "");
        vscode.postMessage({
          type: "saveProjectRuntimeConfiguration",
          config: {
            protocol: protocolElement.value === "pva" ? "pva" : "ca",
            EPICS_CA_ADDR_LIST: getAddressValues(),
            EPICS_CA_AUTO_ADDR_LIST: autoAddrListElement.value === "No" ? "No" : "Yes",
          },
        });
      });

      window.addEventListener("message", (event) => {
        const message = event.data;
        if (message?.type !== "saveProjectRuntimeConfigurationResult") {
          return;
        }

        setStatus(message.message, message.success ? "success" : "error");
      });

      for (const addressValue of initialConfig.EPICS_CA_ADDR_LIST || []) {
        createAddressRow(addressValue);
      }
      ensureAtLeastOneAddressRow();
      protocolElement.value = initialConfig.protocol === "pva" ? "pva" : "ca";
      autoAddrListElement.value =
        initialConfig.EPICS_CA_AUTO_ADDR_LIST === "No" ? "No" : "Yes";
    </script>
  </body>
</html>`;
}

function createNonce() {
  return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function analyzeRuntimeDocument(document, defaultProtocol, databaseHelpers = {}) {
  if (!document) {
    return {
      definitions: [],
      diagnostics: [],
      lineReferences: [],
    };
  }

  if (isDatabaseRuntimeDocument(document)) {
    return {
      definitions: extractDatabaseMonitorDefinitions(
        document.getText(),
        defaultProtocol,
        databaseHelpers,
      ),
      diagnostics: [],
      lineReferences: [],
    };
  }

  if (isStrictMonitorDocument(document)) {
    return analyzeStrictMonitorDocument(document, defaultProtocol);
  }

  return {
    definitions: extractLineMonitorDefinitions(document.getText(), defaultProtocol),
    diagnostics: [],
    lineReferences: [],
  };
}

function analyzeStrictMonitorDocument(document, defaultProtocol) {
  return analyzeStrictMonitorText(document?.getText(), defaultProtocol);
}

function extractDatabaseMonitorDefinitions(text, defaultProtocol, databaseHelpers = {}) {
  const declarations =
    typeof databaseHelpers.extractRecordDeclarations === "function"
      ? databaseHelpers.extractRecordDeclarations(text)
      : extractDatabaseMonitorDeclarationsFallback(text);
  const macroDefinitions = createDatabaseMonitorMacroDefinitions(
    typeof databaseHelpers.extractDatabaseTocMacroAssignments === "function"
      ? databaseHelpers.extractDatabaseTocMacroAssignments(text)
      : new Map(),
  );
  const definitions = [];
  const seen = new Set();
  const macroExpansionCache = new Map();
  const protocol = normalizeRuntimeProtocol(defaultProtocol);

  for (const declaration of declarations) {
    const pvName = normalizeDatabaseMonitorPvName(
      expandDatabaseMonitorValue(
        declaration.name,
        macroDefinitions,
        macroExpansionCache,
        [],
      ),
      declaration.name,
    );
    if (!pvName || seen.has(pvName)) {
      continue;
    }

    seen.add(pvName);
    definitions.push({
      pvName,
      protocol,
      pvRequest: "",
      tocRecordName: declaration.name,
      recordType: declaration.recordType,
    });
  }

  return definitions;
}

function extractDatabaseMonitorDeclarationsFallback(text) {
  const declarations = [];

  for (const match of String(text || "").matchAll(
    /\brecord\s*\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)+)"\s*\)/g,
  )) {
    declarations.push({
      recordType: String(match[1] || "").trim(),
      name: String(match[2] || "").trim(),
    });
  }

  return declarations;
}

function createDatabaseMonitorMacroDefinitions(assignments) {
  const macroDefinitions = new Map();
  if (!assignments || typeof assignments.entries !== "function") {
    return macroDefinitions;
  }

  for (const [macroName, assignment] of assignments.entries()) {
    if (!assignment?.hasAssignment) {
      continue;
    }

    macroDefinitions.set(macroName, {
      name: macroName,
      value: assignment.value || "",
    });
  }

  return macroDefinitions;
}

function expandDatabaseMonitorValue(text, macroDefinitions, cache, stack) {
  const sourceText = String(text || "");
  let expandedText = "";
  let cursor = 0;

  for (const match of sourceText.matchAll(/\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}/g)) {
    const matchText = match[0];
    const matchIndex = match.index ?? 0;
    expandedText += sourceText.slice(cursor, matchIndex);

    const macroName = match[1] || match[3];
    const defaultValue = match[2];
    expandedText += resolveDatabaseMonitorMacro(
      macroName,
      defaultValue,
      matchText,
      macroDefinitions,
      cache,
      stack,
    );
    cursor = matchIndex + matchText.length;
  }

  expandedText += sourceText.slice(cursor);
  return expandedText;
}

function resolveDatabaseMonitorMacro(
  macroName,
  defaultValue,
  originalText,
  macroDefinitions,
  cache,
  stack,
) {
  const cacheKey = `${macroName}\u0000${defaultValue !== undefined ? defaultValue : ""}\u0000${originalText}`;
  if (cache?.has(cacheKey)) {
    return cache.get(cacheKey);
  }

  if (stack.includes(macroName)) {
    return originalText;
  }

  const definition = macroDefinitions.get(macroName);
  let resolvedValue;

  if (definition) {
    resolvedValue = expandDatabaseMonitorValue(
      definition.value,
      macroDefinitions,
      cache,
      [...stack, macroName],
    );
  } else if (defaultValue !== undefined) {
    resolvedValue = expandDatabaseMonitorValue(
      defaultValue,
      macroDefinitions,
      cache,
      stack,
    );
  } else {
    resolvedValue = originalText;
  }

  cache?.set(cacheKey, resolvedValue);
  return resolvedValue;
}

function normalizeDatabaseMonitorPvName(expandedPvName, fallbackPvName) {
  const normalizedExpandedPvName = String(expandedPvName || "").trim();
  if (!normalizedExpandedPvName || /\s/.test(normalizedExpandedPvName)) {
    return String(fallbackPvName || "").trim();
  }

  return normalizedExpandedPvName;
}

function analyzeStrictMonitorText(text, defaultProtocol) {
  const lines = String(text || "").split(/\r?\n/);
  const definitions = [];
  const diagnostics = [];
  const lineReferences = Array.from({ length: lines.length }, () => undefined);
  const seen = new Set();
  const macroDefinitions = new Map();
  const parsedLines = [];
  const protocol = normalizeRuntimeProtocol(defaultProtocol);

  for (let lineNumber = 0; lineNumber < lines.length; lineNumber += 1) {
    const lineText = lines[lineNumber];
    const trimmedLine = lineText.trim();
    const startCharacter = lineText.indexOf(trimmedLine);
    const endCharacter = startCharacter + trimmedLine.length;

    if (!trimmedLine) {
      parsedLines.push({ type: "blank" });
      continue;
    }

    if (trimmedLine.startsWith("#")) {
      parsedLines.push({ type: "comment" });
      continue;
    }

    const macroMatch = trimmedLine.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\S+)\s*$/);
    if (macroMatch) {
      const macroName = macroMatch[1];
      if (macroDefinitions.has(macroName)) {
        diagnostics.push({
          lineNumber,
          startCharacter,
          endCharacter,
          message: `Duplicate monitor macro "${macroName}".`,
        });
      } else {
        macroDefinitions.set(macroName, {
          name: macroName,
          value: macroMatch[2],
          lineNumber,
          startCharacter: startCharacter + trimmedLine.indexOf(macroName),
          endCharacter: startCharacter + trimmedLine.indexOf(macroName) + macroName.length,
        });
      }
      parsedLines.push({ type: "macro" });
      continue;
    }

    if (trimmedLine.includes("=")) {
      diagnostics.push({
        lineNumber,
        startCharacter,
        endCharacter,
        message: 'Monitor macro definitions must be exactly "NAME = value" with no extra text.',
      });
      parsedLines.push({ type: "invalid" });
      continue;
    }

    if (/\s/.test(trimmedLine)) {
      diagnostics.push({
        lineNumber,
        startCharacter,
        endCharacter,
        message: "Monitor record lines must contain exactly one record name with no extra text.",
      });
      parsedLines.push({ type: "invalid" });
      continue;
    }

    parsedLines.push({
      type: "record",
      rawPvName: trimmedLine,
      lineNumber,
      startCharacter,
      endCharacter,
    });
  }

  const macroExpansionCache = new Map();
  for (const parsedLine of parsedLines) {
    if (parsedLine.type !== "record") {
      continue;
    }

    const expansion = expandStrictMonitorValue(
      parsedLine.rawPvName,
      macroDefinitions,
      macroExpansionCache,
      [],
    );
    if (expansion.errors.length > 0) {
      for (const error of expansion.errors) {
        diagnostics.push({
          lineNumber: parsedLine.lineNumber,
          startCharacter: parsedLine.startCharacter,
          endCharacter: parsedLine.endCharacter,
          message: error,
        });
      }
      continue;
    }

    const expandedPvName = String(expansion.value || "").trim();
    if (!expandedPvName || /\s/.test(expandedPvName)) {
      diagnostics.push({
        lineNumber: parsedLine.lineNumber,
        startCharacter: parsedLine.startCharacter,
        endCharacter: parsedLine.endCharacter,
        message: "Monitor record name resolves to invalid text.",
      });
      continue;
    }

    const reference = {
      pvName: expandedPvName,
      protocol,
      pvRequest: "",
      startCharacter: parsedLine.startCharacter,
      endCharacter: parsedLine.endCharacter,
    };
    lineReferences[parsedLine.lineNumber] = reference;

    const key = `${reference.protocol}:${reference.pvName}:`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    definitions.push({
      pvName: reference.pvName,
      protocol: reference.protocol,
      pvRequest: "",
    });
  }

  return {
    definitions,
    diagnostics,
    lineReferences,
  };
}

function extractLineMonitorDefinitions(text, defaultProtocol) {
  const definitions = [];
  const seen = new Set();

  for (const rawLine of String(text || "").split(/\r?\n/)) {
    const reference = parseLineMonitorReference(rawLine, defaultProtocol);
    if (!reference) {
      continue;
    }

    const key = `${reference.protocol}:${reference.pvName}:${reference.pvRequest}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    definitions.push({
      pvName: reference.pvName,
      protocol: reference.protocol,
      pvRequest: reference.pvRequest,
    });
  }

  return definitions;
}

function expandStrictMonitorValue(text, macroDefinitions, cache, stack) {
  const sourceText = String(text || "");
  let expandedText = "";
  const errors = [];
  let cursor = 0;

  for (const match of sourceText.matchAll(/\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}/g)) {
    const matchText = match[0];
    const matchIndex = match.index ?? 0;
    expandedText += sourceText.slice(cursor, matchIndex);

    const macroName = match[1] || match[3];
    const defaultValue = match[2];
    const resolved = resolveStrictMonitorMacro(
      macroName,
      macroDefinitions,
      cache,
      stack,
      defaultValue,
    );
    expandedText += resolved.value;
    errors.push(...resolved.errors);
    cursor = matchIndex + matchText.length;
  }

  expandedText += sourceText.slice(cursor);
  return {
    value: expandedText,
    errors,
  };
}

function resolveStrictMonitorMacro(
  macroName,
  macroDefinitions,
  cache,
  stack,
  defaultValue,
) {
  if (cache?.has(macroName)) {
    return cache.get(macroName);
  }

  if (stack.includes(macroName)) {
    return {
      value: "",
      errors: [
        `Circular monitor macro reference: ${[...stack, macroName].join(" -> ")}.`,
      ],
    };
  }

  const definition = macroDefinitions.get(macroName);
  if (!definition) {
    if (defaultValue !== undefined) {
      return expandStrictMonitorValue(
        defaultValue,
        macroDefinitions,
        cache,
        stack,
      );
    }

    return {
      value: "",
      errors: [`Undefined monitor macro "${macroName}".`],
    };
  }

  const resolved = expandStrictMonitorValue(
    definition.value,
    macroDefinitions,
    cache,
    [...stack, macroName],
  );
  cache?.set(macroName, resolved);
  return resolved;
}

function parseLineMonitorReference(lineText, defaultProtocol) {
  const normalizedLine = String(lineText || "");
  const trimmedLine = normalizedLine.trim();
  if (!trimmedLine || trimmedLine.startsWith("#") || trimmedLine.startsWith("//")) {
    return undefined;
  }

  let protocol = normalizeRuntimeProtocol(defaultProtocol);
  let payload = trimmedLine;
  let payloadStart = normalizedLine.indexOf(trimmedLine);

  const protocolMatch = payload.match(/^(ca|pva)\s+/i);
  if (protocolMatch) {
    protocol = protocolMatch[1].toLowerCase();
    payload = payload.slice(protocolMatch[0].length).trim();
    payloadStart += protocolMatch[0].length;
    while (/\s/.test(normalizedLine[payloadStart] || "")) {
      payloadStart += 1;
    }
  }

  if (!payload) {
    return undefined;
  }

  const pvNameMatch = payload.match(/^\S+/);
  if (!pvNameMatch) {
    return undefined;
  }

  const pvName = pvNameMatch[0];
  const pvNameStart = payloadStart;
  const pvNameEnd = pvNameStart + pvName.length;
  const pvRequest = protocol === "pva"
    ? payload.slice(pvName.length).trim()
    : "";

  return {
    pvName,
    protocol,
    pvRequest,
    startCharacter: pvNameStart,
    endCharacter: pvNameEnd,
  };
}

function normalizeRuntimeProtocol(protocol) {
  return String(protocol || DEFAULT_PROTOCOL).toLowerCase() === "pva"
    ? "pva"
    : "ca";
}

function createMonitorDiagnostic(diagnostic) {
  const range = new vscode.Range(
    new vscode.Position(
      diagnostic.lineNumber,
      Math.max(0, diagnostic.startCharacter || 0),
    ),
    new vscode.Position(
      diagnostic.lineNumber,
      Math.max(
        Math.max(0, diagnostic.startCharacter || 0) + 1,
        diagnostic.endCharacter || 0,
      ),
    ),
  );
  const vscodeDiagnostic = new vscode.Diagnostic(
    range,
    diagnostic.message,
    vscode.DiagnosticSeverity.Error,
  );
  vscodeDiagnostic.source = "vscode-epics";
  return vscodeDiagnostic;
}

function formatRuntimeHoverValue(value, maxLength = RUNTIME_HOVER_VALUE_MAX_LENGTH) {
  return truncateText(formatRuntimeDisplayValue(value), maxLength);
}

function extractPvaRuntimeValue(pvaData) {
  if (pvaData === undefined || pvaData === null) {
    return undefined;
  }

  if (typeof pvaData !== "object" || Array.isArray(pvaData)) {
    return pvaData;
  }

  if (Object.prototype.hasOwnProperty.call(pvaData, "value")) {
    return pvaData.value;
  }

  return undefined;
}

function hasPvaRuntimeDataWithoutValue(pvaData) {
  return (
    Boolean(pvaData) &&
    typeof pvaData === "object" &&
    !Array.isArray(pvaData) &&
    !Object.prototype.hasOwnProperty.call(pvaData, "value") &&
    Object.keys(pvaData).length > 0
  );
}

function getPvaRuntimeDisplayValue(pvaData, cachedEnumChoices) {
  const value = resolvePvaRuntimeValue(pvaData, cachedEnumChoices);
  if (value !== undefined) {
    return value;
  }

  return hasPvaRuntimeDataWithoutValue(pvaData)
    ? PVA_HAS_DATA_WITHOUT_VALUE_TEXT
    : undefined;
}

function resolvePvaRuntimeValue(pvaData, cachedEnumChoices) {
  const value = extractPvaRuntimeValue(pvaData);
  if (value === undefined) {
    return undefined;
  }

  if (!isPvaEnumIndexValue(value)) {
    return value;
  }

  return {
    index: getPvaEnumIndex(value),
    choices: getPvaEnumChoices(value, cachedEnumChoices),
  };
}

function isPvaEnumIndexValue(value) {
  return (
    Boolean(value) &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    Number.isInteger(Number(value.index))
  );
}

function isPvaEnumLikeValue(value) {
  return isPvaEnumIndexValue(value) && Array.isArray(value.choices);
}

function getPvaEnumIndex(value) {
  return Number(value?.index);
}

function getPvaEnumChoices(value, fallbackChoices) {
  if (Array.isArray(value?.choices)) {
    return value.choices.map((choice) => String(choice ?? ""));
  }

  if (Array.isArray(fallbackChoices)) {
    return fallbackChoices.map((choice) => String(choice ?? ""));
  }

  return [];
}

function formatPvaEnumLikeValue(value) {
  const index = getPvaEnumIndex(value);
  const choices = getPvaEnumChoices(value);
  const choiceLabel = choices[index];
  if (choiceLabel === undefined) {
    return `[${index}] Illegal_Value`;
  }

  if (choiceLabel === "") {
    return `[${index}] ""`;
  }

  return `[${index}] ${choiceLabel}`;
}

function updatePvaEnumChoicesCache(entry, pvaData) {
  if (!entry) {
    return;
  }

  const value = extractPvaRuntimeValue(pvaData);
  if (isPvaEnumIndexValue(value) && Array.isArray(value?.choices)) {
    entry.pvaEnumChoices = getPvaEnumChoices(value);
  }
}

function formatRuntimeDisplayValue(value) {
  if (value === undefined) {
    return "";
  }

  if (isPvaEnumLikeValue(value)) {
    return formatPvaEnumLikeValue(value);
  }

  if (typeof value === "string") {
    return value;
  }

  try {
    return JSON.stringify(value);
  } catch (error) {
    return String(value);
  }
}

function getRuntimePutValueKind(dbrType, runtimeLibrary) {
  if (dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_STRING) {
    return "string";
  }

  if (
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_INT ||
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_FLOAT ||
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_ENUM ||
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_CHAR ||
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_LONG ||
    dbrType === runtimeLibrary.Channel_DBR_TYPES.DBR_DOUBLE
  ) {
    return "number";
  }

  return undefined;
}

function validateRuntimePutInput(value, putSupport) {
  return parseRuntimePutInput(value, putSupport).error;
}

function parseRuntimePutInput(value, putSupport) {
  const valueKind = putSupport?.valueKind || putSupport;
  if (valueKind === "string") {
    return {
      value: String(value ?? ""),
    };
  }

  if (valueKind === "pva-enum" || valueKind === "ca-enum") {
    return parsePvaEnumRuntimePutInput(value, putSupport?.enumChoices || []);
  }

  const trimmedValue = String(value ?? "").trim();
  if (!trimmedValue) {
    return {
      error: "A numeric value is required.",
    };
  }

  const parsedNumber = Number(trimmedValue);
  if (!Number.isFinite(parsedNumber)) {
    return {
      error: `Invalid numeric value: ${value}`,
    };
  }

  return {
    value: parsedNumber,
  };
}

function parsePvaEnumRuntimePutInput(value, enumChoices) {
  const trimmedValue = String(value ?? "").trim();
  if (!trimmedValue) {
    return {
      error: "An enum index or choice name is required.",
    };
  }

  const parsedNumber = Number(trimmedValue);
  if (Number.isFinite(parsedNumber)) {
    if (!Number.isInteger(parsedNumber)) {
      return {
        error: "Enter a whole-number enum index.",
      };
    }

    return {
      value: parsedNumber,
    };
  }

  const normalizedChoice = stripOptionalWrappingQuotes(trimmedValue);
  const choiceIndex = enumChoices.findIndex(
    (choice) => choice === normalizedChoice,
  );
  if (choiceIndex >= 0) {
    return {
      value: choiceIndex,
    };
  }

  return {
    error: enumChoices.length
      ? `Enter a valid enum index or one of: ${enumChoices.join(", ")}.`
      : "Enter a valid enum index.",
  };
}

async function showRuntimePutInput(entry, putSupport) {
  if (putSupport?.valueKind === "pva-enum" || putSupport?.valueKind === "ca-enum") {
    return showRuntimeEnumPutInput(entry, putSupport);
  }

  return vscode.window.showInputBox({
    prompt: `Put ${entry.pvName} (${putSupport.typeLabel})`,
    value: putSupport.initialValue,
    ignoreFocusOut: false,
    validateInput: (rawValue) =>
      validateRuntimePutInput(rawValue, putSupport),
  });
}

function showRuntimeEnumPutInput(entry, putSupport) {
  return new Promise((resolve) => {
    const quickPick = vscode.window.createQuickPick();
    let settled = false;
    const enumChoices = Array.isArray(putSupport?.enumChoices)
      ? putSupport.enumChoices
      : [];
    const currentIndex = Number(putSupport?.initialValue);

    quickPick.title = `Put ${entry.pvName} (${putSupport.typeLabel})`;
    quickPick.placeholder =
      "Select a choice or type an index / choice name and press Enter";
    quickPick.ignoreFocusOut = false;
    quickPick.matchOnDescription = true;
    quickPick.matchOnDetail = true;
    const items = enumChoices.map((choice, index) => {
      const choiceLabel = choice === "" ? '""' : choice;
      return {
        label: `[${index}] ${choiceLabel}`,
        description: index === currentIndex ? "Current" : undefined,
        detail: `Put index ${index}`,
        inputValue: String(index),
      };
    });
    quickPick.items = items;
    const currentItem = items.find((item) => item.inputValue === putSupport?.initialValue);
    if (currentItem) {
      quickPick.activeItems = [currentItem];
    }

    const finish = (value) => {
      if (settled) {
        return;
      }

      settled = true;
      quickPick.dispose();
      resolve(value);
    };

    quickPick.onDidAccept(() => {
      const [selectedItem] = quickPick.selectedItems || [];
      if (selectedItem?.inputValue !== undefined) {
        finish(selectedItem.inputValue);
        return;
      }

      const rawValue = String(quickPick.value || "").trim();
      finish(rawValue || undefined);
    });
    quickPick.onDidHide(() => {
      finish(undefined);
    });
    quickPick.show();
  });
}

function stripOptionalWrappingQuotes(value) {
  const text = String(value ?? "");
  if (
    text.length >= 2 &&
    ((text.startsWith('"') && text.endsWith('"')) ||
      (text.startsWith("'") && text.endsWith("'")))
  ) {
    return text.slice(1, -1);
  }

  return text;
}

function isSuccessfulRuntimePutResult(result, runtimeLibrary, protocol = "ca") {
  if (protocol === "pva") {
    return (
      Boolean(result) &&
      (result.type === runtimeLibrary.PVA_STATUS_TYPE.OK ||
        result.type === runtimeLibrary.PVA_STATUS_TYPE.OKOK)
    );
  }

  return result === runtimeLibrary.ECA_VALUES.ECA_NORMAL;
}

function getRuntimePutFailureMessage(result, runtimeLibrary, protocol = "ca") {
  if (result === undefined) {
    return "The put operation did not complete.";
  }

  if (protocol === "pva") {
    const statusType = runtimeLibrary.PVA_STATUS_TYPE?.[result?.type];
    if (result?.message && statusType) {
      return `${statusType}: ${result.message}`;
    }
    if (result?.message) {
      return result.message;
    }
    if (statusType) {
      return statusType;
    }

    return "EPICS PVA put failed.";
  }

  const ecaName = runtimeLibrary.ECA_VALUES?.[result];
  if (ecaName) {
    return `${ecaName} (${result})`;
  }

  return `EPICS CA put failed with status ${result}.`;
}

function escapeInlineCode(value) {
  return String(value ?? "").replace(/`/g, "\\`");
}

function escapeMarkdownText(value) {
  return String(value ?? "").replace(/([\\`*_{}\[\]()#+\-.!])/g, "\\$1");
}

module.exports = {
  registerRuntimeMonitor,
};
