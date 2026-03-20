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
const DEFAULT_PROTOCOL = "ca";
const DEFAULT_CHANNEL_CREATION_TIMEOUT_SECONDS = 0;
const DEFAULT_MONITOR_SUBSCRIBE_TIMEOUT_SECONDS = 0;
const DEFAULT_LOG_LEVEL = "ERROR";
const DATABASE_RUNTIME_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const LINE_RUNTIME_EXTENSIONS = new Set([".monitor", ".txt"]);
const STATUS_BAR_PRIORITY = 110;
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

    const isRunning = this.isDocumentMonitoringRunning(document);
    this.statusBarItem.text = isRunning ? "$(primitive-square) EPICS" : "$(play) EPICS";
    this.statusBarItem.command = isRunning
      ? STOP_ACTIVE_FILE_RUNTIME_COMMAND
      : START_ACTIVE_FILE_RUNTIME_COMMAND;
    this.statusBarItem.tooltip = isRunning
      ? `Stop EPICS runtime monitoring for ${path.basename(document.uri.fsPath)}`
      : `Start EPICS runtime monitoring for ${path.basename(document.uri.fsPath)}`;
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

    return markdown;
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

    const analysis = analyzeRuntimeDocument(
      document,
      this.getDefaultProtocol(),
      this.getDatabaseRuntimeHelpers(),
    );
    if (analysis.diagnostics?.length) {
      vscode.window.showErrorMessage(
        `Cannot start EPICS runtime for ${path.basename(document.uri.fsPath)} until the .monitor file errors are fixed.`,
      );
      return;
    }

    const definitions = analysis.definitions;
    if (!definitions.length) {
      vscode.window.showWarningMessage(
        `No EPICS monitor targets were found in ${path.basename(document.uri.fsPath)}.`,
      );
      return;
    }

    const sourceUri = document.uri.toString();
    const sourceLabel = path.basename(document.uri.fsPath);
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
      entry.valueText = formatRuntimeValue(monitor.getPvaData());
      return;
    }

    entry.valueText = formatRuntimeValue(monitor.getChannel().getDbrData()?.value);
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

  getDatabaseRuntimeHelpers() {
    return {
      extractDatabaseTocMacroAssignments: this.extractDatabaseTocMacroAssignments,
      extractRecordDeclarations: this.extractRecordDeclarations,
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
    return truncateText(value, RUNTIME_VALUE_DISPLAY_MAX_LENGTH);
  }

  try {
    return truncateText(JSON.stringify(value), RUNTIME_VALUE_DISPLAY_MAX_LENGTH);
  } catch (error) {
    return truncateText(String(value), RUNTIME_VALUE_DISPLAY_MAX_LENGTH);
  }
}

function getMonitorHoverValue(entry) {
  if (!entry) {
    return undefined;
  }

  try {
    if (entry.protocol === "pva") {
      return entry.monitor?.getPvaData();
    }

    return entry.monitor?.getChannel().getDbrData()?.value;
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

  if (typeof value === "string") {
    return truncateText(value, DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH);
  }

  try {
    return truncateText(
      JSON.stringify(value),
      DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH,
    );
  } catch (error) {
    return truncateText(String(value), DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH);
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
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  const extension = path.extname(document.uri.fsPath).toLowerCase();
  return (
    isDatabaseRuntimeDocument(document) ||
    document.languageId === "monitor" ||
    document.languageId === "plaintext" ||
    LINE_RUNTIME_EXTENSIONS.has(extension)
  );
}

function isStrictMonitorDocument(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  return (
    document.languageId === "monitor" ||
    path.extname(document.uri.fsPath).toLowerCase() === ".monitor"
  );
}

function isDatabaseRuntimeDocument(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  return (
    document.languageId === "database" ||
    DATABASE_RUNTIME_EXTENSIONS.has(path.extname(document.uri.fsPath).toLowerCase())
  );
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

function extractDatabaseMonitorDefinitions(text, databaseHelpers = {}) {
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
      protocol: "ca",
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
  return truncateText(String(value ?? ""), maxLength);
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
