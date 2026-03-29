const fs = require("fs");
const os = require("os");
const path = require("path");
const vscode = require("vscode");
const { formatMonitorText } = require("./formatters");

const ADD_RUNTIME_MONITOR_COMMAND = "vscode-epics.addRuntimeMonitor";
const REMOVE_RUNTIME_MONITOR_COMMAND = "vscode-epics.removeRuntimeMonitor";
const CLEAR_RUNTIME_MONITORS_COMMAND = "vscode-epics.clearRuntimeMonitors";
const RESTART_RUNTIME_CONTEXT_COMMAND = "vscode-epics.restartRuntimeContext";
const STOP_RUNTIME_CONTEXT_COMMAND = "vscode-epics.stopRuntimeContext";
const START_ACTIVE_FILE_RUNTIME_COMMAND = "vscode-epics.startActiveFileRuntimeContext";
const STOP_ACTIVE_FILE_RUNTIME_COMMAND = "vscode-epics.stopActiveFileRuntimeContext";
const START_DATABASE_MONITOR_CHANNELS_COMMAND =
  "vscode-epics.startDatabaseMonitorChannels";
const STOP_DATABASE_MONITOR_CHANNELS_COMMAND =
  "vscode-epics.stopDatabaseMonitorChannels";
const PUT_RUNTIME_VALUE_COMMAND = "vscode-epics.putRuntimeValue";
const OPEN_PROJECT_RUNTIME_CONFIGURATION_COMMAND =
  "vscode-epics.openProjectRuntimeConfiguration";
const OPEN_RUNTIME_WIDGET_COMMAND = "vscode-epics.openRuntimeWidget";
const OPEN_PROBE_WIDGET_COMMAND = "vscode-epics.openProbeWidget";
const OPEN_PVLIST_WIDGET_COMMAND = "vscode-epics.openPvlistWidget";
const OPEN_MONITOR_WIDGET_COMMAND = "vscode-epics.openMonitorWidget";
const SET_IOC_SHELL_TERMINAL_COMMAND = "vscode-epics.setIocShellTerminal";
const RUN_IOC_SHELL_COMMAND = "vscode-epics.runIocShellCommand";
const RUN_IOC_SHELL_DBL_COMMAND = "vscode-epics.runIocShellDbl";
const RUN_IOC_SHELL_DBPR_CURRENT_RECORD_COMMAND =
  "vscode-epics.runIocShellDbprCurrentRecord";
const START_ACTIVE_STARTUP_IOC_COMMAND = "vscode-epics.startActiveStartupIoc";
const STOP_ACTIVE_STARTUP_IOC_COMMAND = "vscode-epics.stopActiveStartupIoc";
const MANAGE_PROJECT_STARTUP_IOC_COMMAND = "vscode-epics.manageProjectStartupIoc";
const START_PROJECT_STARTUP_IOC_COMMAND = "vscode-epics.startProjectStartupIoc";
const STOP_PROJECT_STARTUP_IOC_COMMAND = "vscode-epics.stopProjectStartupIoc";
const SHOW_ACTIVE_STARTUP_IOC_TERMINAL_COMMAND =
  "vscode-epics.showActiveStartupIocTerminal";
const OPEN_ACTIVE_STARTUP_IOC_COMMANDS_COMMAND =
  "vscode-epics.openActiveStartupIocCommands";
const OPEN_ACTIVE_STARTUP_IOC_VARIABLES_COMMAND =
  "vscode-epics.openActiveStartupIocVariables";
const OPEN_ACTIVE_STARTUP_IOC_ENVIRONMENT_COMMAND =
  "vscode-epics.openActiveStartupIocEnvironment";
const DEFAULT_PROTOCOL = "ca";
const DEFAULT_CHANNEL_CREATION_TIMEOUT_SECONDS = 0;
const DEFAULT_MONITOR_SUBSCRIBE_TIMEOUT_SECONDS = 0;
const DEFAULT_LOG_LEVEL = "ERROR";
const DATABASE_RUNTIME_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const LINE_RUNTIME_EXTENSIONS = new Set([".pvlist", ".txt"]);
const PROBE_RUNTIME_EXTENSIONS = new Set([".probe"]);
const STATUS_BAR_PRIORITY = 110;
const MOUSE_DOUBLE_CLICK_INTERVAL_MS = 400;
const PROJECT_RUNTIME_CONFIG_FILE_NAME = ".epics-workbench-config.json";
const PROJECT_RUNTIME_CONFIGURATION_VIEW_TYPE =
  "epicsWorkbench.projectRuntimeConfiguration";
const PROBE_CUSTOM_EDITOR_VIEW_TYPE = "epicsWorkbench.probeEditor";
const RUNTIME_WIDGET_VIEW_TYPE = "epicsWorkbench.runtimeWidget";
const PROBE_WIDGET_VIEW_TYPE = "epicsWorkbench.probeWidget";
const PVLIST_WIDGET_VIEW_TYPE = "epicsWorkbench.pvlistWidget";
const MONITOR_WIDGET_VIEW_TYPE = "epicsWorkbench.monitorWidget";
const IOC_RUNTIME_COMMANDS_VIEW_TYPE = "epicsWorkbench.iocRuntimeCommands";
const IOC_RUNTIME_VARIABLES_VIEW_TYPE = "epicsWorkbench.iocRuntimeVariables";
const IOC_RUNTIME_ENVIRONMENT_VIEW_TYPE = "epicsWorkbench.iocRuntimeEnvironment";
const PROJECT_RUNTIME_CONFIGURATION_PROTOCOL_VALUES = ["ca", "pva"];
const PROJECT_RUNTIME_CONFIGURATION_AUTO_ADDR_LIST_VALUES = ["Yes", "No"];
const PVA_HAS_DATA_WITHOUT_VALUE_TEXT = "Has data, but no value";
const RUNTIME_VALUE_DISPLAY_MAX_LENGTH = 120;
const RUNTIME_HOVER_VALUE_MAX_LENGTH = 240;
const DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH = 18;
const PROBE_FIELD_CONNECT_BATCH_SIZE = 8;
const PROBE_OVERLAY_MAX_LINES = 160;
const DEFAULT_MONITOR_WIDGET_BUFFER_SIZE = 500;
const EPICS_CA_EPOCH_OFFSET_SECONDS = 631152000;
const STARTUP_TERMINAL_OUTPUT_MAX_LENGTH = 250000;
const IOC_RUNTIME_COMMANDS_FETCH_TIMEOUT_MS = 5000;
const STARTUP_RUNNING_WATERMARK_LINE_INTERVAL = 6;
const CONTEXT_INITIALIZATION_CANCELLED_MESSAGE =
  "EPICS runtime context initialization was cancelled.";
const ACTIVE_DATABASE_HAS_TOC_CONTEXT_KEY =
  "epicsWorkbench.activeDatabaseHasToc";
const ACTIVE_DATABASE_MONITORING_RUNNING_CONTEXT_KEY =
  "epicsWorkbench.activeDatabaseMonitoringRunning";
const ACTIVE_STARTUP_CAN_START_IOC_CONTEXT_KEY =
  "epicsWorkbench.activeStartupCanStartIoc";
const ACTIVE_STARTUP_CAN_STOP_IOC_CONTEXT_KEY =
  "epicsWorkbench.activeStartupCanStopIoc";
const STARTUP_IOC_PICKER_START_BUTTON = {
  iconPath: new vscode.ThemeIcon("play"),
  tooltip: "Start",
};
const STARTUP_IOC_PICKER_STOP_BUTTON = {
  iconPath: new vscode.ThemeIcon("stop-circle"),
  tooltip: "Stop",
};
const STARTUP_IOC_PICKER_BRING_TO_FRONT_BUTTON = {
  iconPath: new vscode.ThemeIcon("arrow-up"),
  tooltip: "Bring to Front",
};
const STARTUP_IOC_PICKER_OPEN_FILE_BUTTON = {
  iconPath: new vscode.ThemeIcon("go-to-file"),
  tooltip: "Open File",
};
const EPICS_PROJECT_MARKER_SEGMENTS = [
  ["Makefile"],
  ["configure", "RELEASE"],
  ["configure", "RULES_TOP"],
];
function registerRuntimeMonitor(extensionContext, databaseHelpers = {}) {
  const controller = new EpicsRuntimeMonitorController(databaseHelpers);
  const probeCustomEditorProvider = new EpicsProbeCustomEditorProvider(controller);
  const diagnostics = vscode.languages.createDiagnosticCollection("vscode-epics-pvlist");
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
  const startupRunningDecorationType = vscode.window.createTextEditorDecorationType({
    isWholeLine: true,
    after: {
      color: "transparent",
      margin: "0.9em 0 0.9em 5ch",
      fontStyle: "italic",
      fontWeight: "700",
      textDecoration:
        "none; display: block; font-size: 3.2em; letter-spacing: 0.08em; -webkit-text-stroke: 1.4px rgba(220, 38, 38, 0.7);",
    },
  });
  const probeDecorationTypes = Array.from(
    { length: PROBE_OVERLAY_MAX_LINES },
    () =>
      vscode.window.createTextEditorDecorationType({
        after: {
          color: new vscode.ThemeColor("descriptionForeground"),
        },
      }),
  );

  controller.attachStatusBar(statusBarItem);
  controller.attachDiagnosticsCollection(diagnostics);
  controller.attachHoverDecorationType(hoverDecorationType);
  controller.attachDatabaseTocValueDecorationType(databaseTocValueDecorationType);
  controller.attachStartupRunningDecorationType(startupRunningDecorationType);
  controller.attachProbeDecorationTypes(probeDecorationTypes);

  extensionContext.subscriptions.push(
    controller,
    probeCustomEditorProvider,
    diagnostics,
    statusBarItem,
    hoverDecorationType,
    databaseTocValueDecorationType,
    startupRunningDecorationType,
    ...probeDecorationTypes,
  );
  extensionContext.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      PROBE_CUSTOM_EDITOR_VIEW_TYPE,
      probeCustomEditorProvider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
        },
        supportsMultipleEditorsPerDocument: true,
      },
    ),
  );
  extensionContext.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: "pvlist" },
      new EpicsMonitorFormattingProvider(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      ADD_RUNTIME_MONITOR_COMMAND,
      async () => {
        await controller.openRuntimeWidget();
        await controller.addMonitorInteractive();
      },
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
      START_DATABASE_MONITOR_CHANNELS_COMMAND,
      async () => controller.startDatabaseMonitorChannels(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      STOP_DATABASE_MONITOR_CHANNELS_COMMAND,
      async () => controller.stopDatabaseMonitorChannels(),
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
    vscode.commands.registerCommand(
      OPEN_RUNTIME_WIDGET_COMMAND,
      async () => controller.openRuntimeWidget(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_PROBE_WIDGET_COMMAND,
      async (options) => controller.openProbeWidget(options),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_PVLIST_WIDGET_COMMAND,
      async (options) => controller.openPvlistWidget(options),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_MONITOR_WIDGET_COMMAND,
      async (options) => controller.openMonitorWidget(options),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      SET_IOC_SHELL_TERMINAL_COMMAND,
      async () => controller.setIocShellTerminalInteractive(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      RUN_IOC_SHELL_COMMAND,
      async () => controller.runIocShellCommandInteractive(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      RUN_IOC_SHELL_DBL_COMMAND,
      async () => controller.sendNamedIocShellCommand("dbl"),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      RUN_IOC_SHELL_DBPR_CURRENT_RECORD_COMMAND,
      async () => controller.runDbprForActiveRecord(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      START_ACTIVE_STARTUP_IOC_COMMAND,
      async () => controller.startActiveStartupIoc(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      STOP_ACTIVE_STARTUP_IOC_COMMAND,
      async () => controller.stopActiveStartupIoc(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      MANAGE_PROJECT_STARTUP_IOC_COMMAND,
      async (resourceUri) => controller.showProjectStartupIocPicker(resourceUri),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      START_PROJECT_STARTUP_IOC_COMMAND,
      async (resourceUri) => controller.startProjectStartupIoc(resourceUri),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      STOP_PROJECT_STARTUP_IOC_COMMAND,
      async (resourceUri) => controller.stopProjectStartupIoc(resourceUri),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      SHOW_ACTIVE_STARTUP_IOC_TERMINAL_COMMAND,
      async () => controller.showActiveStartupIocTerminal(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_ACTIVE_STARTUP_IOC_COMMANDS_COMMAND,
      async () => controller.openActiveStartupIocCommands(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_ACTIVE_STARTUP_IOC_VARIABLES_COMMAND,
      async () => controller.openActiveStartupIocVariables(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_ACTIVE_STARTUP_IOC_ENVIRONMENT_COMMAND,
      async () => controller.openActiveStartupIocEnvironment(),
    ),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor((editor) => {
      controller.handleActiveEditorChange(editor);
      controller.refreshMonitorHoverDecorationsForEditor(editor);
      controller.refreshDatabaseTocValueDecorationsForEditor(editor);
      controller.refreshStartupRunningDecorationsForEditor(editor);
      controller.refreshProbeDecorationsForEditor(editor);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidChangeActiveTerminal((terminal) => {
      controller.handleActiveTerminalChanged(terminal);
    }),
  );
  extensionContext.subscriptions.push(
    vscode.window.onDidCloseTerminal((terminal) => {
      controller.handleTerminalClosed(terminal);
    }),
  );
  if (typeof vscode.window.onDidStartTerminalShellExecution === "function") {
    extensionContext.subscriptions.push(
      vscode.window.onDidStartTerminalShellExecution((event) => {
        controller.handleTerminalShellExecutionStarted(event);
      }),
    );
  }
  if (typeof vscode.window.onDidChangeTerminalShellIntegration === "function") {
    extensionContext.subscriptions.push(
      vscode.window.onDidChangeTerminalShellIntegration((event) => {
        controller.handleTerminalShellIntegrationChanged(event);
      }),
    );
  }
  if (typeof vscode.window.onDidEndTerminalShellExecution === "function") {
    extensionContext.subscriptions.push(
      vscode.window.onDidEndTerminalShellExecution((event) => {
        controller.handleTerminalShellExecutionEnded(event);
      }),
    );
  }
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
      controller.refreshStartupRunningDecorationsForDocument(event.document);
      controller.refreshProbeDecorationsForDocument(event.document);
      controller.refreshProbePanels();
    }),
  );
  extensionContext.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((document) => {
      controller.refreshMonitorDiagnosticsForDocument(document);
      controller.refreshMonitorHoverDecorationsForDocument(document);
      controller.refreshDatabaseTocValueDecorationsForDocument(document);
      controller.refreshStartupRunningDecorationsForDocument(document);
      controller.refreshProbeDecorationsForDocument(document);
      controller.refreshProbePanels();
    }),
  );
  extensionContext.subscriptions.push(
    vscode.workspace.onDidCloseTextDocument((document) => {
      controller.clearMonitorDiagnosticsForDocument(document);
      controller.clearMonitorHoverDecorationsForDocument(document);
      controller.clearDatabaseTocValueDecorationsForDocument(document);
      controller.clearStartupRunningDecorationsForDocument(document);
      controller.clearProbeDecorationsForDocument(document);
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
  controller.updateDatabaseEditorContextKeys(vscode.window.activeTextEditor);
  for (const document of vscode.workspace.textDocuments) {
    controller.refreshMonitorDiagnosticsForDocument(document);
  }
  controller.refreshVisibleMonitorHoverDecorations();
  return controller;
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
    this.startupRunningDecorationType = undefined;
    this.probeDecorationTypes = [];
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
    this.getFieldNamesForRecordType =
      typeof databaseHelpers.getFieldNamesForRecordType === "function"
        ? databaseHelpers.getFieldNamesForRecordType
        : undefined;
    this.getFieldTypeForRecordType =
      typeof databaseHelpers.getFieldTypeForRecordType === "function"
        ? databaseHelpers.getFieldTypeForRecordType
        : undefined;
    this.probeSessions = new Map();
    this.probeWebviews = new Map();
    this.runtimeWidgetPanel = undefined;
    this.iocRuntimeCommandsPanel = undefined;
    this.iocRuntimeCommandsPanelState = undefined;
    this.iocRuntimeVariablesPanel = undefined;
    this.iocRuntimeVariablesPanelState = undefined;
    this.iocRuntimeEnvironmentPanel = undefined;
    this.iocRuntimeEnvironmentPanelState = undefined;
    this.probeWidgets = new Map();
    this.pvlistWidgets = new Map();
    this.monitorWidgets = new Map();
    this.activeWidgetMenuContext = undefined;
    this.iocShellTerminal = undefined;
    this.lastIocShellCommand = "";
    this.iocStartupTerminalByDocumentPath = new Map();
    this.iocStartupDocumentPathByTerminal = new Map();
    this.terminalActiveExecutionCount = new Map();
    this.terminalLastKnownCwd = new Map();
    this.startupTerminalOutputByTerminal = new Map();
    this.iocRuntimeCommandsByTerminal = new Map();
    this.iocRuntimeCommandHelpByTerminal = new Map();
    this.iocRuntimeVariablesByTerminal = new Map();
    this.iocRuntimeVariablesReloadTimer = undefined;
    this.iocRuntimeEnvironmentByTerminal = new Map();
    this.iocRuntimeEnvironmentReloadTimer = undefined;
  }

  getTreeItem(element) {
    if (element?.type === "context") {
      return this.createContextTreeItem();
    }

    return this.createMonitorTreeItem(element);
  }

  getChildren(element) {
    if (!element) {
      return [this.contextNode, ...this.monitorEntries.filter((entry) => !entry.hidden)];
    }

    return [];
  }

  dispose() {
    this.stopContextInternal();
    this.disposeHoverRefreshTimer();
    this.probeSessions.clear();
    for (const probeWebviewState of this.probeWebviews.values()) {
      for (const panel of probeWebviewState.panels) {
        panel.dispose();
      }
    }
    this.probeWebviews.clear();
    this.runtimeWidgetPanel?.dispose();
    this.runtimeWidgetPanel = undefined;
    this.iocRuntimeCommandsPanel?.dispose();
    this.iocRuntimeCommandsPanel = undefined;
    this.iocRuntimeCommandsPanelState = undefined;
    this.iocRuntimeVariablesPanel?.dispose();
    this.iocRuntimeVariablesPanel = undefined;
    this.iocRuntimeVariablesPanelState = undefined;
    this.iocRuntimeEnvironmentPanel?.dispose();
    this.iocRuntimeEnvironmentPanel = undefined;
    this.iocRuntimeEnvironmentPanelState = undefined;
    for (const widgetState of this.probeWidgets.values()) {
      widgetState.panel.dispose();
    }
    this.probeWidgets.clear();
    for (const widgetState of this.pvlistWidgets.values()) {
      widgetState.panel.dispose();
    }
    this.pvlistWidgets.clear();
    for (const widgetState of this.monitorWidgets.values()) {
      widgetState.panel.dispose();
    }
    this.monitorWidgets.clear();
    this.iocStartupTerminalByDocumentPath.clear();
    this.iocStartupDocumentPathByTerminal.clear();
    this.terminalActiveExecutionCount.clear();
    this.terminalLastKnownCwd.clear();
    this.startupTerminalOutputByTerminal.clear();
    this.iocRuntimeCommandsByTerminal.clear();
    this.iocRuntimeCommandHelpByTerminal.clear();
    this.iocRuntimeVariablesByTerminal.clear();
    this.disposeIocRuntimeVariablesReloadTimer();
    this.iocRuntimeEnvironmentByTerminal.clear();
    this.disposeIocRuntimeEnvironmentReloadTimer();
    this.runtimeConfigurationPanel?.dispose();
    this.runtimeConfigurationPanel = undefined;
    this.statusBarItem?.dispose();
    this._onDidChangeTreeData.dispose();
  }

  handleActiveTerminalChanged(terminal) {
    if (!this.iocShellTerminal && terminal) {
      this.iocShellTerminal = terminal;
      this.refresh(this.contextNode);
    }
    this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
  }

  handleTerminalClosed(terminal) {
    this.untrackStartupIocTerminal(terminal);
    this.terminalActiveExecutionCount.delete(terminal);
    this.terminalLastKnownCwd.delete(terminal);
    this.startupTerminalOutputByTerminal.delete(terminal);
    this.iocRuntimeCommandsByTerminal.delete(terminal);
    this.iocRuntimeCommandHelpByTerminal.delete(terminal);
    this.iocRuntimeVariablesByTerminal.delete(terminal);
    this.iocRuntimeEnvironmentByTerminal.delete(terminal);
    if (this.iocShellTerminal === terminal) {
      this.iocShellTerminal = undefined;
      this.refresh(this.contextNode);
    }
    if (this.iocRuntimeCommandsPanelState?.terminal === terminal) {
      this.iocRuntimeCommandsPanelState = {
        ...this.iocRuntimeCommandsPanelState,
        message: "The tracked IOC terminal has closed.",
        isLoading: false,
      };
      void this.postIocRuntimeCommandsState();
    }
    if (this.iocRuntimeVariablesPanelState?.terminal === terminal) {
      this.iocRuntimeVariablesPanelState = {
        ...this.iocRuntimeVariablesPanelState,
        message: "The tracked IOC terminal has closed.",
        isLoading: false,
      };
      void this.postIocRuntimeVariablesState();
    }
    if (this.iocRuntimeEnvironmentPanelState?.terminal === terminal) {
      this.iocRuntimeEnvironmentPanelState = {
        ...this.iocRuntimeEnvironmentPanelState,
        message: "The tracked IOC terminal has closed.",
        isLoading: false,
      };
      void this.postIocRuntimeEnvironmentState();
    }
    this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
  }

  handleTerminalShellIntegrationChanged(event) {
    const terminal = event?.terminal;
    const cwdUri = event?.shellIntegration?.cwd;
    if (terminal && cwdUri?.scheme === "file") {
      this.terminalLastKnownCwd.set(terminal, normalizeFsPath(cwdUri.fsPath));
    }
  }

  handleTerminalShellExecutionStarted(event) {
    const terminal = event?.terminal;
    if (!terminal) {
      return;
    }

    this.terminalActiveExecutionCount.set(
      terminal,
      this.getTerminalActiveExecutionCount(terminal) + 1,
    );
    if (event?.execution?.cwd?.scheme === "file") {
      this.terminalLastKnownCwd.set(
        terminal,
        normalizeFsPath(event.execution.cwd.fsPath),
      );
    }
    const documentPath = resolveStartupCommandDocumentPath(
      event?.execution?.commandLine?.value,
      event?.execution?.cwd,
    );
    if (!documentPath) {
      return;
    }

    this.trackStartupIocTerminal(documentPath, event.terminal);
    if (this.iocRuntimeCommandsPanelState?.startupDocumentPath === documentPath) {
      this.iocRuntimeCommandsPanelState = {
        ...this.iocRuntimeCommandsPanelState,
        terminal: event.terminal,
        message: "IOC is running.",
        isLoading: false,
      };
      void this.postIocRuntimeCommandsState();
    }
    if (this.iocRuntimeVariablesPanelState?.startupDocumentPath === documentPath) {
      this.iocRuntimeVariablesPanelState = {
        ...this.iocRuntimeVariablesPanelState,
        terminal: event.terminal,
        message: "IOC is running. Loading IOC runtime variables...",
        isLoading: true,
      };
      void this.postIocRuntimeVariablesState();
      this.scheduleIocRuntimeVariablesPanelReload(1500);
    }
    if (this.iocRuntimeEnvironmentPanelState?.startupDocumentPath === documentPath) {
      this.iocRuntimeEnvironmentPanelState = {
        ...this.iocRuntimeEnvironmentPanelState,
        terminal: event.terminal,
        message: "IOC is running. Loading IOC runtime environment...",
        isLoading: true,
      };
      void this.postIocRuntimeEnvironmentState();
      this.scheduleIocRuntimeEnvironmentPanelReload(1500);
    }
    void this.captureStartupTerminalExecutionOutput(event.terminal, event.execution);
  }

  handleTerminalShellExecutionEnded(event) {
    const terminal = event?.terminal;
    if (!terminal) {
      return;
    }

    const nextExecutionCount = Math.max(
      0,
      this.getTerminalActiveExecutionCount(terminal) - 1,
    );
    if (nextExecutionCount > 0) {
      this.terminalActiveExecutionCount.set(terminal, nextExecutionCount);
    } else {
      this.terminalActiveExecutionCount.delete(terminal);
    }
    if (event?.execution?.cwd?.scheme === "file") {
      this.terminalLastKnownCwd.set(
        terminal,
        normalizeFsPath(event.execution.cwd.fsPath),
      );
    }

    const trackedDocumentPath = this.iocStartupDocumentPathByTerminal.get(terminal);
    const endedDocumentPath = resolveStartupCommandDocumentPath(
      event?.execution?.commandLine?.value,
      event?.execution?.cwd,
    );
    if (trackedDocumentPath && (!endedDocumentPath || trackedDocumentPath === endedDocumentPath)) {
      this.untrackStartupIocTerminal(terminal);
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      this.refresh(this.contextNode);
      if (this.iocRuntimeCommandsPanelState?.terminal === terminal) {
        this.iocRuntimeCommandsPanelState = {
          ...this.iocRuntimeCommandsPanelState,
          message: "The tracked IOC process has stopped.",
          isLoading: false,
        };
        void this.postIocRuntimeCommandsState();
      }
      if (this.iocRuntimeVariablesPanelState?.terminal === terminal) {
        this.iocRuntimeVariablesPanelState = {
          ...this.iocRuntimeVariablesPanelState,
          message: "The tracked IOC process has stopped.",
          isLoading: false,
        };
        void this.postIocRuntimeVariablesState();
      }
      if (this.iocRuntimeEnvironmentPanelState?.terminal === terminal) {
        this.iocRuntimeEnvironmentPanelState = {
          ...this.iocRuntimeEnvironmentPanelState,
          message: "The tracked IOC process has stopped.",
          isLoading: false,
        };
        void this.postIocRuntimeEnvironmentState();
      }
    }
  }

  async setIocShellTerminalInteractive() {
    const terminal = await this.pickIocShellTerminal();
    if (!terminal) {
      return;
    }

    this.iocShellTerminal = terminal;
    this.refresh(this.contextNode);
    vscode.window.setStatusBarMessage(
      `EPICS IOC terminal: ${terminal.name}`,
      3000,
    );
  }

  async runIocShellCommandInteractive(initialValue = "") {
    const terminal = await this.resolveIocShellTerminal(true);
    if (!terminal) {
      return;
    }

    const commandText = await vscode.window.showInputBox({
      prompt: `Send IOC shell command to terminal "${terminal.name}"`,
      placeHolder: "dbl",
      value: initialValue || this.lastIocShellCommand || "",
      ignoreFocusOut: true,
    });
    if (commandText === undefined) {
      return;
    }

    await this.sendIocShellCommand(commandText);
  }

  async sendNamedIocShellCommand(commandText) {
    await this.sendIocShellCommand(commandText);
  }

  async runDbprForActiveRecord() {
    const recordName = this.resolveActiveIocShellRecordName();
    if (!recordName) {
      vscode.window.showWarningMessage(
        "No EPICS record could be resolved from the active editor for dbpr.",
      );
      return;
    }

    await this.sendIocShellCommand(`dbpr ${recordName}`);
  }

  async sendIocShellCommand(commandText, terminalOverride) {
    const trimmedCommand = String(commandText || "").trim();
    if (!trimmedCommand) {
      return false;
    }

    const terminal =
      terminalOverride && vscode.window.terminals.includes(terminalOverride)
        ? terminalOverride
        : await this.resolveIocShellTerminal(true);
    if (!terminal) {
      return false;
    }

    this.iocShellTerminal = terminal;
    this.lastIocShellCommand = trimmedCommand;
    terminal.sendText(trimmedCommand, true);
    this.refresh(this.contextNode);
    vscode.window.setStatusBarMessage(`IOC> ${trimmedCommand}`, 2500);
    return true;
  }

  async captureStartupTerminalExecutionOutput(terminal, execution) {
    if (!terminal || !execution || typeof execution.read !== "function") {
      return;
    }

    try {
      for await (const chunk of execution.read()) {
        this.appendStartupTerminalOutput(terminal, chunk);
      }
    } catch (_error) {
      // Ignore shell-integration output capture failures. IOC control still works.
    }
  }

  appendStartupTerminalOutput(terminal, chunk) {
    if (!terminal || !chunk) {
      return;
    }

    const previousText = this.startupTerminalOutputByTerminal.get(terminal) || "";
    const nextText = `${previousText}${String(chunk)}`;
    const trimmedText =
      nextText.length > STARTUP_TERMINAL_OUTPUT_MAX_LENGTH
        ? nextText.slice(nextText.length - STARTUP_TERMINAL_OUTPUT_MAX_LENGTH)
        : nextText;
    this.startupTerminalOutputByTerminal.set(terminal, trimmedText);
  }

  parseIocRuntimeCommandsFromHelpOutput(outputText) {
    const sanitizedText = stripAnsiTerminalText(outputText)
      .replace(/\r/g, "")
      .replace(/\u0000/g, "");
    const terminatorIndex = sanitizedText.indexOf("Type 'help <glob>'");
    const relevantText =
      terminatorIndex >= 0 ? sanitizedText.slice(0, terminatorIndex) : sanitizedText;
    const commands = [];
    const seenCommands = new Set();

    for (const line of relevantText.split("\n")) {
      for (const token of line.split(/\s+/)) {
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(token)) {
          continue;
        }
        if (seenCommands.has(token)) {
          continue;
        }
        seenCommands.add(token);
        commands.push(token);
      }
    }

    return {
      commands,
      isComplete: terminatorIndex >= 0,
    };
  }

  async fetchIocRuntimeCommandsForTerminal(terminal) {
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      return [];
    }

    const cachedCommands = this.iocRuntimeCommandsByTerminal.get(terminal);
    if (Array.isArray(cachedCommands) && cachedCommands.length) {
      return cachedCommands;
    }

    const helpOutput = await this.captureIocShellCommandOutput("help", terminal);
    if (helpOutput === undefined) {
      return [];
    }
    const parsed = this.parseIocRuntimeCommandsFromHelpOutput(helpOutput);
    if (parsed.commands.length) {
      this.iocRuntimeCommandsByTerminal.set(terminal, parsed.commands);
    }
    return parsed.commands;
  }

  parseIocRuntimeCommandHelpOutput(outputText, requestedCommandNames) {
    const sanitizedText = stripAnsiTerminalText(outputText)
      .replace(/\r/g, "")
      .replace(/\u0000/g, "");
    const normalizedCommandNames = Array.isArray(requestedCommandNames)
      ? requestedCommandNames.map((name) => String(name || "").trim()).filter(Boolean)
      : [];
    const remainingCommandNames = new Set(normalizedCommandNames);
    const helpByCommand = new Map();
    let activeCommandName = undefined;
    let activeLines = [];

    const flushActiveHelp = () => {
      if (!activeCommandName) {
        return;
      }
      const helpText = activeLines.join("\n").trim();
      helpByCommand.set(
        activeCommandName,
        helpText || `No detailed help is available for ${activeCommandName}.`,
      );
      remainingCommandNames.delete(activeCommandName);
      activeCommandName = undefined;
      activeLines = [];
    };

    for (const rawLine of sanitizedText.split("\n")) {
      const line = String(rawLine || "");
      const trimmedLine = line.trim();
      const matchingCommandName = normalizedCommandNames.find((commandName) =>
        trimmedLine === commandName || trimmedLine.startsWith(`${commandName} `),
      );
      if (matchingCommandName) {
        flushActiveHelp();
        activeCommandName = matchingCommandName;
        activeLines = [trimmedLine];
        continue;
      }
      if (activeCommandName) {
        activeLines.push(line.replace(/\s+$/g, ""));
      }
    }
    flushActiveHelp();

    for (const commandName of remainingCommandNames) {
      helpByCommand.set(
        commandName,
        `No detailed help is available for ${commandName}.`,
      );
    }

    return helpByCommand;
  }

  async fetchIocRuntimeCommandHelpBatch(commandNames, terminal) {
    const normalizedCommandNames = Array.isArray(commandNames)
      ? commandNames.map((name) => String(name || "").trim()).filter(Boolean)
      : [];
    if (!normalizedCommandNames.length || !terminal || !vscode.window.terminals.includes(terminal)) {
      return new Map();
    }

    let helpByCommand = this.iocRuntimeCommandHelpByTerminal.get(terminal);
    if (!helpByCommand) {
      helpByCommand = new Map();
      this.iocRuntimeCommandHelpByTerminal.set(terminal, helpByCommand);
    }

    const missingCommandNames = normalizedCommandNames.filter(
      (commandName) => !helpByCommand.has(commandName),
    );
    if (missingCommandNames.length) {
      const helpOutput = await this.captureIocShellCommandOutput(
        `help ${missingCommandNames.join(" ")}`,
        terminal,
      );
      const parsedHelp =
        helpOutput === undefined
          ? new Map(
              missingCommandNames.map((commandName) => [
                commandName,
                `No detailed help is available for ${commandName}.`,
              ]),
            )
          : this.parseIocRuntimeCommandHelpOutput(helpOutput, missingCommandNames);
      for (const [commandName, helpText] of parsedHelp.entries()) {
        helpByCommand.set(commandName, helpText);
      }
    }

    const result = new Map();
    for (const commandName of normalizedCommandNames) {
      result.set(
        commandName,
        helpByCommand.get(commandName) ||
          `No detailed help is available for ${commandName}.`,
      );
    }
    return result;
  }

  async prefetchIocRuntimeCommandHelpForPanel(commandNames, terminal, startupDocumentPath) {
    const normalizedCommandNames = Array.isArray(commandNames)
      ? commandNames.map((name) => String(name || "").trim()).filter(Boolean)
      : [];
    if (!normalizedCommandNames.length || !terminal || !vscode.window.terminals.includes(terminal)) {
      return;
    }

    const batchSize = 24;
    for (let index = 0; index < normalizedCommandNames.length; index += batchSize) {
      const batch = normalizedCommandNames.slice(index, index + batchSize);
      const helpMap = await this.fetchIocRuntimeCommandHelpBatch(batch, terminal);
      if (!this.iocRuntimeCommandsPanel?.webview) {
        return;
      }
      const panelState = this.iocRuntimeCommandsPanelState;
      if (
        !panelState ||
        panelState.terminal !== terminal ||
        panelState.startupDocumentPath !== startupDocumentPath
      ) {
        return;
      }
      for (const [commandName, helpText] of helpMap.entries()) {
        await this.iocRuntimeCommandsPanel.webview.postMessage({
          type: "iocRuntimeCommandHelp",
          commandName,
          helpText:
            helpText || `No detailed help is available for ${commandName}.`,
        });
      }
    }
  }

  async fetchIocRuntimeCommandHelp(commandName, terminal) {
    const normalizedCommandName = String(commandName || "").trim();
    if (!normalizedCommandName || !terminal || !vscode.window.terminals.includes(terminal)) {
      return undefined;
    }
    const helpMap = await this.fetchIocRuntimeCommandHelpBatch(
      [normalizedCommandName],
      terminal,
    );
    return helpMap.get(normalizedCommandName);
  }

  parseIocRuntimeVariablesOutput(outputText) {
    const sanitizedText = stripAnsiTerminalText(outputText)
      .replace(/\r/g, "")
      .replace(/\u0000/g, "");
    const variables = [];
    const seenVariableNames = new Set();

    for (const rawLine of sanitizedText.split("\n")) {
      const trimmedLine = String(rawLine || "").trim();
      if (!trimmedLine || /^epics>/.test(trimmedLine)) {
        continue;
      }
      const match = trimmedLine.match(/^(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
      if (!match) {
        continue;
      }
      const variableType = String(match[1] || "").trim();
      const variableName = String(match[2] || "").trim();
      const currentValue = String(match[3] || "").trim();
      if (!variableName || seenVariableNames.has(variableName)) {
        continue;
      }
      seenVariableNames.add(variableName);
      variables.push({
        variableType,
        variableName,
        currentValue,
      });
    }

    return variables;
  }

  async fetchIocRuntimeVariablesForTerminal(terminal) {
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      return [];
    }

    const cachedVariables = this.iocRuntimeVariablesByTerminal.get(terminal);
    if (Array.isArray(cachedVariables) && cachedVariables.length) {
      return cachedVariables;
    }

    const variableOutput = await this.captureIocShellCommandOutput("var", terminal);
    if (variableOutput === undefined) {
      return [];
    }
    const parsedVariables = this.parseIocRuntimeVariablesOutput(variableOutput);
    if (parsedVariables.length) {
      this.iocRuntimeVariablesByTerminal.set(terminal, parsedVariables);
    }
    return parsedVariables;
  }

  parseIocRuntimeEnvironmentOutput(outputText) {
    const sanitizedText = stripAnsiTerminalText(outputText)
      .replace(/\r/g, "")
      .replace(/\u0000/g, "");
    const entries = [];
    const seenNames = new Set();

    for (const rawLine of sanitizedText.split("\n")) {
      const trimmedLine = String(rawLine || "").trim();
      if (!trimmedLine || /^epics>/.test(trimmedLine)) {
        continue;
      }
      const separatorIndex = trimmedLine.indexOf("=");
      if (separatorIndex <= 0) {
        continue;
      }
      const variableName = trimmedLine.slice(0, separatorIndex).trim();
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(variableName) || seenNames.has(variableName)) {
        continue;
      }
      seenNames.add(variableName);
      entries.push({
        variableName,
        currentValue: trimmedLine.slice(separatorIndex + 1),
      });
    }

    return entries;
  }

  async fetchIocRuntimeEnvironmentForTerminal(terminal) {
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      return [];
    }

    const cachedEntries = this.iocRuntimeEnvironmentByTerminal.get(terminal);
    if (Array.isArray(cachedEntries) && cachedEntries.length) {
      return cachedEntries;
    }

    const outputText = await this.captureIocShellCommandOutput("epicsEnvShow", terminal);
    if (outputText === undefined) {
      return [];
    }
    const parsedEntries = this.parseIocRuntimeEnvironmentOutput(outputText);
    if (parsedEntries.length) {
      this.iocRuntimeEnvironmentByTerminal.set(terminal, parsedEntries);
    }
    return parsedEntries;
  }

  disposeIocRuntimeVariablesReloadTimer() {
    if (!this.iocRuntimeVariablesReloadTimer) {
      return;
    }
    clearTimeout(this.iocRuntimeVariablesReloadTimer);
    this.iocRuntimeVariablesReloadTimer = undefined;
  }

  scheduleIocRuntimeVariablesPanelReload(delayMs = 1500) {
    this.disposeIocRuntimeVariablesReloadTimer();
    this.iocRuntimeVariablesReloadTimer = setTimeout(() => {
      this.iocRuntimeVariablesReloadTimer = undefined;
      void this.reloadIocRuntimeVariablesPanel();
    }, Math.max(0, Number(delayMs) || 0));
  }

  disposeIocRuntimeEnvironmentReloadTimer() {
    if (!this.iocRuntimeEnvironmentReloadTimer) {
      return;
    }
    clearTimeout(this.iocRuntimeEnvironmentReloadTimer);
    this.iocRuntimeEnvironmentReloadTimer = undefined;
  }

  scheduleIocRuntimeEnvironmentPanelReload(delayMs = 1500) {
    this.disposeIocRuntimeEnvironmentReloadTimer();
    this.iocRuntimeEnvironmentReloadTimer = setTimeout(() => {
      this.iocRuntimeEnvironmentReloadTimer = undefined;
      void this.reloadIocRuntimeEnvironmentPanel();
    }, Math.max(0, Number(delayMs) || 0));
  }

  formatIocShellQuotedString(value) {
    return `"${String(value ?? "")
      .replace(/\\/g, "\\\\")
      .replace(/"/g, '\\"')
      .replace(/\r/g, "\\r")
      .replace(/\n/g, "\\n")}"`;
  }

  async captureIocShellCommandOutput(commandText, terminal) {
    const trimmedCommand = String(commandText || "").trim();
    if (!trimmedCommand) {
      return undefined;
    }
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      return undefined;
    }

    const outputFilePath = path.join(
      os.tmpdir(),
      `epics-workbench-ioc-output-${Date.now()}-${createNonce()}.txt`,
    );
    try {
      await fs.promises.rm(outputFilePath, { force: true });
    } catch (_error) {
      // Ignore cleanup failures before capture.
    }

    const redirectedCommand = `${trimmedCommand} > ${outputFilePath}`;
    const sent = await this.sendIocShellCommand(redirectedCommand, terminal);
    if (!sent) {
      return undefined;
    }

    const deadline = Date.now() + IOC_RUNTIME_COMMANDS_FETCH_TIMEOUT_MS;
    let outputText = "";
    while (Date.now() < deadline) {
      try {
        outputText = await fs.promises.readFile(outputFilePath, "utf8");
        break;
      } catch (error) {
        if (error?.code !== "ENOENT") {
          throw error;
        }
      }
      await new Promise((resolve) => setTimeout(resolve, 120));
    }

    try {
      await fs.promises.rm(outputFilePath, { force: true });
    } catch (_error) {
      // Ignore temp-file cleanup failures after capture.
    }

    if (outputText === undefined) {
      return undefined;
    }

    return outputText;
  }

  async runIocShellCommandWithCapturedOutput(commandText, terminal) {
    const trimmedCommand = String(commandText || "").trim();
    const outputText = await this.captureIocShellCommandOutput(trimmedCommand, terminal);
    if (outputText === undefined) {
      return undefined;
    }

    const document = await vscode.workspace.openTextDocument({
      language: "plaintext",
      content: `# ${trimmedCommand}\n\n${outputText || ""}`,
    });
    await vscode.window.showTextDocument(document, vscode.ViewColumn.Active, false);
    return outputText;
  }

  trackStartupIocTerminal(documentPath, terminal) {
    const normalizedDocumentPath = normalizeFsPath(documentPath);
    const existingTerminal = this.iocStartupTerminalByDocumentPath.get(normalizedDocumentPath);
    if (existingTerminal && existingTerminal !== terminal) {
      this.iocStartupDocumentPathByTerminal.delete(existingTerminal);
    }

    const previousDocumentPath = this.iocStartupDocumentPathByTerminal.get(terminal);
    if (previousDocumentPath && previousDocumentPath !== normalizedDocumentPath) {
      this.iocStartupTerminalByDocumentPath.delete(previousDocumentPath);
    }

    this.iocStartupTerminalByDocumentPath.set(normalizedDocumentPath, terminal);
    this.iocStartupDocumentPathByTerminal.set(terminal, normalizedDocumentPath);
    this.iocShellTerminal = terminal;
    this.refresh(this.contextNode);
    this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
  }

  untrackStartupIocTerminal(terminal) {
    if (!terminal) {
      return;
    }

    const documentPath = this.iocStartupDocumentPathByTerminal.get(terminal);
    if (documentPath) {
      this.iocStartupDocumentPathByTerminal.delete(terminal);
      const currentTerminal = this.iocStartupTerminalByDocumentPath.get(documentPath);
      if (currentTerminal === terminal) {
        this.iocStartupTerminalByDocumentPath.delete(documentPath);
      }
    }
  }

  getTerminalActiveExecutionCount(terminal) {
    return Number(this.terminalActiveExecutionCount.get(terminal) || 0);
  }

  getTerminalCurrentDirectory(terminal) {
    const shellIntegrationCwd = terminal?.shellIntegration?.cwd;
    if (shellIntegrationCwd?.scheme === "file") {
      return normalizeFsPath(shellIntegrationCwd.fsPath);
    }

    return this.terminalLastKnownCwd.get(terminal) || "";
  }

  findReusableStartupIocTerminal(workingDirectory, startupDocumentPath) {
    const normalizedWorkingDirectory = normalizeFsPath(workingDirectory);
    const normalizedStartupDocumentPath = normalizeFsPath(startupDocumentPath);
    const trackedTerminal = this.iocStartupTerminalByDocumentPath.get(normalizedStartupDocumentPath);
    const trackedTerminals = new Set(this.iocStartupDocumentPathByTerminal.keys());

    for (const terminal of vscode.window.terminals || []) {
      if (trackedTerminal && terminal === trackedTerminal) {
        continue;
      }
      if (trackedTerminals.has(terminal)) {
        continue;
      }
      if (this.getTerminalActiveExecutionCount(terminal) > 0) {
        continue;
      }
      if (this.getTerminalCurrentDirectory(terminal) !== normalizedWorkingDirectory) {
        continue;
      }
      return terminal;
    }

    return undefined;
  }

  async showStartupIocNotification(kind, startupDocumentPath, terminal) {
    if (!startupDocumentPath || !terminal) {
      return;
    }

    const startupFileName = path.basename(startupDocumentPath);
    const terminalName = terminal.name || "Terminal";
    const showTerminalAction = "Show Terminal";
    const message =
      kind === "started"
        ? `Started ${startupFileName} in terminal "${terminalName}".`
        : `Stopped ${startupFileName} in terminal "${terminalName}".`;
    const selectedAction = await vscode.window.showInformationMessage(
      message,
      showTerminalAction,
    );
    if (selectedAction === showTerminalAction) {
      this.iocShellTerminal = terminal;
      terminal.show(true);
      this.refresh(this.contextNode);
    }
  }

  async resolveIocShellTerminal(promptIfNeeded = true) {
    if (
      this.iocShellTerminal &&
      vscode.window.terminals.includes(this.iocShellTerminal)
    ) {
      return this.iocShellTerminal;
    }

    if (vscode.window.activeTerminal) {
      this.iocShellTerminal = vscode.window.activeTerminal;
      this.refresh(this.contextNode);
      return this.iocShellTerminal;
    }

    if (vscode.window.terminals.length === 1) {
      this.iocShellTerminal = vscode.window.terminals[0];
      this.refresh(this.contextNode);
      return this.iocShellTerminal;
    }

    if (!promptIfNeeded) {
      return undefined;
    }

    const pickedTerminal = await this.pickIocShellTerminal();
    if (pickedTerminal) {
      this.iocShellTerminal = pickedTerminal;
      this.refresh(this.contextNode);
    }
    return pickedTerminal;
  }

  async pickIocShellTerminal() {
    const terminals = vscode.window.terminals || [];
    if (terminals.length === 0) {
      vscode.window.showWarningMessage(
        "No VS Code terminal is available. Start the IOC in an integrated terminal first.",
      );
      return undefined;
    }

    if (terminals.length === 1) {
      return terminals[0];
    }

    const selected = await vscode.window.showQuickPick(
      terminals.map((terminal) => ({
        label: terminal.name,
        terminal,
      })),
      {
        placeHolder: "Select the IOC shell terminal",
        ignoreFocusOut: true,
      },
    );
    return selected?.terminal;
  }

  async openStartupDocumentPath(startupDocumentPath, preserveFocus = true) {
    if (!startupDocumentPath) {
      return undefined;
    }

    const document = await vscode.workspace.openTextDocument(startupDocumentPath);
    await vscode.window.showTextDocument(document, {
      preview: false,
      preserveFocus,
    });
    return document;
  }

  registerWidgetMenuContext(panel, kind, state) {
    if (!panel) {
      return;
    }

    const updateContext = (isActive) => {
      if (!isActive) {
        if (this.activeWidgetMenuContext?.panel === panel) {
          this.activeWidgetMenuContext = undefined;
        }
        return;
      }

      this.activeWidgetMenuContext = {
        panel,
        kind,
        state,
      };
    };

    updateContext(panel.active);
    panel.onDidChangeViewState((event) => {
      updateContext(event.webviewPanel.active);
    });
    panel.onDidDispose(() => {
      if (this.activeWidgetMenuContext?.panel === panel) {
        this.activeWidgetMenuContext = undefined;
      }
    });
  }

  getActiveWidgetMenuContext() {
    const context = this.activeWidgetMenuContext;
    if (!context?.panel || context.panel.visible === false) {
      return undefined;
    }
    return context;
  }

  getProbeWidgetCommandOptionsFromActiveWidget() {
    const context = this.getActiveWidgetMenuContext();
    if (!context) {
      return undefined;
    }

    if (context.kind === "probe") {
      return {
        recordName: String(context.state?.recordName || "").trim(),
      };
    }

    if (context.kind === "pvlist") {
      const rows = Array.isArray(context.state?.rows) ? context.state.rows : [];
      const firstChannelName = rows
        .map((row) => String(row?.channelName || "").trim())
        .find(Boolean);
      return {
        recordName: firstChannelName || "",
      };
    }

    if (context.kind === "monitor") {
      const firstChannelName = (context.state?.channelRows || [])
        .map((row) => String(row?.channelName || "").trim())
        .find(Boolean);
      return {
        recordName: firstChannelName || "",
      };
    }

    return undefined;
  }

  getPvlistWidgetCommandOptionsFromActiveWidget() {
    const context = this.getActiveWidgetMenuContext();
    if (!context) {
      return undefined;
    }

    if (context.kind === "probe") {
      const recordName = String(context.state?.recordName || "").trim();
      return {
        sourceKind: "pvlist",
        sourceLabel: context.state?.panel?.title || "EPICS Probe",
        sourceText: recordName ? `${recordName}\n` : "",
      };
    }

    if (context.kind === "pvlist") {
      const sourceModel = context.state?.sourceModel;
      return {
        sourceKind: "pvlist",
        sourceLabel: sourceModel?.sourceLabel || context.state?.panel?.title || "EPICS PvList",
        sourceDocumentUri: sourceModel?.sourceDocumentUri || "",
        sourceText: buildPvlistWidgetFileText(
          sourceModel,
          context.state?.macroValues,
        ),
      };
    }

    if (context.kind === "monitor") {
      const channelNames = (context.state?.channelRows || [])
        .map((row) => String(row?.channelName || "").trim())
        .filter(Boolean);
      return {
        sourceKind: "pvlist",
        sourceLabel: context.state?.sourceLabel || context.state?.panel?.title || "EPICS Monitor",
        sourceText: channelNames.join("\n"),
      };
    }

    return undefined;
  }

  getMonitorWidgetCommandOptionsFromActiveWidget() {
    const context = this.getActiveWidgetMenuContext();
    if (!context) {
      return undefined;
    }

    if (context.kind === "probe") {
      const recordName = String(context.state?.recordName || "").trim();
      return {
        sourceLabel: context.state?.panel?.title || "EPICS Probe",
        initialChannels: recordName ? [recordName] : [],
      };
    }

    if (context.kind === "pvlist") {
      const channelNames = (context.state?.rows || [])
        .map((row) => String(row?.channelName || "").trim())
        .filter(Boolean);
      return {
        sourceLabel: context.state?.sourceModel?.sourceLabel || context.state?.panel?.title || "EPICS PvList",
        initialChannels: channelNames,
      };
    }

    if (context.kind === "monitor") {
      const channelNames = (context.state?.channelRows || [])
        .map((row) => String(row?.channelName || "").trim())
        .filter(Boolean);
      return {
        sourceLabel: context.state?.sourceLabel || context.state?.panel?.title || "EPICS Monitor",
        initialChannels: channelNames,
      };
    }

    return undefined;
  }

  getTrackedStartupIocTerminal(startupDocumentPath) {
    const normalizedStartupDocumentPath = normalizeFsPath(startupDocumentPath);
    if (!normalizedStartupDocumentPath) {
      return undefined;
    }

    const terminal = this.iocStartupTerminalByDocumentPath.get(normalizedStartupDocumentPath);
    if (terminal && vscode.window.terminals.includes(terminal)) {
      return terminal;
    }

    if (terminal) {
      this.untrackStartupIocTerminal(terminal);
    }
    return undefined;
  }

  getCandidateRunningStartupIocTerminals(startupDocumentPath) {
    const normalizedStartupDocumentPath = normalizeFsPath(startupDocumentPath);
    if (!normalizedStartupDocumentPath) {
      return [];
    }

    const trackedTerminal = this.getTrackedStartupIocTerminal(normalizedStartupDocumentPath);
    if (trackedTerminal) {
      return [trackedTerminal];
    }

    if (!canInferRunningStartupIocByWorkingDirectory(normalizedStartupDocumentPath)) {
      return [];
    }

    const workingDirectory = normalizeFsPath(path.dirname(normalizedStartupDocumentPath));
    return (vscode.window.terminals || []).filter((terminal) =>
      this.getTerminalActiveExecutionCount(terminal) > 0 &&
      this.getTerminalCurrentDirectory(terminal) === workingDirectory,
    );
  }

  async resolveRunningStartupIocTerminal(
    startupDocumentPath,
    {
      promptIfAmbiguous = false,
      placeHolder = "Select the running IOC terminal",
    } = {},
  ) {
    const candidates = this.getCandidateRunningStartupIocTerminals(startupDocumentPath);
    if (!candidates.length) {
      return undefined;
    }

    if (candidates.length === 1 || !promptIfAmbiguous) {
      const terminal = candidates[0];
      this.trackStartupIocTerminal(startupDocumentPath, terminal);
      return terminal;
    }

    const selected = await vscode.window.showQuickPick(
      candidates.map((terminal) => ({
        label: terminal.name || "Terminal",
        description: this.getTerminalCurrentDirectory(terminal),
        terminal,
      })),
      {
        placeHolder,
        ignoreFocusOut: true,
      },
    );
    if (!selected?.terminal) {
      return undefined;
    }

    this.trackStartupIocTerminal(startupDocumentPath, selected.terminal);
    return selected.terminal;
  }

  async resolveEpicsProjectRootPath(resourceUri) {
    const resourceFsPath = getCommandResourceFsPath(resourceUri);
    const containingRootPath = findContainingEpicsProjectRootPath(
      resourceFsPath || vscode.window.activeTextEditor?.document?.uri?.fsPath,
    );
    if (containingRootPath) {
      return containingRootPath;
    }

    const candidateRootPaths = getWorkspaceEpicsProjectRootPaths();
    if (candidateRootPaths.length === 1) {
      return candidateRootPaths[0];
    }
    if (!candidateRootPaths.length) {
      return undefined;
    }

    const selected = await vscode.window.showQuickPick(
      candidateRootPaths.map((candidateRootPath) => ({
        label: path.basename(candidateRootPath),
        description: candidateRootPath,
        rootPath: candidateRootPath,
      })),
      {
        placeHolder: "Select the EPICS project for IOC startup commands",
        ignoreFocusOut: true,
      },
    );
    return selected?.rootPath;
  }

  async resolveProjectStartupDocumentPaths(resourceUri) {
    const projectRootPath = await this.resolveEpicsProjectRootPath(resourceUri);
    if (!projectRootPath) {
      vscode.window.showWarningMessage(
        "No EPICS project root is available for IOC startup commands.",
      );
      return undefined;
    }

    const startupDocumentPaths = findProjectStcmdLikeFilePaths(projectRootPath);
    if (!startupDocumentPaths.length) {
      vscode.window.showWarningMessage(
        `No st.cmd-like files were found under ${path.join(projectRootPath, "iocBoot")}.`,
      );
      return undefined;
    }

    return {
      projectRootPath,
      startupDocumentPaths,
    };
  }

  createProjectStartupIocQuickPickItems(projectRootPath, startupDocumentPaths) {
    return startupDocumentPaths.map((startupDocumentPath) => {
      const runningTerminal = this.getCandidateRunningStartupIocTerminals(startupDocumentPath)[0];
      const startupValidation = resolveStartupExecutableValidation(startupDocumentPath);
      const executableErrorMessage =
        startupValidation.executableText && !startupValidation.executablePath
          ? `Error: executable ${startupValidation.executableName} missing`
          : "";
      return {
        label: getIocBootDisplayPath(projectRootPath, startupDocumentPath),
        description: runningTerminal ? "Running" : "",
        detail: executableErrorMessage || (runningTerminal
          ? `Terminal: ${runningTerminal.name || "Terminal"}`
          : "Not running"),
        buttons: executableErrorMessage
          ? [
            STARTUP_IOC_PICKER_OPEN_FILE_BUTTON,
            ...(runningTerminal ? [STARTUP_IOC_PICKER_BRING_TO_FRONT_BUTTON] : []),
          ]
          : runningTerminal
            ? [
              STARTUP_IOC_PICKER_OPEN_FILE_BUTTON,
              STARTUP_IOC_PICKER_BRING_TO_FRONT_BUTTON,
              STARTUP_IOC_PICKER_STOP_BUTTON,
            ]
            : [
              STARTUP_IOC_PICKER_OPEN_FILE_BUTTON,
              STARTUP_IOC_PICKER_START_BUTTON,
              STARTUP_IOC_PICKER_STOP_BUTTON,
            ],
        startupDocumentPath,
      };
    });
  }

  async bringStartupIocToFront(
    startupDocumentPath,
    {
      promptIfAmbiguous = true,
      notifyIfMissing = true,
    } = {},
  ) {
    await this.openStartupDocumentPath(startupDocumentPath, true);
    const terminal = await this.resolveRunningStartupIocTerminal(
      startupDocumentPath,
      {
        promptIfAmbiguous,
        placeHolder: `Select the running IOC terminal for ${path.basename(startupDocumentPath)}`,
      },
    );
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      if (notifyIfMissing) {
        vscode.window.showInformationMessage(
          `IOC is not running for ${path.basename(startupDocumentPath)}.`,
        );
      }
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      return undefined;
    }

    this.iocShellTerminal = terminal;
    terminal.show(true);
    this.refresh(this.contextNode);
    return terminal;
  }

  async stopStartupIocForDocumentPath(
    startupDocumentPath,
    {
      promptIfAmbiguous = true,
      notifyIfMissing = true,
    } = {},
  ) {
    await this.openStartupDocumentPath(startupDocumentPath, true);
    const terminal = await this.resolveRunningStartupIocTerminal(
      startupDocumentPath,
      {
        promptIfAmbiguous,
        placeHolder: `Select the running IOC terminal for ${path.basename(startupDocumentPath)}`,
      },
    );
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      if (notifyIfMissing) {
        vscode.window.showInformationMessage(
          `IOC is not running for ${path.basename(startupDocumentPath)}.`,
        );
      }
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      return undefined;
    }

    this.iocShellTerminal = terminal;
    terminal.show(true);
    terminal.sendText("\u0003", false);
    this.untrackStartupIocTerminal(terminal);
    this.terminalActiveExecutionCount.delete(terminal);
    this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
    this.refresh(this.contextNode);
    vscode.window.setStatusBarMessage(
      `Sent interrupt to IOC for ${path.basename(startupDocumentPath)}`,
      3000,
    );
    return terminal;
  }

  async startOrBringStartupIocToFront(
    startupDocumentPath,
    {
      notify = false,
    } = {},
  ) {
    const runningTerminal = await this.bringStartupIocToFront(
      startupDocumentPath,
      {
        promptIfAmbiguous: true,
        notifyIfMissing: false,
      },
    );
    if (runningTerminal) {
      return runningTerminal;
    }

    const startedTerminal = await this.startStartupIocForDocumentPath(
      startupDocumentPath,
      {
        showTerminal: true,
        notify,
      },
    );
    if (!startedTerminal) {
      return undefined;
    }

    await this.openStartupDocumentPath(startupDocumentPath, true);
    return startedTerminal;
  }

  async showProjectStartupIocPicker(resourceUri) {
    const resolvedProjectStartupPaths =
      await this.resolveProjectStartupDocumentPaths(resourceUri);
    if (!resolvedProjectStartupPaths) {
      return;
    }

    const { projectRootPath, startupDocumentPaths } = resolvedProjectStartupPaths;
    const quickPick = vscode.window.createQuickPick();
    let isDisposed = false;

    const refreshItems = () => {
      if (isDisposed) {
        return;
      }
      quickPick.items = this.createProjectStartupIocQuickPickItems(
        projectRootPath,
        startupDocumentPaths,
      );
    };

    const runPickerAction = async (callback) => {
      if (isDisposed) {
        return;
      }
      quickPick.busy = true;
      quickPick.enabled = false;
      try {
        await callback();
      } finally {
        if (!isDisposed) {
          refreshItems();
          quickPick.busy = false;
          quickPick.enabled = true;
        }
      }
    };

    quickPick.title = "EPICS IOC Start/Stop";
    quickPick.placeholder = "Use the item buttons to open the startup file, start the IOC, stop it, or bring its terminal to front";
    quickPick.ignoreFocusOut = false;
    quickPick.matchOnDescription = true;
    quickPick.matchOnDetail = true;
    refreshItems();

    const hideDisposable = quickPick.onDidHide(() => {
      isDisposed = true;
      hideDisposable.dispose();
      acceptDisposable.dispose();
      selectionDisposable.dispose();
      triggerButtonDisposable.dispose();
      quickPick.dispose();
    });
    const acceptDisposable = quickPick.onDidAccept(() => {});
    const selectionDisposable = quickPick.onDidChangeSelection(() => {});
    const triggerButtonDisposable = quickPick.onDidTriggerItemButton(async (event) => {
      const startupDocumentPath = event.item?.startupDocumentPath;
      if (!startupDocumentPath) {
        return;
      }

      if (event.button === STARTUP_IOC_PICKER_OPEN_FILE_BUTTON) {
        await this.openStartupDocumentPath(startupDocumentPath, false);
        return;
      }

      await runPickerAction(async () => {
        if (event.button === STARTUP_IOC_PICKER_START_BUTTON) {
          await this.startOrBringStartupIocToFront(startupDocumentPath);
          return;
        }
        if (event.button === STARTUP_IOC_PICKER_BRING_TO_FRONT_BUTTON) {
          await this.bringStartupIocToFront(startupDocumentPath);
          return;
        }
        if (event.button === STARTUP_IOC_PICKER_STOP_BUTTON) {
          await this.stopStartupIocForDocumentPath(startupDocumentPath);
        }
      });
    });

    quickPick.show();
  }

  async pickProjectStartupScript(resourceUri, actionLabel) {
    const resolvedProjectStartupPaths =
      await this.resolveProjectStartupDocumentPaths(resourceUri);
    if (!resolvedProjectStartupPaths) {
      return undefined;
    }

    const { startupDocumentPaths } = resolvedProjectStartupPaths;

    const selected = await vscode.window.showQuickPick(
      startupDocumentPaths.map((startupDocumentPath) => {
        const runningTerminal = this.getCandidateRunningStartupIocTerminals(startupDocumentPath)[0];
        const startupValidation = resolveStartupExecutableValidation(startupDocumentPath);
        const executableErrorMessage =
          startupValidation.executableText && !startupValidation.executablePath
            ? `Error: executable ${startupValidation.executableName} missing`
            : "";
        return {
          label: startupDocumentPath,
          description: runningTerminal ? "Running" : "",
          detail: executableErrorMessage || (runningTerminal
            ? `Terminal: ${runningTerminal.name || "Terminal"}`
            : ""),
          startupDocumentPath,
        };
      }),
      {
        placeHolder: `Select the IOC startup file to ${actionLabel}`,
        ignoreFocusOut: true,
      },
    );
    return selected?.startupDocumentPath;
  }

  async startProjectStartupIoc(resourceUri) {
    const startupDocumentPath = await this.pickProjectStartupScript(resourceUri, "start");
    if (!startupDocumentPath) {
      return;
    }

    await this.startOrBringStartupIocToFront(startupDocumentPath);
  }

  async stopProjectStartupIoc(resourceUri) {
    const startupDocumentPath = await this.pickProjectStartupScript(resourceUri, "stop");
    if (!startupDocumentPath) {
      return;
    }

    await this.stopStartupIocForDocumentPath(startupDocumentPath);
  }

  resolveActiveStartupIocDocument(editor = vscode.window.activeTextEditor) {
    const document = editor?.document;
    const documentPath = getIocBootStartupDocumentPath(document);
    if (!document || !documentPath) {
      return undefined;
    }
    return {
      document,
      documentPath,
    };
  }

  async startStartupIocForDocumentPath(
    startupDocumentPath,
    { showTerminal = true, notify = true } = {},
  ) {
    const normalizedStartupDocumentPath = normalizeFsPath(startupDocumentPath);
    if (!normalizedStartupDocumentPath) {
      return undefined;
    }

    const existingTerminal = this.iocStartupTerminalByDocumentPath.get(normalizedStartupDocumentPath);
    if (existingTerminal && vscode.window.terminals.includes(existingTerminal)) {
      this.iocShellTerminal = existingTerminal;
      if (showTerminal) {
        existingTerminal.show(true);
      }
      this.refresh(this.contextNode);
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      if (notify) {
        vscode.window.showInformationMessage(
          `IOC is already running for ${path.basename(normalizedStartupDocumentPath)} in terminal "${existingTerminal.name}".`,
        );
      }
      return existingTerminal;
    }

    const openDocument = vscode.workspace.textDocuments.find(
      (candidate) =>
        candidate.uri?.scheme === "file" &&
        normalizeFsPath(candidate.uri.fsPath) === normalizedStartupDocumentPath,
    );
    const startupDocument =
      openDocument || (await vscode.workspace.openTextDocument(normalizedStartupDocumentPath));
    const startupValidation = resolveStartupExecutableValidation(
      normalizedStartupDocumentPath,
      startupDocument.getText(),
    );
    if (startupValidation.executableText && !startupValidation.executablePath) {
      vscode.window.showErrorMessage(
        `Error: executable ${startupValidation.executableName} missing`,
      );
      return undefined;
    }
    if (startupDocument.isDirty) {
      const saved = await startupDocument.save();
      if (!saved) {
        if (notify) {
          vscode.window.showWarningMessage(
            "Save the startup file before starting the IOC.",
          );
        }
        return undefined;
      }
    }

    const workingDirectory = path.dirname(normalizedStartupDocumentPath);
    const commandText = `./${path.basename(normalizedStartupDocumentPath)}`;
    const terminal =
      this.findReusableStartupIocTerminal(
        workingDirectory,
        normalizedStartupDocumentPath,
      ) ||
      vscode.window.createTerminal({
        name: `IOC: ${path.basename(workingDirectory)}`,
        cwd: workingDirectory,
      });
    this.iocRuntimeCommandsByTerminal.delete(terminal);
    this.iocRuntimeCommandHelpByTerminal.delete(terminal);
    this.iocRuntimeVariablesByTerminal.delete(terminal);
    this.iocRuntimeEnvironmentByTerminal.delete(terminal);
    this.trackStartupIocTerminal(normalizedStartupDocumentPath, terminal);
    if (showTerminal) {
      terminal.show(true);
    }
    if (terminal.shellIntegration) {
      try {
        terminal.shellIntegration.executeCommand(commandText);
      } catch (error) {
        terminal.sendText(commandText, true);
      }
    } else {
      terminal.sendText(commandText, true);
    }
    vscode.window.setStatusBarMessage(
      `Starting IOC: ${commandText}`,
      3000,
    );
    if (notify) {
      void this.showStartupIocNotification(
        "started",
        normalizedStartupDocumentPath,
        terminal,
      );
    }
    return terminal;
  }

  async startActiveStartupIoc() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open an IOC startup file under iocBoot to start the IOC.",
      );
      return;
    }

    await this.startOrBringStartupIocToFront(startupTarget.documentPath, {
      notify: true,
    });
  }

  async stopActiveStartupIoc() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open an IOC startup file under iocBoot to stop the IOC.",
      );
      return;
    }

    const terminal = await this.stopStartupIocForDocumentPath(startupTarget.documentPath);
    if (terminal) {
      void this.showStartupIocNotification(
        "stopped",
        startupTarget.documentPath,
        terminal,
      );
    }
  }

  async showActiveStartupIocTerminal() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open an IOC startup file under iocBoot to show its running terminal.",
      );
      return;
    }

    const terminal = await this.bringStartupIocToFront(startupTarget.documentPath);
    if (terminal) {
      vscode.window.setStatusBarMessage(
        `Showing IOC terminal "${terminal.name}"`,
        2500,
      );
    }
  }

  resolveActiveIocShellRecordName() {
    const editor = vscode.window.activeTextEditor;
    const document = editor?.document;
    const position = editor?.selection?.active;
    if (!document || !position) {
      return undefined;
    }

    if (isDatabaseRuntimeDocument(document)) {
      return this.resolveDatabaseIocShellRecordName(document, position);
    }

    if (isStrictMonitorDocument(document)) {
      return analyzeStrictMonitorText(document.getText(), this.getDefaultProtocol())
        .lineReferences[position.line]
        ?.pvName;
    }

    if (isProbeDocument(document)) {
      const lineText = String(document.lineAt(position.line).text || "").trim();
      if (lineText && !lineText.startsWith("#") && !/\s/.test(lineText)) {
        return lineText;
      }
    }

    if (document.languageId === "startup") {
      const lineText = String(document.lineAt(position.line).text || "");
      const match = lineText.match(/dbpf\(\s*"([^"\n]+)"/);
      if (match?.[1]) {
        return match[1];
      }
    }

    return undefined;
  }

  resolveDatabaseIocShellRecordName(document, position) {
    const text = document.getText();
    const offset = document.offsetAt(position);
    const macroDefinitions = createDatabaseMonitorMacroDefinitions(
      this.extractDatabaseTocMacroAssignments?.(text) || new Map(),
    );
    const macroExpansionCache = new Map();
    const expandRecordName = (recordName) =>
      normalizeDatabaseMonitorPvName(
        expandDatabaseMonitorValue(
          recordName,
          macroDefinitions,
          macroExpansionCache,
          [],
        ),
        recordName,
      );

    for (const tocEntry of this.extractDatabaseTocEntries?.(text) || []) {
      if (offset < tocEntry.nameStart || offset > tocEntry.nameEnd) {
        continue;
      }
      return expandRecordName(tocEntry.recordName);
    }

    for (const declaration of this.extractRecordDeclarations?.(text) || []) {
      if (offset < declaration.recordStart || offset > declaration.recordEnd) {
        continue;
      }
      return expandRecordName(declaration.name);
    }

    return undefined;
  }

  async openRuntimeWidget() {
    if (this.runtimeWidgetPanel) {
      this.runtimeWidgetPanel.reveal(vscode.ViewColumn.Active, false);
      await this.postRuntimeWidgetState();
      return;
    }

    const panel = vscode.window.createWebviewPanel(
      RUNTIME_WIDGET_VIEW_TYPE,
      "EPICS Runtime",
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );
    this.runtimeWidgetPanel = panel;

    panel.webview.html = buildRuntimeWidgetHtml(
      panel.webview,
      this.buildRuntimeWidgetWebviewState(),
    );

    panel.onDidDispose(() => {
      if (this.runtimeWidgetPanel === panel) {
        this.runtimeWidgetPanel = undefined;
      }
    });

    panel.webview.onDidReceiveMessage(async (message) => {
      if (!message?.type) {
        return;
      }

      if (message.type === "openProjectRuntimeConfiguration") {
        await this.openProjectRuntimeConfiguration();
        return;
      }

      if (message.type === "addRuntimeMonitor") {
        await this.addMonitorInteractive();
        return;
      }

      if (message.type === "restartRuntimeContext") {
        await this.restartContext();
        return;
      }

      if (message.type === "stopRuntimeContext") {
        await this.stopContext();
        return;
      }

      if (message.type === "clearRuntimeMonitors") {
        await this.clearMonitors();
        return;
      }

      if (message.type === "setIocShellTerminal") {
        await this.setIocShellTerminalInteractive();
        return;
      }

      if (message.type === "runIocShellCommand") {
        await this.runIocShellCommandInteractive();
        return;
      }

      if (message.type === "runIocShellDbl") {
        await this.sendNamedIocShellCommand("dbl");
        return;
      }

      if (message.type === "runIocShellDbprCurrentRecord") {
        await this.runDbprForActiveRecord();
        return;
      }

      if (message.type === "removeRuntimeMonitor" && message.key) {
        const entry = this.monitorEntries.find((candidate) => candidate.key === message.key);
        if (entry) {
          await this.removeMonitor(entry);
        }
      }
    });

    await this.postRuntimeWidgetState();
  }

  buildIocRuntimeCommandsState(panelState = this.iocRuntimeCommandsPanelState) {
    if (!panelState) {
      return {
        startupFileName: "",
        startupDocumentPath: "",
        terminalName: "",
        commands: [],
        message: "No running IOC startup file is selected.",
        isLoading: false,
        isRunning: false,
      };
    }

    const trackedTerminal = panelState.startupDocumentPath
      ? this.iocStartupTerminalByDocumentPath.get(panelState.startupDocumentPath)
      : undefined;
    const isRunning =
      Boolean(panelState.terminal) &&
      vscode.window.terminals.includes(panelState.terminal) &&
      trackedTerminal === panelState.terminal;

    return {
      startupFileName: path.basename(panelState.startupDocumentPath || ""),
      startupDocumentPath: panelState.startupDocumentPath || "",
      terminalName: panelState.terminal?.name || "",
      commands: Array.isArray(panelState.commands) ? panelState.commands : [],
      message: panelState.message || "",
      isLoading: Boolean(panelState.isLoading),
      isRunning,
    };
  }

  async postIocRuntimeCommandsState(state = this.buildIocRuntimeCommandsState()) {
    if (!this.iocRuntimeCommandsPanel?.webview) {
      return;
    }

    await this.iocRuntimeCommandsPanel.webview.postMessage({
      type: "iocRuntimeCommandsState",
      state,
    });
  }

  buildIocRuntimeVariablesState(panelState = this.iocRuntimeVariablesPanelState) {
    if (!panelState) {
      return {
        startupFileName: "",
        startupDocumentPath: "",
        terminalName: "",
        variables: [],
        message: "No running IOC startup file is selected.",
        isLoading: false,
        isRunning: false,
      };
    }

    const trackedTerminal = panelState.startupDocumentPath
      ? this.iocStartupTerminalByDocumentPath.get(panelState.startupDocumentPath)
      : undefined;
    const isRunning =
      Boolean(panelState.terminal) &&
      vscode.window.terminals.includes(panelState.terminal) &&
      trackedTerminal === panelState.terminal;

    return {
      startupFileName: path.basename(panelState.startupDocumentPath || ""),
      startupDocumentPath: panelState.startupDocumentPath || "",
      terminalName: panelState.terminal?.name || "",
      variables: Array.isArray(panelState.variables) ? panelState.variables : [],
      message: panelState.message || "",
      isLoading: Boolean(panelState.isLoading),
      isRunning,
    };
  }

  async postIocRuntimeVariablesState(state = this.buildIocRuntimeVariablesState()) {
    if (!this.iocRuntimeVariablesPanel?.webview) {
      return;
    }

    await this.iocRuntimeVariablesPanel.webview.postMessage({
      type: "iocRuntimeVariablesState",
      state,
    });
  }

  buildIocRuntimeEnvironmentState(panelState = this.iocRuntimeEnvironmentPanelState) {
    if (!panelState) {
      return {
        startupFileName: "",
        startupDocumentPath: "",
        terminalName: "",
        entries: [],
        message: "No running IOC startup file is selected.",
        isLoading: false,
        isRunning: false,
      };
    }

    const trackedTerminal = panelState.startupDocumentPath
      ? this.iocStartupTerminalByDocumentPath.get(panelState.startupDocumentPath)
      : undefined;
    const isRunning =
      Boolean(panelState.terminal) &&
      vscode.window.terminals.includes(panelState.terminal) &&
      trackedTerminal === panelState.terminal;

    return {
      startupFileName: path.basename(panelState.startupDocumentPath || ""),
      startupDocumentPath: panelState.startupDocumentPath || "",
      terminalName: panelState.terminal?.name || "",
      entries: Array.isArray(panelState.entries) ? panelState.entries : [],
      message: panelState.message || "",
      isLoading: Boolean(panelState.isLoading),
      isRunning,
    };
  }

  async postIocRuntimeEnvironmentState(state = this.buildIocRuntimeEnvironmentState()) {
    if (!this.iocRuntimeEnvironmentPanel?.webview) {
      return;
    }

    await this.iocRuntimeEnvironmentPanel.webview.postMessage({
      type: "iocRuntimeEnvironmentState",
      state,
    });
  }

  async reloadIocRuntimeVariablesPanel(panelState = this.iocRuntimeVariablesPanelState) {
    if (!panelState) {
      return;
    }

    const terminal = panelState.terminal;
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      this.iocRuntimeVariablesPanelState = {
        ...panelState,
        variables: [],
        message: "The tracked IOC terminal is no longer running.",
        isLoading: false,
      };
      await this.postIocRuntimeVariablesState();
      return;
    }

    this.iocRuntimeVariablesPanelState = {
      ...panelState,
      message: "Loading IOC runtime variables...",
      isLoading: true,
    };
    await this.postIocRuntimeVariablesState();

    const variables = await this.fetchIocRuntimeVariablesForTerminal(terminal);
    const activeState = this.iocRuntimeVariablesPanelState;
    if (
      !activeState ||
      activeState.terminal !== terminal ||
      activeState.startupDocumentPath !== panelState.startupDocumentPath
    ) {
      return;
    }

    this.iocRuntimeVariablesPanelState = {
      ...activeState,
      variables,
      message: variables.length
        ? `Loaded ${variables.length} IOC runtime variables from var output.`
        : "Could not read IOC runtime variables from the running terminal.",
      isLoading: false,
    };
    if (this.iocRuntimeVariablesPanel) {
      this.iocRuntimeVariablesPanel.title = `IOC Runtime Variables: ${path.basename(activeState.startupDocumentPath || "")}`;
    }
    await this.postIocRuntimeVariablesState();
  }

  async reloadIocRuntimeEnvironmentPanel(panelState = this.iocRuntimeEnvironmentPanelState) {
    if (!panelState) {
      return;
    }

    const terminal = panelState.terminal;
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      this.iocRuntimeEnvironmentPanelState = {
        ...panelState,
        entries: [],
        message: "The tracked IOC terminal is no longer running.",
        isLoading: false,
      };
      await this.postIocRuntimeEnvironmentState();
      return;
    }

    this.iocRuntimeEnvironmentPanelState = {
      ...panelState,
      message: "Loading IOC runtime environment...",
      isLoading: true,
    };
    await this.postIocRuntimeEnvironmentState();

    const entries = await this.fetchIocRuntimeEnvironmentForTerminal(terminal);
    const activeState = this.iocRuntimeEnvironmentPanelState;
    if (
      !activeState ||
      activeState.terminal !== terminal ||
      activeState.startupDocumentPath !== panelState.startupDocumentPath
    ) {
      return;
    }

    this.iocRuntimeEnvironmentPanelState = {
      ...activeState,
      entries,
      message: entries.length
        ? `Loaded ${entries.length} IOC runtime environment values from epicsEnvShow output.`
        : "Could not read IOC runtime environment values from the running terminal.",
      isLoading: false,
    };
    if (this.iocRuntimeEnvironmentPanel) {
      this.iocRuntimeEnvironmentPanel.title = `IOC Runtime Environment: ${path.basename(activeState.startupDocumentPath || "")}`;
    }
    await this.postIocRuntimeEnvironmentState();
  }

  async openActiveStartupIocCommands() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open a running IOC startup file under iocBoot to view IOC runtime commands.",
      );
      return;
    }

    const terminal = this.iocStartupTerminalByDocumentPath.get(startupTarget.documentPath);
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      vscode.window.showWarningMessage(
        `No tracked IOC terminal is running for ${path.basename(startupTarget.documentPath)}.`,
      );
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      return;
    }

    if (!this.iocRuntimeCommandsPanel) {
      const panel = vscode.window.createWebviewPanel(
        IOC_RUNTIME_COMMANDS_VIEW_TYPE,
        `IOC Runtime Commands: ${path.basename(startupTarget.documentPath)}`,
        vscode.ViewColumn.Active,
        {
          enableScripts: true,
          retainContextWhenHidden: true,
        },
      );
      this.iocRuntimeCommandsPanel = panel;
      panel.webview.html = buildIocRuntimeCommandsHtml(
        panel.webview,
        this.buildIocRuntimeCommandsState({
          startupDocumentPath: startupTarget.documentPath,
          terminal,
          commands: [],
          message: "Loading IOC runtime commands...",
          isLoading: true,
        }),
      );
      panel.onDidDispose(() => {
        if (this.iocRuntimeCommandsPanel === panel) {
          this.iocRuntimeCommandsPanel = undefined;
          this.iocRuntimeCommandsPanelState = undefined;
        }
      });
      panel.webview.onDidReceiveMessage(async (message) => {
        if (!message?.type) {
          return;
        }

        if (message.type === "showIocRuntimeCommandsTerminal") {
          const panelState = this.iocRuntimeCommandsPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeCommandsState({
              ...this.buildIocRuntimeCommandsState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.show(true);
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "stopIocRuntimeCommandsStartup") {
          const panelState = this.iocRuntimeCommandsPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeCommandsState({
              ...this.buildIocRuntimeCommandsState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.sendText("\u0003", false);
          this.iocRuntimeCommandsPanelState = {
            ...panelState,
            message: "Stopping IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeCommandsState();
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "startIocRuntimeCommandsStartup") {
          const panelState = this.iocRuntimeCommandsPanelState;
          const startupDocumentPath = panelState?.startupDocumentPath;
          if (!startupDocumentPath) {
            vscode.window.showWarningMessage(
              "No startup file is associated with this IOC commands page.",
            );
            return;
          }

          this.iocRuntimeCommandsPanelState = {
            ...panelState,
            message: "Starting IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeCommandsState();
          await this.startStartupIocForDocumentPath(startupDocumentPath, {
            showTerminal: false,
            notify: true,
          });
          return;
        }

        if (message.type === "requestIocRuntimeCommandHelp" && message.commandName) {
          const panelState = this.iocRuntimeCommandsPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            return;
          }

          const commandList = Array.isArray(panelState?.commands) ? panelState.commands : [];
          const hoveredCommandIndex = commandList.indexOf(message.commandName);
          const requestedCommandNames =
            hoveredCommandIndex >= 0
              ? commandList.slice(
                  hoveredCommandIndex,
                  Math.min(hoveredCommandIndex + 8, commandList.length),
                )
              : [message.commandName];
          const helpMap = await this.fetchIocRuntimeCommandHelpBatch(
            requestedCommandNames,
            targetTerminal,
          );
          const helpText = helpMap.get(message.commandName);
          if (!this.iocRuntimeCommandsPanel?.webview) {
            return;
          }
          for (const [commandName, commandHelpText] of helpMap.entries()) {
            await this.iocRuntimeCommandsPanel.webview.postMessage({
              type: "iocRuntimeCommandHelp",
              commandName,
              helpText:
                commandHelpText ||
                `No detailed help is available for ${commandName}.`,
            });
          }
          return;
        }

        if (message.type === "sendIocRuntimeCommand" && message.commandName) {
          const panelState = this.iocRuntimeCommandsPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeCommandsState({
              ...this.buildIocRuntimeCommandsState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          const argumentText = String(message.argumentsText || "").trim();
          const fullCommand = `${message.commandName}(${argumentText})`;
          if (message.captureOutput) {
            const outputText = await this.runIocShellCommandWithCapturedOutput(
              fullCommand,
              targetTerminal,
            );
            if (outputText === undefined) {
              vscode.window.showWarningMessage(
                `Could not capture IOC output for ${fullCommand}.`,
              );
            }
            return;
          }

          await this.sendIocShellCommand(fullCommand, targetTerminal);
          return;
        }

        if (message.type === "sendCustomIocRuntimeCommand") {
          const panelState = this.iocRuntimeCommandsPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeCommandsState({
              ...this.buildIocRuntimeCommandsState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          const commandText = String(message.commandText || "").trim();
          if (!commandText) {
            return;
          }
          if (message.captureOutput) {
            const outputText = await this.runIocShellCommandWithCapturedOutput(
              commandText,
              targetTerminal,
            );
            if (outputText === undefined) {
              vscode.window.showWarningMessage(
                `Could not capture IOC output for ${commandText}.`,
              );
            }
            return;
          }
          await this.sendIocShellCommand(commandText, targetTerminal);
        }
      });
    } else {
      this.iocRuntimeCommandsPanel.reveal(vscode.ViewColumn.Active, false);
      this.iocRuntimeCommandsPanel.title = `IOC Runtime Commands: ${path.basename(startupTarget.documentPath)}`;
    }

    this.iocRuntimeCommandsPanelState = {
      startupDocumentPath: startupTarget.documentPath,
      terminal,
      commands: [],
      message: "Loading IOC runtime commands...",
      isLoading: true,
    };
    await this.postIocRuntimeCommandsState();

    const commands = await this.fetchIocRuntimeCommandsForTerminal(terminal);
    this.iocRuntimeCommandsPanelState = {
      startupDocumentPath: startupTarget.documentPath,
      terminal,
      commands,
      message: commands.length
        ? `Loaded ${commands.length} IOC runtime commands from help output.`
        : "Could not read IOC command names from the running terminal.",
      isLoading: false,
    };
    this.iocRuntimeCommandsPanel.title = `IOC Runtime Commands: ${path.basename(startupTarget.documentPath)}`;
    await this.postIocRuntimeCommandsState();
    void this.prefetchIocRuntimeCommandHelpForPanel(
      commands,
      terminal,
      startupTarget.documentPath,
    );
  }

  async openActiveStartupIocVariables() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open a running IOC startup file under iocBoot to view IOC runtime variables.",
      );
      return;
    }

    const terminal = this.iocStartupTerminalByDocumentPath.get(startupTarget.documentPath);
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      vscode.window.showWarningMessage(
        `No tracked IOC terminal is running for ${path.basename(startupTarget.documentPath)}.`,
      );
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      return;
    }

    if (!this.iocRuntimeVariablesPanel) {
      const panel = vscode.window.createWebviewPanel(
        IOC_RUNTIME_VARIABLES_VIEW_TYPE,
        `IOC Runtime Variables: ${path.basename(startupTarget.documentPath)}`,
        vscode.ViewColumn.Active,
        {
          enableScripts: true,
          retainContextWhenHidden: true,
        },
      );
      this.iocRuntimeVariablesPanel = panel;
      panel.webview.html = buildIocRuntimeVariablesHtml(
        panel.webview,
        this.buildIocRuntimeVariablesState({
          startupDocumentPath: startupTarget.documentPath,
          terminal,
          variables: [],
          message: "Loading IOC runtime variables...",
          isLoading: true,
        }),
      );
      panel.onDidDispose(() => {
        if (this.iocRuntimeVariablesPanel === panel) {
          this.iocRuntimeVariablesPanel = undefined;
          this.iocRuntimeVariablesPanelState = undefined;
        }
      });
      panel.webview.onDidReceiveMessage(async (message) => {
        if (!message?.type) {
          return;
        }

        if (message.type === "showIocRuntimeVariablesTerminal") {
          const panelState = this.iocRuntimeVariablesPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeVariablesState({
              ...this.buildIocRuntimeVariablesState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.show(true);
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "stopIocRuntimeVariablesStartup") {
          const panelState = this.iocRuntimeVariablesPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeVariablesState({
              ...this.buildIocRuntimeVariablesState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.sendText("\u0003", false);
          this.iocRuntimeVariablesPanelState = {
            ...panelState,
            message: "Stopping IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeVariablesState();
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "startIocRuntimeVariablesStartup") {
          const panelState = this.iocRuntimeVariablesPanelState;
          const startupDocumentPath = panelState?.startupDocumentPath;
          if (!startupDocumentPath) {
            vscode.window.showWarningMessage(
              "No startup file is associated with this IOC variables page.",
            );
            return;
          }

          this.iocRuntimeVariablesPanelState = {
            ...panelState,
            message: "Starting IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeVariablesState();
          const startedTerminal = await this.startStartupIocForDocumentPath(
            startupDocumentPath,
            {
              showTerminal: false,
              notify: true,
            },
          );
          if (startedTerminal) {
            this.scheduleIocRuntimeVariablesPanelReload(1500);
          }
          return;
        }

        if (message.type === "setIocRuntimeVariable" && message.variableName) {
          const panelState = this.iocRuntimeVariablesPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeVariablesState({
              ...this.buildIocRuntimeVariablesState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          const variableName = String(message.variableName || "").trim();
          const valueText = String(message.valueText || "").trim();
          if (!variableName || !valueText) {
            vscode.window.showWarningMessage(
              `Enter a value before setting ${variableName || "the IOC variable"}.`,
            );
            return;
          }

          await this.sendIocShellCommand(`var ${variableName} ${valueText}`, targetTerminal);
          const cachedVariables = this.iocRuntimeVariablesByTerminal.get(targetTerminal);
          if (Array.isArray(cachedVariables)) {
            this.iocRuntimeVariablesByTerminal.set(
              targetTerminal,
              cachedVariables.map((entry) =>
                entry.variableName === variableName
                  ? { ...entry, currentValue: valueText }
                  : entry,
              ),
            );
          }
          if (panelState && Array.isArray(panelState.variables)) {
            this.iocRuntimeVariablesPanelState = {
              ...panelState,
              variables: panelState.variables.map((entry) =>
                entry.variableName === variableName
                  ? { ...entry, currentValue: valueText }
                  : entry,
              ),
              message: `Sent var ${variableName} ${valueText}`,
              isLoading: false,
            };
            await this.postIocRuntimeVariablesState();
          }
        }
      });
    } else {
      this.iocRuntimeVariablesPanel.reveal(vscode.ViewColumn.Active, false);
      this.iocRuntimeVariablesPanel.title = `IOC Runtime Variables: ${path.basename(startupTarget.documentPath)}`;
    }

    this.iocRuntimeVariablesPanelState = {
      startupDocumentPath: startupTarget.documentPath,
      terminal,
      variables: [],
      message: "Loading IOC runtime variables...",
      isLoading: true,
    };
    await this.postIocRuntimeVariablesState();
    await this.reloadIocRuntimeVariablesPanel();
  }

  async openActiveStartupIocEnvironment() {
    const startupTarget = this.resolveActiveStartupIocDocument();
    if (!startupTarget) {
      vscode.window.showWarningMessage(
        "Open a running IOC startup file under iocBoot to view IOC runtime environment values.",
      );
      return;
    }

    const terminal = this.iocStartupTerminalByDocumentPath.get(startupTarget.documentPath);
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      vscode.window.showWarningMessage(
        `No tracked IOC terminal is running for ${path.basename(startupTarget.documentPath)}.`,
      );
      this.updateStartupEditorContextKeys(vscode.window.activeTextEditor);
      return;
    }

    if (!this.iocRuntimeEnvironmentPanel) {
      const panel = vscode.window.createWebviewPanel(
        IOC_RUNTIME_ENVIRONMENT_VIEW_TYPE,
        `IOC Runtime Environment: ${path.basename(startupTarget.documentPath)}`,
        vscode.ViewColumn.Active,
        {
          enableScripts: true,
          retainContextWhenHidden: true,
        },
      );
      this.iocRuntimeEnvironmentPanel = panel;
      panel.webview.html = buildIocRuntimeEnvironmentHtml(
        panel.webview,
        this.buildIocRuntimeEnvironmentState({
          startupDocumentPath: startupTarget.documentPath,
          terminal,
          entries: [],
          message: "Loading IOC runtime environment...",
          isLoading: true,
        }),
      );
      panel.onDidDispose(() => {
        if (this.iocRuntimeEnvironmentPanel === panel) {
          this.iocRuntimeEnvironmentPanel = undefined;
          this.iocRuntimeEnvironmentPanelState = undefined;
        }
      });
      panel.webview.onDidReceiveMessage(async (message) => {
        if (!message?.type) {
          return;
        }

        if (message.type === "showIocRuntimeEnvironmentTerminal") {
          const panelState = this.iocRuntimeEnvironmentPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeEnvironmentState({
              ...this.buildIocRuntimeEnvironmentState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.show(true);
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "stopIocRuntimeEnvironmentStartup") {
          const panelState = this.iocRuntimeEnvironmentPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeEnvironmentState({
              ...this.buildIocRuntimeEnvironmentState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          this.iocShellTerminal = targetTerminal;
          targetTerminal.sendText("\u0003", false);
          this.iocRuntimeEnvironmentPanelState = {
            ...panelState,
            message: "Stopping IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeEnvironmentState();
          this.refresh(this.contextNode);
          return;
        }

        if (message.type === "startIocRuntimeEnvironmentStartup") {
          const panelState = this.iocRuntimeEnvironmentPanelState;
          const startupDocumentPath = panelState?.startupDocumentPath;
          if (!startupDocumentPath) {
            vscode.window.showWarningMessage(
              "No startup file is associated with this IOC runtime environment page.",
            );
            return;
          }

          this.iocRuntimeEnvironmentPanelState = {
            ...panelState,
            message: "Starting IOC...",
            isLoading: false,
          };
          await this.postIocRuntimeEnvironmentState();
          const startedTerminal = await this.startStartupIocForDocumentPath(
            startupDocumentPath,
            {
              showTerminal: false,
              notify: true,
            },
          );
          if (startedTerminal) {
            this.scheduleIocRuntimeEnvironmentPanelReload(1500);
          }
          return;
        }

        if (message.type === "setIocRuntimeEnvironmentValue" && message.variableName) {
          const panelState = this.iocRuntimeEnvironmentPanelState;
          const targetTerminal = panelState?.terminal;
          if (!targetTerminal || !vscode.window.terminals.includes(targetTerminal)) {
            vscode.window.showWarningMessage(
              "The tracked IOC terminal is no longer running.",
            );
            await this.postIocRuntimeEnvironmentState({
              ...this.buildIocRuntimeEnvironmentState(panelState),
              message: "The tracked IOC terminal is no longer running.",
            });
            return;
          }

          const variableName = String(message.variableName || "").trim();
          const valueText = String(message.valueText || "");
          if (!variableName) {
            return;
          }

          await this.sendIocShellCommand(
            `epicsEnvSet ${variableName} ${this.formatIocShellQuotedString(valueText)}`,
            targetTerminal,
          );
          const cachedEntries = this.iocRuntimeEnvironmentByTerminal.get(targetTerminal);
          if (Array.isArray(cachedEntries)) {
            this.iocRuntimeEnvironmentByTerminal.set(
              targetTerminal,
              cachedEntries.map((entry) =>
                entry.variableName === variableName
                  ? { ...entry, currentValue: valueText }
                  : entry,
              ),
            );
          }
          if (panelState && Array.isArray(panelState.entries)) {
            this.iocRuntimeEnvironmentPanelState = {
              ...panelState,
              entries: panelState.entries.map((entry) =>
                entry.variableName === variableName
                  ? { ...entry, currentValue: valueText }
                  : entry,
              ),
              message: `Sent epicsEnvSet ${variableName} ${this.formatIocShellQuotedString(valueText)}`,
              isLoading: false,
            };
            await this.postIocRuntimeEnvironmentState();
          }
        }
      });
    } else {
      this.iocRuntimeEnvironmentPanel.reveal(vscode.ViewColumn.Active, false);
      this.iocRuntimeEnvironmentPanel.title = `IOC Runtime Environment: ${path.basename(startupTarget.documentPath)}`;
    }

    this.iocRuntimeEnvironmentPanelState = {
      startupDocumentPath: startupTarget.documentPath,
      terminal,
      entries: [],
      message: "Loading IOC runtime environment...",
      isLoading: true,
    };
    await this.postIocRuntimeEnvironmentState();
    await this.reloadIocRuntimeEnvironmentPanel();
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

  attachStartupRunningDecorationType(decorationType) {
    this.startupRunningDecorationType = decorationType;
  }

  attachProbeDecorationTypes(decorationTypes) {
    this.probeDecorationTypes = Array.isArray(decorationTypes)
      ? decorationTypes
      : [];
  }

  async resolveProbeCustomEditor(document, webviewPanel) {
    if (!document?.uri || !webviewPanel?.webview) {
      return;
    }

    const sourceUri = document.uri.toString();
    let probeWebviewState = this.probeWebviews.get(sourceUri);
    if (!probeWebviewState) {
      probeWebviewState = {
        document,
        panels: new Set(),
      };
      this.probeWebviews.set(sourceUri, probeWebviewState);
    } else {
      probeWebviewState.document = document;
    }
    probeWebviewState.panels.add(webviewPanel);

    webviewPanel.webview.options = {
      enableScripts: true,
    };
    webviewPanel.webview.html = buildProbeCustomEditorHtml(
      webviewPanel.webview,
      this.buildProbeWebviewState(document),
    );

    webviewPanel.onDidDispose(() => {
      const currentState = this.probeWebviews.get(sourceUri);
      if (!currentState) {
        return;
      }
      currentState.panels.delete(webviewPanel);
      if (!currentState.panels.size) {
        this.probeWebviews.delete(sourceUri);
      }
    });

    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      if (!message?.type) {
        return;
      }

      if (message.type === "startProbeRuntime") {
        await this.startProbeDocumentRuntime(document);
        return;
      }

      if (message.type === "stopProbeRuntime") {
        await this.stopProbeDocumentRuntime(document);
        return;
      }

      if (message.type === "putProbeValue" && message.key) {
        await this.putRuntimeValue({ key: message.key });
      }
    });

    if (
      this.contextStatus !== "stopped" &&
      !this.probeSessions.has(sourceUri) &&
      analyzeProbeDocument(document).recordName
    ) {
      void this.startProbeDocumentRuntime(document);
    }

    void this.postProbeWebviewState(document, webviewPanel);
  }

  async openProbeWidget(options = {}) {
    const recordName = String(options?.recordName || "").trim();
    const widgetId = createNonce();
    const sourceUri = `probe-widget:${widgetId}`;
    const panel = vscode.window.createWebviewPanel(
      PROBE_WIDGET_VIEW_TYPE,
      recordName ? `Probe: ${recordName}` : "EPICS Probe",
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );

    const widgetState = {
      widgetId,
      sourceUri,
      recordName,
      panel,
    };
    this.probeWidgets.set(sourceUri, widgetState);
    this.registerWidgetMenuContext(panel, "probe", widgetState);

    panel.webview.html = buildProbeWidgetHtml(
      panel.webview,
      this.buildProbeWidgetWebviewState(widgetState),
    );

    panel.onDidDispose(() => {
      this.probeWidgets.delete(sourceUri);
      this.disposeProbeSession(sourceUri);
      void this.removeEntriesBySourceUri(sourceUri);
    });

    panel.webview.onDidReceiveMessage(async (message) => {
      if (!message?.type) {
        return;
      }

      if (message.type === "updateProbeWidgetRecordName") {
        await this.updateProbeWidgetRecordName(widgetState, message.recordName);
        return;
      }

      if (message.type === "putProbeValue" && message.key) {
        await this.putRuntimeValue({ key: message.key });
        return;
      }

      if (message.type === "processProbeWidget") {
        await this.processProbeWidget(widgetState);
      }
    });

    if (recordName) {
      await this.startProbeWidgetSession(widgetState);
    } else {
      await this.postProbeWidgetState(widgetState);
    }
  }

  async updateProbeWidgetRecordName(widgetState, nextRecordName) {
    if (!widgetState?.sourceUri || !widgetState.panel) {
      return;
    }

    const recordName = String(nextRecordName || "").trim();
    widgetState.recordName = recordName;
    widgetState.panel.title = recordName ? `Probe: ${recordName}` : "EPICS Probe";

    this.disposeProbeSession(widgetState.sourceUri);
    await this.removeEntriesBySourceUri(widgetState.sourceUri);

    if (!recordName) {
      await this.postProbeWidgetState(widgetState);
      return;
    }

    await this.startProbeWidgetSession(widgetState);
  }

  async startProbeWidgetSession(widgetState) {
    if (!widgetState?.sourceUri || !widgetState.recordName) {
      return;
    }

    await this.startProbeRuntimeSession({
      sourceUri: widgetState.sourceUri,
      sourceLabel: widgetState.panel?.title || "EPICS Probe",
      recordName: widgetState.recordName,
      progressTitle: `Starting EPICS probe for ${widgetState.recordName}`,
      showProgress: false,
    });
    await this.postProbeWidgetState(widgetState);
  }

  async openPvlistWidget(options = {}) {
    const widgetId = createNonce();
    const sourceUri = `pvlist-widget:${widgetId}`;
    const sourceModel = buildPvlistWidgetSourceModel(
      options,
      this.extractRecordDeclarations,
      this.extractDatabaseTocMacroAssignments,
    );
    if (sourceModel?.sourceKind === "pvlist" && (sourceModel.diagnostics || []).length > 0) {
      vscode.window.showErrorMessage(
        `Cannot open PvList widget: ${sourceModel.diagnostics[0]?.message || "invalid .pvlist format."}`,
      );
      return;
    }
    const panel = vscode.window.createWebviewPanel(
      PVLIST_WIDGET_VIEW_TYPE,
      sourceModel?.sourceLabel ? `PvList: ${sourceModel.sourceLabel}` : "EPICS PvList",
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );

    const widgetState = {
      widgetId,
      sourceUri,
      panel,
      sourceModel,
      macroValues: new Map(sourceModel?.macroValues || []),
      rows: [],
    };
    widgetState.rows = buildPvlistWidgetMonitorPlan(
      widgetState.sourceModel,
      widgetState.macroValues,
      this.getDefaultProtocol(),
    ).rows;
    this.pvlistWidgets.set(sourceUri, widgetState);
    this.registerWidgetMenuContext(panel, "pvlist", widgetState);

    panel.webview.html = buildPvlistWidgetHtml(
      panel.webview,
      this.buildPvlistWidgetWebviewState(widgetState),
    );

    panel.onDidDispose(() => {
      this.pvlistWidgets.delete(sourceUri);
      void this.removeEntriesBySourceUri(sourceUri);
    });

    panel.webview.onDidReceiveMessage(async (message) => {
      if (!message?.type) {
        return;
      }

      if (message.type === "replacePvlistWidgetChannels") {
        await this.replacePvlistWidgetChannels(widgetState, message.text);
        return;
      }

      if (message.type === "savePvlistWidget") {
        await this.savePvlistWidget(widgetState);
        return;
      }

      if (message.type === "updatePvlistWidgetMacro" && message.name) {
        await this.updatePvlistWidgetMacro(widgetState, message.name, message.value);
        return;
      }

      if (message.type === "putPvlistValue" && message.key) {
        await this.putRuntimeValueFromInput({ key: message.key }, message.value);
      }
    });

    await this.applyPvlistWidgetMonitoring(widgetState);
  }

  async openMonitorWidget(options = {}) {
    const widgetId = createNonce();
    const sourceUri = `monitor-widget:${widgetId}`;
    const initialChannels = Array.isArray(options?.initialChannels)
      ? options.initialChannels
      : [];
    const panel = vscode.window.createWebviewPanel(
      MONITOR_WIDGET_VIEW_TYPE,
      getMonitorWidgetPanelTitle(initialChannels),
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
      },
    );

    const widgetState = {
      widgetId,
      sourceUri,
      panel,
      sourceLabel: String(options?.sourceLabel || "EPICS Monitor"),
      channelRows: [],
      historyLines: [],
      bufferSize: DEFAULT_MONITOR_WIDGET_BUFFER_SIZE,
      sessionsByRowId: new Map(),
      lastError: "",
    };
    this.monitorWidgets.set(sourceUri, widgetState);
    this.registerWidgetMenuContext(panel, "monitor", widgetState);

    for (const channelName of initialChannels) {
      this.addMonitorWidgetRow(widgetState, channelName);
    }
    if (!widgetState.channelRows.length) {
      this.addMonitorWidgetRow(widgetState, "");
    }

    panel.webview.html = buildMonitorWidgetHtml(
      panel.webview,
      this.buildMonitorWidgetWebviewState(widgetState),
    );

    panel.onDidDispose(() => {
      this.monitorWidgets.delete(sourceUri);
      void this.disposeMonitorWidget(widgetState);
    });

    panel.webview.onDidReceiveMessage(async (message) => {
      if (!message?.type) {
        return;
      }

      if (message.type === "addMonitorWidgetChannelRow") {
        this.addMonitorWidgetRow(widgetState, "");
        await this.postMonitorWidgetState(widgetState);
        return;
      }

      if (message.type === "updateMonitorWidgetChannel" && message.rowId) {
        await this.updateMonitorWidgetChannel(
          widgetState,
          String(message.rowId),
          message.channelName,
        );
        return;
      }

      if (message.type === "updateMonitorWidgetBufferSize") {
        await this.updateMonitorWidgetBufferSize(widgetState, message.value);
        return;
      }

      if (message.type === "exportMonitorWidgetData") {
        await this.exportMonitorWidgetData(widgetState);
      }
    });

    await this.startMonitorWidgetRows(widgetState);
  }

  addMonitorWidgetRow(widgetState, initialChannelName = "") {
    const row = {
      id: createNonce(),
      channelName: String(initialChannelName || "").trim(),
      connectionAttemptId: 0,
      status: String(initialChannelName || "").trim() ? "connecting" : "idle",
      lastError: "",
    };
    widgetState.channelRows.push(row);
    this.updateMonitorWidgetPanelTitle(widgetState);
    return row;
  }

  async updateMonitorWidgetChannel(widgetState, rowId, nextChannelName) {
    const row = widgetState?.channelRows?.find((candidate) => candidate.id === rowId);
    if (!row) {
      return;
    }

    const normalizedChannelName = String(nextChannelName || "").trim();
    if (row.channelName === normalizedChannelName) {
      return;
    }

    await this.stopMonitorWidgetRowSession(widgetState, row);
    row.channelName = normalizedChannelName;
    row.lastError = "";
    row.status = normalizedChannelName ? "connecting" : "idle";
    this.updateMonitorWidgetPanelTitle(widgetState);
    await this.postMonitorWidgetState(widgetState);

    if (normalizedChannelName) {
      await this.startMonitorWidgetRow(widgetState, row);
    }
  }

  async updateMonitorWidgetBufferSize(widgetState, nextValue) {
    const parsed = Number.parseInt(String(nextValue || "").trim(), 10);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      vscode.window.showErrorMessage("Buffer size must be a positive integer.");
      await this.postMonitorWidgetState(widgetState);
      return;
    }

    widgetState.bufferSize = parsed;
    trimMonitorWidgetHistory(widgetState);
    await this.postMonitorWidgetState(widgetState);
  }

  async exportMonitorWidgetData(widgetState) {
    const targetUri = await vscode.window.showSaveDialog({
      defaultUri: getDefaultMonitorWidgetSaveUri(widgetState),
      filters: {
        "Text": ["txt"],
      },
      saveLabel: "Export Monitor Data",
    });
    if (!targetUri) {
      return;
    }

    const fileText = buildMonitorWidgetExportText(widgetState);
    await vscode.workspace.fs.writeFile(targetUri, Buffer.from(fileText, "utf8"));
    vscode.window.showInformationMessage(
      `Exported monitor data to ${path.basename(targetUri.fsPath || targetUri.path || "monitor-data.txt")}.`,
    );
  }

  async startMonitorWidgetRows(widgetState, runtimeContext) {
    if (!widgetState) {
      return;
    }

    const activeRows = widgetState.channelRows.filter((row) => String(row.channelName || "").trim());
    if (!activeRows.length) {
      await this.postMonitorWidgetState(widgetState);
      return;
    }

    let activeRuntimeContext = runtimeContext;
    try {
      activeRuntimeContext = activeRuntimeContext || await this.ensureRuntimeContext();
      widgetState.lastError = "";
    } catch (error) {
      widgetState.lastError = getErrorMessage(error);
      await this.postMonitorWidgetState(widgetState);
      return;
    }

    await Promise.allSettled(
      activeRows.map((row) => this.startMonitorWidgetRow(widgetState, row, activeRuntimeContext)),
    );
    await this.postMonitorWidgetState(widgetState);
  }

  async startMonitorWidgetRow(widgetState, row, runtimeContext) {
    if (!widgetState || !row) {
      return;
    }

    const rawChannelName = String(row.channelName || "").trim();
    if (!rawChannelName) {
      return;
    }

    const attemptId = Number(row.connectionAttemptId || 0) + 1;
    row.connectionAttemptId = attemptId;
    row.status = "connecting";
    row.lastError = "";
    void this.postMonitorWidgetState(widgetState);

    const definition = parseMonitorWidgetChannelReference(
      rawChannelName,
      this.getDefaultProtocol(),
    );
    let session = undefined;

    try {
      const activeRuntimeContext = runtimeContext || await this.ensureRuntimeContext();
      if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
        return;
      }

      const channel = await activeRuntimeContext.createChannel(
        definition.pvName,
        definition.protocol,
        this.getChannelCreationTimeoutSeconds(),
      );
      if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
        try {
          await channel?.destroyHard?.();
        } catch (error) {
          // Ignore cleanup failures for superseded monitor widget channel attempts.
        }
        return;
      }
      if (!channel) {
        throw new Error(
          `Failed to create ${definition.protocol.toUpperCase()} channel for ${definition.pvName}.`,
        );
      }

      session = {
        rowId: row.id,
        rawChannelName,
        pvName: definition.pvName,
        protocol: definition.protocol,
        pvRequest: definition.pvRequest,
        channel,
        monitor: undefined,
        caEnumChoices: undefined,
        pvaEnumChoices: undefined,
      };

      channel.setDestroySoftCallback(() => {
        if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
          return;
        }
        row.status = "disconnected";
        void this.postMonitorWidgetState(widgetState);
      });
      channel.setDestroyHardCallback(() => {
        if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
          return;
        }
        widgetState.sessionsByRowId.delete(row.id);
        row.status = row.channelName ? "stopped" : "idle";
        void this.postMonitorWidgetState(widgetState);
      });

      const monitor =
        definition.protocol === "pva"
          ? await channel.createMonitorPva(
            this.getMonitorSubscribeTimeoutSeconds(),
            definition.pvRequest,
            (activeMonitor) => {
              this.handleMonitorWidgetUpdate(widgetState, row, session, activeMonitor);
            },
          )
          : (() => {
            const caMonitorOptions = getCaEnumMonitorOptions(channel);
            if (caMonitorOptions) {
              return channel.createMonitor(
                this.getMonitorSubscribeTimeoutSeconds(),
                (activeMonitor) => {
                  this.handleMonitorWidgetUpdate(widgetState, row, session, activeMonitor);
                },
                caMonitorOptions.dbrType,
                caMonitorOptions.valueCount,
              );
            }

            return channel.createMonitor(
              this.getMonitorSubscribeTimeoutSeconds(),
              (activeMonitor) => {
                this.handleMonitorWidgetUpdate(widgetState, row, session, activeMonitor);
              },
            );
          })();

      if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
        await this.cleanupMonitorWidgetSession({
          ...session,
          monitor: await monitor,
        });
        return;
      }

      session.monitor = await monitor;
      widgetState.sessionsByRowId.set(row.id, session);
      row.status = getRuntimeMonitorState(session.monitor) === "SUBSCRIBED"
        ? "subscribed"
        : "connecting";
      void this.postMonitorWidgetState(widgetState);
    } catch (error) {
      if (!this.isCurrentMonitorWidgetAttempt(widgetState, row, attemptId)) {
        return;
      }
      row.status = "error";
      row.lastError = getErrorMessage(error);
      widgetState.sessionsByRowId.delete(row.id);
      await this.cleanupMonitorWidgetSession(session);
      void this.postMonitorWidgetState(widgetState);
    }
  }

  handleMonitorWidgetUpdate(widgetState, row, session, monitor) {
    if (!widgetState || !row || !session) {
      return;
    }

    const activeSession = widgetState.sessionsByRowId.get(row.id);
    if (activeSession && activeSession !== session) {
      return;
    }

    if (!activeSession && session.monitor) {
      widgetState.sessionsByRowId.set(row.id, session);
    }

    row.status = "subscribed";
    row.lastError = "";

    const runtimeLibrary = this.requireRuntimeLibrarySafe();
    const formattedLine = buildMonitorWidgetHistoryLine(session, monitor, runtimeLibrary);
    if (formattedLine) {
      widgetState.historyLines.push(formattedLine);
      trimMonitorWidgetHistory(widgetState);
    }

    void this.postMonitorWidgetState(widgetState);
  }

  async stopMonitorWidgetRowSession(widgetState, row) {
    if (!widgetState || !row) {
      return;
    }

    row.connectionAttemptId = Number(row.connectionAttemptId || 0) + 1;
    const session = widgetState.sessionsByRowId.get(row.id);
    widgetState.sessionsByRowId.delete(row.id);
    await this.cleanupMonitorWidgetSession(session);
    row.status = row.channelName ? "stopped" : "idle";
  }

  stopMonitorWidgetSessions(widgetState) {
    if (!widgetState) {
      return;
    }

    for (const row of widgetState.channelRows || []) {
      row.connectionAttemptId = Number(row.connectionAttemptId || 0) + 1;
      row.status = row.channelName ? "stopped" : "idle";
      row.lastError = "";
    }

    for (const session of widgetState.sessionsByRowId.values()) {
      try {
        session.monitor?.destroyHard?.();
      } catch (error) {
        // Ignore best-effort monitor cleanup failures while stopping the context.
      }
      try {
        session.channel?.destroyHard?.();
      } catch (error) {
        // Ignore best-effort channel cleanup failures while stopping the context.
      }
    }

    widgetState.sessionsByRowId.clear();
  }

  async restartMonitorWidgets(runtimeContext) {
    await Promise.allSettled(
      [...this.monitorWidgets.values()].map((widgetState) =>
        this.restartMonitorWidgetSessions(widgetState, runtimeContext)),
    );
  }

  async restartMonitorWidgetSessions(widgetState, runtimeContext) {
    this.stopMonitorWidgetSessions(widgetState);
    await this.startMonitorWidgetRows(widgetState, runtimeContext);
  }

  async disposeMonitorWidget(widgetState) {
    if (!widgetState) {
      return;
    }
    for (const row of widgetState.channelRows || []) {
      row.connectionAttemptId = Number(row.connectionAttemptId || 0) + 1;
    }
    const sessions = [...widgetState.sessionsByRowId.values()];
    widgetState.sessionsByRowId.clear();
    await Promise.allSettled(sessions.map((session) => this.cleanupMonitorWidgetSession(session)));
  }

  async cleanupMonitorWidgetSession(session) {
    if (!session) {
      return;
    }

    try {
      session.monitor?.destroyHard?.();
    } catch (error) {
      // Ignore monitor teardown errors while cleaning up widget sessions.
    }

    try {
      await session.channel?.destroyHard?.();
    } catch (error) {
      // Ignore channel teardown errors while cleaning up widget sessions.
    }
  }

  isCurrentMonitorWidgetAttempt(widgetState, row, attemptId) {
    return (
      Boolean(widgetState) &&
      this.monitorWidgets.get(widgetState.sourceUri) === widgetState &&
      Boolean(row) &&
      widgetState.channelRows.includes(row) &&
      Number(row.connectionAttemptId || 0) === Number(attemptId)
    );
  }

  updateMonitorWidgetPanelTitle(widgetState) {
    if (!widgetState?.panel) {
      return;
    }

    widgetState.panel.title = getMonitorWidgetPanelTitle(
      widgetState.channelRows.map((row) => row.channelName),
    );
  }

  async savePvlistWidget(widgetState) {
    const sourceModel = widgetState?.sourceModel;
    if (!sourceModel) {
      return;
    }

    const fileText = buildPvlistWidgetFileText(sourceModel, widgetState.macroValues);
    const currentUriText = sourceModel.sourceDocumentUri
      ? String(sourceModel.sourceDocumentUri)
      : "";
    let targetUri =
      sourceModel.sourceKind === "pvlist" && currentUriText
        ? safeParseUri(currentUriText)
        : undefined;

    if (!targetUri) {
      targetUri = await vscode.window.showSaveDialog({
        defaultUri: getDefaultPvlistWidgetSaveUri(sourceModel),
        filters: {
          "PV List": ["pvlist"],
        },
        saveLabel: "Save PvList",
      });
    }

    if (!targetUri) {
      return;
    }

    await vscode.workspace.fs.writeFile(targetUri, Buffer.from(fileText, "utf8"));
    sourceModel.sourceDocumentUri = targetUri.toString();
    sourceModel.sourceLabel = path.basename(targetUri.fsPath || targetUri.path || "EPICS PvList");
    if (widgetState.panel) {
      widgetState.panel.title = `PvList: ${sourceModel.sourceLabel}`;
    }
    await this.postPvlistWidgetState(widgetState);
    vscode.window.showInformationMessage(
      `Saved PvList to ${path.basename(targetUri.fsPath || targetUri.path || "pvlist")}.`,
    );
  }

  async replacePvlistWidgetChannels(widgetState, text) {
    const sourceModel = widgetState?.sourceModel;
    if (!sourceModel) {
      return;
    }

    const nextRawPvNames = parseAddedPvlistChannelLines(text);
    const previousRawPvNames = Array.isArray(sourceModel.rawPvNames) ? sourceModel.rawPvNames : [];
    const nextMacroNames = extractOrderedEpicsMacroNames(nextRawPvNames);
    const previousMacroNames = Array.isArray(sourceModel.macroNames) ? sourceModel.macroNames : [];
    const didChannelsChange =
      previousRawPvNames.length !== nextRawPvNames.length ||
      previousRawPvNames.some((entry, index) => String(entry || "") !== nextRawPvNames[index]);
    const didMacrosChange =
      previousMacroNames.length !== nextMacroNames.length ||
      previousMacroNames.some((entry, index) => String(entry || "") !== nextMacroNames[index]);

    if (!didChannelsChange && !didMacrosChange) {
      await this.postPvlistWidgetState(widgetState);
      return;
    }

    const previousMacroValues = widgetState?.macroValues instanceof Map
      ? widgetState.macroValues
      : new Map();
    sourceModel.rawPvNames = [...nextRawPvNames];
    sourceModel.macroNames = [...nextMacroNames];
    const nextMacroValues = new Map(
      nextMacroNames.map((macroName) => [macroName, previousMacroValues.get(macroName) || ""]),
    );
    sourceModel.macroValues = nextMacroValues;
    widgetState.macroValues = nextMacroValues;

    await this.applyPvlistWidgetMonitoring(widgetState);
  }

  async updatePvlistWidgetMacro(widgetState, macroName, macroValue) {
    if (!widgetState?.macroValues || !macroName) {
      return;
    }

    const normalizedMacroName = String(macroName);
    const normalizedMacroValue = String(macroValue ?? "");
    if ((widgetState.macroValues.get(normalizedMacroName) || "") === normalizedMacroValue) {
      return;
    }

    widgetState.macroValues.set(normalizedMacroName, normalizedMacroValue);
    await this.applyPvlistWidgetMonitoring(widgetState);
  }

  async applyPvlistWidgetMonitoring(widgetState) {
    if (!widgetState?.sourceUri) {
      return;
    }

    const plan = buildPvlistWidgetMonitorPlan(
      widgetState.sourceModel,
      widgetState.macroValues,
      this.getDefaultProtocol(),
    );
    widgetState.rows = plan.rows;

    const queuedEntries = await this.syncPvlistWidgetMonitorEntries(widgetState, plan);

    if (queuedEntries.length) {
      try {
        const runtimeContext = await this.ensureRuntimeContext();
        void this.connectEntriesInParallel(queuedEntries, runtimeContext);
      } catch (error) {
        if (!isContextInitializationCancelledError(error)) {
          // Keep the widget visible even when runtime context startup fails.
        }
      }
    }

    await this.postPvlistWidgetState(widgetState);
  }

  async syncPvlistWidgetMonitorEntries(widgetState, plan) {
    const sourceUri = widgetState?.sourceUri;
    if (!sourceUri) {
      return [];
    }

    const desiredDefinitions = (plan?.definitions || []).map((definition) => ({
      ...definition,
      sourceKind: "pvlistWidget",
      sourceUri,
      sourceLabel: widgetState.sourceModel?.sourceLabel || "EPICS PvList",
      hidden: true,
    }));
    const desiredKeys = new Set(desiredDefinitions.map((definition) => createMonitorKey(definition)));
    const existingEntries = this.monitorEntries.filter(
      (entry) => entry.sourceUri === sourceUri && entry.sourceKind === "pvlistWidget",
    );
    const removableEntries = existingEntries.filter((entry) => !desiredKeys.has(entry.key));

    if (removableEntries.length) {
      await this.removeSpecificMonitorEntries(removableEntries);
    }

    return desiredDefinitions
      .map((definition) => this.queueMonitorEntry(definition, { hidden: true }))
      .filter((entry) => Boolean(entry) && !entry.channel && !entry.monitor);
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
      this.updateDatabaseEditorContextKeys(editor);
      this.updateStartupEditorContextKeys(editor);
      return;
    }

    const document = editor?.document;
    this.updateDatabaseEditorContextKeys(editor);
    this.updateStartupEditorContextKeys(editor);
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

    if (!isStrictMonitorDocument(document) && !isProbeDocument(document)) {
      this.monitorDiagnostics.delete(document.uri);
      return;
    }

    const analysis = isProbeDocument(document)
      ? analyzeProbeDocument(document)
      : analyzeStrictMonitorDocument(
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
    if (
      !this.hoverDecorationType &&
      !this.databaseTocValueDecorationType &&
      !this.startupRunningDecorationType &&
      !this.probeDecorationTypes.length
    ) {
      return;
    }

    this.reconcileMonitorStates();

    for (const editor of vscode.window.visibleTextEditors) {
      this.refreshMonitorHoverDecorationsForEditor(editor);
      this.refreshDatabaseTocValueDecorationsForEditor(editor);
      this.refreshStartupRunningDecorationsForEditor(editor);
      this.refreshProbeDecorationsForEditor(editor);
    }
    this.refreshProbePanels();
  }

  reconcileMonitorStates() {
    for (const entry of this.monitorEntries) {
      if (entry.sourceKind === "probe") {
        continue;
      }

      const channelState = getRuntimeChannelState(entry.channel);
      const monitorState = getRuntimeMonitorState(entry.monitor);
      const hasLiveValue =
        entry.status === "subscribed" &&
        Boolean(entry.monitor) &&
        entry.lastUpdated instanceof Date;
      let didChange = false;
      let shouldRecover = false;

      if (monitorState === "SUBSCRIBED" && entry.monitor) {
        if (entry.status !== "subscribed" || entry.lastError) {
          entry.status = "subscribed";
          entry.lastError = undefined;
          entry.hasEverSubscribed = true;
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
        if (!hasLiveValue && (entry.status !== "connecting" || entry.lastError)) {
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
        if (!hasLiveValue && (entry.status !== "connecting" || entry.lastError)) {
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
      if (
        shouldRecover &&
        (entry.sourceKind !== "probe" || !entry.hasEverSubscribed)
      ) {
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

  refreshStartupRunningDecorationsForDocument(document) {
    if (!this.startupRunningDecorationType || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        this.refreshStartupRunningDecorationsForEditor(editor);
      }
    }
  }

  clearStartupRunningDecorationsForDocument(document) {
    if (!this.startupRunningDecorationType || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        editor.setDecorations(this.startupRunningDecorationType, []);
      }
    }
  }

  refreshProbeDecorationsForDocument(document) {
    if (!isProbeDocument(document) || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        this.refreshProbeDecorationsForEditor(editor);
      }
    }
  }

  clearProbeDecorationsForDocument(document) {
    if (!this.probeDecorationTypes.length || !document?.uri) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      if (editor.document.uri.toString() === document.uri.toString()) {
        this.clearProbeDecorationsForEditor(editor);
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

  refreshStartupRunningDecorationsForEditor(editor) {
    if (!this.startupRunningDecorationType || !editor) {
      return;
    }

    const startupTarget = this.resolveActiveStartupIocDocument(editor);
    if (!startupTarget) {
      editor.setDecorations(this.startupRunningDecorationType, []);
      return;
    }

    const terminal = this.iocStartupTerminalByDocumentPath.get(startupTarget.documentPath);
    if (!terminal || !vscode.window.terminals.includes(terminal)) {
      editor.setDecorations(this.startupRunningDecorationType, []);
      return;
    }

    const decorationOptions = [];
    const lineCount = Math.max(Number(editor.document.lineCount) || 0, 1);
    for (
      let lineIndex = 0;
      lineIndex < lineCount;
      lineIndex += STARTUP_RUNNING_WATERMARK_LINE_INTERVAL
    ) {
      const line = editor.document.lineAt(Math.min(lineIndex, lineCount - 1));
      decorationOptions.push({
        range: new vscode.Range(line.range.start, line.range.start),
        renderOptions: {
          after: {
            contentText: "Running ...",
          },
        },
      });
    }

    editor.setDecorations(this.startupRunningDecorationType, decorationOptions);
  }

  refreshProbeDecorationsForEditor(editor) {
    if (!this.probeDecorationTypes.length || !editor) {
      return;
    }

    const document = editor.document;
    if (!isProbeDocument(document) || !document?.uri) {
      this.clearProbeDecorationsForEditor(editor);
      return;
    }

    const analysis = analyzeProbeDocument(document);
    if (analysis.diagnostics?.length || !analysis.recordName) {
      this.clearProbeDecorationsForEditor(editor);
      return;
    }

    const session = this.probeSessions.get(document.uri.toString());
    if (!session || this.contextStatus === "stopped") {
      this.clearProbeDecorationsForEditor(editor);
      return;
    }

    const state = this.buildProbePanelState(session);
    const anchorPosition = buildProbeOverlayAnchorPosition(document, analysis);
    if (!anchorPosition) {
      this.clearProbeDecorationsForEditor(editor);
      return;
    }

    const overlayLines = buildProbeOverlayLines(state);
    const range = new vscode.Range(anchorPosition, anchorPosition);
    const hoverMessage = new vscode.MarkdownString(
      buildProbeOverlayHoverMarkdown(state),
    );

    this.probeDecorationTypes.forEach((decorationType, index) => {
      const lineText = overlayLines[index];
      if (!lineText) {
        editor.setDecorations(decorationType, []);
        return;
      }

      editor.setDecorations(decorationType, [
        {
          range,
          hoverMessage,
          renderOptions: {
            after: {
              contentText: lineText,
              margin: "0 0 0 1ch",
              textDecoration: "none; display: block;",
            },
          },
        },
      ]);
    });
  }

  clearProbeDecorationsForEditor(editor) {
    if (!editor) {
      return;
    }

    for (const decorationType of this.probeDecorationTypes) {
      editor.setDecorations(decorationType, []);
    }
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

    this.disposeProbeSession(document.uri.toString());
    await this.removeEntriesBySourceUri(document.uri.toString());
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async handleDocumentChanged(event) {
    const document = event?.document;
    if (!document?.uri || !event?.contentChanges?.length) {
      return;
    }

    if (!this.hasEntriesForSourceUri(document.uri.toString())) {
      if (
        vscode.window.activeTextEditor?.document?.uri?.toString() ===
        document.uri.toString()
      ) {
        this.updateDatabaseEditorContextKeys(vscode.window.activeTextEditor);
      }
      return;
    }

    if (isProbeDocument(document)) {
      await this.startProbeDocumentRuntime(document);
      this.updateDatabaseEditorContextKeys(vscode.window.activeTextEditor);
      return;
    }

    await this.removeEntriesBySourceUri(document.uri.toString());
    this.updateDatabaseEditorContextKeys(vscode.window.activeTextEditor);
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
          await this.restartMonitorWidgets(runtimeContext);
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

  async startDatabaseMonitorChannels() {
    const document = vscode.window.activeTextEditor?.document;
    if (!isDatabaseRuntimeDocument(document)) {
      return;
    }

    if (!this.databaseDocumentHasToc(document)) {
      vscode.window.showWarningMessage(
        "Update Table of Contents before starting database channel monitoring.",
      );
      return;
    }

    await this.startActiveFileRuntimeContext();
  }

  async stopDatabaseMonitorChannels() {
    const document = vscode.window.activeTextEditor?.document;
    if (!isDatabaseRuntimeDocument(document)) {
      return;
    }

    await this.stopActiveFileRuntimeContext();
  }

  async startActiveFileRuntimeContext() {
    const editor = vscode.window.activeTextEditor;
    const document = editor?.document;
    if (!isRuntimeDocument(document)) {
      vscode.window.showWarningMessage(
        "Open a database, template, .pvlist, .probe, or plain text file to start file runtime monitoring.",
      );
      return;
    }

    const workspaceFolder =
      this.getWorkspaceFolderForDocument(document) || this.getDefaultWorkspaceFolder();
    if (workspaceFolder) {
      this.runtimeWorkspaceFolder = workspaceFolder;
    }

    if (isProbeDocument(document)) {
      await this.startProbeDocumentRuntime(document);
      this.handleActiveEditorChange(vscode.window.activeTextEditor);
      return;
    }

    const analysis = analyzeRuntimeDocument(
      document,
      this.getDefaultProtocol(),
      this.getDatabaseRuntimeHelpers(),
    );
    const sourceLabel = getRuntimeDocumentLabel(document);
    if (analysis.diagnostics?.length) {
      vscode.window.showErrorMessage(
        `Cannot start EPICS runtime for ${sourceLabel} until the file errors are fixed.`,
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

    if (isProbeDocument(document)) {
      this.disposeProbeSession(document.uri.toString());
    }
    await this.removeEntriesBySourceUri(document.uri.toString());
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async startProbeDocumentRuntime(document) {
    const sourceUri = document?.uri?.toString();
    if (!sourceUri) {
      return;
    }

    const analysis = analyzeProbeDocument(document);
    const sourceLabel = getRuntimeDocumentLabel(document);
    if (analysis.diagnostics?.length) {
      vscode.window.showErrorMessage(
        `Cannot start EPICS runtime for ${sourceLabel} until the probe file errors are fixed.`,
      );
      return;
    }
    if (!analysis.recordName) {
      vscode.window.showWarningMessage(
        `No EPICS probe target was found in ${sourceLabel}.`,
      );
      return;
    }

    await this.startProbeRuntimeSession({
      sourceUri,
      sourceLabel,
      recordName: analysis.recordName,
      progressTitle: `Starting EPICS probe for ${sourceLabel}`,
      showProgress: true,
    });
    this.refreshProbePanels();
  }

  async startProbeRuntimeSession({
    sourceUri,
    sourceLabel,
    recordName,
    progressTitle,
    showProgress = true,
  }) {
    if (!sourceUri || !recordName) {
      return undefined;
    }

    this.probeSessions.delete(sourceUri);
    await this.removeEntriesBySourceUri(sourceUri);

    const session = {
      sourceUri,
      sourceLabel,
      recordName,
      recordType: undefined,
      fieldEntriesStarted: false,
      fieldEntryStartInProgress: false,
    };
    this.probeSessions.set(sourceUri, session);

    const bootstrapDefinitions = [
      {
        pvName: recordName,
        protocol: this.getDefaultProtocol(),
        pvRequest: "",
        hidden: true,
        sourceKind: "probe",
        probeRole: "main",
      },
      {
        pvName: `${recordName}.RTYP`,
        protocol: this.getDefaultProtocol(),
        pvRequest: "",
        hidden: true,
        sourceKind: "probe",
        probeRole: "recordType",
        probeFieldName: "RTYP",
      },
    ];

    const startWork = async () => {
      const queuedEntries = bootstrapDefinitions
        .map((definition) =>
          this.queueMonitorEntry(
            {
              ...definition,
              sourceUri,
              sourceLabel,
              recordType: session.recordType,
            },
            {
              hidden: true,
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
    };

    if (showProgress) {
      await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Notification,
          title: progressTitle || `Starting EPICS probe for ${sourceLabel}`,
        },
        startWork,
      );
    } else {
      await startWork();
    }

    return session;
  }

  async stopProbeDocumentRuntime(document) {
    const sourceUri = document?.uri?.toString();
    if (!sourceUri) {
      return;
    }

    this.disposeProbeSession(sourceUri);
    await this.removeEntriesBySourceUri(sourceUri);
    this.refreshProbePanels();
  }

  disposeProbeSession(sourceUri) {
    if (!sourceUri) {
      return;
    }
    const session = this.probeSessions.get(sourceUri);
    if (!session) {
      return;
    }
    this.probeSessions.delete(sourceUri);
  }

  async ensureProbeFieldEntries(session) {
    if (!session || session.fieldEntriesStarted || session.fieldEntryStartInProgress) {
      return;
    }
    const recordType = this.getProbeResolvedRecordType(session);
    if (!recordType) {
      return;
    }
    session.recordType = recordType;
    session.fieldEntryStartInProgress = true;
    try {
      const fieldNames = (this.getFieldNamesForRecordType?.(recordType) || []).filter(
        (fieldName) => String(fieldName || "").toUpperCase() !== "RTYP",
      );
      const definitions = fieldNames.map((fieldName) => ({
        pvName: `${session.recordName}.${fieldName}`,
        protocol: this.getDefaultProtocol(),
        pvRequest: "",
        hidden: true,
        sourceKind: "probe",
        probeRole: "field",
        probeFieldName: fieldName,
        recordType,
        sourceUri: session.sourceUri,
        sourceLabel: session.sourceLabel,
      }));
      const queuedEntries = definitions
        .map((definition) => this.queueMonitorEntry(definition, { hidden: true }))
        .filter(Boolean);
      session.fieldEntriesStarted = true;
      if (queuedEntries.length) {
        const runtimeContext = await this.ensureRuntimeContext();
        void this.connectEntriesInBatches(queuedEntries, runtimeContext);
      }
    } finally {
      session.fieldEntryStartInProgress = false;
    }
  }

  getProbeResolvedRecordType(session) {
    if (session?.recordType) {
      return session.recordType;
    }
    const typeEntry = this.monitorEntries.find(
      (entry) =>
        entry.sourceUri === session?.sourceUri &&
        entry.sourceKind === "probe" &&
        entry.probeRole === "recordType",
    );
    const recordType = String(typeEntry?.valueText || "").trim();
    return recordType || undefined;
  }

  refreshProbePanels() {
    for (const session of this.probeSessions.values()) {
      if (!session.fieldEntriesStarted && this.getProbeResolvedRecordType(session)) {
        void this.ensureProbeFieldEntries(session);
      }
    }
    this.refreshVisibleProbeDecorations();
    this.refreshProbeWebviews();
    this.refreshProbeWidgets();
  }

  refreshProbeWebviews() {
    for (const probeWebviewState of this.probeWebviews.values()) {
      for (const panel of probeWebviewState.panels) {
        void this.postProbeWebviewState(probeWebviewState.document, panel);
      }
    }
  }

  refreshProbeWidgets() {
    for (const widgetState of this.probeWidgets.values()) {
      void this.postProbeWidgetState(widgetState);
    }
  }

  refreshPvlistWidgets() {
    for (const widgetState of this.pvlistWidgets.values()) {
      void this.postPvlistWidgetState(widgetState);
    }
  }

  refreshMonitorWidgets() {
    for (const widgetState of this.monitorWidgets.values()) {
      void this.postMonitorWidgetState(widgetState);
    }
  }

  refreshVisibleProbeDecorations() {
    if (!this.probeDecorationTypes.length) {
      return;
    }

    for (const editor of vscode.window.visibleTextEditors) {
      this.refreshProbeDecorationsForEditor(editor);
    }
  }

  buildProbePanelState(session) {
    const mainEntry = this.monitorEntries.find(
      (entry) =>
        entry.sourceUri === session.sourceUri &&
        entry.sourceKind === "probe" &&
        entry.probeRole === "main",
    );
    const recordType = this.getProbeResolvedRecordType(session);
    const fieldEntries = this.monitorEntries.filter(
      (entry) =>
        entry.sourceUri === session.sourceUri &&
        entry.sourceKind === "probe" &&
        entry.probeRole === "field" &&
        entry.hasEverSubscribed,
    );
    const orderedFieldNames = recordType
      ? (this.getFieldNamesForRecordType?.(recordType) || []).map((fieldName) =>
        String(fieldName || "").toUpperCase(),
      )
      : [];
    const fieldOrder = new Map(
      orderedFieldNames.map((fieldName, index) => [fieldName, index]),
    );
    fieldEntries.sort((left, right) => {
      const leftOrder = fieldOrder.get(String(left.probeFieldName || "").toUpperCase());
      const rightOrder = fieldOrder.get(String(right.probeFieldName || "").toUpperCase());
      if (leftOrder !== undefined && rightOrder !== undefined) {
        return leftOrder - rightOrder;
      }
      if (leftOrder !== undefined) {
        return -1;
      }
      if (rightOrder !== undefined) {
        return 1;
      }
      return String(left.probeFieldName || "").localeCompare(
        String(right.probeFieldName || ""),
      );
    });
    return {
      recordName: session.recordName,
      recordType: recordType || "(not loaded)",
      sourceLabel: session.sourceLabel,
      runtimeStatus: this.contextStatus,
      value: getProbeEntryDisplayValue(mainEntry),
      valueKey: mainEntry?.key,
      valueCanPut: this.canPutRuntimeValue(mainEntry),
      lastUpdated: mainEntry?.lastUpdated ? mainEntry.lastUpdated.toLocaleTimeString() : "Waiting for data",
      access: getProbeAccessLabel(mainEntry, this.requireRuntimeLibrarySafe()),
      fieldStatusText: !recordType
        ? "Waiting for the record type before loading fields..."
        : session.fieldEntryStartInProgress || !session.fieldEntriesStarted
          ? "Connecting field channels..."
          : "No connectable fields have reported data yet.",
      fields: fieldEntries.map((entry) => ({
        key: entry.key,
        fieldName: entry.probeFieldName,
        pvName: entry.pvName,
        value: getProbeEntryDisplayValue(entry),
        updated: entry.lastUpdated ? entry.lastUpdated.toLocaleTimeString() : "",
        canPut: this.canPutRuntimeValue(entry),
      })),
    };
  }

  buildProbeWebviewState(document) {
    const sourceUri = document?.uri?.toString();
    const analysis = analyzeProbeDocument(document);
    const session = sourceUri ? this.probeSessions.get(sourceUri) : undefined;
    const panelState = session ? this.buildProbePanelState(session) : undefined;

    let message = "";
    if (analysis.diagnostics?.length) {
      message = analysis.diagnostics[0]?.message || "";
    } else if (!analysis.recordName) {
      message = "No EPICS probe target was found in this file.";
    } else if (this.contextStatus === "stopped") {
      message = "Start the probe to create the EPICS context and connect channels.";
    } else if (!panelState) {
      message = "Probe is starting...";
    }

    return {
      sourceLabel: getRuntimeDocumentLabel(document),
      recordName: analysis.recordName || "",
      contextStatus: this.contextStatus,
      contextError: this.contextError || "",
      canStart: !analysis.diagnostics?.length && Boolean(analysis.recordName),
      canStop: Boolean(session),
      message,
      state: panelState || undefined,
      diagnostics: analysis.diagnostics || [],
    };
  }

  async postProbeWebviewState(document, webviewPanel) {
    if (!document?.uri || !webviewPanel?.webview) {
      return;
    }

    const payload = this.buildProbeWebviewState(document);
    await webviewPanel.webview.postMessage({
      type: "probeRuntimeState",
      state: payload,
    });
  }

  buildProbeWidgetWebviewState(widgetState) {
    const sourceUri = widgetState?.sourceUri;
    const session = sourceUri ? this.probeSessions.get(sourceUri) : undefined;
    const panelState = session ? this.buildProbePanelState(session) : undefined;
    const recordName = String(widgetState?.recordName || "").trim();

    let message = "";
    if (!recordName) {
      message = "Enter a channel name and press Enter to start the probe.";
    } else if (this.contextStatus === "stopped" && !session) {
      message = "Press Enter after changing the channel name to start the probe.";
    } else if (!panelState) {
      message = "Probe is starting...";
    }

    return {
      sourceLabel: widgetState?.panel?.title || "EPICS Probe",
      recordName,
      contextStatus: this.contextStatus,
      contextError: this.contextError || "",
      canProcess: Boolean(session && panelState?.recordName),
      message,
      state: panelState || undefined,
    };
  }

  buildPvlistWidgetWebviewState(widgetState) {
    const sourceUri = widgetState?.sourceUri;
    const entries = this.monitorEntries.filter(
      (entry) => entry.sourceUri === sourceUri && entry.sourceKind === "pvlistWidget",
    );
    const entryByPvName = new Map(entries.map((entry) => [entry.pvName, entry]));
    const macros = (widgetState?.sourceModel?.macroNames || []).map((macroName) => ({
      name: macroName,
      value: widgetState?.macroValues?.get(macroName) || "",
    }));
    const rows = (widgetState?.rows || []).map((row) => {
      const entry = row?.pvName ? entryByPvName.get(row.pvName) : undefined;
      return {
        id: row.id,
        channelName: row.channelName,
        key: entry?.key,
        canPut: this.canPutRuntimeValue(entry),
        value: row.valueText || getProbeEntryDisplayValue(entry),
      };
    });

    let message = "";
    if ((widgetState?.sourceModel?.diagnostics || []).length > 0) {
      message = widgetState.sourceModel.diagnostics[0]?.message || "";
    } else if (this.contextError) {
      message = this.contextError;
    } else if (!rows.length) {
      message = "No resolvable PV list entries are available with the current macro values.";
    }

    return {
      sourceLabel: widgetState?.sourceModel?.sourceLabel || "EPICS PvList",
      contextStatus: this.contextStatus,
      contextError: this.contextError || "",
      rawPvNames: Array.isArray(widgetState?.sourceModel?.rawPvNames)
        ? widgetState.sourceModel.rawPvNames
        : [],
      macros,
      rows,
      message,
    };
  }

  buildMonitorWidgetWebviewState(widgetState) {
    const rows = (widgetState?.channelRows || []).map((row) => ({
      id: row.id,
      channelName: row.channelName,
      statusText: row.lastError
        ? `(${row.lastError})`
        : row.status === "connecting"
          ? "(connecting...)"
          : row.status === "disconnected"
            ? "(disconnected)"
            : row.status === "stopped"
              ? "(stopped)"
              : row.status === "subscribed"
                ? ""
                : "",
    }));

    let message = "";
    if (widgetState?.lastError) {
      message = widgetState.lastError;
    } else if (this.contextError) {
      message = this.contextError;
    } else if (this.contextStatus === "stopped" && rows.some((row) => String(row.channelName || "").trim())) {
      message = "Runtime context is stopped.";
    } else if (!rows.some((row) => String(row.channelName || "").trim())) {
      message = "Add channels to start monitoring.";
    } else if (!widgetState?.historyLines?.length) {
      message = "Waiting for monitor data.";
    }

    return {
      sourceLabel: widgetState?.sourceLabel || "EPICS Monitor",
      contextStatus: this.contextStatus,
      contextError: this.contextError || "",
      bufferSize: widgetState?.bufferSize || DEFAULT_MONITOR_WIDGET_BUFFER_SIZE,
      rows,
      historyText: (widgetState?.historyLines || []).join("\n"),
      message,
    };
  }

  buildRuntimeWidgetWebviewState() {
    const monitors = this.monitorEntries
      .filter((entry) => !entry.hidden)
      .map((entry) => ({
        key: entry.key,
        pvName: entry.pvName,
        protocolText: String(entry.protocol || "").toUpperCase(),
        status: entry.status,
        description: buildMonitorDescription(entry),
        sourceLabel: entry.sourceLabel || "",
        valueText: entry.valueText || "",
        updatedText:
          entry.lastUpdated instanceof Date
            ? entry.lastUpdated.toLocaleTimeString()
            : "",
        errorText: entry.lastError || "",
      }));

    let message = "";
    if (this.contextError) {
      message = this.contextError;
    } else if (!monitors.length) {
      message = "No runtime monitors are active. Add one or start runtime for the active file.";
    }

    return {
      contextStatus: this.contextStatus,
      contextDescription: this.getContextDescription(),
      defaultProtocol: String(this.getDefaultProtocol() || "").toUpperCase(),
      channelTimeoutText: `${String(this.getChannelCreationTimeoutSeconds() ?? "none")}s`,
      monitorTimeoutText: `${String(this.getMonitorSubscribeTimeoutSeconds() ?? "none")}s`,
      iocShellTerminalName: this.iocShellTerminal?.name || "",
      monitorCount: monitors.length,
      monitors,
      message,
    };
  }

  async postRuntimeWidgetState() {
    if (!this.runtimeWidgetPanel?.webview) {
      return;
    }

    const payload = this.buildRuntimeWidgetWebviewState();
    await this.runtimeWidgetPanel.webview.postMessage({
      type: "runtimeWidgetState",
      state: payload,
    });
  }

  async postProbeWidgetState(widgetState) {
    if (!widgetState?.panel?.webview) {
      return;
    }

    const payload = this.buildProbeWidgetWebviewState(widgetState);
    await widgetState.panel.webview.postMessage({
      type: "probeWidgetState",
      state: payload,
    });
  }

  async postPvlistWidgetState(widgetState) {
    if (!widgetState?.panel?.webview) {
      return;
    }

    const payload = this.buildPvlistWidgetWebviewState(widgetState);
    await widgetState.panel.webview.postMessage({
      type: "pvlistWidgetState",
      state: payload,
    });
  }

  async postMonitorWidgetState(widgetState) {
    if (!widgetState?.panel?.webview) {
      return;
    }

    const payload = this.buildMonitorWidgetWebviewState(widgetState);
    await widgetState.panel.webview.postMessage({
      type: "monitorWidgetState",
      state: payload,
    });
  }

  async processProbeWidget(widgetState) {
    const recordName = String(widgetState?.recordName || "").trim();
    if (!recordName) {
      return;
    }

    let channel;
    try {
      const runtimeContext = await this.ensureRuntimeContext();
      const runtimeLibrary = this.requireRuntimeLibrary();
      const protocol = this.getDefaultProtocol();
      channel = await runtimeContext.createChannel(
        `${recordName}.PROC`,
        protocol,
        this.getChannelCreationTimeoutSeconds(),
      );
      if (!channel) {
        throw new Error(`Failed to create ${protocol.toUpperCase()} channel for ${recordName}.PROC.`);
      }
      const result = protocol === "pva"
        ? await channel.putPva("value", [1])
        : await channel.put(1, undefined, true);
      if (!isSuccessfulRuntimePutResult(result, runtimeLibrary, protocol)) {
        throw new Error(
          getRuntimePutFailureMessage(result, runtimeLibrary, protocol),
        );
      }
    } catch (error) {
      vscode.window.showErrorMessage(
        `Failed to process ${recordName}: ${getErrorMessage(error)}`,
      );
    } finally {
      try {
        await channel?.destroyHard?.();
      } catch (error) {
        try {
          await channel?.destroy?.();
        } catch (destroyError) {
          // Ignore best-effort cleanup failures for ad-hoc PROC puts.
        }
      }
    }
  }

  requireRuntimeLibrarySafe() {
    try {
      return this.requireRuntimeLibrary();
    } catch (error) {
      return undefined;
    }
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

    await this.putRuntimeValueFromInput(entry, input, putSupport);
  }

  async putRuntimeValueFromInput(target, input, resolvedPutSupport) {
    const entry = this.resolveRuntimePutEntry(target);
    if (!entry) {
      return;
    }

    const putSupport = resolvedPutSupport || this.getRuntimePutSupport(entry);
    if (!putSupport.canPut) {
      vscode.window.showWarningMessage(putSupport.reason);
      return;
    }

    if (!this.activePutRequestKeys.has(entry.key)) {
      this.activePutRequestKeys.add(entry.key);
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
      hidden: Boolean(options.hidden || definition.hidden),
      sourceKind: definition.sourceKind,
      probeRole: definition.probeRole,
      probeFieldName: definition.probeFieldName,
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
      softDisconnectTimer: undefined,
      hasEverSubscribed: false,
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

  async connectEntriesInBatches(
    entries,
    runtimeContext,
    batchSize = PROBE_FIELD_CONNECT_BATCH_SIZE,
  ) {
    const normalizedBatchSize =
      Math.max(1, Number(batchSize) || PROBE_FIELD_CONNECT_BATCH_SIZE);
    const pendingEntries = Array.isArray(entries) ? entries.filter(Boolean) : [];
    for (let index = 0; index < pendingEntries.length; index += normalizedBatchSize) {
      const batch = pendingEntries.slice(index, index + normalizedBatchSize);
      await Promise.allSettled(
        batch.map((entry) => this.connectEntry(entry, runtimeContext)),
      );
    }
  }

  async connectEntry(entry, runtimeContext) {
    if (!entry || !this.monitorEntries.includes(entry)) {
      return;
    }

    const attemptId = Number(entry.connectionAttemptId || 0) + 1;
    entry.connectionAttemptId = attemptId;
    this.clearSoftDisconnectTimer(entry);
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
        await this.cleanupConnectionResources(entry, monitor, channel);
      }
      this.handleActiveEditorChange(vscode.window.activeTextEditor);
    }
  }

  attachChannelToEntry(entry, channel, attemptId) {
    entry.channel = channel;
    entry.serverAddress = channel.getServerAddress();

    if (entry?.sourceKind === "probe") {
      return;
    }

    channel.setDestroySoftCallback(() => {
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      this.scheduleSoftDisconnect(entry, attemptId);
    });
    channel.setDestroyHardCallback(() => {
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      this.clearSoftDisconnectTimer(entry);
      entry.status = "destroyed";
      entry.channel = undefined;
      entry.monitor = undefined;
      entry.lastError = "Channel was destroyed.";
      this.refresh(entry);
    });
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
      this.clearSoftDisconnectTimer(entry);
      entry.status = "subscribed";
      entry.lastError = undefined;
      entry.hasEverSubscribed = true;
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
    this.clearSoftDisconnectTimer(entry);

    entry.monitor = undefined;
    entry.channel = undefined;
    entry.serverAddress = undefined;

    try {
      monitor?.destroyHard();
    } catch (error) {
      // Ignore monitor teardown errors while cleaning up the UI state.
    }

    if (entry?.sourceKind !== "probe") {
      try {
        await channel?.destroyHard();
      } catch (error) {
        // Ignore channel teardown errors while cleaning up the UI state.
      }
    }

    entry.status = "stopped";
  }

  async cleanupConnectionResources(entry, monitor, channel) {
    try {
      monitor?.destroyHard();
    } catch (error) {
      // Ignore teardown errors for abandoned async connection attempts.
    }

    if (entry?.sourceKind !== "probe") {
      try {
        await channel?.destroyHard();
      } catch (error) {
        // Ignore teardown errors for abandoned async connection attempts.
      }
    }
  }

  handleMonitorUpdate(entry, monitor) {
    if (entry.monitor && entry.monitor !== monitor) {
      return;
    }

    this.clearSoftDisconnectTimer(entry);
    entry.status = "subscribed";
    entry.lastError = undefined;
    entry.hasEverSubscribed = true;
    this.updateEntryValue(entry, monitor);
    this.refresh(entry);
  }

  scheduleSoftDisconnect(entry, attemptId) {
    this.clearSoftDisconnectTimer(entry);
    entry.softDisconnectTimer = setTimeout(() => {
      entry.softDisconnectTimer = undefined;
      if (!this.isCurrentConnectionAttempt(entry, attemptId)) {
        return;
      }
      if (entry.status === "subscribed") {
        return;
      }
      entry.status = "disconnected";
      entry.lastError = "Channel disconnected. Waiting for recovery.";
      this.refresh(entry);
    }, 1500);
  }

  clearSoftDisconnectTimer(entry) {
    if (!entry?.softDisconnectTimer) {
      return;
    }
    clearTimeout(entry.softDisconnectTimer);
    entry.softDisconnectTimer = undefined;
  }

  updateEntryValue(entry, monitor) {
    let runtimeValue;
    if (entry.protocol === "pva") {
      updatePvaEnumChoicesCache(entry, monitor.getPvaData());
      runtimeValue = getPvaRuntimeDisplayValue(
        monitor.getPvaData(),
        entry.pvaEnumChoices,
      );
    } else {
      updateCaEnumChoicesCache(entry, monitor.getChannel().getDbrData?.());
      runtimeValue = getCaRuntimeDisplayValue(
        entry,
        monitor.getChannel().getDbrData?.(),
      );
    }

    if (shouldIgnoreTransientProbeFieldValue(entry, runtimeValue)) {
      return;
    }

    entry.lastUpdated = new Date();
    entry.valueText = formatRuntimeValue(runtimeValue);
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
      this.clearSoftDisconnectTimer(entry);
      entry.channel = undefined;
      entry.monitor = undefined;
      if (entry.status !== "error") {
        entry.status = "stopped";
      }
    }

    for (const widgetState of this.monitorWidgets.values()) {
      this.stopMonitorWidgetSessions(widgetState);
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
    void this.postRuntimeWidgetState();
    this.refreshProbePanels();
    this.refreshPvlistWidgets();
    this.refreshMonitorWidgets();
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
    const iocTerminalName = this.iocShellTerminal?.name;
    if (this.contextStatus === "connected") {
      const activeDescription = `${this.monitorEntries.filter((entry) => entry.monitor).length} active`;
      return iocTerminalName
        ? `${activeDescription} | IOC: ${iocTerminalName}`
        : activeDescription;
    }

    if (this.contextStatus === "connecting") {
      return iocTerminalName ? `connecting | IOC: ${iocTerminalName}` : "connecting";
    }

    if (this.contextStatus === "error") {
      return iocTerminalName ? `error | IOC: ${iocTerminalName}` : "error";
    }

    return iocTerminalName ? `stopped | IOC: ${iocTerminalName}` : "stopped";
  }

  getContextTooltip() {
    const lines = [
      `Status: ${this.contextStatus}`,
      `Default protocol: ${this.getDefaultProtocol()}`,
      `Channel timeout: ${String(this.getChannelCreationTimeoutSeconds() ?? "none")}s`,
      `Monitor timeout: ${String(this.getMonitorSubscribeTimeoutSeconds() ?? "none")}s`,
    ];
    if (this.iocShellTerminal?.name) {
      lines.push(`IOC shell terminal: ${this.iocShellTerminal.name}`);
    }
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

    const visibleEntries = this.monitorEntries.filter((candidate) => !candidate.hidden);
    if (!visibleEntries.length) {
      return undefined;
    }

    if (visibleEntries.length === 1) {
      return visibleEntries[0];
    }

    const selected = await vscode.window.showQuickPick(
      visibleEntries.map((candidate) => ({
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

  databaseDocumentHasToc(document) {
    if (!isDatabaseRuntimeDocument(document)) {
      return false;
    }

    const tocEntries = this.extractDatabaseTocEntries?.(document.getText()) || [];
    return tocEntries.length > 0;
  }

  updateDatabaseEditorContextKeys(editor) {
    const document = editor?.document;
    const hasToc = this.databaseDocumentHasToc(document);
    const isRunning =
      Boolean(document?.uri) && this.hasEntriesForSourceUri(document.uri.toString());

    void vscode.commands.executeCommand(
      "setContext",
      ACTIVE_DATABASE_HAS_TOC_CONTEXT_KEY,
      hasToc,
    );
    void vscode.commands.executeCommand(
      "setContext",
      ACTIVE_DATABASE_MONITORING_RUNNING_CONTEXT_KEY,
      hasToc && isRunning,
    );
  }

  updateStartupEditorContextKeys(editor) {
    const startupTarget = this.resolveActiveStartupIocDocument(editor);
    const isRunning =
      Boolean(startupTarget) &&
      this.getCandidateRunningStartupIocTerminals(startupTarget.documentPath).length > 0;

    void vscode.commands.executeCommand(
      "setContext",
      ACTIVE_STARTUP_CAN_START_IOC_CONTEXT_KEY,
      Boolean(startupTarget) && !isRunning,
    );
    void vscode.commands.executeCommand(
      "setContext",
      ACTIVE_STARTUP_CAN_STOP_IOC_CONTEXT_KEY,
      Boolean(startupTarget) && isRunning,
    );
  }

  async removeSpecificMonitorEntries(entries) {
    const removableEntries = Array.isArray(entries) ? entries.filter(Boolean) : [];
    if (!removableEntries.length) {
      return;
    }

    const removableKeys = new Set(removableEntries.map((entry) => entry.key));
    for (const entry of removableEntries) {
      await this.disconnectEntry(entry);
    }

    this.monitorEntries = this.monitorEntries.filter(
      (entry) => !removableKeys.has(entry.key),
    );
    this.refreshProbePanels();
    if (!this.monitorEntries.length) {
      this.stopContextInternal();
    }
    this.refresh();
    this.handleActiveEditorChange(vscode.window.activeTextEditor);
  }

  async removeEntriesBySourceUri(sourceUri) {
    if (!sourceUri) {
      return;
    }

    const entries = this.monitorEntries.filter(
      (entry) => entry.sourceUri === sourceUri,
    );
    await this.removeSpecificMonitorEntries(entries);
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

  if (isDatabaseRuntimeDocument(document) || isStrictMonitorDocument(document) || isProbeDocument(document)) {
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

  if (document.languageId === "pvlist") {
    return true;
  }

  return (
    document.uri.scheme === "file" &&
    path.extname(document.uri.fsPath).toLowerCase() === ".pvlist"
  );
}

function isProbeDocument(document) {
  if (!document?.uri) {
    return false;
  }

  if (document.languageId === "probe") {
    return true;
  }

  return (
    document.uri.scheme === "file" &&
    PROBE_RUNTIME_EXTENSIONS.has(path.extname(document.uri.fsPath).toLowerCase())
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

function normalizeFsPath(fsPath) {
  if (!fsPath) {
    return "";
  }
  return path.resolve(String(fsPath)).replace(/\\/g, "/");
}

function isPathUnderIocBoot(fsPath) {
  const normalizedPath = normalizeFsPath(fsPath);
  if (!normalizedPath) {
    return false;
  }

  return /(^|\/)iocBoot(\/|$)/.test(normalizedPath);
}

function readFirstLineFromFile(fsPath) {
  try {
    const text = fs.readFileSync(fsPath, "utf8");
    return String(text || "").split(/\r?\n/, 1)[0] || "";
  } catch (error) {
    return "";
  }
}

function extractShebangExecutable(text) {
  const firstLine = String(text || "").split(/\r?\n/, 1)[0] || "";
  if (!firstLine.startsWith("#!")) {
    return "";
  }

  return extractCommandLineExecutable(firstLine.slice(2).trim());
}

function resolveExistingExecutablePath(executableText, baseDirectory = "") {
  const normalizedExecutable = String(executableText || "").trim();
  if (!normalizedExecutable) {
    return undefined;
  }

  const hasPathSeparator =
    normalizedExecutable.includes("/") || normalizedExecutable.includes("\\");
  if (path.isAbsolute(normalizedExecutable)) {
    return fs.existsSync(normalizedExecutable) ? normalizedExecutable : undefined;
  }

  if (hasPathSeparator) {
    const candidatePath = baseDirectory
      ? path.resolve(baseDirectory, normalizedExecutable)
      : path.resolve(normalizedExecutable);
    return fs.existsSync(candidatePath) ? candidatePath : undefined;
  }

  const pathEntries = String(process.env.PATH || "")
    .split(path.delimiter)
    .filter(Boolean);
  for (const pathEntry of pathEntries) {
    const candidatePath = path.join(pathEntry, normalizedExecutable);
    if (fs.existsSync(candidatePath)) {
      return candidatePath;
    }
  }

  return undefined;
}

function isStcmdLikeFilePath(fsPath, documentText) {
  const normalizedPath = normalizeFsPath(fsPath);
  if (!normalizedPath || !fs.existsSync(normalizedPath)) {
    return false;
  }

  let stats;
  try {
    stats = fs.statSync(normalizedPath);
  } catch (error) {
    return false;
  }
  if (!stats.isFile()) {
    return false;
  }

  if (!isPathUnderIocBoot(normalizedPath)) {
    return false;
  }

  const extension = path.extname(normalizedPath).toLowerCase();
  if (extension !== ".cmd" && extension !== ".iocsh") {
    return false;
  }

  return Boolean(
    resolveStartupExecutableValidation(normalizedPath, documentText).executableText,
  );
}

function resolveStartupExecutableValidation(fsPath, documentText) {
  const normalizedPath = normalizeFsPath(fsPath);
  const executableText = extractShebangExecutable(
    documentText === undefined ? readFirstLineFromFile(normalizedPath) : documentText,
  );
  const executableName = path.basename(executableText || "");
  return {
    executableText,
    executableName,
    executablePath: executableText
      ? resolveExistingExecutablePath(executableText, path.dirname(normalizedPath))
      : undefined,
  };
}

function getIocBootRelativeDisplayPath(projectRootPath, startupDocumentPath) {
  const iocBootRootPath = normalizeFsPath(path.join(projectRootPath, "iocBoot"));
  const relativePath = path.relative(iocBootRootPath, startupDocumentPath).replace(/\\/g, "/");
  return relativePath.replace(/^(\.\.\/)+/, "");
}

function getIocBootDisplayPath(projectRootPath, startupDocumentPath) {
  const relativePath = getIocBootRelativeDisplayPath(projectRootPath, startupDocumentPath);
  return relativePath ? `iocBoot/${relativePath}` : "iocBoot";
}

function getCommandResourceFsPath(resourceUri) {
  const candidateUri =
    resourceUri?.scheme
      ? resourceUri
      : Array.isArray(resourceUri) && resourceUri[0]?.scheme
        ? resourceUri[0]
        : undefined;
  return candidateUri?.scheme === "file" ? normalizeFsPath(candidateUri.fsPath) : "";
}

function isExistingFile(filePath) {
  try {
    return fs.statSync(filePath).isFile();
  } catch (error) {
    return false;
  }
}

function isFsPathEpicsProjectRoot(rootPath) {
  if (!rootPath || !fs.existsSync(rootPath)) {
    return false;
  }

  try {
    if (!fs.statSync(rootPath).isDirectory()) {
      return false;
    }
  } catch (error) {
    return false;
  }

  return EPICS_PROJECT_MARKER_SEGMENTS.every((pathSegments) =>
    isExistingFile(path.join(rootPath, ...pathSegments)),
  );
}

function findContainingEpicsProjectRootPath(filePath) {
  if (!filePath) {
    return undefined;
  }

  let currentPath = normalizeFsPath(filePath);
  try {
    if (fs.statSync(currentPath).isFile()) {
      currentPath = normalizeFsPath(path.dirname(currentPath));
    }
  } catch (error) {
    return undefined;
  }

  while (currentPath) {
    if (isFsPathEpicsProjectRoot(currentPath)) {
      return currentPath;
    }

    const parentPath = normalizeFsPath(path.dirname(currentPath));
    if (parentPath === currentPath) {
      return undefined;
    }
    currentPath = parentPath;
  }

  return undefined;
}

function getWorkspaceEpicsProjectRootPaths() {
  const roots = new Set();
  for (const workspaceFolder of vscode.workspace.workspaceFolders || []) {
    const containingRootPath = findContainingEpicsProjectRootPath(workspaceFolder.uri?.fsPath);
    if (containingRootPath) {
      roots.add(containingRootPath);
      continue;
    }

    const workspaceRootPath = normalizeFsPath(workspaceFolder.uri?.fsPath);
    if (isFsPathEpicsProjectRoot(workspaceRootPath)) {
      roots.add(workspaceRootPath);
    }
  }
  return [...roots].sort((left, right) => left.localeCompare(right));
}

function findProjectStcmdLikeFilePaths(projectRootPath) {
  const iocBootRootPath = normalizeFsPath(path.join(projectRootPath, "iocBoot"));
  if (!iocBootRootPath || !fs.existsSync(iocBootRootPath)) {
    return [];
  }

  const startupDocumentPaths = [];
  const pendingDirectories = [iocBootRootPath];
  while (pendingDirectories.length > 0) {
    const currentDirectory = pendingDirectories.pop();
    let directoryEntries = [];
    try {
      directoryEntries = fs.readdirSync(currentDirectory, { withFileTypes: true });
    } catch (error) {
      continue;
    }

    for (const directoryEntry of directoryEntries) {
      const entryPath = path.join(currentDirectory, directoryEntry.name);
      if (directoryEntry.isDirectory()) {
        pendingDirectories.push(entryPath);
        continue;
      }
      if (!directoryEntry.isFile()) {
        continue;
      }
      if (isStcmdLikeFilePath(entryPath)) {
        startupDocumentPaths.push(normalizeFsPath(entryPath));
      }
    }
  }

  return startupDocumentPaths.sort((left, right) =>
    getIocBootRelativeDisplayPath(projectRootPath, left).localeCompare(
      getIocBootRelativeDisplayPath(projectRootPath, right),
    ),
  );
}

function findSiblingStcmdLikeFilePaths(directoryPath) {
  const normalizedDirectoryPath = normalizeFsPath(directoryPath);
  if (!normalizedDirectoryPath || !fs.existsSync(normalizedDirectoryPath)) {
    return [];
  }

  let directoryEntries = [];
  try {
    directoryEntries = fs.readdirSync(normalizedDirectoryPath, { withFileTypes: true });
  } catch (error) {
    return [];
  }

  const startupDocumentPaths = [];
  for (const directoryEntry of directoryEntries) {
    if (!directoryEntry?.isFile?.()) {
      continue;
    }

    const entryPath = path.join(normalizedDirectoryPath, directoryEntry.name);
    if (isStcmdLikeFilePath(entryPath)) {
      startupDocumentPaths.push(normalizeFsPath(entryPath));
    }
  }

  return startupDocumentPaths.sort((left, right) => left.localeCompare(right));
}

function canInferRunningStartupIocByWorkingDirectory(startupDocumentPath) {
  const normalizedStartupDocumentPath = normalizeFsPath(startupDocumentPath);
  if (!normalizedStartupDocumentPath) {
    return false;
  }

  const siblingStartupPaths = findSiblingStcmdLikeFilePaths(
    path.dirname(normalizedStartupDocumentPath),
  );
  if (siblingStartupPaths.length !== 1) {
    return false;
  }

  return siblingStartupPaths[0] === normalizedStartupDocumentPath;
}

function getIocBootStartupDocumentPath(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  const documentPath = document.uri.fsPath;
  if (!isStcmdLikeFilePath(documentPath, document.getText())) {
    return undefined;
  }

  return normalizeFsPath(documentPath);
}

function extractCommandLineExecutable(commandLine) {
  const trimmedCommandLine = String(commandLine || "").trim();
  if (!trimmedCommandLine) {
    return "";
  }

  const match = trimmedCommandLine.match(/^(?:"([^"]+)"|'([^']+)'|(\S+))/);
  return match?.[1] || match?.[2] || match?.[3] || "";
}

function resolveStartupCommandDocumentPath(commandLine, cwdUri) {
  const executable = extractCommandLineExecutable(commandLine);
  if (!executable) {
    return undefined;
  }

  const executableText = String(executable).trim();
  let candidatePath = executableText;
  if (!path.isAbsolute(candidatePath)) {
    if (cwdUri?.scheme !== "file") {
      return undefined;
    }
    candidatePath = path.resolve(cwdUri.fsPath, candidatePath);
  }

  if (!isStcmdLikeFilePath(candidatePath)) {
    return undefined;
  }

  return normalizeFsPath(candidatePath);
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
  const rawProtocol = rawConfig?.protocol;
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
  const normalized = normalizeProjectRuntimeConfiguration(config);
  const serialized = {
    EPICS_CA_ADDR_LIST: normalized.EPICS_CA_ADDR_LIST,
    EPICS_CA_AUTO_ADDR_LIST: normalized.EPICS_CA_AUTO_ADDR_LIST,
  };
  if (normalized.protocol === "pva") {
    serialized.protocol = "pva";
  }
  return `${JSON.stringify(serialized, null, 2)}\n`;
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

function buildProbeCustomEditorHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Probe</title>
    <style>
      :root {
        color-scheme: light dark;
      }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .toolbar {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px 16px;
        border-bottom: 1px solid var(--vscode-panel-border);
        position: sticky;
        top: 0;
        background: var(--vscode-editor-background);
        z-index: 1;
      }
      button {
        border: 1px solid var(--vscode-button-border, transparent);
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        padding: 6px 12px;
        border-radius: 4px;
        cursor: pointer;
      }
      button.secondary {
        background: transparent;
        color: var(--vscode-foreground);
      }
      button:disabled {
        opacity: 0.5;
        cursor: default;
      }
      .status {
        color: var(--vscode-descriptionForeground);
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 20px 24px 28px;
      }
      .content.overlay-active {
        display: flex;
      }
      h1, h2, p {
        margin: 0;
      }
      .message {
        margin-top: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .error {
        margin-top: 12px;
        color: var(--vscode-errorForeground);
      }
      .meta-grid {
        display: grid;
        grid-template-columns: max-content 1fr;
        gap: 10px 16px;
        margin-top: 16px;
      }
      .meta-label {
        color: var(--vscode-descriptionForeground);
      }
      .fields {
        margin-top: 24px;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        margin-top: 10px;
      }
      th, td {
        text-align: left;
        padding: 8px 10px;
        border-bottom: 1px solid var(--vscode-panel-border);
        vertical-align: top;
      }
      th {
        color: var(--vscode-descriptionForeground);
        font-weight: 600;
      }
      td:first-child, th:first-child {
        width: 18ch;
      }
      code {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
      }
      .put-target {
        cursor: pointer;
      }
      .put-target:hover {
        text-decoration: underline;
      }
      .empty {
        margin-top: 12px;
        color: var(--vscode-descriptionForeground);
      }
    </style>
  </head>
  <body>
    <div class="toolbar">
      <button id="startButton">Start Probe</button>
      <button id="stopButton" class="secondary">Stop Probe</button>
      <span id="statusText" class="status"></span>
    </div>
    <div id="content" class="content"></div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const doubleClickIntervalMs = ${JSON.stringify(MOUSE_DOUBLE_CLICK_INTERVAL_MS)};
      const initialState = ${initialStateJson};
      const content = document.getElementById("content");
      const startButton = document.getElementById("startButton");
      const stopButton = document.getElementById("stopButton");
      const statusText = document.getElementById("statusText");
      let pendingPutClick = undefined;

      startButton.addEventListener("click", () => {
        vscode.postMessage({ type: "startProbeRuntime" });
      });
      stopButton.addEventListener("click", () => {
        vscode.postMessage({ type: "stopProbeRuntime" });
      });

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function render(payload) {
        const state = payload?.state;
        const title = escapeHtml(payload?.recordName || payload?.sourceLabel || "EPICS Probe");
        startButton.disabled = !payload?.canStart || payload?.canStop;
        stopButton.disabled = !payload?.canStop;
        statusText.textContent =
          payload?.contextStatus === "running"
            ? "Runtime active"
            : payload?.contextStatus === "error"
              ? "Runtime error"
              : "Runtime stopped";

        const message = payload?.message
          ? '<p class="' + (payload?.diagnostics?.length ? 'error' : 'message') + '">' + escapeHtml(payload.message) + '</p>'
          : '';

        if (!state) {
          content.innerHTML = '<h1>' + title + '</h1>' + message;
          return;
        }

        const fieldRows = (state.fields || []).map((field) => {
          const putClass = field.canPut ? 'put-target' : '';
          const putTitle = field.canPut ? ' title="Double-click to put a new value"' : '';
          return '<tr>' +
            '<td><code>' + escapeHtml(field.fieldName) + '</code></td>' +
            '<td class="' + putClass + '" data-key="' + escapeHtml(field.key) + '" data-can-put="' + (field.canPut ? "true" : "false") + '"' + putTitle + '>' + escapeHtml(field.value) + '</td>' +
            '</tr>';
        }).join('');

        const valueClass = state.valueCanPut ? 'put-target' : '';
        const valueTitle = state.valueCanPut ? ' title="Double-click to put a new value"' : '';
        content.innerHTML =
          '<h1>' + title + '</h1>' +
          message +
          '<div class="meta-grid">' +
            '<div class="meta-label">Value</div><div class="' + valueClass + '" data-key="' + escapeHtml(state.valueKey || '') + '" data-can-put="' + (state.valueCanPut ? "true" : "false") + '"' + valueTitle + '>' + escapeHtml(state.value) + '</div>' +
            '<div class="meta-label">Record Type</div><div><code>' + escapeHtml(state.recordType) + '</code></div>' +
            '<div class="meta-label">Last Update</div><div>' + escapeHtml(state.lastUpdated) + '</div>' +
            '<div class="meta-label">Permission</div><div>' + escapeHtml(state.access) + '</div>' +
          '</div>' +
          '<div class="fields">' +
            '<h2>Fields</h2>' +
            (fieldRows
              ? '<table><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>' + fieldRows + '</tbody></table>'
              : '<p class="empty">' + escapeHtml(state.fieldStatusText || 'No fields loaded.') + '</p>') +
          '</div>';
      }

      content.addEventListener('click', (event) => {
        const eventTarget = event.target instanceof Element
          ? event.target
          : event.target?.parentElement;
        const target = eventTarget?.closest?.('[data-key][data-can-put="true"]');
        if (!target) {
          pendingPutClick = undefined;
          return;
        }
        const key = target.getAttribute('data-key');
        if (!key) {
          pendingPutClick = undefined;
          return;
        }
        const now = Date.now();
        const isDoubleClick =
          pendingPutClick?.key === key &&
          now - pendingPutClick.time <= doubleClickIntervalMs;
        pendingPutClick = isDoubleClick
          ? undefined
          : {
              key,
              time: now,
            };
        if (!isDoubleClick) {
          return;
        }
        vscode.postMessage({
          type: 'putProbeValue',
          key,
        });
      });

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'probeRuntimeState') {
          render(event.data.state);
        }
      });

      render(initialState);
    </script>
  </body>
</html>`;
}

function buildRuntimeWidgetHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Runtime</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .toolbar {
        position: sticky;
        top: 0;
        z-index: 1;
        display: flex;
        flex-wrap: wrap;
        gap: 10px 12px;
        align-items: center;
        justify-content: space-between;
        padding: 12px 16px;
        border-bottom: 1px solid var(--vscode-panel-border);
        background: var(--vscode-editor-background);
      }
      .toolbar-left, .toolbar-right {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .toolbar-title {
        font-weight: 600;
        margin-right: 8px;
      }
      .toolbar-status {
        color: var(--vscode-descriptionForeground);
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 20px 24px 28px;
      }
      .action-button {
        padding: 6px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 4px;
        background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
        color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
        cursor: pointer;
        font: inherit;
      }
      .action-button:hover {
        background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground));
      }
      .action-button.primary {
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
      }
      .action-button.primary:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .section + .section {
        margin-top: 24px;
      }
      .section-title {
        margin: 0 0 12px;
        font-size: 1.05rem;
        font-weight: 600;
      }
      .summary-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px;
      }
      .summary-card {
        border: 1px solid var(--vscode-panel-border);
        border-radius: 8px;
        padding: 12px 14px;
        background: color-mix(in srgb, var(--vscode-editor-background) 88%, var(--vscode-editorHoverWidget-background, var(--vscode-sideBar-background)) 12%);
      }
      .summary-label {
        color: var(--vscode-descriptionForeground);
        font-size: 0.9rem;
        margin-bottom: 6px;
      }
      .summary-value {
        font-weight: 600;
        word-break: break-word;
      }
      .message {
        margin-top: 14px;
        color: var(--vscode-descriptionForeground);
      }
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
      }
      th, td {
        text-align: left;
        padding: 10px 10px;
        border-bottom: 1px solid var(--vscode-panel-border);
        vertical-align: top;
      }
      th {
        color: var(--vscode-descriptionForeground);
        font-weight: 600;
      }
      th:nth-child(1), td:nth-child(1) {
        width: 22ch;
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
      }
      th:nth-child(2), td:nth-child(2) {
        width: 14ch;
      }
      th:nth-child(5), td:nth-child(5) {
        width: 8ch;
      }
      .monitor-meta {
        color: var(--vscode-descriptionForeground);
        margin-top: 4px;
        font-size: 0.9rem;
      }
      .status-badge {
        display: inline-flex;
        align-items: center;
        border: 1px solid var(--vscode-panel-border);
        border-radius: 999px;
        padding: 2px 8px;
        font-size: 0.85rem;
      }
      .empty {
        color: var(--vscode-descriptionForeground);
      }
      .remove-button {
        padding: 4px 8px;
        min-width: 0;
      }
    </style>
  </head>
  <body>
    <div class="toolbar">
      <div class="toolbar-left">
        <div class="toolbar-title">EPICS Runtime</div>
        <div id="statusText" class="toolbar-status"></div>
      </div>
      <div class="toolbar-right">
        <button class="action-button primary" data-action="addRuntimeMonitor">Add Monitor</button>
        <button class="action-button" data-action="restartRuntimeContext">Restart</button>
        <button class="action-button" data-action="stopRuntimeContext">Stop</button>
        <button class="action-button" data-action="clearRuntimeMonitors">Clear</button>
        <button class="action-button" data-action="openProjectRuntimeConfiguration">Config</button>
        <button class="action-button" data-action="setIocShellTerminal">Set IOC Terminal</button>
        <button class="action-button" data-action="runIocShellCommand">IOC Command...</button>
        <button class="action-button" data-action="runIocShellDbl">dbl</button>
        <button class="action-button" data-action="runIocShellDbprCurrentRecord">dbpr</button>
      </div>
    </div>
    <div id="content" class="content"></div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const initialState = ${initialStateJson};
      const content = document.getElementById("content");
      const statusText = document.getElementById("statusText");

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function render(state) {
        const payload = state || {};
        statusText.textContent = payload.contextDescription || payload.contextStatus || "";

        const monitorRows = Array.isArray(payload.monitors) ? payload.monitors : [];
        const monitorsHtml = monitorRows.length
          ? '<table><thead><tr><th>PV</th><th>Status</th><th>Value</th><th>Source</th><th></th></tr></thead><tbody>' +
              monitorRows.map((monitor) => {
                const badgeText = escapeHtml((monitor.protocolText ? monitor.protocolText + ' ' : '') + (monitor.status || ''));
                const valueText = monitor.valueText || monitor.errorText || '';
                const metaLines = [
                  monitor.description || '',
                  monitor.updatedText ? ('Updated: ' + monitor.updatedText) : ''
                ].filter(Boolean);
                return '<tr>' +
                  '<td><div>' + escapeHtml(monitor.pvName || '') + '</div>' +
                    (metaLines.length
                      ? '<div class="monitor-meta">' + escapeHtml(metaLines.join(' | ')) + '</div>'
                      : '') +
                  '</td>' +
                  '<td><span class="status-badge">' + badgeText + '</span></td>' +
                  '<td>' + escapeHtml(valueText) + '</td>' +
                  '<td>' + escapeHtml(monitor.sourceLabel || '') + '</td>' +
                  '<td><button class="action-button remove-button" data-remove-key="' + escapeHtml(monitor.key || '') + '">Remove</button></td>' +
                '</tr>';
              }).join('') +
            '</tbody></table>'
          : '<div class="empty">No runtime monitors are active.</div>';

        const messageHtml = payload.message
          ? '<div class="message">' + escapeHtml(payload.message) + '</div>'
          : '';

        content.innerHTML =
          '<div class="section">' +
            '<div class="summary-grid">' +
              '<div class="summary-card"><div class="summary-label">Status</div><div class="summary-value">' + escapeHtml(payload.contextStatus || 'stopped') + '</div></div>' +
              '<div class="summary-card"><div class="summary-label">Default Protocol</div><div class="summary-value">' + escapeHtml(payload.defaultProtocol || '') + '</div></div>' +
              '<div class="summary-card"><div class="summary-label">Channel Timeout</div><div class="summary-value">' + escapeHtml(payload.channelTimeoutText || '') + '</div></div>' +
              '<div class="summary-card"><div class="summary-label">Monitor Timeout</div><div class="summary-value">' + escapeHtml(payload.monitorTimeoutText || '') + '</div></div>' +
              '<div class="summary-card"><div class="summary-label">IOC Terminal</div><div class="summary-value">' + escapeHtml(payload.iocShellTerminalName || '(not set)') + '</div></div>' +
              '<div class="summary-card"><div class="summary-label">Monitors</div><div class="summary-value">' + escapeHtml(String(payload.monitorCount || 0)) + '</div></div>' +
            '</div>' +
            messageHtml +
          '</div>' +
          '<div class="section">' +
            '<div class="section-title">Monitor Entries</div>' +
            monitorsHtml +
          '</div>';

        for (const button of content.querySelectorAll('[data-remove-key]')) {
          button.addEventListener('click', () => {
            vscode.postMessage({
              type: 'removeRuntimeMonitor',
              key: button.dataset.removeKey,
            });
          });
        }
      }

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'runtimeWidgetState') {
          render(event.data.state);
        }
      });

      for (const button of document.querySelectorAll('[data-action]')) {
        button.addEventListener('click', () => {
          vscode.postMessage({ type: button.dataset.action });
        });
      }

      render(initialState);
    </script>
  </body>
</html>`;
}

function buildIocRuntimeCommandsHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>IOC Runtime Commands</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .header {
        padding: 16px 20px 14px;
        border-bottom: 1px solid var(--vscode-panel-border);
      }
      .header-top {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        justify-content: space-between;
      }
      .title {
        margin: 0;
        font-size: 1.35rem;
        font-weight: 600;
      }
      .meta {
        margin-top: 8px;
        color: var(--vscode-descriptionForeground);
        line-height: 1.5;
      }
      .toolbar {
        margin-top: 14px;
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .toolbar-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .filter-input {
        width: 100%;
        min-width: 26ch;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .custom-command-bar {
        margin-top: 12px;
        display: grid;
        grid-template-columns: minmax(28ch, 1fr) auto;
        gap: 10px;
        align-items: center;
      }
      .custom-command-input {
        width: 100%;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .header-button {
        padding: 8px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 6px;
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        cursor: pointer;
        font: inherit;
      }
      .header-button:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .header-button:disabled,
      .send-button:disabled,
      .args-input:disabled {
        opacity: 0.55;
        cursor: default;
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 18px 20px 24px;
      }
      .message {
        margin-bottom: 16px;
        color: var(--vscode-descriptionForeground);
      }
      .message.error {
        color: var(--vscode-errorForeground);
      }
      .filter-summary {
        margin-bottom: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .command-list {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      .command-card {
        display: flex;
        flex-direction: column;
        gap: 0;
        border: 1px solid var(--vscode-panel-border);
        border-radius: 8px;
        background: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-editorHoverWidget-background, var(--vscode-sideBar-background)) 10%);
        overflow: hidden;
      }
      .command-row {
        display: grid;
        grid-template-columns: minmax(18ch, 22ch) minmax(24ch, 1fr) auto auto;
        gap: 10px;
        align-items: center;
        padding: 10px 12px 8px;
      }
      .command-name {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
        font-size: 0.98rem;
        font-weight: 600;
      }
      .command-help {
        padding: 0 12px 12px;
        color: var(--vscode-descriptionForeground);
        white-space: pre-wrap;
        line-height: 1.45;
        font-size: 0.93rem;
      }
      .command-help::before {
        content: "";
        display: block;
        margin-bottom: 8px;
        border-top: 1px solid var(--vscode-panel-border);
      }
      .args-input {
        width: 100%;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .send-button {
        padding: 8px 14px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 6px;
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        cursor: pointer;
        font: inherit;
      }
      .send-button:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .capture-label {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        color: var(--vscode-descriptionForeground);
        white-space: nowrap;
      }
      .empty {
        color: var(--vscode-descriptionForeground);
      }
      body.stopped .meta,
      body.stopped .toolbar,
      body.stopped .content {
        opacity: 0.45;
        filter: grayscale(0.25);
      }
      .stopped-watermark {
        display: none;
        position: fixed;
        inset: 170px 0 0 0;
        pointer-events: none;
        overflow: hidden;
        z-index: 2;
      }
      body.stopped .stopped-watermark {
        display: block;
      }
      .stopped-watermark-grid {
        position: absolute;
        inset: 0 -10% 0 -10%;
        display: grid;
        grid-template-columns: repeat(4, minmax(18ch, 1fr));
        gap: 34px 48px;
        align-content: start;
        transform: rotate(-18deg);
      }
      .stopped-watermark-text {
        color: transparent;
        font-size: 3.1rem;
        font-style: italic;
        font-weight: 700;
        letter-spacing: 0.08em;
        -webkit-text-stroke: 1.3px rgba(220, 38, 38, 0.65);
        user-select: none;
      }
      @media (max-width: 900px) {
        .custom-command-bar {
          grid-template-columns: 1fr;
        }
        .command-row {
          grid-template-columns: 1fr;
        }
        .stopped-watermark-grid {
          grid-template-columns: repeat(2, minmax(18ch, 1fr));
        }
      }
    </style>
  </head>
  <body>
    <div class="stopped-watermark" aria-hidden="true">
      <div class="stopped-watermark-grid">
        ${Array.from({ length: 24 }, () => '<div class="stopped-watermark-text">Stopped</div>').join("")}
      </div>
    </div>
    <div class="header">
      <div class="header-top">
        <h1 class="title">IOC Runtime Commands</h1>
        <div class="toolbar-actions">
          <button id="showTerminalButton" class="header-button">Show Running Terminal</button>
          <button id="startIocButton" class="header-button">Start IOC</button>
          <button id="stopIocButton" class="header-button">Stop IOC</button>
        </div>
      </div>
      <div id="meta" class="meta"></div>
      <div class="toolbar">
        <input id="filterInput" class="filter-input" type="text" spellcheck="false" placeholder="Filter commands and arguments (space-separated AND terms)" />
      </div>
      <div class="custom-command-bar">
        <input id="customCommandInput" class="custom-command-input" type="text" spellcheck="false" placeholder="Run any command" />
        <label class="capture-label"><input id="customCommandCapture" type="checkbox" /> Capture output</label>
        <button id="customCommandSendButton" class="send-button">Send</button>
      </div>
    </div>
    <div class="content">
      <div id="message" class="message"></div>
      <div id="filterSummary" class="filter-summary"></div>
      <div id="commands" class="command-list"></div>
    </div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const messageNode = document.getElementById("message");
      const commandsNode = document.getElementById("commands");
      const metaNode = document.getElementById("meta");
      const filterInput = document.getElementById("filterInput");
      const customCommandInput = document.getElementById("customCommandInput");
      const customCommandCapture = document.getElementById("customCommandCapture");
      const customCommandSendButton = document.getElementById("customCommandSendButton");
      const filterSummaryNode = document.getElementById("filterSummary");
      const showTerminalButton = document.getElementById("showTerminalButton");
      const startIocButton = document.getElementById("startIocButton");
      const stopIocButton = document.getElementById("stopIocButton");
      const initialState = ${initialStateJson};
      let currentState = initialState || {};
      let argumentsByCommand = new Map();
      let captureByCommand = new Map();
      let helpByCommand = new Map();
      let pendingHelpCommands = new Set();
      let customCommandText = "";
      let customCommandCaptureOutput = false;

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function normalizeFilterTerms(value) {
        return String(value || "")
          .toLowerCase()
          .split(/\\s+/)
          .map((term) => term.trim())
          .filter(Boolean);
      }

      function getFilteredCommands(state) {
        const payload = state || {};
        const commandList = Array.isArray(payload.commands) ? payload.commands : [];
        const filterTerms = normalizeFilterTerms(filterInput.value);
        if (!filterTerms.length) {
          return commandList;
        }

        return commandList.filter((commandName) => {
          const argumentText = argumentsByCommand.get(commandName) || "";
          const haystack = (
            String(commandName || "") + " " + String(argumentText || "")
          ).toLowerCase();
          return filterTerms.every((term) => haystack.includes(term));
        });
      }

      function sendCommand(commandName) {
        const input = document.querySelector('[data-command-input="' + CSS.escape(commandName) + '"]');
        const captureCheckbox = document.querySelector('[data-command-capture="' + CSS.escape(commandName) + '"]');
        const argumentsText = input instanceof HTMLInputElement ? input.value : "";
        const captureOutput = captureCheckbox instanceof HTMLInputElement
          ? captureCheckbox.checked
          : Boolean(captureByCommand.get(commandName));
        argumentsByCommand.set(commandName, argumentsText);
        captureByCommand.set(commandName, captureOutput);
        vscode.postMessage({
          type: "sendIocRuntimeCommand",
          commandName,
          argumentsText,
          captureOutput,
        });
      }

      function requestCommandHelp(commandName) {
        if (!commandName || pendingHelpCommands.has(commandName) || helpByCommand.has(commandName)) {
          return;
        }
        if (currentState?.isRunning === false) {
          return;
        }
        pendingHelpCommands.add(commandName);
        vscode.postMessage({
          type: "requestIocRuntimeCommandHelp",
          commandName,
        });
      }

      function getCommandHelpText(commandName) {
        const helpText = helpByCommand.get(commandName);
        if (helpText) {
          return helpText;
        }
        if (currentState?.isRunning === false) {
          return "Help is unavailable while the IOC is stopped.";
        }
        return "Loading help...";
      }

      function updateCommandPresentation(commandName) {
        if (!commandName) {
          return;
        }
        const inputElement = commandsNode.querySelector('[data-command-input="' + CSS.escape(commandName) + '"]');
        if (inputElement instanceof HTMLInputElement) {
          inputElement.placeholder = buildArgumentPlaceholder(commandName);
        }
        const helpElement = commandsNode.querySelector('[data-command-help="' + CSS.escape(commandName) + '"]');
        if (helpElement) {
          helpElement.textContent = getCommandHelpText(commandName);
        }
      }

      function extractParameterHints(commandName) {
        const helpText = helpByCommand.get(commandName);
        if (!helpText) {
          return [];
        }
        const firstNonEmptyLine = String(helpText)
          .split(/\\r?\\n/)
          .map((line) => line.trim())
          .find(Boolean);
        if (!firstNonEmptyLine) {
          return [];
        }
        const normalizedPrefix = String(commandName || "").trim() + " ";
        const signatureText = firstNonEmptyLine.startsWith(normalizedPrefix)
          ? firstNonEmptyLine.slice(normalizedPrefix.length)
          : firstNonEmptyLine === String(commandName || "").trim()
            ? ""
            : firstNonEmptyLine;
        const hints = [];
        const signaturePattern = /'([^']*)'|"([^"]*)"|([^\\s]+)/g;
        let match;
        while ((match = signaturePattern.exec(signatureText))) {
          const value = (match[1] || match[2] || match[3] || "").trim();
          if (value) {
            hints.push(value);
          }
        }
        return hints;
      }

      function countFilledParameters(value) {
        const text = String(value || "");
        const segments = [];
        let current = "";
        let quote = "";
        for (const character of text) {
          if (quote) {
            current += character;
            if (character === quote) {
              quote = "";
            }
            continue;
          }
          if (character === "'" || character === '"') {
            quote = character;
            current += character;
            continue;
          }
          if (character === ",") {
            segments.push(current);
            current = "";
            continue;
          }
          current += character;
        }
        segments.push(current);
        return segments.filter((segment) => String(segment || "").trim()).length;
      }

      function buildArgumentPlaceholder(commandName) {
        const parameterHints = extractParameterHints(commandName);
        if (!parameterHints.length) {
          return currentState?.isRunning === false
            ? "arguments, separated by commas"
            : "Loading parameter hints...";
        }
        const filledCount = countFilledParameters(argumentsByCommand.get(commandName) || "");
        const remainingHints = parameterHints.slice(filledCount);
        return remainingHints.length
          ? remainingHints.join(", ")
          : "";
      }

      function render(state) {
        currentState = state || {};
        const payload = currentState;
        const fullCommandList = Array.isArray(payload.commands) ? payload.commands : [];
        const commandList = getFilteredCommands(payload);
        metaNode.innerHTML =
          '<div>Startup file: <code>' + escapeHtml(payload.startupFileName || "") + '</code></div>' +
          '<div>Terminal: <code>' + escapeHtml(payload.terminalName || "") + '</code></div>';
        messageNode.textContent = payload.message || "";
        document.body.classList.toggle('stopped', payload.isRunning === false);
        filterInput.disabled = payload.isRunning === false;
        customCommandInput.disabled = payload.isRunning === false;
        customCommandCapture.disabled = payload.isRunning === false;
        customCommandSendButton.disabled = payload.isRunning === false || !String(customCommandText || "").trim();
        if (document.activeElement !== customCommandInput) {
          customCommandInput.value = customCommandText;
        }
        customCommandCapture.checked = customCommandCaptureOutput;
        startIocButton.style.display = payload.isRunning === false ? '' : 'none';
        startIocButton.disabled = payload.isRunning !== false;
        stopIocButton.disabled = payload.isRunning === false;
        stopIocButton.style.display = payload.isRunning === false ? 'none' : '';
        filterSummaryNode.textContent = fullCommandList.length
          ? ('Showing ' + String(commandList.length) + ' of ' + String(fullCommandList.length) + ' commands.')
          : '';

        if (!commandList.length) {
          commandsNode.innerHTML = '<div class="empty">' +
            escapeHtml(payload.isLoading ? 'Loading IOC command names...' : 'No IOC command names match the current filter.') +
            '</div>';
          return;
        }

        commandsNode.innerHTML = commandList.map((commandName) =>
          '<div class="command-card">' +
            '<div class="command-row">' +
              '<div class="command-name" data-help-command="' + escapeHtml(commandName) + '"><code>' + escapeHtml(commandName) + '</code></div>' +
              '<input class="args-input" data-command-input="' + escapeHtml(commandName) + '" type="text" spellcheck="false" placeholder="' + escapeHtml(buildArgumentPlaceholder(commandName)) + '" value="' + escapeHtml(argumentsByCommand.get(commandName) || "") + '"' + (payload.isRunning === false ? ' disabled' : '') + ' />' +
              '<label class="capture-label"><input data-command-capture="' + escapeHtml(commandName) + '" type="checkbox"' + (captureByCommand.get(commandName) ? ' checked' : '') + (payload.isRunning === false ? ' disabled' : '') + ' /> Capture output</label>' +
              '<button class="send-button" data-send-command="' + escapeHtml(commandName) + '"' + (payload.isRunning === false ? ' disabled' : '') + '>Send</button>' +
            '</div>' +
            '<div class="command-help" data-command-help="' + escapeHtml(commandName) + '">' + escapeHtml(getCommandHelpText(commandName)) + '</div>' +
          '</div>'
        ).join('');
      }

      commandsNode.addEventListener('click', (event) => {
        const target = event.target instanceof Element ? event.target.closest('[data-send-command]') : undefined;
        const commandName = target?.getAttribute('data-send-command');
        if (!commandName) {
          return;
        }
        sendCommand(commandName);
      });

      commandsNode.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter') {
          return;
        }
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const commandName = target?.getAttribute('data-command-input');
        if (!commandName) {
          return;
        }
        event.preventDefault();
        sendCommand(commandName);
      });

      commandsNode.addEventListener('input', (event) => {
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const commandName = target?.getAttribute('data-command-input');
        if (!commandName) {
          return;
        }
        argumentsByCommand.set(commandName, target.value);
        updateCommandPresentation(commandName);
      });

      commandsNode.addEventListener('change', (event) => {
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const commandName = target?.getAttribute('data-command-capture');
        if (!commandName) {
          return;
        }
        captureByCommand.set(commandName, target.checked);
      });

      filterInput.addEventListener('input', () => {
        render(currentState);
      });

      customCommandInput.addEventListener('input', () => {
        customCommandText = customCommandInput.value;
        customCommandSendButton.disabled =
          currentState?.isRunning === false || !String(customCommandText || "").trim();
      });

      customCommandCapture.addEventListener('change', () => {
        customCommandCaptureOutput = Boolean(customCommandCapture.checked);
      });

      function sendCustomCommand() {
        const trimmedCommand = String(customCommandInput.value || "").trim();
        if (!trimmedCommand || currentState?.isRunning === false) {
          return;
        }
        customCommandText = trimmedCommand;
        vscode.postMessage({
          type: 'sendCustomIocRuntimeCommand',
          commandText: trimmedCommand,
          captureOutput: customCommandCaptureOutput,
        });
      }

      customCommandInput.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter') {
          return;
        }
        event.preventDefault();
        sendCustomCommand();
      });

      customCommandSendButton.addEventListener('click', () => {
        sendCustomCommand();
      });

      showTerminalButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'showIocRuntimeCommandsTerminal' });
      });

      startIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'startIocRuntimeCommandsStartup' });
      });

      stopIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'stopIocRuntimeCommandsStartup' });
      });

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'iocRuntimeCommandsState') {
          render(event.data.state);
          return;
        }
        if (event.data?.type === 'iocRuntimeCommandHelp' && event.data.commandName) {
          pendingHelpCommands.delete(event.data.commandName);
          helpByCommand.set(event.data.commandName, event.data.helpText || 'No detailed help is available.');
          updateCommandPresentation(event.data.commandName);
        }
      });

      try {
        render(initialState);
      } catch (error) {
        messageNode.classList.add('error');
        messageNode.textContent = 'IOC Runtime Commands page error: ' + String(error && error.message ? error.message : error);
      }
    </script>
  </body>
</html>`;
}

function buildIocRuntimeVariablesHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>IOC Runtime Variables</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .header {
        padding: 16px 20px 14px;
        border-bottom: 1px solid var(--vscode-panel-border);
      }
      .header-top {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        justify-content: space-between;
      }
      .title {
        margin: 0;
        font-size: 1.35rem;
        font-weight: 600;
      }
      .meta {
        margin-top: 8px;
        color: var(--vscode-descriptionForeground);
        line-height: 1.5;
      }
      .toolbar {
        margin-top: 14px;
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .toolbar-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .filter-input {
        width: 100%;
        min-width: 26ch;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .header-button,
      .set-button {
        padding: 8px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 6px;
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        cursor: pointer;
        font: inherit;
      }
      .header-button:hover,
      .set-button:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .header-button:disabled,
      .set-button:disabled,
      .value-input:disabled {
        opacity: 0.55;
        cursor: default;
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 18px 20px 24px;
      }
      .message {
        margin-bottom: 16px;
        color: var(--vscode-descriptionForeground);
      }
      .message.error {
        color: var(--vscode-errorForeground);
      }
      .filter-summary {
        margin-bottom: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .warning-note {
        margin-bottom: 12px;
        color: var(--vscode-errorForeground);
      }
      .variable-list {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      .variable-card {
        display: flex;
        flex-direction: column;
        gap: 0;
        border: 1px solid var(--vscode-panel-border);
        border-radius: 8px;
        background: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-editorHoverWidget-background, var(--vscode-sideBar-background)) 10%);
        overflow: hidden;
      }
      .variable-row {
        display: grid;
        grid-template-columns: minmax(18ch, 22ch) minmax(24ch, 1fr) auto;
        gap: 10px;
        align-items: center;
        padding: 10px 12px 8px;
      }
      .variable-name {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
        font-size: 0.98rem;
        font-weight: 600;
      }
      .value-input {
        width: 100%;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .variable-meta {
        padding: 0 12px 12px;
        color: var(--vscode-descriptionForeground);
        white-space: pre-wrap;
        line-height: 1.45;
        font-size: 0.93rem;
      }
      .variable-meta::before {
        content: "";
        display: block;
        margin-bottom: 8px;
        border-top: 1px solid var(--vscode-panel-border);
      }
      .empty {
        color: var(--vscode-descriptionForeground);
      }
      body.stopped .meta,
      body.stopped .toolbar,
      body.stopped .content {
        opacity: 0.45;
        filter: grayscale(0.25);
      }
      .stopped-watermark {
        display: none;
        position: fixed;
        inset: 170px 0 0 0;
        pointer-events: none;
        overflow: hidden;
        z-index: 2;
      }
      body.stopped .stopped-watermark {
        display: block;
      }
      .stopped-watermark-grid {
        position: absolute;
        inset: 0 -10% 0 -10%;
        display: grid;
        grid-template-columns: repeat(4, minmax(18ch, 1fr));
        gap: 34px 48px;
        align-content: start;
        transform: rotate(-18deg);
      }
      .stopped-watermark-text {
        color: transparent;
        font-size: 3.1rem;
        font-style: italic;
        font-weight: 700;
        letter-spacing: 0.08em;
        -webkit-text-stroke: 1.3px rgba(220, 38, 38, 0.65);
        user-select: none;
      }
      @media (max-width: 900px) {
        .variable-row {
          grid-template-columns: 1fr;
        }
        .stopped-watermark-grid {
          grid-template-columns: repeat(2, minmax(18ch, 1fr));
        }
      }
    </style>
  </head>
  <body>
    <div class="stopped-watermark" aria-hidden="true">
      <div class="stopped-watermark-grid">
        ${Array.from({ length: 24 }, () => '<div class="stopped-watermark-text">Stopped</div>').join("")}
      </div>
    </div>
    <div class="header">
      <div class="header-top">
        <h1 class="title">IOC Runtime Variables</h1>
        <div class="toolbar-actions">
          <button id="showTerminalButton" class="header-button">Show Running Terminal</button>
          <button id="startIocButton" class="header-button">Start IOC</button>
          <button id="stopIocButton" class="header-button">Stop IOC</button>
        </div>
      </div>
      <div id="meta" class="meta"></div>
      <div class="toolbar">
        <input id="filterInput" class="filter-input" type="text" spellcheck="false" placeholder="Filter variables by name, type, or value" />
      </div>
    </div>
    <div class="content">
      <div id="message" class="message"></div>
      <div id="filterSummary" class="filter-summary"></div>
      <div class="warning-note">Note: the current value may not reflect the real value.</div>
      <div id="variables" class="variable-list"></div>
    </div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const messageNode = document.getElementById("message");
      const variablesNode = document.getElementById("variables");
      const metaNode = document.getElementById("meta");
      const filterInput = document.getElementById("filterInput");
      const filterSummaryNode = document.getElementById("filterSummary");
      const showTerminalButton = document.getElementById("showTerminalButton");
      const startIocButton = document.getElementById("startIocButton");
      const stopIocButton = document.getElementById("stopIocButton");
      const initialState = ${initialStateJson};
      let currentState = initialState || {};
      let valuesByVariable = new Map();

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function normalizeFilterTerms(value) {
        return String(value || "")
          .toLowerCase()
          .split(/\\s+/)
          .map((term) => term.trim())
          .filter(Boolean);
      }

      function getFilteredVariables(state) {
        const payload = state || {};
        const variableList = Array.isArray(payload.variables) ? payload.variables : [];
        const filterTerms = normalizeFilterTerms(filterInput.value);
        if (!filterTerms.length) {
          return variableList;
        }

        return variableList.filter((entry) => {
          const variableName = String(entry?.variableName || "");
          const variableType = String(entry?.variableType || "");
          const currentValue = String(entry?.currentValue || "");
          const draftValue = valuesByVariable.get(variableName) || "";
          const haystack = (
            variableName + " " + variableType + " " + currentValue + " " + draftValue
          ).toLowerCase();
          return filterTerms.every((term) => haystack.includes(term));
        });
      }

      function sendVariable(variableName) {
        const input = document.querySelector('[data-variable-input="' + CSS.escape(variableName) + '"]');
        const valueText = input instanceof HTMLInputElement ? input.value.trim() : "";
        valuesByVariable.delete(variableName);
        vscode.postMessage({
          type: "setIocRuntimeVariable",
          variableName,
          valueText,
        });
      }

      function render(state) {
        currentState = state || {};
        const payload = currentState;
        const fullVariableList = Array.isArray(payload.variables) ? payload.variables : [];
        const variableList = getFilteredVariables(payload);
        metaNode.innerHTML =
          '<div>Startup file: <code>' + escapeHtml(payload.startupFileName || "") + '</code></div>' +
          '<div>Terminal: <code>' + escapeHtml(payload.terminalName || "") + '</code></div>';
        messageNode.textContent = payload.message || "";
        document.body.classList.toggle('stopped', payload.isRunning === false);
        filterInput.disabled = payload.isRunning === false;
        startIocButton.style.display = payload.isRunning === false ? '' : 'none';
        startIocButton.disabled = payload.isRunning !== false;
        stopIocButton.style.display = payload.isRunning === false ? 'none' : '';
        stopIocButton.disabled = payload.isRunning === false;
        filterSummaryNode.textContent = fullVariableList.length
          ? ('Showing ' + String(variableList.length) + ' of ' + String(fullVariableList.length) + ' variables.')
          : '';

        if (!variableList.length) {
          variablesNode.innerHTML = '<div class="empty">' +
            escapeHtml(payload.isLoading ? 'Loading IOC runtime variables...' : 'No IOC runtime variables match the current filter.') +
            '</div>';
          return;
        }

        for (const entry of fullVariableList) {
          const variableName = String(entry?.variableName || "");
          if (!variableName || valuesByVariable.has(variableName)) {
            continue;
          }
          valuesByVariable.set(variableName, String(entry?.currentValue || ""));
        }

        variablesNode.innerHTML = variableList.map((entry) => {
          const variableName = String(entry?.variableName || "");
          const variableType = String(entry?.variableType || "");
          const currentValue = String(entry?.currentValue || "");
          return '<div class="variable-card">' +
            '<div class="variable-row">' +
              '<div class="variable-name"><code>' + escapeHtml(variableName) + '</code></div>' +
              '<input class="value-input" data-variable-input="' + escapeHtml(variableName) + '" type="text" spellcheck="false" value="' + escapeHtml(valuesByVariable.get(variableName) || "") + '"' + (payload.isRunning === false ? ' disabled' : '') + ' />' +
              '<button class="set-button" data-set-variable="' + escapeHtml(variableName) + '"' + (payload.isRunning === false ? ' disabled' : '') + '>Set</button>' +
            '</div>' +
            '<div class="variable-meta">Type: <code>' + escapeHtml(variableType || "unknown") + '</code> | Current value: <code>' + escapeHtml(currentValue) + '</code></div>' +
          '</div>';
        }).join('');
      }

      variablesNode.addEventListener('click', (event) => {
        const target = event.target instanceof Element ? event.target.closest('[data-set-variable]') : undefined;
        const variableName = target?.getAttribute('data-set-variable');
        if (!variableName) {
          return;
        }
        sendVariable(variableName);
      });

      variablesNode.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter') {
          return;
        }
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const variableName = target?.getAttribute('data-variable-input');
        if (!variableName) {
          return;
        }
        event.preventDefault();
        sendVariable(variableName);
      });

      variablesNode.addEventListener('input', (event) => {
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const variableName = target?.getAttribute('data-variable-input');
        if (!variableName) {
          return;
        }
        valuesByVariable.set(variableName, target.value);
      });

      filterInput.addEventListener('input', () => {
        render(currentState);
      });

      showTerminalButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'showIocRuntimeVariablesTerminal' });
      });

      startIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'startIocRuntimeVariablesStartup' });
      });

      stopIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'stopIocRuntimeVariablesStartup' });
      });

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'iocRuntimeVariablesState') {
          render(event.data.state);
        }
      });

      try {
        render(initialState);
      } catch (error) {
        messageNode.classList.add('error');
        messageNode.textContent = 'IOC Runtime Variables page error: ' + String(error && error.message ? error.message : error);
      }
    </script>
  </body>
</html>`;
}

function buildIocRuntimeEnvironmentHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>IOC Runtime Environment</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .header {
        padding: 16px 20px 14px;
        border-bottom: 1px solid var(--vscode-panel-border);
      }
      .header-top {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        justify-content: space-between;
      }
      .title {
        margin: 0;
        font-size: 1.35rem;
        font-weight: 600;
      }
      .meta {
        margin-top: 8px;
        color: var(--vscode-descriptionForeground);
        line-height: 1.5;
      }
      .toolbar {
        margin-top: 14px;
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .toolbar-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        align-items: center;
      }
      .filter-input {
        width: 100%;
        min-width: 26ch;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .header-button,
      .set-button {
        padding: 8px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 6px;
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        cursor: pointer;
        font: inherit;
      }
      .header-button:hover,
      .set-button:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .header-button:disabled,
      .set-button:disabled,
      .value-input:disabled {
        opacity: 0.55;
        cursor: default;
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 18px 20px 24px;
      }
      .message {
        margin-bottom: 16px;
        color: var(--vscode-descriptionForeground);
      }
      .message.error {
        color: var(--vscode-errorForeground);
      }
      .filter-summary {
        margin-bottom: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .warning-note {
        margin-bottom: 12px;
        color: var(--vscode-errorForeground);
      }
      .entry-list {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      .entry-card {
        display: flex;
        flex-direction: column;
        gap: 0;
        border: 1px solid var(--vscode-panel-border);
        border-radius: 8px;
        background: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-editorHoverWidget-background, var(--vscode-sideBar-background)) 10%);
        overflow: hidden;
      }
      .entry-row {
        display: grid;
        grid-template-columns: minmax(18ch, 24ch) minmax(28ch, 1fr) auto;
        gap: 10px;
        align-items: center;
        padding: 10px 12px 8px;
      }
      .entry-name {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
        font-size: 0.98rem;
        font-weight: 600;
      }
      .value-input {
        width: 100%;
        box-sizing: border-box;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        border-radius: 6px;
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        padding: 8px 10px;
        font: inherit;
      }
      .entry-meta {
        padding: 0 12px 12px;
        color: var(--vscode-descriptionForeground);
        white-space: pre-wrap;
        line-height: 1.45;
        font-size: 0.93rem;
      }
      .entry-meta::before {
        content: "";
        display: block;
        margin-bottom: 8px;
        border-top: 1px solid var(--vscode-panel-border);
      }
      .empty {
        color: var(--vscode-descriptionForeground);
      }
      body.stopped .meta,
      body.stopped .toolbar,
      body.stopped .content {
        opacity: 0.45;
        filter: grayscale(0.25);
      }
      .stopped-watermark {
        display: none;
        position: fixed;
        inset: 170px 0 0 0;
        pointer-events: none;
        overflow: hidden;
        z-index: 2;
      }
      body.stopped .stopped-watermark {
        display: block;
      }
      .stopped-watermark-grid {
        position: absolute;
        inset: 0 -10% 0 -10%;
        display: grid;
        grid-template-columns: repeat(4, minmax(18ch, 1fr));
        gap: 34px 48px;
        align-content: start;
        transform: rotate(-18deg);
      }
      .stopped-watermark-text {
        color: transparent;
        font-size: 3.1rem;
        font-style: italic;
        font-weight: 700;
        letter-spacing: 0.08em;
        -webkit-text-stroke: 1.3px rgba(220, 38, 38, 0.65);
        user-select: none;
      }
      @media (max-width: 900px) {
        .entry-row {
          grid-template-columns: 1fr;
        }
        .stopped-watermark-grid {
          grid-template-columns: repeat(2, minmax(18ch, 1fr));
        }
      }
    </style>
  </head>
  <body>
    <div class="stopped-watermark" aria-hidden="true">
      <div class="stopped-watermark-grid">
        ${Array.from({ length: 24 }, () => '<div class="stopped-watermark-text">Stopped</div>').join("")}
      </div>
    </div>
    <div class="header">
      <div class="header-top">
        <h1 class="title">IOC Runtime Environment</h1>
        <div class="toolbar-actions">
          <button id="showTerminalButton" class="header-button">Show Running Terminal</button>
          <button id="startIocButton" class="header-button">Start IOC</button>
          <button id="stopIocButton" class="header-button">Stop IOC</button>
        </div>
      </div>
      <div id="meta" class="meta"></div>
      <div class="toolbar">
        <input id="filterInput" class="filter-input" type="text" spellcheck="false" placeholder="Filter environment variables by name or value" />
      </div>
    </div>
    <div class="content">
      <div id="message" class="message"></div>
      <div id="filterSummary" class="filter-summary"></div>
      <div class="warning-note">Note: the current value may not reflect the real value.</div>
      <div id="entries" class="entry-list"></div>
    </div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const messageNode = document.getElementById("message");
      const entriesNode = document.getElementById("entries");
      const metaNode = document.getElementById("meta");
      const filterInput = document.getElementById("filterInput");
      const filterSummaryNode = document.getElementById("filterSummary");
      const showTerminalButton = document.getElementById("showTerminalButton");
      const startIocButton = document.getElementById("startIocButton");
      const stopIocButton = document.getElementById("stopIocButton");
      const initialState = ${initialStateJson};
      let currentState = initialState || {};
      let valuesByName = new Map();

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function normalizeFilterTerms(value) {
        return String(value || "")
          .toLowerCase()
          .split(/\\s+/)
          .map((term) => term.trim())
          .filter(Boolean);
      }

      function getFilteredEntries(state) {
        const payload = state || {};
        const entryList = Array.isArray(payload.entries) ? payload.entries : [];
        const filterTerms = normalizeFilterTerms(filterInput.value);
        if (!filterTerms.length) {
          return entryList;
        }

        return entryList.filter((entry) => {
          const variableName = String(entry?.variableName || "");
          const currentValue = String(entry?.currentValue || "");
          const draftValue = valuesByName.get(variableName) || "";
          const haystack = (
            variableName + " " + currentValue + " " + draftValue
          ).toLowerCase();
          return filterTerms.every((term) => haystack.includes(term));
        });
      }

      function sendEntry(variableName) {
        const input = document.querySelector('[data-entry-input="' + CSS.escape(variableName) + '"]');
        const valueText = input instanceof HTMLInputElement ? input.value : "";
        valuesByName.set(variableName, valueText);
        vscode.postMessage({
          type: "setIocRuntimeEnvironmentValue",
          variableName,
          valueText,
        });
      }

      function render(state) {
        currentState = state || {};
        const payload = currentState;
        const fullEntryList = Array.isArray(payload.entries) ? payload.entries : [];
        const entryList = getFilteredEntries(payload);
        metaNode.innerHTML =
          '<div>Startup file: <code>' + escapeHtml(payload.startupFileName || "") + '</code></div>' +
          '<div>Terminal: <code>' + escapeHtml(payload.terminalName || "") + '</code></div>';
        messageNode.textContent = payload.message || "";
        document.body.classList.toggle('stopped', payload.isRunning === false);
        filterInput.disabled = payload.isRunning === false;
        startIocButton.style.display = payload.isRunning === false ? '' : 'none';
        startIocButton.disabled = payload.isRunning !== false;
        stopIocButton.style.display = payload.isRunning === false ? 'none' : '';
        stopIocButton.disabled = payload.isRunning === false;
        filterSummaryNode.textContent = fullEntryList.length
          ? ('Showing ' + String(entryList.length) + ' of ' + String(fullEntryList.length) + ' environment variables.')
          : '';

        if (!entryList.length) {
          entriesNode.innerHTML = '<div class="empty">' +
            escapeHtml(payload.isLoading ? 'Loading IOC runtime environment...' : 'No IOC runtime environment variables match the current filter.') +
            '</div>';
          return;
        }

        for (const entry of fullEntryList) {
          const variableName = String(entry?.variableName || "");
          if (!variableName || valuesByName.has(variableName)) {
            continue;
          }
          valuesByName.set(variableName, String(entry?.currentValue || ""));
        }

        entriesNode.innerHTML = entryList.map((entry) => {
          const variableName = String(entry?.variableName || "");
          const currentValue = String(entry?.currentValue || "");
          return '<div class="entry-card">' +
            '<div class="entry-row">' +
              '<div class="entry-name"><code>' + escapeHtml(variableName) + '</code></div>' +
              '<input class="value-input" data-entry-input="' + escapeHtml(variableName) + '" type="text" spellcheck="false" value="' + escapeHtml(valuesByName.get(variableName) || "") + '"' + (payload.isRunning === false ? ' disabled' : '') + ' />' +
              '<button class="set-button" data-set-entry="' + escapeHtml(variableName) + '"' + (payload.isRunning === false ? ' disabled' : '') + '>Set</button>' +
            '</div>' +
            '<div class="entry-meta">Current value: <code>' + escapeHtml(currentValue) + '</code></div>' +
          '</div>';
        }).join('');
      }

      entriesNode.addEventListener('click', (event) => {
        const target = event.target instanceof Element ? event.target.closest('[data-set-entry]') : undefined;
        const variableName = target?.getAttribute('data-set-entry');
        if (!variableName) {
          return;
        }
        sendEntry(variableName);
      });

      entriesNode.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter') {
          return;
        }
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const variableName = target?.getAttribute('data-entry-input');
        if (!variableName) {
          return;
        }
        event.preventDefault();
        sendEntry(variableName);
      });

      entriesNode.addEventListener('input', (event) => {
        const target = event.target instanceof HTMLInputElement ? event.target : undefined;
        const variableName = target?.getAttribute('data-entry-input');
        if (!variableName) {
          return;
        }
        valuesByName.set(variableName, target.value);
      });

      filterInput.addEventListener('input', () => {
        render(currentState);
      });

      showTerminalButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'showIocRuntimeEnvironmentTerminal' });
      });

      startIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'startIocRuntimeEnvironmentStartup' });
      });

      stopIocButton.addEventListener('click', () => {
        vscode.postMessage({ type: 'stopIocRuntimeEnvironmentStartup' });
      });

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'iocRuntimeEnvironmentState') {
          render(event.data.state);
        }
      });

      try {
        render(initialState);
      } catch (error) {
        messageNode.classList.add('error');
        messageNode.textContent = 'IOC Runtime Environment page error: ' + String(error && error.message ? error.message : error);
      }
    </script>
  </body>
</html>`;
}

function buildProbeWidgetHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Probe</title>
    <style>
      :root {
        color-scheme: light dark;
      }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .toolbar {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px 16px;
        border-bottom: 1px solid var(--vscode-panel-border);
        position: sticky;
        top: 0;
        background: var(--vscode-editor-background);
        z-index: 1;
      }
      .toolbar label {
        color: var(--vscode-descriptionForeground);
      }
      .toolbar input {
        min-width: 280px;
        padding: 6px 10px;
        border: 1px solid var(--vscode-input-border, transparent);
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        border-radius: 4px;
        font: inherit;
      }
      button {
        border: 1px solid var(--vscode-button-border, transparent);
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        padding: 6px 12px;
        border-radius: 4px;
        cursor: pointer;
      }
      button:disabled {
        opacity: 0.5;
        cursor: default;
      }
      .status {
        color: var(--vscode-descriptionForeground);
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 20px 24px 28px;
      }
      h1, h2, p {
        margin: 0;
      }
      .message {
        margin-top: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .error {
        margin-top: 12px;
        color: var(--vscode-errorForeground);
      }
      .meta-grid {
        display: grid;
        grid-template-columns: max-content 1fr;
        gap: 10px 16px;
        margin-top: 16px;
      }
      .meta-label {
        color: var(--vscode-descriptionForeground);
      }
      .fields {
        margin-top: 24px;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        margin-top: 10px;
      }
      th, td {
        text-align: left;
        padding: 8px 10px;
        border-bottom: 1px solid var(--vscode-panel-border);
        vertical-align: top;
      }
      th {
        color: var(--vscode-descriptionForeground);
        font-weight: 600;
      }
      td:first-child, th:first-child {
        width: 18ch;
      }
      code {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
      }
      .put-target {
        cursor: pointer;
      }
      .put-target:hover {
        text-decoration: underline;
      }
      .empty {
        margin-top: 12px;
        color: var(--vscode-descriptionForeground);
      }
    </style>
  </head>
  <body>
    <div class="toolbar">
      <label for="channelInput">Channel</label>
      <input id="channelInput" type="text" spellcheck="false" />
      <button id="processButton" type="button">Process</button>
      <span id="statusText" class="status"></span>
    </div>
    <div id="content" class="content"></div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const doubleClickIntervalMs = ${JSON.stringify(MOUSE_DOUBLE_CLICK_INTERVAL_MS)};
      const initialState = ${initialStateJson};
      const content = document.getElementById("content");
      const channelInput = document.getElementById("channelInput");
      const processButton = document.getElementById("processButton");
      const statusText = document.getElementById("statusText");
      let pendingPutClick = undefined;

      channelInput.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") {
          return;
        }
        vscode.postMessage({
          type: "updateProbeWidgetRecordName",
          recordName: channelInput.value,
        });
      });

      processButton.addEventListener("click", () => {
        vscode.postMessage({ type: "processProbeWidget" });
      });

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function render(payload) {
        const state = payload?.state;
        const title = escapeHtml(payload?.recordName || "EPICS Probe");
        if (document.activeElement !== channelInput) {
          channelInput.value = payload?.recordName || "";
        }
        processButton.disabled = !payload?.canProcess;
        statusText.textContent =
          payload?.contextStatus === "running" && payload?.recordName
            ? "Running"
            : payload?.recordName
              ? "Stopped"
              : "Idle";

        const message = payload?.message
          ? '<p class="message">' + escapeHtml(payload.message) + '</p>'
          : '';

        if (!state) {
          content.innerHTML = '<h1>' + title + '</h1>' + message;
          return;
        }

        const fieldRows = (state.fields || []).map((field) => {
          const putClass = field.canPut ? 'put-target' : '';
          const putTitle = field.canPut ? ' title="Double-click to put a new value"' : '';
          return '<tr>' +
            '<td><code>' + escapeHtml(field.fieldName) + '</code></td>' +
            '<td class="' + putClass + '" data-key="' + escapeHtml(field.key) + '" data-can-put="' + (field.canPut ? "true" : "false") + '"' + putTitle + '>' + escapeHtml(field.value) + '</td>' +
            '</tr>';
        }).join('');

        const valueClass = state.valueCanPut ? 'put-target' : '';
        const valueTitle = state.valueCanPut ? ' title="Double-click to put a new value"' : '';
        content.innerHTML =
          '<h1>' + title + '</h1>' +
          message +
          '<div class="meta-grid">' +
            '<div class="meta-label">Value</div><div class="' + valueClass + '" data-key="' + escapeHtml(state.valueKey || '') + '" data-can-put="' + (state.valueCanPut ? "true" : "false") + '"' + valueTitle + '>' + escapeHtml(state.value) + '</div>' +
            '<div class="meta-label">Record Type</div><div><code>' + escapeHtml(state.recordType) + '</code></div>' +
            '<div class="meta-label">Last Update</div><div>' + escapeHtml(state.lastUpdated) + '</div>' +
            '<div class="meta-label">Permission</div><div>' + escapeHtml(state.access) + '</div>' +
          '</div>' +
          '<div class="fields">' +
            '<h2>Fields</h2>' +
            (fieldRows
              ? '<table><thead><tr><th>Field</th><th>Value</th></tr></thead><tbody>' + fieldRows + '</tbody></table>'
              : '<p class="empty">' + escapeHtml(state.fieldStatusText || 'No fields loaded.') + '</p>') +
          '</div>';
      }

      content.addEventListener('click', (event) => {
        const eventTarget = event.target instanceof Element
          ? event.target
          : event.target?.parentElement;
        const target = eventTarget?.closest?.('[data-key][data-can-put="true"]');
        if (!target) {
          pendingPutClick = undefined;
          return;
        }
        const key = target.getAttribute('data-key');
        if (!key) {
          pendingPutClick = undefined;
          return;
        }
        const now = Date.now();
        const isDoubleClick =
          pendingPutClick?.key === key &&
          now - pendingPutClick.time <= doubleClickIntervalMs;
        pendingPutClick = isDoubleClick
          ? undefined
          : { key, time: now };
        if (!isDoubleClick) {
          return;
        }
        vscode.postMessage({
          type: 'putProbeValue',
          key,
        });
      });

      window.addEventListener('message', (event) => {
        if (event.data?.type === 'probeWidgetState') {
          render(event.data.state);
        }
      });

      render(initialState);
    </script>
  </body>
</html>`;
}

function buildMonitorWidgetHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Monitor</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body {
        display: flex;
        flex-direction: column;
      }
      .header {
        border-bottom: 1px solid var(--vscode-panel-border);
        background: var(--vscode-editor-background);
        padding: 12px 16px;
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      .toolbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
      }
      .toolbar-left, .toolbar-right {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .toolbar-title { font-weight: 600; }
      .toolbar-status { color: var(--vscode-descriptionForeground); }
      .action-button {
        padding: 6px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 4px;
        background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
        color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
        cursor: pointer;
        font: inherit;
      }
      .action-button:hover {
        background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground));
      }
      .channel-rows {
        display: flex;
        flex-direction: column;
        gap: 8px;
      }
      .channel-row {
        display: flex;
        align-items: center;
        gap: 12px;
      }
      .channel-label {
        color: var(--vscode-descriptionForeground);
        min-width: 6ch;
      }
      .channel-input, .config-input {
        flex: 1;
        padding: 6px 10px;
        border: 1px solid var(--vscode-input-border, transparent);
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        border-radius: 4px;
        font: inherit;
        box-sizing: border-box;
      }
      .channel-status {
        color: var(--vscode-descriptionForeground);
        min-width: 16ch;
      }
      .content {
        flex: 1;
        overflow: auto;
        padding: 16px 24px 24px;
      }
      .message {
        margin-bottom: 12px;
        color: var(--vscode-descriptionForeground);
      }
      .history {
        margin: 0;
        white-space: pre;
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
        font-size: var(--vscode-editor-font-size, inherit);
        line-height: 1.45;
      }
      .empty {
        color: var(--vscode-descriptionForeground);
      }
      .overlay-page {
        min-height: 100%;
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .overlay-title {
        margin: 0;
        font-size: 1.2rem;
        font-weight: 600;
      }
      .overlay-hint {
        color: var(--vscode-descriptionForeground);
      }
      .config-row {
        display: flex;
        align-items: center;
        gap: 12px;
        max-width: 320px;
      }
    </style>
  </head>
  <body>
    <div class="header">
      <div class="toolbar">
        <div class="toolbar-left">
          <div class="toolbar-title">Monitor</div>
          <div id="statusText" class="toolbar-status"></div>
        </div>
        <div class="toolbar-right">
          <button id="addChannelButton" class="action-button">Add Channel</button>
          <button id="configureButton" class="action-button">Configure</button>
          <button id="exportButton" class="action-button">Export Data</button>
          <button id="doneButton" class="action-button" hidden>Done</button>
        </div>
      </div>
      <div id="channelRows" class="channel-rows"></div>
    </div>
    <div id="content" class="content"></div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      const initialState = ${initialStateJson};
      const content = document.getElementById("content");
      const channelRows = document.getElementById("channelRows");
      const statusText = document.getElementById("statusText");
      const addChannelButton = document.getElementById("addChannelButton");
      const configureButton = document.getElementById("configureButton");
      const exportButton = document.getElementById("exportButton");
      const doneButton = document.getElementById("doneButton");
      let currentState = initialState;
      let overlayMode = false;
      let editingState = undefined;
      let suppressFocusSync = false;
      let isRendering = false;

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function captureEditingState() {
        const activeElement = document.activeElement;
        if (!activeElement || !(activeElement instanceof HTMLInputElement)) {
          return;
        }
        if (activeElement.classList.contains("channel-input")) {
          editingState = {
            type: "channel",
            rowId: activeElement.dataset.rowId,
            value: activeElement.value,
            selectionStart: activeElement.selectionStart,
            selectionEnd: activeElement.selectionEnd,
          };
          return;
        }
        if (activeElement.classList.contains("config-input")) {
          editingState = {
            type: "buffer",
            value: activeElement.value,
            selectionStart: activeElement.selectionStart,
            selectionEnd: activeElement.selectionEnd,
          };
        }
      }

      function restoreActiveInput() {
        if (!editingState?.type) {
          return;
        }

        let input = undefined;
        if (editingState.type === "channel") {
          input = [...channelRows.querySelectorAll(".channel-input")].find(
            (candidate) => candidate.dataset.rowId === editingState.rowId,
          );
        } else if (editingState.type === "buffer") {
          input = content.querySelector(".config-input");
        }
        if (!input) {
          return;
        }

        const fallbackCaret = String(editingState.value || "").length;
        const selectionStart =
          typeof editingState.selectionStart === "number"
            ? editingState.selectionStart
            : fallbackCaret;
        const selectionEnd =
          typeof editingState.selectionEnd === "number"
            ? editingState.selectionEnd
            : fallbackCaret;
        suppressFocusSync = true;
        input.focus();
        suppressFocusSync = false;
        input.setSelectionRange(selectionStart, selectionEnd);
      }

      function updateEditingSelection(input) {
        if (!editingState) {
          return;
        }
        if (editingState.type === "channel" && editingState.rowId === input.dataset.rowId) {
          editingState = {
            ...editingState,
            value: input.value,
            selectionStart: input.selectionStart,
            selectionEnd: input.selectionEnd,
          };
        } else if (editingState.type === "buffer" && input.classList.contains("config-input")) {
          editingState = {
            ...editingState,
            value: input.value,
            selectionStart: input.selectionStart,
            selectionEnd: input.selectionEnd,
          };
        }
      }

      function commitBufferSizeAndClose() {
        const input = content.querySelector(".config-input");
        vscode.postMessage({
          type: "updateMonitorWidgetBufferSize",
          value: input ? input.value : currentState?.bufferSize,
        });
        overlayMode = false;
        editingState = undefined;
        render(currentState);
      }

      addChannelButton.addEventListener("click", () => {
        vscode.postMessage({ type: "addMonitorWidgetChannelRow" });
      });

      configureButton.addEventListener("click", () => {
        overlayMode = true;
        render(currentState);
      });

      exportButton.addEventListener("click", () => {
        vscode.postMessage({ type: "exportMonitorWidgetData" });
      });

      doneButton.addEventListener("click", () => {
        commitBufferSizeAndClose();
      });

      channelRows.addEventListener("focusin", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("channel-input") || suppressFocusSync) {
          return;
        }
        editingState = {
          type: "channel",
          rowId: input.dataset.rowId,
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      });

      channelRows.addEventListener("input", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("channel-input")) {
          return;
        }
        editingState = {
          type: "channel",
          rowId: input.dataset.rowId,
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      });

      channelRows.addEventListener("keydown", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("channel-input")) {
          return;
        }
        if (event.key !== "Enter") {
          return;
        }
        event.preventDefault();
        vscode.postMessage({
          type: "updateMonitorWidgetChannel",
          rowId: input.dataset.rowId,
          channelName: input.value,
        });
      });

      channelRows.addEventListener("keyup", (event) => {
        const input = event.target;
        if (input instanceof HTMLInputElement && input.classList.contains("channel-input")) {
          updateEditingSelection(input);
        }
      });

      channelRows.addEventListener("mouseup", (event) => {
        const input = event.target;
        if (input instanceof HTMLInputElement && input.classList.contains("channel-input")) {
          updateEditingSelection(input);
        }
      });

      channelRows.addEventListener("focusout", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("channel-input")) {
          return;
        }
        if (isRendering) {
          return;
        }
        if (editingState?.type === "channel" && editingState.rowId === input.dataset.rowId) {
          editingState = undefined;
        }
      });

      content.addEventListener("focusin", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("config-input") || suppressFocusSync) {
          return;
        }
        editingState = {
          type: "buffer",
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      });

      content.addEventListener("input", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("config-input")) {
          return;
        }
        editingState = {
          type: "buffer",
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      });

      content.addEventListener("keydown", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("config-input")) {
          return;
        }
        if (event.key !== "Enter") {
          return;
        }
        event.preventDefault();
        commitBufferSizeAndClose();
      });

      content.addEventListener("keyup", (event) => {
        const input = event.target;
        if (input instanceof HTMLInputElement && input.classList.contains("config-input")) {
          updateEditingSelection(input);
        }
      });

      content.addEventListener("mouseup", (event) => {
        const input = event.target;
        if (input instanceof HTMLInputElement && input.classList.contains("config-input")) {
          updateEditingSelection(input);
        }
      });

      content.addEventListener("focusout", (event) => {
        const input = event.target;
        if (!(input instanceof HTMLInputElement) || !input.classList.contains("config-input")) {
          return;
        }
        if (isRendering) {
          return;
        }
        if (editingState?.type === "buffer") {
          editingState = undefined;
        }
      });

      function render(payload) {
        captureEditingState();
        currentState = payload || {};
        isRendering = true;

        const wasPinnedToBottom =
          !overlayMode &&
          content.scrollTop + content.clientHeight >= content.scrollHeight - 16;
        const previousScrollTop = content.scrollTop;

        statusText.textContent =
          payload?.contextStatus === "connected"
            ? "Running"
            : payload?.contextStatus === "connecting"
              ? "Connecting"
              : payload?.contextStatus === "error"
                ? "Error"
                : "";
        addChannelButton.hidden = overlayMode;
        configureButton.hidden = overlayMode;
        exportButton.hidden = overlayMode;
        doneButton.hidden = !overlayMode;

        const rows = Array.isArray(payload?.rows) ? payload.rows : [];
        channelRows.innerHTML = rows.map((row, index) => {
          const currentValue =
            editingState?.type === "channel" && editingState.rowId === row.id
              ? editingState.value
              : row.channelName || "";
          return '<div class="channel-row">' +
            '<div class="channel-label">Channel ' + String(index + 1) + '</div>' +
            '<input class="channel-input" data-row-id="' + escapeHtml(row.id) + '" type="text" spellcheck="false" value="' + escapeHtml(currentValue) + '" />' +
            '<div class="channel-status">' + escapeHtml(row.statusText || "") + '</div>' +
          '</div>';
        }).join("");

        const messageHtml = payload?.message
          ? '<div class="message">' + escapeHtml(payload.message) + '</div>'
          : "";
        const historyHtml =
          payload?.historyText
            ? '<pre class="history">' + escapeHtml(payload.historyText) + '</pre>'
            : '<div class="empty">No monitor events have been recorded yet.</div>';
        const configValue =
          editingState?.type === "buffer"
            ? editingState.value
            : String(payload?.bufferSize || "");

        content.innerHTML = overlayMode
          ? '<div class="overlay-page">' +
              '<div><div class="overlay-title">Monitor Configuration</div><div class="overlay-hint">Buffer size controls how many monitor lines are kept in this widget.</div></div>' +
              '<div class="config-row"><label for="bufferSizeInput">Buffer size</label><input id="bufferSizeInput" class="config-input" type="text" spellcheck="false" value="' + escapeHtml(configValue) + '" /></div>' +
            '</div>'
          : messageHtml + historyHtml;

        restoreActiveInput();
        isRendering = false;

        if (!overlayMode) {
          if (wasPinnedToBottom) {
            content.scrollTop = content.scrollHeight;
          } else {
            content.scrollTop = previousScrollTop;
          }
        }
      }

      window.addEventListener("message", (event) => {
        if (event.data?.type === "monitorWidgetState") {
          render(event.data.state);
        }
      });

      render(initialState);
    </script>
  </body>
</html>`;
}

function buildPvlistWidgetSourceModel(
  options = {},
  extractRecordDeclarations,
  extractDatabaseTocMacroAssignments,
) {
  const sourceKind = options?.sourceKind === "pvlist" ? "pvlist" : "database";
  const sourceLabel = String(options?.sourceLabel || "EPICS PvList");
  const sourceText = String(options?.sourceText || "");
  const sourceDocumentUri = options?.sourceDocumentUri
    ? String(options.sourceDocumentUri)
    : "";

  if (sourceKind === "pvlist") {
    return {
      ...parsePvlistWidgetSourceText(sourceText, sourceLabel),
      sourceDocumentUri,
    };
  }

  const declarations =
    typeof extractRecordDeclarations === "function"
      ? extractRecordDeclarations(sourceText)
      : extractDatabaseMonitorDeclarationsFallback(sourceText);
  const rawPvNames = [];
  const seen = new Set();
  for (const declaration of declarations) {
    const name = String(declaration?.name || "").trim();
    if (!name || seen.has(name)) {
      continue;
    }
    seen.add(name);
    rawPvNames.push(name);
  }

  const { macroNames, macroValues } = buildDatabasePvlistMacroState(
    sourceText,
    extractDatabaseTocMacroAssignments,
  );

  return {
    sourceKind,
    sourceLabel,
    sourceDocumentUri,
    sourceText,
    rawPvNames,
    macroNames,
    macroValues,
    diagnostics: [],
  };
}

function buildDatabasePvlistMacroState(sourceText, extractDatabaseTocMacroAssignments) {
  const macroNames = [];
  const macroValues = new Map();
  const macroAssignments =
    typeof extractDatabaseTocMacroAssignments === "function"
      ? extractDatabaseTocMacroAssignments(sourceText)
      : new Map();

  if (macroAssignments instanceof Map) {
    for (const [macroName, assignment] of macroAssignments.entries()) {
      const normalizedMacroName = String(macroName || "").trim();
      if (!normalizedMacroName || macroValues.has(normalizedMacroName)) {
        continue;
      }
      macroNames.push(normalizedMacroName);
      macroValues.set(
        normalizedMacroName,
        assignment?.hasAssignment ? String(assignment?.value || "") : "",
      );
    }
  }

  for (const macroName of extractOrderedEpicsMacroNames([sourceText])) {
    if (!macroName || macroValues.has(macroName)) {
      continue;
    }
    macroNames.push(macroName);
    macroValues.set(macroName, "");
  }

  return { macroNames, macroValues };
}

function parsePvlistWidgetSourceText(text, sourceLabel) {
  const lines = String(text || "").split(/\r?\n/);
  const diagnostics = [];
  const rawPvNames = [];
  const macroNames = [];
  const macroValues = new Map();
  const seenMacroNames = new Set();
  const seenPvNames = new Set();

  for (let lineNumber = 0; lineNumber < lines.length; lineNumber += 1) {
    const lineText = lines[lineNumber];
    const trimmedLine = lineText.trim();
    const startCharacter = lineText.indexOf(trimmedLine);
    const endCharacter = startCharacter + trimmedLine.length;

    if (!trimmedLine || trimmedLine.startsWith("#")) {
      continue;
    }

    const macroMatch = trimmedLine.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (macroMatch) {
      const macroName = macroMatch[1];
      if (macroValues.has(macroName)) {
        diagnostics.push({
          lineNumber,
          startCharacter,
          endCharacter,
          message: `Duplicate pvlist macro "${macroName}".`,
        });
      } else {
        macroValues.set(macroName, macroMatch[2] || "");
      }
      if (!seenMacroNames.has(macroName)) {
        seenMacroNames.add(macroName);
        macroNames.push(macroName);
      }
      continue;
    }

    if (trimmedLine.includes("=")) {
      diagnostics.push({
        lineNumber,
        startCharacter,
        endCharacter,
        message: 'PV list macro definitions must be exactly "NAME = value" with no extra text.',
      });
      continue;
    }

    if (/\s/.test(trimmedLine)) {
      diagnostics.push({
        lineNumber,
        startCharacter,
        endCharacter,
        message: "PV list lines must contain exactly one record name with no extra text.",
      });
      continue;
    }

    if (!seenPvNames.has(trimmedLine)) {
      seenPvNames.add(trimmedLine);
      rawPvNames.push(trimmedLine);
    }
    for (const macroName of extractOrderedEpicsMacroNames([trimmedLine])) {
      if (!seenMacroNames.has(macroName)) {
        seenMacroNames.add(macroName);
        macroNames.push(macroName);
        macroValues.set(macroName, "");
      }
    }
  }

  return {
    sourceKind: "pvlist",
    sourceLabel,
    sourceText: String(text || ""),
    rawPvNames,
    macroNames,
    macroValues,
    diagnostics,
  };
}

function parseAddedPvlistChannelLines(text) {
  const entries = [];
  const seen = new Set();
  for (const rawLine of String(text || "").split(/\r?\n/)) {
    const line = String(rawLine || "").trim();
    if (!line || line.startsWith("#") || /\s/.test(line)) {
      continue;
    }
    if (seen.has(line)) {
      continue;
    }
    seen.add(line);
    entries.push(line);
  }
  return entries;
}

function extractOrderedEpicsMacroNames(texts) {
  const names = [];
  const seen = new Set();
  for (const text of Array.isArray(texts) ? texts : [texts]) {
    const sourceText = String(text || "");
    for (const match of sourceText.matchAll(/\$\(([^)=\s]+)(?:=[^)]*)?\)|\$\{([^}\s]+)\}/g)) {
      const macroName = match[1] || match[2];
      if (!macroName || seen.has(macroName)) {
        continue;
      }
      seen.add(macroName);
      names.push(macroName);
    }
  }
  return names;
}

function buildPvlistWidgetMonitorPlan(sourceModel, macroValues, defaultProtocol) {
  const rows = [];
  const definitions = [];
  const seen = new Set();
  const protocol = normalizeRuntimeProtocol(defaultProtocol);

  if (!sourceModel) {
    return { rows, definitions };
  }

  const sourceKind = sourceModel.sourceKind === "pvlist" ? "pvlist" : "database";
  const assignedMacroDefinitions = new Map();
  for (const macroName of sourceModel.macroNames || []) {
    const value = String(macroValues?.get(macroName) || "");
    if (!value) {
      continue;
    }
    assignedMacroDefinitions.set(macroName, {
      name: macroName,
      value,
    });
  }

  const strictMacroCache = new Map();
  const databaseAssignments = new Map(
    [...assignedMacroDefinitions.entries()].map(([macroName, definition]) => [
      macroName,
      { hasAssignment: true, value: definition.value },
    ]),
  );
  const databaseMacroDefinitions = createDatabaseMonitorMacroDefinitions(databaseAssignments);
  const databaseMacroCache = new Map();

  (sourceModel.rawPvNames || []).forEach((rawPvName, index) => {
    let pvName = "";
    let valueText = "";
    let unresolved = false;
    if (sourceKind === "pvlist") {
      const expansion = expandStrictMonitorValue(
        rawPvName,
        assignedMacroDefinitions,
        strictMacroCache,
        [],
      );
      if (expansion.errors.length > 0) {
        unresolved = true;
        valueText = `(${expansion.errors[0]})`;
      } else {
        pvName = String(expansion.value || "").trim();
        if (!pvName || /\s/.test(pvName) || hasUnresolvedEpicsMacroText(pvName)) {
          unresolved = true;
          valueText = "(set macros)";
        }
      }
    } else {
      pvName = normalizeDatabaseMonitorPvName(
        expandDatabaseMonitorValue(
          rawPvName,
          databaseMacroDefinitions,
          databaseMacroCache,
          [],
        ),
        rawPvName,
      );
      if (!pvName || /\s/.test(pvName) || hasUnresolvedEpicsMacroText(pvName)) {
        unresolved = true;
        valueText = "(set macros)";
      }
    }

    rows.push({
      id: `pvlist-row:${index}`,
      rawPvName,
      channelName: unresolved ? rawPvName : pvName,
      pvName: unresolved ? undefined : pvName,
      valueText,
    });

    if (unresolved || !pvName) {
      return;
    }

    const key = `${protocol}:${pvName}:`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    definitions.push({
      pvName,
      protocol,
      pvRequest: "",
    });
  });

  return { rows, definitions };
}

function buildPvlistWidgetFileText(sourceModel, macroValues, eol = "\n") {
  const macroNames = Array.isArray(sourceModel?.macroNames) ? sourceModel.macroNames : [];
  const rawPvNames = Array.isArray(sourceModel?.rawPvNames) ? sourceModel.rawPvNames : [];
  const lines = [
    "# this is a pvlist file for EPICS Workbench in VSCode",
    "",
  ];

  if (macroNames.length > 0) {
    for (const macroName of macroNames) {
      lines.push(`${macroName} = ${String(macroValues?.get(macroName) || "")}`);
    }
    lines.push("");
  }

  lines.push(...rawPvNames.map((entry) => String(entry || "").trim()).filter(Boolean));
  return lines.join(eol);
}

function getMonitorWidgetPanelTitle(channelNames) {
  const normalizedChannels = (Array.isArray(channelNames) ? channelNames : [])
    .map((channelName) => String(channelName || "").trim())
    .filter(Boolean);
  if (normalizedChannels.length === 0) {
    return "EPICS Monitor";
  }
  if (normalizedChannels.length === 1) {
    return `Monitor: ${normalizedChannels[0]}`;
  }
  return `Monitor: ${normalizedChannels[0]} (+${normalizedChannels.length - 1})`;
}

function trimMonitorWidgetHistory(widgetState) {
  const bufferSize = Math.max(
    1,
    Number(widgetState?.bufferSize) || DEFAULT_MONITOR_WIDGET_BUFFER_SIZE,
  );
  while ((widgetState?.historyLines?.length || 0) > bufferSize) {
    widgetState.historyLines.shift();
  }
}

function buildMonitorWidgetExportText(widgetState, eol = "\n") {
  const channelNames = (widgetState?.channelRows || [])
    .map((row) => String(row?.channelName || "").trim())
    .filter(Boolean);
  const lines = [
    channelNames.length === 1
      ? `# monitor data for channel ${channelNames[0]}`
      : channelNames.length > 1
        ? `# monitor data for channels ${channelNames.join(", ")}`
        : "# monitor data exported from EPICS Monitor widget",
    "",
    ...((widgetState?.historyLines || []).map((line) => String(line || ""))),
  ];
  return lines.join(eol);
}

function getDefaultMonitorWidgetSaveUri(widgetState) {
  const firstChannelName = (widgetState?.channelRows || [])
    .map((row) => String(row?.channelName || "").trim())
    .find(Boolean);
  const sanitizedBaseName = String(firstChannelName || "epics-monitor-data")
    .replace(/^ca:\/\//i, "")
    .replace(/^pva:\/\//i, "")
    .replace(/[^\w.-]+/g, "_")
    .replace(/^_+|_+$/g, "") || "epics-monitor-data";
  const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri;
  if (workspaceRoot?.scheme === "file") {
    return vscode.Uri.joinPath(workspaceRoot, `${sanitizedBaseName}.txt`);
  }
  return undefined;
}

function parseMonitorWidgetChannelReference(channelName, defaultProtocol) {
  let pvName = String(channelName || "").trim();
  let protocol = normalizeRuntimeProtocol(defaultProtocol);
  if (/^ca:\/\//i.test(pvName)) {
    protocol = "ca";
    pvName = pvName.replace(/^ca:\/\//i, "");
  } else if (/^pva:\/\//i.test(pvName)) {
    protocol = "pva";
    pvName = pvName.replace(/^pva:\/\//i, "");
  }

  return {
    protocol,
    pvName: pvName.trim(),
    pvRequest: "",
  };
}

function buildMonitorWidgetHistoryLine(session, monitor, runtimeLibrary) {
  if (!session || !monitor) {
    return "";
  }

  if (session.protocol === "pva") {
    const pvaData = monitor.getPvaData?.();
    updatePvaEnumChoicesCache(session, pvaData);
    const valueText = formatRuntimeDisplayValue(
      getPvaRuntimeDisplayValue(pvaData, session.pvaEnumChoices),
    );
    const timestampText = formatPvaMonitorTimestamp(pvaData);
    const alarmText = formatPvaMonitorAlarmText(pvaData);
    return [
      String(session.pvName || "").padEnd(28),
      timestampText,
      valueText,
      alarmText,
    ].filter(Boolean).join(" ").trimEnd();
  }

  const dbrData = monitor.getChannel?.().getDbrData?.();
  updateCaEnumChoicesCache(session, dbrData);
  const valueText = formatRuntimeDisplayValue(
    getCaRuntimeDisplayValue(session, dbrData),
  );
  const timestampText = formatCaMonitorTimestamp(dbrData);
  const alarmText = formatCaMonitorAlarmText(dbrData, runtimeLibrary);
  return [
    String(session.pvName || "").padEnd(28),
    timestampText,
    valueText,
    alarmText,
  ].filter(Boolean).join(" ").trimEnd();
}

function formatCaMonitorTimestamp(dbrData) {
  const secondsSinceEpoch = Number(dbrData?.secondsSinceEpoch);
  if (!Number.isFinite(secondsSinceEpoch)) {
    return "";
  }
  const nanoSeconds = Number(dbrData?.nanoSeconds || 0);
  return formatMonitorTimestamp(
    secondsSinceEpoch + EPICS_CA_EPOCH_OFFSET_SECONDS,
    nanoSeconds,
  );
}

function formatPvaMonitorTimestamp(pvaData) {
  const timeStamp = pvaData?.timeStamp;
  const secondsPastEpoch = timeStamp?.secondsPastEpoch;
  if (secondsPastEpoch === undefined || secondsPastEpoch === null) {
    return "";
  }
  const seconds = typeof secondsPastEpoch === "bigint"
    ? Number(secondsPastEpoch)
    : Number(secondsPastEpoch);
  if (!Number.isFinite(seconds)) {
    return "";
  }
  return formatMonitorTimestamp(seconds, Number(timeStamp?.nanoseconds || 0));
}

function formatMonitorTimestamp(epochSeconds, nanoseconds) {
  const seconds = Number(epochSeconds);
  if (!Number.isFinite(seconds)) {
    return "";
  }
  const normalizedNanoseconds = Math.max(0, Number(nanoseconds || 0));
  const date = new Date((seconds * 1000) + Math.floor(normalizedNanoseconds / 1e6));
  const microseconds = Math.floor(normalizedNanoseconds / 1000)
    .toString()
    .padStart(6, "0");
  return [
    date.getFullYear().toString().padStart(4, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-") +
    " " +
    [
      String(date.getHours()).padStart(2, "0"),
      String(date.getMinutes()).padStart(2, "0"),
      String(date.getSeconds()).padStart(2, "0"),
    ].join(":") +
    `.${microseconds}`;
}

function formatCaMonitorAlarmText(dbrData, runtimeLibrary) {
  const statusText = formatCaAlarmStatusText(
    Number(dbrData?.status),
    runtimeLibrary,
  );
  const severityText = formatCaAlarmSeverityText(
    Number(dbrData?.severity),
    runtimeLibrary,
  );
  return [statusText, severityText].filter(Boolean).join(" ");
}

function formatPvaMonitorAlarmText(pvaData) {
  const severityValue = Number(pvaData?.alarm?.severity);
  if (!Number.isFinite(severityValue)) {
    return "";
  }
  return formatCaAlarmSeverityText(severityValue);
}

function formatCaAlarmStatusText(value, runtimeLibrary) {
  const statusName = resolveEnumName(value, runtimeLibrary?.CA_ALARM_STATUS, {
    0: "NO_ALARM",
    1: "READ",
    2: "WRITE",
    3: "HIHI",
    4: "HIGH",
    5: "LOLO",
    6: "LOW",
    7: "STATE",
    8: "COS",
    9: "COMM",
    10: "TIMEOUT",
    11: "HWLIMIT",
    12: "CALC",
    13: "SCAN",
    14: "LINK",
    15: "SOFT",
    16: "BAD_SUB",
    17: "UDF",
    18: "DISABLE",
    19: "SIMM",
    20: "READ_ACCESS",
    21: "WRITE_ACCESS",
  });
  if (!statusName) {
    return "";
  }
  return statusName === "NO_ALARM" ? statusName : `${statusName}_ALARM`;
}

function formatCaAlarmSeverityText(value, runtimeLibrary) {
  const severityName = resolveEnumName(value, runtimeLibrary?.CA_ALARM_SEVRITY, {
    0: "NO_ALARM",
    1: "MINOR",
    2: "MAJOR",
    3: "INVALID",
  });
  if (!severityName) {
    return "";
  }
  return severityName === "NO_ALARM" ? severityName : `${severityName}_ALARM`;
}

function resolveEnumName(value, runtimeEnum, fallbackNames) {
  if (!Number.isFinite(value)) {
    return "";
  }
  if (runtimeEnum && Object.prototype.hasOwnProperty.call(runtimeEnum, value)) {
    return String(runtimeEnum[value] || "");
  }
  return String(fallbackNames?.[value] || "");
}

function safeParseUri(value) {
  try {
    return value ? vscode.Uri.parse(String(value)) : undefined;
  } catch (error) {
    return undefined;
  }
}

function getDefaultPvlistWidgetSaveUri(sourceModel) {
  const sourceUri = safeParseUri(sourceModel?.sourceDocumentUri);
  if (sourceUri?.scheme === "file") {
    const sourcePath = sourceUri.fsPath;
    if (sourceModel?.sourceKind === "pvlist") {
      return sourceUri;
    }
    return vscode.Uri.file(
      `${sourcePath.replace(/\.[^./\\]+$/, "")}.pvlist`,
    );
  }

  const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri;
  if (workspaceRoot?.scheme === "file") {
    const fallbackBaseName = String(sourceModel?.sourceLabel || "pvlist")
      .replace(/\.[^./\\]+$/, "")
      .replace(/[^\w.-]+/g, "_")
      .replace(/^_+|_+$/g, "") || "pvlist";
    return vscode.Uri.joinPath(workspaceRoot, `${fallbackBaseName}.pvlist`);
  }

  return undefined;
}

function hasUnresolvedEpicsMacroText(value) {
  return /\$\(([^)=\s]+)(?:=[^)]*)?\)|\$\{([^}\s]+)\}/.test(String(value || ""));
}

function buildPvlistWidgetHtml(webview, initialState = {}) {
  const nonce = createNonce();
  const initialStateJson = JSON.stringify(initialState).replace(/</g, "\\u003c");
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS PvList</title>
    <style>
      :root { color-scheme: light dark; }
      html, body {
        margin: 0;
        height: 100%;
        background: var(--vscode-editor-background);
        color: var(--vscode-foreground);
        font-family: var(--vscode-font-family);
      }
      body { display: flex; flex-direction: column; }
      .toolbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 12px 16px;
        border-bottom: 1px solid var(--vscode-panel-border);
        position: sticky;
        top: 0;
        background: var(--vscode-editor-background);
        z-index: 1;
      }
      .toolbar-left, .toolbar-right {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .toolbar-title { font-weight: 600; }
      .toolbar-status { color: var(--vscode-descriptionForeground); }
      .content {
        flex: 1;
        overflow: auto;
        padding: 20px 24px 28px;
      }
      h2 { margin: 0 0 12px; font-size: 1.1rem; }
      .section + .section { margin-top: 24px; }
      .message { margin-bottom: 16px; color: var(--vscode-descriptionForeground); }
      .bulk-editor {
        display: flex;
        flex-direction: column;
        gap: 10px;
        flex: 1;
        min-height: 0;
      }
      .bulk-input {
        width: 100%;
        min-height: 65vh;
        flex: 1;
        resize: vertical;
        padding: 10px 12px;
        border: 1px solid var(--vscode-input-border, transparent);
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        border-radius: 4px;
        font-family: inherit;
        font-size: inherit;
        line-height: 1.4;
        box-sizing: border-box;
      }
      .bulk-actions {
        display: flex;
        justify-content: flex-start;
      }
      .overlay-page {
        flex: 1;
        min-height: 0;
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .overlay-title {
        margin: 0;
        font-size: 1.2rem;
        font-weight: 600;
      }
      .overlay-hint {
        color: var(--vscode-descriptionForeground);
      }
      .action-button {
        padding: 6px 12px;
        border: 1px solid var(--vscode-button-border, transparent);
        border-radius: 4px;
        background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
        color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
        cursor: pointer;
        font: inherit;
      }
      .action-button:hover {
        background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground));
      }
      .macros {
        display: grid;
        grid-template-columns: max-content minmax(220px, 420px);
        gap: 10px 16px;
        align-items: center;
      }
      .macro-name {
        color: var(--vscode-descriptionForeground);
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
      }
      .macro-input, .inline-put-input {
        padding: 6px 10px;
        border: 1px solid var(--vscode-input-border, transparent);
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        border-radius: 4px;
        font: inherit;
        box-sizing: border-box;
      }
      .inline-put-input {
        width: 100%;
        margin: 0;
        min-height: calc(1em + 2px);
        padding: 0 6px;
        line-height: 1.4;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
      }
      th, td {
        text-align: left;
        padding: 8px 10px;
        border-bottom: 1px solid var(--vscode-panel-border);
        vertical-align: top;
      }
      th {
        color: var(--vscode-descriptionForeground);
        font-weight: 600;
      }
      td:first-child, th:first-child {
        width: 38ch;
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
      }
      .value-cell {
        cursor: pointer;
        position: relative;
      }
      .value-cell:hover { text-decoration: underline; }
      .value-cell.readonly { cursor: default; }
      .value-cell.readonly:hover { text-decoration: none; }
      .value-display {
        min-height: 1.4em;
        white-space: pre-wrap;
      }
      .value-display.hidden {
        visibility: hidden;
      }
      .value-edit-shell {
        position: absolute;
        inset: 8px 10px;
        display: flex;
        align-items: center;
        pointer-events: none;
      }
      .value-edit-shell .inline-put-input { pointer-events: auto; }
      .empty { color: var(--vscode-descriptionForeground); }
    </style>
  </head>
	  <body>
	    <div class="toolbar">
	      <div class="toolbar-left">
          <div class="toolbar-title">PvList</div>
          <div id="statusText" class="toolbar-status"></div>
        </div>
	      <div class="toolbar-right">
          <button id="addChannelsButton" class="action-button">Configure Channels</button>
          <button id="doneButton" class="action-button" hidden>Done</button>
	      </div>
	    </div>
	    <div id="content" class="content"></div>
	    <script nonce="${nonce}">
	      const vscode = acquireVsCodeApi();
      const doubleClickIntervalMs = ${JSON.stringify(MOUSE_DOUBLE_CLICK_INTERVAL_MS)};
	      const initialState = ${initialStateJson};
	      const content = document.getElementById("content");
	      const statusText = document.getElementById("statusText");
	      const addChannelsButton = document.getElementById("addChannelsButton");
	      const doneButton = document.getElementById("doneButton");
	      let currentState = initialState;
      let pendingClick = undefined;
      let editingState = undefined;
      let draftState = undefined;
      let isRendering = false;
      let suppressFocusSync = false;
      let overlayMode = false;

      function captureActiveEditingState() {
        const activeElement = document.activeElement;
        if (!activeElement || !(activeElement instanceof HTMLInputElement)) {
          return;
        }
        if (activeElement.classList.contains("macro-input")) {
          editingState = {
            type: "macro",
            name: activeElement.dataset.macroName,
            value: activeElement.value,
            selectionStart: activeElement.selectionStart,
            selectionEnd: activeElement.selectionEnd,
          };
          return;
        }
        if (activeElement.classList.contains("inline-put-input")) {
          editingState = {
            type: "value",
            key: activeElement.dataset.putKey,
            value: activeElement.value,
            selectionStart: activeElement.selectionStart,
            selectionEnd: activeElement.selectionEnd,
          };
        }
      }

      function captureActiveDraftState() {
        const activeElement = document.activeElement;
        if (!activeElement || !(activeElement instanceof HTMLTextAreaElement)) {
          return;
        }
        if (!activeElement.classList.contains("bulk-input")) {
          return;
        }
        draftState = {
          type: activeElement.dataset.draftType,
          value: activeElement.value,
          selectionStart: activeElement.selectionStart,
          selectionEnd: activeElement.selectionEnd,
          scrollTop: activeElement.scrollTop,
          scrollLeft: activeElement.scrollLeft,
        };
      }

      function updateActiveDraftSelection(input) {
        if (draftState?.type !== input.dataset.draftType) {
          return;
        }
        draftState = {
          ...draftState,
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
          scrollTop: input.scrollTop,
          scrollLeft: input.scrollLeft,
        };
      }

      function focusActiveDraftInput() {
        if (!draftState?.type) {
          return;
        }
        const draftInput = [...content.querySelectorAll(".bulk-input")].find(
          (candidate) => candidate.dataset.draftType === draftState.type,
        );
        if (!draftInput) {
          return;
        }
        const fallbackCaret = String(draftState.value || "").length;
        const selectionStart =
          typeof draftState.selectionStart === "number"
            ? draftState.selectionStart
            : fallbackCaret;
        const selectionEnd =
          typeof draftState.selectionEnd === "number"
            ? draftState.selectionEnd
            : fallbackCaret;
        const scrollTop =
          typeof draftState.scrollTop === "number"
            ? draftState.scrollTop
            : 0;
        const scrollLeft =
          typeof draftState.scrollLeft === "number"
            ? draftState.scrollLeft
            : 0;
        suppressFocusSync = true;
        draftInput.focus();
        suppressFocusSync = false;
        draftInput.setSelectionRange(selectionStart, selectionEnd);
        draftInput.scrollTop = scrollTop;
        draftInput.scrollLeft = scrollLeft;
      }

      addChannelsButton.addEventListener("click", () => {
        overlayMode = true;
        editingState = undefined;
        pendingClick = undefined;
        render(currentState);
      });

      doneButton.addEventListener("click", () => {
        overlayMode = false;
        render(currentState);
      });

	      function updateActiveMacroSelection(input) {
        if (editingState?.type !== "macro" || editingState.name !== input.dataset.macroName) {
          return;
        }
        editingState = {
          ...editingState,
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      }

      function focusActiveMacroInput() {
        if (editingState?.type !== "macro") {
          return;
        }
        const macroInput = [...content.querySelectorAll(".macro-input")].find(
          (candidate) => candidate.dataset.macroName === editingState.name,
        );
        if (!macroInput) {
          return;
        }
        const fallbackCaret = String(editingState.value || "").length;
        const selectionStart =
          typeof editingState.selectionStart === "number"
            ? editingState.selectionStart
            : fallbackCaret;
        const selectionEnd =
          typeof editingState.selectionEnd === "number"
            ? editingState.selectionEnd
            : fallbackCaret;
        suppressFocusSync = true;
        macroInput.focus();
        suppressFocusSync = false;
        macroInput.setSelectionRange(selectionStart, selectionEnd);
      }

      function updateActiveValueSelection(input) {
        if (editingState?.type !== "value" || editingState.key !== input.dataset.putKey) {
          return;
        }
        editingState = {
          ...editingState,
          value: input.value,
          selectionStart: input.selectionStart,
          selectionEnd: input.selectionEnd,
        };
      }

      function focusActiveValueInput() {
        if (editingState?.type !== "value") {
          return;
        }
        const valueInput = [...content.querySelectorAll(".inline-put-input")].find(
          (candidate) => candidate.dataset.putKey === editingState.key,
        );
        if (!valueInput) {
          return;
        }
        const fallbackCaret = String(editingState.value || "").length;
        const selectionStart =
          typeof editingState.selectionStart === "number"
            ? editingState.selectionStart
            : fallbackCaret;
        const selectionEnd =
          typeof editingState.selectionEnd === "number"
            ? editingState.selectionEnd
            : fallbackCaret;
        suppressFocusSync = true;
        valueInput.focus();
        suppressFocusSync = false;
        valueInput.setSelectionRange(selectionStart, selectionEnd);
      }

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function render(payload) {
        captureActiveEditingState();
        captureActiveDraftState();
        currentState = payload || {};
        statusText.textContent = payload?.contextStatus || "";
        addChannelsButton.hidden = overlayMode;
        doneButton.hidden = true;
        const messageHtml = payload?.message
          ? '<div class="message">' + escapeHtml(payload.message) + '</div>'
          : "";
        const channelDraftValue =
          draftState?.type === "channels"
            ? draftState.value
            : (Array.isArray(payload?.rawPvNames) ? payload.rawPvNames.join("\\n") : "");
        const macros = Array.isArray(payload?.macros) ? payload.macros : [];
        const rows = Array.isArray(payload?.rows) ? payload.rows : [];
        const macrosHtml = macros.length
          ? '<div class="macros">' + macros.map((macro) => {
              const currentValue =
                editingState?.type === "macro" && editingState.name === macro.name
                  ? editingState.value
                  : macro.value || "";
              return '<div class="macro-name">' + escapeHtml(macro.name) + '</div>' +
                '<input class="macro-input" data-macro-name="' + escapeHtml(macro.name) +
                '" type="text" spellcheck="false" value="' + escapeHtml(currentValue) + '" />';
            }).join("") + '</div>'
          : '<div class="empty">No macros are used by this source.</div>';
        const rowsHtml = rows.length
          ? '<table><thead><tr><th>Channel</th><th>Value</th></tr></thead><tbody>' +
            rows.map((row) => {
              const isEditingValue =
                editingState?.type === "value" && editingState.key === row.key;
              const valueHtml =
                '<div class="value-display ' + (isEditingValue ? 'hidden' : '') + '">' +
                  escapeHtml(row.value || "") +
                '</div>' +
                (isEditingValue
                  ? '<div class="value-edit-shell"><input class="inline-put-input" data-put-key="' +
                    escapeHtml(row.key || "") +
                    '" type="text" spellcheck="false" value="' + escapeHtml(editingState.value || "") + '" /></div>'
                  : '');
              const canPut = Boolean(row.canPut && row.key);
              return '<tr>' +
                '<td>' + escapeHtml(row.channelName || "") + '</td>' +
                '<td class="value-cell ' + (canPut ? "" : "readonly") + '" data-put-key="' +
                escapeHtml(row.key || "") + '">' + valueHtml + '</td>' +
                '</tr>';
            }).join("") + '</tbody></table>'
          : '<div class="empty">No channel rows are available.</div>';

        isRendering = true;
        content.className = overlayMode ? "content overlay-active" : "content";
        content.innerHTML = overlayMode
          ? '<div class="overlay-page">' +
              '<div><div class="overlay-title">Edit Channels</div><div class="overlay-hint">One channel per line. The full list here becomes the widget channel list, in this order.</div></div>' +
              '<div class="bulk-editor">' +
                '<textarea class="bulk-input" data-draft-type="channels" spellcheck="false" placeholder="One channel per line">' +
                  escapeHtml(channelDraftValue) +
                '</textarea>' +
                '<div class="bulk-actions"><button class="action-button" data-action="apply-channels">OK</button></div>' +
              '</div>' +
            '</div>'
          : messageHtml +
            '<div class="section"><h2>Macros</h2>' + macrosHtml + '</div>' +
            '<div class="section"><h2>Channels</h2>' + rowsHtml + '</div>';
        isRendering = false;

        for (const draftInput of content.querySelectorAll(".bulk-input")) {
          draftInput.addEventListener("focus", () => {
            if (suppressFocusSync) {
              return;
            }
            draftState = {
              type: draftInput.dataset.draftType,
              value: draftInput.value,
              selectionStart: draftInput.selectionStart,
              selectionEnd: draftInput.selectionEnd,
              scrollTop: draftInput.scrollTop,
              scrollLeft: draftInput.scrollLeft,
            };
          });
          draftInput.addEventListener("input", (event) => {
            draftState = {
              type: draftInput.dataset.draftType,
              value: event.target.value,
              selectionStart: event.target.selectionStart,
              selectionEnd: event.target.selectionEnd,
              scrollTop: event.target.scrollTop,
              scrollLeft: event.target.scrollLeft,
            };
          });
          draftInput.addEventListener("keyup", () => {
            updateActiveDraftSelection(draftInput);
          });
          draftInput.addEventListener("mouseup", () => {
            updateActiveDraftSelection(draftInput);
          });
          draftInput.addEventListener("select", () => {
            updateActiveDraftSelection(draftInput);
          });
          draftInput.addEventListener("scroll", () => {
            updateActiveDraftSelection(draftInput);
          });
          draftInput.addEventListener("blur", () => {
            if (isRendering) {
              return;
            }
            if (draftState?.type === draftInput.dataset.draftType) {
              draftState = undefined;
            }
          });
        }

        for (const actionButton of content.querySelectorAll(".action-button")) {
          actionButton.addEventListener("click", () => {
            const draftInput = content.querySelector('.bulk-input[data-draft-type="channels"]');
            const text = draftInput?.value || "";
            overlayMode = false;
            draftState = undefined;
            render(currentState);
            vscode.postMessage({
              type: "replacePvlistWidgetChannels",
              text,
            });
          });
        }

        for (const macroInput of content.querySelectorAll(".macro-input")) {
          macroInput.addEventListener("focus", () => {
            if (suppressFocusSync) {
              return;
            }
            editingState = {
              type: "macro",
              name: macroInput.dataset.macroName,
              value: macroInput.value,
              selectionStart: macroInput.selectionStart,
              selectionEnd: macroInput.selectionEnd,
            };
          });
          macroInput.addEventListener("input", (event) => {
            editingState = {
              type: "macro",
              name: macroInput.dataset.macroName,
              value: event.target.value,
              selectionStart: event.target.selectionStart,
              selectionEnd: event.target.selectionEnd,
            };
          });
          macroInput.addEventListener("keyup", () => {
            updateActiveMacroSelection(macroInput);
          });
          macroInput.addEventListener("mouseup", () => {
            updateActiveMacroSelection(macroInput);
          });
          macroInput.addEventListener("select", () => {
            updateActiveMacroSelection(macroInput);
          });
          macroInput.addEventListener("blur", () => {
            if (isRendering) {
              return;
            }
            if (editingState?.type === "macro" && editingState.name === macroInput.dataset.macroName) {
              editingState = undefined;
            }
          });
          macroInput.addEventListener("keydown", (event) => {
            if (event.key !== "Enter") {
              return;
            }
            const macroName = macroInput.dataset.macroName;
            editingState = undefined;
            vscode.postMessage({
              type: "updatePvlistWidgetMacro",
              name: macroName,
              value: macroInput.value,
            });
          });
        }

        for (const valueCell of content.querySelectorAll(".value-cell")) {
          valueCell.addEventListener("click", () => {
            const key = valueCell.dataset.putKey;
            if (!key || valueCell.classList.contains("readonly")) {
              pendingClick = undefined;
              return;
            }
            if (editingState?.type === "value") {
              if (editingState.key === key) {
                focusActiveValueInput();
              }
              pendingClick = undefined;
              return;
            }
            const now = Date.now();
            if (
              pendingClick &&
              pendingClick.key === key &&
              now - pendingClick.timestamp <= doubleClickIntervalMs
            ) {
              const row = rows.find((candidate) => candidate.key === key);
              editingState = {
                type: "value",
                key,
                value: row?.value || "",
                selectionStart: undefined,
                selectionEnd: undefined,
              };
              pendingClick = undefined;
              render(currentState);
              focusActiveValueInput();
              return;
            }
            pendingClick = { key, timestamp: now };
          });
        }

        for (const input of content.querySelectorAll(".inline-put-input")) {
          input.addEventListener("input", (event) => {
            editingState = {
              type: "value",
              key: input.dataset.putKey,
              value: event.target.value,
              selectionStart: event.target.selectionStart,
              selectionEnd: event.target.selectionEnd,
            };
          });
          input.addEventListener("keyup", () => {
            updateActiveValueSelection(input);
          });
          input.addEventListener("mouseup", () => {
            updateActiveValueSelection(input);
          });
          input.addEventListener("select", () => {
            updateActiveValueSelection(input);
          });
          input.addEventListener("blur", () => {
            if (editingState?.type !== "value" || editingState.key !== input.dataset.putKey) {
              return;
            }
            updateActiveValueSelection(input);
            requestAnimationFrame(() => {
              if (
                editingState?.type === "value" &&
                editingState.key === input.dataset.putKey &&
                document.activeElement !== input
              ) {
                focusActiveValueInput();
              }
            });
          });
          input.addEventListener("keydown", (event) => {
            updateActiveValueSelection(input);
            if (event.key === "Enter") {
              event.preventDefault();
              event.stopPropagation();
              const key = input.dataset.putKey;
              const value = input.value;
              editingState = undefined;
              pendingClick = undefined;
              input.blur();
              render(currentState);
              vscode.postMessage({
                type: "putPvlistValue",
                key,
                value,
              });
              return;
            }
            if (event.key === "Escape") {
              event.preventDefault();
              event.stopPropagation();
              editingState = undefined;
              pendingClick = undefined;
              render(currentState);
              return;
            }
          });
        }

        if (editingState?.type === "macro") {
          focusActiveMacroInput();
        }

        if (editingState?.type === "value") {
          focusActiveValueInput();
        } else if (draftState?.type) {
          focusActiveDraftInput();
        }
      }

      window.addEventListener("message", (event) => {
        if (event.data?.type === "pvlistWidgetState") {
          render(event.data.state);
        }
      });

      render(initialState);
    </script>
  </body>
</html>`;
}

function stripAnsiTerminalText(value) {
  return String(value || "").replace(
    /\u001b(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])/g,
    "",
  );
}

function quoteShellArgument(value) {
  return `'${String(value || "").replace(/'/g, `'\"'\"'`)}'`;
}

function createNonce() {
  return Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
}

class EpicsProbeCustomEditorProvider {
  constructor(controller) {
    this.controller = controller;
  }

  async resolveCustomTextEditor(document, webviewPanel) {
    await this.controller.resolveProbeCustomEditor(document, webviewPanel);
  }
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

  if (isProbeDocument(document)) {
    return analyzeProbeDocument(document);
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

function analyzeProbeDocument(document) {
  const text = document?.getText?.() || "";
  const lines = String(text).split(/\r?\n/);
  const diagnostics = [];
  const recordLines = [];

  for (let lineNumber = 0; lineNumber < lines.length; lineNumber += 1) {
    const lineText = lines[lineNumber];
    const trimmedLine = lineText.trim();
    const startCharacter = lineText.indexOf(trimmedLine);
    const endCharacter = startCharacter + trimmedLine.length;

    if (!trimmedLine) {
      continue;
    }

    if (trimmedLine.startsWith("#")) {
      continue;
    }

    recordLines.push({
      lineNumber,
      startCharacter,
      endCharacter,
      value: trimmedLine,
    });
  }

  if (recordLines.length > 1) {
    for (const recordLine of recordLines) {
      diagnostics.push({
        lineNumber: recordLine.lineNumber,
        startCharacter: recordLine.startCharacter,
        endCharacter: recordLine.endCharacter,
        message: "Probe files must contain exactly one non-empty record-name line.",
      });
    }
  }

  const recordLine = recordLines[0];
  if (recordLine) {
    if (/\$\(|\$\{/.test(recordLine.value)) {
      diagnostics.push({
        lineNumber: recordLine.lineNumber,
        startCharacter: recordLine.startCharacter,
        endCharacter: recordLine.endCharacter,
        message: "Probe record names cannot contain EPICS macros.",
      });
    }

    if (/\s/.test(recordLine.value)) {
      diagnostics.push({
        lineNumber: recordLine.lineNumber,
        startCharacter: recordLine.startCharacter,
        endCharacter: recordLine.endCharacter,
        message: "Probe files allow only one record name with no extra text.",
      });
    }
  }

  return {
    recordName:
      diagnostics.length === 0 && recordLine
        ? recordLine.value
        : undefined,
    recordLineNumber: recordLine?.lineNumber,
    recordType: undefined,
    definitions: [],
    diagnostics,
    lineReferences: [],
  };
}

function buildProbeOverlayAnchorPosition(document, analysis) {
  const recordLineNumber = Number(analysis?.recordLineNumber);
  if (!Number.isInteger(recordLineNumber) || recordLineNumber < 0) {
    return undefined;
  }

  const line = document.lineAt(recordLineNumber);
  return new vscode.Position(recordLineNumber, line.text.length);
}

function buildProbeOverlayLines(state) {
  if (!state) {
    return [];
  }

  const lines = [
    `Probe`,
    `Value: ${state.value}`,
    `Type: ${state.recordType}`,
    `Updated: ${state.lastUpdated}`,
    `Access: ${state.access}`,
  ];

  if (state.fields.length) {
    const fieldLines = state.fields
      .slice(0, Math.max(0, PROBE_OVERLAY_MAX_LINES - 6))
      .map((field) => {
      const valueText = truncateText(String(field.value || ""), 28);
      return `${field.fieldName}: ${valueText}`;
      });
    lines.push("Fields:");
    lines.push(...fieldLines);
    if (state.fields.length > fieldLines.length) {
      lines.push(`+${state.fields.length - fieldLines.length} more fields`);
    }
  } else if (state.fieldStatusText) {
    lines.push(state.fieldStatusText);
  }

  return lines;
}

function buildProbeOverlayHoverMarkdown(state) {
  if (!state) {
    return "Probe session is not available.";
  }

  const lines = [
    `**${escapeMarkdownText(state.recordName || "EPICS Probe")}**`,
    "",
    `Value: \`${escapeInlineCode(state.value)}\``,
    `Record Type: \`${escapeInlineCode(state.recordType)}\``,
    `Last Update: ${escapeMarkdownText(state.lastUpdated)}`,
    `Permission: ${escapeMarkdownText(state.access)}`,
  ];

  if (state.fields.length) {
    lines.push("", "**Fields**");
    for (const field of state.fields) {
      const updatedSuffix = field.updated
        ? ` (${escapeMarkdownText(field.updated)})`
        : "";
      lines.push(
        `- \`${escapeInlineCode(field.fieldName)}\`: \`${escapeInlineCode(field.value)}\`${updatedSuffix}`,
      );
    }
  } else if (state.fieldStatusText) {
    lines.push("", escapeMarkdownText(state.fieldStatusText));
  }

  return lines.join("\n");
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

    const macroMatch = trimmedLine.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (macroMatch) {
      const macroName = macroMatch[1];
      if (macroDefinitions.has(macroName)) {
        diagnostics.push({
          lineNumber,
          startCharacter,
          endCharacter,
          message: `Duplicate pvlist macro "${macroName}".`,
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
        message: 'PV list macro definitions must be exactly "NAME = value" with no extra text.',
      });
      parsedLines.push({ type: "invalid" });
      continue;
    }

    if (/\s/.test(trimmedLine)) {
      diagnostics.push({
        lineNumber,
        startCharacter,
        endCharacter,
        message: "PV list lines must contain exactly one record name with no extra text.",
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
        message: "PV list entry resolves to invalid text.",
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
        `Circular pvlist macro reference: ${[...stack, macroName].join(" -> ")}.`,
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
      errors: [`Undefined pvlist macro "${macroName}".`],
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

function shouldIgnoreTransientProbeFieldValue(entry, value) {
  if (entry?.sourceKind !== "probe" || entry?.probeRole !== "field") {
    return false;
  }

  if (value === undefined || value === null) {
    return true;
  }

  if (Array.isArray(value) && value.length === 0) {
    return true;
  }

  return false;
}

function getProbeEntryDisplayValue(entry) {
  if (!entry) {
    return "(connecting...)";
  }

  if (
    entry.status === "connecting" ||
    entry.status === "pending" ||
    entry.status === "disconnected" ||
    entry.status === "stopped"
  ) {
    return "(connecting...)";
  }

  if (entry.status === "error" || entry.status === "destroyed") {
    return "(connecting...)";
  }

  if (entry.valueText) {
    return entry.valueText;
  }

  const value = getMonitorHoverValue(entry);
  if (value === undefined) {
    return entry.valueText || "(connecting...)";
  }

  return formatRuntimeDisplayValue(value);
}

function getProbeAccessLabel(entry, runtimeLibrary) {
  if (!entry) {
    return "(connecting...)";
  }

  if (
    entry.status === "connecting" ||
    entry.status === "pending" ||
    entry.status === "disconnected" ||
    entry.status === "stopped"
  ) {
    return "(connecting...)";
  }

  if (entry.protocol === "pva") {
    return canPutProbeEntry(entry) ? "Read/Write" : "Read only";
  }

  if (!runtimeLibrary || !entry.channel?.getAccessRight) {
    return canPutProbeEntry(entry) ? "Read/Write" : "Read only";
  }

  const accessRight = entry.channel.getAccessRight?.();
  if (accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.READ_ONLY) {
    return "Read only";
  }
  if (
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.WRITE_ONLY ||
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.READ_WRITE
  ) {
    return "Read/Write";
  }
  if (
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NOT_AVAILABLE ||
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NO_ACCESS
  ) {
    return "No access";
  }

  return canPutProbeEntry(entry) ? "Read/Write" : "Read only";
}

function canPutProbeEntry(entry) {
  if (!entry) {
    return false;
  }

  if (entry.probeRole === "recordType") {
    return false;
  }

  if (
    entry.status === "connecting" ||
    entry.status === "pending" ||
    entry.status === "disconnected" ||
    entry.status === "stopped" ||
    entry.status === "error" ||
    entry.status === "destroyed"
  ) {
    return false;
  }

  if (entry.protocol === "pva") {
    const currentValue = getMonitorHoverValue(entry);
    if (currentValue === undefined || Array.isArray(currentValue)) {
      return false;
    }
    if (isPvaEnumLikeValue(currentValue)) {
      return true;
    }
    return typeof currentValue === "string" || typeof currentValue === "number";
  }

  const valueCount = Number(entry.channel?.getValueCount?.() || 0);
  if (valueCount > 1) {
    return false;
  }

  const runtimeLibrary = safeRequireRuntimeLibrary();
  if (!runtimeLibrary || !entry.channel?.getAccessRight) {
    return false;
  }

  const accessRight = entry.channel.getAccessRight?.();
  return !(
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NOT_AVAILABLE ||
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.NO_ACCESS ||
    accessRight === runtimeLibrary.Channel_ACCESS_RIGHTS.READ_ONLY
  );
}

function safeRequireRuntimeLibrary() {
  try {
    return require("epics-tca");
  } catch (error) {
    return undefined;
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
  getDefaultProjectRuntimeConfiguration,
  loadProjectRuntimeConfiguration,
  createRuntimeEnvironmentFromProjectConfiguration,
  normalizeRuntimeProtocol,
  resolveStartupExecutableValidation,
  safeRequireRuntimeLibrary,
  formatRuntimeValue,
  getCaRuntimeDisplayValue,
  getPvaRuntimeDisplayValue,
};
