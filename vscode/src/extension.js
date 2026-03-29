const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");
const childProcess = require("child_process");
const vscode = require("vscode");
const { collectEpicsBuildApplication } = require("../scripts/epics-build-model");
const { createDatabaseExcelTools } = require("./databaseExcel");
const { createDatabaseExcelImportTools } = require("./databaseExcelImport");
const {
  formatDatabaseText,
  formatMakefileText,
  formatProtocolText,
  formatStartupText,
  formatSubstitutionText,
  splitSubstitutionCommaSeparatedItems,
} = require("./formatters");
const { createDatabaseTocTools } = require("./databaseToc");
const {
  formatSequencerText,
  getSequencerDefinitionLocation,
  getSequencerReferenceLocations,
  getSequencerSymbolHover,
} = require("./sequencer");
const {
  registerRuntimeMonitor,
  getDefaultProjectRuntimeConfiguration,
  loadProjectRuntimeConfiguration,
  createRuntimeEnvironmentFromProjectConfiguration,
  normalizeRuntimeProtocol,
  safeRequireRuntimeLibrary,
  formatRuntimeValue,
  getCaRuntimeDisplayValue,
  getPvaRuntimeDisplayValue,
} = require("./runtimeMonitor");
const { registerTdmIntegration } = require("./tdmIntegration");

const LANGUAGE_IDS = {
  database: "database",
  startup: "startup",
  substitutions: "substitutions",
  dbd: "database definition",
  source: "epics-source",
  proto: "proto",
  sequencer: "sequencer",
  pvlist: "pvlist",
  probe: "probe",
};

const DATABASE_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const SUBSTITUTION_EXTENSIONS = new Set([".sub", ".subs", ".substitutions"]);
const STARTUP_EXTENSIONS = new Set([".cmd", ".iocsh"]);
const DBD_EXTENSIONS = new Set([".dbd"]);
const PVLIST_EXTENSIONS = new Set([".pvlist"]);
const PROBE_EXTENSIONS = new Set([".probe"]);
const PROTOCOL_EXTENSIONS = new Set([".proto"]);
const SOURCE_EXTENSIONS = new Set([
  ".c",
  ".cc",
  ".cpp",
  ".cxx",
  ".h",
  ".hh",
  ".hpp",
  ".hxx",
]);
const EPICS_SOURCE_DOCUMENT_SELECTORS = [...SOURCE_EXTENSIONS].map((extension) => ({
  scheme: "file",
  pattern: `**/*${extension}`,
}));
const INDEX_GLOB =
  "**/*.{db,vdb,template,sub,subs,substitutions,cmd,iocsh,dbd,pvlist,probe,proto}";
const SOURCE_INDEX_GLOB = "**/*.{c,cc,cpp,cxx,h,hh,hpp,hxx}";
const PROJECT_INDEX_GLOBS = [
  "**/Makefile",
  "**/configure/RELEASE",
  "**/configure/RELEASE.local",
];
const EPICS_PROJECT_MARKER_SEGMENTS = [
  ["Makefile"],
  ["configure", "RELEASE"],
  ["configure", "RULES_TOP"],
];
const INDEX_EXCLUDE_GLOB = "**/{.git,node_modules,out,dist}/**";
const OPEN_RECORD_LOCATION_COMMAND = "vscode-epics.openRecordLocation";
const INSERT_RECORD_TAIL_COMMAND = "vscode-epics.insertRecordTail";
const INSERT_FIELD_TAIL_COMMAND = "vscode-epics.insertFieldTail";
const INSERT_STARTUP_COMMAND_TAIL_COMMAND = "vscode-epics.insertStartupCommandTail";
const INSERT_DBD_DEVICE_TAIL_COMMAND = "vscode-epics.insertDbdDeviceTail";
const INSERT_DBD_DRIVER_TAIL_COMMAND = "vscode-epics.insertDbdDriverTail";
const INSERT_DBD_REGISTRAR_TAIL_COMMAND = "vscode-epics.insertDbdRegistrarTail";
const INSERT_DBD_FUNCTION_TAIL_COMMAND = "vscode-epics.insertDbdFunctionTail";
const INSERT_DBD_VARIABLE_TAIL_COMMAND = "vscode-epics.insertDbdVariableTail";
const INSERT_DBLOAD_RECORDS_MACRO_TAIL_COMMAND = "vscode-epics.insertDbLoadRecordsMacroTail";
const COLLAPSE_ALL_RECORDS_COMMAND = "vscode-epics.collapseAllRecords";
const EXPAND_ALL_RECORDS_COMMAND = "vscode-epics.expandAllRecords";
const GENERATE_DATABASE_TOC_COMMAND = "vscode-epics.generateDatabaseToc";
const FORMAT_DATABASE_FILE_COMMAND = "vscode-epics.formatDatabaseFile";
const FORMAT_ACTIVE_EPICS_FILE_COMMAND = "vscode-epics.formatActiveEpicsFile";
const COPY_ALL_RECORD_NAMES_COMMAND = "vscode-epics.copyAllRecordNames";
const COPY_AS_MONITOR_FILE_COMMAND = "vscode-epics.copyAsMonitorFile";
const COPY_AS_EXPANDED_DB_COMMAND = "vscode-epics.copyAsExpandedDb";
const ADD_DB_TO_MAKEFILE_COMMAND = "vscode-epics.addDbToMakefile";
const BUILD_WITH_MAKEFILE_COMMAND = "vscode-epics.buildWithMakefile";
const CLEAN_AND_BUILD_WITH_MAKEFILE_COMMAND =
  "vscode-epics.cleanAndBuildWithMakefile";
const BUILD_PROJECT_COMMAND = "vscode-epics.buildProject";
const CLEAN_AND_BUILD_PROJECT_COMMAND = "vscode-epics.cleanAndBuildProject";
const EXPORT_DATABASE_TO_EXCEL_COMMAND = "vscode-epics.exportDatabaseToExcel";
const IMPORT_DATABASE_FROM_EXCEL_COMMAND = "vscode-epics.importDatabaseFromExcel";
const IMPORT_DATABASE_FROM_EXCEL_SHORT_COMMAND =
  "vscode-epics.importDatabaseFromExcelShort";
const START_DATABASE_MONITOR_CHANNELS_COMMAND =
  "vscode-epics.startDatabaseMonitorChannels";
const STOP_DATABASE_MONITOR_CHANNELS_COMMAND =
  "vscode-epics.stopDatabaseMonitorChannels";
const EPICS_PROJECT_CONTEXT_KEY = "epicsWorkbench.isEpicsProject";
const ACTIVE_CAN_ADD_DB_TO_MAKEFILE_CONTEXT_KEY =
  "epicsWorkbench.activeCanAddDbToMakefile";
const ACTIVE_CAN_BUILD_WITH_MAKEFILE_CONTEXT_KEY =
  "epicsWorkbench.activeCanBuildWithMakefile";
const OPEN_EXCEL_IMPORT_PREVIEW_COMMAND = "vscode-epics.openExcelImportPreview";
const OPEN_IN_PROBE_COMMAND = "vscode-epics.openInProbe";
const OPEN_IN_PVLIST_COMMAND = "vscode-epics.openInPvList";
const OPEN_IN_MONITOR_COMMAND = "vscode-epics.openInMonitor";
const OPEN_IN_CHANNEL_GRAPH_COMMAND = "vscode-epics.openInChannelGraph";
const OPEN_PROBE_WIDGET_COMMAND = "vscode-epics.openProbeWidget";
const OPEN_PVLIST_WIDGET_COMMAND = "vscode-epics.openPvlistWidget";
const OPEN_MONITOR_WIDGET_COMMAND = "vscode-epics.openMonitorWidget";
const CHANNEL_GRAPH_VIEW_TYPE = "epicsWorkbench.channelGraph";
const UPDATE_MENU_FIELD_VALUE_COMMAND = "vscode-epics.updateMenuFieldValue";
const EXCEL_IMPORT_PREVIEW_VIEW_TYPE = "vscode-epics.excelImportPreview";
const STREAM_PROTOCOL_PATH_VARIABLE = "STREAM_PROTOCOL_PATH";
const PROJECT_RUNTIME_CONFIG_FILE_NAME = ".epics-workbench-config.json";
const DBD_DEVICE_LINK_TYPES = ["INST_IO"];
const TRIGGER_SUGGEST_COMMAND = "editor.action.triggerSuggest";
const RECORD_PREVIEW_MAX_LINES = 100;
const RECORD_PREVIEW_MAX_CHARACTERS = 12000;
const MAKEFILE_INSTALLABLE_DATABASE_EXTENSIONS = new Set([".db", ".template"]);
const STARTUP_COMMAND_TEMPLATES = new Map([
  ["dbLoadRecords", '("${1}")'],
  ["dbLoadTemplate", '("${1}")'],
]);
const DATABASE_SEMANTIC_TOKEN_TYPES = [
  "type",
  "variable",
  "property",
  "number",
  "enumMember",
  "string",
  "macro",
];
const DATABASE_SEMANTIC_TOKEN_MODIFIERS = ["declaration"];

const COMPLETION_TRIGGER_CHARACTERS = ["(", "\"", "$", "{", "/", ".", ":", ","];
const NUMERIC_DBF_TYPES = new Set([
  "DBF_SHORT",
  "DBF_ENUM",
  "DBF_UCHAR",
  "DBF_UINT64",
  "DBF_ULONG",
  "DBF_USHORT",
  "DBF_DOUBLE",
  "DBF_INT64",
  "DBF_LONG",
]);
const INTEGER_DBF_TYPES = new Set([
  "DBF_SHORT",
  "DBF_ENUM",
  "DBF_UCHAR",
  "DBF_UINT64",
  "DBF_ULONG",
  "DBF_USHORT",
  "DBF_INT64",
  "DBF_LONG",
]);
const LINK_DBF_TYPES = new Set([
  "DBF_FWDLINK",
  "DBF_INLINK",
  "DBF_OUTLINK",
]);
const EMPTY_DEFAULT_DBF_TYPES = new Set([
  "DBF_STRING",
  "DBF_FWDLINK",
  "DBF_INLINK",
  "DBF_OUTLINK",
]);
const DBD_DEVICE_DECLARATION_CACHE = new Map();

const STATIC_FIELD_VALUE_ENUMS = {
  SCAN: [
    "Passive",
    "Event",
    "I/O Intr",
    ".1 second",
    ".2 second",
    ".5 second",
    "1 second",
    "2 second",
    "5 second",
    "10 second",
    "15 second",
    "30 second",
    "1 minute",
    "5 minutes",
    "10 minutes",
    "30 minutes",
    "1 hour",
  ],
  OMSL: ["supervisory", "Supervisory", "closed_loop"],
  SELM: ["All", "Specified", "Mask"],
  OOPT: [
    "Every Time",
    "On Change",
    "When Zero",
    "When Non-zero",
    "Transition To Zero",
    "Transition To Non-zero",
  ],
  DOPT: ["Use CALC", "Use OCAL"],
  OIF: ["Full", "Incremental"],
  DTYP: ["Soft Channel", "stream"],
  PINI: ["YES", "NO"],
  HHSV: ["NO_ALARM", "MINOR", "MAJOR"],
  HSV: ["NO_ALARM", "MINOR", "MAJOR"],
  LSV: ["NO_ALARM", "MINOR", "MAJOR"],
  LLSV: ["NO_ALARM", "MINOR", "MAJOR"],
  ZSV: ["NO_ALARM", "MINOR", "MAJOR"],
  OSV: ["NO_ALARM", "MINOR", "MAJOR"],
  COSV: ["NO_ALARM", "MINOR", "MAJOR"],
  UNSV: ["NO_ALARM", "MINOR", "MAJOR"],
};

const COMMON_RECORD_FIELDS = [
  "NAME",
  "DESC",
  "ASG",
  "SCAN",
  "PINI",
  "PHAS",
  "DTYP",
  "DISA",
  "SDIS",
  "DISP",
  "PROC",
  "STAT",
  "SEVR",
  "UDF",
  "FLNK",
  "VAL",
  "PREC",
  "EGU",
  "INP",
  "OUT",
  "TSEL",
  "DOL",
  "OMSL",
  "SIML",
  "SIMM",
  "SIMS",
  "SIOL",
  "HOPR",
  "LOPR",
];
let pendingRecordTypeSuggest;
let pendingRecordTypeSuggestTimer;
let pendingRecordNameSuggest;
let pendingRecordNameSuggestTimer;
let pendingFieldNameSuggest;
let pendingFieldNameSuggestTimer;
let recordTypeSuggestAnchor;
let fieldNameSuggestAnchor;
let recordTemplateFields = new Map();
let recordTemplateStaticData = {
  fieldOrderByRecordType: new Map(),
  fieldTypesByRecordType: new Map(),
  fieldMenuChoicesByRecordType: new Map(),
  fieldInitialValuesByRecordType: new Map(),
};
const {
  buildOpenRecordCommandUri,
  extractDatabaseTocEntries,
  extractDatabaseTocMacroAssignments,
  findRecordDeclarationByTypeAndName,
  removeDatabaseTocBlock,
  upsertDatabaseTocText,
} = createDatabaseTocTools({
  openRecordLocationCommand: OPEN_RECORD_LOCATION_COMMAND,
  getRecordTemplateStaticData: () => recordTemplateStaticData,
  getDefaultFieldValue,
  extractRecordDeclarations,
  extractFieldDeclarationsInRecord,
  maskDatabaseComments,
  extractMacroNames,
  compareLabels,
  escapeRegExp,
});
const { buildDatabaseWorkbookBuffer } = createDatabaseExcelTools({
  extractRecordDeclarations,
  extractFieldDeclarationsInRecord,
});
const { importDatabaseWorkbookBuffer } = createDatabaseExcelImportTools();

function activate(context) {
  const runtimeMonitorController = registerRuntimeMonitor(context, {
    extractDatabaseTocEntries,
    extractDatabaseTocMacroAssignments,
    extractRecordDeclarations,
    getFieldNamesForRecordType: getRuntimeProbeFieldNamesForRecordType,
    getFieldTypeForRecordType: getRuntimeProbeFieldTypeForRecordType,
  });
  registerTdmIntegration(context);
  const staticData = loadStaticData(context.extensionPath);
  recordTemplateFields = new Map(staticData.recordTemplateFields || []);
  recordTemplateStaticData = {
    fieldOrderByRecordType: staticData.fieldOrderByRecordType,
    fieldTypesByRecordType: staticData.fieldTypesByRecordType,
    fieldMenuChoicesByRecordType: staticData.fieldMenuChoicesByRecordType,
    fieldInitialValuesByRecordType: staticData.fieldInitialValuesByRecordType,
  };
  const makeBuildOutputChannel = vscode.window.createOutputChannel("EPICS Build");
  context.subscriptions.push(makeBuildOutputChannel);
  const refreshActiveMakefileContextKeys = () => {
    updateActiveMakefileContextKeys(vscode.window.activeTextEditor);
  };
  void updateEpicsProjectContext();
  refreshActiveMakefileContextKeys();
  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      void updateEpicsProjectContext();
      refreshActiveMakefileContextKeys();
    }),
  );
  registerWatcher(context.subscriptions, "**/Makefile", () => {
    void updateEpicsProjectContext();
    refreshActiveMakefileContextKeys();
  });
  registerWatcher(context.subscriptions, "**/configure/RELEASE", () => {
    void updateEpicsProjectContext();
  });
  registerWatcher(context.subscriptions, "**/configure/RULES_TOP", () => {
    void updateEpicsProjectContext();
  });
  const workspaceIndex = new WorkspaceIndex(staticData);
  const diagnostics = vscode.languages.createDiagnosticCollection("vscode-epics");
  const databaseSemanticTokensLegend = new vscode.SemanticTokensLegend(
    DATABASE_SEMANTIC_TOKEN_TYPES,
    DATABASE_SEMANTIC_TOKEN_MODIFIERS,
  );
  const databaseRecordDecorationTypes = createDatabaseRecordDecorationTypes();
  const databaseValueDecorationTypes = createDatabaseValueDecorationTypes();
  const refreshDocumentDiagnostics = (document) =>
    void refreshDiagnostics(document, diagnostics, workspaceIndex);
  const refreshDocumentDecorations = (document) =>
    updateDatabaseRecordDecorationsForDocument(
      document,
      databaseRecordDecorationTypes,
    );
  const refreshDatabaseValueDecorations = (document) =>
    void updateDatabaseValueDecorationsForDocument(
      document,
      workspaceIndex,
      databaseValueDecorationTypes,
    );
  const refreshOpenDiagnostics = () => {
    for (const document of vscode.workspace.textDocuments) {
      refreshDocumentDiagnostics(document);
    }
  };
  const refreshDependentOpenDiagnostics = () => {
    for (const document of vscode.workspace.textDocuments) {
      if (isStartupDocument(document) || isSubstitutionsDocument(document)) {
        refreshDocumentDiagnostics(document);
      }
    }
  };
  context.subscriptions.push(workspaceIndex);
  context.subscriptions.push(
    workspaceIndex.onDidChange(() => {
      refreshOpenDiagnostics();
    }),
  );
  context.subscriptions.push(diagnostics);
  context.subscriptions.push(...databaseRecordDecorationTypes);
  context.subscriptions.push(...Object.values(databaseValueDecorationTypes));
  context.subscriptions.push(
    vscode.commands.registerCommand(
      OPEN_RECORD_LOCATION_COMMAND,
      async ({ absolutePath, line }) => {
        if (!absolutePath || !fs.existsSync(absolutePath)) {
          return;
        }

        const document = await vscode.workspace.openTextDocument(
          vscode.Uri.file(absolutePath),
        );
        const lineIndex = Math.max(0, Number(line || 1) - 1);
        const selection = new vscode.Range(lineIndex, 0, lineIndex, 0);
        await vscode.window.showTextDocument(document, {
          preview: false,
          selection,
        });
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_RECORD_TAIL_COMMAND,
      async ({ recordType }) => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || !isDatabaseDocument(editor.document) || !recordType) {
          return;
        }

        const snippet = buildRecordTemplateTailSnippet(
          recordType,
          recordTemplateFields.get(recordType),
        );
        if (!snippet) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString(snippet),
          getRecordTailInsertionRange(editor.document, editor.selection.active),
        );
        schedulePendingRecordNameSuggest(editor.document.uri.toString());
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_FIELD_TAIL_COMMAND,
      async ({ recordType, fieldName }) => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || !isDatabaseDocument(editor.document) || !fieldName) {
          return;
        }

        const snippet = buildFieldCompletionTailSnippet(
          recordTemplateStaticData,
          recordType,
          fieldName,
        );
        await editor.insertSnippet(
          new vscode.SnippetString(snippet),
          getFieldTailInsertionRange(editor.document, editor.selection.active),
        );
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_STARTUP_COMMAND_TAIL_COMMAND,
      async ({ commandName }) => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || !isStartupDocument(editor.document) || !commandName) {
          return;
        }

        const snippet = STARTUP_COMMAND_TEMPLATES.get(commandName);
        if (!snippet) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString(snippet),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand(TRIGGER_SUGGEST_COMMAND);
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBLOAD_RECORDS_MACRO_TAIL_COMMAND,
      async ({ absolutePath }) => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || !isStartupDocument(editor.document)) {
          return;
        }

        const fileText = absolutePath ? readTextFile(absolutePath) : undefined;
        const macroNames = fileText
          ? extractMacroNames(maskDatabaseComments(fileText))
          : [];
        const snippet = buildDbLoadRecordsCompletionTailSnippet(macroNames);
        await editor.insertSnippet(
          new vscode.SnippetString(snippet),
          getDbLoadRecordsTailInsertionRange(
            editor.document,
            editor.selection.active,
          ),
        );
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBD_DEVICE_TAIL_COMMAND,
      async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.languageId !== LANGUAGE_IDS.dbd) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString("(${0})"),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.triggerSuggest");
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBD_DRIVER_TAIL_COMMAND,
      async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.languageId !== LANGUAGE_IDS.dbd) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString("(${0})"),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.triggerSuggest");
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBD_REGISTRAR_TAIL_COMMAND,
      async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.languageId !== LANGUAGE_IDS.dbd) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString("(${0})"),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.triggerSuggest");
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBD_FUNCTION_TAIL_COMMAND,
      async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.languageId !== LANGUAGE_IDS.dbd) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString("(${0})"),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.triggerSuggest");
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      INSERT_DBD_VARIABLE_TAIL_COMMAND,
      async () => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.languageId !== LANGUAGE_IDS.dbd) {
          return;
        }

        await editor.insertSnippet(
          new vscode.SnippetString("(${0})"),
          new vscode.Range(
            editor.selection.active.line,
            editor.selection.active.character,
            editor.selection.active.line,
            editor.selection.active.character,
          ),
        );
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.triggerSuggest");
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(COLLAPSE_ALL_RECORDS_COMMAND, async () => {
      await foldAllRecordsInActiveEditor("editor.fold");
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(EXPAND_ALL_RECORDS_COMMAND, async () => {
      await foldAllRecordsInActiveEditor("editor.unfold");
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(GENERATE_DATABASE_TOC_COMMAND, async () => {
      await generateDatabaseTocInActiveEditor();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(FORMAT_DATABASE_FILE_COMMAND, async () => {
      await formatDatabaseFileInActiveEditor();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(FORMAT_ACTIVE_EPICS_FILE_COMMAND, async () => {
      await formatActiveEpicsFileInActiveEditor();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(COPY_ALL_RECORD_NAMES_COMMAND, async () => {
      await copyAllRecordNamesInActiveEditor();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(COPY_AS_MONITOR_FILE_COMMAND, async () => {
      await copyDatabaseAsMonitorFileInActiveEditor();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(COPY_AS_EXPANDED_DB_COMMAND, async () => {
      await copySubstitutionsAsExpandedDatabaseInActiveEditor(workspaceIndex);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(ADD_DB_TO_MAKEFILE_COMMAND, async (resourceUri) => {
      await addDbToMakefileForDocument(resourceUri);
      refreshActiveMakefileContextKeys();
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(BUILD_WITH_MAKEFILE_COMMAND, async (resourceUri) => {
      await buildWithLocalMakefile(resourceUri, makeBuildOutputChannel);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(CLEAN_AND_BUILD_WITH_MAKEFILE_COMMAND, async (resourceUri) => {
      await cleanWithLocalMakefile(resourceUri, makeBuildOutputChannel);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(BUILD_PROJECT_COMMAND, async (resourceUri) => {
      await buildEpicsProject(resourceUri, makeBuildOutputChannel);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(CLEAN_AND_BUILD_PROJECT_COMMAND, async (resourceUri) => {
      await cleanEpicsProject(resourceUri, makeBuildOutputChannel);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(EXPORT_DATABASE_TO_EXCEL_COMMAND, async (resourceUri) => {
      await exportDatabaseToExcelResource(resourceUri, workspaceIndex);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      IMPORT_DATABASE_FROM_EXCEL_COMMAND,
      async (resourceUri) => {
        await importDatabaseFromExcelResource(resourceUri);
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      IMPORT_DATABASE_FROM_EXCEL_SHORT_COMMAND,
      async (resourceUri) => {
        await importDatabaseFromExcelResource(resourceUri);
      },
    ),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(OPEN_EXCEL_IMPORT_PREVIEW_COMMAND, async () => {
      await openExcelImportPreviewPanel(context);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_PROBE_COMMAND, async () => {
      await openInProbeFromActiveEditor(workspaceIndex, runtimeMonitorController);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_PVLIST_COMMAND, async () => {
      await openInPvlistFromActiveEditor(workspaceIndex, runtimeMonitorController);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_MONITOR_COMMAND, async () => {
      await openInMonitorFromActiveEditor(workspaceIndex, runtimeMonitorController);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_CHANNEL_GRAPH_COMMAND, async () => {
      await openInChannelGraphFromActiveEditor(workspaceIndex);
    }),
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      UPDATE_MENU_FIELD_VALUE_COMMAND,
      async ({ uri, fieldName, start, end, value }) => {
        if (!uri || !fieldName || typeof start !== "number" || typeof end !== "number") {
          return;
        }

        const documentUri = vscode.Uri.parse(uri);
        const document = await vscode.workspace.openTextDocument(documentUri);
        const editor =
          vscode.window.visibleTextEditors.find(
            (candidate) =>
              candidate.document.uri.toString() === documentUri.toString(),
          ) ||
          (await vscode.window.showTextDocument(document, { preview: false }));
        const liveRange =
          resolveMenuFieldValueRange(document, fieldName, start, end) ||
          new vscode.Range(document.positionAt(start), document.positionAt(end));

        await editor.edit((editBuilder) => {
          editBuilder.replace(liveRange, String(value ?? ""));
        });

        const hoverPosition = liveRange.start;
        editor.selection = new vscode.Selection(hoverPosition, hoverPosition);
        editor.revealRange(new vscode.Range(hoverPosition, hoverPosition));
        await vscode.commands.executeCommand("editor.action.hideHover");
        await new Promise((resolve) => setTimeout(resolve, 25));
        await vscode.commands.executeCommand("editor.action.showHover");
      },
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerCompletionItemProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.dbd },
        { language: "makefile" },
        { scheme: "file", pattern: "**/Makefile" },
      ],
      new EpicsCompletionProvider(workspaceIndex),
      ...COMPLETION_TRIGGER_CHARACTERS,
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDefinitionProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.sequencer },
        { language: "makefile" },
        { scheme: "file", pattern: "**/Makefile" },
      ],
      new EpicsDefinitionProvider(workspaceIndex),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerReferenceProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.dbd },
        { language: LANGUAGE_IDS.pvlist },
        { language: LANGUAGE_IDS.probe },
        { language: LANGUAGE_IDS.sequencer },
        ...EPICS_SOURCE_DOCUMENT_SELECTORS,
      ],
      new EpicsReferenceProvider(workspaceIndex),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerRenameProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.dbd },
        { language: LANGUAGE_IDS.pvlist },
        { language: LANGUAGE_IDS.probe },
        ...EPICS_SOURCE_DOCUMENT_SELECTORS,
      ],
      new EpicsRenameProvider(workspaceIndex),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerHoverProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.pvlist },
        { language: LANGUAGE_IDS.sequencer },
        { language: "makefile" },
        { scheme: "file", pattern: "**/Makefile" },
        { scheme: "file", pattern: "**/envPaths*" },
      ],
      new EpicsHoverProvider(workspaceIndex),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerCodeActionsProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
      ],
      new EpicsCodeActionProvider(workspaceIndex),
      {
        providedCodeActionKinds: [vscode.CodeActionKind.QuickFix],
      },
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentSemanticTokensProvider(
      { language: LANGUAGE_IDS.database },
      new EpicsDatabaseSemanticTokensProvider(
        workspaceIndex,
        databaseSemanticTokensLegend,
      ),
      databaseSemanticTokensLegend,
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentSymbolProvider(
      { language: LANGUAGE_IDS.database },
      new EpicsDatabaseDocumentSymbolProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: LANGUAGE_IDS.database },
      new EpicsDatabaseFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: LANGUAGE_IDS.substitutions },
      new EpicsSubstitutionsFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: LANGUAGE_IDS.startup },
      new EpicsStartupFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      [
        { language: "makefile" },
        { scheme: "file", pattern: "**/Makefile" },
      ],
      new EpicsMakefileFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: LANGUAGE_IDS.proto },
      new EpicsProtocolFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerDocumentFormattingEditProvider(
      { language: LANGUAGE_IDS.sequencer },
      new EpicsSequencerFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((document) => {
      refreshActiveMakefileContextKeys();
      if (isProjectModelDocument(document)) {
        workspaceIndex.markDirty();
        refreshOpenDiagnostics();
        return;
      }

      refreshDocumentDiagnostics(document);
      refreshDocumentDecorations(document);
      refreshDatabaseValueDecorations(document);
    }),
  );
  context.subscriptions.push(
    vscode.workspace.onDidChangeTextDocument((event) => {
      refreshActiveMakefileContextKeys();
      if (isProjectModelDocument(event.document)) {
        workspaceIndex.markDirty();
        refreshOpenDiagnostics();
        return;
      }

      queueRecordTypeSuggest(event);
      queueFieldNameSuggest(event);
      refreshDocumentDiagnostics(event.document);
      refreshDocumentDecorations(event.document);
      refreshDatabaseValueDecorations(event.document);
      if (
        isDatabaseDocument(event.document) ||
        isSubstitutionsDocument(event.document)
      ) {
        refreshDependentOpenDiagnostics();
      }
    }),
  );
  context.subscriptions.push(
    vscode.workspace.onDidCloseTextDocument((document) => {
      refreshActiveMakefileContextKeys();
      diagnostics.delete(document.uri);
      if (
        isDatabaseDocument(document) ||
        isSubstitutionsDocument(document)
      ) {
        refreshDependentOpenDiagnostics();
      }
    }),
  );
  context.subscriptions.push(
    vscode.workspace.onDidSaveTextDocument((document) => {
      refreshActiveMakefileContextKeys();
      if (isProjectModelDocument(document)) {
        workspaceIndex.markDirty();
        refreshOpenDiagnostics();
        return;
      }

      refreshDocumentDiagnostics(document);
      refreshDocumentDecorations(document);
      refreshDatabaseValueDecorations(document);
      if (
        isDatabaseDocument(document) ||
        isSubstitutionsDocument(document)
      ) {
        refreshDependentOpenDiagnostics();
      }
    }),
  );
  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => {
      refreshActiveMakefileContextKeys();
      updateDatabaseRecordDecorationsForVisibleEditors(
        databaseRecordDecorationTypes,
      );
      void updateDatabaseValueDecorationsForVisibleEditors(
        workspaceIndex,
        databaseValueDecorationTypes,
      );
    }),
  );
  context.subscriptions.push(
    vscode.window.onDidChangeVisibleTextEditors(() => {
      updateDatabaseRecordDecorationsForVisibleEditors(
        databaseRecordDecorationTypes,
      );
      void updateDatabaseValueDecorationsForVisibleEditors(
        workspaceIndex,
        databaseValueDecorationTypes,
      );
    }),
  );

  for (const document of vscode.workspace.textDocuments) {
    refreshDocumentDiagnostics(document);
  }
  updateDatabaseRecordDecorationsForVisibleEditors(databaseRecordDecorationTypes);
  void updateDatabaseValueDecorationsForVisibleEditors(
    workspaceIndex,
    databaseValueDecorationTypes,
  );
}

function deactivate() {}

class EpicsCompletionProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async provideCompletionItems(document, position) {
    const anchoredRecordTypeContext = getAnchoredRecordTypeContext(document, position);
    const anchoredFieldNameContext = getAnchoredFieldNameContext(document, position);
    const context =
      anchoredRecordTypeContext ||
      anchoredFieldNameContext ||
      getCompletionContext(document, position);
    if (!context) {
      return undefined;
    }

    switch (context.type) {
      case "recordType":
        return buildRecordTypeItems(context);

      case "startupCommand":
        return buildStartupCommandItems(context);

      case "recordName":
        return buildRecordNameItems(document, context);

      default:
        break;
    }

    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );

    switch (context.type) {
      case "dbdDeviceKeyword":
        return buildDbdDeviceKeywordItems(context);

      case "dbdDeviceRecordType":
        return buildDbdDeviceRecordTypeItems(snapshot, context);

      case "dbdDeviceLinkType":
        return buildDbdDeviceLinkTypeItems(context);

      case "dbdDeviceSupportName":
        return buildDbdDeviceSupportNameItems(snapshot, context);

      case "dbdDeviceChoiceName":
        return buildDbdDeviceChoiceNameItems(context);

      case "dbdDriverKeyword":
        return buildDbdDriverKeywordItems(context);

      case "dbdDriverName":
        return buildDbdDriverNameItems(snapshot, context);

      case "dbdRegistrarKeyword":
        return buildDbdRegistrarKeywordItems(context);

      case "dbdRegistrarName":
        return buildDbdRegistrarNameItems(snapshot, document, context);

      case "dbdFunctionKeyword":
        return buildDbdFunctionKeywordItems(context);

      case "dbdFunctionName":
        return buildDbdFunctionNameItems(snapshot, document, context);

      case "dbdVariableKeyword":
        return buildDbdVariableKeywordItems(context);

      case "dbdVariableName":
        return buildDbdVariableNameItems(snapshot, document, context);

      case "fieldName":
        return buildFieldNameItems(snapshot, document, position, context);

      case "startupLoadedRecordName":
        return buildStartupLoadedRecordNameItems(snapshot, document, position, context);

      case "startupLoadMacroTail":
        return buildStartupLoadMacroTailItems(snapshot, document, position, context);

      case "startupLoadMacros":
        return buildStartupLoadMacroItems(snapshot, document, position, context);

      case "fieldValue":
        return buildFieldValueItems(snapshot, context);

      case "macroName":
        return buildLabelItems(
          getMacroCompletionLabels(snapshot, document, position),
          {
          kind: vscode.CompletionItemKind.Variable,
          range: context.range,
          detail: "EPICS macro",
          },
        );

      case "filePath":
        return buildFilePathItems(snapshot, document, context);

      case "makefileDbd":
      case "makefileLib":
        return buildMakefileReferenceItems(snapshot, document, context);

      default:
        return undefined;
    }
  }
}

class EpicsDatabaseDocumentSymbolProvider {
  provideDocumentSymbols(document) {
    if (!isDatabaseDocument(document)) {
      return [];
    }

    const text = document.getText();
    return extractRecordDeclarations(text)
      .map((declaration) =>
        createRecordDocumentSymbol(document, declaration, text.length),
      )
      .filter(Boolean);
  }
}

class EpicsDefinitionProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async provideDefinition(document, position) {
    if (isSequencerDocument(document)) {
      return getSequencerDefinitionLocation(document, position);
    }

    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    const navigationTarget = getNavigationTarget(snapshot, document, position);
    if (!navigationTarget) {
      return undefined;
    }

    return new vscode.Location(
      vscode.Uri.file(navigationTarget.absolutePath),
      new vscode.Position(
        Math.max(0, Number(navigationTarget.line || 1) - 1),
        Math.max(0, Number(navigationTarget.character || 1) - 1),
      ),
    );
  }
}

class EpicsReferenceProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async provideReferences(document, position, context) {
    if (isSequencerDocument(document)) {
      return getSequencerReferenceLocations(
        document,
        position,
        Boolean(context?.includeDeclaration),
      );
    }

    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    return getEpicsReferenceLocations(
      snapshot,
      document,
      position,
      true,
    );
  }
}

class EpicsRenameProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async prepareRename(document, position) {
    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    const symbol = getEpicsSemanticSymbolAtPosition(snapshot, document, position);
    if (!symbol || symbol.readOnly) {
      throw new Error("Nothing renameable at the current cursor position.");
    }

    return {
      range: symbol.range,
      placeholder: symbol.name,
    };
  }

  async provideRenameEdits(document, position, newName) {
    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    return buildEpicsRenameWorkspaceEdit(snapshot, document, position, newName);
  }
}

class EpicsHoverProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async provideHover(document, position) {
    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    return getHover(snapshot, document, position);
  }
}

class EpicsCodeActionProvider {
  constructor(workspaceIndex) {
    this.workspaceIndex = workspaceIndex;
  }

  async provideCodeActions(document, range, context) {
    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    return getEpicsCodeActions(snapshot, document, range, context);
  }
}

class EpicsDatabaseSemanticTokensProvider {
  constructor(workspaceIndex, legend) {
    this.workspaceIndex = workspaceIndex;
    this.legend = legend;
  }

  async provideDocumentSemanticTokens(document) {
    if (!isDatabaseDocument(document)) {
      return new vscode.SemanticTokensBuilder(this.legend).build();
    }

    const snapshot = mergeSnapshotWithDocument(
      await this.workspaceIndex.getSnapshot(),
      document,
    );
    return buildDatabaseSemanticTokens(snapshot, document, this.legend);
  }
}

class EpicsDatabaseFormattingProvider {
  provideDocumentFormattingEdits(document, options) {
    const formattedText = formatDatabaseText(document.getText(), options);
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

class EpicsSubstitutionsFormattingProvider {
  provideDocumentFormattingEdits(document, options) {
    const formattedText = formatSubstitutionText(document.getText(), options);
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

class EpicsProtocolFormattingProvider {
  provideDocumentFormattingEdits(document, options) {
    const formattedText = formatProtocolText(document.getText(), options);
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

class EpicsStartupFormattingProvider {
  provideDocumentFormattingEdits(document) {
    const formattedText = formatStartupText(document.getText());
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

class EpicsMakefileFormattingProvider {
  provideDocumentFormattingEdits(document) {
    const formattedText = formatMakefileText(document.getText());
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

class EpicsSequencerFormattingProvider {
  provideDocumentFormattingEdits(document) {
    const formattedText = formatSequencerText(document.getText());
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

class WorkspaceIndex {
  constructor(staticData) {
    this.staticData = staticData;
    this.snapshot = createEmptySnapshot(staticData);
    this.dirty = true;
    this.rebuildPromise = undefined;
    this.disposables = [];
    this.buildModelCache = new EpicsBuildModelCache();
    this.changeEmitter = new vscode.EventEmitter();
    this.disposables.push(this.changeEmitter);

    registerWatcher(this.disposables, INDEX_GLOB, () => this.handleWorkspaceChange());
    registerWatcher(
      this.disposables,
      SOURCE_INDEX_GLOB,
      () => this.handleWorkspaceChange(),
    );
    for (const glob of PROJECT_INDEX_GLOBS) {
      registerWatcher(this.disposables, glob, () => this.handleWorkspaceChange());
    }
  }

  onDidChange(listener, thisArg) {
    return this.changeEmitter.event(listener, thisArg);
  }

  async getSnapshot() {
    await this.ensureFresh();
    return this.snapshot;
  }

  async ensureFresh() {
    if (!this.dirty) {
      return;
    }

    if (!this.rebuildPromise) {
      this.rebuildPromise = this.rebuild();
    }

    await this.rebuildPromise;
  }

  markDirty() {
    this.dirty = true;
  }

  handleWorkspaceChange() {
    this.markDirty();
    this.changeEmitter.fire();
  }

  async rebuild() {
    const snapshot = createEmptySnapshot(this.staticData);
    const uris = await collectWorkspaceUris();
    const projectFiles = [];

    for (const uri of uris) {
      const text = await readWorkspaceFile(uri);
      if (text === undefined) {
        continue;
      }

      if (isIndexedContentFile(uri)) {
        applyParsedData(snapshot, parseDocumentText(uri, text));
        snapshot.workspaceFiles.push(createWorkspaceFileEntry(uri));
      }

      if (isProjectModelUri(uri)) {
        projectFiles.push({ uri, text });
      }
    }

    snapshot.projectModel = await buildProjectModel(projectFiles, this.buildModelCache);
    snapshot.workspaceFilesByAbsolutePath = buildWorkspaceFileLookup(snapshot.workspaceFiles);
    snapshot.workspaceFiles.sort((left, right) =>
      compareLabels(left.relativePath, right.relativePath),
    );

    this.snapshot = snapshot;
    this.dirty = false;
    this.rebuildPromise = undefined;
  }

  dispose() {
    this.buildModelCache.dispose();
    this.disposables.forEach((disposable) => disposable.dispose());
  }
}

class EpicsBuildModelCache {
  constructor() {
    this.entries = new Map();
  }

  async getApplication(rootPath, filesByAbsolutePath) {
    const normalizedRootPath = normalizeFsPath(rootPath);
    const signature = computeProjectBuildSignature(normalizedRootPath, filesByAbsolutePath);
    const cached = this.entries.get(normalizedRootPath);
    if (cached?.signature === signature) {
      return cached.application;
    }

    try {
      const application = await collectEpicsBuildApplication(normalizedRootPath);
      if (application) {
        this.entries.set(normalizedRootPath, {
          signature,
          application,
        });
      }
      return application;
    } catch (error) {
      return cached?.application;
    }
  }

  dispose() {
    this.entries.clear();
  }
}

function computeProjectBuildSignature(rootPath, filesByAbsolutePath) {
  const hash = crypto.createHash("sha1");
  const relevantFiles = [...filesByAbsolutePath.entries()]
    .filter(([filePath]) => isPathWithinRoot(filePath, rootPath))
    .sort((left, right) => compareLabels(left[0], right[0]));

  for (const [filePath, file] of relevantFiles) {
    hash.update(filePath);
    hash.update("\0");
    hash.update(String(file?.text || ""));
    hash.update("\0");
  }

  return hash.digest("hex");
}

function loadStaticData(extensionPath) {
  const grammar = readJsonFile(path.join(extensionPath, "syntaxes", "db.tmLanguage.json"));
  const embeddedRecordFields =
    readJsonFile(path.join(extensionPath, "data", "embedded-record-fields.json")) || {};
  const embeddedRecordFieldTypes =
    readJsonFile(path.join(extensionPath, "data", "embedded-record-field-types.json")) || {};
  const embeddedRecordFieldMenus =
    readJsonFile(path.join(extensionPath, "data", "embedded-record-field-menus.json")) || {};
  const embeddedRecordFieldInitials =
    readJsonFile(path.join(extensionPath, "data", "embedded-record-field-initials.json")) || {};
  const recordTemplateFields =
    readJsonFile(path.join(extensionPath, "data", "record-template-fields.json")) || {};

  const recordTypes = new Set(
    extractAlternationTerms(
      grammar?.repository?.record_types?.patterns?.[0]?.match,
    ),
  );
  const allFields = new Set(
    extractAlternationTerms(
      grammar?.repository?.field_types?.patterns?.[0]?.match,
    ),
  );
  const fieldsByRecordType = new Map();
  const fieldTypesByRecordType = createFieldTypeMap(embeddedRecordFieldTypes);
  const fieldMenuChoicesByRecordType = createFieldMenuChoiceMap(embeddedRecordFieldMenus);
  const fieldInitialValuesByRecordType = createFieldInitialValueMap(
    embeddedRecordFieldInitials,
  );

  for (const [recordType, fieldNames] of Object.entries(recordTemplateFields)) {
    recordTypes.add(recordType);
    addToMapOfSets(fieldsByRecordType, recordType, fieldNames);
    for (const fieldName of fieldNames) {
      allFields.add(fieldName);
    }
  }

  for (const [recordType, fieldNames] of Object.entries(embeddedRecordFields)) {
    recordTypes.add(recordType);
    addToMapOfSets(fieldsByRecordType, recordType, fieldNames);
    for (const fieldName of fieldNames) {
      allFields.add(fieldName);
    }
  }

  for (const fieldName of COMMON_RECORD_FIELDS) {
    allFields.add(fieldName);
  }

  return {
    recordTypes,
    allFields,
    fieldsByRecordType,
    fieldOrderByRecordType: new Map(Object.entries(embeddedRecordFields)),
    recordTemplateFields: new Map(Object.entries(recordTemplateFields)),
    fieldTypesByRecordType,
    fieldMenuChoicesByRecordType,
    fieldInitialValuesByRecordType,
  };
}

function getCompletionContext(document, position) {
  const linePrefix = document.lineAt(position).text.slice(0, position.character);

  const macroContext = getMacroContext(document, position, linePrefix);
  if (macroContext) {
    return macroContext;
  }

  if (document.languageId === LANGUAGE_IDS.database) {
    const recordNameContext = getRegexContext(
      position,
      linePrefix,
      /record\(\s*[A-Za-z0-9_]+\s*,\s*"([^"\n]*)$/,
      "recordName",
    );
    if (recordNameContext) {
      return recordNameContext;
    }

    const fieldNameContext = getRegexContext(
      position,
      linePrefix,
      /field\(\s*(?:"?([A-Za-z0-9_]*))$/,
      "fieldName",
    );
    if (fieldNameContext) {
      const textBefore = document.getText(
        new vscode.Range(new vscode.Position(0, 0), position),
      );
      fieldNameContext.recordType =
        findEnclosingRecordDeclaration(
          document.getText(),
          document.offsetAt(position),
        )?.recordType || findEnclosingRecordType(textBefore);
      return fieldNameContext;
    }

    const fieldValueContext = getFieldValueContext(position, linePrefix);
    if (fieldValueContext) {
      const textBefore = document.getText(
        new vscode.Range(new vscode.Position(0, 0), position),
      );
      fieldValueContext.recordType = findEnclosingRecordType(textBefore);
      return fieldValueContext;
    }

    const includeContext = getRegexContext(
      position,
      linePrefix,
      /include\s+"([^"\n]*)$/,
      "filePath",
    );
    if (includeContext) {
      includeContext.fileKind = "databaseInclude";
      return includeContext;
    }
  }

  if (document.languageId === LANGUAGE_IDS.startup) {
    const startupCommandContext = getStartupCommandContext(position, linePrefix);
    if (startupCommandContext) {
      return startupCommandContext;
    }

    const startupLoadedRecordNameContext = getStartupLoadedRecordNameContext(
      position,
      linePrefix,
    );
    if (startupLoadedRecordNameContext) {
      return startupLoadedRecordNameContext;
    }

    const startupIncludeContext = getRegexContext(
      position,
      linePrefix,
      /^[ \t]*<\s*"?([^"\n]*)$/,
      "filePath",
    );
    if (startupIncludeContext) {
      startupIncludeContext.fileKind = "startupInclude";
      return startupIncludeContext;
    }

    const startupDirectoryContext = getRegexContext(
      position,
      linePrefix,
      /cd(?:\s+|\(\s*)"([^"\n]*)$/,
      "filePath",
    );
    if (startupDirectoryContext) {
      startupDirectoryContext.fileKind = "startupDirectory";
      return startupDirectoryContext;
    }

    const dbLoadDatabaseContext = getRegexContext(
      position,
      linePrefix,
      /dbLoadDatabase(?:\(\s*|\s+)"([^"\n]*)$/,
      "filePath",
    );
    if (dbLoadDatabaseContext) {
      dbLoadDatabaseContext.fileKind = "dbLoadDatabase";
      return dbLoadDatabaseContext;
    }

    const dbLoadRecordsContext = getRegexContext(
      position,
      linePrefix,
      /dbLoadRecords\(\s*"([^"\n]*)$/,
      "filePath",
    );
    if (dbLoadRecordsContext) {
      dbLoadRecordsContext.fileKind = "dbLoadRecords";
      return dbLoadRecordsContext;
    }

    const dbLoadRecordsMacroTailContext = getStartupLoadMacroTailCompletionContext(
      position,
      linePrefix,
    );
    if (dbLoadRecordsMacroTailContext) {
      return dbLoadRecordsMacroTailContext;
    }

    const dbLoadRecordsMacroContext = getStartupLoadMacroCompletionContext(
      position,
      linePrefix,
    );
    if (dbLoadRecordsMacroContext) {
      return dbLoadRecordsMacroContext;
    }

    const dbLoadTemplateContext = getRegexContext(
      position,
      linePrefix,
      /dbLoadTemplate\(\s*"([^"\n]*)$/,
      "filePath",
    );
    if (dbLoadTemplateContext) {
      dbLoadTemplateContext.fileKind = "dbLoadTemplate";
      return dbLoadTemplateContext;
    }

    const includeContext = getRegexContext(
      position,
      linePrefix,
      /include\s+"([^"\n]*)$/,
      "filePath",
    );
    if (includeContext) {
      includeContext.fileKind = "startupInclude";
      return includeContext;
    }
  }

  if (document.languageId === LANGUAGE_IDS.dbd) {
    const dbdKeywordDeviceContext = getRegexContext(
      position,
      linePrefix,
      /^\s*(dev[A-Za-z_]*)$/,
      "dbdDeviceKeyword",
    );
    if (dbdKeywordDeviceContext) {
      return dbdKeywordDeviceContext;
    }

    const dbdKeywordDriverContext = getRegexContext(
      position,
      linePrefix,
      /^\s*(drv[A-Za-z_]*)$/,
      "dbdDriverKeyword",
    );
    if (dbdKeywordDriverContext) {
      return dbdKeywordDriverContext;
    }

    const dbdKeywordRegistrarContext = getRegexContext(
      position,
      linePrefix,
      /^\s*(reg[A-Za-z_]*)$/,
      "dbdRegistrarKeyword",
    );
    if (dbdKeywordRegistrarContext) {
      return dbdKeywordRegistrarContext;
    }

    const dbdKeywordFunctionContext = getRegexContext(
      position,
      linePrefix,
      /^\s*(fun[A-Za-z_]*)$/,
      "dbdFunctionKeyword",
    );
    if (dbdKeywordFunctionContext) {
      return dbdKeywordFunctionContext;
    }

    const dbdKeywordVariableContext = getRegexContext(
      position,
      linePrefix,
      /^\s*(var[A-Za-z_]*)$/,
      "dbdVariableKeyword",
    );
    if (dbdKeywordVariableContext) {
      return dbdKeywordVariableContext;
    }

    const dbdDeviceContext = getDbdDeviceCompletionContext(position, linePrefix);
    if (dbdDeviceContext) {
      return dbdDeviceContext;
    }

    const dbdDriverContext = getDbdDriverCompletionContext(position, linePrefix);
    if (dbdDriverContext) {
      return dbdDriverContext;
    }

    const dbdRegistrarContext = getDbdRegistrarCompletionContext(position, linePrefix);
    if (dbdRegistrarContext) {
      return dbdRegistrarContext;
    }

    const dbdFunctionContext = getDbdFunctionCompletionContext(position, linePrefix);
    if (dbdFunctionContext) {
      return dbdFunctionContext;
    }

    const dbdVariableContext = getDbdVariableCompletionContext(position, linePrefix);
    if (dbdVariableContext) {
      return dbdVariableContext;
    }

    const includeContext = getRegexContext(
      position,
      linePrefix,
      /include\s+"([^"\n]*)$/,
      "filePath",
    );
    if (includeContext) {
      includeContext.fileKind = "dbdInclude";
      return includeContext;
    }
  }

  if (isSourceMakefileDocument(document)) {
    const makefileContext = getMakefileReferenceContext(position, linePrefix);
    if (makefileContext) {
      return makefileContext;
    }
  }

  return undefined;
}

function getStartupCommandContext(position, linePrefix) {
  const match = linePrefix.match(/^\s*(db[A-Za-z]*)$/);
  if (!match) {
    return undefined;
  }

  const partial = match[1] || "";
  return {
    type: "startupCommand",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getStartupLoadedRecordNameContext(position, linePrefix) {
  const match = linePrefix.match(/dbpf\(\s*"([^"\n]*)$/);
  if (!match) {
    return undefined;
  }

  const partial = match[1] || "";
  return {
    type: "startupLoadedRecordName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getStartupLoadMacroCompletionContext(position, linePrefix) {
  const match = linePrefix.match(
    /dbLoadRecords\(\s*"([^"\n]+)"\s*,\s*"([^"\n]*)$/,
  );
  if (!match) {
    return undefined;
  }

  const macroText = match[2] || "";
  const lastCommaIndex = macroText.lastIndexOf(",");
  const segmentStartIndex = lastCommaIndex >= 0 ? lastCommaIndex + 1 : 0;
  const segmentText = macroText.slice(segmentStartIndex);
  const leadingWhitespaceLength = segmentText.match(/^\s*/)?.[0].length || 0;
  const partial = segmentText.slice(leadingWhitespaceLength);
  const completedText = macroText.slice(0, segmentStartIndex);
  return {
    type: "startupLoadMacros",
    path: match[1],
    macroText,
    completedText,
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getStartupLoadMacroTailCompletionContext(position, linePrefix) {
  const match = linePrefix.match(/dbLoadRecords\(\s*"([^"\n]+)"\s*$/);
  if (!match) {
    return undefined;
  }

  return {
    type: "startupLoadMacroTail",
    path: match[1],
    range: new vscode.Range(
      position.line,
      position.character,
      position.line,
      position.character,
    ),
  };
}

function getAnchoredFieldNameContext(document, position) {
  if (!isDatabaseDocument(document) || !fieldNameSuggestAnchor) {
    return undefined;
  }

  if (
    fieldNameSuggestAnchor.uri !== document.uri.toString() ||
    fieldNameSuggestAnchor.line !== position.line ||
    position.character < fieldNameSuggestAnchor.character
  ) {
    return undefined;
  }

  const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
  const match = linePrefix.match(/field\(\s*(?:"?([A-Za-z0-9_]*))$/);
  if (!match) {
    fieldNameSuggestAnchor = undefined;
    return undefined;
  }

  const partial = match[1] || "";
  return {
    type: "fieldName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
    recordType:
      fieldNameSuggestAnchor.recordType ||
      findEnclosingRecordDeclaration(
        document.getText(),
        document.offsetAt(position),
      )?.recordType,
  };
}

function getAnchoredRecordTypeContext(document, position) {
  if (!isDatabaseDocument(document) || !recordTypeSuggestAnchor) {
    return undefined;
  }

  if (
    recordTypeSuggestAnchor.uri !== document.uri.toString() ||
    recordTypeSuggestAnchor.line !== position.line ||
    position.character < recordTypeSuggestAnchor.character
  ) {
    return undefined;
  }

  const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
  const match = linePrefix.match(/record\(\s*([A-Za-z0-9_]*)$/);
  if (!match) {
    recordTypeSuggestAnchor = undefined;
    return undefined;
  }

  const partial = match[1] || "";
  return {
    type: "recordType",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getRegexContext(position, linePrefix, regex, type, options = {}) {
  const match = linePrefix.match(regex);
  if (!match) {
    return undefined;
  }

  const capturedText = match[1] || "";
  const partial = options.replaceWithEmptyPartial ? "" : capturedText;
  return {
    type,
    partial,
    range: new vscode.Range(
      position.line,
      position.character - capturedText.length,
      position.line,
      position.character,
    ),
  };
}

function queueRecordTypeSuggest(event) {
  if (!event || !isDatabaseDocument(event.document) || !event.contentChanges?.length) {
    return;
  }

  const activeEditor = vscode.window.activeTextEditor;
  if (
    !activeEditor ||
    activeEditor.document.uri.toString() !== event.document.uri.toString()
  ) {
    return;
  }

  for (const change of event.contentChanges) {
    if (
      (change.text !== "(" && change.text !== "()") ||
      change.rangeLength !== 0 ||
      change.range.start.line !== change.range.end.line
    ) {
      continue;
    }

    const cursorCharacter = change.range.start.character + 1;
    const linePrefix = event.document.lineAt(change.range.start.line).text.slice(
      0,
      cursorCharacter,
    );
    if (!/record\(\s*$/.test(linePrefix)) {
      continue;
    }

    pendingRecordTypeSuggest = {
      uri: event.document.uri.toString(),
      line: change.range.start.line,
      character: cursorCharacter,
    };
    schedulePendingRecordTypeSuggest();
    return;
  }
}

function queueFieldNameSuggest(event) {
  if (!event || !isDatabaseDocument(event.document) || !event.contentChanges?.length) {
    return;
  }

  const activeEditor = vscode.window.activeTextEditor;
  if (
    !activeEditor ||
    activeEditor.document.uri.toString() !== event.document.uri.toString()
  ) {
    return;
  }

  for (const change of event.contentChanges) {
    if (
      (change.text !== "(" && change.text !== "()") ||
      change.rangeLength !== 0 ||
      change.range.start.line !== change.range.end.line
    ) {
      continue;
    }

    const cursorCharacter = change.range.start.character + 1;
    const linePrefix = event.document.lineAt(change.range.start.line).text.slice(
      0,
      cursorCharacter,
    );
    if (!/field\(\s*$/.test(linePrefix)) {
      continue;
    }

    pendingFieldNameSuggest = {
      uri: event.document.uri.toString(),
      line: change.range.start.line,
      character: cursorCharacter,
    };
    schedulePendingFieldNameSuggest();
    return;
  }
}

function schedulePendingRecordTypeSuggest() {
  clearTimeout(pendingRecordTypeSuggestTimer);
  attemptPendingRecordTypeSuggest(0);
}

function schedulePendingRecordNameSuggest(uri) {
  pendingRecordNameSuggest = { uri };
  clearTimeout(pendingRecordNameSuggestTimer);
  attemptPendingRecordNameSuggest(0);
}

function schedulePendingFieldNameSuggest() {
  clearTimeout(pendingFieldNameSuggestTimer);
  attemptPendingFieldNameSuggest(0);
}

function attemptPendingRecordTypeSuggest(attemptIndex) {
  const retryDelays = [0, 20, 60, 150];
  if (!pendingRecordTypeSuggest) {
    return;
  }

  pendingRecordTypeSuggestTimer = setTimeout(() => {
    if (!pendingRecordTypeSuggest) {
      return;
    }

    const activeEditor = vscode.window.activeTextEditor;
    if (!activeEditor) {
      scheduleNextPendingRecordTypeSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    if (activeEditor.document.uri.toString() !== pendingRecordTypeSuggest.uri) {
      scheduleNextPendingRecordTypeSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const selection = activeEditor.selection;
    if (
      !selection ||
      !selection.isEmpty ||
      selection.active.line !== pendingRecordTypeSuggest.line ||
      selection.active.character !== pendingRecordTypeSuggest.character
    ) {
      scheduleNextPendingRecordTypeSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const linePrefix = activeEditor.document.lineAt(selection.active.line).text.slice(
      0,
      selection.active.character,
    );
    if (!/record\(\s*$/.test(linePrefix)) {
      scheduleNextPendingRecordTypeSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    pendingRecordTypeSuggest = undefined;
    clearTimeout(pendingRecordTypeSuggestTimer);
    pendingRecordTypeSuggestTimer = undefined;
    recordTypeSuggestAnchor = {
      uri: activeEditor.document.uri.toString(),
      line: selection.active.line,
      character: selection.active.character,
    };
    void vscode.commands.executeCommand("editor.action.triggerSuggest");
  }, retryDelays[Math.min(attemptIndex, retryDelays.length - 1)]);
}

function scheduleNextPendingRecordTypeSuggestAttempt(attemptIndex, retryDelays) {
  if (!pendingRecordTypeSuggest) {
    clearTimeout(pendingRecordTypeSuggestTimer);
    pendingRecordTypeSuggestTimer = undefined;
    return;
  }

  if (attemptIndex >= retryDelays.length - 1) {
    pendingRecordTypeSuggest = undefined;
    clearTimeout(pendingRecordTypeSuggestTimer);
    pendingRecordTypeSuggestTimer = undefined;
    return;
  }

  attemptPendingRecordTypeSuggest(attemptIndex + 1);
}

function attemptPendingRecordNameSuggest(attemptIndex) {
  const retryDelays = [0, 20, 60, 150];
  if (!pendingRecordNameSuggest) {
    return;
  }

  pendingRecordNameSuggestTimer = setTimeout(() => {
    if (!pendingRecordNameSuggest) {
      return;
    }

    const activeEditor = vscode.window.activeTextEditor;
    if (!activeEditor) {
      scheduleNextPendingRecordNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    if (activeEditor.document.uri.toString() !== pendingRecordNameSuggest.uri) {
      scheduleNextPendingRecordNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const selection = activeEditor.selection;
    if (!selection) {
      scheduleNextPendingRecordNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const linePrefix = activeEditor.document.lineAt(selection.active.line).text.slice(
      0,
      selection.active.character,
    );
    if (!/record\(\s*[A-Za-z0-9_]+\s*,\s*"([^"\n]*)$/.test(linePrefix)) {
      scheduleNextPendingRecordNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    pendingRecordNameSuggest = undefined;
    clearTimeout(pendingRecordNameSuggestTimer);
    pendingRecordNameSuggestTimer = undefined;
    void vscode.commands.executeCommand("editor.action.triggerSuggest");
  }, retryDelays[Math.min(attemptIndex, retryDelays.length - 1)]);
}

function scheduleNextPendingRecordNameSuggestAttempt(attemptIndex, retryDelays) {
  if (!pendingRecordNameSuggest) {
    clearTimeout(pendingRecordNameSuggestTimer);
    pendingRecordNameSuggestTimer = undefined;
    return;
  }

  if (attemptIndex >= retryDelays.length - 1) {
    pendingRecordNameSuggest = undefined;
    clearTimeout(pendingRecordNameSuggestTimer);
    pendingRecordNameSuggestTimer = undefined;
    return;
  }

  attemptPendingRecordNameSuggest(attemptIndex + 1);
}

function attemptPendingFieldNameSuggest(attemptIndex) {
  const retryDelays = [0, 20, 60, 150];
  if (!pendingFieldNameSuggest) {
    return;
  }

  pendingFieldNameSuggestTimer = setTimeout(() => {
    if (!pendingFieldNameSuggest) {
      return;
    }

    const activeEditor = vscode.window.activeTextEditor;
    if (!activeEditor) {
      scheduleNextPendingFieldNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    if (activeEditor.document.uri.toString() !== pendingFieldNameSuggest.uri) {
      scheduleNextPendingFieldNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const selection = activeEditor.selection;
    if (
      !selection ||
      !selection.isEmpty ||
      selection.active.line !== pendingFieldNameSuggest.line ||
      selection.active.character !== pendingFieldNameSuggest.character
    ) {
      scheduleNextPendingFieldNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    const linePrefix = activeEditor.document.lineAt(selection.active.line).text.slice(
      0,
      selection.active.character,
    );
    if (!/field\(\s*$/.test(linePrefix)) {
      scheduleNextPendingFieldNameSuggestAttempt(attemptIndex, retryDelays);
      return;
    }

    pendingFieldNameSuggest = undefined;
    clearTimeout(pendingFieldNameSuggestTimer);
    pendingFieldNameSuggestTimer = undefined;
    const recordDeclaration = findEnclosingRecordDeclaration(
      activeEditor.document.getText(),
      activeEditor.document.offsetAt(selection.active),
    );
    fieldNameSuggestAnchor = {
      uri: activeEditor.document.uri.toString(),
      line: selection.active.line,
      character: selection.active.character,
      recordType: recordDeclaration?.recordType,
    };
    void vscode.commands.executeCommand("editor.action.triggerSuggest");
  }, retryDelays[Math.min(attemptIndex, retryDelays.length - 1)]);
}

function scheduleNextPendingFieldNameSuggestAttempt(attemptIndex, retryDelays) {
  if (!pendingFieldNameSuggest) {
    clearTimeout(pendingFieldNameSuggestTimer);
    pendingFieldNameSuggestTimer = undefined;
    return;
  }

  if (attemptIndex >= retryDelays.length - 1) {
    pendingFieldNameSuggest = undefined;
    clearTimeout(pendingFieldNameSuggestTimer);
    pendingFieldNameSuggestTimer = undefined;
    return;
  }

  attemptPendingFieldNameSuggest(attemptIndex + 1);
}

function getRecordTailInsertionRange(document, position) {
  const lineText = document.lineAt(position.line).text;
  let endCharacter = position.character;

  while (
    endCharacter < lineText.length &&
    /[A-Za-z0-9_]/.test(lineText[endCharacter])
  ) {
    endCharacter += 1;
  }

  if (lineText[endCharacter] === ")") {
    endCharacter += 1;
  }

  return new vscode.Range(
    position.line,
    position.character,
    position.line,
    endCharacter,
  );
}

function buildRecordTemplateTailSnippet(recordType, fieldNames) {
  if (!Array.isArray(fieldNames) || fieldNames.length === 0) {
    return undefined;
  }

  const lines = [`, "${buildSnippetPlaceholder(1, "")}") {`];
  let placeholderIndex = 2;

  for (const fieldName of fieldNames) {
    const defaultValue = getDefaultFieldValue(
      recordTemplateStaticData,
      recordType,
      fieldName,
    );
    lines.push(
      `    field(${fieldName}, "${buildSnippetPlaceholder(placeholderIndex, defaultValue)}")`,
    );
    placeholderIndex += 1;
  }

  lines.push(`}\${0}`);
  return lines.join("\n");
}

function getMacroContext(document, position, linePrefix) {
  const macroMatch =
    linePrefix.match(/\$\(([A-Za-z0-9_]*)$/) ||
    linePrefix.match(/\$\{([A-Za-z0-9_]*)$/) ||
    (document.languageId === LANGUAGE_IDS.startup
      ? linePrefix.match(/\$([A-Za-z0-9_]*)$/)
      : undefined);

  if (!macroMatch) {
    return undefined;
  }

  const partial = macroMatch[1] || "";
  return {
    type: "macroName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getFieldValueContext(position, linePrefix) {
  const match = linePrefix.match(
    /field\(\s*(?:"([A-Za-z0-9_]+)"|([A-Za-z0-9_]+))\s*,\s*"([^"\n]*)$/,
  );
  if (!match) {
    return undefined;
  }

  const fieldName = match[1] || match[2];
  const partial = match[3] || "";
  return {
    type: "fieldValue",
    fieldName,
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getDbdDeviceCompletionContext(position, linePrefix) {
  const deviceRecordTypeMatch = linePrefix.match(/^\s*device\(\s*([A-Za-z0-9_]*)$/);
  if (deviceRecordTypeMatch) {
    const partial = deviceRecordTypeMatch[1] || "";
    return {
      type: "dbdDeviceRecordType",
      partial,
      range: new vscode.Range(
        position.line,
        position.character - partial.length,
        position.line,
        position.character,
      ),
    };
  }

  const deviceLinkTypeMatch = linePrefix.match(
    /^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]*)$/,
  );
  if (deviceLinkTypeMatch) {
    const partial = deviceLinkTypeMatch[1] || "";
    return {
      type: "dbdDeviceLinkType",
      partial,
      range: new vscode.Range(
        position.line,
        position.character - partial.length,
        position.line,
        position.character,
      ),
    };
  }

  const deviceSupportNameMatch = linePrefix.match(
    /^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]*)$/,
  );
  if (deviceSupportNameMatch) {
    const partial = deviceSupportNameMatch[1] || "";
    return {
      type: "dbdDeviceSupportName",
      partial,
      range: new vscode.Range(
        position.line,
        position.character - partial.length,
        position.line,
        position.character,
      ),
    };
  }

  const deviceChoiceNameMatch = linePrefix.match(
    /^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]+)\s*,\s*"([^"\n]*)$/,
  );
  if (deviceChoiceNameMatch) {
    const supportName = deviceChoiceNameMatch[1];
    const partial = deviceChoiceNameMatch[2] || "";
    return {
      type: "dbdDeviceChoiceName",
      supportName,
      partial,
      range: new vscode.Range(
        position.line,
        position.character - partial.length,
        position.line,
        position.character,
      ),
    };
  }

  return undefined;
}

function getDbdDriverCompletionContext(position, linePrefix) {
  const driverNameMatch = linePrefix.match(/^\s*driver\(\s*([A-Za-z0-9_]*)$/);
  if (!driverNameMatch) {
    return undefined;
  }

  const partial = driverNameMatch[1] || "";
  return {
    type: "dbdDriverName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getDbdRegistrarCompletionContext(position, linePrefix) {
  const registrarNameMatch = linePrefix.match(/^\s*registrar\(\s*([A-Za-z0-9_]*)$/);
  if (!registrarNameMatch) {
    return undefined;
  }

  const partial = registrarNameMatch[1] || "";
  return {
    type: "dbdRegistrarName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getDbdFunctionCompletionContext(position, linePrefix) {
  const functionNameMatch = linePrefix.match(/^\s*function\(\s*([A-Za-z0-9_]*)$/);
  if (!functionNameMatch) {
    return undefined;
  }

  const partial = functionNameMatch[1] || "";
  return {
    type: "dbdFunctionName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getDbdVariableCompletionContext(position, linePrefix) {
  const variableNameMatch = linePrefix.match(/^\s*variable\(\s*([A-Za-z0-9_]*)$/);
  if (!variableNameMatch) {
    return undefined;
  }

  const partial = variableNameMatch[1] || "";
  return {
    type: "dbdVariableName",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function getMakefileReferenceContext(position, linePrefix) {
  const prefixMatch = linePrefix.match(
    /^\s*(?:[A-Za-z0-9_.-]+_)?(DBD|LIBS)\s*(?:\+?=|:=|\?=)\s*/,
  );
  if (!prefixMatch) {
    return undefined;
  }

  const referenceKind = prefixMatch[1] === "DBD" ? "dbd" : "lib";
  const valuePrefix = linePrefix.slice(prefixMatch[0].length);
  if (valuePrefix.includes("#")) {
    return undefined;
  }

  const tokenMatch = valuePrefix.match(/(?:^|\s)([^\s]*)$/);
  const partial = tokenMatch ? tokenMatch[1] || "" : "";
  if (containsMakeVariableReference(partial)) {
    return undefined;
  }

  return {
    type: referenceKind === "dbd" ? "makefileDbd" : "makefileLib",
    partial,
    range: new vscode.Range(
      position.line,
      position.character - partial.length,
      position.line,
      position.character,
    ),
  };
}

function buildLabelItems(labels, options) {
  const seen = new Set();
  const items = [];

  for (const label of labels) {
    if (!label || seen.has(label)) {
      continue;
    }

    seen.add(label);
    const item = new vscode.CompletionItem(label, options.kind);
    item.range = options.range;
    item.detail = options.detail;
    items.push(item);
  }

  return items;
}

function buildRecordNameItems(document, context) {
  const names = extractRecordDeclarations(document.getText())
    .map((declaration) => declaration.name)
    .filter((name) => name && matchesCompletionQuery(name, context.partial))
    .sort(compareLabels);

  const items = [];
  const seen = new Set();
  for (const name of names) {
    if (seen.has(name)) {
      continue;
    }

    seen.add(name);
    const item = new vscode.CompletionItem(
      name,
      vscode.CompletionItemKind.Reference,
    );
    item.range = context.range;
    item.detail = "Existing record name in this file";
    item.filterText = buildFilterText(name);
    items.push(item);
  }

  return items;
}

function buildStartupLoadedRecordNameItems(snapshot, document, position, context) {
  const names = getStartupLoadedRecordNames(snapshot, document, position).filter((name) =>
    matchesCompletionQuery(name, context.partial),
  );
  const items = [];
  const seen = new Set();

  for (const name of names) {
    if (!name || seen.has(name)) {
      continue;
    }

    seen.add(name);
    const item = new vscode.CompletionItem(
      name,
      vscode.CompletionItemKind.Reference,
    );
    item.range = context.range;
    item.detail = "Record loaded by this startup script";
    item.filterText = buildFilterText(name);
    item.sortText = buildSortText(name, context.partial);
    items.push(item);
  }

  return items;
}

function buildStartupCommandItems(context) {
  const items = [];

  for (const commandName of [...STARTUP_COMMAND_TEMPLATES.keys()].sort(compareLabels)) {
    if (!matchesCompletionQuery(commandName, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(
      commandName,
      vscode.CompletionItemKind.Function,
    );
    item.range = context.range;
    item.insertText = commandName;
    item.detail = "EPICS startup command";
    item.filterText = buildFilterText(commandName);
    item.sortText = buildSortText(commandName, context.partial);
    item.command = {
      command: INSERT_STARTUP_COMMAND_TAIL_COMMAND,
      title: "Insert EPICS startup command tail",
      arguments: [{ commandName }],
    };
    items.push(item);
  }

  return items;
}

function buildRecordTypeItems(context) {
  const items = [];

  for (const recordType of [...recordTemplateFields.keys()].sort(compareLabels)) {
    if (!matchesCompletionQuery(recordType, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(
      recordType,
      vscode.CompletionItemKind.Class,
    );
    item.range = context.range;
    item.insertText = recordType;
    item.detail = "EPICS record type";
    item.filterText = buildFilterText(recordType);
    item.sortText = buildSortText(recordType, context.partial);
    item.command = {
      command: INSERT_RECORD_TAIL_COMMAND,
      title: "Insert EPICS record tail",
      arguments: [{ recordType }],
    };
    items.push(item);
  }

  return items;
}

function buildDbdDeviceRecordTypeItems(snapshot, context) {
  const items = [];

  for (const recordType of [...snapshot.recordTypes].sort(compareLabels)) {
    if (!matchesCompletionQuery(recordType, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(
      recordType,
      vscode.CompletionItemKind.Class,
    );
    item.range = context.range;
    item.insertText = `${recordType}, `;
    item.detail = "EPICS record type";
    item.filterText = buildFilterText(recordType);
    item.sortText = buildSortText(recordType, context.partial);
    item.command = {
      command: TRIGGER_SUGGEST_COMMAND,
      title: "Trigger Suggest",
    };
    items.push(item);
  }

  return items;
}

function buildDbdDeviceKeywordItems(context) {
  if (!matchesCompletionQuery("device", context.partial)) {
    return [];
  }

  const item = new vscode.CompletionItem(
    "device",
    vscode.CompletionItemKind.Keyword,
  );
  item.range = context.range;
  item.insertText = "device";
  item.detail = "EPICS device support definition";
  item.filterText = buildFilterText("device");
  item.sortText = buildSortText("device", context.partial);
  item.preselect = true;
  item.command = {
    command: INSERT_DBD_DEVICE_TAIL_COMMAND,
    title: "Insert DBD device tail",
  };
  return [item];
}

function buildDbdDeviceLinkTypeItems(context) {
  const items = [];

  for (const linkType of DBD_DEVICE_LINK_TYPES) {
    if (!matchesCompletionQuery(linkType, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(
      linkType,
      vscode.CompletionItemKind.EnumMember,
    );
    item.range = context.range;
    item.insertText = `${linkType}, `;
    item.detail = "Device link type";
    item.documentation = new vscode.MarkdownString(
      "Modern EPICS device supports typically use `INST_IO` here.",
    );
    item.filterText = buildFilterText(linkType);
    item.sortText = buildSortText(linkType, context.partial);
    item.command = {
      command: TRIGGER_SUGGEST_COMMAND,
      title: "Trigger Suggest",
    };
    items.push(item);
  }

  return items;
}

function buildDbdDeviceSupportNameItems(snapshot, context) {
  const items = [];

  for (const [supportName, definitions] of [...snapshot.deviceSupportDefinitionsByName.entries()].sort(
    ([left], [right]) => compareLabels(left, right),
  )) {
    if (!matchesCompletionQuery(supportName, context.partial)) {
      continue;
    }

    const definition = definitions[0];
    const item = new vscode.CompletionItem(
      supportName,
      vscode.CompletionItemKind.Struct,
    );
    item.range = context.range;
    item.insertText = `${supportName}, "`;
    item.detail = definition
      ? `epicsExportAddress(${definition.exportType}, ${supportName})`
      : "Exported device support structure";
    if (definition) {
      item.documentation = new vscode.MarkdownString(
        `Defined in \`${definition.relativePath}:${definition.line}\``,
      );
    }
    item.filterText = buildFilterText(supportName);
    item.sortText = buildSortText(supportName, context.partial);
    item.command = {
      command: TRIGGER_SUGGEST_COMMAND,
      title: "Trigger Suggest",
    };
    items.push(item);
  }

  return items;
}

function buildDbdDeviceChoiceNameItems(context) {
  const suggestions = [];
  const inferredName = inferDeviceChoiceName(context.supportName);
  if (inferredName) {
    suggestions.push({
      label: inferredName,
      detail: "Suggested DTYP name from exported structure",
    });
  }

  suggestions.push({
    label: "device_name",
    detail: "Used as DTYP in database records",
  });

  const items = [];
  const seen = new Set();
  for (const suggestion of suggestions) {
    if (!suggestion.label || seen.has(suggestion.label)) {
      continue;
    }
    if (!matchesCompletionQuery(suggestion.label, context.partial)) {
      continue;
    }

    seen.add(suggestion.label);
    const item = new vscode.CompletionItem(
      suggestion.label,
      vscode.CompletionItemKind.Value,
    );
    item.range = context.range;
    item.insertText = `${suggestion.label}"`;
    item.detail = suggestion.detail;
    item.documentation = new vscode.MarkdownString(
      "This string becomes the `DTYP` value used in EPICS database records.",
    );
    item.filterText = buildFilterText(suggestion.label);
    item.sortText = buildSortText(suggestion.label, context.partial);
    if (suggestion.label === inferredName) {
      item.preselect = true;
    }
    items.push(item);
  }

  return items;
}

function buildDbdDriverKeywordItems(context) {
  if (!matchesCompletionQuery("driver", context.partial)) {
    return [];
  }

  const item = new vscode.CompletionItem(
    "driver",
    vscode.CompletionItemKind.Keyword,
  );
  item.range = context.range;
  item.insertText = "driver";
  item.detail = "EPICS driver support definition";
  item.filterText = buildFilterText("driver");
  item.sortText = buildSortText("driver", context.partial);
  item.preselect = true;
  item.command = {
    command: INSERT_DBD_DRIVER_TAIL_COMMAND,
    title: "Insert DBD driver tail",
  };
  return [item];
}

function buildDbdDriverNameItems(snapshot, context) {
  const items = [];

  for (const [driverName, definitions] of [...snapshot.driverDefinitionsByName.entries()].sort(
    ([left], [right]) => compareLabels(left, right),
  )) {
    if (!matchesCompletionQuery(driverName, context.partial)) {
      continue;
    }

    const definition = definitions[0];
    const item = new vscode.CompletionItem(
      driverName,
      vscode.CompletionItemKind.Struct,
    );
    item.range = context.range;
    item.insertText = driverName;
    item.detail = definition
      ? `epicsExportAddress(${definition.exportType}, ${driverName})`
      : "Exported driver support structure";
    if (definition) {
      item.documentation = new vscode.MarkdownString(
        `Defined in \`${definition.relativePath}:${definition.line}\``,
      );
    }
    item.filterText = buildFilterText(driverName);
    item.sortText = buildSortText(driverName, context.partial);
    items.push(item);
  }

  return items;
}

function buildDbdRegistrarKeywordItems(context) {
  if (!matchesCompletionQuery("registrar", context.partial)) {
    return [];
  }

  const item = new vscode.CompletionItem(
    "registrar",
    vscode.CompletionItemKind.Keyword,
  );
  item.range = context.range;
  item.insertText = "registrar";
  item.detail = "EPICS registrar definition";
  item.filterText = buildFilterText("registrar");
  item.sortText = buildSortText("registrar", context.partial);
  item.preselect = true;
  item.command = {
    command: INSERT_DBD_REGISTRAR_TAIL_COMMAND,
    title: "Insert DBD registrar tail",
  };
  return [item];
}

function buildDbdRegistrarNameItems(snapshot, document, context) {
  const items = [];
  const documentDirectory = getDocumentDirectoryPath(document);

  for (const [registrarName, definitions] of [...snapshot.registrarDefinitionsByName.entries()].sort(
    ([left], [right]) => compareLabels(left, right),
  )) {
    const localDefinitions = definitions.filter((definition) =>
      isDefinitionInDirectory(definition, documentDirectory),
    );
    if (localDefinitions.length === 0) {
      continue;
    }

    if (!matchesCompletionQuery(registrarName, context.partial)) {
      continue;
    }

    const definition = localDefinitions[0];
    const item = new vscode.CompletionItem(
      registrarName,
      vscode.CompletionItemKind.Function,
    );
    item.range = context.range;
    item.insertText = registrarName;
    item.detail = `epicsExportRegistrar(${registrarName})`;
    item.documentation = new vscode.MarkdownString(
      `Defined in \`${definition.relativePath}:${definition.line}\``,
    );
    item.filterText = buildFilterText(registrarName);
    item.sortText = buildSortText(registrarName, context.partial);
    items.push(item);
  }

  return items;
}

function buildDbdFunctionKeywordItems(context) {
  if (!matchesCompletionQuery("function", context.partial)) {
    return [];
  }

  const item = new vscode.CompletionItem(
    "function",
    vscode.CompletionItemKind.Keyword,
  );
  item.range = context.range;
  item.insertText = "function";
  item.detail = "EPICS function definition";
  item.filterText = buildFilterText("function");
  item.sortText = buildSortText("function", context.partial);
  item.preselect = true;
  item.command = {
    command: INSERT_DBD_FUNCTION_TAIL_COMMAND,
    title: "Insert DBD function tail",
  };
  return [item];
}

function buildDbdFunctionNameItems(snapshot, document, context) {
  const items = [];
  const documentDirectory = getDocumentDirectoryPath(document);

  for (const [functionName, definitions] of [...snapshot.functionDefinitionsByName.entries()].sort(
    ([left], [right]) => compareLabels(left, right),
  )) {
    const localDefinitions = definitions.filter((definition) =>
      isDefinitionInDirectory(definition, documentDirectory),
    );
    if (localDefinitions.length === 0) {
      continue;
    }

    if (!matchesCompletionQuery(functionName, context.partial)) {
      continue;
    }

    const definition = localDefinitions[0];
    const item = new vscode.CompletionItem(
      functionName,
      vscode.CompletionItemKind.Function,
    );
    item.range = context.range;
    item.insertText = functionName;
    item.detail = `epicsRegisterFunction(${functionName})`;
    item.documentation = new vscode.MarkdownString(
      `Defined in \`${definition.relativePath}:${definition.line}\``,
    );
    item.filterText = buildFilterText(functionName);
    item.sortText = buildSortText(functionName, context.partial);
    items.push(item);
  }

  return items;
}

function buildDbdVariableKeywordItems(context) {
  if (!matchesCompletionQuery("variable", context.partial)) {
    return [];
  }

  const item = new vscode.CompletionItem(
    "variable",
    vscode.CompletionItemKind.Keyword,
  );
  item.range = context.range;
  item.insertText = "variable";
  item.detail = "EPICS variable definition";
  item.filterText = buildFilterText("variable");
  item.sortText = buildSortText("variable", context.partial);
  item.preselect = true;
  item.command = {
    command: INSERT_DBD_VARIABLE_TAIL_COMMAND,
    title: "Insert DBD variable tail",
  };
  return [item];
}

function buildDbdVariableNameItems(snapshot, document, context) {
  const items = [];
  const documentDirectory = getDocumentDirectoryPath(document);

  for (const [variableName, definitions] of [...snapshot.variableDefinitionsByName.entries()].sort(
    ([left], [right]) => compareLabels(left, right),
  )) {
    const localDefinitions = definitions.filter((definition) =>
      isDefinitionInDirectory(definition, documentDirectory),
    );
    if (localDefinitions.length === 0) {
      continue;
    }

    if (!matchesCompletionQuery(variableName, context.partial)) {
      continue;
    }

    const definition = localDefinitions[0];
    const item = new vscode.CompletionItem(
      variableName,
      vscode.CompletionItemKind.Variable,
    );
    item.range = context.range;
    item.insertText = `${variableName}, ${definition.exportType}`;
    item.detail = `epicsExportAddress(${definition.exportType}, ${variableName})`;
    item.documentation = new vscode.MarkdownString(
      `Defined in \`${definition.relativePath}:${definition.line}\``,
    );
    item.filterText = buildFilterText(variableName);
    item.sortText = buildSortText(variableName, context.partial);
    items.push(item);
  }

  return items;
}

function buildFieldNameItems(snapshot, document, position, context) {
  const fieldNames = getAvailableFieldNamesForRecordInstance(
    snapshot,
    document,
    position,
    context.recordType,
  );
  const items = [];

  for (const fieldName of fieldNames) {
    if (!matchesCompletionQuery(fieldName, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(
      fieldName,
      vscode.CompletionItemKind.Field,
    );
    item.range = context.range;
    item.insertText = fieldName;
    item.detail = buildFieldCompletionDetail(snapshot, context.recordType, fieldName);
    item.filterText = buildFilterText(fieldName);
    item.sortText = buildSortText(fieldName, context.partial);
    item.command = {
      command: INSERT_FIELD_TAIL_COMMAND,
      title: "Insert EPICS field tail",
      arguments: [
        {
          recordType: context.recordType,
          fieldName,
        },
      ],
    };
    items.push(item);
  }

  return items;
}

function buildFieldValueItems(snapshot, context) {
  if (isLinkField(context.fieldName)) {
    const linkItems = buildLinkTargetItems(snapshot, context);
    if (linkItems.length > 0) {
      return linkItems;
    }
  }

  const items = [];
  const valueLabels = getFieldValueLabels(snapshot, context);

  for (const label of valueLabels) {
    const item = new vscode.CompletionItem(
      label,
      vscode.CompletionItemKind.EnumMember,
    );
    item.range = context.range;
    item.detail = `${context.fieldName} value`;
    item.filterText = buildFilterText(label);
    items.push(item);
  }

  items.sort((left, right) => compareLabels(left.label, right.label));
  return items;
}

function buildFilePathItems(snapshot, document, context) {
  const items = [];
  const labels = new Set();
  const allowedExtensions = getAllowedExtensionsForFileContext(context.fileKind);
  const startupState =
    isStartupDocument(document)
      ? createStartupExecutionState(
          snapshot,
          document,
          document.offsetAt(context.range.end),
        )
      : undefined;
  const startupBaseDirectory = startupState?.currentDirectory;

  if (context.fileKind === "dbLoadRecords" && startupState?.currentDirectory) {
    return getDbLoadRecordsFilesystemPathItems(startupState, context);
  }

  for (const filesystemEntry of getFilesystemPathEntries(startupState, context)) {
    const item = new vscode.CompletionItem(
      filesystemEntry.insertPath,
      filesystemEntry.kind,
    );
    item.range = context.range;
    item.detail = filesystemEntry.detail;
    item.documentation = filesystemEntry.documentation;
    item.filterText = buildFilterText(filesystemEntry.insertPath);
    item.sortText = `0-${filesystemEntry.insertPath}`;
    applyDbLoadRecordsPathCompletion(item, context, filesystemEntry.absolutePath);
    items.push(item);
    labels.add(filesystemEntry.insertPath);
  }

  for (const projectEntry of getProjectFilePathEntries(
    snapshot,
    document,
    context.fileKind,
    startupBaseDirectory,
  )) {
    if (
      context.partial &&
      !matchesCompletionQuery(projectEntry.insertPath, context.partial)
    ) {
      continue;
    }

    const item = new vscode.CompletionItem(
      projectEntry.insertPath,
      vscode.CompletionItemKind.File,
    );
    item.range = context.range;
    item.detail = projectEntry.detail;
    item.documentation = projectEntry.documentation;
    item.filterText = buildFilterText(projectEntry.insertPath);
    item.sortText = `0-${projectEntry.insertPath}`;
    applyDbLoadRecordsPathCompletion(item, context, projectEntry.absolutePath);
    items.push(item);
    labels.add(projectEntry.insertPath);
  }

  for (const entry of snapshot.workspaceFiles) {
    if (!allowedExtensions.has(entry.extension)) {
      continue;
    }

    const insertPath = startupBaseDirectory
      ? getRelativePathFromBaseDirectory(startupBaseDirectory, entry.uri.fsPath)
      : getInsertPathForDocument(document, entry);
    if (labels.has(insertPath)) {
      continue;
    }

    if (
      context.partial &&
      !matchesCompletionQuery(insertPath, context.partial)
    ) {
      continue;
    }

    const item = new vscode.CompletionItem(
      insertPath,
      vscode.CompletionItemKind.File,
    );
    item.range = context.range;
    item.detail = entry.relativePath;
    applyDbLoadRecordsPathCompletion(item, context, entry.absolutePath);
    items.push(item);
  }

  items.sort((left, right) => compareLabels(left.label, right.label));
  return items;
}

function buildStartupLoadMacroItems(snapshot, document, position, context) {
  const resolvedMacroData = resolveStartupLoadMacroData(
    snapshot,
    document,
    position,
    context.path,
  );
  if (!resolvedMacroData) {
    return [];
  }

  const { resolvedFile, macroNames } = resolvedMacroData;
  if (macroNames.length === 0) {
    return [];
  }

  const completedAssignments = extractNamedAssignments(context.completedText);
  const remainingMacroNames = macroNames.filter(
    (macroName) => !completedAssignments.has(macroName),
  );
  if (remainingMacroNames.length === 0) {
    return [];
  }

  if (completedAssignments.size === 0 && !context.partial.trim()) {
    const assignmentLabel = buildDbLoadRecordsMacroAssignmentsLabel(remainingMacroNames);
    const item = new vscode.CompletionItem(
      assignmentLabel,
      vscode.CompletionItemKind.Value,
    );
    item.range = getDbLoadRecordsMacroValueInsertionRange(document, context.range);
    item.insertText = new vscode.SnippetString(
      `${buildDbLoadRecordsMacroAssignmentsSnippet(remainingMacroNames)}")${buildSnippetPlaceholder(0, "")}`,
    );
    item.detail = `Macros used by ${path.posix.basename(normalizePath(context.path))}`;
    item.documentation = new vscode.MarkdownString(
      `Loaded from \`${resolvedFile.absolutePath}\``,
    );
    item.filterText = buildFilterText(assignmentLabel);
    item.sortText = buildSortText(assignmentLabel, context.partial);
    item.preselect = true;
    return [item];
  }

  const items = [];
  for (const macroName of remainingMacroNames) {
    const assignmentLabel = `${macroName}=`;
    if (!matchesCompletionQuery(assignmentLabel, context.partial)) {
      continue;
    }

    const isLastRemaining = remainingMacroNames.length === 1;
    const item = new vscode.CompletionItem(
      assignmentLabel,
      vscode.CompletionItemKind.Variable,
    );
    item.range = isLastRemaining
      ? getDbLoadRecordsMacroValueInsertionRange(document, context.range)
      : context.range;
    item.insertText = new vscode.SnippetString(
      isLastRemaining
        ? `${macroName}=${buildSnippetPlaceholder(1, "")}")${buildSnippetPlaceholder(0, "")}`
        : `${macroName}=${buildSnippetPlaceholder(1, "")},`,
    );
    item.detail = `Remaining macro for ${path.posix.basename(normalizePath(context.path))}`;
    item.documentation = new vscode.MarkdownString(
      `Loaded from \`${resolvedFile.absolutePath}\``,
    );
    item.filterText = buildFilterText(assignmentLabel);
    item.sortText = buildSortText(assignmentLabel, context.partial);
    item.preselect = true;
    if (!isLastRemaining) {
      item.command = {
        command: TRIGGER_SUGGEST_COMMAND,
        title: "Trigger Suggest",
      };
    }
    items.push(item);
  }

  items.sort((left, right) => compareLabels(left.label, right.label));
  return items;
}

function buildStartupLoadMacroTailItems(snapshot, document, position, context) {
  const resolvedMacroData = resolveStartupLoadMacroData(
    snapshot,
    document,
    position,
    context.path,
  );
  if (!resolvedMacroData) {
    return [];
  }

  const { resolvedFile, macroNames } = resolvedMacroData;
  if (macroNames.length === 0) {
    return [];
  }

  const assignmentLabel = buildDbLoadRecordsMacroAssignmentsLabel(macroNames);
  const item = new vscode.CompletionItem(
    assignmentLabel,
    vscode.CompletionItemKind.Value,
  );
  item.range = getDbLoadRecordsTailInsertionRange(document, context.range.start);
  item.insertText = new vscode.SnippetString(
    buildDbLoadRecordsCompletionTailSnippet(macroNames),
  );
  item.detail = `Macros used by ${path.posix.basename(normalizePath(context.path))}`;
  item.documentation = new vscode.MarkdownString(
    `Loaded from \`${resolvedFile.absolutePath}\``,
  );
  item.filterText = buildFilterText(assignmentLabel);
  item.sortText = `0-${assignmentLabel}`;
  item.preselect = true;
  return [item];
}

function resolveStartupLoadMacroData(snapshot, document, position, filePath) {
  const state = createStartupExecutionState(
    snapshot,
    document,
    document.offsetAt(position),
  );
  const resolution = resolveStartupPath(
    snapshot,
    document,
    {
      command: "dbLoadRecords",
      path: filePath,
    },
    state,
  );
  const resolvedFile = resolution
    ? getReadableStartupFileResolution(document, resolution)
    : undefined;
  if (!resolvedFile?.text) {
    return undefined;
  }

  return {
    resolvedFile,
    macroNames: extractMacroNames(maskDatabaseComments(resolvedFile.text)),
  };
}

function applyDbLoadRecordsPathCompletion(item, context, absolutePath) {
  if (
    !item ||
    context.fileKind !== "dbLoadRecords" ||
    item.kind !== vscode.CompletionItemKind.File
  ) {
    return;
  }

  item.command = {
    command: INSERT_DBLOAD_RECORDS_MACRO_TAIL_COMMAND,
    title: "Insert dbLoadRecords macro tail",
    arguments: [{ absolutePath }],
  };
}

function buildMakefileReferenceItems(snapshot, document, context) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const resourceMap =
    context.type === "makefileDbd"
      ? project?.availableDbds || snapshot.projectModel.availableDbds
      : project?.availableLibs || snapshot.projectModel.availableLibs;
  const kind =
    context.type === "makefileDbd"
      ? vscode.CompletionItemKind.File
      : vscode.CompletionItemKind.Module;
  const items = [];

  for (const entry of resourceMap.values()) {
    if (context.partial && !matchesCompletionQuery(entry.name, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(entry.name, kind);
    item.range = context.range;
    item.detail = entry.detail;
    item.documentation = entry.documentation;
    item.filterText = buildFilterText(entry.name);
    item.sortText = buildSortText(entry.name, context.partial);
    items.push(item);
  }

  items.sort((left, right) => compareLabels(left.label, right.label));
  return items;
}

function getNavigationTarget(snapshot, document, position) {
  if (isDatabaseDocument(document)) {
    return getDatabaseNavigationTarget(snapshot, document, position);
  }

  if (isStartupDocument(document)) {
    return getStartupNavigationTarget(snapshot, document, position);
  }

  if (isSubstitutionsDocument(document)) {
    return getSubstitutionNavigationTarget(snapshot, document, position);
  }

  if (isMakefileDocument(document)) {
    return getMakefileNavigationTarget(snapshot, document, position);
  }

  return undefined;
}

function getDatabaseNavigationTarget(snapshot, document, position) {
  if (!isDatabaseDocument(document) || document.uri.scheme !== "file") {
    return undefined;
  }

  const text = document.getText();
  const offset = document.offsetAt(position);
  for (const entry of extractDatabaseTocEntries(text)) {
    if (offset < entry.linkStart || offset > entry.linkEnd) {
      continue;
    }

    const declaration = findRecordDeclarationByTypeAndName(
      text,
      entry.recordType,
      entry.recordName,
    );
    if (!declaration) {
      return undefined;
    }

    const recordPosition = document.positionAt(declaration.recordStart);
    return {
      absolutePath: normalizeFsPath(document.uri.fsPath),
      line: recordPosition.line + 1,
      character: recordPosition.character + 1,
    };
  }

  const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (!fieldDeclaration || !LINK_DBF_TYPES.has(fieldDeclaration.dbfType)) {
    return undefined;
  }

  const recordName = resolveLinkedRecordName(
    snapshot,
    document,
    fieldDeclaration.value,
  );
  if (!recordName) {
    return undefined;
  }

  const definitions = getRecordDefinitionsForName(snapshot, document, recordName);
  if (definitions.length === 0) {
    return undefined;
  }

  const definition = definitions[0];
  return {
    absolutePath: definition.absolutePath,
    line: definition.line,
    character: 1,
  };
}

function getHover(snapshot, document, position) {
  if (isDatabaseDocument(document)) {
    const databaseTocHover = getDatabaseTocEntryHover(document, position);
    if (databaseTocHover) {
      return databaseTocHover;
    }

    const deviceTypeFieldHover = getDeviceTypeFieldHover(
      snapshot,
      document,
      position,
    );
    if (deviceTypeFieldHover) {
      return deviceTypeFieldHover;
    }

    const menuFieldHover = getMenuFieldHover(snapshot, document, position);
    if (menuFieldHover) {
      return menuFieldHover;
    }

    const streamProtocolHover = getStreamProtocolFileHover(
      snapshot,
      document,
      position,
    );
    if (streamProtocolHover) {
      return streamProtocolHover;
    }

    const linkedRecordHover = getLinkedRecordHover(snapshot, document, position);
    if (linkedRecordHover) {
      return linkedRecordHover;
    }
  }

  if (isStartupDocument(document)) {
    const startupRecordHover = getStartupDbpfRecordHover(snapshot, document, position);
    if (startupRecordHover) {
      return startupRecordHover;
    }

    const startupLoadFileHover = getStartupLoadFileHover(
      snapshot,
      document,
      position,
    );
    if (startupLoadFileHover) {
      return startupLoadFileHover;
    }
  }

  if (isSubstitutionsDocument(document)) {
    const substitutionTemplateHover = getSubstitutionTemplateHover(
      snapshot,
      document,
      position,
    );
    if (substitutionTemplateHover) {
      return substitutionTemplateHover;
    }
  }

  if (isPvlistDocument(document)) {
    const pvlistRecordHover = getPvlistRecordHover(snapshot, document, position);
    if (pvlistRecordHover) {
      return pvlistRecordHover;
    }
  }

  if (isMakefileDocument(document)) {
    const makefileDatabaseHover = getMakefileDatabaseFileHover(
      snapshot,
      document,
      position,
    );
    if (makefileDatabaseHover) {
      return makefileDatabaseHover;
    }
  }

  if (isSequencerDocument(document)) {
    const sequencerHover = getSequencerSymbolHover(document, position);
    if (sequencerHover) {
      return sequencerHover;
    }
  }

  return getVariableHover(snapshot, document, position);
}

function buildDatabaseSemanticTokens(snapshot, document, legend) {
  const builder = new vscode.SemanticTokensBuilder(legend);
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    pushDatabaseSemanticToken(
      builder,
      document,
      recordDeclaration.recordTypeStart,
      recordDeclaration.recordTypeEnd,
      "type",
    );
    pushDatabaseSemanticToken(
      builder,
      document,
      recordDeclaration.nameStart,
      recordDeclaration.nameEnd,
      "variable",
      ["declaration"],
    );

    const fieldDeclarations = extractFieldDeclarationsInRecord(text, recordDeclaration);
    const fieldTypes =
      snapshot.fieldTypesByRecordType.get(recordDeclaration.recordType) || new Map();
    const isStreamRecord = fieldDeclarations.some(
      (fieldDeclaration) =>
        fieldDeclaration.fieldName === "DTYP" &&
        fieldDeclaration.value.trim().toLowerCase() === "stream",
    );

    for (const fieldDeclaration of fieldDeclarations) {
      pushDatabaseSemanticToken(
        builder,
        document,
        fieldDeclaration.fieldNameStart,
        fieldDeclaration.fieldNameEnd,
        "property",
      );

      const dbfType = fieldTypes.get(fieldDeclaration.fieldName);
      if (!dbfType || fieldDeclaration.valueStart >= fieldDeclaration.valueEnd) {
        continue;
      }

      if (LINK_DBF_TYPES.has(dbfType)) {
        pushDatabaseLinkValueSemanticTokens(
          builder,
          snapshot,
          document,
          fieldDeclaration,
          isStreamRecord,
        );
        continue;
      }

      if (dbfType === "DBF_MENU") {
        pushDatabaseSemanticToken(
          builder,
          document,
          fieldDeclaration.valueStart,
          fieldDeclaration.valueEnd,
          "enumMember",
        );
        continue;
      }

      if (NUMERIC_DBF_TYPES.has(dbfType)) {
        pushDatabaseSemanticToken(
          builder,
          document,
          fieldDeclaration.valueStart,
          fieldDeclaration.valueEnd,
          "number",
        );
        continue;
      }

      pushDatabaseSemanticToken(
        builder,
        document,
        fieldDeclaration.valueStart,
        fieldDeclaration.valueEnd,
        "string",
      );
    }
  }

  return builder.build();
}

function pushDatabaseLinkValueSemanticTokens(
  builder,
  snapshot,
  document,
  fieldDeclaration,
  isStreamRecord,
) {
  const tokenRanges = [];

  if (isStreamRecord) {
    const protocolRange = getStreamProtocolReferenceRange(fieldDeclaration);
    if (protocolRange) {
      tokenRanges.push({ ...protocolRange, type: "string", modifiers: [] });
    }
  }

  const linkedRecordRange = getLinkedRecordSemanticTokenRange(
    snapshot,
    document,
    fieldDeclaration,
  );
  if (linkedRecordRange) {
    tokenRanges.push({ ...linkedRecordRange, type: "variable", modifiers: [] });
  }

  for (const macroRange of extractMacroOffsets(
    fieldDeclaration.value,
    fieldDeclaration.valueStart,
  )) {
    tokenRanges.push({
      start: macroRange.start,
      end: macroRange.end,
      type: "macro",
      modifiers: [],
    });
  }

  tokenRanges.sort(
    (left, right) => left.start - right.start || left.end - right.end,
  );
  let lastEnd = -1;
  for (const tokenRange of tokenRanges) {
    if (tokenRange.start < lastEnd) {
      continue;
    }

    pushDatabaseSemanticToken(
      builder,
      document,
      tokenRange.start,
      tokenRange.end,
      tokenRange.type,
      tokenRange.modifiers,
    );
    lastEnd = tokenRange.end;
  }
}

function getLinkedRecordSemanticTokenRange(snapshot, document, fieldDeclaration) {
  const fieldValue = String(fieldDeclaration.value || "");
  const trimmedValue = fieldValue.trim();
  if (!trimmedValue || trimmedValue.startsWith("@")) {
    return undefined;
  }

  const leadingWhitespaceLength = fieldValue.match(/^\s*/)[0].length;
  const firstTokenMatch = fieldValue.slice(leadingWhitespaceLength).match(/^[^\s]+/);
  if (!firstTokenMatch) {
    return undefined;
  }

  const firstToken = firstTokenMatch[0];
  let tokenText = firstToken;
  const lastDotIndex = firstToken.lastIndexOf(".");
  if (lastDotIndex > 0) {
    const suffix = firstToken.slice(lastDotIndex + 1);
    if (/^[A-Z0-9_]+$/.test(suffix)) {
      tokenText = firstToken.slice(0, lastDotIndex);
    }
  }

  const normalizedRecordName = normalizeLinkedRecordCandidate(tokenText);
  if (
    !normalizedRecordName ||
    getRecordDefinitionsForName(snapshot, document, normalizedRecordName).length === 0
  ) {
    return undefined;
  }

  const tokenStart = fieldDeclaration.valueStart + leadingWhitespaceLength;
  return {
    start: tokenStart,
    end: tokenStart + tokenText.length,
  };
}

function getStreamProtocolReferenceRange(fieldDeclaration) {
  const match = String(fieldDeclaration.value || "").match(/^\s*@([^\s"'`]+)/);
  if (!match || containsEpicsMacroReference(match[1])) {
    return undefined;
  }

  const protocolPath = match[1];
  const protocolStart =
    fieldDeclaration.valueStart + match[0].indexOf(protocolPath);
  return {
    start: protocolStart,
    end: protocolStart + protocolPath.length,
  };
}

function pushDatabaseSemanticToken(
  builder,
  document,
  startOffset,
  endOffset,
  tokenType,
  tokenModifiers = [],
) {
  if (
    typeof startOffset !== "number" ||
    typeof endOffset !== "number" ||
    endOffset <= startOffset
  ) {
    return;
  }

  const start = document.positionAt(startOffset);
  const end = document.positionAt(endOffset);
  if (start.line !== end.line) {
    return;
  }

  builder.push(new vscode.Range(start, end), tokenType, tokenModifiers);
}

function getDatabaseTocEntryHover(document, position) {
  if (!isDatabaseDocument(document)) {
    return undefined;
  }

  const offset = document.offsetAt(position);
  const text = document.getText();
  for (const entry of extractDatabaseTocEntries(text)) {
    if (offset < entry.hoverStart || offset > entry.hoverEnd) {
      continue;
    }

    const declaration = findRecordDeclarationByTypeAndName(
      text,
      entry.recordType,
      entry.recordName,
    );
    if (!declaration) {
      return undefined;
    }

    const markdown = new vscode.MarkdownString();
    markdown.isTrusted = true;
    const line = getLineNumberAtOffset(text, declaration.recordStart);
    markdown.appendMarkdown(
      `**Record:** \`${escapeInlineCode(declaration.name)}\``,
    );
    markdown.appendMarkdown(
      `\n\nType: \`${escapeInlineCode(declaration.recordType)}\``,
    );
    markdown.appendMarkdown(
      `\n\nLocation: ${createRecordLocationLink(
        {
          absolutePath: normalizeFsPath(document.uri.fsPath),
          line,
        },
        `line ${line}`,
      )}`,
    );
    markdown.appendMarkdown("\n\nPreview:");
    markdown.appendCodeblock(buildRecordPreview(text, declaration), "db");

    return new vscode.Hover(
      markdown,
      new vscode.Range(
        document.positionAt(entry.hoverStart),
        document.positionAt(entry.hoverEnd),
      ),
    );
  }

  return undefined;
}

function getDeviceTypeFieldHover(snapshot, document, position) {
  if (!isDatabaseDocument(document) || document.uri.scheme !== "file") {
    return undefined;
  }

  const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (!fieldDeclaration || fieldDeclaration.fieldName !== "DTYP") {
    return undefined;
  }

  const matches = resolveDeviceTypeHoverMatches(
    snapshot,
    document,
    fieldDeclaration,
  );
  if (matches.length === 0) {
    return undefined;
  }

  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(
    `**Device support:** \`${escapeInlineCode(fieldDeclaration.value.trim())}\``,
  );
  markdown.appendMarkdown(
    `\n\nRecord type: \`${escapeInlineCode(fieldDeclaration.recordType)}\``,
  );
  markdown.appendMarkdown("\n\nMatched `device(...)` declarations:");

  const displayedMatches = matches.slice(0, 5);
  for (const match of displayedMatches) {
    const locationLabel = `${match.relativePath || match.absolutePath}:${match.line}`;
    markdown.appendMarkdown("\n\n---\n\n");
    markdown.appendCodeblock(match.declarationText, "dbd");
    markdown.appendMarkdown(`\n\nDBD: ${createFileLocationLink(
      match.absolutePath,
      match.line,
      locationLabel,
    )}`);
    markdown.appendMarkdown(
      `\n\nSupport: \`${escapeInlineCode(match.supportName)}\``,
    );
    markdown.appendMarkdown(
      `\n\nLink type: \`${escapeInlineCode(match.linkType)}\``,
    );
    markdown.appendMarkdown(
      `\n\nSearch root: \`${escapeInlineCode(match.searchLabel)}\``,
    );
  }

  if (matches.length > displayedMatches.length) {
    markdown.appendMarkdown(
      `\n\n${matches.length - displayedMatches.length} more matching device declarations omitted.`,
    );
  }

  return new vscode.Hover(markdown, fieldDeclaration.range);
}

function resolveDeviceTypeHoverMatches(snapshot, document, fieldDeclaration) {
  const recordType = String(fieldDeclaration.recordType || "").trim();
  const choiceName = String(fieldDeclaration.value || "").trim();
  if (!recordType || !choiceName) {
    return [];
  }

  const matches = [];
  for (const candidate of collectDeviceTypeHoverDbdFiles(snapshot, document)) {
    for (const declaration of getDbdDeviceDeclarationsForFile(candidate.absolutePath)) {
      if (
        declaration.recordType !== recordType ||
        declaration.choiceName !== choiceName
      ) {
        continue;
      }

      matches.push({
        ...declaration,
        absolutePath: candidate.absolutePath,
        relativePath: candidate.relativePath,
        searchLabel: candidate.searchLabel,
      });
    }
  }

  return matches;
}

function collectDeviceTypeHoverDbdFiles(snapshot, document) {
  const entries = [];
  const seen = new Set();
  const documentDirectory = normalizeFsPath(path.dirname(document.uri.fsPath));
  addDeviceTypeHoverDbdDirectoryEntries(
    entries,
    seen,
    documentDirectory,
    "current folder",
  );

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (!project?.rootPath) {
    return entries;
  }

  addDeviceTypeHoverDbdDirectoryEntries(
    entries,
    seen,
    path.join(project.rootPath, "dbd"),
    "project dbd",
  );

  for (const releaseRoot of resolveReleaseModuleRoots(
    project.rootPath,
    project.releaseVariables,
  )) {
    addDeviceTypeHoverDbdDirectoryEntries(
      entries,
      seen,
      path.join(releaseRoot.rootPath, "dbd"),
      `${releaseRoot.variableName} dbd`,
    );
  }

  return entries;
}

function addDeviceTypeHoverDbdDirectoryEntries(entries, seen, directoryPath, searchLabel) {
  const normalizedDirectoryPath = normalizeFsPath(directoryPath);
  if (!normalizedDirectoryPath || !fs.existsSync(normalizedDirectoryPath)) {
    return;
  }

  let directoryEntries;
  try {
    directoryEntries = fs.readdirSync(normalizedDirectoryPath, { withFileTypes: true });
  } catch (error) {
    return;
  }

  for (const entry of directoryEntries) {
    if (
      !entry.isFile() ||
      path.extname(entry.name).toLowerCase() !== ".dbd"
    ) {
      continue;
    }

    const absolutePath = normalizeFsPath(
      path.join(normalizedDirectoryPath, entry.name),
    );
    if (seen.has(absolutePath)) {
      continue;
    }

    seen.add(absolutePath);
    entries.push({
      absolutePath,
      relativePath: normalizePath(
        vscode.workspace.asRelativePath(vscode.Uri.file(absolutePath), false),
      ),
      searchLabel,
    });
  }
}

function getDbdDeviceDeclarationsForFile(absolutePath) {
  const normalizedPath = normalizeFsPath(absolutePath);
  if (!normalizedPath) {
    return [];
  }

  const openDocument = vscode.workspace.textDocuments.find(
    (document) =>
      document.uri.scheme === "file" &&
      normalizeFsPath(document.uri.fsPath) === normalizedPath,
  );

  let cacheTag;
  let text;
  if (openDocument) {
    cacheTag = `open:${openDocument.version}`;
    text = openDocument.getText();
  } else {
    let stats;
    try {
      stats = fs.statSync(normalizedPath);
    } catch (error) {
      return [];
    }

    if (!stats.isFile()) {
      return [];
    }

    cacheTag = `fs:${stats.size}:${stats.mtimeMs}`;
    text = readTextFile(normalizedPath);
  }

  const cached = DBD_DEVICE_DECLARATION_CACHE.get(normalizedPath);
  if (cached?.cacheTag === cacheTag) {
    return cached.declarations;
  }

  const declarations = extractDbdDeviceDeclarations(text || "").map((declaration) => ({
    ...declaration,
    absolutePath: normalizedPath,
    relativePath: normalizePath(
      vscode.workspace.asRelativePath(vscode.Uri.file(normalizedPath), false),
    ),
    line: getLineNumberAtOffset(text || "", declaration.start),
  }));

  DBD_DEVICE_DECLARATION_CACHE.set(normalizedPath, {
    cacheTag,
    declarations,
  });
  return declarations;
}

function getMenuFieldHover(snapshot, document, position) {
  if (!isDatabaseDocument(document) || document.uri.scheme !== "file") {
    return undefined;
  }

  const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (!fieldDeclaration || fieldDeclaration.dbfType !== "DBF_MENU") {
    return undefined;
  }

  const choices = getMenuFieldChoices(
    snapshot,
    fieldDeclaration.recordType,
    fieldDeclaration.fieldName,
  );
  if (choices.length === 0) {
    return undefined;
  }

  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(
    `**${escapeInlineCode(fieldDeclaration.fieldName)}** menu choices`,
  );
  markdown.appendMarkdown(
    `\n\nCurrent value: \`${escapeInlineCode(fieldDeclaration.value)}\``,
  );

  for (const choice of choices) {
    markdown.appendMarkdown("\n\n- ");
    if (choice === fieldDeclaration.value) {
      markdown.appendMarkdown(`\`${escapeInlineCode(choice)}\``);
      markdown.appendMarkdown(" (current)");
      continue;
    }

    markdown.appendMarkdown(
      createMenuFieldChoiceLink(document, fieldDeclaration, choice),
    );
  }

  return new vscode.Hover(markdown, fieldDeclaration.range);
}

function getLinkedRecordHover(snapshot, document, position) {
  const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (!fieldDeclaration || !LINK_DBF_TYPES.has(fieldDeclaration.dbfType)) {
    return undefined;
  }

  const match = resolveLinkedRecordDefinitionMatch(
    snapshot,
    document,
    fieldDeclaration.value,
  );
  if (!match) {
    return undefined;
  }

  return createLinkedRecordHover(
    match.recordName,
    fieldDeclaration.fieldName,
    match.definitions,
    fieldDeclaration.range,
  );
}

function getPvlistRecordHover(snapshot, document, position) {
  if (!isPvlistDocument(document)) {
    return undefined;
  }

  const line = document.lineAt(position.line);
  const trimmedLine = String(line.text || "").trim();
  if (
    !trimmedLine ||
    trimmedLine.startsWith("#") ||
    /^[A-Za-z_][A-Za-z0-9_]*\s*=/.test(trimmedLine) ||
    trimmedLine.includes("=") ||
    /\s/.test(trimmedLine)
  ) {
    return undefined;
  }

  const trimmedStart = line.firstNonWhitespaceCharacterIndex;
  const trimmedEnd = trimmedStart + trimmedLine.length;
  if (position.character < trimmedStart || position.character > trimmedEnd) {
    return undefined;
  }

  const expandedRecordName = resolvePvlistHoverRecordName(document.getText(), trimmedLine);
  if (!expandedRecordName) {
    return undefined;
  }

  const definitions = getRecordDefinitionsForName(snapshot, document, expandedRecordName);
  if (definitions.length === 0) {
    return undefined;
  }

  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(`**Record:** \`${escapeInlineCode(expandedRecordName)}\``);
  if (expandedRecordName !== trimmedLine) {
    markdown.appendMarkdown(
      `\n\nPV List entry: \`${escapeInlineCode(trimmedLine)}\``,
    );
  }

  const displayedDefinitions = definitions.slice(0, 3);
  for (const definition of displayedDefinitions) {
    const location = `${definition.relativePath || definition.absolutePath}:${definition.line}`;
    markdown.appendMarkdown("\n\n---\n\n");
    markdown.appendMarkdown(
      `Type: \`${escapeInlineCode(definition.recordType || "unknown")}\``,
    );
    if (definition.absolutePath) {
      markdown.appendMarkdown(
        `\n\nLocation: ${createRecordLocationLink(definition, location)}`,
      );
    } else {
      markdown.appendMarkdown(`\n\nLocation: \`${escapeInlineCode(location)}\``);
    }
    markdown.appendMarkdown("\n\nPreview:");
    markdown.appendCodeblock(
      definition.preview || `record(${definition.recordType}, "${definition.name}")`,
      "db",
    );
  }

  if (definitions.length > displayedDefinitions.length) {
    markdown.appendMarkdown(
      `\n\n${definitions.length - displayedDefinitions.length} more matching record definitions omitted.`,
    );
  }

  return new vscode.Hover(
    markdown,
    new vscode.Range(
      new vscode.Position(position.line, trimmedStart),
      new vscode.Position(position.line, trimmedEnd),
    ),
  );
}

function getStreamProtocolFileHover(snapshot, document, position) {
  const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (
    !fieldDeclaration ||
    !LINK_DBF_TYPES.has(fieldDeclaration.dbfType) ||
    !isStreamDeviceRecordField(document, position)
  ) {
    return undefined;
  }

  const protocolReference = getStreamProtocolReferenceAtPosition(
    fieldDeclaration,
    document,
    position,
  );
  if (!protocolReference) {
    return undefined;
  }

  const resolutions = resolveStreamProtocolFileReferences(
    snapshot,
    document,
    protocolReference.protocolPath,
  );
  if (resolutions.length === 0) {
    return undefined;
  }

  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(
    `**StreamDevice protocol:** \`${escapeInlineCode(protocolReference.protocolPath)}\``,
  );
  markdown.appendMarkdown(
    `\n\nReferenced by \`${escapeInlineCode(fieldDeclaration.fieldName)}\``,
  );

  for (const resolution of resolutions) {
    markdown.appendMarkdown("\n\n- ");
    markdown.appendMarkdown(
      createProtocolFileLink(
        resolution.absolutePath,
        resolution.protocolRelativePath || resolution.absolutePath,
      ),
    );
    markdown.appendMarkdown(
      ` from \`${escapeInlineCode(resolution.startupFileRelativePath)}\``,
    );
  }

  return new vscode.Hover(markdown, protocolReference.range);
}

function getStartupDbpfRecordHover(snapshot, document, position) {
  const argument = getStartupDbpfArgumentAtPosition(document, position);
  if (!argument) {
    return undefined;
  }

  const loadedDefinitionsByName = getStartupLoadedRecordDefinitionMap(
    snapshot,
    document,
    position,
  );
  const recordName = extractLinkedRecordCandidates(argument.value).find((candidate) =>
    loadedDefinitionsByName.has(candidate),
  );
  if (!recordName) {
    return undefined;
  }

  const definitions = loadedDefinitionsByName.get(recordName) || [];
  if (definitions.length === 0) {
    return undefined;
  }

  return createLinkedRecordHover(
    recordName,
    "dbpf",
    definitions,
    argument.range,
  );
}

function getStartupLoadFileHover(snapshot, document, position) {
  if (!isStartupDocument(document) || getVariableReferenceAtPosition(document, position)) {
    return undefined;
  }

  const statement = getStartupStatementAtPosition(document, position);
  if (!statement || statement.kind !== "load" || statement.command !== "dbLoadRecords") {
    return undefined;
  }

  const state = createStartupExecutionState(snapshot, document, statement.start);
  const resolution = resolveStartupPath(snapshot, document, statement, state);
  if (!resolution || resolution.isDirectory) {
    return undefined;
  }

  const resolvedFile = getReadableStartupFileResolution(document, resolution);
  if (!resolvedFile?.text) {
    return undefined;
  }

  return createStartupLoadFileHover(document, statement, resolvedFile);
}

function getSubstitutionTemplateHover(snapshot, document, position) {
  if (!isSubstitutionsDocument(document)) {
    return undefined;
  }

  const reference = getSubstitutionTemplateReferenceAtPosition(document, position);
  if (!reference) {
    return undefined;
  }

  const absolutePaths = resolveSubstitutionTemplateAbsolutePathsForDocument(
    snapshot,
    document,
    reference.templatePath,
  );
  if (absolutePaths.length === 0) {
    return undefined;
  }

  return createSubstitutionTemplateHover(document, reference, absolutePaths);
}

function getStartupDbpfArgumentAtPosition(document, position) {
  if (!isStartupDocument(document)) {
    return undefined;
  }

  const lineText = `${document.lineAt(position.line).text}\n`;
  const sanitizedLineText = maskDatabaseComments(lineText).slice(0, -1);
  const regex = /dbpf\(\s*"((?:[^"\\]|\\.)*)"/g;
  let match;

  while ((match = regex.exec(sanitizedLineText))) {
    const valueStart = match.index + match[0].length - match[1].length - 1;
    const valueEnd = valueStart + match[1].length;
    if (position.character < valueStart || position.character > valueEnd) {
      continue;
    }

    return {
      value: match[1],
      range: new vscode.Range(
        position.line,
        valueStart,
        position.line,
        valueEnd,
      ),
    };
  }

  return undefined;
}

function createStartupLoadFileHover(document, statement, resolvedFile) {
  const absolutePath = normalizeFsPath(resolvedFile.absolutePath);
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown("**EPICS database file**");
  appendDatabaseFileHoverSummary(markdown, absolutePath, resolvedFile.text);

  return new vscode.Hover(
    markdown,
    new vscode.Range(
      document.positionAt(statement.pathStart),
      document.positionAt(statement.pathEnd),
    ),
  );
}

function createSubstitutionTemplateHover(document, reference, absolutePaths) {
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;

  if (absolutePaths.length === 1) {
    const absolutePath = absolutePaths[0];
    const templateText = readTextFile(absolutePath);
    markdown.appendMarkdown("**EPICS database/template file**");

    if (templateText === undefined) {
      markdown.appendMarkdown(
        `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
      );
    } else {
      appendDatabaseFileHoverSummary(markdown, absolutePath, templateText);
    }
  } else {
    markdown.appendMarkdown("**EPICS database/template file candidates**");
    markdown.appendMarkdown(`\n\nMatches: \`${absolutePaths.length}\``);

    for (const [index, absolutePath] of absolutePaths.entries()) {
      const templateText = readTextFile(absolutePath);
      markdown.appendMarkdown(`\n\n---\n\n**Candidate ${index + 1}**`);
      if (templateText === undefined) {
        markdown.appendMarkdown(
          `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
        );
      } else {
        appendDatabaseFileHoverSummary(markdown, absolutePath, templateText);
      }
    }
  }

  return new vscode.Hover(
    markdown,
    new vscode.Range(
      document.positionAt(reference.start),
      document.positionAt(reference.end),
    ),
  );
}

function getEpicsReferenceLocations(snapshot, document, position, includeDeclaration) {
  const symbol = getEpicsSemanticSymbolAtPosition(snapshot, document, position);
  if (!symbol) {
    return [];
  }

  return collectEpicsSymbolOccurrences(
    snapshot,
    symbol,
    document,
    includeDeclaration,
  ).map((occurrence) => new vscode.Location(occurrence.uri, occurrence.range));
}

function getEpicsSemanticSymbolAtPosition(snapshot, document, position) {
  return (
    getRecordSymbolAtPosition(snapshot, document, position) ||
    getMacroSymbolAtPosition(document, position) ||
    getRecordTypeSymbolAtPosition(document, position) ||
    getFieldSymbolAtPosition(document, position) ||
    getDbdNamedSymbolAtPosition(document, position) ||
    getSourceNamedSymbolAtPosition(document, position)
  );
}

function getRecordSymbolAtPosition(snapshot, document, position) {
  if (isDatabaseDocument(document)) {
    return getDatabaseRecordSymbolAtPosition(snapshot, document, position);
  }

  if (isStartupDocument(document)) {
    return getStartupRecordSymbolAtPosition(snapshot, document, position);
  }

  if (isPvlistDocument(document)) {
    return getPvlistRecordSymbolAtPosition(document, position);
  }

  if (isProbeDocument(document)) {
    return getProbeRecordSymbolAtPosition(document, position);
  }

  return undefined;
}

function getRecordTypeSymbolAtPosition(document, position) {
  if (isDatabaseDocument(document)) {
    return getDatabaseRecordTypeSymbolAtPosition(document, position);
  }

  if (document?.languageId === LANGUAGE_IDS.dbd) {
    return getDbdRecordTypeSymbolAtPosition(document, position);
  }

  return undefined;
}

function getDatabaseRecordTypeSymbolAtPosition(document, position) {
  const offset = document.offsetAt(position);
  for (const declaration of extractRecordDeclarations(document.getText())) {
    if (offset < declaration.recordTypeStart || offset > declaration.recordTypeEnd) {
      continue;
    }

    return createEpicsSemanticSymbol(
      "recordType",
      declaration.recordType,
      document,
      declaration.recordTypeStart,
      declaration.recordTypeEnd,
    );
  }

  return undefined;
}

function getDbdRecordTypeSymbolAtPosition(document, position) {
  const offset = document.offsetAt(position);
  const text = document.getText();

  for (const declaration of extractDbdRecordTypeDeclarations(text)) {
    if (offset >= declaration.nameStart && offset <= declaration.nameEnd) {
      return createEpicsSemanticSymbol(
        "recordType",
        declaration.name,
        document,
        declaration.nameStart,
        declaration.nameEnd,
      );
    }
  }

  for (const entry of extractDbdDeviceDeclarations(text)) {
    if (offset < entry.recordTypeStart || offset > entry.recordTypeEnd) {
      continue;
    }

    return createEpicsSemanticSymbol(
      "recordType",
      entry.recordType,
      document,
      entry.recordTypeStart,
      entry.recordTypeEnd,
    );
  }

  return undefined;
}

function getFieldSymbolAtPosition(document, position) {
  if (isDatabaseDocument(document)) {
    return getDatabaseFieldSymbolAtPosition(document, position);
  }

  if (document?.languageId === LANGUAGE_IDS.dbd) {
    return getDbdFieldSymbolAtPosition(document, position);
  }

  return undefined;
}

function getDatabaseFieldSymbolAtPosition(document, position) {
  const offset = document.offsetAt(position);
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    if (offset < recordDeclaration.recordStart || offset > recordDeclaration.recordEnd) {
      continue;
    }

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(text, recordDeclaration)) {
      if (
        offset >= fieldDeclaration.fieldNameStart &&
        offset <= fieldDeclaration.fieldNameEnd
      ) {
        return createEpicsSemanticSymbol(
          "field",
          fieldDeclaration.fieldName,
          document,
          fieldDeclaration.fieldNameStart,
          fieldDeclaration.fieldNameEnd,
          undefined,
          { recordType: recordDeclaration.recordType },
        );
      }
    }
  }

  return undefined;
}

function getDbdFieldSymbolAtPosition(document, position) {
  const offset = document.offsetAt(position);
  const text = document.getText();

  for (const recordTypeDeclaration of extractDbdRecordTypeDeclarations(text)) {
    if (offset < recordTypeDeclaration.blockStart || offset > recordTypeDeclaration.blockEnd) {
      continue;
    }

    for (const fieldDeclaration of extractDbdFieldDeclarationsInRecordType(
      text,
      recordTypeDeclaration,
    )) {
      if (
        offset >= fieldDeclaration.fieldNameStart &&
        offset <= fieldDeclaration.fieldNameEnd
      ) {
        return createEpicsSemanticSymbol(
          "field",
          fieldDeclaration.fieldName,
          document,
          fieldDeclaration.fieldNameStart,
          fieldDeclaration.fieldNameEnd,
          undefined,
          { recordType: recordTypeDeclaration.name },
        );
      }
    }
  }

  return undefined;
}

function getDbdNamedSymbolAtPosition(document, position) {
  if (document?.languageId !== LANGUAGE_IDS.dbd) {
    return undefined;
  }

  const offset = document.offsetAt(position);
  const text = document.getText();

  for (const entry of extractDbdDeviceDeclarations(text)) {
    if (offset < entry.supportNameStart || offset > entry.supportNameEnd) {
      continue;
    }

    return createEpicsSemanticSymbol(
      "deviceSupport",
      entry.supportName,
      document,
      entry.supportNameStart,
      entry.supportNameEnd,
    );
  }

  for (const [kind, keyword] of [
    ["driver", "driver"],
    ["registrar", "registrar"],
    ["function", "function"],
    ["variable", "variable"],
  ]) {
    for (const entry of extractDbdNamedEntries(text, keyword)) {
      if (offset < entry.nameStart || offset > entry.nameEnd) {
        continue;
      }

      return createEpicsSemanticSymbol(
        kind,
        entry.name,
        document,
        entry.nameStart,
        entry.nameEnd,
      );
    }
  }

  return undefined;
}

function getSourceNamedSymbolAtPosition(document, position) {
  if (!isSourceDocument(document)) {
    return undefined;
  }

  const offset = document.offsetAt(position);
  const text = document.getText();

  for (const entry of extractSourceNamedSymbolOccurrences(text)) {
    if (offset < entry.nameStart || offset > entry.nameEnd) {
      continue;
    }

    return createEpicsSemanticSymbol(
      entry.kind,
      entry.name,
      document,
      entry.nameStart,
      entry.nameEnd,
    );
  }

  return undefined;
}

function getDatabaseRecordSymbolAtPosition(snapshot, document, position) {
  const offset = document.offsetAt(position);
  const text = document.getText();
  const macroAssignments = extractDatabaseTocMacroAssignments(text);

  for (const declaration of extractRecordDeclarations(text)) {
    if (offset >= declaration.nameStart && offset <= declaration.nameEnd) {
      return createEpicsSemanticSymbol(
        "record",
        declaration.name,
        document,
        declaration.nameStart,
        declaration.nameEnd,
        getDatabaseRecordSearchNames(text, declaration.name),
      );
    }

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(text, declaration)) {
      const candidateRanges = getLinkedRecordCandidateRanges(
        fieldDeclaration.value,
        fieldDeclaration.valueStart,
        {
          allowMacroReferences: true,
          macroAssignments,
        },
      ).filter((candidate) => offset >= candidate.start && offset <= candidate.end);
      if (candidateRanges.length === 0) {
        continue;
      }

      const searchNames = getShortestRecordCandidateNames(candidateRanges);
      const declarationMatch = [...searchNames]
        .map((recordName) => findDatabaseDeclarationByAnyRecordName(text, recordName))
        .find(Boolean);
      const hasDefinition = [...searchNames].some(
        (candidateName) =>
          getRecordDefinitionsForName(snapshot, document, candidateName).length > 0,
      );
      if (!hasDefinition) {
        continue;
      }

      return createEpicsSemanticSymbol(
        "record",
        declarationMatch?.declaration?.name || [...searchNames][0],
        document,
        candidateRanges.reduce((minimum, candidate) => Math.min(minimum, candidate.start), Number.POSITIVE_INFINITY),
        candidateRanges.reduce((maximum, candidate) => Math.max(maximum, candidate.end), 0),
        searchNames,
      );
    }
  }

  const fallbackFieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
    snapshot,
    document,
    position,
  );
  if (fallbackFieldDeclaration) {
    const fallbackSymbol = resolveDatabaseFieldRecordSymbol(
      snapshot,
      document,
      text,
      fallbackFieldDeclaration,
      macroAssignments,
    );
    if (fallbackSymbol) {
      return fallbackSymbol;
    }
  }

  return undefined;
}

function resolveDatabaseFieldRecordSymbol(
  snapshot,
  document,
  documentText,
  fieldDeclaration,
  macroAssignments,
) {
  const candidateRanges = getLinkedRecordCandidateRanges(
    fieldDeclaration.value,
    fieldDeclaration.valueStart,
    {
      allowMacroReferences: true,
      macroAssignments,
    },
  );
  const searchNames = getShortestRecordCandidateNames(candidateRanges);
  if (searchNames.size === 0) {
    return undefined;
  }

  const declarationMatch = [...searchNames]
    .map((recordName) => findDatabaseDeclarationByAnyRecordName(documentText, recordName))
    .find(Boolean);
  const hasDefinition = [...searchNames].some(
    (candidateName) =>
      getRecordDefinitionsForName(snapshot, document, candidateName).length > 0,
  );
  if (!hasDefinition) {
    return undefined;
  }

  const shortestRanges = candidateRanges.filter((candidate) => searchNames.has(candidate.name));
  return createEpicsSemanticSymbol(
    "record",
    declarationMatch?.declaration?.name || [...searchNames][0],
    document,
    shortestRanges.reduce((minimum, candidate) => Math.min(minimum, candidate.start), Number.POSITIVE_INFINITY),
    shortestRanges.reduce((maximum, candidate) => Math.max(maximum, candidate.end), 0),
    searchNames,
  );

  return undefined;
}

function getStartupRecordSymbolAtPosition(snapshot, document, position) {
  const argument = getStartupDbpfArgumentAtPosition(document, position);
  if (!argument) {
    return undefined;
  }

  const baseOffset = document.offsetAt(argument.range.start);
  const candidateRange = getLinkedRecordCandidateRangeAtOffset(
    argument.value,
    baseOffset,
    document.offsetAt(position),
  );
  if (!candidateRange) {
    return undefined;
  }

  const loadedDefinitionsByName = getStartupLoadedRecordDefinitionMap(
    snapshot,
    document,
    argument.range.end,
  );
  if (!loadedDefinitionsByName.has(candidateRange.name)) {
    return undefined;
  }

  return createEpicsSemanticSymbol(
    "record",
    candidateRange.name,
    document,
    candidateRange.start,
    candidateRange.end,
    new Set([candidateRange.name]),
  );
}

function getPvlistRecordSymbolAtPosition(document, position) {
  const lineInfo = getPvlistRecordReferenceForLine(document, position.line);
  if (!lineInfo || !lineInfo.range.contains(position)) {
    return undefined;
  }

  return createEpicsSemanticSymbol(
    "record",
    lineInfo.name,
    document,
    document.offsetAt(lineInfo.range.start),
    document.offsetAt(lineInfo.range.end),
    new Set([lineInfo.name]),
  );
}

function getProbeRecordSymbolAtPosition(document, position) {
  const lineInfo = getProbeRecordReferenceForLine(document, position.line);
  if (!lineInfo || !lineInfo.range.contains(position)) {
    return undefined;
  }

  return createEpicsSemanticSymbol(
    "record",
    lineInfo.name,
    document,
    document.offsetAt(lineInfo.range.start),
    document.offsetAt(lineInfo.range.end),
    new Set([lineInfo.name]),
  );
}

function getMacroSymbolAtPosition(document, position) {
  const offset = document.offsetAt(position);

  const assignmentSymbol =
    getStartupMacroSymbolAtOffset(document, offset) ||
    getSubstitutionsMacroSymbolAtOffset(document, offset) ||
    getPvlistMacroSymbolAtOffset(document, offset);
  if (assignmentSymbol) {
    return assignmentSymbol;
  }

  return getGenericMacroReferenceSymbolAtOffset(document, offset);
}

function getStartupMacroSymbolAtOffset(document, offset) {
  if (!isStartupDocument(document)) {
    return undefined;
  }

  for (const statement of extractStartupStatements(document.getText())) {
    if (
      statement.kind === "envSet" &&
      statement.nameStart !== undefined &&
      statement.nameEnd !== undefined &&
      offset >= statement.nameStart &&
      offset <= statement.nameEnd
    ) {
      return createEpicsSemanticSymbol(
        "macro",
        statement.name,
        document,
        statement.nameStart,
        statement.nameEnd,
      );
    }

    if (
      statement.kind === "load" &&
      statement.command === "dbLoadRecords" &&
      statement.macros &&
      statement.macroValueStart !== undefined
    ) {
      const extracted = extractNamedAssignmentsWithRanges(
        statement.macros,
        statement.macroValueStart,
      );
      for (const [macroName, range] of extracted.nameRanges.entries()) {
        if (offset < range.start || offset > range.end) {
          continue;
        }

        return createEpicsSemanticSymbol(
          "macro",
          macroName,
          document,
          range.start,
          range.end,
        );
      }
    }
  }

  return undefined;
}

function getSubstitutionsMacroSymbolAtOffset(document, offset) {
  if (!isSubstitutionsDocument(document)) {
    return undefined;
  }

  for (const block of extractSubstitutionBlocksWithRanges(document.getText())) {
    if (block.kind === "global") {
      const extracted = extractNamedAssignmentsWithRanges(block.body, block.bodyStart);
      for (const [macroName, range] of extracted.nameRanges.entries()) {
        if (offset < range.start || offset > range.end) {
          continue;
        }

        return createEpicsSemanticSymbol(
          "macro",
          macroName,
          document,
          range.start,
          range.end,
        );
      }
      continue;
    }

    if (block.kind !== "file") {
      continue;
    }

    const parsedRows = parseSubstitutionFileBlockRowsDetailed(block.body, block.bodyStart);
    if (parsedRows.kind === "pattern") {
      for (const range of extractSubstitutionPatternColumnRanges(block.body, block.bodyStart)) {
        if (offset < range.start || offset > range.end) {
          continue;
        }

        return createEpicsSemanticSymbol(
          "macro",
          range.name,
          document,
          range.start,
          range.end,
        );
      }
    }

    for (const row of parsedRows.rows) {
      for (const [macroName, range] of row.nameRanges.entries()) {
        if (offset < range.start || offset > range.end) {
          continue;
        }

        return createEpicsSemanticSymbol(
          "macro",
          macroName,
          document,
          range.start,
          range.end,
        );
      }
    }
  }

  return undefined;
}

function getPvlistMacroSymbolAtOffset(document, offset) {
  if (!isPvlistDocument(document)) {
    return undefined;
  }

  let lineStart = 0;
  for (const rawLine of document.getText().split(/\r?\n/)) {
    const match = rawLine.match(/^(\s*)([A-Za-z_][A-Za-z0-9_]*)\s*=/);
    if (match) {
      const start = lineStart + match[1].length;
      const end = start + match[2].length;
      if (offset >= start && offset <= end) {
        return createEpicsSemanticSymbol("macro", match[2], document, start, end);
      }
    }

    lineStart += rawLine.length + 1;
  }

  return undefined;
}

function getGenericMacroReferenceSymbolAtOffset(document, offset) {
  const maskedText = maskDatabaseComments(document.getText());
  const regex = /\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}/g;
  let match;

  while ((match = regex.exec(maskedText))) {
    const macroName = match[1] || match[2];
    const nameStart = match.index + 2;
    const nameEnd = nameStart + macroName.length;
    if (offset < nameStart || offset > nameEnd) {
      continue;
    }

    return createEpicsSemanticSymbol(
      "macro",
      macroName,
      document,
      nameStart,
      nameEnd,
    );
  }

  return undefined;
}

function createEpicsSemanticSymbol(
  kind,
  name,
  document,
  startOffset,
  endOffset,
  searchNames,
  extra = undefined,
) {
  return {
    kind,
    name,
    searchNames:
      searchNames instanceof Set && searchNames.size > 0
        ? new Set(searchNames)
        : new Set([name]),
    range: new vscode.Range(
      document.positionAt(startOffset),
      document.positionAt(endOffset),
    ),
    ...(extra && typeof extra === "object" ? extra : {}),
  };
}

function collectEpicsSymbolOccurrences(
  snapshot,
  symbol,
  currentDocument,
  includeDeclaration,
) {
  const occurrences = [];
  const seen = new Set();

  for (const document of collectSemanticReferenceDocuments(snapshot, currentDocument)) {
    switch (symbol.kind) {
      case "record":
        collectRecordOccurrencesInDocument(
          snapshot,
          document,
          symbol.searchNames || new Set([symbol.name]),
          includeDeclaration,
          occurrences,
          seen,
        );
        break;

      case "macro":
        collectMacroOccurrencesInDocument(
          document,
          symbol.name,
          includeDeclaration,
          occurrences,
          seen,
        );
        break;

      case "recordType":
        collectRecordTypeOccurrencesInDocument(
          document,
          symbol.name,
          includeDeclaration,
          occurrences,
          seen,
        );
        break;

      case "field":
        collectFieldOccurrencesInDocument(
          document,
          symbol.recordType,
          symbol.name,
          includeDeclaration,
          occurrences,
          seen,
        );
        break;

      case "deviceSupport":
      case "driver":
      case "registrar":
      case "function":
      case "variable":
        collectNamedDbdSymbolOccurrencesInDocument(
          document,
          symbol.kind,
          symbol.name,
          includeDeclaration,
          occurrences,
          seen,
        );
        break;

      default:
        break;
    }
  }

  return occurrences;
}

function collectSemanticReferenceDocuments(snapshot, currentDocument) {
  const documentsByUri = new Map();
  const addDocument = (document) => {
    if (!document?.uri || !isSemanticReferenceDocument(document)) {
      return;
    }

    documentsByUri.set(document.uri.toString(), document);
  };

  addDocument(currentDocument);
  for (const document of vscode.workspace.textDocuments) {
    addDocument(document);
  }

  for (const entry of snapshot.workspaceFiles) {
    if (!entry.uri || documentsByUri.has(entry.uri.toString())) {
      continue;
    }

    const languageId = getEpicsLanguageIdForUri(entry.uri);
    if (!isSemanticReferenceLanguageId(languageId)) {
      continue;
    }

    const text = entry.absolutePath ? readTextFile(entry.absolutePath) : undefined;
    if (text === undefined) {
      continue;
    }

    documentsByUri.set(
      entry.uri.toString(),
      createSyntheticTextDocument(entry.uri, languageId, text),
    );
  }

  return [...documentsByUri.values()];
}

function isSemanticReferenceLanguageId(languageId) {
  return new Set([
    LANGUAGE_IDS.database,
    LANGUAGE_IDS.startup,
    LANGUAGE_IDS.substitutions,
    LANGUAGE_IDS.dbd,
    LANGUAGE_IDS.source,
    LANGUAGE_IDS.pvlist,
    LANGUAGE_IDS.probe,
  ]).has(languageId);
}

function isSemanticReferenceDocument(document) {
  return (
    Boolean(document) &&
    (
      isSemanticReferenceLanguageId(document.languageId) ||
      isSourceDocument(document)
    )
  );
}

function collectRecordOccurrencesInDocument(
  snapshot,
  document,
  recordNames,
  includeDeclaration,
  occurrences,
  seen,
) {
  const searchNames =
    recordNames instanceof Set ? recordNames : new Set([String(recordNames || "")]);

  if (isDatabaseDocument(document)) {
    const text = document.getText();
    const macroAssignments = extractDatabaseTocMacroAssignments(text);

    if (includeDeclaration) {
      for (const declaration of extractRecordDeclarations(text)) {
        const declarationNames = getDatabaseRecordSearchNames(text, declaration.name);
        if (![...declarationNames].some((candidateName) => searchNames.has(candidateName))) {
          continue;
        }

        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(declaration.nameStart),
            document.positionAt(declaration.nameEnd),
          ),
        );
      }
    }

    for (const declaration of extractRecordDeclarations(text)) {
      for (const fieldDeclaration of extractFieldDeclarationsInRecord(text, declaration)) {
        for (const candidate of getLinkedRecordCandidateRanges(
          fieldDeclaration.value,
          fieldDeclaration.valueStart,
          {
            allowMacroReferences: true,
            macroAssignments,
          },
        )) {
          if (!searchNames.has(candidate.name)) {
            continue;
          }

          addSemanticOccurrence(
            occurrences,
            seen,
            document.uri,
            new vscode.Range(
              document.positionAt(candidate.start),
              document.positionAt(candidate.end),
            ),
          );
        }
      }
    }
    return;
  }

  if (isStartupDocument(document)) {
    const lineCount = document.getText().split(/\r?\n/).length;
    const regex = /dbpf\(\s*"((?:[^"\\]|\\.)*)"/g;

    for (let line = 0; line < lineCount; line += 1) {
      const lineText = `${document.lineAt(line).text}\n`;
      const sanitizedLineText = maskDatabaseComments(lineText).slice(0, -1);
      let match;
      while ((match = regex.exec(sanitizedLineText))) {
        const valueStart = match.index + match[0].length - match[1].length - 1;
        const absoluteValueStart = document.offsetAt(new vscode.Position(line, valueStart));
        const loadedDefinitionsByName = getStartupLoadedRecordDefinitionMap(
          snapshot,
          document,
          new vscode.Position(line, valueStart + match[1].length),
        );
        for (const candidate of getLinkedRecordCandidateRanges(match[1], absoluteValueStart)) {
          if (
            !searchNames.has(candidate.name) ||
            ![...searchNames].some((recordName) => loadedDefinitionsByName.has(recordName))
          ) {
            continue;
          }

          addSemanticOccurrence(
            occurrences,
            seen,
            document.uri,
            new vscode.Range(
              document.positionAt(candidate.start),
              document.positionAt(candidate.end),
            ),
          );
        }
      }
    }
    return;
  }

  if (isPvlistDocument(document)) {
    const lineCount = document.getText().split(/\r?\n/).length;
    for (let line = 0; line < lineCount; line += 1) {
      const recordReference = getPvlistRecordReferenceForLine(document, line);
      if (!recordReference || !searchNames.has(recordReference.name)) {
        continue;
      }

      addSemanticOccurrence(occurrences, seen, document.uri, recordReference.range);
    }
    return;
  }

  if (isProbeDocument(document)) {
    const lineCount = document.getText().split(/\r?\n/).length;
    for (let line = 0; line < lineCount; line += 1) {
      const recordReference = getProbeRecordReferenceForLine(document, line);
      if (!recordReference || !searchNames.has(recordReference.name)) {
        continue;
      }

      addSemanticOccurrence(occurrences, seen, document.uri, recordReference.range);
    }
  }
}

function collectRecordTypeOccurrencesInDocument(
  document,
  recordTypeName,
  includeDeclaration,
  occurrences,
  seen,
) {
  if (!recordTypeName) {
    return;
  }

  if (isDatabaseDocument(document)) {
    for (const declaration of extractRecordDeclarations(document.getText())) {
      if (declaration.recordType !== recordTypeName) {
        continue;
      }

      addSemanticOccurrence(
        occurrences,
        seen,
        document.uri,
        new vscode.Range(
          document.positionAt(declaration.recordTypeStart),
          document.positionAt(declaration.recordTypeEnd),
        ),
      );
    }
    return;
  }

  if (document?.languageId === LANGUAGE_IDS.dbd) {
    const text = document.getText();
    if (includeDeclaration) {
      for (const declaration of extractDbdRecordTypeDeclarations(text)) {
        if (declaration.name !== recordTypeName) {
          continue;
        }

        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(declaration.nameStart),
            document.positionAt(declaration.nameEnd),
          ),
        );
      }
    }

    for (const entry of extractDbdDeviceDeclarations(text)) {
      if (entry.recordType !== recordTypeName) {
        continue;
      }

      addSemanticOccurrence(
        occurrences,
        seen,
        document.uri,
        new vscode.Range(
          document.positionAt(entry.recordTypeStart),
          document.positionAt(entry.recordTypeEnd),
        ),
      );
    }
  }
}

function collectFieldOccurrencesInDocument(
  document,
  recordTypeName,
  fieldName,
  includeDeclaration,
  occurrences,
  seen,
) {
  if (!recordTypeName || !fieldName) {
    return;
  }

  if (isDatabaseDocument(document)) {
    const text = document.getText();
    for (const declaration of extractRecordDeclarations(text)) {
      if (declaration.recordType !== recordTypeName) {
        continue;
      }

      for (const fieldDeclaration of extractFieldDeclarationsInRecord(
        text,
        declaration,
        fieldName,
      )) {
        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(fieldDeclaration.fieldNameStart),
            document.positionAt(fieldDeclaration.fieldNameEnd),
          ),
        );
      }
    }
    return;
  }

  if (document?.languageId !== LANGUAGE_IDS.dbd || !includeDeclaration) {
    return;
  }

  const text = document.getText();
  for (const declaration of extractDbdRecordTypeDeclarations(text)) {
    if (declaration.name !== recordTypeName) {
      continue;
    }

    for (const fieldDeclaration of extractDbdFieldDeclarationsInRecordType(
      text,
      declaration,
      fieldName,
    )) {
      addSemanticOccurrence(
        occurrences,
        seen,
        document.uri,
        new vscode.Range(
          document.positionAt(fieldDeclaration.fieldNameStart),
          document.positionAt(fieldDeclaration.fieldNameEnd),
        ),
      );
    }
  }
}

function collectNamedDbdSymbolOccurrencesInDocument(
  document,
  kind,
  symbolName,
  includeDeclaration,
  occurrences,
  seen,
) {
  if (!symbolName) {
    return;
  }

  if (document?.languageId === LANGUAGE_IDS.dbd) {
    const text = document.getText();

    if (kind === "deviceSupport") {
      for (const entry of extractDbdDeviceDeclarations(text)) {
        if (entry.supportName !== symbolName) {
          continue;
        }

        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(entry.supportNameStart),
            document.positionAt(entry.supportNameEnd),
          ),
        );
      }
      return;
    }

    const keyword = getDbdNamedSymbolKeyword(kind);
    if (!keyword) {
      return;
    }

    for (const entry of extractDbdNamedEntries(text, keyword)) {
      if (entry.name !== symbolName) {
        continue;
      }

      addSemanticOccurrence(
        occurrences,
        seen,
        document.uri,
        new vscode.Range(
          document.positionAt(entry.nameStart),
          document.positionAt(entry.nameEnd),
        ),
      );
    }
    return;
  }

  if (!isSourceDocument(document) || !includeDeclaration) {
    return;
  }

  for (const entry of extractSourceNamedSymbolOccurrences(document.getText())) {
    if (entry.kind !== kind || entry.name !== symbolName) {
      continue;
    }

    addSemanticOccurrence(
      occurrences,
      seen,
      document.uri,
      new vscode.Range(
        document.positionAt(entry.nameStart),
        document.positionAt(entry.nameEnd),
      ),
    );
  }
}

function collectMacroOccurrencesInDocument(
  document,
  macroName,
  includeDeclaration,
  occurrences,
  seen,
) {
  if (includeDeclaration) {
    collectMacroDefinitionOccurrences(document, macroName, occurrences, seen);
  }
  collectMacroReferenceOccurrences(document, macroName, occurrences, seen);
}

function collectMacroDefinitionOccurrences(document, macroName, occurrences, seen) {
  if (isStartupDocument(document)) {
    for (const statement of extractStartupStatements(document.getText())) {
      if (
        statement.kind === "envSet" &&
        statement.name === macroName &&
        statement.nameStart !== undefined &&
        statement.nameEnd !== undefined
      ) {
        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(statement.nameStart),
            document.positionAt(statement.nameEnd),
          ),
        );
      }

      if (
        statement.kind === "load" &&
        statement.command === "dbLoadRecords" &&
        statement.macros &&
        statement.macroValueStart !== undefined
      ) {
        const extracted = extractNamedAssignmentsWithRanges(
          statement.macros,
          statement.macroValueStart,
        );
        const range = extracted.nameRanges.get(macroName);
        if (!range) {
          continue;
        }

        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(range.start),
            document.positionAt(range.end),
          ),
        );
      }
    }
  }

  if (isSubstitutionsDocument(document)) {
    for (const block of extractSubstitutionBlocksWithRanges(document.getText())) {
      if (block.kind === "global") {
        const extracted = extractNamedAssignmentsWithRanges(block.body, block.bodyStart);
        const range = extracted.nameRanges.get(macroName);
        if (range) {
          addSemanticOccurrence(
            occurrences,
            seen,
            document.uri,
            new vscode.Range(
              document.positionAt(range.start),
              document.positionAt(range.end),
            ),
          );
        }
        continue;
      }

      if (block.kind !== "file") {
        continue;
      }

      const parsedRows = parseSubstitutionFileBlockRowsDetailed(block.body, block.bodyStart);
      if (parsedRows.kind === "pattern") {
        for (const range of extractSubstitutionPatternColumnRanges(block.body, block.bodyStart)) {
          if (range.name !== macroName) {
            continue;
          }

          addSemanticOccurrence(
            occurrences,
            seen,
            document.uri,
            new vscode.Range(
              document.positionAt(range.start),
              document.positionAt(range.end),
            ),
          );
        }
      }

      for (const row of parsedRows.rows) {
        const range = row.nameRanges.get(macroName);
        if (!range) {
          continue;
        }

        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(range.start),
            document.positionAt(range.end),
          ),
        );
      }
    }
  }

  if (isPvlistDocument(document)) {
    let lineStart = 0;
    for (const rawLine of document.getText().split(/\r?\n/)) {
      const match = rawLine.match(/^(\s*)([A-Za-z_][A-Za-z0-9_]*)\s*=/);
      if (match?.[2] === macroName) {
        const start = lineStart + match[1].length;
        const end = start + macroName.length;
        addSemanticOccurrence(
          occurrences,
          seen,
          document.uri,
          new vscode.Range(
            document.positionAt(start),
            document.positionAt(end),
          ),
        );
      }

      lineStart += rawLine.length + 1;
    }
  }
}

function collectMacroReferenceOccurrences(document, macroName, occurrences, seen) {
  const maskedText = maskDatabaseComments(document.getText());
  const regex = /\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}/g;
  let match;

  while ((match = regex.exec(maskedText))) {
    const matchedName = match[1] || match[2];
    if (matchedName !== macroName) {
      continue;
    }

    const start = match.index + 2;
    const end = start + matchedName.length;
    addSemanticOccurrence(
      occurrences,
      seen,
      document.uri,
      new vscode.Range(
        document.positionAt(start),
        document.positionAt(end),
      ),
    );
  }
}

function getLinkedRecordCandidateRangeAtOffset(fieldValue, baseOffset, offset) {
  return getLinkedRecordCandidateRanges(fieldValue, baseOffset)
    .filter((candidate) => offset >= candidate.start && offset <= candidate.end)
    .sort(
      (left, right) =>
        (left.end - left.start) - (right.end - right.start) ||
        compareLabels(left.name, right.name),
    )[0];
}

function getLinkedRecordCandidateRangeForValue(fieldValue, baseOffset, recordName) {
  return getLinkedRecordCandidateRanges(fieldValue, baseOffset).find(
    (candidate) => candidate.name === recordName,
  );
}

function getLinkedRecordCandidateRanges(fieldValue, baseOffset = 0, options = undefined) {
  if (!fieldValue) {
    return [];
  }

  const text = String(fieldValue || "");
  const leadingWhitespaceLength = text.match(/^\s*/)?.[0]?.length || 0;
  const trimmedText = text.slice(leadingWhitespaceLength);
  const tokenMatch = trimmedText.match(/^[^\s]+/);
  if (!tokenMatch) {
    return [];
  }

  let token = tokenMatch[0];
  let tokenStart = baseOffset + leadingWhitespaceLength;
  if (token.startsWith("@")) {
    return [];
  }

  const protocolMatch = token.match(/^(?:ca|pva):\/\//i);
  if (protocolMatch) {
    tokenStart += protocolMatch[0].length;
    token = token.slice(protocolMatch[0].length);
  }

  const normalizedToken = token.replace(/[),;]+$/, "");
  if (!normalizedToken) {
    return [];
  }

  const sourceLastDotIndex = normalizedToken.lastIndexOf(".");
  const sourceFieldSuffix =
    sourceLastDotIndex > 0 ? normalizedToken.slice(sourceLastDotIndex + 1) : undefined;
  const sourceRecordEnd =
    sourceLastDotIndex > 0 && /^[A-Z0-9_]+$/.test(sourceFieldSuffix)
      ? tokenStart + sourceLastDotIndex
      : undefined;

  const ranges = [];
  const addRange = (name, start, end) => {
    if (
      !name ||
      ranges.some(
        (existing) => existing.name === name && existing.start === start && existing.end === end,
      )
    ) {
      return;
    }

    ranges.push({ name, start, end });
  };

  const addTokenVariants = (tokenValue, start, end) => {
    if (!tokenValue) {
      return;
    }

    addRange(tokenValue, start, end);

    const lastDotIndex = tokenValue.lastIndexOf(".");
    if (lastDotIndex > 0) {
      const suffix = tokenValue.slice(lastDotIndex + 1);
      if (/^[A-Z0-9_]+$/.test(suffix) && sourceRecordEnd !== undefined) {
        addRange(
          tokenValue.slice(0, lastDotIndex),
          start,
          sourceRecordEnd,
        );
      }
    }
  };

  if (containsEpicsMacroReference(normalizedToken)) {
    if (!options?.allowMacroReferences) {
      return [];
    }

    const expandedToken = resolveDatabaseRecordNameFromToc(
      normalizedToken,
      options?.macroAssignments instanceof Map ? options.macroAssignments : new Map(),
    );
    if (expandedToken && expandedToken !== normalizedToken) {
      addTokenVariants(expandedToken, tokenStart, tokenStart + normalizedToken.length);
    }
    addTokenVariants(normalizedToken, tokenStart, tokenStart + normalizedToken.length);

    return ranges;
  }

  addTokenVariants(normalizedToken, tokenStart, tokenStart + normalizedToken.length);
  return ranges;
}

function getPvlistRecordReferenceForLine(document, lineNumber) {
  if (!isPvlistDocument(document) || lineNumber < 0 || lineNumber >= document.lineCount) {
    return undefined;
  }

  const line = document.lineAt(lineNumber);
  const rawLine = String(line.text || "");
  const trimmedLine = rawLine.trim();
  if (
    !trimmedLine ||
    trimmedLine.startsWith("#") ||
    /^[A-Za-z_][A-Za-z0-9_]*\s*=/.test(trimmedLine) ||
    containsEpicsMacroReference(trimmedLine)
  ) {
    return undefined;
  }

  const leadingWhitespaceLength = rawLine.length - rawLine.trimStart().length;
  const protocolPrefixMatch = trimmedLine.match(/^(?:ca|pva):\/\//i);
  const baseOffset =
    document.offsetAt(new vscode.Position(lineNumber, leadingWhitespaceLength)) +
    (protocolPrefixMatch ? protocolPrefixMatch[0].length : 0);
  const candidate = getLinkedRecordCandidateRanges(
    protocolPrefixMatch ? trimmedLine.slice(protocolPrefixMatch[0].length) : trimmedLine,
    baseOffset,
  )[0];
  if (!candidate) {
    return undefined;
  }

  return {
    name: candidate.name,
    range: new vscode.Range(
      document.positionAt(candidate.start),
      document.positionAt(candidate.end),
    ),
  };
}

function getProbeRecordReferenceForLine(document, lineNumber) {
  if (!isProbeDocument(document) || lineNumber < 0 || lineNumber >= document.lineCount) {
    return undefined;
  }

  const line = document.lineAt(lineNumber);
  const rawLine = String(line.text || "");
  const trimmedLine = rawLine.trim();
  if (
    !trimmedLine ||
    trimmedLine.startsWith("#") ||
    /\s/.test(trimmedLine) ||
    containsEpicsMacroReference(trimmedLine)
  ) {
    return undefined;
  }

  const leadingWhitespaceLength = rawLine.length - rawLine.trimStart().length;
  const candidate = getLinkedRecordCandidateRanges(
    trimmedLine,
    document.offsetAt(new vscode.Position(lineNumber, leadingWhitespaceLength)),
  )[0];
  if (!candidate) {
    return undefined;
  }

  return {
    name: candidate.name,
    range: new vscode.Range(
      document.positionAt(candidate.start),
      document.positionAt(candidate.end),
    ),
  };
}

function getShortestRecordCandidateNames(candidateRanges) {
  if (!Array.isArray(candidateRanges) || candidateRanges.length === 0) {
    return new Set();
  }

  const minimumLength = candidateRanges.reduce(
    (minimum, candidate) => Math.min(minimum, candidate.end - candidate.start),
    Number.POSITIVE_INFINITY,
  );
  return new Set(
    candidateRanges
      .filter((candidate) => candidate.end - candidate.start === minimumLength)
      .map((candidate) => candidate.name)
      .filter(Boolean),
  );
}

function getDatabaseRecordSearchNames(documentText, recordName) {
  const normalizedRecordName = String(recordName || "").trim();
  const names = new Set();
  if (!normalizedRecordName) {
    return names;
  }

  names.add(normalizedRecordName);
  if (!documentText || !containsEpicsMacroReference(normalizedRecordName)) {
    return names;
  }

  const expandedName = resolveDatabaseRecordNameFromToc(
    normalizedRecordName,
    extractDatabaseTocMacroAssignments(documentText),
  );
  if (expandedName && expandedName !== normalizedRecordName) {
    names.add(expandedName);
  }

  return names;
}

function findDatabaseDeclarationByAnyRecordName(documentText, recordName) {
  for (const declaration of extractRecordDeclarations(documentText)) {
    const searchNames = getDatabaseRecordSearchNames(documentText, declaration.name);
    if (!searchNames.has(recordName)) {
      continue;
    }

    return {
      declaration,
      searchNames,
    };
  }

  return undefined;
}

function extractSubstitutionPatternColumnRanges(body, baseOffset = 0) {
  if (!/^\s*pattern\b/.test(body)) {
    return [];
  }

  const segments = extractTopLevelBraceSegments(body, baseOffset);
  if (segments.length === 0) {
    return [];
  }

  const headerSegment = segments[0];
  const regex = /"((?:[^"\\]|\\.)*)"|([A-Za-z_][A-Za-z0-9_]*)/g;
  const ranges = [];
  let match;

  while ((match = regex.exec(headerSegment.text))) {
    const name = match[1] || match[2];
    const offset = match[1] ? 1 : 0;
    ranges.push({
      name,
      start: headerSegment.contentStart + match.index + offset,
      end: headerSegment.contentStart + match.index + offset + name.length,
    });
  }

  return ranges;
}

function addSemanticOccurrence(occurrences, seen, uri, range) {
  const key = `${uri.toString()}:${range.start.line}:${range.start.character}:${range.end.line}:${range.end.character}`;
  if (seen.has(key)) {
    return;
  }

  seen.add(key);
  occurrences.push({ uri, range });
}

function buildEpicsRenameWorkspaceEdit(snapshot, document, position, newName) {
  const symbol = getEpicsSemanticSymbolAtPosition(snapshot, document, position);
  if (!symbol) {
    throw new Error("Nothing renameable at the current cursor position.");
  }

  const normalizedName = validateEpicsRename(snapshot, symbol, document, newName);

  const occurrences = collectEpicsSymbolOccurrences(
    snapshot,
    symbol,
    document,
    true,
  );
  if (occurrences.length === 0) {
    return undefined;
  }

  const edit = new vscode.WorkspaceEdit();
  for (const occurrence of occurrences) {
    edit.replace(occurrence.uri, occurrence.range, normalizedName);
  }

  return edit;
}

function validateEpicsRename(snapshot, symbol, document, newName) {
  const trimmedName = String(newName || "").trim();
  if (!trimmedName) {
    throw new Error("The new name must not be empty.");
  }

  switch (symbol.kind) {
    case "macro": {
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmedName)) {
        throw new Error("EPICS macro names must match [A-Za-z_][A-Za-z0-9_]*.");
      }

      const conflicts = [];
      collectMacroDefinitionOccurrences(document, trimmedName, conflicts, new Set());
      if (conflicts.some((occurrence) => !rangesEqual(occurrence.range, symbol.range))) {
        throw new Error(`Macro "${trimmedName}" is already defined in this file.`);
      }
      break;
    }

    case "record": {
      if (/[\s",]/.test(trimmedName)) {
        throw new Error("EPICS record names must not contain spaces, quotes, or commas.");
      }

      const conflictSearchNames = isDatabaseDocument(document)
        ? getDatabaseRecordSearchNames(document.getText(), trimmedName)
        : new Set([trimmedName]);
      const currentSearchNames = symbol.searchNames || new Set([symbol.name]);

      for (const referenceDocument of collectSemanticReferenceDocuments(snapshot, document)) {
        if (!isDatabaseDocument(referenceDocument)) {
          continue;
        }

        const referenceText = referenceDocument.getText();
        for (const declaration of extractRecordDeclarations(referenceText)) {
          const declarationNames = getDatabaseRecordSearchNames(referenceText, declaration.name);
          const hasConflict = [...declarationNames].some((candidateName) =>
            conflictSearchNames.has(candidateName),
          );
          const isCurrentDeclaration =
            referenceDocument.uri.toString() === document.uri.toString() &&
            rangesEqual(
              new vscode.Range(
                referenceDocument.positionAt(declaration.nameStart),
                referenceDocument.positionAt(declaration.nameEnd),
              ),
              symbol.range,
            ) &&
            [...declarationNames].some((candidateName) => currentSearchNames.has(candidateName));
          if (hasConflict && !isCurrentDeclaration) {
            throw new Error(`Record "${trimmedName}" is already defined in the workspace.`);
          }
        }
      }
      break;
    }

    case "recordType":
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmedName)) {
        throw new Error("EPICS record types must match [A-Za-z_][A-Za-z0-9_]*.");
      }
      if (!hasWorkspaceDbdRecordTypeDeclaration(snapshot, document, symbol.name)) {
        throw new Error(
          `Record type "${symbol.name}" can only be renamed when it is declared in a workspace .dbd file.`,
        );
      }
      if (trimmedName !== symbol.name && snapshot.recordTypes.has(trimmedName)) {
        throw new Error(`Record type "${trimmedName}" already exists in the workspace.`);
      }
      break;

    case "field":
      if (!/^[A-Z][A-Z0-9_]*$/.test(trimmedName)) {
        throw new Error("EPICS field names must match [A-Z][A-Z0-9_]*.");
      }
      if (!hasWorkspaceDbdFieldDeclaration(snapshot, document, symbol.recordType, symbol.name)) {
        throw new Error(
          `Field "${symbol.name}" can only be renamed when "${symbol.recordType}.${symbol.name}" is declared in a workspace .dbd file.`,
        );
      }
      if (
        trimmedName !== symbol.name &&
        snapshot.fieldsByRecordType.get(symbol.recordType)?.has(trimmedName)
      ) {
        throw new Error(
          `Field "${trimmedName}" already exists on record type "${symbol.recordType}".`,
        );
      }
      break;

    case "deviceSupport":
    case "driver":
    case "registrar":
    case "function":
    case "variable":
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmedName)) {
        throw new Error(`${formatEpicsSymbolKindLabel(symbol.kind)} names must match [A-Za-z_][A-Za-z0-9_]*.`);
      }
      if (
        trimmedName !== symbol.name &&
        getNamedSymbolDefinitionMap(snapshot, symbol.kind)?.has(trimmedName)
      ) {
        throw new Error(
          `${formatEpicsSymbolKindLabel(symbol.kind)} "${trimmedName}" already exists in the workspace.`,
        );
      }
      break;

    default:
      break;
  }

  return trimmedName;
}

function getNamedSymbolDefinitionMap(snapshot, kind) {
  switch (kind) {
    case "deviceSupport":
      return snapshot.deviceSupportDefinitionsByName;
    case "driver":
      return snapshot.driverDefinitionsByName;
    case "registrar":
      return snapshot.registrarDefinitionsByName;
    case "function":
      return snapshot.functionDefinitionsByName;
    case "variable":
      return snapshot.variableDefinitionsByName;
    default:
      return undefined;
  }
}

function formatEpicsSymbolKindLabel(kind) {
  return {
    deviceSupport: "Device support",
    driver: "Driver",
    registrar: "Registrar",
    function: "Function",
    variable: "Variable",
  }[kind] || "EPICS symbol";
}

function rangesEqual(left, right) {
  return (
    Boolean(left && right) &&
    left.start.line === right.start.line &&
    left.start.character === right.start.character &&
    left.end.line === right.end.line &&
    left.end.character === right.end.character
  );
}

function hasWorkspaceDbdRecordTypeDeclaration(snapshot, currentDocument, recordTypeName) {
  return collectSemanticReferenceDocuments(snapshot, currentDocument)
    .filter((document) => document?.languageId === LANGUAGE_IDS.dbd)
    .some((document) =>
      extractDbdRecordTypeDeclarations(document.getText()).some(
        (declaration) => declaration.name === recordTypeName,
      ),
    );
}

function hasWorkspaceDbdFieldDeclaration(snapshot, currentDocument, recordTypeName, fieldName) {
  return collectSemanticReferenceDocuments(snapshot, currentDocument)
    .filter((document) => document?.languageId === LANGUAGE_IDS.dbd)
    .some((document) =>
      extractDbdRecordTypeDeclarations(document.getText()).some((declaration) =>
        declaration.name === recordTypeName &&
        extractDbdFieldDeclarationsInRecordType(
          document.getText(),
          declaration,
          fieldName,
        ).length > 0,
      ),
    );
}

function getEpicsCodeActions(snapshot, document, range, context) {
  const actions = [];
  for (const diagnostic of context?.diagnostics || []) {
    if (
      diagnostic.code === "epics.startup.missingDbLoadRecordsMacros" &&
      isStartupDocument(document)
    ) {
      const action = createStartupLoadMacroQuickFix(snapshot, document, diagnostic);
      if (action) {
        actions.push(action);
      }
      continue;
    }

    if (
      diagnostic.code === "epics.database.duplicateRecordName" &&
      isDatabaseDocument(document)
    ) {
      const action = createDuplicateRecordNameQuickFix(document, diagnostic);
      if (action) {
        actions.push(action);
      }
      continue;
    }

    if (
      diagnostic.code === "epics.database.invalidFieldName" &&
      isDatabaseDocument(document)
    ) {
      const action = createInvalidFieldNameQuickFix(snapshot, document, diagnostic);
      if (action) {
        actions.push(action);
      }
      continue;
    }

    if (
      diagnostic.code === "epics.database.invalidMenuFieldValue" &&
      isDatabaseDocument(document)
    ) {
      const action = createInvalidMenuFieldValueQuickFix(snapshot, document, diagnostic);
      if (action) {
        actions.push(action);
      }
      continue;
    }

    if (
      diagnostic.code === "epics.startup.unknownIocRegistrationFunction" &&
      isStartupDocument(document)
    ) {
      actions.push(...createUnknownIocRegistrationQuickFixes(snapshot, document, diagnostic));
    }
  }

  return actions;
}

function createStartupLoadMacroQuickFix(snapshot, document, diagnostic) {
  const statement = extractStartupStatements(document.getText()).find(
    (candidate) =>
      candidate.kind === "load" &&
      candidate.command === "dbLoadRecords" &&
      candidate.pathStart === document.offsetAt(diagnostic.range.start) &&
      candidate.pathEnd === document.offsetAt(diagnostic.range.end),
  );
  if (!statement) {
    return undefined;
  }

  const resolvedMacroData = resolveStartupLoadMacroData(
    snapshot,
    document,
    diagnostic.range.end,
    statement.path,
  );
  if (!resolvedMacroData || resolvedMacroData.macroNames.length === 0) {
    return undefined;
  }

  const providedMacroNames = extractAssignedMacroNames(statement.macros);
  const missingMacroNames = resolvedMacroData.macroNames.filter(
    (macroName) => !providedMacroNames.has(macroName),
  );
  if (missingMacroNames.length === 0) {
    return undefined;
  }

  const action = new vscode.CodeAction(
    `Add missing dbLoadRecords macros: ${missingMacroNames.join(", ")}`,
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.isPreferred = true;
  action.edit = new vscode.WorkspaceEdit();

  if (statement.macroValueStart !== undefined && statement.macroValueEnd !== undefined) {
    const existingMacros = String(statement.macros || "");
    const appendedMacros = buildDbLoadRecordsMacroAssignmentsLabel(missingMacroNames);
    const separator = existingMacros.trim() ? "," : "";
    action.edit.replace(
      document.uri,
      new vscode.Range(
        document.positionAt(statement.macroValueStart),
        document.positionAt(statement.macroValueEnd),
      ),
      `${existingMacros}${separator}${appendedMacros}`,
    );
  } else {
    action.edit.insert(
      document.uri,
      document.positionAt(statement.pathEnd + 1),
      `, "${buildDbLoadRecordsMacroAssignmentsLabel(missingMacroNames)}"`,
    );
  }

  return action;
}

function createDuplicateRecordNameQuickFix(document, diagnostic) {
  const currentName = document.getText(diagnostic.range);
  if (!currentName) {
    return undefined;
  }

  const newName = suggestUniqueRecordName(document.getText(), currentName);
  const action = new vscode.CodeAction(
    `Rename duplicate record to "${newName}"`,
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.isPreferred = true;
  action.edit = new vscode.WorkspaceEdit();
  action.edit.replace(document.uri, diagnostic.range, newName);
  return action;
}

function createInvalidFieldNameQuickFix(snapshot, document, diagnostic) {
  const invalidFieldName = document.getText(diagnostic.range);
  if (!invalidFieldName) {
    return undefined;
  }

  const recordDeclaration = findEnclosingRecordDeclaration(
    document.getText(),
    document.offsetAt(diagnostic.range.start),
  );
  if (!recordDeclaration) {
    return undefined;
  }

  const replacement = findBestMatchingLabel(
    getFieldNamesForRecordType(snapshot, recordDeclaration.recordType).filter(
      (fieldName) => fieldName !== invalidFieldName,
    ),
    invalidFieldName,
  );
  if (!replacement) {
    return undefined;
  }

  const action = new vscode.CodeAction(
    `Replace invalid field with "${replacement}"`,
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.isPreferred = true;
  action.edit = new vscode.WorkspaceEdit();
  action.edit.replace(document.uri, diagnostic.range, replacement);
  return action;
}

function createInvalidMenuFieldValueQuickFix(snapshot, document, diagnostic) {
  const invalidValue = document.getText(diagnostic.range);
  const text = document.getText();
  const recordDeclaration = findEnclosingRecordDeclaration(
    text,
    document.offsetAt(diagnostic.range.start),
  );
  if (!recordDeclaration) {
    return undefined;
  }

  const fieldDeclaration = extractFieldDeclarationsInRecord(text, recordDeclaration).find(
    (candidate) =>
      candidate.valueStart === document.offsetAt(diagnostic.range.start) &&
      candidate.valueEnd === document.offsetAt(diagnostic.range.end),
  );
  if (!fieldDeclaration) {
    return undefined;
  }

  const allowedChoices = getMenuFieldChoices(
    snapshot,
    recordDeclaration.recordType,
    fieldDeclaration.fieldName,
  ).filter((choice) => choice !== invalidValue);
  const replacement =
    findBestMatchingLabel(allowedChoices, invalidValue) ||
    getFieldInitialValue(snapshot, recordDeclaration.recordType, fieldDeclaration.fieldName) ||
    allowedChoices[0];
  if (!replacement) {
    return undefined;
  }

  const action = new vscode.CodeAction(
    `Replace with menu value "${replacement}"`,
    vscode.CodeActionKind.QuickFix,
  );
  action.diagnostics = [diagnostic];
  action.isPreferred = true;
  action.edit = new vscode.WorkspaceEdit();
  action.edit.replace(document.uri, diagnostic.range, replacement);
  return action;
}

function createUnknownIocRegistrationQuickFixes(snapshot, document, diagnostic) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (!project || project.iocsByName.size === 0) {
    return [];
  }

  const currentName = document.getText(diagnostic.range);
  const actions = [];
  for (const iocName of [...project.iocsByName.keys()].sort(compareLabels)) {
    const replacement = `${iocName}_registerRecordDeviceDriver`;
    if (replacement === currentName) {
      continue;
    }

    const action = new vscode.CodeAction(
      `Replace with "${replacement}"`,
      vscode.CodeActionKind.QuickFix,
    );
    action.diagnostics = [diagnostic];
    action.edit = new vscode.WorkspaceEdit();
    action.edit.replace(document.uri, diagnostic.range, replacement);
    if (actions.length === 0) {
      action.isPreferred = true;
    }
    actions.push(action);
  }

  return actions;
}

function suggestUniqueRecordName(text, recordName) {
  const existingNames = new Set(
    extractRecordDeclarations(text).map((declaration) => declaration.name),
  );
  let suffix = 1;
  let candidate = `${recordName}_${suffix}`;
  while (existingNames.has(candidate)) {
    suffix += 1;
    candidate = `${recordName}_${suffix}`;
  }

  return candidate;
}

function createSyntheticTextDocument(uri, languageId, text) {
  const documentText = String(text || "");
  const lineOffsets = createLineOffsets(documentText);

  return {
    uri,
    languageId,
    fileName: uri?.fsPath || uri?.path || "",
    lineCount: lineOffsets.length,
    getText(range) {
      if (!range) {
        return documentText;
      }

      return documentText.slice(this.offsetAt(range.start), this.offsetAt(range.end));
    },
    positionAt(offset) {
      return positionAtOffset(documentText, lineOffsets, offset);
    },
    offsetAt(position) {
      return offsetAtPosition(documentText, lineOffsets, position);
    },
    lineAt(line) {
      const safeLine = Math.max(0, Math.min(line, lineOffsets.length - 1));
      const start = lineOffsets[safeLine];
      const end =
        safeLine + 1 < lineOffsets.length
          ? Math.max(start, lineOffsets[safeLine + 1] - 1)
          : documentText.length;
      const textValue =
        documentText[end - 1] === "\r"
          ? documentText.slice(start, end - 1)
          : documentText.slice(start, end);
      return {
        text: textValue,
        range: new vscode.Range(
          new vscode.Position(safeLine, 0),
          new vscode.Position(safeLine, textValue.length),
        ),
      };
    },
  };
}

function createLineOffsets(text) {
  const offsets = [0];
  for (let index = 0; index < text.length; index += 1) {
    if (text[index] === "\n") {
      offsets.push(index + 1);
    }
  }

  return offsets;
}

function positionAtOffset(text, lineOffsets, offset) {
  const safeOffset = Math.max(0, Math.min(offset, text.length));
  let low = 0;
  let high = lineOffsets.length;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (lineOffsets[mid] > safeOffset) {
      high = mid;
    } else {
      low = mid + 1;
    }
  }

  const line = Math.max(0, low - 1);
  return new vscode.Position(line, safeOffset - lineOffsets[line]);
}

function offsetAtPosition(text, lineOffsets, position) {
  if (!position) {
    return 0;
  }

  const safeLine = Math.max(0, Math.min(position.line, lineOffsets.length - 1));
  const lineStart = lineOffsets[safeLine];
  const lineEnd =
    safeLine + 1 < lineOffsets.length
      ? lineOffsets[safeLine + 1]
      : text.length;
  return Math.max(lineStart, Math.min(lineStart + position.character, lineEnd));
}

function getMakefileDatabaseFileHover(snapshot, document, position) {
  if (!isMakefileDocument(document)) {
    return undefined;
  }

  const reference = getMakefileReferenceAtPosition(document, position);
  if (!reference || reference.kind !== "dbFile") {
    return undefined;
  }

  const target = getMakefileDatabaseReferenceTarget(
    snapshot,
    document,
    reference,
  );
  const absolutePath = target?.absolutePath;
  if (!absolutePath) {
    return undefined;
  }

  const text = readTextFile(absolutePath);
  const extension = path.extname(absolutePath).toLowerCase();
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;

  if (DATABASE_EXTENSIONS.has(extension)) {
    markdown.appendMarkdown("**EPICS database/template file**");
    if (path.posix.basename(reference.name) !== path.basename(absolutePath)) {
      markdown.appendMarkdown(
        `\n\nInstalled name: \`${escapeInlineCode(reference.name)}\``,
      );
    }

    if (text === undefined) {
      markdown.appendMarkdown(
        `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
      );
    } else {
      appendDatabaseFileHoverSummary(markdown, absolutePath, text);
    }
  } else if (SUBSTITUTION_EXTENSIONS.has(extension)) {
    markdown.appendMarkdown("**EPICS substitutions file**");
    if (path.posix.basename(reference.name) !== path.basename(absolutePath)) {
      markdown.appendMarkdown(
        `\n\nInstalled name: \`${escapeInlineCode(reference.name)}\``,
      );
    }

    if (text === undefined) {
      markdown.appendMarkdown(
        `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
      );
    } else {
      appendSubstitutionFileHoverSummary(markdown, absolutePath, text);
    }
  } else {
    markdown.appendMarkdown("**EPICS file**");
    markdown.appendMarkdown(
      `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
    );
  }

  return new vscode.Hover(
    markdown,
    new vscode.Range(
      document.positionAt(reference.start),
      document.positionAt(reference.end),
    ),
  );
}

function appendDatabaseFileHoverSummary(markdown, absolutePath, text) {
  const macroNames = extractMacroNames(maskDatabaseComments(text)).sort(compareLabels);
  const recordDeclarations = extractRecordDeclarations(text);
  const recordPreviewNames = recordDeclarations
    .slice(0, 100)
    .map((declaration) => declaration.name);

  markdown.appendMarkdown(
    `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
  );
  markdown.appendMarkdown(`\n\nRecords: \`${recordDeclarations.length}\``);
  if (macroNames.length > 0) {
    markdown.appendMarkdown(
      `\n\nMacros: ${macroNames
        .map((macroName) => `\`${escapeInlineCode(macroName)}\``)
        .join(", ")}`,
    );
  } else {
    markdown.appendMarkdown("\n\nMacros: none");
  }

  if (recordPreviewNames.length > 0) {
    markdown.appendMarkdown("\n\nRecord name preview:");
    markdown.appendCodeblock(recordPreviewNames.join("\n"), "text");
    if (recordDeclarations.length > recordPreviewNames.length) {
      markdown.appendMarkdown(
        `\n\n${recordDeclarations.length - recordPreviewNames.length} more record names omitted.`,
      );
    }
  }
}

function appendSubstitutionFileHoverSummary(markdown, absolutePath, text) {
  const blocks = extractSubstitutionBlocks(text);
  const expansionCount = blocks.reduce(
    (count, block) => count + parseSubstitutionFileBlockRows(block.body).length,
    0,
  );
  const lines = String(text).split(/\r?\n/);
  const previewLineLimit = 200;
  const previewLines = lines.slice(0, previewLineLimit);

  markdown.appendMarkdown(
    `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
  );
  markdown.appendMarkdown(`\n\nBlocks: \`${blocks.length}\``);
  markdown.appendMarkdown(`\n\nExpansions: \`${expansionCount}\``);
  if (previewLines.length > 0) {
    markdown.appendMarkdown("\n\nContent preview:");
    markdown.appendCodeblock(previewLines.join("\n"), LANGUAGE_IDS.substitutions);
    if (lines.length > previewLines.length) {
      markdown.appendMarkdown(
        `\n\n${lines.length - previewLines.length} more lines omitted.`,
      );
    }
  }
}

function getVariableHover(snapshot, document, position) {
  const reference = getVariableReferenceAtPosition(document, position);
  if (!reference) {
    return undefined;
  }

  if (isMakefileDocument(document)) {
    const resolved = resolveMakefileVariableValue(snapshot, document, reference.variableName);
    if (!resolved?.resolvedValue) {
      return undefined;
    }

    return createVariableHover(reference, resolved.resolvedValue, resolved.absolutePath);
  }

  if (isStartupStateDocument(document)) {
    const statement = getStartupStatementAtPosition(document, position);
    const untilOffset = statement ? statement.start : document.offsetAt(position);
    const state = createStartupExecutionState(snapshot, document, untilOffset);
    const resolvedValue =
      state.envVariables.get(reference.variableName) ||
      process.env[reference.variableName];
    if (!resolvedValue) {
      return undefined;
    }

    const absolutePath = computeAbsoluteVariablePath(
      resolvedValue,
      state.currentDirectory || path.dirname(document.uri.fsPath),
    );
    return createVariableHover(
      reference,
      resolvedValue,
      absolutePath,
      state.envVariableSources?.get(reference.variableName),
    );
  }

  return undefined;
}

function createVariableHover(reference, resolvedValue, absolutePath, sourceInfo) {
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(`**${escapeInlineCode(reference.variableName)}**`);
  markdown.appendMarkdown(`\n\nResolved value: \`${escapeInlineCode(resolvedValue)}\``);

  if (absolutePath && absolutePath !== resolvedValue) {
    markdown.appendMarkdown(`\n\nAbsolute path: \`${escapeInlineCode(absolutePath)}\``);
  }

  if (sourceInfo?.sourcePath) {
    const linkLabel = sourceInfo.line
      ? `${sourceInfo.sourcePath}:${sourceInfo.line}`
      : sourceInfo.sourcePath;
    markdown.appendMarkdown(
      `\n\nDefined in: ${
        sourceInfo.line
          ? createRecordLocationLink(
            {
              absolutePath: sourceInfo.sourcePath,
              line: sourceInfo.line,
            },
            linkLabel,
          )
          : createProtocolFileLink(sourceInfo.sourcePath, linkLabel)
      }`,
    );
    if (sourceInfo.sourceKind) {
      markdown.appendMarkdown(`\n\nSource: \`${escapeInlineCode(sourceInfo.sourceKind)}\``);
    }
    if (sourceInfo.rawValue !== undefined) {
      markdown.appendMarkdown(
        `\n\nAssigned value: \`${escapeInlineCode(sourceInfo.rawValue)}\``,
      );
    }
  }

  return new vscode.Hover(markdown, reference.range);
}

function createLinkedRecordHover(
  recordName,
  fieldName,
  definitions,
  range,
) {
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown(`**Linked record:** \`${escapeInlineCode(recordName)}\``);
  markdown.appendMarkdown(`\n\nReferenced by \`${escapeInlineCode(fieldName)}\``);

  const displayedDefinitions = definitions.slice(0, 3);
  for (const definition of displayedDefinitions) {
    const location = `${definition.relativePath || definition.absolutePath}:${definition.line}`;
    markdown.appendMarkdown("\n\n---\n\n");
    markdown.appendMarkdown(
      `Type: \`${escapeInlineCode(definition.recordType || "unknown")}\``,
    );
    if (definition.absolutePath) {
      markdown.appendMarkdown(
        `\n\nLocation: ${createRecordLocationLink(definition, location)}`,
      );
    } else {
      markdown.appendMarkdown(
        `\n\nLocation: \`${escapeInlineCode(location)}\``,
      );
    }
    markdown.appendMarkdown("\n\nPreview:");
    markdown.appendCodeblock(
      definition.preview || `record(${definition.recordType}, "${definition.name}")`,
      "db",
    );
  }

  if (definitions.length > displayedDefinitions.length) {
    markdown.appendMarkdown(
      `\n\n${definitions.length - displayedDefinitions.length} more matching record definitions omitted.`,
    );
  }

  return new vscode.Hover(markdown, range);
}

function createRecordLocationLink(definition, label) {
  return createFileLocationLink(definition.absolutePath, definition.line, label);
}

function createFileLocationLink(absolutePath, line, label) {
  const commandUri = buildOpenRecordCommandUri(absolutePath, line);
  return `[${escapeMarkdownLinkLabel(label)}](${commandUri})`;
}

function createProtocolFileLink(absolutePath, label) {
  return createFileLocationLink(absolutePath, 1, label);
}

function createMenuFieldChoiceLink(document, fieldDeclaration, choice) {
  const commandArguments = encodeURIComponent(
    JSON.stringify([
      {
        uri: document.uri.toString(),
        fieldName: fieldDeclaration.fieldName,
        start: fieldDeclaration.valueStart,
        end: fieldDeclaration.valueEnd,
        value: choice,
      },
    ]),
  );
  const commandUri = `command:${UPDATE_MENU_FIELD_VALUE_COMMAND}?${commandArguments}`;
  return `[${escapeMarkdownLinkLabel(choice)}](${commandUri})`;
}

function resolveMenuFieldValueRange(document, fieldName, start, end) {
  const declarations = extractFieldDeclarations(document.getText(), fieldName);
  const matchingDeclaration = declarations.find(
    (declaration) =>
      declaration.valueStart === start ||
      (start >= declaration.valueStart && start <= declaration.valueEnd),
  );
  if (matchingDeclaration) {
    return new vscode.Range(
      document.positionAt(matchingDeclaration.valueStart),
      document.positionAt(matchingDeclaration.valueEnd),
    );
  }

  return new vscode.Range(document.positionAt(start), document.positionAt(end));
}

function getRecordScopedFieldDeclarationAtPosition(snapshot, document, position) {
  if (!isDatabaseDocument(document)) {
    return undefined;
  }

  const text = document.getText();
  const offset = document.offsetAt(position);
  for (const recordDeclaration of extractRecordDeclarations(text)) {
    if (
      offset < recordDeclaration.recordStart ||
      offset > recordDeclaration.recordEnd
    ) {
      continue;
    }

    const fieldTypes = snapshot.fieldTypesByRecordType.get(recordDeclaration.recordType);
    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      if (
        offset >= fieldDeclaration.valueStart &&
        offset <= fieldDeclaration.valueEnd
      ) {
        return {
          ...fieldDeclaration,
          recordName: recordDeclaration.name,
          recordType: recordDeclaration.recordType,
          dbfType: fieldTypes?.get(fieldDeclaration.fieldName),
          range: new vscode.Range(
            document.positionAt(fieldDeclaration.valueStart),
            document.positionAt(fieldDeclaration.valueEnd),
          ),
        };
      }
    }
  }

  return undefined;
}

function isStreamDeviceRecordField(document, position) {
  const recordDeclaration = findEnclosingRecordDeclaration(
    document.getText(),
    document.offsetAt(position),
  );
  if (!recordDeclaration) {
    return false;
  }

  const dtypField = extractFieldDeclarationsInRecord(
    document.getText(),
    recordDeclaration,
    "DTYP",
  )[0];
  return dtypField?.value?.trim().toLowerCase() === "stream";
}

function getStreamProtocolReferenceAtPosition(fieldDeclaration, document, position) {
  const match = String(fieldDeclaration.value || "").match(/^\s*@([^\s"'`]+)/);
  if (!match) {
    return undefined;
  }

  const protocolPath = match[1];
  if (!protocolPath || containsEpicsMacroReference(protocolPath)) {
    return undefined;
  }

  const protocolStart =
    fieldDeclaration.valueStart + match[0].indexOf(protocolPath);
  const protocolEnd = protocolStart + protocolPath.length;
  const offset = document.offsetAt(position);
  if (offset < protocolStart || offset > protocolEnd) {
    return undefined;
  }

  return {
    protocolPath,
    range: new vscode.Range(
      document.positionAt(protocolStart),
      document.positionAt(protocolEnd),
    ),
  };
}

function resolveStreamProtocolFileReferences(snapshot, document, protocolPath) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (!project || !protocolPath) {
    return [];
  }

  const resolutions = [];
  const seen = new Set();
  for (const definition of collectStreamProtocolPathDefinitions(snapshot, project)) {
    for (const searchDirectory of definition.searchDirectories) {
      const absolutePath = normalizeFsPath(
        path.isAbsolute(protocolPath)
          ? protocolPath
          : path.resolve(searchDirectory, protocolPath),
      );
      if (!fs.existsSync(absolutePath)) {
        continue;
      }

      let stats;
      try {
        stats = fs.statSync(absolutePath);
      } catch (error) {
        continue;
      }

      if (!stats.isFile()) {
        continue;
      }

      const resolutionKey = `${definition.startupFileAbsolutePath}:${absolutePath}`;
      if (seen.has(resolutionKey)) {
        continue;
      }

      seen.add(resolutionKey);
      resolutions.push({
        absolutePath,
        protocolRelativePath: normalizePath(path.relative(project.rootPath, absolutePath)),
        startupFileAbsolutePath: definition.startupFileAbsolutePath,
        startupFileRelativePath: definition.startupFileRelativePath,
      });
    }
  }

  return resolutions.sort((left, right) =>
    compareLabels(
      `${left.protocolRelativePath} ${left.startupFileRelativePath}`,
      `${right.protocolRelativePath} ${right.startupFileRelativePath}`,
    ),
  );
}

function collectStreamProtocolPathDefinitions(snapshot, project) {
  const definitions = [];

  const startupFilePaths =
    Array.isArray(project.startupEntryPoints) && project.startupEntryPoints.length > 0
      ? project.startupEntryPoints
      : findProjectIocBootFilePaths(project.rootPath);

  for (const startupFilePath of startupFilePaths) {
    const text = readTextFile(startupFilePath);
    if (!text || !text.includes(STREAM_PROTOCOL_PATH_VARIABLE)) {
      continue;
    }

    const pseudoDocument = { uri: vscode.Uri.file(startupFilePath) };
    const state = createInitialStartupExecutionState(snapshot, pseudoDocument);

    for (const statement of extractStartupStatements(text)) {
      if (
        statement.kind === "envSet" &&
        statement.name === STREAM_PROTOCOL_PATH_VARIABLE
      ) {
        definitions.push({
          startupFileAbsolutePath: normalizeFsPath(startupFilePath),
          startupFileRelativePath: normalizePath(
            path.relative(project.rootPath, startupFilePath),
          ),
          searchDirectories: splitStreamProtocolSearchDirectories(
            expandStartupValue(statement.value, state.envVariables),
            state.currentDirectory || normalizeFsPath(path.dirname(startupFilePath)),
          ),
        });
      }

      applyStartupStatement(snapshot, pseudoDocument, statement, state);
    }
  }

  return definitions.filter(
    (definition) => definition.searchDirectories.length > 0,
  );
}

function splitStreamProtocolSearchDirectories(value, baseDirectory) {
  return String(value || "")
    .split(path.delimiter)
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) =>
      normalizeFsPath(
        path.isAbsolute(entry)
          ? entry
          : path.resolve(baseDirectory, entry),
      ),
    );
}

function findProjectIocBootFilePaths(rootPath) {
  const iocBootPath = normalizeFsPath(path.join(rootPath, "iocBoot"));
  if (!fs.existsSync(iocBootPath)) {
    return [];
  }

  const filePaths = [];
  const pendingDirectories = [iocBootPath];

  while (pendingDirectories.length > 0) {
    const directoryPath = pendingDirectories.pop();
    let entries;
    try {
      entries = fs.readdirSync(directoryPath, { withFileTypes: true });
    } catch (error) {
      continue;
    }

    for (const entry of entries) {
      const absolutePath = normalizeFsPath(path.join(directoryPath, entry.name));
      if (entry.isDirectory()) {
        pendingDirectories.push(absolutePath);
        continue;
      }

      if (entry.isFile()) {
        filePaths.push(absolutePath);
      }
    }
  }

  return filePaths.sort(compareLabels);
}

function resolveLinkedRecordName(snapshot, document, fieldValue) {
  for (const candidate of getPreferredLinkedRecordCandidateNames(document, fieldValue)) {
    if (getRecordDefinitionsForName(snapshot, document, candidate).length > 0) {
      return candidate;
    }
  }

  return undefined;
}

function getPreferredLinkedRecordCandidateNames(document, fieldValue) {
  const macroAssignments = isDatabaseDocument(document)
    ? extractDatabaseTocMacroAssignments(document.getText())
    : new Map();
  const candidateRanges = getLinkedRecordCandidateRanges(fieldValue, 0, {
    allowMacroReferences: true,
    macroAssignments,
  });
  const preferred = [];
  const fallback = [];
  const seen = new Set();

  for (const candidate of candidateRanges) {
    const name = String(candidate?.name || "").trim();
    if (!name || seen.has(name)) {
      continue;
    }
    seen.add(name);
    if (containsEpicsMacroReference(name)) {
      fallback.push(name);
    } else {
      preferred.push(name);
    }
  }

  return [...preferred, ...fallback];
}

function resolveLinkedRecordDefinitionMatch(snapshot, document, fieldValue) {
  for (const candidate of getPreferredLinkedRecordCandidateNames(document, fieldValue)) {
    const definitions = getRecordDefinitionsForName(snapshot, document, candidate);
    if (definitions.length > 0) {
      return {
        recordName: candidate,
        definitions,
      };
    }
  }

  return undefined;
}

function extractLinkedRecordCandidates(fieldValue) {
  if (!fieldValue) {
    return [];
  }

  const trimmedValue = fieldValue.trim();
  if (!trimmedValue || trimmedValue.startsWith("@")) {
    return [];
  }

  const firstToken = trimmedValue.split(/\s+/)[0];
  const candidates = [];
  const addCandidate = (candidate) => {
    const normalizedCandidate = normalizeLinkedRecordCandidate(candidate);
    if (!normalizedCandidate || candidates.includes(normalizedCandidate)) {
      return;
    }

    candidates.push(normalizedCandidate);
  };

  addCandidate(firstToken);

  const lastDotIndex = firstToken.lastIndexOf(".");
  if (lastDotIndex > 0) {
    const suffix = firstToken.slice(lastDotIndex + 1);
    if (/^[A-Z0-9_]+$/.test(suffix)) {
      addCandidate(firstToken.slice(0, lastDotIndex));
    }
  }

  return candidates;
}

function normalizeLinkedRecordCandidate(candidate) {
  return String(candidate || "")
    .trim()
    .replace(/[),;]+$/, "");
}

function resolvePvlistHoverRecordName(documentText, rawLine) {
  const expandedLine = expandPvlistHoverValue(
    rawLine,
    extractPvlistHoverMacroDefinitions(documentText),
    new Set(),
  );
  if (!expandedLine) {
    return undefined;
  }

  const channelText = stripRuntimeProtocolPrefix(expandedLine.trim());
  for (const candidate of extractLinkedRecordCandidates(channelText)) {
    if (candidate) {
      return candidate;
    }
  }

  return undefined;
}

function extractPvlistHoverMacroDefinitions(text) {
  const macroDefinitions = new Map();
  for (const rawLine of String(text || "").split(/\r?\n/)) {
    const trimmedLine = String(rawLine || "").trim();
    if (!trimmedLine || trimmedLine.startsWith("#")) {
      continue;
    }

    const macroMatch = trimmedLine.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!macroMatch) {
      continue;
    }

    const macroName = macroMatch[1];
    if (!macroDefinitions.has(macroName)) {
      macroDefinitions.set(macroName, macroMatch[2] || "");
    }
  }
  return macroDefinitions;
}

function expandPvlistHoverValue(text, macroDefinitions, stack) {
  let unresolved = false;
  const expanded = String(text || "").replace(
    /\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}/g,
    (match, parenName, defaultValue, braceName) => {
      const macroName = parenName || braceName;
      const resolved = resolvePvlistHoverMacroValue(
        macroName,
        defaultValue,
        macroDefinitions,
        stack,
      );
      if (resolved === undefined) {
        unresolved = true;
        return "";
      }
      return resolved;
    },
  );
  return unresolved ? undefined : expanded;
}

function resolvePvlistHoverMacroValue(
  macroName,
  defaultValue,
  macroDefinitions,
  stack,
) {
  if (stack.has(macroName)) {
    return undefined;
  }

  if (!macroDefinitions.has(macroName)) {
    return defaultValue;
  }

  const nextStack = new Set(stack);
  nextStack.add(macroName);
  return expandPvlistHoverValue(
    macroDefinitions.get(macroName),
    macroDefinitions,
    nextStack,
  );
}

function stripRuntimeProtocolPrefix(value) {
  const text = String(value || "");
  if (/^pva:\/\//i.test(text)) {
    return text.replace(/^pva:\/\//i, "");
  }
  if (/^ca:\/\//i.test(text)) {
    return text.replace(/^ca:\/\//i, "");
  }
  return text;
}

function getRecordDefinitionsForName(snapshot, document, recordName) {
  if (!recordName) {
    return [];
  }

  const definitions = [];
  const seen = new Set();
  const currentPath =
    document?.uri?.scheme === "file"
      ? normalizeFsPath(document.uri.fsPath)
      : undefined;
  const addDefinition = (definition) => {
    if (!definition?.absolutePath) {
      return;
    }
    const key = `${definition.absolutePath}:${definition.line}:${definition.name}`;
    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    definitions.push(definition);
  };

  const currentDefinitions = createRecordDefinitions(
    document.uri,
    document.getText(),
    extractRecordDeclarations(document.getText()).filter(
      (declaration) => getDatabaseRecordSearchNames(document.getText(), declaration.name).has(recordName),
    ),
  );

  for (const definition of currentDefinitions) {
    addDefinition(definition);
  }

  for (const definition of collectRecordDefinitionsFromSearchPaths(
    snapshot,
    document,
    recordName,
  )) {
    if (currentPath && definition.absolutePath === currentPath) {
      continue;
    }

    addDefinition(definition);
  }

  for (const definition of snapshot.recordDefinitionsByName.get(recordName) || []) {
    if (currentPath && definition.absolutePath === currentPath) {
      continue;
    }

    addDefinition(definition);
  }

  return definitions;
}

function collectRecordDefinitionsFromSearchPaths(snapshot, document, recordName) {
  if (!document?.uri || document.uri.scheme !== "file" || !recordName) {
    return [];
  }

  const searchDirectories = [];
  const seenDirectories = new Set();
  const addDirectory = (directoryPath, recursive) => {
    const normalizedPath = normalizeFsPath(directoryPath);
    if (!normalizedPath || seenDirectories.has(`${normalizedPath}:${recursive ? 1 : 0}`)) {
      return;
    }
    if (!fs.existsSync(normalizedPath)) {
      return;
    }

    let stats;
    try {
      stats = fs.statSync(normalizedPath);
    } catch (error) {
      return;
    }

    if (!stats.isDirectory()) {
      return;
    }

    seenDirectories.add(`${normalizedPath}:${recursive ? 1 : 0}`);
    searchDirectories.push({ directoryPath: normalizedPath, recursive });
  };

  addDirectory(path.dirname(document.uri.fsPath), false);

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (project?.rootPath) {
    addDirectory(path.join(project.rootPath, "db"), true);
    addDirectory(path.join(project.rootPath, "Db"), true);

    for (const releaseRoot of resolveReleaseModuleRoots(
      project.rootPath,
      project.releaseVariables,
    )) {
      addDirectory(path.join(releaseRoot.rootPath, "db"), true);
      addDirectory(path.join(releaseRoot.rootPath, "Db"), true);
    }
  }

  const definitions = [];
  for (const searchDirectory of searchDirectories) {
    for (const filePath of collectDatabaseFilePaths(searchDirectory.directoryPath, searchDirectory.recursive)) {
      if (normalizeFsPath(filePath) === normalizeFsPath(document.uri.fsPath)) {
        continue;
      }

      const text = readTextFile(filePath);
      if (text === undefined) {
        continue;
      }

      const matchingDeclarations = extractRecordDeclarations(text).filter(
        (declaration) => getDatabaseRecordSearchNames(text, declaration.name).has(recordName),
      );
      if (!matchingDeclarations.length) {
        continue;
      }

      definitions.push(
        ...createRecordDefinitions(vscode.Uri.file(filePath), text, matchingDeclarations),
      );
    }
  }

  return definitions;
}

function collectDatabaseFilePaths(directoryPath, recursive) {
  const normalizedDirectory = normalizeFsPath(directoryPath);
  if (!normalizedDirectory || !fs.existsSync(normalizedDirectory)) {
    return [];
  }

  const results = [];
  const visitDirectory = (currentDirectory) => {
    let entries;
    try {
      entries = fs.readdirSync(currentDirectory, { withFileTypes: true });
    } catch (error) {
      return;
    }

    entries.sort((left, right) => compareLabels(left.name, right.name));
    for (const entry of entries) {
      const entryPath = normalizeFsPath(path.join(currentDirectory, entry.name));
      if (entry.isDirectory()) {
        if (recursive) {
          visitDirectory(entryPath);
        }
        continue;
      }

      if (DATABASE_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
        results.push(entryPath);
      }
    }
  };

  visitDirectory(normalizedDirectory);
  return results;
}

function getVariableReferenceAtPosition(document, position) {
  const lineText = document.lineAt(position.line).text;
  const regex = isMakefileDocument(document)
    ? /\$\(([A-Za-z_][A-Za-z0-9_.-]*)\)|\$\{([A-Za-z_][A-Za-z0-9_.-]*)\}|\$([A-Za-z_][A-Za-z0-9_.-]*)/g
    : /\$\(([A-Za-z_][A-Za-z0-9_]*)(?:=[^)]*)?\)|\$\{([A-Za-z_][A-Za-z0-9_]*)(?:=[^}]*)?\}|\$([A-Za-z_][A-Za-z0-9_]*)/g;
  let match;

  while ((match = regex.exec(lineText))) {
    const matchStart = match.index;
    const matchEnd = match.index + match[0].length;
    if (position.character < matchStart || position.character >= matchEnd) {
      continue;
    }

    return {
      variableName: match[1] || match[2] || match[3],
      range: new vscode.Range(
        position.line,
        matchStart,
        position.line,
        matchEnd,
      ),
    };
  }

  return undefined;
}

function resolveMakefileVariableValue(snapshot, document, variableName) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  const assignments = parseMakeAssignments(document.getText());
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const releaseVariables = project ? project.releaseVariables : new Map();
  const baseDirectory = normalizeFsPath(path.dirname(document.uri.fsPath));
  const cache = new Map();
  const resolving = new Set();

  const resolveVariable = (name) => {
    if (cache.has(name)) {
      return cache.get(name);
    }

    if (resolving.has(name)) {
      return undefined;
    }

    let rawValue;
    if (assignments.has(name)) {
      rawValue = assignments.get(name).join(" ");
    } else if (releaseVariables.has(name)) {
      rawValue = releaseVariables.get(name);
    } else if (process.env[name]) {
      rawValue = process.env[name];
    } else {
      return undefined;
    }

    resolving.add(name);
    const expandedValue = rawValue.replace(
      /\$\(([^)]+)\)|\$\{([^}]+)\}|\$([A-Za-z_][A-Za-z0-9_.-]*)/g,
      (_, parenthesizedName, bracedName, bareName) => {
        const nestedName = parenthesizedName || bracedName || bareName;
        return resolveVariable(nestedName)?.resolvedValue || `$(${nestedName})`;
      },
    );
    resolving.delete(name);

    const absolutePath = computeAbsoluteVariablePath(expandedValue, baseDirectory);
    const result = {
      resolvedValue: absolutePath || expandedValue,
      absolutePath,
    };
    cache.set(name, result);
    return result;
  };

  return resolveVariable(variableName);
}

function getStartupNavigationTarget(snapshot, document, position) {
  const statement = getStartupStatementAtPosition(document, position);
  if (!statement) {
    return undefined;
  }

  const state = createStartupExecutionState(snapshot, document, statement.start);
  const resolution = resolveStartupPath(snapshot, document, statement, state);
  if (!resolution || resolution.isDirectory) {
    return undefined;
  }

  return getNavigationTargetFromStartupResolution(resolution);
}

function getSubstitutionNavigationTarget(snapshot, document, position) {
  const reference = getSubstitutionTemplateReferenceAtPosition(document, position);
  if (!reference) {
    return undefined;
  }

  const absolutePath = resolveSubstitutionTemplateAbsolutePathForDocument(
    snapshot,
    document,
    reference.templatePath,
  );
  if (!absolutePath) {
    return undefined;
  }

  return { absolutePath };
}

function getMakefileNavigationTarget(snapshot, document, position) {
  const reference = getMakefileReferenceAtPosition(document, position);
  if (!reference) {
    return undefined;
  }

  if (reference.kind === "dbFile") {
    return getMakefileDatabaseReferenceTarget(snapshot, document, reference);
  }

  const localTarget = getLocalMakefileNavigationTarget(document, reference);
  if (localTarget) {
    return localTarget;
  }

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (!project || !isProjectResourceMakefileReference(reference)) {
    return undefined;
  }

  const entry =
    reference.kind === "dbd"
      ? project.availableDbds.get(reference.name)
      : project.availableLibs.get(reference.name);
  if (!entry || !entry.absolutePath || !fs.existsSync(entry.absolutePath)) {
    return undefined;
  }

  return {
    absolutePath: entry.absolutePath,
  };
}

function getMacroCompletionLabels(snapshot, document, position) {
  const labels = new Set(snapshot.macros);

  if (document.languageId === LANGUAGE_IDS.startup) {
    const state = createStartupExecutionState(
      snapshot,
      document,
      document.offsetAt(position),
    );
    for (const variableName of state.envVariables.keys()) {
      labels.add(variableName);
    }
  }

  return [...labels].sort(compareLabels);
}

function getStartupLoadedRecordNames(snapshot, document, position) {
  return [...getStartupLoadedRecordDefinitionMap(snapshot, document, position).keys()].sort(
    compareLabels,
  );
}

function getStartupLoadedRecordDefinitionMap(snapshot, document, position) {
  if (!isStartupDocument(document)) {
    return new Map();
  }

  const definitionsByName = new Map();
  const untilOffset = document.offsetAt(position);
  const state = createInitialStartupExecutionState(snapshot, document);

  for (const statement of extractStartupStatements(document.getText())) {
    if (statement.start >= untilOffset) {
      break;
    }

    if (
      statement.kind === "load" &&
      ["dbLoadRecords", "dbLoadTemplate"].includes(statement.command)
    ) {
      for (const definition of getLoadedRecordDefinitionsForStartupStatement(
        snapshot,
        document,
        statement,
        state,
      )) {
        addToMapOfArrays(definitionsByName, definition.name, definition);
      }
    }

    applyStartupStatement(snapshot, document, statement, state);
  }

  return definitionsByName;
}

function getLoadedRecordDefinitionsForStartupStatement(snapshot, document, statement, state) {
  const resolution = resolveStartupPath(snapshot, document, statement, state);
  if (!resolution) {
    return [];
  }

  switch (statement.command) {
    case "dbLoadRecords":
      return getLoadedRecordDefinitionsFromDatabaseFile(
        document,
        statement.path,
        resolution,
        parseStartupLoadMacroAssignments(statement.macros, state?.envVariables),
        state?.envVariables,
      );

    case "dbLoadTemplate":
      return getLoadedRecordDefinitionsFromSubstitutionFile(
        snapshot,
        document,
        resolution,
        state,
      );

    default:
      return [];
  }
}

function getLoadedRecordDefinitionsFromDatabaseFile(
  document,
  displayPath,
  resolution,
  loadMacros,
  envVariables,
) {
  const resolvedFile = getReadableStartupFileResolution(document, resolution);
  if (!resolvedFile?.text) {
    return [];
  }

  return createStartupLoadedRecordDefinitions(
    document,
    resolvedFile,
    extractRecordDeclarations(resolvedFile.text),
    [loadMacros, envVariables, process.env],
    displayPath,
  );
}

function getLoadedRecordDefinitionsFromSubstitutionFile(
  snapshot,
  document,
  resolution,
  state,
) {
  const resolvedFile = getReadableStartupFileResolution(document, resolution);
  if (!resolvedFile?.text) {
    return [];
  }

  const definitions = [];
  for (const load of parseSubstitutionLoads(resolvedFile.text)) {
    const templateAbsolutePath = resolveSubstitutionTemplateAbsolutePath(
      snapshot,
      document,
      state,
      resolvedFile.absolutePath,
      load.templatePath,
    );
    if (!templateAbsolutePath) {
      continue;
    }

    const templateText = readTextFile(templateAbsolutePath);
    if (templateText === undefined) {
      continue;
    }

    const rows = load.rows.length > 0 ? load.rows : [new Map()];
    const declarations = extractRecordDeclarations(templateText);
    for (const rowMacros of rows) {
      definitions.push(
        ...createStartupLoadedRecordDefinitions(
          document,
          {
            absolutePath: templateAbsolutePath,
            text: templateText,
          },
          declarations,
          [rowMacros, state?.envVariables, process.env],
          load.templatePath,
        ),
      );
    }
  }

  return definitions;
}

function createStartupLoadedRecordDefinitions(
  startupDocument,
  resolvedFile,
  declarations,
  macroSources,
  displayPath,
) {
  const definitions = [];

  for (const declaration of declarations) {
    const expandedName = expandEpicsValue(declaration.name, macroSources);
    if (!expandedName) {
      continue;
    }

    definitions.push({
      name: expandedName,
      recordType: declaration.recordType,
      absolutePath: resolvedFile.absolutePath,
      relativePath:
        displayPath || getRelativePathFromDocument(startupDocument, resolvedFile.absolutePath),
      line: getLineNumberAtOffset(resolvedFile.text, declaration.recordStart),
      preview: buildStartupLoadedRecordPreview(
        resolvedFile.text,
        declaration,
        expandedName,
      ),
    });
  }

  return definitions;
}

function buildStartupLoadedRecordPreview(text, declaration, expandedName) {
  const preview = buildRecordPreview(text, declaration);
  const escapedRecordName = escapeDoubleQuotedString(expandedName);

  return preview.replace(
    /record\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"/,
    `record($1, "${escapedRecordName}"`,
  );
}

function getReadableStartupFileResolution(document, resolution) {
  const navigationTarget = getNavigationTargetFromStartupResolution(resolution);
  if (!navigationTarget?.absolutePath) {
    return undefined;
  }

  const absolutePath = normalizeFsPath(navigationTarget.absolutePath);
  const text = readTextFile(absolutePath);
  if (text === undefined) {
    return undefined;
  }

  return {
    absolutePath,
    text,
  };
}

function getReadableAbsolutePathForArtifact(artifact) {
  if (!artifact) {
    return undefined;
  }

  if (artifact.absoluteRuntimePath && isExistingFile(artifact.absoluteRuntimePath)) {
    return normalizeFsPath(artifact.absoluteRuntimePath);
  }

  if (!artifact.sourceRelativePath) {
    return undefined;
  }

  const sourcePath = normalizeFsPath(
    path.join(findRootPathForArtifact(artifact), artifact.sourceRelativePath),
  );
  return isExistingFile(sourcePath) ? sourcePath : undefined;
}

function resolveSubstitutionTemplateAbsolutePath(
  snapshot,
  document,
  state,
  substitutionFilePath,
  templatePath,
) {
  const expandedTemplatePath = expandEpicsValue(templatePath, [
    state?.envVariables,
    process.env,
  ]);
  const syntheticLoad = {
    command: "dbLoadRecords",
    path: expandedTemplatePath,
  };
  const startupResolution = resolveStartupPath(snapshot, document, syntheticLoad, state);
  const startupTarget = startupResolution
    ? getNavigationTargetFromStartupResolution(startupResolution)
    : undefined;
  if (startupTarget?.absolutePath) {
    return normalizeFsPath(startupTarget.absolutePath);
  }

  if (!substitutionFilePath) {
    return undefined;
  }

  const fallbackPath = normalizeFsPath(
    path.resolve(path.dirname(substitutionFilePath), expandedTemplatePath),
  );
  return readTextFile(fallbackPath) === undefined ? undefined : fallbackPath;
}

function resolveSubstitutionTemplateAbsolutePathForDocument(
  snapshot,
  document,
  templatePath,
) {
  return resolveSubstitutionTemplateAbsolutePathsForDocument(
    snapshot,
    document,
    templatePath,
  )[0];
}

function resolveSubstitutionTemplateAbsolutePathsForDocument(
  snapshot,
  document,
  templatePath,
) {
  if (!document?.uri || document.uri.scheme !== "file" || !templatePath) {
    return [];
  }

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const releaseVariables = project ? project.releaseVariables : new Map();
  const expandedTemplatePath = expandEpicsValue(templatePath, [
    releaseVariables,
    process.env,
  ]);
  if (!expandedTemplatePath) {
    return [];
  }

  const absolutePaths = [];
  const seen = new Set();
  const addCandidatePath = (candidatePath) => {
    if (!candidatePath) {
      return;
    }

    const normalizedPath = normalizeFsPath(candidatePath);
    if (seen.has(normalizedPath)) {
      return;
    }

    try {
      const stats = fs.statSync(normalizedPath);
      if (!stats.isFile()) {
        return;
      }
    } catch (error) {
      return;
    }

    seen.add(normalizedPath);
    absolutePaths.push(normalizedPath);
  };

  addCandidatePath(
    path.resolve(path.dirname(document.uri.fsPath), expandedTemplatePath),
  );

  if (project?.rootPath) {
    for (const candidatePath of getReleaseTemplateCandidatePaths(
      project.rootPath,
      expandedTemplatePath,
    )) {
      addCandidatePath(candidatePath);
    }
  }

  for (const releaseRoot of getProjectReleaseSearchRoots(project)) {
    for (const candidatePath of getReleaseTemplateCandidatePaths(
      releaseRoot,
      expandedTemplatePath,
    )) {
      addCandidatePath(candidatePath);
    }
  }

  return absolutePaths;
}

function getProjectReleaseSearchRoots(project) {
  if (!project) {
    return [];
  }

  const roots = [];
  const seen = new Set();

  for (const rawValue of project.releaseVariables.values()) {
    const expandedValue = expandEpicsValue(rawValue, [
      project.releaseVariables,
      process.env,
    ]);
    const absolutePath = computeAbsoluteVariablePath(expandedValue, project.rootPath);
    if (!absolutePath || seen.has(absolutePath) || !fs.existsSync(absolutePath)) {
      continue;
    }

    let stats;
    try {
      stats = fs.statSync(absolutePath);
    } catch (error) {
      continue;
    }

    if (!stats.isDirectory()) {
      continue;
    }

    seen.add(absolutePath);
    roots.push(absolutePath);
  }

  return roots;
}

function getReleaseTemplateCandidatePaths(releaseRoot, templatePath) {
  const normalizedTemplatePath = normalizePath(templatePath);
  const candidates = [];

  if (/^(?:db|Db)\//.test(normalizedTemplatePath)) {
    candidates.push(normalizeFsPath(path.resolve(releaseRoot, normalizedTemplatePath)));
  } else {
    candidates.push(normalizeFsPath(path.resolve(releaseRoot, "db", normalizedTemplatePath)));
    candidates.push(normalizeFsPath(path.resolve(releaseRoot, "Db", normalizedTemplatePath)));
  }

  return [...new Set(candidates)];
}

function parseStartupLoadMacroAssignments(rawAssignments, envVariables) {
  const assignments = extractNamedAssignments(rawAssignments);
  const values = new Map();

  for (const [name, rawValue] of assignments.entries()) {
    values.set(name, expandEpicsValue(rawValue, [envVariables, process.env]));
  }

  return values;
}

function getFieldNamesForRecordType(snapshot, recordType) {
  if (recordType) {
    if (snapshot.fieldsByRecordType.has(recordType)) {
      return [...snapshot.fieldsByRecordType.get(recordType)].sort(compareLabels);
    }
  }

  const fallbackLabels = new Set(COMMON_RECORD_FIELDS);
  for (const fieldName of snapshot.allFields) {
    fallbackLabels.add(fieldName);
  }

  return [...fallbackLabels].sort(compareLabels);
}

function getRuntimeProbeFieldNamesForRecordType(recordType) {
  const labels = [];
  const seen = new Set();
  const addFieldName = (fieldName) => {
    const normalizedFieldName = String(fieldName || "");
    if (!normalizedFieldName || seen.has(normalizedFieldName)) {
      return;
    }
    seen.add(normalizedFieldName);
    labels.push(normalizedFieldName);
  };

  if (recordType) {
    for (const fieldName of recordTemplateStaticData.fieldOrderByRecordType.get(recordType) || []) {
      addFieldName(fieldName);
    }
    for (const fieldName of recordTemplateFields.get(recordType) || []) {
      addFieldName(fieldName);
    }
    for (const fieldName of recordTemplateStaticData.fieldTypesByRecordType.get(recordType)?.keys() || []) {
      addFieldName(fieldName);
    }
  }

  if (labels.length === 0) {
    for (const fieldName of COMMON_RECORD_FIELDS) {
      addFieldName(fieldName);
    }
  }

  return labels;
}

function getRuntimeProbeFieldTypeForRecordType(recordType, fieldName) {
  if (!recordType || !fieldName) {
    return undefined;
  }

  return recordTemplateStaticData.fieldTypesByRecordType
    .get(recordType)
    ?.get(fieldName);
}

function getAvailableFieldNamesForRecordInstance(
  snapshot,
  document,
  position,
  recordType,
) {
  const fieldNames = getFieldNamesForRecordType(snapshot, recordType);
  if (!document || !recordType) {
    return fieldNames;
  }

  const recordDeclaration = findEnclosingRecordDeclaration(
    document.getText(),
    document.offsetAt(position),
  );
  if (!recordDeclaration) {
    return fieldNames;
  }

  const existingFields = new Set(
    extractFieldDeclarationsInRecord(document.getText(), recordDeclaration).map(
      (fieldDeclaration) => fieldDeclaration.fieldName,
    ),
  );

  return fieldNames.filter((fieldName) => !existingFields.has(fieldName));
}

function getFieldTailInsertionRange(document, position) {
  const lineText = document.lineAt(position.line).text;
  let endCharacter = position.character;

  while (
    endCharacter < lineText.length &&
    /[A-Za-z0-9_]/.test(lineText[endCharacter])
  ) {
    endCharacter += 1;
  }

  if (lineText[endCharacter] === "\"") {
    endCharacter += 1;
  }

  if (lineText[endCharacter] === ")") {
    endCharacter += 1;
  }

  return new vscode.Range(
    position.line,
    position.character,
    position.line,
    endCharacter,
  );
}

function getDbLoadRecordsTailInsertionRange(document, position) {
  const lineText = document.lineAt(position.line).text;
  const lineSuffix = lineText.slice(position.character);
  const suffixMatch = lineSuffix.match(/^"\s*(?:,\s*"[^"\n]*")?\s*\)/);

  if (suffixMatch) {
    return new vscode.Range(
      position.line,
      position.character,
      position.line,
      position.character + suffixMatch[0].length,
    );
  }

  let endCharacter = position.character;
  if (lineText[endCharacter] === "\"") {
    endCharacter += 1;
  }

  while (endCharacter < lineText.length && /\s/.test(lineText[endCharacter])) {
    endCharacter += 1;
  }

  if (lineText[endCharacter] === ")") {
    endCharacter += 1;
  }

  return new vscode.Range(
    position.line,
    position.character,
    position.line,
    endCharacter,
  );
}

function getDbLoadRecordsMacroValueInsertionRange(document, range) {
  const lineText = document.lineAt(range.end.line).text;
  const lineSuffix = lineText.slice(range.end.character);
  const suffixMatch = lineSuffix.match(/^"\s*\)/);
  const endCharacter = suffixMatch
    ? range.end.character + suffixMatch[0].length
    : range.end.character;

  return new vscode.Range(
    range.start.line,
    range.start.character,
    range.end.line,
    endCharacter,
  );
}

function buildFieldCompletionTailSnippet(snapshot, recordType, fieldName) {
  const defaultValue = getDefaultFieldValue(snapshot, recordType, fieldName);
  const valuePlaceholder = buildSnippetPlaceholder(1, defaultValue);

  return `, "${valuePlaceholder}")`;
}

function buildDbLoadRecordsMacroAssignmentsLabel(macroNames) {
  const names = Array.isArray(macroNames)
    ? macroNames.filter(Boolean)
    : [];
  return names.map((macroName) => `${macroName}=`).join(",");
}

function buildDbLoadRecordsMacroAssignmentsSnippet(macroNames) {
  const names = Array.isArray(macroNames)
    ? macroNames.filter(Boolean)
    : [];
  return names
    .map((macroName, index) => `${macroName}=${buildSnippetPlaceholder(index + 1, "")}`)
    .join(",");
}

function buildDbLoadRecordsCompletionTailSnippet(macroNames) {
  const assignmentSnippet = buildDbLoadRecordsMacroAssignmentsSnippet(macroNames);
  if (!assignmentSnippet) {
    return `")${buildSnippetPlaceholder(0, "")}`;
  }
  return `", "${assignmentSnippet}")${buildSnippetPlaceholder(0, "")}`;
}

function buildFieldCompletionDetail(snapshot, recordType, fieldName) {
  const dbfType = snapshot.fieldTypesByRecordType.get(recordType)?.get(fieldName);
  if (dbfType && recordType) {
    return `${fieldName} for ${recordType} (${dbfType})`;
  }

  if (recordType) {
    return `Field for ${recordType}`;
  }

  return "EPICS field";
}

function getFieldInitialValue(snapshot, recordType, fieldName) {
  return getDefaultFieldValue(snapshot, recordType, fieldName);
}

function getDefaultFieldValue(snapshot, recordType, fieldName) {
  const dbfType = snapshot.fieldTypesByRecordType.get(recordType)?.get(fieldName);
  const explicitInitialValue = snapshot.fieldInitialValuesByRecordType
    .get(recordType)
    ?.get(fieldName);
  if (explicitInitialValue !== undefined) {
    return resolveExplicitFieldInitialValue(
      snapshot,
      recordType,
      fieldName,
      dbfType,
      explicitInitialValue,
    );
  }

  if (!dbfType) {
    return "";
  }

  if (NUMERIC_DBF_TYPES.has(dbfType)) {
    return "0";
  }

  if (EMPTY_DEFAULT_DBF_TYPES.has(dbfType)) {
    return "";
  }

  if (dbfType === "DBF_MENU") {
    return getMenuFieldChoices(snapshot, recordType, fieldName)[0] || "";
  }

  return "";
}

function resolveExplicitFieldInitialValue(
  snapshot,
  recordType,
  fieldName,
  dbfType,
  explicitInitialValue,
) {
  if (dbfType !== "DBF_MENU") {
    return explicitInitialValue;
  }

  const choices = getMenuFieldChoices(snapshot, recordType, fieldName);
  if (choices.length === 0) {
    return explicitInitialValue;
  }

  if (choices.includes(explicitInitialValue)) {
    return explicitInitialValue;
  }

  const trimmedValue = String(explicitInitialValue).trim();
  if (/^\d+$/.test(trimmedValue)) {
    const choiceIndex = Number(trimmedValue);
    if (choiceIndex >= 0 && choiceIndex < choices.length) {
      return choices[choiceIndex];
    }
  }

  return explicitInitialValue;
}

function buildSnippetPlaceholder(index, defaultValue) {
  if (defaultValue === "") {
    return `\${${index}}`;
  }

  return `\${${index}:${escapeSnippetPlaceholderValue(defaultValue)}}`;
}

function escapeSnippetPlaceholderValue(value) {
  return String(value).replace(/[$}\\]/g, "\\$&");
}

function getFieldValueLabels(snapshot, context) {
  const labels = new Set();
  const menuChoices = getMenuFieldChoices(
    snapshot,
    context.recordType,
    context.fieldName,
  );

  if (menuChoices.length > 0) {
    for (const choice of menuChoices) {
      labels.add(choice);
    }
  } else {
    for (const label of STATIC_FIELD_VALUE_ENUMS[context.fieldName] || []) {
      labels.add(label);
    }

    const observedValues = snapshot.fieldValuesByField.get(context.fieldName);
    if (observedValues) {
      for (const value of observedValues) {
        labels.add(value);
      }
    }
  }

  return [...labels]
    .filter((label) => matchesCompletionQuery(label, context.partial))
    .sort(compareLabels);
}

function getMenuFieldChoices(snapshot, recordType, fieldName) {
  if (!recordType || !fieldName) {
    return [];
  }

  const fieldMenus = snapshot.fieldMenuChoicesByRecordType.get(recordType);
  if (!fieldMenus) {
    return [];
  }

  return fieldMenus.get(fieldName) || [];
}

function buildLinkTargetItems(snapshot, context) {
  const items = [];

  for (const [recordName, recordType] of snapshot.recordsByName.entries()) {
    if (!matchesCompletionQuery(recordName, context.partial)) {
      continue;
    }

    const item = new vscode.CompletionItem(recordName, vscode.CompletionItemKind.Reference);
    item.range = context.range;
    item.insertText = recordName;
    item.detail = "EPICS link target";
    item.documentation = recordType
      ? `EPICS record target of type \`${recordType}\``
      : "EPICS record target";
    item.filterText = buildFilterText(recordName);
    item.sortText = buildSortText(recordName, context.partial);
    item.preselect = true;
    items.push(item);
  }

  items.sort((left, right) => compareLabels(left.label, right.label));
  return items;
}

function getAllowedExtensionsForFileContext(fileKind) {
  switch (fileKind) {
    case "startupDirectory":
      return new Set();
    case "dbLoadDatabase":
      return DBD_EXTENSIONS;
    case "dbLoadRecords":
    case "databaseInclude":
      return DATABASE_EXTENSIONS;
    case "dbLoadTemplate":
      return SUBSTITUTION_EXTENSIONS;
    case "dbdInclude":
      return DBD_EXTENSIONS;
    case "startupInclude":
    default:
      return new Set([
        ...DATABASE_EXTENSIONS,
        ...SUBSTITUTION_EXTENSIONS,
        ...STARTUP_EXTENSIONS,
        ...DBD_EXTENSIONS,
      ]);
  }
}

function getInsertPathForDocument(document, fileEntry) {
  if (document.uri.scheme !== "file") {
    return fileEntry.relativePath;
  }

  const sourceFolder = vscode.workspace.getWorkspaceFolder(document.uri);
  const targetFolder = vscode.workspace.getWorkspaceFolder(fileEntry.uri);

  if (
    sourceFolder &&
    targetFolder &&
    sourceFolder.uri.toString() === targetFolder.uri.toString()
  ) {
    const workspaceRoot = sourceFolder.uri.fsPath;
    const sourceDir = normalizePath(
      path.relative(workspaceRoot, path.dirname(document.uri.fsPath)),
    );
    const targetPath = normalizePath(
      path.relative(workspaceRoot, fileEntry.uri.fsPath),
    );
    const relativePath = path.posix.relative(sourceDir || ".", targetPath);
    return relativePath || path.posix.basename(targetPath);
  }

  return fileEntry.relativePath;
}

function getFilesystemPathEntries(startupState, context) {
  if (!startupState || !startupState.currentDirectory || !documentPathKindUsesFilesystem(context.fileKind)) {
    return [];
  }

  const expandedPartial = expandStartupValue(context.partial || "", startupState.envVariables);
  if (containsMakeVariableReference(expandedPartial)) {
    return [];
  }

  const normalizedPartial = normalizePath(expandedPartial || "");
  const hasDirectorySeparator = /[\\/]/.test(normalizedPartial);
  const hasTrailingSeparator = /[\\/]$/.test(normalizedPartial);
  const relativeDirectory = hasDirectorySeparator
    ? hasTrailingSeparator
      ? normalizedPartial.replace(/[\\/]+$/, "")
      : path.posix.dirname(normalizedPartial)
    : ".";
  const namePrefix = hasTrailingSeparator
    ? ""
    : hasDirectorySeparator
      ? path.posix.basename(normalizedPartial)
      : normalizedPartial;
  const absoluteDirectory = normalizeFsPath(
    path.resolve(
      startupState.currentDirectory,
      relativeDirectory === "." ? "" : relativeDirectory,
    ),
  );

  if (!fs.existsSync(absoluteDirectory)) {
    return [];
  }

  let directoryEntries;
  try {
    directoryEntries = fs.readdirSync(absoluteDirectory, { withFileTypes: true });
  } catch (error) {
    return [];
  }

  const allowedExtensions = getAllowedExtensionsForFileContext(context.fileKind);
  const items = [];

  for (const directoryEntry of directoryEntries) {
    if (
      namePrefix &&
      !matchesCompletionQuery(directoryEntry.name, namePrefix)
    ) {
      continue;
    }

    const absoluteEntryPath = path.join(absoluteDirectory, directoryEntry.name);
    const relativeEntryPath = getRelativePathFromBaseDirectory(
      startupState.currentDirectory,
      absoluteEntryPath,
    );

    if (directoryEntry.isDirectory()) {
      items.push({
        insertPath: `${normalizePath(relativeEntryPath)}/`,
        absolutePath: normalizeFsPath(absoluteEntryPath),
        kind: vscode.CompletionItemKind.Folder,
        detail: normalizePath(relativeEntryPath),
        documentation: `Directory from ${startupState.currentDirectory}`,
      });
      continue;
    }

    if (context.fileKind === "startupDirectory") {
      continue;
    }

    if (context.fileKind === "startupInclude") {
      items.push({
        insertPath: normalizePath(relativeEntryPath),
        absolutePath: normalizeFsPath(absoluteEntryPath),
        kind: vscode.CompletionItemKind.File,
        detail: normalizePath(relativeEntryPath),
        documentation: `File from ${startupState.currentDirectory}`,
      });
      continue;
    }

    const extension = path.extname(directoryEntry.name).toLowerCase();
    if (!allowedExtensions.has(extension)) {
      continue;
    }

    items.push({
      insertPath: normalizePath(relativeEntryPath),
      absolutePath: normalizeFsPath(absoluteEntryPath),
      kind: vscode.CompletionItemKind.File,
      detail: normalizePath(relativeEntryPath),
      documentation: `File from ${startupState.currentDirectory}`,
    });
  }

  return items;
}

function getDbLoadRecordsFilesystemPathItems(startupState, context) {
  if (!startupState?.currentDirectory) {
    return [];
  }

  const allowedExtensions = getAllowedExtensionsForFileContext("dbLoadRecords");
  const currentDirectory = normalizeFsPath(startupState.currentDirectory);
  const dbDirectory = normalizeFsPath(path.join(currentDirectory, "db"));
  const partialText = String(context?.partial || "");
  const normalizedPartial = normalizePath(partialText);
  const namePrefix = normalizedPartial.includes("/")
    ? path.posix.basename(normalizedPartial)
    : normalizedPartial;
  const items = [];
  const seenInsertPaths = new Set();

  const collectFiles = (directoryPath, locationLabel) => {
    if (!directoryPath || !fs.existsSync(directoryPath)) {
      return;
    }

    let stats;
    try {
      stats = fs.statSync(directoryPath);
    } catch (error) {
      return;
    }
    if (!stats.isDirectory()) {
      return;
    }

    let directoryEntries;
    try {
      directoryEntries = fs.readdirSync(directoryPath, { withFileTypes: true });
    } catch (error) {
      return;
    }

    directoryEntries.sort((left, right) => compareLabels(left.name, right.name));
    for (const directoryEntry of directoryEntries) {
      if (!directoryEntry.isFile()) {
        continue;
      }

      const extension = path.extname(directoryEntry.name).toLowerCase();
      if (!allowedExtensions.has(extension)) {
        continue;
      }

      if (namePrefix && !matchesCompletionQuery(directoryEntry.name, namePrefix)) {
        continue;
      }

      const absoluteEntryPath = normalizeFsPath(path.join(directoryPath, directoryEntry.name));
      const insertPath = normalizePath(
        getRelativePathFromBaseDirectory(currentDirectory, absoluteEntryPath),
      );
      if (!insertPath || seenInsertPaths.has(insertPath)) {
        continue;
      }
      seenInsertPaths.add(insertPath);

      items.push({
        label: directoryEntry.name,
        insertPath,
        absolutePath: absoluteEntryPath,
        detail: locationLabel,
        documentation: `File from ${locationLabel}`,
      });
    }
  };

  collectFiles(currentDirectory, currentDirectory);
  collectFiles(dbDirectory, dbDirectory);

  return items.map((entry) => {
    const item = new vscode.CompletionItem(
      entry.label,
      vscode.CompletionItemKind.File,
    );
    item.range = context.range;
    item.insertText = entry.insertPath;
    item.detail = entry.detail;
    item.documentation = entry.documentation;
    item.filterText = buildFilterText(entry.label);
    item.sortText = `0-${entry.label}`;
    applyDbLoadRecordsPathCompletion(item, context, entry.absolutePath);
    return item;
  });
}

function documentPathKindUsesFilesystem(fileKind) {
  return [
    "startupInclude",
    "startupDirectory",
    "dbLoadDatabase",
    "dbLoadRecords",
    "dbLoadTemplate",
  ].includes(fileKind);
}

function createEmptySnapshot(staticData) {
  return {
    recordTypes: new Set(staticData.recordTypes),
    allFields: new Set(staticData.allFields),
    fieldsByRecordType: cloneMapOfSets(staticData.fieldsByRecordType),
    fieldTypesByRecordType: cloneMapOfMaps(staticData.fieldTypesByRecordType),
    fieldMenuChoicesByRecordType: cloneMapOfMaps(
      staticData.fieldMenuChoicesByRecordType,
    ),
    fieldInitialValuesByRecordType: cloneMapOfMaps(
      staticData.fieldInitialValuesByRecordType,
    ),
    deviceSupportDefinitionsByName: new Map(),
    driverDefinitionsByName: new Map(),
    registrarDefinitionsByName: new Map(),
    functionDefinitionsByName: new Map(),
    variableDefinitionsByName: new Map(),
    recordsByName: new Map(),
    recordDefinitionsByName: new Map(),
    macros: new Set(),
    fieldValuesByField: new Map(),
    workspaceFiles: [],
    workspaceFilesByAbsolutePath: new Map(),
    projectModel: createEmptyProjectModel(),
  };
}

function mergeSnapshotWithDocument(snapshot, document) {
  const merged = {
    recordTypes: new Set(snapshot.recordTypes),
    allFields: new Set(snapshot.allFields),
    fieldsByRecordType: cloneMapOfSets(snapshot.fieldsByRecordType),
    fieldTypesByRecordType: cloneMapOfMaps(snapshot.fieldTypesByRecordType),
    fieldMenuChoicesByRecordType: cloneMapOfMaps(
      snapshot.fieldMenuChoicesByRecordType,
    ),
    fieldInitialValuesByRecordType: cloneMapOfMaps(
      snapshot.fieldInitialValuesByRecordType,
    ),
    deviceSupportDefinitionsByName: cloneMapOfArrays(
      snapshot.deviceSupportDefinitionsByName,
    ),
    driverDefinitionsByName: cloneMapOfArrays(snapshot.driverDefinitionsByName),
    registrarDefinitionsByName: cloneMapOfArrays(snapshot.registrarDefinitionsByName),
    functionDefinitionsByName: cloneMapOfArrays(snapshot.functionDefinitionsByName),
    variableDefinitionsByName: cloneMapOfArrays(snapshot.variableDefinitionsByName),
    recordsByName: new Map(snapshot.recordsByName),
    recordDefinitionsByName: cloneMapOfArrays(snapshot.recordDefinitionsByName),
    macros: new Set(snapshot.macros),
    fieldValuesByField: cloneMapOfSets(snapshot.fieldValuesByField),
    workspaceFiles: snapshot.workspaceFiles,
    workspaceFilesByAbsolutePath: snapshot.workspaceFilesByAbsolutePath,
    projectModel: snapshot.projectModel,
  };

  applyParsedData(
    merged,
    parseDocumentText(document.uri, document.getText(), document.languageId),
  );

  if (isIndexedContentDocument(document)) {
    const workspaceFileEntry = createWorkspaceFileEntry(document.uri);
    merged.workspaceFilesByAbsolutePath = new Map(snapshot.workspaceFilesByAbsolutePath);
    merged.workspaceFilesByAbsolutePath.set(
      normalizeFsPath(document.uri.fsPath),
      workspaceFileEntry,
    );
  }

  return merged;
}

function applyParsedData(snapshot, parsedData) {
  for (const record of parsedData.records) {
    if (!snapshot.recordsByName.has(record.name)) {
      snapshot.recordsByName.set(record.name, record.recordType);
    }
  }

  for (const recordDefinition of parsedData.recordDefinitions) {
    const searchNames =
      recordDefinition.searchNames instanceof Set && recordDefinition.searchNames.size > 0
        ? recordDefinition.searchNames
        : new Set([recordDefinition.name]);
    for (const searchName of searchNames) {
      addToMapOfArrays(
        snapshot.recordDefinitionsByName,
        searchName,
        recordDefinition,
      );
    }
  }

  for (const macroName of parsedData.macros) {
    snapshot.macros.add(macroName);
  }

  for (const recordType of parsedData.recordTypes) {
    snapshot.recordTypes.add(recordType);
  }

  for (const fieldName of parsedData.allFields) {
    snapshot.allFields.add(fieldName);
  }

  for (const [recordType, fieldNames] of parsedData.fieldsByRecordType.entries()) {
    addToMapOfSets(snapshot.fieldsByRecordType, recordType, fieldNames);
  }

  for (const [recordType, fieldTypes] of parsedData.fieldTypesByRecordType.entries()) {
    addToMapOfMaps(snapshot.fieldTypesByRecordType, recordType, fieldTypes);
  }

  for (const [recordType, fieldInitialValues] of parsedData.fieldInitialValuesByRecordType.entries()) {
    addToMapOfMaps(
      snapshot.fieldInitialValuesByRecordType,
      recordType,
      fieldInitialValues,
    );
  }

  for (const deviceSupportDefinition of parsedData.deviceSupportDefinitions) {
    addToMapOfArrays(
      snapshot.deviceSupportDefinitionsByName,
      deviceSupportDefinition.name,
      deviceSupportDefinition,
    );
  }

  for (const driverDefinition of parsedData.driverDefinitions) {
    addToMapOfArrays(
      snapshot.driverDefinitionsByName,
      driverDefinition.name,
      driverDefinition,
    );
  }

  for (const registrarDefinition of parsedData.registrarDefinitions) {
    addToMapOfArrays(
      snapshot.registrarDefinitionsByName,
      registrarDefinition.name,
      registrarDefinition,
    );
  }

  for (const functionDefinition of parsedData.functionDefinitions) {
    addToMapOfArrays(
      snapshot.functionDefinitionsByName,
      functionDefinition.name,
      functionDefinition,
    );
  }

  for (const variableDefinition of parsedData.variableDefinitions) {
    addToMapOfArrays(
      snapshot.variableDefinitionsByName,
      variableDefinition.name,
      variableDefinition,
    );
  }

  for (const [fieldName, fieldValues] of parsedData.fieldValuesByField.entries()) {
    addToMapOfSets(snapshot.fieldValuesByField, fieldName, fieldValues);
  }
}

function parseDocumentText(uri, text, languageId) {
  const recordDeclarations = extractRecordDeclarations(text);
  const data = {
    records: recordDeclarations.map((declaration) => ({
      recordType: declaration.recordType,
      name: declaration.name,
    })),
    recordDefinitions: createRecordDefinitions(uri, text, recordDeclarations),
    macros: new Set(extractMacroNames(text)),
    recordTypes: new Set(),
    allFields: new Set(),
    fieldsByRecordType: new Map(),
    fieldTypesByRecordType: new Map(),
    fieldMenuChoicesByRecordType: new Map(),
    fieldInitialValuesByRecordType: new Map(),
    deviceSupportDefinitions: [],
    driverDefinitions: [],
    registrarDefinitions: [],
    functionDefinitions: [],
    variableDefinitions: [],
    fieldValuesByField: extractFieldValues(text),
  };

  if (languageId === LANGUAGE_IDS.startup || hasExtension(uri, STARTUP_EXTENSIONS)) {
    for (const macroName of extractStartupMacros(text)) {
      data.macros.add(macroName);
    }
  }

  if (
    languageId === LANGUAGE_IDS.substitutions ||
    hasExtension(uri, SUBSTITUTION_EXTENSIONS)
  ) {
    for (const macroName of extractSubstitutionMacros(text)) {
      data.macros.add(macroName);
    }
  }

  if (languageId === LANGUAGE_IDS.dbd || hasExtension(uri, DBD_EXTENSIONS)) {
    const dbdFieldsByRecordType = extractDbdFieldsByRecordType(text);
    for (const [recordType, fieldNames] of dbdFieldsByRecordType.entries()) {
      data.recordTypes.add(recordType);
      addToMapOfSets(data.fieldsByRecordType, recordType, fieldNames);
      for (const fieldName of fieldNames) {
        data.allFields.add(fieldName);
      }
    }

    const dbdFieldTypesByRecordType = extractDbdFieldTypesByRecordType(text);
    for (const [recordType, fieldTypes] of dbdFieldTypesByRecordType.entries()) {
      addToMapOfMaps(data.fieldTypesByRecordType, recordType, fieldTypes);
    }

    const dbdFieldInitialValuesByRecordType =
      extractDbdFieldInitialValuesByRecordType(text);
    for (const [recordType, fieldInitialValues] of dbdFieldInitialValuesByRecordType.entries()) {
      addToMapOfMaps(
        data.fieldInitialValuesByRecordType,
        recordType,
        fieldInitialValues,
      );
    }
  }

  if (hasExtension(uri, SOURCE_EXTENSIONS)) {
    data.deviceSupportDefinitions = extractDeviceSupportDefinitions(uri, text);
    data.driverDefinitions = extractDriverDefinitions(uri, text);
    data.registrarDefinitions = extractRegistrarDefinitions(uri, text);
    data.functionDefinitions = extractFunctionDefinitions(uri, text);
    data.variableDefinitions = extractVariableDefinitions(uri, text);
  }

  return data;
}

function createEmptyProjectModel() {
  return {
    applications: [],
    runtimeArtifacts: [],
    startupEntryPoints: [],
    iocsByName: new Map(),
    releaseVariables: new Map(),
    availableDbds: new Map(),
    availableLibs: new Map(),
  };
}

async function buildProjectModel(projectFiles, buildModelCache) {
  const projectModel = createEmptyProjectModel();
  const filesByAbsolutePath = new Map();
  const rootPaths = new Set();

  for (const file of projectFiles) {
    filesByAbsolutePath.set(normalizeFsPath(file.uri.fsPath), file);

    if (
      ["RELEASE", "RELEASE.local"].includes(path.basename(file.uri.fsPath)) &&
      path.basename(path.dirname(file.uri.fsPath)) === "configure"
    ) {
      rootPaths.add(normalizeFsPath(path.dirname(path.dirname(file.uri.fsPath))));
    }
  }

  for (const rootPath of [...rootPaths].sort(compareLabels)) {
    const buildApplication =
      buildModelCache && typeof buildModelCache.getApplication === "function"
        ? await buildModelCache.getApplication(rootPath, filesByAbsolutePath)
        : undefined;
    const application = buildApplication
      ? normalizeBuildModelApplication(buildApplication)
      : buildProjectApplicationFromFiles(rootPath, filesByAbsolutePath);
    if (!application) {
      continue;
    }

    projectModel.applications.push(application);

    for (const artifact of application.runtimeArtifacts) {
      projectModel.runtimeArtifacts.push(artifact);
    }

    if (Array.isArray(application.startupEntryPoints)) {
      projectModel.startupEntryPoints.push(...application.startupEntryPoints);
    }

    for (const [iocName, iocInfo] of application.iocsByName.entries()) {
      if (!projectModel.iocsByName.has(iocName)) {
        projectModel.iocsByName.set(iocName, iocInfo);
      }
    }

    for (const [variableName, variableValue] of application.releaseVariables.entries()) {
      if (!projectModel.releaseVariables.has(variableName)) {
        projectModel.releaseVariables.set(variableName, variableValue);
      }
    }

    mergeProjectResourceMap(projectModel.availableDbds, application.availableDbds);
    mergeProjectResourceMap(projectModel.availableLibs, application.availableLibs);
  }

  return projectModel;
}

function normalizeBuildModelApplication(application) {
  const rootPath = normalizeFsPath(application?.rootPath);
  const rootUri = vscode.Uri.file(rootPath);
  const rootRelativePath = normalizePath(vscode.workspace.asRelativePath(rootUri, false));

  return {
    rootPath,
    rootUri,
    rootRelativePath,
    releaseVariables: new Map(Object.entries(application?.releaseVariables || {})),
    iocsByName: new Map(
      (Array.isArray(application?.iocs) ? application.iocs : [])
        .filter((iocInfo) => iocInfo?.name)
        .map((iocInfo) => [iocInfo.name, iocInfo]),
    ),
    runtimeArtifacts: Array.isArray(application?.runtimeArtifacts)
      ? application.runtimeArtifacts.map((artifact) => ({
          ...artifact,
          absoluteRuntimePath: normalizeFsPath(artifact.absoluteRuntimePath),
        }))
      : [],
    startupEntryPoints: Array.isArray(application?.startupEntryPoints)
      ? application.startupEntryPoints.map((entryPath) => normalizeFsPath(entryPath))
      : [],
    availableDbds: new Map(
      (Array.isArray(application?.availableDbds) ? application.availableDbds : [])
        .filter((entry) => entry?.name)
        .map((entry) => [entry.name, entry]),
    ),
    availableLibs: new Map(
      (Array.isArray(application?.availableLibs) ? application.availableLibs : [])
        .filter((entry) => entry?.name)
        .map((entry) => [entry.name, entry]),
    ),
    buildInfo: application?.buildInfo || undefined,
  };
}

function buildProjectApplicationFromFiles(rootPath, filesByAbsolutePath) {
  const rootUri = vscode.Uri.file(rootPath);
  const rootRelativePath = normalizePath(vscode.workspace.asRelativePath(rootUri, false));
  const releaseVariables = new Map();
  const iocsByName = new Map();
  const runtimeArtifacts = [];
  const releaseFilePaths = [
    normalizeFsPath(path.join(rootPath, "configure", "RELEASE")),
    normalizeFsPath(path.join(rootPath, "configure", "RELEASE.local")),
  ];
  const releaseText = releaseFilePaths
    .map((releaseFilePath) => filesByAbsolutePath.get(releaseFilePath)?.text)
    .filter(Boolean)
    .join("\n");

  if (releaseText) {
    for (const [variableName, values] of parseMakeAssignments(releaseText).entries()) {
      if (values.length > 0) {
        releaseVariables.set(variableName, values.join(" "));
      }
    }
  }

  for (const [filePath, file] of filesByAbsolutePath.entries()) {
    if (!isPathWithinRoot(filePath, rootPath) || path.basename(filePath) !== "Makefile") {
      continue;
    }

    const relativePath = normalizePath(path.relative(rootPath, filePath));
    const srcMatch = relativePath.match(/^([^/]+App)\/src\/Makefile$/);
    if (srcMatch) {
      const appDirName = srcMatch[1];
      const assignments = parseMakeAssignments(file.text);
      const iocNames = assignments.get("PROD_IOC") || [];
      const dbdNames = assignments.get("DBD") || [];

      for (const iocName of iocNames) {
        iocsByName.set(iocName, {
          name: iocName,
          appDirName,
          makefileRelativePath: relativePath,
          registerFunctionName: `${iocName}_registerRecordDeviceDriver`,
        });
      }

      for (const dbdName of dbdNames) {
        runtimeArtifacts.push(
          createRuntimeArtifact({
            rootPath,
            appDirName,
            kind: "dbd",
            runtimeRelativePath: normalizePath(path.posix.join("dbd", dbdName)),
            sourceRelativePath: relativePath,
            detail: `Generated DBD from ${relativePath}`,
            documentation: `Generated by ${relativePath}`,
          }),
        );
      }

      continue;
    }

    const dbMatch = relativePath.match(/^([^/]+App)\/(Db|db)\/Makefile$/);
    if (!dbMatch) {
      continue;
    }

    const appDirName = dbMatch[1];
    const dbDirRelativePath = normalizePath(path.posix.join(appDirName, dbMatch[2]));
    const assignments = parseMakeAssignments(file.text);
    const installedDbNames = assignments.get("DB") || [];
    const templateMappings = extractTemplateMappings(assignments);

    for (const installedDbName of installedDbNames) {
      const sourceFileName = resolveDatabaseSourceFileName(
        installedDbName,
        templateMappings,
      );
      const sourceRelativePath = normalizePath(
        path.posix.join(dbDirRelativePath, sourceFileName),
      );
      const artifactKind = getRuntimeArtifactKind(installedDbName);

      if (!artifactKind) {
        continue;
      }

      runtimeArtifacts.push(
        createRuntimeArtifact({
          rootPath,
          appDirName,
          kind: artifactKind,
          runtimeRelativePath: normalizePath(path.posix.join("db", installedDbName)),
          sourceRelativePath,
          detail: `Installed ${artifactKind} from ${sourceRelativePath}`,
          documentation: `Declared in ${relativePath}`,
        }),
      );
    }
  }

  const releaseRoots = resolveReleaseModuleRoots(rootPath, releaseVariables);
  const availableDbds = buildProjectDbdEntries({
    rootPath,
    runtimeArtifacts,
    releaseRoots,
    rootRelativePath,
  });
  const availableLibs = buildProjectLibEntries({
    rootPath,
    releaseRoots,
    rootRelativePath,
  });

  return {
    rootPath,
    rootUri,
    rootRelativePath,
    releaseVariables,
    iocsByName,
    runtimeArtifacts,
    startupEntryPoints: findProjectIocBootFilePaths(rootPath),
    availableDbds,
    availableLibs,
  };
}

async function refreshDiagnostics(document, diagnostics, workspaceIndex) {
  if (
    !isDatabaseDocument(document) &&
    !isStartupDocument(document) &&
    !isSubstitutionsDocument(document) &&
    !isSourceMakefileDocument(document)
  ) {
    diagnostics.delete(document.uri);
    return;
  }

  const snapshot = mergeSnapshotWithDocument(
    await workspaceIndex.getSnapshot(),
    document,
  );

  if (isDatabaseDocument(document)) {
    diagnostics.set(document.uri, createDatabaseDiagnostics(document, snapshot));
    return;
  }

  if (isSubstitutionsDocument(document)) {
    diagnostics.set(document.uri, createSubstitutionDiagnostics(document, snapshot));
    return;
  }

  if (isSourceMakefileDocument(document)) {
    diagnostics.set(document.uri, createSourceMakefileDiagnostics(document, snapshot));
    return;
  }

  diagnostics.set(document.uri, createStartupDiagnostics(document, snapshot));
}

function createDatabaseDiagnostics(document, snapshot) {
  return [
    ...createUnmatchedDelimiterDiagnostics(document),
    ...createDuplicateRecordDiagnostics(document),
    ...createDuplicateFieldDiagnostics(document),
    ...createInvalidFieldDiagnostics(document, snapshot),
    ...createInvalidNumericFieldValueDiagnostics(document, snapshot),
    ...createInvalidMenuFieldValueDiagnostics(document, snapshot),
    ...createDescFieldLengthDiagnostics(document),
    ...createMakefileInclusionDiagnostics(document),
  ];
}

function createUnmatchedDelimiterDiagnostics(document) {
  const diagnostics = [];
  const text = document.getText();
  const delimiterStack = [];
  const matchingDelimiters = new Map([
    [")", "("],
    ["}", "{"],
  ]);
  let inString = false;
  let escaped = false;
  let inComment = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];

    if (inComment) {
      if (character === "\n") {
        inComment = false;
      }
      continue;
    }

    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inString = false;
      }
      continue;
    }

    if (character === "#") {
      inComment = true;
      continue;
    }

    if (character === "\"") {
      inString = true;
      continue;
    }

    if (character === "(" || character === "{") {
      delimiterStack.push({ character, index });
      continue;
    }

    if (character !== ")" && character !== "}") {
      continue;
    }

    const expectedOpening = matchingDelimiters.get(character);
    const lastOpening = delimiterStack[delimiterStack.length - 1];
    if (!lastOpening || lastOpening.character !== expectedOpening) {
      diagnostics.push(
        createDiagnostic(
          document.positionAt(index),
          document.positionAt(index + 1),
          `Unmatched "${character}".`,
        ),
      );
      continue;
    }

    delimiterStack.pop();
  }

  for (const unmatchedOpening of delimiterStack) {
    const expectedClosing = unmatchedOpening.character === "(" ? ")" : "}";
    diagnostics.push(
      createDiagnostic(
        document.positionAt(unmatchedOpening.index),
        document.positionAt(unmatchedOpening.index + 1),
        `Unmatched "${unmatchedOpening.character}"; missing "${expectedClosing}".`,
      ),
    );
  }

  return diagnostics;
}

function createDuplicateRecordDiagnostics(document) {
  const declarations = extractRecordDeclarations(document.getText());
  const declarationsByName = new Map();
  const diagnostics = [];

  for (const declaration of declarations) {
    addToMapOfArrays(declarationsByName, declaration.name, declaration);
  }

  for (const [recordName, duplicates] of declarationsByName.entries()) {
    if (duplicates.length < 2) {
      continue;
    }

    for (const declaration of duplicates) {
      const start = document.positionAt(declaration.nameStart);
      const end = document.positionAt(declaration.nameEnd);
      const diagnostic = new vscode.Diagnostic(
        new vscode.Range(start, end),
        `Duplicate record name "${recordName}" in this file.`,
        vscode.DiagnosticSeverity.Error,
      );
      diagnostic.source = "vscode-epics";
      diagnostic.code = "epics.database.duplicateRecordName";
      diagnostics.push(diagnostic);
    }
  }

  return diagnostics;
}

function createDuplicateFieldDiagnostics(document) {
  const diagnostics = [];
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    const fieldsByName = new Map();

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      addToMapOfArrays(fieldsByName, fieldDeclaration.fieldName, fieldDeclaration);
    }

    for (const [fieldName, duplicates] of fieldsByName.entries()) {
      if (duplicates.length < 2) {
        continue;
      }

      for (const duplicate of duplicates) {
        diagnostics.push(
          createDiagnostic(
            document.positionAt(duplicate.fieldNameStart),
            document.positionAt(duplicate.fieldNameEnd),
            `Duplicate field "${fieldName}" in record "${recordDeclaration.name}".`,
          ),
        );
      }
    }
  }

  return diagnostics;
}

function createInvalidFieldDiagnostics(document, snapshot) {
  const diagnostics = [];
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    const allowedFields = snapshot.fieldsByRecordType.get(recordDeclaration.recordType);
    if (!allowedFields || allowedFields.size === 0) {
      continue;
    }

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      if (allowedFields.has(fieldDeclaration.fieldName)) {
        continue;
      }

      diagnostics.push(
        Object.assign(
          createDiagnostic(
            document.positionAt(fieldDeclaration.fieldNameStart),
            document.positionAt(fieldDeclaration.fieldNameEnd),
            `Field "${fieldDeclaration.fieldName}" is not valid for record type "${recordDeclaration.recordType}".`,
          ),
          {
            code: "epics.database.invalidFieldName",
          },
        ),
      );
    }
  }

  return diagnostics;
}

function createInvalidNumericFieldValueDiagnostics(document, snapshot) {
  const diagnostics = [];
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    const fieldTypes = snapshot.fieldTypesByRecordType.get(recordDeclaration.recordType);
    if (!fieldTypes || fieldTypes.size === 0) {
      continue;
    }

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      const dbfType = fieldTypes.get(fieldDeclaration.fieldName);
      if (!NUMERIC_DBF_TYPES.has(dbfType)) {
        continue;
      }

      if (isSkippableNumericFieldValue(fieldDeclaration.value)) {
        continue;
      }

      if (isValidNumericFieldValue(fieldDeclaration.value, dbfType)) {
        continue;
      }

      diagnostics.push(
        createDiagnostic(
          document.positionAt(fieldDeclaration.valueStart),
          document.positionAt(fieldDeclaration.valueEnd),
          `Field "${fieldDeclaration.fieldName}" expects a ${dbfType} numeric value.`,
        ),
      );
    }
  }

  return diagnostics;
}

function createInvalidMenuFieldValueDiagnostics(document, snapshot) {
  const diagnostics = [];
  const text = document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    const fieldTypes = snapshot.fieldTypesByRecordType.get(recordDeclaration.recordType);
    if (!fieldTypes || fieldTypes.size === 0) {
      continue;
    }

    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      const dbfType = fieldTypes.get(fieldDeclaration.fieldName);
      if (dbfType !== "DBF_MENU") {
        continue;
      }

      if (containsEpicsMacroReference(fieldDeclaration.value)) {
        continue;
      }

      const allowedChoices = getMenuFieldChoices(
        snapshot,
        recordDeclaration.recordType,
        fieldDeclaration.fieldName,
      );
      if (allowedChoices.length === 0 || allowedChoices.includes(fieldDeclaration.value)) {
        continue;
      }

      diagnostics.push(
        Object.assign(
          createDiagnostic(
            document.positionAt(fieldDeclaration.valueStart),
            document.positionAt(fieldDeclaration.valueEnd),
            `Field "${fieldDeclaration.fieldName}" must be one of the menu choices for "${recordDeclaration.recordType}".`,
          ),
          {
            code: "epics.database.invalidMenuFieldValue",
          },
        ),
      );
    }
  }

  return diagnostics;
}

function createDescFieldLengthDiagnostics(document) {
  const diagnostics = [];

  for (const field of extractFieldDeclarations(document.getText(), "DESC")) {
    const descriptionLength = getEscapedStringLength(field.value);
    if (descriptionLength <= 40) {
      continue;
    }

    diagnostics.push(
      createDiagnostic(
        document.positionAt(field.valueStart),
        document.positionAt(field.valueEnd),
        `DESC must be 40 characters or fewer; found ${descriptionLength}.`,
      ),
    );
  }

  return diagnostics;
}

function createStartupDiagnostics(document, snapshot) {
  const diagnostics = [];
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const state = createInitialStartupExecutionState(snapshot, document);
  const statements = extractStartupStatements(document.getText());

  for (const statement of statements) {
    switch (statement.kind) {
      case "include": {
        const resolution = resolveStartupPath(snapshot, document, statement, state);
        if (!resolution) {
          diagnostics.push(
            createDiagnostic(
              document.positionAt(statement.pathStart),
              document.positionAt(statement.pathEnd),
              buildStartupPathDiagnosticMessage(snapshot, document, statement, project),
            ),
          );
          break;
        }

        applyIncludedStartupState(resolution.absolutePath, state);
        break;
      }

      case "envSet":
        state.envVariables.set(
          statement.name,
          expandStartupValue(statement.value, state.envVariables),
        );
        break;

      case "cd": {
        const resolution = resolveStartupPath(snapshot, document, statement, state);
        if (!resolution || !resolution.isDirectory) {
          diagnostics.push(
            createDiagnostic(
              document.positionAt(statement.pathStart),
              document.positionAt(statement.pathEnd),
              buildStartupPathDiagnosticMessage(snapshot, document, statement, project),
            ),
          );
          break;
        }

        state.currentDirectory = resolution.absolutePath;
        break;
      }

      case "load": {
        const resolution = resolveStartupPath(snapshot, document, statement, state);
        if (!resolution) {
          diagnostics.push(
            createDiagnostic(
              document.positionAt(statement.pathStart),
              document.positionAt(statement.pathEnd),
              buildStartupPathDiagnosticMessage(snapshot, document, statement, project),
            ),
          );
          break;
        }

        if (statement.command === "dbLoadRecords") {
          diagnostics.push(
            ...createStartupLoadMacroDiagnostics(
              document,
              statement,
              getReadableStartupFileResolution(document, resolution),
            ),
          );
        }
        break;
      }

      case "register":
        if (!project || project.iocsByName.has(statement.iocName)) {
          continue;
        }

        diagnostics.push(
          Object.assign(
            createDiagnostic(
              document.positionAt(statement.nameStart),
              document.positionAt(statement.nameEnd),
              `Unknown IOC registration function "${statement.functionName}" for this EPICS application.`,
            ),
            {
              code: "epics.startup.unknownIocRegistrationFunction",
            },
          ),
        );
        break;

      default:
        break;
    }
  }

  return diagnostics;
}

function createStartupLoadMacroDiagnostics(document, statement, resolvedFile) {
  if (!resolvedFile?.text) {
    return [];
  }

  const requiredMacroNames = extractRequiredMacroNames(
    maskDatabaseComments(resolvedFile.text),
  );
  if (requiredMacroNames.length === 0) {
    return [];
  }

  const providedMacroNames = extractAssignedMacroNames(statement.macros);
  const missingMacroNames = requiredMacroNames.filter(
    (macroName) => !providedMacroNames.has(macroName),
  );
  if (missingMacroNames.length === 0) {
    return [];
  }

  return [
    Object.assign(
      createDiagnostic(
        document.positionAt(statement.pathStart),
        document.positionAt(statement.pathEnd),
        `dbLoadRecords is missing macro assignments for "${path.posix.basename(
          normalizePath(statement.path),
        )}": ${missingMacroNames.join(", ")}.`,
      ),
      {
        code: "epics.startup.missingDbLoadRecordsMacros",
      },
    ),
  ];
}

function createSubstitutionDiagnostics(document, snapshot) {
  if (!isSubstitutionsDocument(document)) {
    return [];
  }

  const diagnostics = [...createMakefileInclusionDiagnostics(document)];
  let globalMacros = new Map();
  const text = document.getText();

  for (const block of extractSubstitutionBlocksWithRanges(text)) {
    if (block.kind === "global") {
      globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body));
      continue;
    }

    if (
      block.kind !== "file" ||
      !block.templatePath ||
      block.templatePathStart === undefined ||
      block.templatePathEnd === undefined
    ) {
      continue;
    }

    const absolutePath = resolveSubstitutionTemplateAbsolutePathForDocument(
      snapshot,
      document,
      block.templatePath,
    );
    if (!absolutePath) {
      diagnostics.push(
        createDiagnostic(
          document.positionAt(block.templatePathStart),
          document.positionAt(block.templatePathEnd),
          `Cannot resolve substitutions database/template file "${block.templatePath}" from the local folder or configure/RELEASE db directories.`,
        ),
      );
      continue;
    }

    const templateText = readTextFile(absolutePath);
    if (templateText === undefined) {
      continue;
    }

    const templateBaseName = path.posix.basename(normalizePath(block.templatePath));
    const maskedTemplateText = maskDatabaseComments(templateText);
    const requiredMacroNames = extractRequiredMacroNames(maskedTemplateText);
    const templateMacroNames = new Set(extractMacroNames(maskedTemplateText));
    const parsedRows = parseSubstitutionFileBlockRowsDetailed(block.body, block.bodyStart);

    if (parsedRows.kind === "pattern") {
      const effectiveNames = new Set([
        ...globalMacros.keys(),
        ...parsedRows.columns,
      ]);
      const missingMacroNames = requiredMacroNames.filter(
        (macroName) => !effectiveNames.has(macroName),
      );
      const excessiveMacroNames = parsedRows.columns.filter(
        (macroName) => !templateMacroNames.has(macroName),
      );

      if (missingMacroNames.length > 0 || excessiveMacroNames.length > 0) {
        const details = [];
        if (missingMacroNames.length > 0) {
          details.push(`missing: ${missingMacroNames.join(", ")}`);
        }
        if (excessiveMacroNames.length > 0) {
          details.push(`excessive: ${excessiveMacroNames.join(", ")}`);
        }

        diagnostics.push(
          createDiagnostic(
            document.positionAt(parsedRows.headerRangeStart),
            document.positionAt(parsedRows.headerRangeEnd),
            `Pattern macros for "${templateBaseName}" are ${details.join("; ")}.`,
          ),
        );
      }

      const rowsEligibleForDuplicateCheck = [];
      for (const row of parsedRows.rows) {
        const missingColumnNames = parsedRows.columns.filter(
          (_, index) => row.values[index] === undefined,
        );
        if (missingColumnNames.length > 0) {
          diagnostics.push(
            createDiagnostic(
              document.positionAt(row.rangeStart),
              document.positionAt(row.rangeEnd),
              `Pattern row for "${templateBaseName}" is missing values for: ${missingColumnNames.join(", ")}.`,
            ),
          );
          continue;
        }

        if (row.values.length > parsedRows.columns.length) {
          diagnostics.push(
            createDiagnostic(
              document.positionAt(row.rangeStart),
              document.positionAt(row.rangeEnd),
              `Pattern row for "${templateBaseName}" has ${row.values.length - parsedRows.columns.length} extra value(s).`,
            ),
          );
          continue;
        }

        rowsEligibleForDuplicateCheck.push(row);
      }

      if (missingMacroNames.length === 0 && excessiveMacroNames.length === 0) {
        diagnostics.push(
          ...createSubstitutionDuplicateRowDiagnostics(
            document,
            templateBaseName,
            rowsEligibleForDuplicateCheck,
            globalMacros,
          ),
        );
      }

      continue;
    }

    const rowsEligibleForDuplicateCheck = [];
    for (const row of parsedRows.rows) {
      const effectiveNames = new Set([
        ...globalMacros.keys(),
        ...row.assignments.keys(),
      ]);
      const missingMacroNames = requiredMacroNames.filter(
        (macroName) => !effectiveNames.has(macroName),
      );
      if (missingMacroNames.length > 0) {
        diagnostics.push(
          createDiagnostic(
            document.positionAt(row.rangeStart),
            document.positionAt(row.rangeEnd),
            `Macro assignments for "${templateBaseName}" are missing: ${missingMacroNames.join(", ")}.`,
          ),
        );
      }

      if (missingMacroNames.length === 0) {
        rowsEligibleForDuplicateCheck.push(row);
      }

      for (const macroName of row.assignments.keys()) {
        if (templateMacroNames.has(macroName)) {
          continue;
        }

        const range = row.nameRanges.get(macroName) || {
          start: row.rangeStart,
          end: row.rangeEnd,
        };
        diagnostics.push(
          createDiagnostic(
            document.positionAt(range.start),
            document.positionAt(range.end),
            `Macro "${macroName}" is not used by "${templateBaseName}".`,
          ),
        );
      }
    }

    diagnostics.push(
      ...createSubstitutionDuplicateRowDiagnostics(
        document,
        templateBaseName,
        rowsEligibleForDuplicateCheck,
        globalMacros,
      ),
    );
  }

  return diagnostics;
}

function createSubstitutionDuplicateRowDiagnostics(
  document,
  templateBaseName,
  rows,
  globalMacros,
) {
  if (rows.length < 2) {
    return [];
  }

  const rowIndexesBySignature = new Map();
  const duplicateRowIndexes = new Set();

  rows.forEach((row, rowIndex) => {
    const effectiveAssignments = [...mergeMacroMaps(globalMacros, row.assignments).entries()]
      .sort((left, right) => compareLabels(left[0], right[0]));
    const signature = JSON.stringify(effectiveAssignments);
    const existingRowIndexes = rowIndexesBySignature.get(signature) || [];
    for (const existingRowIndex of existingRowIndexes) {
      duplicateRowIndexes.add(existingRowIndex);
      duplicateRowIndexes.add(rowIndex);
    }
    addToMapOfArrays(rowIndexesBySignature, signature, rowIndex);
  });

  const diagnostics = [];
  for (const rowIndex of [...duplicateRowIndexes].sort(
    (left, right) => rows[left].rangeStart - rows[right].rangeStart,
  )) {
    diagnostics.push(
      createDiagnostic(
        document.positionAt(rows[rowIndex].rangeStart),
        document.positionAt(rows[rowIndex].rangeEnd),
        `Duplicate substitutions macro assignments for "${templateBaseName}".`,
      ),
    );
  }

  return diagnostics;
}

function createSourceMakefileDiagnostics(document, snapshot) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (!project) {
    return [];
  }

  const diagnostics = [];
  const references = extractMakefileReferences(document.getText());

  for (const reference of references) {
    if (!isProjectResourceMakefileReference(reference)) {
      continue;
    }

    if (!isConcreteMakefileReferenceToken(reference.name, reference.kind)) {
      continue;
    }

    const resourceMap =
      reference.kind === "dbd" ? project.availableDbds : project.availableLibs;
    if (resourceMap.has(reference.name)) {
      continue;
    }
    if (
      reference.kind === "dbd" &&
      hasLocalMakefileDbdReference(document, reference.name)
    ) {
      continue;
    }

    const message =
      reference.kind === "dbd"
        ? `Unknown DBD "${reference.name}". It was not found in the current Makefile folder, this project's dbd outputs, or the module roots from RELEASE.`
        : `Unknown library "${reference.name}". It was not found in this project's lib directories or the module roots from RELEASE.`;
    diagnostics.push(
      createDiagnostic(
        document.positionAt(reference.start),
        document.positionAt(reference.end),
        message,
      ),
    );
  }

  return diagnostics;
}

function extractRecords(text) {
  return extractRecordDeclarations(text).map((declaration) => ({
    recordType: declaration.recordType,
    name: declaration.name,
  }));
}

function createRecordDefinitions(uri, text, declarations) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const workspaceFileEntry = createWorkspaceFileEntry(uri);
  return declarations.map((declaration) => ({
    name: declaration.name,
    searchNames: getDatabaseRecordSearchNames(text, declaration.name),
    recordType: declaration.recordType,
    absolutePath: normalizeFsPath(uri.fsPath),
    relativePath: workspaceFileEntry.relativePath,
    line: getLineNumberAtOffset(text, declaration.recordStart),
    preview: buildRecordPreview(text, declaration),
  }));
}

function createRecordDocumentSymbol(document, declaration, textLength) {
  const symbolRange = new vscode.Range(
    document.positionAt(Math.max(0, declaration.recordStart)),
    document.positionAt(
      Math.min(Math.max(declaration.recordStart, declaration.recordEnd), textLength),
    ),
  );
  const selectionRange = new vscode.Range(
    document.positionAt(Math.max(0, declaration.nameStart)),
    document.positionAt(
      Math.min(Math.max(declaration.nameStart, declaration.nameEnd), textLength),
    ),
  );
  const symbolName = `${declaration.recordType} ${declaration.name}`;

  return new vscode.DocumentSymbol(
    symbolName,
    "",
    vscode.SymbolKind.Variable,
    symbolRange,
    selectionRange,
  );
}

async function foldAllRecordsInActiveEditor(commandName) {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isDatabaseDocument(editor.document)) {
    return;
  }

  const selectionLines = extractRecordDeclarations(editor.document.getText())
    .map((declaration) => editor.document.positionAt(declaration.recordStart).line)
    .filter((line, index, lines) => index === 0 || line !== lines[index - 1]);

  if (!selectionLines.length) {
    return;
  }

  await vscode.commands.executeCommand(commandName, {
    levels: 1,
    selectionLines,
  });
}

async function generateDatabaseTocInActiveEditor() {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isDatabaseDocument(editor.document)) {
    return;
  }

  const document = editor.document;
  const originalText = document.getText();
  const nextText = upsertDatabaseTocText(originalText, getDocumentEol(document));
  if (nextText === originalText) {
    return;
  }

  const fullRange = new vscode.Range(
    new vscode.Position(0, 0),
    document.positionAt(originalText.length),
  );
  await editor.edit((editBuilder) => {
    editBuilder.replace(fullRange, nextText);
  });
}

async function formatDatabaseFileInActiveEditor() {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isDatabaseDocument(editor.document)) {
    return;
  }

  await vscode.commands.executeCommand("editor.action.formatDocument");
}

async function formatActiveEpicsFileInActiveEditor() {
  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  if (
    !document ||
    (
      !isDatabaseDocument(document) &&
      !isSubstitutionsDocument(document) &&
      !isStartupDocument(document) &&
      !isMakefileDocument(document) &&
      !isProtocolDocument(document) &&
      document.languageId !== LANGUAGE_IDS.dbd
    )
  ) {
    return;
  }

  await vscode.commands.executeCommand("editor.action.formatDocument");
}

async function copyAllRecordNamesInActiveEditor() {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isDatabaseDocument(editor.document)) {
    return;
  }

  const recordNames = extractUniqueRecordNames(editor.document.getText());
  if (recordNames.length === 0) {
    vscode.window.showWarningMessage(
      "No EPICS record names were found in the active database file.",
    );
    return;
  }

  await vscode.env.clipboard.writeText(
    buildRecordNamesClipboardText(recordNames, getDocumentEol(editor.document)),
  );
}

async function copyDatabaseAsMonitorFileInActiveEditor() {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isDatabaseDocument(editor.document)) {
    return;
  }

  const text = editor.document.getText();
  const recordNames = extractUniqueRecordNames(text);
  if (recordNames.length === 0) {
    vscode.window.showWarningMessage(
      "No EPICS record names were found in the active database file.",
    );
    return;
  }

  const macroNames = extractRecordNameMacroNames(recordNames);
  const monitorText = buildMonitorFileText(
    recordNames,
    macroNames,
    getDocumentEol(editor.document),
  );
  await vscode.env.clipboard.writeText(monitorText);

  const recordLabel = `${recordNames.length} record${recordNames.length === 1 ? "" : "s"}`;
  const macroLabel = `${macroNames.length} macro${macroNames.length === 1 ? "" : "s"}`;
  vscode.window.showInformationMessage(
    `Copied ${recordLabel} and ${macroLabel} as a .pvlist file.`,
  );
}

async function exportDatabaseToExcelResource(resourceUri, workspaceIndex) {
  const document = await resolveDocumentForCommand(resourceUri);
  if (!document || (!isDatabaseDocument(document) && !isSubstitutionsDocument(document))) {
    return;
  }

  let sourceText = document.getText();
  if (isSubstitutionsDocument(document)) {
    const expandedSource = await resolveExpandedSubstitutionsDatabaseSource(
      workspaceIndex,
      document,
      "export Excel",
    );
    if (!expandedSource) {
      return;
    }
    sourceText = expandedSource.text;
  }

  const workbookBuffer = buildDatabaseWorkbookBuffer(sourceText);
  const sourcePath = document.uri.scheme === "file" ? document.uri.fsPath : "database";
  const defaultUri = document.uri.scheme === "file"
    ? vscode.Uri.file(
      `${sourcePath.replace(/\.[^./\\]+$/, "")}.xlsx`,
    )
    : undefined;
  const targetUri = await vscode.window.showSaveDialog({
    defaultUri,
    filters: {
      "Excel Workbook": ["xlsx"],
    },
    saveLabel: "Export to Excel",
  });
  if (!targetUri) {
    return;
  }

  await vscode.workspace.fs.writeFile(targetUri, workbookBuffer);
  const targetLabel = path.basename(targetUri.fsPath || targetUri.path || "database.xlsx");
  const openChoice = `Open ${targetLabel}`;
  const selectedChoice = await vscode.window.showInformationMessage(
    `Exported ${path.basename(sourcePath)} to ${targetLabel}.`,
    openChoice,
  );
  if (selectedChoice === openChoice) {
    await vscode.env.openExternal(targetUri);
  }
}

async function importDatabaseFromExcelResource(resourceUri) {
  let targetUri =
    resourceUri instanceof vscode.Uri
      ? resourceUri
      : Array.isArray(resourceUri) && resourceUri[0] instanceof vscode.Uri
        ? resourceUri[0]
        : resourceUri?.fsPath
          ? vscode.Uri.file(resourceUri.fsPath)
          : undefined;
  if (
    targetUri &&
    path.extname(targetUri.fsPath || targetUri.path).toLowerCase() !== ".xlsx"
  ) {
    targetUri = undefined;
  }
  if (!targetUri) {
    const selectedUris = await vscode.window.showOpenDialog({
      canSelectFiles: true,
      canSelectMany: false,
      filters: {
        "Excel Workbook": ["xlsx"],
      },
      openLabel: "Import Excel as EPICS DB",
    });
    targetUri = selectedUris?.[0];
  }
  if (!targetUri || path.extname(targetUri.fsPath || targetUri.path).toLowerCase() !== ".xlsx") {
    return;
  }

  let workbookBuffer;
  let fileStats;
  try {
    workbookBuffer = Buffer.from(await vscode.workspace.fs.readFile(targetUri));
    fileStats = await vscode.workspace.fs.stat(targetUri);
  } catch (error) {
    vscode.window.showErrorMessage(
      `Failed to read ${path.basename(targetUri.fsPath || targetUri.path)}: ${getErrorMessage(error)}`,
    );
    return;
  }

  const importedSheets = importDatabaseWorkbookBuffer(workbookBuffer, {
    sourceFileName: path.basename(targetUri.fsPath || targetUri.path),
    sourceCreatedAt: fileStats?.ctime ? new Date(fileStats.ctime) : undefined,
    sourceModifiedAt: fileStats?.mtime ? new Date(fileStats.mtime) : undefined,
    importedAt: new Date(),
  });
  if (!importedSheets.length) {
    vscode.window.showWarningMessage(
      `No EPICS-style sheets were found in ${path.basename(targetUri.fsPath || targetUri.path)}.`,
    );
    return;
  }

  await openImportedDatabaseTabs(importedSheets);
  const sheetLabel = `${importedSheets.length} sheet${importedSheets.length === 1 ? "" : "s"}`;
  vscode.window.showInformationMessage(
    `Imported ${sheetLabel} from ${path.basename(targetUri.fsPath || targetUri.path)}.`,
  );
}

async function openExcelImportPreviewPanel(context) {
  const panel = vscode.window.createWebviewPanel(
    EXCEL_IMPORT_PREVIEW_VIEW_TYPE,
    "EPICS Excel Import Preview",
    vscode.ViewColumn.Active,
    {
      enableScripts: true,
      retainContextWhenHidden: true,
    },
  );
  panel.webview.html = buildExcelImportPreviewWebviewHtml(panel.webview);
  panel.webview.onDidReceiveMessage(async (message) => {
    if (message?.type !== "importExcelPreviewWorkbook") {
      return;
    }

    let workbookBuffer;
    try {
      workbookBuffer = Buffer.from(String(message.base64 || ""), "base64");
    } catch (error) {
      await panel.webview.postMessage({
        type: "excelImportPreviewResult",
        success: false,
        message: `Failed to decode workbook: ${getErrorMessage(error)}`,
      });
      return;
    }

    let importedSheets;
    try {
      importedSheets = importDatabaseWorkbookBuffer(workbookBuffer, {
        sourceFileName: message.name || "dropped.xlsx",
        sourceModifiedAt: message.lastModified ? new Date(message.lastModified) : undefined,
        importedAt: new Date(),
      });
    } catch (error) {
      await panel.webview.postMessage({
        type: "excelImportPreviewResult",
        success: false,
        message: getErrorMessage(error),
      });
      return;
    }

    if (!importedSheets.length) {
      await panel.webview.postMessage({
        type: "excelImportPreviewResult",
        success: false,
        message: `No EPICS-style sheets were found in ${message.name || "the dropped workbook"}.`,
      });
      return;
    }

    await openImportedDatabaseTabs(importedSheets);
    const sheetLabel = `${importedSheets.length} sheet${importedSheets.length === 1 ? "" : "s"}`;
    await panel.webview.postMessage({
      type: "excelImportPreviewResult",
      success: true,
      message: `Imported ${sheetLabel} from ${message.name || "the dropped workbook"}.`,
    });
  });
}

async function openImportedDatabaseTabs(importedSheets) {
  for (let index = 0; index < importedSheets.length; index += 1) {
    const importedSheet = importedSheets[index];
    await openUntitledImportedDatabaseDocument(
      importedSheet.suggestedFileName,
      importedSheet.text,
      index + 1 < importedSheets.length,
    );
  }
}

async function updateEpicsProjectContext() {
  let isEpicsProject = false;
  for (const workspaceFolder of vscode.workspace.workspaceFolders || []) {
    if (await isWorkspaceFolderEpicsProject(workspaceFolder)) {
      isEpicsProject = true;
      break;
    }
  }
  await vscode.commands.executeCommand("setContext", EPICS_PROJECT_CONTEXT_KEY, isEpicsProject);
}

function updateActiveMakefileContextKeys(editor) {
  const document = editor?.document;
  const pendingInsertion = getPendingMakefileDbInsertion(document);
  const buildDirectory = resolveLocalMakeBuildDirectoryForDocument(document);

  void vscode.commands.executeCommand(
    "setContext",
    ACTIVE_CAN_ADD_DB_TO_MAKEFILE_CONTEXT_KEY,
    Boolean(pendingInsertion),
  );
  void vscode.commands.executeCommand(
    "setContext",
    ACTIVE_CAN_BUILD_WITH_MAKEFILE_CONTEXT_KEY,
    Boolean(buildDirectory),
  );
}

async function isWorkspaceFolderEpicsProject(workspaceFolder) {
  for (const pathSegments of EPICS_PROJECT_MARKER_SEGMENTS) {
    const candidateUri = vscode.Uri.joinPath(workspaceFolder.uri, ...pathSegments);
    if (!(await doesUriExist(candidateUri))) {
      return false;
    }
  }
  return true;
}

async function doesUriExist(uri) {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch (error) {
    return false;
  }
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

function isMakefileInstallableDatabaseSourceDocument(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  const extension = path.extname(document.uri.fsPath).toLowerCase();
  return (
    MAKEFILE_INSTALLABLE_DATABASE_EXTENSIONS.has(extension) ||
    SUBSTITUTION_EXTENSIONS.has(extension)
  );
}

function getMakefileDbInstallTokenForDocument(document) {
  if (!isMakefileInstallableDatabaseSourceDocument(document)) {
    return undefined;
  }

  const parsedPath = path.parse(document.uri.fsPath);
  const extension = parsedPath.ext.toLowerCase();
  if (SUBSTITUTION_EXTENSIONS.has(extension)) {
    return `${parsedPath.name}.db`;
  }

  return parsedPath.base;
}

function getSiblingMakefilePathForDocument(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  if (isMakefileDocument(document)) {
    return normalizeFsPath(document.uri.fsPath);
  }

  return normalizeFsPath(path.join(path.dirname(document.uri.fsPath), "Makefile"));
}

function getPendingMakefileDbInsertion(document) {
  const makefilePath = getSiblingMakefilePathForDocument(document);
  const installToken = getMakefileDbInstallTokenForDocument(document);
  if (!makefilePath || !installToken || !fs.existsSync(makefilePath)) {
    return undefined;
  }

  const makefileText = readTextFile(makefilePath);
  if (makefileText === undefined) {
    return undefined;
  }

  const installedDbTokens = parseMakeAssignments(makefileText).get("DB") || [];
  if (installedDbTokens.includes(installToken)) {
    return undefined;
  }

  return {
    installToken,
    makefilePath,
  };
}

function createMakefileInclusionDiagnostics(document) {
  const makefilePath = getSiblingMakefilePathForDocument(document);
  const installToken = getMakefileDbInstallTokenForDocument(document);
  if (!makefilePath || !installToken || !fs.existsSync(makefilePath)) {
    return [];
  }

  const makefileText = readTextFile(makefilePath);
  if (makefileText === undefined) {
    return [];
  }

  const installedDbTokens = parseMakeAssignments(makefileText).get("DB") || [];
  if (installedDbTokens.includes(installToken)) {
    return [];
  }

  if (isDatabaseDocument(document) && isReferencedBySiblingSubstitutionsFile(document)) {
    return [];
  }

  const range = getFileWarningRange(document);
  return [
    Object.assign(
      createDiagnostic(
        range.start,
        range.end,
        "This file is not included in Makefile.",
        vscode.DiagnosticSeverity.Warning,
      ),
      {
        code: "epics.makefile.notIncluded",
      },
    ),
  ];
}

function isReferencedBySiblingSubstitutionsFile(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  const extension = path.extname(document.uri.fsPath).toLowerCase();
  if (!MAKEFILE_INSTALLABLE_DATABASE_EXTENSIONS.has(extension)) {
    return false;
  }

  const normalizedTargetPath = normalizeFsPath(document.uri.fsPath);
  const parentDirectory = path.dirname(normalizedTargetPath);
  let siblingNames;
  try {
    siblingNames = fs.readdirSync(parentDirectory);
  } catch (error) {
    return false;
  }

  for (const siblingName of siblingNames) {
    const siblingPath = normalizeFsPath(path.join(parentDirectory, siblingName));
    if (siblingPath === normalizedTargetPath) {
      continue;
    }

    if (!SUBSTITUTION_EXTENSIONS.has(path.extname(siblingName).toLowerCase())) {
      continue;
    }

    const siblingText = readTextFile(siblingPath);
    if (!siblingText) {
      continue;
    }

    for (const block of extractSubstitutionBlocksWithRanges(siblingText)) {
      if (block.kind !== "file" || !block.templatePath) {
        continue;
      }

      if (
        doesSubstitutionTemplatePathReferenceTarget(
          siblingPath,
          block.templatePath,
          normalizedTargetPath,
        )
      ) {
        return true;
      }
    }
  }

  return false;
}

function doesSubstitutionTemplatePathReferenceTarget(
  substitutionsFilePath,
  templatePath,
  normalizedTargetPath,
) {
  const rawTemplatePath = String(templatePath || "").trim();
  if (!rawTemplatePath) {
    return false;
  }

  const directCandidatePath = normalizeFsPath(
    path.isAbsolute(rawTemplatePath)
      ? rawTemplatePath
      : path.resolve(path.dirname(substitutionsFilePath), rawTemplatePath),
  );
  if (directCandidatePath === normalizedTargetPath) {
    return true;
  }

  return getSubstitutionTemplateBasename(rawTemplatePath) === path.basename(normalizedTargetPath);
}

function getSubstitutionTemplateBasename(rawTemplatePath) {
  const normalizedPath = normalizePath(rawTemplatePath);
  const segments = normalizedPath.split("/").filter(Boolean);
  return segments[segments.length - 1] || normalizedPath;
}

function getFileWarningRange(document) {
  const endOffset = Math.min(document.getText().length, 1);
  return new vscode.Range(document.positionAt(0), document.positionAt(endOffset));
}

function resolveLocalMakeBuildDirectoryForDocument(document) {
  const makefilePath = getSiblingMakefilePathForDocument(document);
  if (!makefilePath || !fs.existsSync(makefilePath)) {
    return undefined;
  }

  if (isMakefileDocument(document) || isMakefileInstallableDatabaseSourceDocument(document)) {
    return path.dirname(makefilePath);
  }

  return undefined;
}

async function resolveDocumentForCommand(resourceUri) {
  if (Array.isArray(resourceUri) && resourceUri[0]?.scheme) {
    try {
      return await vscode.workspace.openTextDocument(resourceUri[0]);
    } catch (error) {
      return undefined;
    }
  }

  if (resourceUri?.scheme) {
    try {
      return await vscode.workspace.openTextDocument(resourceUri);
    } catch (error) {
      return undefined;
    }
  }

  return vscode.window.activeTextEditor?.document;
}

function resolveUriForCommand(resourceUri) {
  if (Array.isArray(resourceUri) && resourceUri[0]?.scheme) {
    return resourceUri[0];
  }

  if (resourceUri?.scheme) {
    return resourceUri;
  }

  if (resourceUri?.resourceUri?.scheme) {
    return resourceUri.resourceUri;
  }

  return vscode.window.activeTextEditor?.document?.uri;
}

async function addDbToMakefileForDocument(resourceUri) {
  const sourceDocument = await resolveDocumentForCommand(resourceUri);
  const pendingInsertion = getPendingMakefileDbInsertion(sourceDocument);
  if (!pendingInsertion) {
    vscode.window.showWarningMessage(
      "The active file cannot be added to a local DB += entry right now.",
    );
    return;
  }

  const makefileUri = vscode.Uri.file(pendingInsertion.makefilePath);
  const makefileDocument = await vscode.workspace.openTextDocument(makefileUri);
  const insertion = buildMakefileDbAppendText(
    makefileDocument.getText(),
    pendingInsertion.installToken,
    getDocumentEol(makefileDocument),
  );
  if (!insertion) {
    vscode.window.showErrorMessage(
      "Failed to update the local Makefile. Missing include $(TOP)/configure/CONFIG or include $(TOP)/configure/RULES.",
    );
    return;
  }
  const edit = new vscode.WorkspaceEdit();
  edit.insert(makefileUri, makefileDocument.positionAt(insertion.offset), insertion.text);
  const applied = await vscode.workspace.applyEdit(edit);
  if (!applied) {
    vscode.window.showErrorMessage("Failed to update the local Makefile.");
    return;
  }

  await vscode.window.showTextDocument(makefileDocument, {
    preview: false,
    preserveFocus: false,
  });
  vscode.window.showInformationMessage(
    `Added DB += ${pendingInsertion.installToken} to ${path.basename(pendingInsertion.makefilePath)}.`,
  );
}

function buildMakefileDbAppendText(existingText, installToken, eol) {
  const insertionOffset = findMakefileDbInsertionOffset(existingText);
  if (insertionOffset === undefined) {
    return undefined;
  }

  const normalizedEol = eol === "\r\n" ? "\r\n" : "\n";
  return {
    offset: insertionOffset,
    text: `DB += ${installToken}${normalizedEol}`,
  };
}

function findMakefileDbInsertionOffset(text) {
  const configMatch = findMakefileIncludeLine(text, "CONFIG");
  const rulesMatch = findMakefileIncludeLine(text, "RULES");
  if (!configMatch || !rulesMatch || rulesMatch.index <= configMatch.index) {
    return undefined;
  }

  return rulesMatch.index;
}

function findMakefileIncludeLine(text, includeName) {
  const pattern = new RegExp(
    `^\\s*include\\s+\\$\\(TOP\\)/configure/${includeName}\\s*$`,
    "m",
  );
  return pattern.exec(text);
}

async function buildWithLocalMakefile(resourceUri, outputChannel) {
  const document = await resolveDocumentForCommand(resourceUri);
  const buildDirectory = resolveLocalMakeBuildDirectoryForDocument(document);
  if (!buildDirectory) {
    await buildEpicsProject(resourceUri || document?.uri, outputChannel);
    return;
  }

  await runMakeCommands(
    buildDirectory,
    outputChannel,
    `Build ${path.basename(buildDirectory)}`,
    [["clean"]],
  );
}

async function cleanWithLocalMakefile(resourceUri, outputChannel) {
  const document = await resolveDocumentForCommand(resourceUri);
  const buildDirectory = resolveLocalMakeBuildDirectoryForDocument(document);
  if (!buildDirectory) {
    await cleanEpicsProject(resourceUri || document?.uri, outputChannel);
    return;
  }

  await runMakeCommands(
    buildDirectory,
    outputChannel,
    `Clean ${path.basename(buildDirectory)}`,
    [[]],
  );
}

async function buildEpicsProject(resourceUri, outputChannel) {
  const activeUri = resolveUriForCommand(resourceUri);
  const buildTarget = await resolveEpicsProjectBuildTarget(activeUri);
  if (!buildTarget) {
    vscode.window.showWarningMessage("No EPICS project root is available to build.");
    return;
  }

  await runMakeCommand(
    buildTarget.rootPath,
    outputChannel,
    `Build Project ${buildTarget.label}`,
  );
}

async function cleanEpicsProject(resourceUri, outputChannel) {
  const activeUri = resolveUriForCommand(resourceUri);
  const buildTarget = await resolveEpicsProjectBuildTarget(activeUri);
  if (!buildTarget) {
    vscode.window.showWarningMessage("No EPICS project root is available to clean.");
    return;
  }

  await runMakeCommands(
    buildTarget.rootPath,
    outputChannel,
    `Clean Project ${buildTarget.label}`,
    [["distclean"]],
  );
}

async function resolveEpicsProjectBuildTarget(activeUri) {
  if (activeUri?.scheme === "file") {
    const containingRootPath = findContainingEpicsProjectRootPath(activeUri.fsPath);
    if (containingRootPath) {
      return {
        rootPath: containingRootPath,
        label: path.basename(containingRootPath),
      };
    }
  }

  const epicsWorkspaceFolders = [];
  for (const workspaceFolder of vscode.workspace.workspaceFolders || []) {
    if (await isWorkspaceFolderEpicsProject(workspaceFolder)) {
      epicsWorkspaceFolders.push(workspaceFolder);
    }
  }

  if (!epicsWorkspaceFolders.length) {
    return undefined;
  }

  if (activeUri) {
    const activeWorkspaceFolder = vscode.workspace.getWorkspaceFolder(activeUri);
    if (
      activeWorkspaceFolder &&
      epicsWorkspaceFolders.some(
        (workspaceFolder) => workspaceFolder.uri.toString() === activeWorkspaceFolder.uri.toString(),
      )
    ) {
      return {
        rootPath: normalizeFsPath(activeWorkspaceFolder.uri.fsPath),
        label: activeWorkspaceFolder.name,
      };
    }
  }

  if (epicsWorkspaceFolders.length === 1) {
    return {
      rootPath: normalizeFsPath(epicsWorkspaceFolders[0].uri.fsPath),
      label: epicsWorkspaceFolders[0].name,
    };
  }

  const selectedWorkspace = await vscode.window.showQuickPick(
    epicsWorkspaceFolders.map((workspaceFolder) => ({
      label: workspaceFolder.name,
      description: workspaceFolder.uri.fsPath,
      workspaceFolder,
    })),
    {
      placeHolder: "Select an EPICS project root to build",
    },
  );
  if (!selectedWorkspace) {
    return undefined;
  }

  return {
    rootPath: normalizeFsPath(selectedWorkspace.workspaceFolder.uri.fsPath),
    label: selectedWorkspace.workspaceFolder.name,
  };
}

async function runMakeCommand(cwd, outputChannel, label) {
  return runMakeCommands(cwd, outputChannel, label, [[]]);
}

async function runMakeCommands(cwd, outputChannel, label, commandArgsList) {
  if (!cwd || !fs.existsSync(cwd)) {
    vscode.window.showErrorMessage(`Build directory does not exist: ${cwd || "<unknown>"}`);
    return;
  }

  outputChannel.show(true);
  outputChannel.appendLine("");
  outputChannel.appendLine(`=== ${label} ===`);
  outputChannel.appendLine(`cwd: ${cwd}`);
  outputChannel.appendLine("");

  const failures = [];
  for (const args of commandArgsList) {
    const result = await runSingleMakeCommand(cwd, outputChannel, args);
    if (result.started === false) {
      vscode.window.showErrorMessage(`Failed to start make: ${result.message}`);
      return;
    }

    if (result.signal) {
      vscode.window.showWarningMessage(`${label} was terminated by signal ${result.signal}.`);
      return;
    }

    if (result.code !== 0) {
      failures.push({
        args,
        code: result.code,
      });
    }
  }

  if (failures.length === 0) {
    vscode.window.showInformationMessage(`${label} finished successfully.`);
    return;
  }

  const failedCommandLabel = formatMakeCommand(failures[0].args);
  vscode.window.showErrorMessage(
    `${label} failed. ${failedCommandLabel} exited with code ${failures[0].code}.`,
  );
}

async function runMakeShellSequence(cwd, outputChannel, label, commandArgsList) {
  if (!cwd || !fs.existsSync(cwd)) {
    vscode.window.showErrorMessage(`Build directory does not exist: ${cwd || "<unknown>"}`);
    return;
  }

  const normalizedCommandArgsList = Array.isArray(commandArgsList)
    ? commandArgsList.map((args) => (Array.isArray(args) ? args : []))
    : [[]];
  const commandLabel = normalizedCommandArgsList
    .map((args) => formatMakeCommand(args))
    .join("; ");
  const invocation = buildMakeShellSequenceInvocation(normalizedCommandArgsList);

  outputChannel.show(true);
  outputChannel.appendLine("");
  outputChannel.appendLine(`=== ${label} ===`);
  outputChannel.appendLine(`cwd: ${cwd}`);
  outputChannel.appendLine("");
  outputChannel.appendLine(`$ ${commandLabel}`);

  const result = await new Promise((resolve) => {
    const child = childProcess.spawn(invocation.command, invocation.args, {
      cwd,
      env: process.env,
      shell: false,
    });

    child.stdout.on("data", (data) => {
      outputChannel.append(String(data));
    });
    child.stderr.on("data", (data) => {
      outputChannel.append(String(data));
    });
    child.on("error", (error) => {
      outputChannel.appendLine(`[error] Failed to start ${commandLabel}: ${error.message}`);
      resolve({
        started: false,
        message: error.message,
      });
    });
    child.on("close", (code, signal) => {
      if (signal) {
        outputChannel.appendLine(`[done] ${commandLabel} terminated by signal ${signal}`);
        outputChannel.appendLine("");
        resolve({
          started: true,
          signal,
        });
        return;
      }

      outputChannel.appendLine(`[done] ${commandLabel} exited with code ${code}`);
      outputChannel.appendLine("");
      resolve({
        started: true,
        code,
      });
    });
  });

  if (result.started === false) {
    vscode.window.showErrorMessage(`Failed to start make: ${result.message}`);
    return;
  }

  if (result.signal) {
    vscode.window.showWarningMessage(`${label} was terminated by signal ${result.signal}.`);
    return;
  }

  if (result.code === 0) {
    vscode.window.showInformationMessage(`${label} finished successfully.`);
    return;
  }

  vscode.window.showErrorMessage(`${label} failed. ${commandLabel} exited with code ${result.code}.`);
}

async function runSingleMakeCommand(cwd, outputChannel, args) {
  const normalizedArgs = Array.isArray(args) ? args : [];
  const commandLabel = formatMakeCommand(normalizedArgs);
  outputChannel.appendLine(`$ ${commandLabel}`);

  return new Promise((resolve) => {
    const child = childProcess.spawn("make", normalizedArgs, {
      cwd,
      env: process.env,
      shell: process.platform === "win32",
    });

    child.stdout.on("data", (data) => {
      outputChannel.append(String(data));
    });
    child.stderr.on("data", (data) => {
      outputChannel.append(String(data));
    });
    child.on("error", (error) => {
      outputChannel.appendLine(`[error] Failed to start ${commandLabel}: ${error.message}`);
      resolve({
        started: false,
        message: error.message,
      });
    });
    child.on("close", (code, signal) => {
      if (signal) {
        outputChannel.appendLine(`[done] ${commandLabel} terminated by signal ${signal}`);
        outputChannel.appendLine("");
        resolve({
          started: true,
          signal,
        });
        return;
      }

      outputChannel.appendLine(`[done] ${commandLabel} exited with code ${code}`);
      outputChannel.appendLine("");
      resolve({
        started: true,
        code,
      });
    });
  });
}

function formatMakeCommand(args) {
  if (!Array.isArray(args) || args.length === 0) {
    return "make";
  }

  return `make ${args.join(" ")}`;
}

function buildMakeShellSequenceInvocation(commandArgsList) {
  if (process.platform === "win32") {
    return {
      command: "cmd.exe",
      args: ["/d", "/s", "/c", buildWindowsMakeShellSequence(commandArgsList)],
    };
  }

  return {
    command: "/bin/sh",
    args: ["-lc", buildPosixMakeShellSequence(commandArgsList)],
  };
}

function buildPosixMakeShellSequence(commandArgsList) {
  return commandArgsList.map((args) => formatMakeCommand(args)).join("; ");
}

function buildWindowsMakeShellSequence(commandArgsList) {
  return commandArgsList.map((args) => formatMakeCommand(args)).join(" & ");
}

async function openUntitledImportedDatabaseDocument(fileName, text, preserveFocus) {
  const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath || os.tmpdir();
  const untitledUri = vscode.Uri.file(path.join(workspaceRoot, fileName)).with({
    scheme: "untitled",
  });
  const document = await vscode.workspace.openTextDocument(untitledUri);
  const editor = await vscode.window.showTextDocument(document, { preview: false, preserveFocus });
  await editor.edit((editBuilder) => {
    editBuilder.replace(
      new vscode.Range(new vscode.Position(0, 0), document.positionAt(document.getText().length)),
      text,
    );
  });
  await vscode.languages.setTextDocumentLanguage(editor.document, LANGUAGE_IDS.database);
}

async function openInProbeFromActiveEditor(workspaceIndex, runtimeMonitorController) {
  const widgetOptions =
    runtimeMonitorController?.getProbeWidgetCommandOptionsFromActiveWidget?.();
  if (widgetOptions) {
    await vscode.commands.executeCommand(OPEN_PROBE_WIDGET_COMMAND, widgetOptions);
    return;
  }

  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  if (!document) {
    return;
  }

  const target = isSubstitutionsDocument(document)
    ? undefined
    : await resolveProbeTargetAtActiveEditor(workspaceIndex);
  await vscode.commands.executeCommand(OPEN_PROBE_WIDGET_COMMAND, {
    recordName: target?.recordName || "",
  });
}

async function openInPvlistFromActiveEditor(workspaceIndex, runtimeMonitorController) {
  const widgetOptions =
    runtimeMonitorController?.getPvlistWidgetCommandOptionsFromActiveWidget?.();
  if (widgetOptions) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, widgetOptions);
    return;
  }

  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  if (!document) {
    return;
  }

  const sourceLabel = getDocumentDisplayLabel(document);
  if (isDatabaseDocument(document)) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "database",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: document.getText(),
    });
    return;
  }

  if (isSubstitutionsDocument(document)) {
    const expandedSource = await resolveExpandedSubstitutionsDatabaseSource(
      workspaceIndex,
      document,
      "open PV List",
    );
    if (!expandedSource) {
      return;
    }

    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "database",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: expandedSource.text,
    });
    return;
  }

  if (isStartupDocument(document)) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "pvlist",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: "",
    });
    return;
  }

  if (isProtocolDocument(document)) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "pvlist",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: "",
    });
    return;
  }

  if (document.languageId === LANGUAGE_IDS.dbd) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "pvlist",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: "",
    });
    return;
  }

  if (isPvlistDocument(document)) {
    await vscode.commands.executeCommand(OPEN_PVLIST_WIDGET_COMMAND, {
      sourceKind: "pvlist",
      sourceLabel,
      sourceDocumentUri: document.uri.toString(),
      sourceText: document.getText(),
    });
  }
}

async function openInMonitorFromActiveEditor(workspaceIndex, runtimeMonitorController) {
  const widgetOptions =
    runtimeMonitorController?.getMonitorWidgetCommandOptionsFromActiveWidget?.();
  if (widgetOptions) {
    await vscode.commands.executeCommand(OPEN_MONITOR_WIDGET_COMMAND, widgetOptions);
    return;
  }

  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  if (!document) {
    return;
  }

  let initialChannels = [];
  if (isSubstitutionsDocument(document)) {
    initialChannels = [];
  } else {
    const initialChannel = await resolveMonitorTargetAtActiveEditor(workspaceIndex);
    initialChannels = initialChannel ? [initialChannel] : [];
  }
  await vscode.commands.executeCommand(OPEN_MONITOR_WIDGET_COMMAND, {
    sourceLabel: getDocumentDisplayLabel(document),
    initialChannels,
  });
}

async function openInChannelGraphFromActiveEditor(workspaceIndex) {
  const source = await resolveChannelGraphSourceAtActiveEditor(workspaceIndex);
  const seedRecordName =
    source?.seedRecordName || await resolveMonitorTargetAtActiveEditor(workspaceIndex);

  const graphSession = {
    originNodeIds: seedRecordName ? [seedRecordName] : [],
    sources: source ? [source] : [],
    mode: source ? "static" : "dynamic",
    runtimeSession: undefined,
    disposed: false,
  };
  if (source) {
    graphSession.originNodeIds = [];
  }
  const getGraphMode = () => (graphSession.sources.length > 0 ? "static" : "dynamic");
  const panel = vscode.window.createWebviewPanel(
    CHANNEL_GRAPH_VIEW_TYPE,
    getChannelGraphPanelTitle(seedRecordName || source?.sourceLabel),
    vscode.ViewColumn.Active,
    {
      enableScripts: true,
      retainContextWhenHidden: true,
    },
  );
  let renderTimer = undefined;
  let htmlInitialized = false;

  const disposeRuntimeSession = async () => {
    if (!graphSession.runtimeSession) {
      return;
    }
    const activeSession = graphSession.runtimeSession;
    graphSession.runtimeSession = undefined;
    await activeSession.dispose();
  };

  const ensureRuntimeSession = () => {
    if (graphSession.runtimeSession) {
      return graphSession.runtimeSession;
    }

    graphSession.runtimeSession = new ChannelGraphRuntimeSession({
      initialOriginNodeIds: graphSession.originNodeIds,
      sourceEntries: graphSession.sources,
      onStateChange: () => {
        scheduleRender();
      },
    });
    void graphSession.runtimeSession.start();
    return graphSession.runtimeSession;
  };

  const buildCurrentGraphState = () => {
    graphSession.mode = getGraphMode();
    const sourceFiles = graphSession.sources.map((entry) => ({
      key: getChannelGraphSourceKey(entry),
      path: String(entry.sourcePath || entry.sourceLabel || ""),
    }));
    if (graphSession.mode === "dynamic") {
      return {
        ...ensureRuntimeSession().buildState(),
        sourceFiles,
      };
    }

    return {
      ...buildChannelGraphState(
      graphSession.sources.map((entry) => entry.sourceText).join("\n\n"),
      buildChannelGraphSourceLabel(graphSession.sources),
      graphSession.originNodeIds[0],
      graphSession.mode,
      graphSession.originNodeIds,
      ),
      sourceFiles,
    };
  };

  const renderPanel = (forceHtml = false) => {
    const graphState = buildCurrentGraphState();
    panel.title = getChannelGraphPanelTitle(
      graphSession.originNodeIds[0] || graphSession.sources[0]?.sourceLabel || graphState.sourceLabel,
    );
    if (!htmlInitialized || forceHtml) {
      htmlInitialized = true;
      panel.webview.html = buildChannelGraphWebviewHtml(panel.webview, graphState);
      return;
    }
    void panel.webview.postMessage({
      type: "setChannelGraphState",
      state: graphState,
    });
  };

  const scheduleRender = () => {
    if (graphSession.disposed || renderTimer) {
      return;
    }
    renderTimer = setTimeout(() => {
      renderTimer = undefined;
      if (!graphSession.disposed) {
        renderPanel();
      }
    }, 50);
  };

  renderPanel(true);
  panel.onDidDispose(() => {
    graphSession.disposed = true;
    if (renderTimer) {
      clearTimeout(renderTimer);
      renderTimer = undefined;
    }
    void disposeRuntimeSession();
  });
  panel.webview.onDidReceiveMessage(async (message) => {
    if (message?.type === "pickChannelGraphDatabaseFiles") {
      const selectedUris = await vscode.window.showOpenDialog({
        canSelectFiles: true,
        canSelectMany: true,
        filters: {
          "EPICS Database": ["db", "vdb", "template"],
        },
        openLabel: "Add Database Files",
      });
      if (!selectedUris?.length) {
        return;
      }

      const existingKeys = new Set(graphSession.sources.map(getChannelGraphSourceKey));
      for (const uri of selectedUris) {
        let sourceText;
        try {
          sourceText = Buffer.from(await vscode.workspace.fs.readFile(uri)).toString("utf8");
        } catch (error) {
          vscode.window.showErrorMessage(
            `Failed to read ${path.basename(uri.fsPath || uri.path)}: ${getErrorMessage(error)}`,
          );
          continue;
        }

        const sourceEntry = {
          sourceLabel: path.basename(uri.fsPath || uri.path),
          sourceText,
          seedRecordName: graphSession.originNodeIds[0],
          sourcePath: uri.fsPath || uri.path,
        };
        const key = getChannelGraphSourceKey(sourceEntry);
        if (existingKeys.has(key)) {
          continue;
        }
        existingKeys.add(key);
        graphSession.sources.push(sourceEntry);
      }

      if (getGraphMode() === "static") {
        await disposeRuntimeSession();
      }
      renderPanel();
      return;
    }

    if (message?.type === "removeChannelGraphDatabaseFile") {
      const key = String(message.key || "");
      if (!key) {
        return;
      }
      const previousMode = getGraphMode();
      graphSession.sources = graphSession.sources.filter(
        (entry) => getChannelGraphSourceKey(entry) !== key,
      );
      if (previousMode !== getGraphMode()) {
        const runtimeSession = ensureRuntimeSession();
        for (const originNodeId of graphSession.originNodeIds) {
          await runtimeSession.addOriginNode(originNodeId);
        }
      }
      renderPanel();
      return;
    }

    if (message?.type === "addChannelGraphOrigin") {
      const nodeId = String(message.nodeId || "").trim();
      if (!nodeId) {
        return;
      }
      if (!graphSession.originNodeIds.includes(nodeId)) {
        graphSession.originNodeIds.push(nodeId);
      }
      if (getGraphMode() === "dynamic") {
        const runtimeSession = ensureRuntimeSession();
        await runtimeSession.addOriginNode(nodeId);
      }
      renderPanel();
      return;
    }

    if (message?.type === "clearChannelGraph") {
      graphSession.originNodeIds = [];
      if (getGraphMode() === "dynamic") {
        const runtimeSession = ensureRuntimeSession();
        await runtimeSession.clearGraph();
      }
      renderPanel();
      return;
    }

    if (message?.type === "expandChannelGraphNode") {
      if (getGraphMode() !== "dynamic" || !message.nodeId) {
        return;
      }
      const runtimeSession = ensureRuntimeSession();
      await runtimeSession.expandNode(String(message.nodeId));
      renderPanel();
    }
  });
}

async function resolveChannelGraphSourceAtActiveEditor(workspaceIndex) {
  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  const position = editor?.selection?.active;
  if (!document) {
    return undefined;
  }

  if (isDatabaseDocument(document)) {
    const text = document.getText();
    const offset = position ? document.offsetAt(position) : -1;
    const tocTarget = position
      ? extractDatabaseTocEntries(text).find(
        (entry) => offset >= entry.nameStart && offset <= entry.nameEnd,
      )
      : undefined;
    const declarationTarget = position
      ? extractRecordDeclarations(text).find(
        (declaration) => offset >= declaration.nameStart && offset <= declaration.nameEnd,
      )
      : undefined;
    if (position) {
      const probeTarget = await resolveProbeTargetAtActiveEditor(workspaceIndex);
      if (
        probeTarget?.recordName &&
        !declarationTarget &&
        !tocTarget &&
        probeTarget.recordName !== resolveProbeTargetFromDatabaseToc(document, position)?.recordName
      ) {
        const externalSource = await resolveChannelGraphSourceForRecordName(
          workspaceIndex,
          document,
          probeTarget.recordName,
        );
        if (externalSource) {
          return externalSource;
        }
      }
    }
    return {
      sourceLabel: getDocumentDisplayLabel(document),
      sourceText: text,
      seedRecordName: tocTarget?.recordName || declarationTarget?.name,
      sourcePath: document.uri.scheme === "file" ? document.uri.fsPath : undefined,
    };
  }

  const seedRecordName = await resolveMonitorTargetAtActiveEditor(workspaceIndex);
  if (!seedRecordName) {
    return undefined;
  }

  return resolveChannelGraphSourceForRecordName(workspaceIndex, document, seedRecordName);
}

async function resolveChannelGraphSourceForRecordName(
  workspaceIndex,
  document,
  recordName,
) {
  if (!recordName) {
    return undefined;
  }

  const snapshot = await workspaceIndex.getSnapshot();
  const definition = getRecordDefinitionsForName(snapshot, document, recordName)[0];
  if (!definition?.absolutePath) {
    return undefined;
  }

  const sourceText = readTextFile(definition.absolutePath);
  if (sourceText === undefined) {
    return undefined;
  }

  return {
    sourceLabel: path.basename(definition.absolutePath),
    sourceText,
    seedRecordName: recordName,
    sourcePath: definition.absolutePath,
  };
}

function buildChannelGraphSourceLabel(sources) {
  const labels = (sources || [])
    .map((entry) => String(entry?.sourceLabel || "").trim())
    .filter(Boolean);
  if (!labels.length) {
    return "EPICS Channel Graph";
  }
  if (labels.length === 1) {
    return labels[0];
  }
  return `${labels[0]} + ${labels.length - 1} more`;
}

function getChannelGraphSourceKey(source) {
  return String(source?.sourcePath || `${source?.sourceLabel || ""}\u0000${source?.sourceText || ""}`);
}

function getChannelGraphPanelTitle(label) {
  const normalized = String(label || "").trim();
  return normalized ? `Channel Graph: ${normalized}` : "EPICS Channel Graph";
}

function resolveChannelGraphRuntimeWorkspaceFolder(sourceEntries) {
  for (const sourceEntry of sourceEntries || []) {
    if (!sourceEntry?.sourcePath) {
      continue;
    }
    const workspaceFolder = vscode.workspace.getWorkspaceFolder(
      vscode.Uri.file(sourceEntry.sourcePath),
    );
    if (workspaceFolder) {
      return workspaceFolder;
    }
  }

  const activeDocument = vscode.window.activeTextEditor?.document;
  if (activeDocument?.uri) {
    return vscode.workspace.getWorkspaceFolder(activeDocument.uri);
  }

  return vscode.workspace.workspaceFolders?.[0];
}

function getChannelGraphCaReadOptions(channel) {
  if (String(channel?.getDbrTypeStr?.() || "") !== "DBR_ENUM") {
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

function updateChannelGraphCaEnumChoicesCache(entry, dbrData) {
  if (!entry || !Array.isArray(dbrData?.strings)) {
    return;
  }
  const validCount = Number(dbrData?.number_of_string_used);
  entry.caEnumChoices = Number.isFinite(validCount) && validCount > 0
    ? dbrData.strings.slice(0, validCount).map((choice) => String(choice ?? ""))
    : dbrData.strings.map((choice) => String(choice ?? ""));
}

function updateChannelGraphPvaEnumChoicesCache(entry, pvaData) {
  const value = pvaData?.value;
  if (
    value &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    Array.isArray(value.choices)
  ) {
    entry.pvaEnumChoices = value.choices.map((choice) => String(choice ?? ""));
  }
}

class ChannelGraphRuntimeSession {
  constructor({ initialOriginNodeIds = [], sourceEntries, onStateChange }) {
    this.originNodeIds = new Set(
      (Array.isArray(initialOriginNodeIds) ? initialOriginNodeIds : [])
        .map((nodeId) => String(nodeId || "").trim())
        .filter(Boolean),
    );
    this.sourceEntries = Array.isArray(sourceEntries) ? sourceEntries : [];
    this.onStateChange = typeof onStateChange === "function" ? onStateChange : () => {};
    this.runtimeLibrary = safeRequireRuntimeLibrary();
    this.protocol = "ca";
    this.status = "idle";
    this.errorMessage = "";
    this.context = undefined;
    this.initializationPromise = undefined;
    this.nodeConnectPromises = new Map();
    this.nodeSessions = new Map();
    this.nodesById = new Map();
    this.edges = [];
    this.seenEdges = new Set();
    this.resolvedNodeIds = new Set();
    this.disposed = false;
    this.stateChangeScheduled = false;
  }

  buildState() {
    const adjacency = Object.create(null);
    for (const edge of this.edges) {
      adjacency[edge.fromId] = adjacency[edge.fromId] || [];
      adjacency[edge.toId] = adjacency[edge.toId] || [];
      if (!adjacency[edge.fromId].includes(edge.toId)) {
        adjacency[edge.fromId].push(edge.toId);
      }
      if (!adjacency[edge.toId].includes(edge.fromId)) {
        adjacency[edge.toId].push(edge.fromId);
      }
    }

    let message = this.errorMessage;
    if (!message) {
      if (this.originNodeIds.size === 0) {
        message = "Enter a channel name to start the Channel Graph.";
      } else if (this.status === "connecting") {
        const firstOriginNodeId = [...this.originNodeIds][0];
        message = `Connecting to EPICS runtime and expanding "${firstOriginNodeId}"...`;
      } else if (this.status === "connected" && this.edges.length === 0) {
        const firstOriginNodeId = [...this.originNodeIds][0];
        message = `No runtime link relationships were found for "${firstOriginNodeId}".`;
      }
    }

    return {
      mode: "dynamic",
      sourceLabel: this.originNodeIds.size > 0
        ? `Runtime (${this.protocol.toUpperCase()}): ${[...this.originNodeIds][0]}`
        : `Runtime (${this.protocol.toUpperCase()})`,
      seedRecordName: [...this.originNodeIds][0],
      originNodeIds: [...this.originNodeIds],
      message,
      allowAddDatabaseFiles: true,
      nodes: [...this.nodesById.values()],
      edges: [...this.edges],
      adjacency,
    };
  }

  scheduleStateChange() {
    if (this.disposed || this.stateChangeScheduled) {
      return;
    }
    this.stateChangeScheduled = true;
    setTimeout(() => {
      this.stateChangeScheduled = false;
      if (!this.disposed) {
        this.onStateChange();
      }
    }, 0);
  }

  ensureNode(nodeId, overrides = {}) {
    const normalizedNodeId = String(nodeId || "").trim();
    if (!normalizedNodeId) {
      return undefined;
    }
    const existingNode = this.nodesById.get(normalizedNodeId);
    if (existingNode) {
      Object.assign(existingNode, overrides);
      return existingNode;
    }
    const node = {
      id: normalizedNodeId,
      name: normalizedNodeId.startsWith("__value__:") ? "" : normalizedNodeId,
      recordType: "",
      scanValue: "",
      valueText: "",
      external: true,
      ...overrides,
    };
    this.nodesById.set(normalizedNodeId, node);
    return node;
  }

  addEdge(edge) {
    const edgeKey = `${edge.fromId}\u0000${edge.toId}\u0000${edge.label}`;
    if (this.seenEdges.has(edgeKey)) {
      return;
    }
    this.seenEdges.add(edgeKey);
    this.edges.push(edge);
  }

  async start() {
    if (this.initializationPromise) {
      return this.initializationPromise;
    }

    this.status = "connecting";
    this.scheduleStateChange();
    this.initializationPromise = (async () => {
      if (!this.runtimeLibrary?.Context) {
        this.status = "error";
        this.errorMessage = "epics-tca is not available for dynamic Channel Graph.";
        this.scheduleStateChange();
        return;
      }

      const workspaceFolder = resolveChannelGraphRuntimeWorkspaceFolder(this.sourceEntries);
      const loadedConfig = workspaceFolder
        ? loadProjectRuntimeConfiguration(
          path.join(workspaceFolder.uri.fsPath, PROJECT_RUNTIME_CONFIG_FILE_NAME),
        )
        : {
          exists: false,
          config: getDefaultProjectRuntimeConfiguration(),
        };
      this.protocol = normalizeRuntimeProtocol(loadedConfig.config.protocol);
      const context = new this.runtimeLibrary.Context(
        createRuntimeEnvironmentFromProjectConfiguration(loadedConfig.config),
        "warning",
      );
      await context.initialize();
      if (this.disposed) {
        try {
          context.destroyHard?.();
        } catch (error) {
          // Ignore cleanup failures after disposal.
        }
        return;
      }
      this.context = context;
      this.status = "connected";
      for (const originNodeId of this.originNodeIds) {
        this.ensureNode(originNodeId, {
          name: originNodeId,
          external: false,
        });
      }
      for (const originNodeId of this.originNodeIds) {
        await this.ensureNodeSession(originNodeId);
        await this.populateRuntimeNodeMetadata(originNodeId);
        await this.expandNode(originNodeId);
      }
      this.scheduleStateChange();
    })().catch((error) => {
      this.status = "error";
      this.errorMessage = `Dynamic resolution failed: ${getErrorMessage(error)}`;
      this.scheduleStateChange();
    });

    return this.initializationPromise;
  }

  async addOriginNode(nodeId) {
    const normalizedNodeId = String(nodeId || "").trim();
    if (!normalizedNodeId) {
      return;
    }

    this.originNodeIds.add(normalizedNodeId);
    this.ensureNode(normalizedNodeId, {
      name: normalizedNodeId,
      external: false,
    });
    this.errorMessage = "";
    await this.start();
    if (!this.context || this.disposed) {
      this.scheduleStateChange();
      return;
    }
    await this.ensureNodeSession(normalizedNodeId);
    await this.populateRuntimeNodeMetadata(normalizedNodeId);
    await this.expandNode(normalizedNodeId);
    this.scheduleStateChange();
  }

  async clearGraph() {
    this.originNodeIds.clear();
    this.resolvedNodeIds.clear();
    this.edges = [];
    this.seenEdges.clear();
    this.nodesById.clear();
    this.errorMessage = "";
    await Promise.all(
      [...this.nodeSessions.values()].map((session) => this.cleanupNodeSession(session)),
    );
    this.nodeSessions.clear();
    this.nodeConnectPromises.clear();
    this.scheduleStateChange();
  }

  async expandNode(nodeId) {
    const normalizedNodeId = String(nodeId || "").trim();
    if (
      !normalizedNodeId ||
      normalizedNodeId.startsWith("__value__:") ||
      this.resolvedNodeIds.has(normalizedNodeId)
    ) {
      return;
    }

    await this.start();
    if (!this.context || this.disposed) {
      return;
    }

    this.resolvedNodeIds.add(normalizedNodeId);
    await this.ensureNodeSession(normalizedNodeId);
    await this.populateRuntimeNodeMetadata(normalizedNodeId);
    const node = this.nodesById.get(normalizedNodeId);
    const recordType = String(node?.recordType || "").trim();
    if (!recordType) {
      this.scheduleStateChange();
      return;
    }

    const linkFieldNames = getRuntimeProbeFieldNamesForRecordType(recordType).filter((fieldName) => {
      const dbfType = getRuntimeProbeFieldTypeForRecordType(recordType, fieldName);
      return LINK_DBF_TYPES.has(dbfType) || isLinkField(fieldName);
    });

    for (const fieldName of linkFieldNames) {
      const fieldValue = await this.readRuntimeFieldValue(normalizedNodeId, fieldName);
      const rawValue = String(fieldValue?.rawValue || "").trim();
      if (!rawValue) {
        continue;
      }

      const target = parseChannelGraphLinkTarget(rawValue);
      if (!target) {
        continue;
      }

      const targetNodeId = target.recordName
        ? target.recordName
        : `__value__:${normalizedNodeId}:${fieldName}:${target.rawValue}`;
      const targetNodeName = target.recordName || target.rawValue;
      if (!targetNodeName) {
        continue;
      }

      this.ensureNode(targetNodeId, {
        name: targetNodeName,
        external: !target.recordName,
      });

      const label = target.targetField
        ? `${fieldName}:${target.targetField}`
        : fieldName;
      const dbfType = getRuntimeProbeFieldTypeForRecordType(recordType, fieldName);
      this.addEdge(
        dbfType === "DBF_INLINK" || isChannelGraphInputField(fieldName)
          ? { fromId: targetNodeId, toId: normalizedNodeId, label }
          : { fromId: normalizedNodeId, toId: targetNodeId, label },
      );

      if (target.recordName) {
        void this.ensureNodeSession(targetNodeId).then(() =>
          this.populateRuntimeNodeMetadata(targetNodeId),
        );
      }
    }

    this.scheduleStateChange();
  }

  async ensureNodeSession(nodeId) {
    const normalizedNodeId = String(nodeId || "").trim();
    if (!normalizedNodeId || normalizedNodeId.startsWith("__value__:")) {
      return false;
    }
    if (this.nodeSessions.has(normalizedNodeId)) {
      return true;
    }
    if (this.nodeConnectPromises.has(normalizedNodeId)) {
      return this.nodeConnectPromises.get(normalizedNodeId);
    }

    const connectPromise = (async () => {
      if (!this.context) {
        return false;
      }

      let channel;
      let session;
      try {
        channel = await this.context.createChannel(normalizedNodeId, this.protocol, 1.5);
        if (!channel) {
          return false;
        }
        session = {
          nodeId: normalizedNodeId,
          channel,
          monitor: undefined,
          protocol: this.protocol,
          caEnumChoices: undefined,
          pvaEnumChoices: undefined,
        };
        const node = this.ensureNode(normalizedNodeId, {
          name: normalizedNodeId,
          external: false,
        });

        channel.setDestroySoftCallback?.(() => {
          if (this.disposed) {
            return;
          }
          const activeNode = this.nodesById.get(normalizedNodeId);
          if (activeNode) {
            activeNode.valueText = "";
          }
          this.scheduleStateChange();
        });
        channel.setDestroyHardCallback?.(() => {
          if (this.disposed) {
            return;
          }
          this.nodeSessions.delete(normalizedNodeId);
          const activeNode = this.nodesById.get(normalizedNodeId);
          if (activeNode) {
            activeNode.valueText = "";
          }
          this.scheduleStateChange();
        });

        const monitorPromise = this.protocol === "pva"
          ? channel.createMonitorPva(1.5, "", (activeMonitor) => {
            this.handleNodeMonitorUpdate(normalizedNodeId, session, activeMonitor);
          })
          : (() => {
            const caReadOptions = getChannelGraphCaReadOptions(channel);
            if (caReadOptions) {
              return channel.createMonitor(
                1.5,
                (activeMonitor) => {
                  this.handleNodeMonitorUpdate(normalizedNodeId, session, activeMonitor);
                },
                caReadOptions.dbrType,
                caReadOptions.valueCount,
              );
            }
            return channel.createMonitor(1.5, (activeMonitor) => {
              this.handleNodeMonitorUpdate(normalizedNodeId, session, activeMonitor);
            });
          })();

        session.monitor = await monitorPromise;
        if (this.disposed) {
          await this.cleanupNodeSession(session);
          return false;
        }
        this.nodeSessions.set(normalizedNodeId, session);
        if (node) {
          node.external = false;
        }
        this.scheduleStateChange();
        return true;
      } catch (error) {
        const node = this.nodesById.get(normalizedNodeId);
        if (node && !this.resolvedNodeIds.has(normalizedNodeId)) {
          node.external = true;
        }
        await this.cleanupNodeSession(session || { channel });
        this.scheduleStateChange();
        return false;
      } finally {
        this.nodeConnectPromises.delete(normalizedNodeId);
      }
    })();

    this.nodeConnectPromises.set(normalizedNodeId, connectPromise);
    return connectPromise;
  }

  handleNodeMonitorUpdate(nodeId, session, monitor) {
    if (this.disposed) {
      return;
    }
    const node = this.nodesById.get(nodeId);
    if (!node) {
      return;
    }

    if (session.protocol === "pva") {
      const pvaData = monitor?.getPvaData?.();
      updateChannelGraphPvaEnumChoicesCache(session, pvaData);
      node.valueText = formatRuntimeValue(
        getPvaRuntimeDisplayValue(pvaData, session.pvaEnumChoices),
      );
    } else {
      const dbrData = monitor?.getChannel?.().getDbrData?.();
      updateChannelGraphCaEnumChoicesCache(session, dbrData);
      node.valueText = formatRuntimeValue(
        getCaRuntimeDisplayValue(session, dbrData),
      );
    }

    this.scheduleStateChange();
  }

  async populateRuntimeNodeMetadata(nodeId) {
    const normalizedNodeId = String(nodeId || "").trim();
    if (!normalizedNodeId || normalizedNodeId.startsWith("__value__:")) {
      return;
    }

    const node = this.ensureNode(normalizedNodeId, {
      name: normalizedNodeId,
    });
    const [recordType, scanValue] = await Promise.all([
      this.readRuntimeFieldValue(normalizedNodeId, "RTYP"),
      this.readRuntimeFieldValue(normalizedNodeId, "SCAN"),
    ]);
    if (node) {
      if (recordType?.displayValue) {
        node.recordType = recordType.displayValue;
      }
      if (scanValue?.displayValue) {
        node.scanValue = scanValue.displayValue;
      }
    }
    this.scheduleStateChange();
  }

  async readRuntimeFieldValue(recordName, fieldName) {
    if (!this.context || this.disposed) {
      return undefined;
    }

    let channel;
    try {
      channel = await this.context.createChannel(
        `${recordName}.${fieldName}`,
        this.protocol,
        1.5,
      );
      if (!channel) {
        return undefined;
      }

      if (this.protocol === "pva") {
        const pvaData = await channel.getPva(1.5, "");
        const displayValue = formatRuntimeValue(
          getPvaRuntimeDisplayValue(pvaData, undefined),
        );
        return {
          rawValue: displayValue,
          displayValue,
        };
      }

      const readOptions = getChannelGraphCaReadOptions(channel);
      const dbrData = await channel.get(
        1.5,
        readOptions?.dbrType,
        readOptions?.valueCount,
      );
      const entry = {
        channel,
        caEnumChoices: undefined,
      };
      updateChannelGraphCaEnumChoicesCache(entry, dbrData);
      const displayValue = formatRuntimeValue(
        getCaRuntimeDisplayValue(entry, dbrData),
      );
      return {
        rawValue: displayValue,
        displayValue,
      };
    } catch (error) {
      return undefined;
    } finally {
      try {
        await channel?.destroyHard?.();
      } catch (error) {
        // Ignore best-effort field read cleanup failures.
      }
    }
  }

  async cleanupNodeSession(session) {
    if (!session) {
      return;
    }
    try {
      await session.monitor?.destroySoft?.();
    } catch (error) {
      // Ignore.
    }
    try {
      await session.channel?.destroyHard?.();
    } catch (error) {
      // Ignore.
    }
  }

  async dispose() {
    this.disposed = true;
    this.onStateChange = () => {};
    await Promise.all([...this.nodeSessions.values()].map((session) => this.cleanupNodeSession(session)));
    this.nodeSessions.clear();
    this.nodeConnectPromises.clear();
    try {
      await this.context?.destroyHard?.();
    } catch (error) {
      // Ignore runtime context cleanup failures.
    }
  }
}

async function resolveMonitorTargetAtActiveEditor(workspaceIndex) {
  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  const position = editor?.selection?.active;
  if (!document) {
    return undefined;
  }

  if (isDatabaseDocument(document) || isStartupDocument(document)) {
    const target = await resolveProbeTargetAtActiveEditor(workspaceIndex);
    return String(target?.recordName || "").trim() || undefined;
  }

  if (isPvlistDocument(document) && position) {
    const trimmedLine = String(document.lineAt(position.line).text || "").trim();
    if (
      trimmedLine &&
      !trimmedLine.startsWith("#") &&
      !/^[A-Za-z_][A-Za-z0-9_]*\s*=/.test(trimmedLine) &&
      !trimmedLine.includes("=") &&
      !/\s/.test(trimmedLine)
    ) {
      return trimmedLine;
    }
    return undefined;
  }

  if (isProbeDocument(document)) {
    for (const rawLine of document.getText().split(/\r?\n/)) {
      const trimmedLine = String(rawLine || "").trim();
      if (!trimmedLine || trimmedLine.startsWith("#")) {
        continue;
      }
      if (/\s/.test(trimmedLine) || /\$\(|\$\{/.test(trimmedLine)) {
        return undefined;
      }
      return trimmedLine;
    }
  }

  return undefined;
}

async function resolveProbeTargetAtActiveEditor(workspaceIndex) {
  const editor = vscode.window.activeTextEditor;
  const document = editor?.document;
  const position = editor?.selection?.active;
  if (!document || !position) {
    return undefined;
  }

  if (isDatabaseDocument(document)) {
    const tocTarget = resolveProbeTargetFromDatabaseToc(document, position);
    if (tocTarget) {
      return tocTarget;
    }

    const text = document.getText();
    const offset = document.offsetAt(position);
    const declaration = extractRecordDeclarations(text).find(
      (candidate) => offset >= candidate.nameStart && offset <= candidate.nameEnd,
    );
    if (declaration) {
      const recordName = resolveProbeRecordNameFromDatabaseText(text, declaration.name);
      if (!recordName) {
        return undefined;
      }

      return {
        recordName,
        recordType: declaration.recordType,
      };
    }

    const snapshot = await workspaceIndex.getSnapshot();
    const fieldDeclaration = getRecordScopedFieldDeclarationAtPosition(
      snapshot,
      document,
      position,
    );
    if (fieldDeclaration) {
      const linkedTarget = resolveProbeTargetFromLinkedDatabaseField(
        snapshot,
        document,
        fieldDeclaration.value,
      );
      if (linkedTarget) {
        return linkedTarget;
      }
    }
  }

  if (isStartupDocument(document)) {
    const argument = getStartupDbpfArgumentAtPosition(document, position);
    if (!argument) {
      return undefined;
    }

    const snapshot = await workspaceIndex.getSnapshot();
    const loadedDefinitionsByName = getStartupLoadedRecordDefinitionMap(
      snapshot,
      document,
      position,
    );
    const recordName = extractLinkedRecordCandidates(argument.value).find((candidate) =>
      loadedDefinitionsByName.has(candidate),
    );
    if (!recordName) {
      return undefined;
    }

    const definitions = loadedDefinitionsByName.get(recordName) || [];
    return {
      recordName,
      recordType: definitions[0]?.recordType,
    };
  }

  if (isSubstitutionsDocument(document)) {
    const snapshot = await workspaceIndex.getSnapshot();
    return resolveProbeTargetFromSubstitutionsDocument(snapshot, document, position);
  }

  return undefined;
}

function resolveProbeTargetFromLinkedDatabaseField(snapshot, document, fieldValue) {
  const match = resolveLinkedRecordDefinitionMatch(snapshot, document, fieldValue);
  if (!match) {
    return undefined;
  }

  if (!containsEpicsMacroReference(match.recordName)) {
    return {
      recordName: match.recordName,
      recordType: match.definitions[0]?.recordType,
    };
  }

  for (const definition of match.definitions) {
    const resolvedRecordName = resolveRuntimeRecordNameForDefinition(document, definition);
    if (!resolvedRecordName) {
      continue;
    }

    return {
      recordName: resolvedRecordName,
      recordType: definition.recordType,
    };
  }

  return undefined;
}

function resolveProbeTargetFromDatabaseToc(document, position) {
  const text = document.getText();
  const offset = document.offsetAt(position);
  const tocEntry = extractDatabaseTocEntries(text).find(
    (entry) => offset >= entry.nameStart && offset <= entry.nameEnd,
  );
  if (!tocEntry) {
    return undefined;
  }

  const recordName = resolveProbeRecordNameFromDatabaseText(text, tocEntry.recordName);
  if (!recordName) {
    return undefined;
  }

  return {
    recordName,
    recordType: tocEntry.recordType,
  };
}

function resolveProbeRecordNameFromDatabaseText(documentText, recordName) {
  const normalizedRecordName = String(recordName || "").trim();
  if (!normalizedRecordName) {
    return undefined;
  }
  if (!containsEpicsMacroReference(normalizedRecordName)) {
    return normalizedRecordName;
  }
  return resolveDatabaseRecordNameFromToc(
    normalizedRecordName,
    extractDatabaseTocMacroAssignments(documentText),
  );
}

function resolveRuntimeRecordNameForDefinition(activeDocument, definition) {
  const rawRecordName = String(definition?.name || "").trim();
  if (!rawRecordName) {
    return undefined;
  }
  if (!containsEpicsMacroReference(rawRecordName)) {
    return rawRecordName;
  }
  if (!definition?.absolutePath) {
    return undefined;
  }

  const definitionText = getDefinitionDocumentText(activeDocument, definition.absolutePath);
  if (definitionText === undefined) {
    return undefined;
  }

  return resolveProbeRecordNameFromDatabaseText(definitionText, rawRecordName);
}

function getDefinitionDocumentText(activeDocument, absolutePath) {
  const normalizedAbsolutePath = normalizeFsPath(absolutePath);
  if (
    activeDocument?.uri?.scheme === "file" &&
    normalizeFsPath(activeDocument.uri.fsPath) === normalizedAbsolutePath
  ) {
    return activeDocument.getText();
  }

  const openDocument = vscode.workspace.textDocuments.find(
    (document) =>
      document.uri?.scheme === "file" &&
      normalizeFsPath(document.uri.fsPath) === normalizedAbsolutePath,
  );
  if (openDocument) {
    return openDocument.getText();
  }

  return readTextFile(normalizedAbsolutePath);
}

function resolveDatabaseRecordNameFromToc(recordName, macroAssignments) {
  const normalizedRecordName = String(recordName || "").trim();
  if (!normalizedRecordName) {
    return undefined;
  }
  if (!containsEpicsMacroReference(normalizedRecordName)) {
    return normalizedRecordName;
  }

  const resolvedMacroValues = new Map();
  if (macroAssignments instanceof Map) {
    for (const [macroName, assignment] of macroAssignments.entries()) {
      if (assignment?.hasAssignment) {
        resolvedMacroValues.set(macroName, String(assignment.value || ""));
      }
    }
  }

  const resolvedRecordName = resolveDatabaseRecordNameFromMacroValues(
    normalizedRecordName,
    resolvedMacroValues,
    new Set(),
  );
  if (!resolvedRecordName || containsEpicsMacroReference(resolvedRecordName)) {
    return undefined;
  }

  return resolvedRecordName;
}

function resolveDatabaseRecordNameFromMacroValues(recordName, macroValues, stack) {
  let unresolved = false;
  const expandedRecordName = String(recordName || "").replace(
    /\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}|\$([A-Za-z_][A-Za-z0-9_.-]*)/g,
    (match, parenName, defaultValue, braceName, bareName) => {
      const macroName = parenName || braceName || bareName;
      if (!macroValues.has(macroName) || stack.has(macroName)) {
        unresolved = true;
        return match;
      }

      const nextStack = new Set(stack);
      nextStack.add(macroName);
      const resolvedValue = resolveDatabaseRecordNameFromMacroValues(
        macroValues.get(macroName),
        macroValues,
        nextStack,
      );
      if (resolvedValue === undefined) {
        unresolved = true;
        return match;
      }
      return resolvedValue;
    },
  );

  if (unresolved) {
    return undefined;
  }

  return expandedRecordName;
}

function resolveProbeTargetFromSubstitutionsDocument(snapshot, document, position) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  const text = document.getText();
  const offset = document.offsetAt(position);
  let globalMacros = new Map();
  let fallbackTarget;
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const releaseVariables = project ? project.releaseVariables : new Map();

  for (const block of extractSubstitutionBlocksWithRanges(text)) {
    if (block.kind === "global") {
      globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body));
      continue;
    }

    if (block.kind !== "file" || !block.templatePath) {
      continue;
    }

    const templateAbsolutePath = resolveSubstitutionTemplateAbsolutePathForDocument(
      snapshot,
      document,
      block.templatePath,
    );
    if (!templateAbsolutePath) {
      continue;
    }

    const templateText = readTextFile(templateAbsolutePath);
    if (templateText === undefined) {
      continue;
    }
    const templateTextWithoutToc = removeDatabaseTocBlock(templateText);
    const parsedRows = parseSubstitutionFileBlockRowsDetailed(block.body, block.bodyStart);
    const candidateRows = parsedRows.rows.length > 0
      ? parsedRows.rows.map((row) => ({
        rangeStart: row.rangeStart,
        rangeEnd: row.rangeEnd,
        macros: mergeMacroMaps(globalMacros, row.assignments),
      }))
      : [{
        rangeStart: block.bodyStart,
        rangeEnd: block.bodyEnd,
        macros: new Map(globalMacros),
      }];
    if (!fallbackTarget && candidateRows.length > 0) {
      const expandedFallbackText = normalizeTextEol(
        expandEpicsValue(
          templateTextWithoutToc,
          [candidateRows[0].macros, releaseVariables, process.env],
        ),
        "\n",
      ).trimEnd();
      const fallbackDeclaration = extractRecordDeclarations(expandedFallbackText)[0];
      if (fallbackDeclaration?.name) {
        fallbackTarget = {
          recordName: fallbackDeclaration.name,
          recordType: fallbackDeclaration.recordType,
        };
      }
    }
    const matchingRow =
      candidateRows.find((row) => offset >= row.rangeStart && offset <= row.rangeEnd) ||
      (block.templatePathStart !== undefined &&
      block.templatePathEnd !== undefined &&
      offset >= block.templatePathStart &&
      offset <= block.templatePathEnd
        ? candidateRows[0]
        : undefined);
    if (!matchingRow) {
      continue;
    }

    const expandedText = normalizeTextEol(
      expandEpicsValue(
        templateTextWithoutToc,
        [matchingRow.macros, releaseVariables, process.env],
      ),
      "\n",
    ).trimEnd();
    const declaration = extractRecordDeclarations(expandedText)[0];
    if (!declaration?.name) {
      continue;
    }

    return {
      recordName: declaration.name,
      recordType: declaration.recordType,
    };
  }

  return fallbackTarget;
}

function buildChannelGraphState(
  sourceText,
  sourceLabel,
  seedRecordName,
  mode = "static",
  originNodeIds = seedRecordName ? [seedRecordName] : [],
) {
  if (!sourceText) {
    return {
      mode,
      sourceLabel: sourceLabel || "EPICS Channel Graph",
      seedRecordName,
      originNodeIds,
      message: "No database source is available for Channel Graph.",
      allowAddDatabaseFiles: true,
      nodes: [],
      edges: [],
      adjacency: {},
    };
  }

  const declarations = extractRecordDeclarations(sourceText);
  if (!declarations.length) {
    return {
      mode,
      sourceLabel: sourceLabel || "EPICS Channel Graph",
      seedRecordName,
      originNodeIds,
      message: `No record declarations were found in ${sourceLabel || "the source file"}.`,
      allowAddDatabaseFiles: true,
      nodes: [],
      edges: [],
      adjacency: {},
    };
  }

  const nodesById = new Map();
  const edges = [];
  const seenEdges = new Set();

  for (const declaration of declarations) {
    const fieldDeclarations = extractFieldDeclarationsInRecord(sourceText, declaration);
    const scanValue = fieldDeclarations.find((field) => field.fieldName === "SCAN")?.value || "";
    nodesById.set(declaration.name, {
      id: declaration.name,
      name: declaration.name,
      recordType: declaration.recordType,
      scanValue,
      external: false,
    });
  }

  for (const declaration of declarations) {
    const fieldDeclarations = extractFieldDeclarationsInRecord(sourceText, declaration);
    for (const fieldDeclaration of fieldDeclarations) {
      const dbfType = getRuntimeProbeFieldTypeForRecordType(
        declaration.recordType,
        fieldDeclaration.fieldName,
      );
      if (!LINK_DBF_TYPES.has(dbfType) && !isLinkField(fieldDeclaration.fieldName)) {
        continue;
      }

      const target = parseChannelGraphLinkTarget(fieldDeclaration.value);
      if (!target) {
        continue;
      }

      const targetNodeId = target.recordName
        ? target.recordName
        : `__value__:${declaration.name}:${fieldDeclaration.fieldName}:${target.rawValue}`;
      const targetNodeName = target.recordName || target.rawValue;
      if (!targetNodeName) {
        continue;
      }

      if (!nodesById.has(targetNodeId)) {
        nodesById.set(targetNodeId, {
          id: targetNodeId,
          name: targetNodeName,
          recordType: "",
          scanValue: "",
          external: true,
        });
      }

      const label = target.targetField
        ? `${fieldDeclaration.fieldName}:${target.targetField}`
        : fieldDeclaration.fieldName;
      const edge = dbfType === "DBF_INLINK" || isChannelGraphInputField(fieldDeclaration.fieldName)
        ? {
          fromId: targetNodeId,
          toId: declaration.name,
          label,
        }
        : {
          fromId: declaration.name,
          toId: targetNodeId,
          label,
        };
      const edgeKey = `${edge.fromId}\u0000${edge.toId}\u0000${edge.label}`;
      if (seenEdges.has(edgeKey)) {
        continue;
      }
      seenEdges.add(edgeKey);
      edges.push(edge);
    }
  }

  const adjacency = Object.create(null);
  for (const edge of edges) {
    adjacency[edge.fromId] = adjacency[edge.fromId] || [];
    adjacency[edge.toId] = adjacency[edge.toId] || [];
    if (!adjacency[edge.fromId].includes(edge.toId)) {
      adjacency[edge.fromId].push(edge.toId);
    }
    if (!adjacency[edge.toId].includes(edge.fromId)) {
      adjacency[edge.toId].push(edge.fromId);
    }
  }

  const nodeIds = [...nodesById.keys()];
  let message = "";
  if (mode === "static" && (!originNodeIds || originNodeIds.length === 0)) {
    message = "Enter a channel name to start the Channel Graph.";
  } else {
    message =
      seedRecordName && !nodesById.has(seedRecordName)
        ? `Seed record "${seedRecordName}" was not found in ${sourceLabel || "the source file"}. Showing the full graph.`
        : edges.length === 0
          ? `No link relationships were found in ${sourceLabel || "the source file"}.`
          : "";
  }
  return {
    mode,
    sourceLabel: sourceLabel || "EPICS Channel Graph",
    seedRecordName,
    originNodeIds,
    message,
    allowAddDatabaseFiles: true,
    nodes: nodeIds.map((id) => nodesById.get(id)),
    edges,
    adjacency,
  };
}

function parseChannelGraphLinkTarget(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed.startsWith("@")) {
    return {
      rawValue: stripOptionalWrappingQuotes(trimmed),
      targetField: "",
    };
  }

  const token = trimmed
    .split(/[\s,]+/)
    .map((part) => String(part || "").trim())
    .find(Boolean);
  if (!token) {
    return undefined;
  }

  const unwrapped = stripOptionalWrappingQuotes(token);
  if (unwrapped === "0" || unwrapped === "1") {
    return {
      rawValue: unwrapped,
      targetField: "",
    };
  }
  const recordName = unwrapped.split(".")[0] || "";
  if (!recordName) {
    return {
      rawValue: stripOptionalWrappingQuotes(trimmed),
      targetField: "",
    };
  }

  return {
    recordName,
    targetField: unwrapped.includes(".") ? unwrapped.split(".").slice(1).join(".") : "",
  };
}

function stripOptionalWrappingQuotes(value) {
  const text = String(value || "");
  if (text.length >= 2) {
    const first = text[0];
    const last = text[text.length - 1];
    if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
      return text.slice(1, -1);
    }
  }
  return text;
}

function isChannelGraphInputField(fieldName) {
  return (
    fieldName === "INP" ||
    /^INP[A-U]$/.test(fieldName) ||
    /^DOL[0-9A-F]$/.test(fieldName)
  );
}

function buildChannelGraphWebviewHtml(webview, graphState) {
  const nonce = String(Date.now());
  const initialState = JSON.stringify(graphState || {}).replace(/</g, "\\u003c");
  const escapedSourceLabel = escapeChannelGraphHtml(
    String(graphState?.sourceLabel || "EPICS Channel Graph"),
  );
  const escapedMessage = graphState?.message
    ? escapeChannelGraphHtml(String(graphState.message))
    : "";
  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta
      http-equiv="Content-Security-Policy"
      content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
    />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>EPICS Channel Graph</title>
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
        position: sticky;
        top: 0;
        z-index: 2;
        padding: 12px 16px;
        border-bottom: 1px solid var(--vscode-panel-border);
        background: var(--vscode-editor-background);
      }
      .toolbar-title {
        font-size: 1.1rem;
        font-weight: 600;
      }
      .toolbar-meta {
        margin-top: 6px;
        color: var(--vscode-descriptionForeground);
      }
      .toolbar-actions {
        margin-top: 8px;
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        align-items: center;
      }
      .toolbar-origin {
        margin-top: 8px;
        display: flex;
        gap: 8px;
        align-items: center;
        flex-wrap: wrap;
      }
      .source-files {
        margin-top: 10px;
        display: flex;
        flex-direction: column;
        gap: 6px;
      }
      .source-file-row {
        display: flex;
        align-items: center;
        gap: 8px;
        max-width: 100%;
      }
      .source-file-path {
        font-family: var(--vscode-editor-font-family, var(--vscode-font-family));
        font-size: 0.9rem;
        color: var(--vscode-descriptionForeground);
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .source-file-remove {
        border: 1px solid var(--vscode-panel-border);
        background: transparent;
        color: var(--vscode-foreground);
        border-radius: 999px;
        width: 22px;
        height: 22px;
        line-height: 18px;
        cursor: pointer;
        padding: 0;
        flex: 0 0 auto;
      }
      .toolbar-button {
        border: 1px solid var(--vscode-button-border, transparent);
        background: var(--vscode-button-background);
        color: var(--vscode-button-foreground);
        border-radius: 6px;
        padding: 6px 12px;
        cursor: pointer;
      }
      .toolbar-button:hover {
        background: var(--vscode-button-hoverBackground);
      }
      .toolbar-button[disabled] {
        opacity: 0.55;
        cursor: default;
      }
      .origin-input {
        min-width: 240px;
        max-width: 380px;
        border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
        background: var(--vscode-input-background);
        color: var(--vscode-input-foreground);
        border-radius: 6px;
        padding: 6px 10px;
        outline: none;
      }
      .message {
        margin-top: 6px;
        color: var(--vscode-descriptionForeground);
      }
      .viewport {
        position: relative;
        flex: 1;
        overflow: auto;
        background:
          radial-gradient(circle at 1px 1px, color-mix(in srgb, var(--vscode-panel-border) 55%, transparent) 1px, transparent 0);
        background-size: 28px 28px;
      }
      .stage {
        position: relative;
        min-width: 1400px;
        min-height: 900px;
      }
      .edges {
        position: absolute;
        inset: 0;
        overflow: visible;
      }
      .node-layer {
        position: absolute;
        inset: 0;
      }
      .graph-node {
        position: absolute;
        min-width: 140px;
        max-width: 220px;
        padding: 10px 14px;
        border-radius: 18px;
        border: 2px solid #f000ff;
        background: color-mix(in srgb, var(--vscode-editor-background) 35%, #8bc7ff 65%);
        color: #000;
        text-align: center;
        user-select: none;
        cursor: grab;
        box-shadow: 0 10px 24px color-mix(in srgb, #000 12%, transparent);
      }
      .graph-node.external {
        border-color: #7ab7ff;
        border-radius: 999px;
        background: color-mix(in srgb, var(--vscode-editor-background) 30%, #9fcfff 70%);
      }
      .graph-node.collapsed {
        border-radius: 999px;
      }
      .graph-node:active {
        cursor: grabbing;
      }
      .node-title {
        font-size: 1.05rem;
        font-weight: 600;
        line-height: 1.2;
      }
      .node-subtitle {
        margin-top: 2px;
        font-size: 0.95rem;
        line-height: 1.2;
      }
      .node-value {
        margin-top: 4px;
        font-size: 1rem;
        line-height: 1.15;
      }
      .edge-label {
        font-size: 0.85rem;
        fill: var(--vscode-foreground);
      }
      .edge-label-bg {
        fill: color-mix(in srgb, var(--vscode-editor-background) 82%, #ffffff 18%);
        stroke: none;
      }
    </style>
  </head>
  <body>
    <div class="toolbar">
      <div class="toolbar-title">EPICS Channel Graph</div>
      <div id="channelGraphSourceLabel" class="toolbar-meta">Source: ${escapedSourceLabel}</div>
      <div id="channelGraphHelpText" class="toolbar-meta">Drag nodes to reposition. Double-click a node to expand one more hop.</div>
      <div class="toolbar-actions">
        <button id="addDatabaseFilesButton" class="toolbar-button" type="button" ${graphState?.allowAddDatabaseFiles === false ? "disabled" : ""}>Add Database Files</button>
        <button id="clearChannelGraphButton" class="toolbar-button" type="button">Clear Graph</button>
      </div>
      <div class="toolbar-origin">
        <input id="channelGraphOriginInput" class="origin-input" type="text" placeholder="Enter channel name" />
        <button id="addChannelGraphOriginButton" class="toolbar-button" type="button">Add Channel</button>
      </div>
      <div id="channelGraphSourceFiles" class="source-files"></div>
      <div id="channelGraphMessage" class="message" ${graphState?.message ? "" : 'style="display:none"'}>${escapedMessage}</div>
    </div>
    <div id="viewport" class="viewport">
      <div id="stage" class="stage">
        <svg id="edges" class="edges">
          <defs>
            <marker id="arrow" markerWidth="10" markerHeight="10" refX="7" refY="3" orient="auto" markerUnits="strokeWidth">
              <path d="M0,0 L0,6 L8,3 z" fill="#3b6cff"></path>
            </marker>
          </defs>
        </svg>
        <div id="nodeLayer" class="node-layer"></div>
      </div>
    </div>
    <script nonce="${nonce}">
      const vscode = acquireVsCodeApi();
      let state = ${initialState};
      const viewport = document.getElementById("viewport");
      const stage = document.getElementById("stage");
      const edgeLayer = document.getElementById("edges");
      const nodeLayer = document.getElementById("nodeLayer");
      const sourceLabelNode = document.getElementById("channelGraphSourceLabel");
      const helpTextNode = document.getElementById("channelGraphHelpText");
      const messageNode = document.getElementById("channelGraphMessage");
      const sourceFilesNode = document.getElementById("channelGraphSourceFiles");
      const addDatabaseFilesButton = document.getElementById("addDatabaseFilesButton");
      const clearChannelGraphButton = document.getElementById("clearChannelGraphButton");
      const originInput = document.getElementById("channelGraphOriginInput");
      const addOriginButton = document.getElementById("addChannelGraphOriginButton");
      let nodeById = new Map((state.nodes || []).map((node) => [node.id, node]));
      function getOriginNodeIds() {
        return Array.isArray(state.originNodeIds)
          ? state.originNodeIds.filter((nodeId) => nodeById.has(nodeId))
          : [];
      }
      function getInitialVisibleNodeIds() {
        const originIds = getOriginNodeIds();
        if (!originIds.length) {
          return [];
        }
        const visibleIds = new Set();
        for (const originNodeId of originIds) {
          visibleIds.add(originNodeId);
          for (const neighborId of state.adjacency?.[originNodeId] || []) {
            visibleIds.add(neighborId);
          }
        }
        return [...visibleIds];
      }

      let visibleNodeIds = new Set(getInitialVisibleNodeIds());
      let expandedNodeIds = new Set(getOriginNodeIds());
      const positions = Object.create(null);
      let dragState = undefined;

      function escapeHtml(value) {
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#39;");
      }

      function isExpandedNode(nodeId) {
        return expandedNodeIds.has(nodeId);
      }

      function renderSourceFiles() {
        if (!sourceFilesNode) {
          return;
        }
        const sourceFiles = Array.isArray(state.sourceFiles) ? state.sourceFiles : [];
        sourceFilesNode.innerHTML = sourceFiles.map((entry) =>
          '<div class="source-file-row">' +
            '<button class="source-file-remove" data-source-key="' + escapeHtml(entry.key) + '" type="button" title="Remove database file">×</button>' +
            '<div class="source-file-path" title="' + escapeHtml(entry.path) + '">' + escapeHtml(entry.path) + '</div>' +
          '</div>'
        ).join("");
        sourceFilesNode.style.display = sourceFiles.length ? "flex" : "none";
      }

      function buildNodeLines(node, expanded) {
        if (node.external || !expanded) {
          return [escapeHtml(node.name)];
        }
        const subtitle = node.scanValue
          ? "(" + escapeHtml(node.recordType || "?") + ") (" + escapeHtml(node.scanValue) + ")"
          : "(" + escapeHtml(node.recordType || "?") + ")";
        const lines = [
          '<div class="node-title">' + escapeHtml(node.name) + '</div>',
          '<div class="node-subtitle">' + subtitle + '</div>',
        ];
        if (String(node.valueText || "").trim()) {
          lines.push('<div class="node-value">' + escapeHtml(node.valueText) + '</div>');
        }
        return lines;
      }

      function ensureInitialPositions() {
        if (state.seedRecordName && nodeById.has(state.seedRecordName)) {
          const levels = new Map([[state.seedRecordName, 0]]);
          const queue = [state.seedRecordName];
          while (queue.length) {
            const current = queue.shift();
            const nextLevel = (levels.get(current) || 0) + 1;
            for (const neighborId of [...(state.adjacency?.[current] || [])].sort()) {
              if (!visibleNodeIds.has(neighborId) || levels.has(neighborId)) {
                continue;
              }
              levels.set(neighborId, nextLevel);
              queue.push(neighborId);
            }
          }

          const nodeIdsByLevel = new Map();
          for (const [nodeId, level] of levels.entries()) {
            const levelNodeIds = nodeIdsByLevel.get(level) || [];
            levelNodeIds.push(nodeId);
            nodeIdsByLevel.set(level, levelNodeIds);
          }

          [...nodeIdsByLevel.keys()].sort((left, right) => left - right).forEach((level) => {
            const levelNodeIds = (nodeIdsByLevel.get(level) || []).sort();
            levelNodeIds.forEach((nodeId, index) => {
              positions[nodeId] = positions[nodeId] || {
                x: 140 + level * 280,
                y: 120 + index * 180,
              };
            });
          });
          return;
        }
        (state.nodes || []).forEach((node, index) => {
          if (positions[node.id]) {
            return;
          }
          const column = index % 4;
          const row = Math.floor(index / 4);
          positions[node.id] = { x: 140 + column * 260, y: 120 + row * 180 };
        });
      }

      function getNodeSize(node) {
        const expanded = isExpandedNode(node.id);
        const titleLength = String(node.name || "").length;
        const subtitleLength = node.external || !expanded
          ? 0
          : String(node.recordType || "").length + String(node.scanValue || "").length + 6;
        const valueLength = node.external || !expanded
          ? 0
          : String(node.valueText || "").length;
        return {
          width: Math.max((Math.max(titleLength, subtitleLength, valueLength) * 9) + 40, (node.external || !expanded) ? 120 : 170),
          height: (node.external || !expanded) ? 54 : (String(node.valueText || "").trim() ? 106 : 84),
        };
      }

      function getVisibleEdges() {
        return (state.edges || []).filter(
          (edge) => visibleNodeIds.has(edge.fromId) && visibleNodeIds.has(edge.toId),
        );
      }

      function getVisibleBounds() {
        const visibleNodes = (state.nodes || []).filter((node) => visibleNodeIds.has(node.id));
        if (!visibleNodes.length) {
          return { width: 1400, height: 900 };
        }
        let maxX = 0;
        let maxY = 0;
        for (const node of visibleNodes) {
          const position = positions[node.id] || { x: 120, y: 120 };
          const size = getNodeSize(node);
          maxX = Math.max(maxX, position.x + size.width + 120);
          maxY = Math.max(maxY, position.y + size.height + 120);
        }
        return {
          width: Math.max(maxX, 1400),
          height: Math.max(maxY, 900),
        };
      }

      function connectionPoint(position, size, angle, invert) {
        const radius = Math.min(size.width, size.height) / 2;
        const factor = invert ? -1 : 1;
        return {
          x: position.x + size.width / 2 + Math.cos(angle) * radius * factor,
          y: position.y + size.height / 2 + Math.sin(angle) * radius * factor,
        };
      }

      function expandNode(nodeId) {
        const neighbors = (state.adjacency?.[nodeId] || []).filter((id) => !visibleNodeIds.has(id));
        expandedNodeIds.add(nodeId);
        if (!neighbors.length) {
          render();
          return;
        }
        const base = positions[nodeId] || { x: 620, y: 260 };
        const radius = 220;
        neighbors.forEach((neighborId, index) => {
          const angle = (Math.PI * 2 * index) / Math.max(neighbors.length, 1);
          if (!positions[neighborId]) {
            positions[neighborId] = {
              x: Math.round(base.x + Math.cos(angle) * radius),
              y: Math.round(base.y + Math.sin(angle) * radius),
            };
          }
          visibleNodeIds.add(neighborId);
        });
        render();
      }

      function revealExpandedNodeNeighbors(targetVisibleNodeIds) {
        for (const expandedNodeId of expandedNodeIds) {
          if (!nodeById.has(expandedNodeId)) {
            continue;
          }
          targetVisibleNodeIds.add(expandedNodeId);
          for (const neighborId of state.adjacency?.[expandedNodeId] || []) {
            if (nodeById.has(neighborId)) {
              targetVisibleNodeIds.add(neighborId);
            }
          }
        }
      }

      function applyGraphState(nextState) {
        state = nextState || {};
        nodeById = new Map((state.nodes || []).map((node) => [node.id, node]));
        const nextOriginNodeIds = getOriginNodeIds();
        const showEmptyStaticView = state.mode !== "dynamic" && nextOriginNodeIds.length === 0;

        const nextVisibleNodeIds = new Set();
        if (!showEmptyStaticView) {
          for (const nodeId of visibleNodeIds) {
            if (nodeById.has(nodeId)) {
              nextVisibleNodeIds.add(nodeId);
            }
          }
          for (const nodeId of getInitialVisibleNodeIds()) {
            nextVisibleNodeIds.add(nodeId);
          }
          revealExpandedNodeNeighbors(nextVisibleNodeIds);
        }
        visibleNodeIds = nextVisibleNodeIds;

        const nextExpandedNodeIds = new Set();
        if (!showEmptyStaticView && nextOriginNodeIds.length > 0) {
          for (const nodeId of expandedNodeIds) {
            if (nodeById.has(nodeId)) {
              nextExpandedNodeIds.add(nodeId);
            }
          }
          for (const originNodeId of nextOriginNodeIds) {
            nextExpandedNodeIds.add(originNodeId);
          }
        }
        expandedNodeIds = nextExpandedNodeIds;

        for (const nodeId of Object.keys(positions)) {
          if (!nodeById.has(nodeId)) {
            delete positions[nodeId];
          }
        }

        if (sourceLabelNode) {
          sourceLabelNode.textContent = "Source: " + String(state.sourceLabel || "EPICS Channel Graph");
        }
        if (helpTextNode) {
          helpTextNode.textContent =
            "Drag nodes to reposition. Double-click a node to expand one more hop.";
        }
        if (messageNode) {
          const nextMessage = String(state.message || "");
          messageNode.textContent = nextMessage;
          messageNode.style.display = nextMessage ? "block" : "none";
        }
        if (addDatabaseFilesButton) {
          addDatabaseFilesButton.disabled = state.allowAddDatabaseFiles === false;
        }
        if (originInput && document.activeElement !== originInput) {
          originInput.value = "";
        }
        renderSourceFiles();
        render();
      }

      function render() {
        ensureInitialPositions();
        const bounds = getVisibleBounds();
        stage.style.width = bounds.width + "px";
        stage.style.height = bounds.height + "px";
        edgeLayer.setAttribute("viewBox", "0 0 " + bounds.width + " " + bounds.height);
        edgeLayer.setAttribute("width", String(bounds.width));
        edgeLayer.setAttribute("height", String(bounds.height));

        const visibleNodes = (state.nodes || []).filter((node) => visibleNodeIds.has(node.id));
        nodeLayer.innerHTML = visibleNodes.map((node) => {
          const position = positions[node.id] || { x: 120, y: 120 };
          const size = getNodeSize(node);
          const expanded = isExpandedNode(node.id);
          const classes = ["graph-node"];
          if (node.external) {
            classes.push("external");
          } else if (!expanded) {
            classes.push("collapsed");
          }
          return '<div class="' + classes.join(" ") + '" data-node-id="' + escapeHtml(node.id) + '" ' +
            'style="left:' + position.x + 'px;top:' + position.y + 'px;width:' + size.width + 'px;min-height:' + size.height + 'px">' +
            buildNodeLines(node, expanded).join("") +
            '</div>';
        }).join("");

        edgeLayer.innerHTML = '<defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="7" refY="3" orient="auto" markerUnits="strokeWidth"><path d="M0,0 L0,6 L8,3 z" fill="#3b6cff"></path></marker></defs>' +
          getVisibleEdges().map((edge) => {
            const fromNode = nodeById.get(edge.fromId);
            const toNode = nodeById.get(edge.toId);
            if (!fromNode || !toNode) {
              return "";
            }
            const fromPos = positions[edge.fromId] || { x: 120, y: 120 };
            const toPos = positions[edge.toId] || { x: 120, y: 120 };
            const fromSize = getNodeSize(fromNode);
            const toSize = getNodeSize(toNode);
            if (edge.fromId === edge.toId) {
              const labelWidth = Math.max(String(edge.label || "").length * 8, 24);
              const startX = fromPos.x + fromSize.width * 0.62;
              const startY = fromPos.y + 10;
              const endX = fromPos.x + fromSize.width * 0.38;
              const endY = fromPos.y + 10;
              const controlY = fromPos.y - Math.max(fromSize.height, 56) * 0.7;
              const labelX = fromPos.x + fromSize.width / 2;
              const labelY = controlY - 8;
              return '<path d="M ' + startX + ' ' + startY +
                ' C ' + (fromPos.x + fromSize.width + 36) + ' ' + controlY +
                ', ' + (fromPos.x - 36) + ' ' + controlY +
                ', ' + endX + ' ' + endY +
                '" stroke="#3b6cff" stroke-width="2" fill="none"></path>' +
                '<line x1="' + (endX + 10) + '" y1="' + (endY - 12) + '" x2="' + endX + '" y2="' + endY + '" stroke="#3b6cff" stroke-width="2" marker-end="url(#arrow)"></line>' +
                '<rect class="edge-label-bg" x="' + (labelX - labelWidth / 2 - 6) + '" y="' + (labelY - 14) + '" width="' + (labelWidth + 12) + '" height="20" rx="8" ry="8"></rect>' +
                '<text class="edge-label" x="' + labelX + '" y="' + labelY + '" text-anchor="middle">' + escapeHtml(edge.label) + '</text>';
            }
            const angle = Math.atan2(
              (toPos.y + toSize.height / 2) - (fromPos.y + fromSize.height / 2),
              (toPos.x + toSize.width / 2) - (fromPos.x + fromSize.width / 2),
            );
            const start = connectionPoint(fromPos, fromSize, angle, false);
            const end = connectionPoint(toPos, toSize, angle, true);
            const labelX = (start.x + end.x) / 2;
            const labelY = (start.y + end.y) / 2 - 4;
            const labelWidth = Math.max(String(edge.label || "").length * 8, 24);
            return '<line x1="' + start.x + '" y1="' + start.y + '" x2="' + end.x + '" y2="' + end.y + '" stroke="#3b6cff" stroke-width="2" marker-end="url(#arrow)"></line>' +
              '<rect class="edge-label-bg" x="' + (labelX - labelWidth / 2 - 6) + '" y="' + (labelY - 14) + '" width="' + (labelWidth + 12) + '" height="20" rx="8" ry="8"></rect>' +
              '<text class="edge-label" x="' + labelX + '" y="' + labelY + '" text-anchor="middle">' + escapeHtml(edge.label) + '</text>';
          }).join("");
      }

      nodeLayer.addEventListener("mousedown", (event) => {
        const target = event.target instanceof Element ? event.target.closest("[data-node-id]") : undefined;
        if (!(target instanceof HTMLElement)) {
          return;
        }
        const nodeId = target.dataset.nodeId;
        if (!nodeId) {
          return;
        }
        const position = positions[nodeId] || { x: 120, y: 120 };
        const rect = stage.getBoundingClientRect();
        const localX = event.clientX - rect.left + viewport.scrollLeft;
        const localY = event.clientY - rect.top + viewport.scrollTop;
        dragState = {
          nodeId,
          offsetX: localX - position.x,
          offsetY: localY - position.y,
        };
      });

      nodeLayer.addEventListener("dblclick", (event) => {
        const target = event.target instanceof Element ? event.target.closest("[data-node-id]") : undefined;
        if (!(target instanceof HTMLElement)) {
          return;
        }
        const nodeId = target.dataset.nodeId;
        if (nodeId) {
          if (state.mode === "dynamic") {
            expandedNodeIds.add(nodeId);
            render();
            vscode.postMessage({
              type: "expandChannelGraphNode",
              nodeId,
            });
            return;
          }
          expandNode(nodeId);
        }
      });

      addDatabaseFilesButton?.addEventListener("click", () => {
        if (state.allowAddDatabaseFiles === false) {
          return;
        }
        vscode.postMessage({ type: "pickChannelGraphDatabaseFiles" });
      });
      sourceFilesNode?.addEventListener("click", (event) => {
        const target = event.target instanceof Element ? event.target.closest("[data-source-key]") : undefined;
        if (!(target instanceof HTMLElement)) {
          return;
        }
        const key = String(target.dataset.sourceKey || "").trim();
        if (!key) {
          return;
        }
        vscode.postMessage({
          type: "removeChannelGraphDatabaseFile",
          key,
        });
      });
      clearChannelGraphButton?.addEventListener("click", () => {
        vscode.postMessage({ type: "clearChannelGraph" });
      });
      function submitOriginInput() {
        const nodeId = String(originInput?.value || "").trim();
        if (!nodeId) {
          return;
        }
        vscode.postMessage({
          type: "addChannelGraphOrigin",
          nodeId,
        });
        if (originInput) {
          originInput.value = "";
        }
      }
      addOriginButton?.addEventListener("click", submitOriginInput);
      originInput?.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          submitOriginInput();
        }
      });
      window.addEventListener("message", (event) => {
        if (event.data?.type === "setChannelGraphState") {
          applyGraphState(event.data.state);
        }
      });

      window.addEventListener("mousemove", (event) => {
        if (!dragState) {
          return;
        }
        const rect = stage.getBoundingClientRect();
        const localX = event.clientX - rect.left + viewport.scrollLeft;
        const localY = event.clientY - rect.top + viewport.scrollTop;
        positions[dragState.nodeId] = {
          x: Math.max(20, Math.round(localX - dragState.offsetX)),
          y: Math.max(20, Math.round(localY - dragState.offsetY)),
        };
        render();
      });

      window.addEventListener("mouseup", () => {
        dragState = undefined;
      });

      applyGraphState(state);
    </script>
  </body>
</html>`;
}

function escapeChannelGraphHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildExcelImportPreviewWebviewHtml(webview) {
  const nonce = String(Date.now());
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>EPICS Excel Import Preview</title>
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
    .panel {
      border: 1px dashed var(--vscode-panel-border);
      border-radius: 12px;
      padding: 28px;
      min-height: 220px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      gap: 16px;
      background: color-mix(in srgb, var(--vscode-editor-background) 86%, var(--vscode-editorHoverWidget-background) 14%);
    }
    .panel.dragover {
      border-color: var(--vscode-focusBorder);
      background: color-mix(in srgb, var(--vscode-editor-background) 70%, var(--vscode-focusBorder) 30%);
    }
    h1 {
      margin: 0;
      font-size: 1.1rem;
    }
    p {
      margin: 0;
      line-height: 1.5;
      opacity: 0.9;
    }
    button {
      width: fit-content;
      padding: 8px 14px;
      border: none;
      border-radius: 999px;
      cursor: pointer;
      color: var(--vscode-button-foreground);
      background: var(--vscode-button-background);
      font: inherit;
    }
    button:hover {
      background: var(--vscode-button-hoverBackground);
    }
    .status {
      min-height: 1.5em;
      color: var(--vscode-descriptionForeground);
    }
  </style>
</head>
<body>
  <div id="drop-panel" class="panel">
    <h1>Drop an EPICS Excel workbook here</h1>
    <p>If the workbook contains EPICS-style sheets whose first row starts with <code>Record</code> and <code>Type</code>, each matching sheet will open as a temporary EPICS database tab.</p>
    <button id="pick-file" type="button">Choose Workbook...</button>
    <input id="file-input" type="file" accept=".xlsx" hidden />
    <div id="status" class="status">Ready.</div>
  </div>
  <script nonce="${nonce}">
    const vscode = acquireVsCodeApi();
    const dropPanel = document.getElementById("drop-panel");
    const statusNode = document.getElementById("status");
    const fileInput = document.getElementById("file-input");
    const pickButton = document.getElementById("pick-file");

    function setStatus(message) {
      statusNode.textContent = message;
    }

    async function postWorkbook(file) {
      if (!file) {
        return;
      }
      setStatus("Importing " + file.name + "...");
      const arrayBuffer = await file.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      let binary = "";
      const chunkSize = 0x8000;
      for (let offset = 0; offset < bytes.length; offset += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
      }
      vscode.postMessage({
        type: "importExcelPreviewWorkbook",
        name: file.name,
        lastModified: file.lastModified,
        base64: btoa(binary),
      });
    }

    dropPanel.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropPanel.classList.add("dragover");
    });
    dropPanel.addEventListener("dragleave", () => {
      dropPanel.classList.remove("dragover");
    });
    dropPanel.addEventListener("drop", async (event) => {
      event.preventDefault();
      dropPanel.classList.remove("dragover");
      const file = event.dataTransfer?.files?.[0];
      if (file) {
        await postWorkbook(file);
      }
    });
    pickButton.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files?.[0];
      if (file) {
        await postWorkbook(file);
      }
      fileInput.value = "";
    });

    window.addEventListener("message", (event) => {
      const message = event.data;
      if (message?.type === "excelImportPreviewResult") {
        setStatus(message.message || (message.success ? "Import complete." : "Import failed."));
      }
    });
  </script>
</body>
</html>`;
}

function buildMonitorFileText(recordNames, macroNames, eol = "\n") {
  const lines = [
    "# this is a pvlist file for EPICS Workbench in VSCode",
    "# Fill in the macro values below, then open this file and click the EPICS play button in the status bar.",
    "# Each non-comment line after the macro block monitors one EPICS record or PV.",
    "",
  ];

  if (macroNames.length > 0) {
    for (const macroName of macroNames) {
      lines.push(`${macroName} = `);
    }
    lines.push("");
  }

  lines.push(...recordNames);
  return lines.join(eol);
}

function extractUniqueRecordNames(text) {
  const names = [];
  const seen = new Set();

  for (const declaration of extractRecordDeclarations(text)) {
    if (!declaration.name || seen.has(declaration.name)) {
      continue;
    }

    seen.add(declaration.name);
    names.push(declaration.name);
  }

  return names;
}

function extractRecordNameMacroNames(recordNames) {
  return extractMacroNames(recordNames.join("\n")).sort(compareLabels);
}

function buildRecordNamesClipboardText(recordNames, eol = "\n") {
  const macroNames = extractMacroNames(recordNames.join("\n"));
  const lines = [];

  if (macroNames.length > 0) {
    for (const macroName of macroNames) {
      lines.push(`${macroName} = `);
    }
    lines.push("");
  }

  lines.push(...recordNames);
  return lines.join(eol);
}

async function copySubstitutionsAsExpandedDatabaseInActiveEditor(workspaceIndex) {
  const editor = vscode.window.activeTextEditor;
  if (!editor || !isSubstitutionsDocument(editor.document)) {
    return;
  }

  const document = editor.document;
  const snapshot = await workspaceIndex.getSnapshot();
  const diagnostics = createSubstitutionDiagnostics(document, snapshot);
  if (diagnostics.length > 0) {
    vscode.window.showErrorMessage(
      "Cannot copy as expanded db until the substitutions file errors are fixed.",
    );
    return;
  }

  const expansion = buildExpandedDatabaseFromSubstitutions(snapshot, document);
  if (expansion.errors.length > 0) {
    vscode.window.showErrorMessage(
      `Cannot copy as expanded db: ${expansion.errors[0]}`,
    );
    return;
  }

  if (!expansion.text) {
    vscode.window.showWarningMessage(
      "No substitutions file loads were found in the active substitutions file.",
    );
    return;
  }

  await vscode.env.clipboard.writeText(expansion.text);
  const loadLabel = `${expansion.sectionCount} expansion${expansion.sectionCount === 1 ? "" : "s"}`;
  vscode.window.showInformationMessage(`Copied ${loadLabel} as expanded db text.`);
}

async function resolveExpandedSubstitutionsDatabaseSource(
  workspaceIndex,
  document,
  actionLabel,
) {
  if (!document || !isSubstitutionsDocument(document)) {
    return undefined;
  }

  const snapshot = await workspaceIndex.getSnapshot();
  const diagnostics = createSubstitutionDiagnostics(document, snapshot);
  if (diagnostics.length > 0) {
    vscode.window.showErrorMessage(
      `Cannot ${actionLabel} until the substitutions file errors are fixed.`,
    );
    return undefined;
  }

  const expansion = buildExpandedDatabaseFromSubstitutions(snapshot, document);
  if (expansion.errors.length > 0) {
    vscode.window.showErrorMessage(
      `Cannot ${actionLabel}: ${expansion.errors[0]}`,
    );
    return undefined;
  }

  if (!expansion.text) {
    vscode.window.showWarningMessage(
      "No substitutions file loads were found in the active substitutions file.",
    );
    return undefined;
  }

  return {
    text: expansion.text,
    recordNames: extractUniqueRecordNames(expansion.text),
  };
}

function buildExpandedDatabaseFromSubstitutions(snapshot, document) {
  const eol = getDocumentEol(document);
  const sections = [];
  const errors = [];
  let globalMacros = new Map();
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const releaseVariables = project ? project.releaseVariables : new Map();

  for (const block of extractSubstitutionBlocksWithRanges(document.getText())) {
    if (block.kind === "global") {
      globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body));
      continue;
    }

    if (block.kind !== "file" || !block.templatePath) {
      continue;
    }

    const templateAbsolutePath = resolveSubstitutionTemplateAbsolutePathForDocument(
      snapshot,
      document,
      block.templatePath,
    );
    if (!templateAbsolutePath) {
      errors.push(`Cannot resolve substitutions file "${block.templatePath}".`);
      continue;
    }

    const templateText = readTextFile(templateAbsolutePath);
    if (templateText === undefined) {
      errors.push(`Cannot read substitutions file "${block.templatePath}".`);
      continue;
    }
    const templateTextWithoutToc = removeDatabaseTocBlock(templateText);

    const parsedRows = parseSubstitutionFileBlockRows(block.body);
    const rows = parsedRows.length > 0
      ? parsedRows.map((rowMacros) => mergeMacroMaps(globalMacros, rowMacros))
      : [new Map(globalMacros)];

    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const rowMacros = rows[rowIndex];
      const expandedText = normalizeTextEol(
        expandEpicsValue(
          templateTextWithoutToc,
          [rowMacros, releaseVariables, process.env],
        ),
        eol,
      ).trimEnd();
      if (!expandedText) {
        continue;
      }

      sections.push(
        buildExpandedSubstitutionSection(
          block.templatePath,
          rowMacros,
          expandedText,
          eol,
        ),
      );
    }
  }

  return {
    text: sections.join(`${eol}${eol}`),
    sectionCount: sections.length,
    errors,
  };
}

function buildExpandedSubstitutionSection(templatePath, rowMacros, expandedText, eol) {
  const lines = [formatExpandedSubstitutionHeaderLine(templatePath)];
  lines.push(...formatExpandedSubstitutionMacroLines(rowMacros));
  lines.push("");
  lines.push(expandedText);
  return lines.join(eol);
}

function formatExpandedSubstitutionHeaderLine(templatePath) {
  const fileName = path.basename(String(templatePath || "").trim()) || "template";
  return `# ------------------- ${fileName} -------------------`;
}

function formatExpandedSubstitutionMacroLines(macros) {
  if (!(macros instanceof Map) || macros.size === 0) {
    return [];
  }

  return [...macros.entries()]
    .sort(([leftName], [rightName]) => compareLabels(leftName, rightName))
    .map(([name, value]) => `# ${name} = ${value}`);
}

function normalizeTextEol(text, eol) {
  return String(text || "").replace(/\r\n/g, "\n").replace(/\n/g, eol);
}

function createDatabaseRecordDecorationTypes() {
  return [
    vscode.window.createTextEditorDecorationType({
      isWholeLine: true,
      light: {
        backgroundColor: "rgba(0, 96, 160, 0.035)",
      },
      dark: {
        backgroundColor: "rgba(128, 192, 255, 0.05)",
      },
    }),
    vscode.window.createTextEditorDecorationType({
      isWholeLine: true,
      light: {
        backgroundColor: "rgba(140, 90, 0, 0.025)",
      },
      dark: {
        backgroundColor: "rgba(255, 210, 120, 0.035)",
      },
    }),
  ];
}

function updateDatabaseRecordDecorationsForVisibleEditors(decorationTypes) {
  for (const editor of vscode.window.visibleTextEditors) {
    updateDatabaseRecordDecorations(editor, decorationTypes);
  }
}

function updateDatabaseRecordDecorationsForDocument(document, decorationTypes) {
  if (!document) {
    return;
  }

  for (const editor of vscode.window.visibleTextEditors) {
    if (editor.document.uri.toString() !== document.uri.toString()) {
      continue;
    }

    updateDatabaseRecordDecorations(editor, decorationTypes);
  }
}

function updateDatabaseRecordDecorations(editor, decorationTypes) {
  if (!editor || !decorationTypes?.length) {
    return;
  }

  if (!isDatabaseDocument(editor.document)) {
    for (const decorationType of decorationTypes) {
      editor.setDecorations(decorationType, []);
    }
    return;
  }

  const text = editor.document.getText();
  const recordDeclarations = extractRecordDeclarations(text);
  const decorationRanges = decorationTypes.map(() => []);

  recordDeclarations.forEach((declaration, index) => {
    const decorationTypeIndex = index % decorationTypes.length;
    const endOffset = Math.max(
      declaration.recordStart,
      declaration.recordEnd - 1,
    );
    decorationRanges[decorationTypeIndex].push(
      new vscode.Range(
        editor.document.positionAt(declaration.recordStart),
        editor.document.positionAt(endOffset),
      ),
    );
  });

  decorationTypes.forEach((decorationType, index) => {
    editor.setDecorations(decorationType, decorationRanges[index]);
  });
}

function createDatabaseValueDecorationTypes() {
  return {
    numeric: vscode.window.createTextEditorDecorationType({
      light: {
        color: "#005cc5",
      },
      dark: {
        color: "#7fb7ff",
      },
    }),
    link: vscode.window.createTextEditorDecorationType({
      light: {
        color: "#8a4600",
      },
      dark: {
        color: "#ffba7a",
      },
    }),
    menu: vscode.window.createTextEditorDecorationType({
      light: {
        color: "#8f2d7a",
      },
      dark: {
        color: "#f2a7e8",
      },
    }),
    other: vscode.window.createTextEditorDecorationType({
      light: {
        color: "#2f6f3e",
      },
      dark: {
        color: "#8ed1a0",
      },
    }),
    linkMacro: vscode.window.createTextEditorDecorationType({
      light: {
        color: "#c73e1d",
      },
      dark: {
        color: "#ff9f7f",
      },
      fontStyle: "italic",
    }),
  };
}

async function updateDatabaseValueDecorationsForVisibleEditors(
  workspaceIndex,
  decorationTypes,
) {
  for (const editor of vscode.window.visibleTextEditors) {
    await updateDatabaseValueDecorations(editor, workspaceIndex, decorationTypes);
  }
}

async function updateDatabaseValueDecorationsForDocument(
  document,
  workspaceIndex,
  decorationTypes,
) {
  if (!document) {
    return;
  }

  for (const editor of vscode.window.visibleTextEditors) {
    if (editor.document.uri.toString() !== document.uri.toString()) {
      continue;
    }

    await updateDatabaseValueDecorations(editor, workspaceIndex, decorationTypes);
  }
}

async function updateDatabaseValueDecorations(
  editor,
  workspaceIndex,
  decorationTypes,
) {
  if (!editor || !decorationTypes) {
    return;
  }

  if (!isDatabaseDocument(editor.document)) {
    clearDatabaseValueDecorations(editor, decorationTypes);
    return;
  }

  const snapshot = mergeSnapshotWithDocument(
    await workspaceIndex.getSnapshot(),
    editor.document,
  );
  const rangesByCategory = {
    numeric: [],
    link: [],
    menu: [],
    other: [],
    linkMacro: [],
  };
  const text = editor.document.getText();

  for (const recordDeclaration of extractRecordDeclarations(text)) {
    const fieldTypes = snapshot.fieldTypesByRecordType.get(recordDeclaration.recordType);
    for (const fieldDeclaration of extractFieldDeclarationsInRecord(
      text,
      recordDeclaration,
    )) {
      const dbfType = fieldTypes?.get(fieldDeclaration.fieldName);
      if (!dbfType || fieldDeclaration.valueStart >= fieldDeclaration.valueEnd) {
        continue;
      }

      if (LINK_DBF_TYPES.has(dbfType)) {
        addLinkValueDecorationRanges(
          editor.document,
          fieldDeclaration,
          rangesByCategory,
        );
        continue;
      }

      const category = classifyDatabaseValueDecorationCategory(dbfType);
      rangesByCategory[category].push(
        new vscode.Range(
          editor.document.positionAt(fieldDeclaration.valueStart),
          editor.document.positionAt(fieldDeclaration.valueEnd),
        ),
      );
    }
  }

  for (const [category, decorationType] of Object.entries(decorationTypes)) {
    editor.setDecorations(decorationType, rangesByCategory[category] || []);
  }
}

function clearDatabaseValueDecorations(editor, decorationTypes) {
  for (const decorationType of Object.values(decorationTypes)) {
    editor.setDecorations(decorationType, []);
  }
}

function classifyDatabaseValueDecorationCategory(dbfType) {
  if (NUMERIC_DBF_TYPES.has(dbfType)) {
    return "numeric";
  }

  if (dbfType === "DBF_MENU") {
    return "menu";
  }

  return "other";
}

function addLinkValueDecorationRanges(document, fieldDeclaration, rangesByCategory) {
  const macros = extractMacroOffsets(fieldDeclaration.value, fieldDeclaration.valueStart);
  if (macros.length === 0) {
    rangesByCategory.link.push(
      new vscode.Range(
        document.positionAt(fieldDeclaration.valueStart),
        document.positionAt(fieldDeclaration.valueEnd),
      ),
    );
    return;
  }

  let segmentStart = fieldDeclaration.valueStart;
  for (const macro of macros) {
    if (segmentStart < macro.start) {
      rangesByCategory.link.push(
        new vscode.Range(
          document.positionAt(segmentStart),
          document.positionAt(macro.start),
        ),
      );
    }

    rangesByCategory.linkMacro.push(
      new vscode.Range(
        document.positionAt(macro.start),
        document.positionAt(macro.end),
      ),
    );
    segmentStart = macro.end;
  }

  if (segmentStart < fieldDeclaration.valueEnd) {
    rangesByCategory.link.push(
      new vscode.Range(
        document.positionAt(segmentStart),
        document.positionAt(fieldDeclaration.valueEnd),
      ),
    );
  }
}

function extractMacroOffsets(value, baseOffset) {
  const ranges = [];
  const regex = /\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}/g;
  let match;

  while ((match = regex.exec(String(value || "")))) {
    ranges.push({
      start: baseOffset + match.index,
      end: baseOffset + match.index + match[0].length,
    });
  }

  return ranges;
}

function getDocumentEol(document) {
  return document.eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
}

function extractRecordDeclarations(text) {
  const records = [];
  const sanitizedText = maskDatabaseComments(text);
  const regex = /record\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"/g;
  let match;

  while ((match = regex.exec(sanitizedText))) {
    const typePrefixMatch = match[0].match(/record\(\s*/);
    const prefixMatch = match[0].match(/record\(\s*[A-Za-z0-9_]+\s*,\s*"/);
    const recordTypeStart = match.index + typePrefixMatch[0].length;
    const recordTypeEnd = recordTypeStart + match[1].length;
    const nameStart = match.index + prefixMatch[0].length;
    const nameEnd = nameStart + match[2].length;
    const recordStart = match.index;
    const recordEnd = findRecordBlockEnd(sanitizedText, recordStart);

    records.push({
      recordType: match[1],
      recordTypeStart,
      recordTypeEnd,
      name: match[2],
      nameStart,
      nameEnd,
      recordStart,
      recordEnd,
    });
  }

  return records;
}

function extractFieldValues(text) {
  const fieldValuesByField = new Map();
  const sanitizedText = maskDatabaseComments(text);
  const regex =
    /field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"/g;
  let match;

  while ((match = regex.exec(sanitizedText))) {
    addToMapOfSets(fieldValuesByField, match[1] || match[2], [match[3]]);
  }

  return fieldValuesByField;
}

function extractFieldDeclarations(text, expectedFieldName) {
  const declarations = [];
  const sanitizedText = maskDatabaseComments(text);
  const regex =
    /field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"/g;
  let match;

  while ((match = regex.exec(sanitizedText))) {
    const fieldName = match[1] || match[2];
    if (expectedFieldName && fieldName !== expectedFieldName) {
      continue;
    }

    const fieldPrefixMatch = match[0].match(/field\(\s*/);
    const fieldToken = match[1] ? `"${match[1]}"` : match[2];
    const prefixMatch = match[0].match(
      /field\(\s*(?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+)\s*,\s*"/,
    );
    const fieldNameStart = match.index + fieldPrefixMatch[0].length;
    const fieldNameEnd = fieldNameStart + fieldToken.length;
    const valueStart = match.index + prefixMatch[0].length;
    const valueEnd = valueStart + match[3].length;
    declarations.push({
      fieldName,
      fieldNameStart,
      fieldNameEnd,
      value: match[3],
      valueStart,
      valueEnd,
    });
  }

  return declarations;
}

function maskDatabaseComments(text) {
  let sanitized = "";
  let inString = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];

    if (inString) {
      sanitized += character;
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inString = false;
      }
      continue;
    }

    if (character === "\"") {
      inString = true;
      sanitized += character;
      continue;
    }

    if (character === "#") {
      while (index < text.length && text[index] !== "\n") {
        sanitized += " ";
        index += 1;
      }

      if (index < text.length && text[index] === "\n") {
        sanitized += "\n";
      }
      continue;
    }

    sanitized += character;
  }

  return sanitized;
}

function extractFieldDeclarationsInRecord(text, recordDeclaration, expectedFieldName) {
  const relativeDeclarations = extractFieldDeclarations(
    text.slice(recordDeclaration.recordStart, recordDeclaration.recordEnd),
    expectedFieldName,
  );

  return relativeDeclarations.map((declaration) => ({
    ...declaration,
    fieldNameStart: declaration.fieldNameStart + recordDeclaration.recordStart,
    fieldNameEnd: declaration.fieldNameEnd + recordDeclaration.recordStart,
    valueStart: declaration.valueStart + recordDeclaration.recordStart,
    valueEnd: declaration.valueEnd + recordDeclaration.recordStart,
  }));
}

function extractMacroNames(text) {
  const names = new Set();
  const regex =
    /\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}/g;
  let match;

  while ((match = regex.exec(text))) {
    names.add(match[1] || match[2]);
  }

  return [...names];
}

function extractRequiredMacroNames(text) {
  const names = new Set();
  const regex =
    /\$\(([^)=,\s]+)(?:=([^)]*))?\)|\$\{([^}=,\s]+)(?:=([^}]*))?\}/g;
  let match;

  while ((match = regex.exec(text))) {
    const parenthesizedName = match[1];
    const parenthesizedDefault = match[2];
    const bracedName = match[3];
    const bracedDefault = match[4];

    if (parenthesizedName !== undefined && parenthesizedDefault === undefined) {
      names.add(parenthesizedName);
      continue;
    }

    if (bracedName !== undefined && bracedDefault === undefined) {
      names.add(bracedName);
    }
  }

  return [...names].sort(compareLabels);
}

function extractStartupMacros(text) {
  const names = new Set();
  const regex = /epicsEnvSet\(\s*"?([A-Za-z_][A-Za-z0-9_]*)"?/g;
  let match;

  while ((match = regex.exec(text))) {
    names.add(match[1]);
  }

  return [...names];
}

function extractSubstitutionMacros(text) {
  const names = new Set();
  let match;
  const patternRegex = /pattern\s*\{([^}]*)\}/g;

  while ((match = patternRegex.exec(text))) {
    splitSymbolList(match[1]).forEach((symbol) => names.add(symbol));
  }

  const assignmentsRegex = /\{\s*([^}]*)\}/g;
  while ((match = assignmentsRegex.exec(text))) {
    const body = match[1];
    const assignmentRegex = /([A-Za-z_][A-Za-z0-9_]*)\s*=/g;
    let assignmentMatch;

    while ((assignmentMatch = assignmentRegex.exec(body))) {
      names.add(assignmentMatch[1]);
    }
  }

  return [...names];
}

function extractDbdRecordTypeDeclarations(text) {
  const declarations = [];
  const recordTypeRegex = /recordtype\(\s*([A-Za-z0-9_]+)\s*\)\s*\{/g;
  let match;

  while ((match = recordTypeRegex.exec(text))) {
    const block = readBalancedBlock(text, recordTypeRegex.lastIndex - 1);
    if (!block) {
      continue;
    }

    const typePrefixMatch = match[0].match(/recordtype\(\s*/);
    const nameStart = match.index + typePrefixMatch[0].length;
    declarations.push({
      name: match[1],
      nameStart,
      nameEnd: nameStart + match[1].length,
      blockStart: match.index,
      blockEnd: block.endIndex,
    });

    recordTypeRegex.lastIndex = block.endIndex;
  }

  return declarations;
}

function extractDbdFieldDeclarationsInRecordType(text, recordTypeDeclaration, expectedFieldName) {
  if (!recordTypeDeclaration) {
    return [];
  }

  const blockText = text.slice(recordTypeDeclaration.blockStart, recordTypeDeclaration.blockEnd);
  const declarations = [];
  const fieldRegex = /field\(\s*([A-Z0-9_]+)\s*,/g;
  let match;

  while ((match = fieldRegex.exec(blockText))) {
    if (expectedFieldName && match[1] !== expectedFieldName) {
      continue;
    }

    const fieldPrefixMatch = match[0].match(/field\(\s*/);
    const fieldNameStart = recordTypeDeclaration.blockStart + match.index + fieldPrefixMatch[0].length;
    declarations.push({
      fieldName: match[1],
      fieldNameStart,
      fieldNameEnd: fieldNameStart + match[1].length,
    });
  }

  return declarations;
}

function extractDbdDeviceDeclarations(text) {
  const declarations = [];
  const regex =
    /device\(\s*([A-Za-z0-9_]+)\s*,\s*([A-Za-z0-9_]+)\s*,\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)/g;
  let match;

  while ((match = regex.exec(text))) {
    const prefixMatch = match[0].match(/device\(\s*/);
    const firstCommaOffset = match[0].indexOf(",");
    const secondCommaOffset = match[0].indexOf(",", firstCommaOffset + 1);
    const thirdCommaOffset = match[0].indexOf(",", secondCommaOffset + 1);
    const recordTypeStart = match.index + prefixMatch[0].length;
    const linkTypeStart = match.index + firstCommaOffset + 1 + (match[0].slice(firstCommaOffset + 1).match(/^\s*/) || [""])[0].length;
    const supportNameStart = match.index + secondCommaOffset + 1 + (match[0].slice(secondCommaOffset + 1).match(/^\s*/) || [""])[0].length;
    const choiceNameStart = match.index + thirdCommaOffset + 1 + (match[0].slice(thirdCommaOffset + 1).match(/^\s*"/) || [""])[0].length;

    declarations.push({
      recordType: match[1],
      linkType: match[2],
      supportName: match[3],
      choiceName: match[4],
      start: match.index,
      end: match.index + match[0].length,
      declarationText: match[0],
      recordTypeStart,
      recordTypeEnd: recordTypeStart + match[1].length,
      supportNameStart,
      supportNameEnd: supportNameStart + match[3].length,
      choiceNameStart,
      choiceNameEnd: choiceNameStart + match[4].length,
    });
  }

  return declarations;
}

function extractDbdNamedEntries(text, keyword) {
  const declarations = [];
  const regex = new RegExp(`${keyword}\\(\\s*([A-Za-z_][A-Za-z0-9_]*)`, "g");
  let match;

  while ((match = regex.exec(text))) {
    const prefixMatch = match[0].match(new RegExp(`${keyword}\\(\\s*`));
    const nameStart = match.index + prefixMatch[0].length;
    declarations.push({
      name: match[1],
      nameStart,
      nameEnd: nameStart + match[1].length,
    });
  }

  return declarations;
}

function extractSourceNamedSymbolOccurrences(text) {
  return [
    ...extractSourceExportAddressOccurrences(text),
    ...extractSourceRegistrarOccurrences(text),
    ...extractSourceFunctionOccurrences(text),
  ];
}

function extractSourceExportAddressOccurrences(text) {
  const occurrences = [];
  const regex =
    /epicsExportAddress\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = regex.exec(text))) {
    const exportType = match[1];
    const name = match[2];
    const secondArgumentStart = match.index + match[0].lastIndexOf(name);
    if (isDeviceSupportExportType(exportType)) {
      occurrences.push({
        kind: "deviceSupport",
        name,
        nameStart: secondArgumentStart,
        nameEnd: secondArgumentStart + name.length,
      });
      continue;
    }

    if (isDriverExportType(exportType)) {
      occurrences.push({
        kind: "driver",
        name,
        nameStart: secondArgumentStart,
        nameEnd: secondArgumentStart + name.length,
      });
      continue;
    }

    if (isVariableExportType(exportType)) {
      occurrences.push({
        kind: "variable",
        name,
        nameStart: secondArgumentStart,
        nameEnd: secondArgumentStart + name.length,
      });
    }
  }

  return occurrences;
}

function extractSourceRegistrarOccurrences(text) {
  const occurrences = [];
  const regex = /epicsExportRegistrar\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = regex.exec(text))) {
    const nameStart = match.index + match[0].lastIndexOf(match[1]);
    occurrences.push({
      kind: "registrar",
      name: match[1],
      nameStart,
      nameEnd: nameStart + match[1].length,
    });
  }

  return occurrences;
}

function extractSourceFunctionOccurrences(text) {
  const occurrences = [];
  const regex = /epicsRegisterFunction\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = regex.exec(text))) {
    const nameStart = match.index + match[0].lastIndexOf(match[1]);
    occurrences.push({
      kind: "function",
      name: match[1],
      nameStart,
      nameEnd: nameStart + match[1].length,
    });
  }

  return occurrences;
}

function getDbdNamedSymbolKeyword(kind) {
  return new Map([
    ["driver", "driver"],
    ["registrar", "registrar"],
    ["function", "function"],
    ["variable", "variable"],
  ]).get(kind);
}

function extractDbdFieldsByRecordType(text) {
  const fieldsByRecordType = new Map();
  for (const declaration of extractDbdRecordTypeDeclarations(text)) {
    for (const fieldDeclaration of extractDbdFieldDeclarationsInRecordType(text, declaration)) {
      addToMapOfSets(fieldsByRecordType, declaration.name, [fieldDeclaration.fieldName]);
    }
  }

  return fieldsByRecordType;
}

function extractDbdFieldTypesByRecordType(text) {
  const fieldTypesByRecordType = new Map();
  const recordTypeRegex = /recordtype\(\s*([A-Za-z0-9_]+)\s*\)\s*\{/g;
  let match;

  while ((match = recordTypeRegex.exec(text))) {
    const block = readBalancedBlock(text, recordTypeRegex.lastIndex - 1);
    if (!block) {
      continue;
    }

    const recordType = match[1];
    const fieldRegex = /field\(\s*([A-Z0-9_]+)\s*,\s*(DBF_[A-Z0-9_]+)\s*\)/g;
    let fieldMatch;

    while ((fieldMatch = fieldRegex.exec(block.body))) {
      addToMapOfMaps(
        fieldTypesByRecordType,
        recordType,
        new Map([[fieldMatch[1], fieldMatch[2]]]),
      );
    }

    recordTypeRegex.lastIndex = block.endIndex;
  }

  return fieldTypesByRecordType;
}

function extractDbdFieldInitialValuesByRecordType(text) {
  const fieldInitialValuesByRecordType = new Map();
  const recordTypeRegex = /recordtype\(\s*([A-Za-z0-9_]+)\s*\)\s*\{/g;
  let match;

  while ((match = recordTypeRegex.exec(text))) {
    const block = readBalancedBlock(text, recordTypeRegex.lastIndex - 1);
    if (!block) {
      continue;
    }

    const recordType = match[1];
    const fieldRegex = /field\(\s*([A-Z0-9_]+)\s*,\s*DBF_[A-Z0-9_]+\s*\)\s*\{/g;
    let fieldMatch;

    while ((fieldMatch = fieldRegex.exec(block.body))) {
      const fieldBlock = readBalancedBlock(block.body, fieldRegex.lastIndex - 1);
      if (!fieldBlock) {
        continue;
      }

      const initialMatch = fieldBlock.body.match(
        /initial\(\s*"((?:[^"\\]|\\.)*)"\s*\)/,
      );
      if (initialMatch) {
        addToMapOfMaps(
          fieldInitialValuesByRecordType,
          recordType,
          new Map([[fieldMatch[1], initialMatch[1]]]),
        );
      }

      fieldRegex.lastIndex = fieldBlock.endIndex;
    }

    recordTypeRegex.lastIndex = block.endIndex;
  }

  return fieldInitialValuesByRecordType;
}

function extractDeviceSupportDefinitions(uri, text) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const definitions = [];
  const definitionRegex =
    /epicsExportAddress\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = definitionRegex.exec(text))) {
    const exportType = match[1];
    const supportName = match[2];
    if (!isDeviceSupportExportType(exportType)) {
      continue;
    }

    definitions.push({
      name: supportName,
      exportType,
      absolutePath: normalizeFsPath(uri.fsPath),
      relativePath: normalizePath(vscode.workspace.asRelativePath(uri, false)),
      line: getLineNumberAtOffset(text, match.index),
    });
  }

  return definitions;
}

function extractDriverDefinitions(uri, text) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const definitions = [];
  const definitionRegex =
    /epicsExportAddress\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = definitionRegex.exec(text))) {
    const exportType = match[1];
    const driverName = match[2];
    if (!isDriverExportType(exportType)) {
      continue;
    }

    definitions.push({
      name: driverName,
      exportType,
      absolutePath: normalizeFsPath(uri.fsPath),
      relativePath: normalizePath(vscode.workspace.asRelativePath(uri, false)),
      line: getLineNumberAtOffset(text, match.index),
    });
  }

  return definitions;
}

function extractRegistrarDefinitions(uri, text) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const definitions = [];
  const definitionRegex =
    /epicsExportRegistrar\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = definitionRegex.exec(text))) {
    const registrarName = match[1];
    definitions.push({
      name: registrarName,
      absolutePath: normalizeFsPath(uri.fsPath),
      relativePath: normalizePath(vscode.workspace.asRelativePath(uri, false)),
      line: getLineNumberAtOffset(text, match.index),
    });
  }

  return definitions;
}

function extractFunctionDefinitions(uri, text) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const definitions = [];
  const definitionRegex =
    /epicsRegisterFunction\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = definitionRegex.exec(text))) {
    const functionName = match[1];
    definitions.push({
      name: functionName,
      absolutePath: normalizeFsPath(uri.fsPath),
      relativePath: normalizePath(vscode.workspace.asRelativePath(uri, false)),
      line: getLineNumberAtOffset(text, match.index),
    });
  }

  return definitions;
}

function extractVariableDefinitions(uri, text) {
  if (!uri || uri.scheme !== "file") {
    return [];
  }

  const definitions = [];
  const definitionRegex =
    /epicsExportAddress\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)/g;
  let match;

  while ((match = definitionRegex.exec(text))) {
    const exportType = match[1];
    const variableName = match[2];
    if (!isVariableExportType(exportType)) {
      continue;
    }

    definitions.push({
      name: variableName,
      exportType,
      absolutePath: normalizeFsPath(uri.fsPath),
      relativePath: normalizePath(vscode.workspace.asRelativePath(uri, false)),
      line: getLineNumberAtOffset(text, match.index),
    });
  }

  return definitions;
}

function isDeviceSupportExportType(exportType) {
  return /(?:^|_)dset$/i.test(String(exportType || ""));
}

function isDriverExportType(exportType) {
  return String(exportType || "") === "drvet";
}

function isVariableExportType(exportType) {
  return (
    Boolean(exportType) &&
    !isDeviceSupportExportType(exportType) &&
    !isDriverExportType(exportType)
  );
}

function readBalancedBlock(text, openingBraceIndex) {
  if (text[openingBraceIndex] !== "{") {
    return undefined;
  }

  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = openingBraceIndex; index < text.length; index += 1) {
    const character = text[index];

    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }

      if (character === "\\") {
        escaped = true;
        continue;
      }

      if (character === "\"") {
        inString = false;
      }

      continue;
    }

    if (character === "\"") {
      inString = true;
      continue;
    }

    if (character === "{") {
      depth += 1;
      continue;
    }

    if (character === "}") {
      depth -= 1;
      if (depth === 0) {
        return {
          body: text.slice(openingBraceIndex + 1, index),
          endIndex: index + 1,
        };
      }
    }
  }

  return undefined;
}

function extractAlternationTerms(pattern) {
  if (!pattern) {
    return [];
  }

  const normalized = pattern
    .replace(/^\\b\(/, "")
    .replace(/\)\\b$/, "")
    .split("|")
    .map((term) => term.replace(/\\b/g, "").trim())
    .filter((term) => /^[A-Za-z0-9_]+$/.test(term));

  return normalized;
}

function findEnclosingRecordType(textBeforePosition) {
  const cursorOffset = textBeforePosition.length;
  const declarations = extractRecordDeclarations(textBeforePosition);

  for (let index = declarations.length - 1; index >= 0; index -= 1) {
    const declaration = declarations[index];
    if (
      declaration.recordStart < cursorOffset &&
      declaration.recordEnd >= cursorOffset
    ) {
      return declaration.recordType;
    }
  }

  return undefined;
}

function findEnclosingRecordDeclaration(text, cursorOffset) {
  const declarations = extractRecordDeclarations(text);

  for (let index = declarations.length - 1; index >= 0; index -= 1) {
    const declaration = declarations[index];
    if (
      declaration.recordStart < cursorOffset &&
      declaration.recordEnd >= cursorOffset
    ) {
      return declaration;
    }
  }

  return undefined;
}

function isLinkField(fieldName) {
  if (!fieldName) {
    return false;
  }

  if (
    ["INP", "OUT", "FLNK", "SELL", "DOL", "SDIS", "SIOL", "TSEL"].includes(
      fieldName,
    )
  ) {
    return true;
  }

  return (
    /^INP[A-U]$/.test(fieldName) ||
    /^OUT[A-U]$/.test(fieldName) ||
    /^DOL[0-9A-F]$/.test(fieldName) ||
    /^LNK[0-9A-F]$/.test(fieldName)
  );
}

function getProjectFilePathEntries(snapshot, document, fileKind, baseDirectory) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const artifacts = project
    ? project.runtimeArtifacts
    : snapshot.projectModel.runtimeArtifacts;
  const entries = [];

  for (const artifact of artifacts) {
    if (!doesArtifactMatchFileKind(artifact, fileKind)) {
      continue;
    }

    entries.push({
      insertPath: baseDirectory
        ? getRelativePathFromBaseDirectory(baseDirectory, artifact.absoluteRuntimePath)
        : getRelativePathFromDocument(document, artifact.absoluteRuntimePath),
      absolutePath: getReadableAbsolutePathForArtifact(artifact),
      detail: artifact.detail,
      documentation: artifact.documentation,
    });
  }

  return entries;
}

function doesArtifactMatchFileKind(artifact, fileKind) {
  switch (fileKind) {
    case "dbLoadDatabase":
      return artifact.kind === "dbd";
    case "dbLoadRecords":
      return artifact.kind === "database";
    case "dbLoadTemplate":
      return artifact.kind === "substitutions";
    default:
      return false;
  }
}

function resolveStartupPath(snapshot, document, load, state) {
  if (document.uri.scheme !== "file") {
    return undefined;
  }

  const envVariables = state?.envVariables || new Map();
  const currentDirectory =
    state?.currentDirectory || normalizeFsPath(path.dirname(document.uri.fsPath));
  const expandedPath = expandStartupValue(load.path, envVariables);
  if (!expandedPath) {
    return undefined;
  }

  const absoluteTargetPath = normalizeFsPath(
    path.isAbsolute(expandedPath)
      ? expandedPath
      : path.resolve(currentDirectory, expandedPath),
  );

  const filesystemResolution = resolveFilesystemStartupPath(absoluteTargetPath, load.command);
  if (filesystemResolution) {
    return filesystemResolution;
  }

  if (snapshot.workspaceFilesByAbsolutePath.has(absoluteTargetPath)) {
    return {
      kind: "workspaceFile",
      entry: snapshot.workspaceFilesByAbsolutePath.get(absoluteTargetPath),
    };
  }

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const artifacts = project
    ? project.runtimeArtifacts
    : snapshot.projectModel.runtimeArtifacts;

  for (const artifact of artifacts) {
    if (!doesArtifactMatchFileKind(artifact, load.command)) {
      continue;
    }

    if (normalizeFsPath(artifact.absoluteRuntimePath) === absoluteTargetPath) {
      return {
        kind: "projectArtifact",
        artifact,
      };
    }
  }

  return undefined;
}

function loadStartupReleaseVariableData(snapshot, document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return {
      values: new Map(),
      sources: new Map(),
    };
  }

  const rootPath = findStartupOwningRootPath(snapshot, document);
  if (!rootPath) {
    return {
      values: new Map(),
      sources: new Map(),
    };
  }

  return loadReleaseVariablesWithSources(rootPath);
}

function findStartupOwningRootPath(snapshot, document) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (project?.rootPath) {
    return normalizeFsPath(project.rootPath);
  }

  let currentDirectory = normalizeFsPath(path.dirname(document.uri.fsPath));
  while (currentDirectory) {
    const releasePath = normalizeFsPath(path.join(currentDirectory, "configure", "RELEASE"));
    const releaseLocalPath = normalizeFsPath(path.join(currentDirectory, "configure", "RELEASE.local"));
    if (isExistingFile(releasePath) || isExistingFile(releaseLocalPath)) {
      return currentDirectory;
    }

    const parentDirectory = normalizeFsPath(path.dirname(currentDirectory));
    if (!parentDirectory || parentDirectory === currentDirectory) {
      break;
    }
    currentDirectory = parentDirectory;
  }

  return undefined;
}

function loadReleaseVariablesWithSources(rootPath) {
  const rawValues = new Map([["TOP", normalizeFsPath(rootPath)]]);
  const sources = new Map();
  const visitedFiles = new Set();
  const resolvedCache = new Map();
  const releaseEntryPaths = [
    normalizeFsPath(path.join(rootPath, "configure", "RELEASE")),
    normalizeFsPath(path.join(rootPath, "configure", "RELEASE.local")),
  ];

  const normalizeResolvedReleaseValue = (value) => {
    const normalizedValue = stripOptionalQuotes(String(value || "").trim());
    if (!normalizedValue || normalizedValue.includes("$(") || normalizedValue.includes("${")) {
      return normalizedValue;
    }
    const absolutePath = computeAbsoluteVariablePath(normalizedValue, rootPath);
    return absolutePath || normalizedValue;
  };

  const resolveReleaseVariableValue = (variableName, stack = new Set()) => {
    if (resolvedCache.has(variableName)) {
      return resolvedCache.get(variableName);
    }

    if (stack.has(variableName)) {
      return rawValues.get(variableName);
    }

    if (!rawValues.has(variableName)) {
      return process.env[variableName];
    }

    stack.add(variableName);
    const rawValue = String(rawValues.get(variableName) || "");
    const expandedValue = rawValue.replace(
      /\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)/g,
      (match, parenthesizedName, parenthesizedDefault, bracedName, bracedDefault, bareName) => {
        const nestedName = parenthesizedName || bracedName || bareName;
        const defaultValue =
          parenthesizedName !== undefined
            ? parenthesizedDefault
            : bracedName !== undefined
              ? bracedDefault
              : undefined;
        const resolvedValue = resolveReleaseVariableValue(nestedName, new Set(stack));
        if (resolvedValue !== undefined) {
          return resolvedValue;
        }
        if (defaultValue !== undefined) {
          return defaultValue;
        }
        return match;
      },
    );
    stack.delete(variableName);

    const normalizedValue = normalizeResolvedReleaseValue(expandedValue);
    resolvedCache.set(variableName, normalizedValue);
    return normalizedValue;
  };

  const expandReleaseInlineValue = (rawValue) => {
    return String(rawValue || "").replace(
      /\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)/g,
      (match, parenthesizedName, parenthesizedDefault, bracedName, bracedDefault, bareName) => {
        const variableName = parenthesizedName || bracedName || bareName;
        const defaultValue =
          parenthesizedName !== undefined
            ? parenthesizedDefault
            : bracedName !== undefined
              ? bracedDefault
              : undefined;
        const resolvedValue = resolveReleaseVariableValue(variableName);
        if (resolvedValue !== undefined) {
          return resolvedValue;
        }
        if (defaultValue !== undefined) {
          return defaultValue;
        }
        return match;
      },
    );
  };

  const loadReleaseFile = (filePath) => {
    const normalizedPath = normalizeFsPath(filePath);
    if (!normalizedPath || visitedFiles.has(normalizedPath) || !isExistingFile(normalizedPath)) {
      return;
    }

    visitedFiles.add(normalizedPath);
    const text = readTextFile(normalizedPath);
    if (text === undefined) {
      return;
    }

    const lines = text.replace(/\\\n/g, " ").split(/\r?\n/);
    for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
      const rawLine = lines[lineIndex];
      const line = stripMakeComments(rawLine).trim();
      if (!line) {
        continue;
      }

      const includeMatch = line.match(/^-?include\s+(.+)$/);
      if (includeMatch) {
        const includeEntries = splitMakeValue(includeMatch[1]);
        for (const includeEntry of includeEntries) {
          const expandedIncludeEntry = stripOptionalQuotes(expandReleaseInlineValue(includeEntry).trim());
          if (!expandedIncludeEntry || expandedIncludeEntry.includes("$(") || expandedIncludeEntry.includes("${")) {
            continue;
          }
          const includePath = normalizeFsPath(
            path.isAbsolute(expandedIncludeEntry)
              ? expandedIncludeEntry
              : path.resolve(path.dirname(normalizedPath), expandedIncludeEntry),
          );
          loadReleaseFile(includePath);
        }
        continue;
      }

      const assignmentMatch = line.match(/^([A-Za-z0-9_.-]+)\s*(\+?=|:=|\?=)\s*(.*)$/);
      if (!assignmentMatch) {
        continue;
      }

      const variableName = assignmentMatch[1];
      const operator = assignmentMatch[2];
      const assignedValue = splitMakeValue(assignmentMatch[3]).join(" ");
      if (operator === "?=" && rawValues.has(variableName)) {
        continue;
      }

      if (operator === "+=" && rawValues.has(variableName)) {
        rawValues.set(
          variableName,
          [String(rawValues.get(variableName) || "").trim(), assignedValue.trim()]
            .filter(Boolean)
            .join(" "),
        );
      } else {
        rawValues.set(variableName, assignedValue);
      }
      sources.set(variableName, {
        sourceKind: "RELEASE",
        sourcePath: normalizedPath,
        line: lineIndex + 1,
        rawValue: assignedValue,
      });
      resolvedCache.clear();
    }
  };

  for (const entryPath of releaseEntryPaths) {
    loadReleaseFile(entryPath);
  }

  const values = new Map();
  for (const variableName of rawValues.keys()) {
    values.set(variableName, resolveReleaseVariableValue(variableName));
  }

  return {
    values,
    sources,
  };
}

function createInitialStartupExecutionState(snapshot, document) {
  const startupReleaseData = loadStartupReleaseVariableData(snapshot, document);
  return {
    envVariables: new Map(startupReleaseData.values),
    envVariableSources: new Map(startupReleaseData.sources),
    currentDirectory:
      document.uri.scheme === "file"
        ? normalizeFsPath(path.dirname(document.uri.fsPath))
        : undefined,
    includedFiles: new Set(),
  };
}

function createStartupExecutionState(snapshot, document, untilOffset) {
  const state = createInitialStartupExecutionState(snapshot, document);

  for (const statement of extractStartupStatements(document.getText())) {
    if (statement.start >= untilOffset) {
      break;
    }

    applyStartupStatement(snapshot, document, statement, state);
  }

  return state;
}

function setStartupEnvVariable(state, variableName, rawValue, sourceInfo) {
  const resolvedValue = expandStartupValue(rawValue, state.envVariables);
  state.envVariables.set(variableName, resolvedValue);
  if (state.envVariableSources) {
    state.envVariableSources.set(variableName, sourceInfo);
  }
}

function applyStartupStatement(snapshot, document, statement, state) {
  switch (statement.kind) {
    case "include": {
      const resolution = resolveStartupPath(snapshot, document, statement, state);
      if (resolution) {
        applyIncludedStartupState(resolution.absolutePath, state);
      }
      break;
    }

    case "envSet":
      setStartupEnvVariable(state, statement.name, statement.value, {
        sourceKind: "startup",
        sourcePath:
          document?.uri?.scheme === "file"
            ? normalizeFsPath(document.uri.fsPath)
            : undefined,
        line: statement.lineNumber,
        rawValue: statement.value,
      });
      break;

    case "cd": {
      const resolution = resolveStartupPath(snapshot, document, statement, state);
      if (resolution && resolution.isDirectory) {
        state.currentDirectory = resolution.absolutePath;
      }
      break;
    }

    default:
      break;
  }
}

function applyIncludedStartupState(absolutePath, state) {
  const normalizedPath = normalizeFsPath(absolutePath);
  if (state.includedFiles.has(normalizedPath)) {
    return;
  }

  const text = readTextFile(normalizedPath);
  if (text === undefined) {
    return;
  }

  state.includedFiles.add(normalizedPath);
  for (const statement of extractStartupStatements(text)) {
    switch (statement.kind) {
      case "include": {
        const expandedPath = expandStartupValue(statement.path, state.envVariables);
        if (!expandedPath) {
          break;
        }

        const absoluteIncludePath = normalizeFsPath(
          path.isAbsolute(expandedPath)
            ? expandedPath
            : path.resolve(state.currentDirectory || path.dirname(normalizedPath), expandedPath),
        );
        applyIncludedStartupState(absoluteIncludePath, state);
        break;
      }

      case "envSet":
        setStartupEnvVariable(state, statement.name, statement.value, {
          sourceKind: /^envPaths(?:\..+)?$/i.test(path.basename(normalizedPath))
            ? "envPaths"
            : "startup include",
          sourcePath: normalizedPath,
          line: statement.lineNumber,
          rawValue: statement.value,
        });
        break;

      case "cd": {
        const expandedPath = expandStartupValue(statement.path, state.envVariables);
        if (!expandedPath) {
          break;
        }

        const absoluteDirectoryPath = normalizeFsPath(
          path.isAbsolute(expandedPath)
            ? expandedPath
            : path.resolve(state.currentDirectory || path.dirname(normalizedPath), expandedPath),
        );
        if (fs.existsSync(absoluteDirectoryPath)) {
          try {
            if (fs.statSync(absoluteDirectoryPath).isDirectory()) {
              state.currentDirectory = absoluteDirectoryPath;
            }
          } catch (error) {
            // Ignore unreadable included-path entries.
          }
        }
        break;
      }

      default:
        break;
    }
  }
}

function resolveFilesystemStartupPath(absolutePath, command) {
  if (!fs.existsSync(absolutePath)) {
    return undefined;
  }

  let stats;
  try {
    stats = fs.statSync(absolutePath);
  } catch (error) {
    return undefined;
  }

  if (command === "cd" || command === "startupDirectory") {
    if (!stats.isDirectory()) {
      return undefined;
    }

    return {
      kind: "filesystem",
      absolutePath,
      isDirectory: true,
    };
  }

  if (stats.isDirectory()) {
    return undefined;
  }

  return {
    kind: "filesystem",
    absolutePath,
    isDirectory: false,
  };
}

function getNavigationTargetFromStartupResolution(resolution) {
  switch (resolution.kind) {
    case "filesystem":
      return { absolutePath: resolution.absolutePath };

    case "workspaceFile":
      return { absolutePath: resolution.entry.absolutePath };

    case "projectArtifact":
      if (
        resolution.artifact.absoluteRuntimePath &&
        fs.existsSync(resolution.artifact.absoluteRuntimePath)
      ) {
        return { absolutePath: resolution.artifact.absoluteRuntimePath };
      }

      if (resolution.artifact.sourceRelativePath) {
        const sourcePath = normalizeFsPath(
          path.join(
            findRootPathForArtifact(resolution.artifact),
            resolution.artifact.sourceRelativePath,
          ),
        );
        if (fs.existsSync(sourcePath) && fs.statSync(sourcePath).isFile()) {
          return { absolutePath: sourcePath };
        }
      }

      return undefined;

    default:
      return undefined;
  }
}

function findRootPathForArtifact(artifact) {
  const runtimePath = normalizeFsPath(artifact.absoluteRuntimePath);
  if (runtimePath.endsWith(`/${artifact.runtimeRelativePath}`)) {
    return runtimePath.slice(0, -artifact.runtimeRelativePath.length - 1);
  }

  return path.dirname(runtimePath);
}

function expandStartupValue(value, envVariables) {
  return expandEpicsValue(value, [envVariables, process.env]);
}

function expandEpicsValue(value, sources) {
  if (!value) {
    return value;
  }

  let expandedValue = String(value);
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const nextValue = expandedValue.replace(
      /\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_]*)/g,
      (match, parenthesizedName, parenthesizedDefault, bracedName, bracedDefault, shellName) => {
        const variableName = parenthesizedName || bracedName || shellName;
        const defaultValue =
          parenthesizedName !== undefined
            ? parenthesizedDefault
            : bracedName !== undefined
              ? bracedDefault
              : undefined;
        const resolvedValue = resolveEpicsValueSource(sources, variableName);
        if (resolvedValue !== undefined) {
          return resolvedValue;
        }

        if (defaultValue !== undefined) {
          return defaultValue;
        }

        return match;
      },
    );

    if (nextValue === expandedValue) {
      break;
    }

    expandedValue = nextValue;
  }

  return expandedValue;
}

function resolveEpicsValueSource(sources, variableName) {
  for (const source of sources || []) {
    if (!source) {
      continue;
    }

    if (typeof source === "function") {
      const resolvedValue = source(variableName);
      if (resolvedValue !== undefined) {
        return resolvedValue;
      }
      continue;
    }

    if (source instanceof Map) {
      if (source.has(variableName)) {
        return source.get(variableName);
      }
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(source, variableName)) {
      return source[variableName];
    }
  }

  return undefined;
}

function computeAbsoluteVariablePath(value, baseDirectory) {
  if (!value || !looksLikePathValue(value) || containsMakeVariableReference(value)) {
    return undefined;
  }

  const expandedHomeValue = value.startsWith("~/")
    ? path.join(process.env.HOME || "", value.slice(2))
    : value;
  const absolutePath = path.isAbsolute(expandedHomeValue)
    ? expandedHomeValue
    : path.resolve(baseDirectory, expandedHomeValue);

  return normalizeFsPath(absolutePath);
}

function looksLikePathValue(value) {
  return (
    typeof value === "string" &&
    (value.startsWith(".") ||
      value.startsWith("/") ||
      value.startsWith("~/") ||
      value.includes("/") ||
      value.includes("\\"))
  );
}

function findProjectForUri(projectModel, uri) {
  if (!uri || uri.scheme !== "file") {
    return undefined;
  }

  const filePath = normalizeFsPath(uri.fsPath);
  let bestMatch;

  for (const application of projectModel.applications) {
    if (!isPathWithinRoot(filePath, application.rootPath)) {
      continue;
    }

    if (!bestMatch || application.rootPath.length > bestMatch.rootPath.length) {
      bestMatch = application;
    }
  }

  return bestMatch;
}

function extractStartupLoadStatements(text) {
  return extractStartupStatements(text)
    .filter((statement) => statement.kind === "load")
    .map((statement) => ({
      command: statement.command,
      path: statement.path,
      pathStart: statement.pathStart,
      pathEnd: statement.pathEnd,
    }));
}

function extractStartupRegisterCalls(text) {
  return extractStartupStatements(text)
    .filter((statement) => statement.kind === "register")
    .map((statement) => ({
      iocName: statement.iocName,
      functionName: statement.functionName,
      nameStart: statement.nameStart,
      nameEnd: statement.nameEnd,
    }));
}

function buildMissingStartupPathMessage(load) {
  switch (load.command) {
    case "startupInclude":
      return `Cannot resolve startup include path "${load.path}".`;
    case "cd":
    case "startupDirectory":
      return `Cannot resolve startup directory "${load.path}".`;
    case "dbLoadDatabase":
      return `Cannot resolve startup DBD path "${load.path}" from the EPICS project model or workspace files.`;
    case "dbLoadRecords":
      return `Cannot resolve startup database path "${load.path}" from the EPICS project model or workspace files.`;
    case "dbLoadTemplate":
      return `Cannot resolve startup substitutions path "${load.path}" from the EPICS project model or workspace files.`;
    default:
      return `Cannot resolve startup path "${load.path}".`;
  }
}

function buildStartupPathDiagnosticMessage(snapshot, document, load, project) {
  if (!project) {
    return buildMissingStartupPathMessage(load);
  }

  const matchingArtifact = findProjectArtifactByFileName(project, load);
  if (matchingArtifact) {
    return `Startup path "${load.path}" does not resolve from this file. The EPICS project model knows this file as "${getRelativePathFromDocument(document, matchingArtifact.absoluteRuntimePath)}".`;
  }

  const sourceCandidate = findProjectSourceCandidate(snapshot, project, load);
  if (sourceCandidate) {
    return `Startup path "${load.path}" does not match any installed EPICS file. Found source file "${sourceCandidate.relativePath}" in the project, but it is not installed by the Makefile rules for this application.`;
  }

  const dbdHint = findProjectDbdHint(project, load);
  if (dbdHint) {
    return dbdHint;
  }

  return buildMissingStartupPathMessage(load);
}

function parseSubstitutionLoads(text) {
  const loads = [];
  let globalMacros = new Map();

  for (const block of extractSubstitutionBlocksWithRanges(text)) {
    if (block.kind === "global") {
      globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body));
      continue;
    }

    if (block.kind !== "file" || !block.templatePath) {
      continue;
    }

    const rows = parseSubstitutionFileBlockRows(block.body).map((rowMacros) =>
      mergeMacroMaps(globalMacros, rowMacros),
    );
    loads.push({
      templatePath: block.templatePath,
      rows,
    });
  }

  return loads;
}

function extractSubstitutionBlocks(text) {
  return extractSubstitutionBlocksWithRanges(text).map((block) => ({
    kind: block.kind,
    templatePath: block.templatePath,
    body: block.body,
  }));
}

function extractSubstitutionBlocksWithRanges(text) {
  const blocks = [];
  const blockPattern =
    /(?:^|\n)\s*(global|file)(?:\s+("(?:[^"\\]|\\.)*"|[^\s{]+))?\s*\{/g;
  let match;

  while ((match = blockPattern.exec(text))) {
    const braceIndex = text.indexOf("{", match.index);
    if (braceIndex < 0) {
      break;
    }

    const blockEnd = findRecordBlockEnd(text, braceIndex);
    const bodyStart = braceIndex + 1;
    const bodyEnd = Math.max(bodyStart, blockEnd - 1);
    const rawTemplatePath = match[2];
    let templatePath;
    let templatePathStart;
    let templatePathEnd;

    if (rawTemplatePath) {
      const rawTemplatePathIndex = match[0].lastIndexOf(rawTemplatePath);
      if (rawTemplatePathIndex >= 0) {
        templatePathStart = match.index + rawTemplatePathIndex;
        templatePathEnd = templatePathStart + rawTemplatePath.length;
        templatePath = rawTemplatePath;

        if (
          rawTemplatePath.length >= 2 &&
          rawTemplatePath.startsWith("\"") &&
          rawTemplatePath.endsWith("\"")
        ) {
          templatePath = rawTemplatePath.slice(1, -1);
          templatePathStart += 1;
          templatePathEnd -= 1;
        }
      } else {
        templatePath = stripOptionalQuotes(rawTemplatePath);
      }
    }

    blocks.push({
      kind: match[1],
      templatePath,
      templatePathStart,
      templatePathEnd,
      body: text.slice(bodyStart, bodyEnd),
      bodyStart,
      bodyEnd,
    });
    blockPattern.lastIndex = blockEnd;
  }

  return blocks;
}

function parseSubstitutionFileBlockRows(body) {
  const parsedRows = parseSubstitutionFileBlockRowsDetailed(body);
  if (parsedRows.kind === "pattern") {
    return parsedRows.rows
      .map((row) => row.assignments)
      .filter((assignments) => assignments.size > 0);
  }

  return parsedRows.rows
    .map((row) => row.assignments)
    .filter((assignments) => assignments.size > 0);
}

function parseSubstitutionFileBlockRowsDetailed(body, baseOffset = 0) {
  const segments = extractTopLevelBraceSegments(body, baseOffset);
  if (segments.length === 0) {
    return {
      kind: "assignments",
      rows: [],
    };
  }

  if (/^\s*pattern\b/.test(body)) {
    const headerSegment = segments[0];
    const columns = tokenizeSubstitutionValues(headerSegment.text);
    return {
      kind: "pattern",
      columns,
      headerRangeStart:
        headerSegment.contentStart < headerSegment.contentEnd
          ? headerSegment.contentStart
          : headerSegment.braceStart,
      headerRangeEnd:
        headerSegment.contentStart < headerSegment.contentEnd
          ? headerSegment.contentEnd
          : headerSegment.braceEnd,
      rows: segments.slice(1).map((segment) => {
        const values = tokenizeSubstitutionValues(segment.text);
        return {
          assignments: createPatternMacroAssignments(columns, values),
          values,
          nameRanges: new Map(),
          rangeStart:
            segment.contentStart < segment.contentEnd
              ? segment.contentStart
              : segment.braceStart,
          rangeEnd:
            segment.contentStart < segment.contentEnd
              ? segment.contentEnd
              : segment.braceEnd,
        };
      }),
    };
  }

  return {
    kind: "assignments",
    rows: segments.map((segment) => {
      const extracted = extractNamedAssignmentsWithRanges(
        segment.text,
        segment.contentStart,
      );
      return {
        assignments: extracted.assignments,
        nameRanges: extracted.nameRanges,
        rangeStart:
          segment.contentStart < segment.contentEnd
            ? segment.contentStart
            : segment.braceStart,
        rangeEnd:
          segment.contentStart < segment.contentEnd
            ? segment.contentEnd
            : segment.braceEnd,
      };
    }),
  };
}

function extractTopLevelBraceContents(text) {
  return extractTopLevelBraceSegments(text).map((segment) => segment.text);
}

function extractTopLevelBraceSegments(text, baseOffset = 0) {
  const segments = [];
  let searchIndex = 0;

  while (searchIndex < text.length) {
    const braceIndex = text.indexOf("{", searchIndex);
    if (braceIndex < 0) {
      break;
    }

    const blockEnd = findRecordBlockEnd(text, braceIndex);
    const contentStart = braceIndex + 1;
    const contentEnd = Math.max(contentStart, blockEnd - 1);
    segments.push({
      text: text.slice(contentStart, contentEnd),
      braceStart: baseOffset + braceIndex,
      braceEnd: baseOffset + blockEnd,
      contentStart: baseOffset + contentStart,
      contentEnd: baseOffset + contentEnd,
    });
    searchIndex = blockEnd;
  }

  return segments;
}

function getSubstitutionTemplateReferenceAtPosition(document, position) {
  if (!isSubstitutionsDocument(document)) {
    return undefined;
  }

  const offset = document.offsetAt(position);
  for (const block of extractSubstitutionBlocksWithRanges(document.getText())) {
    if (
      block.kind !== "file" ||
      !block.templatePath ||
      block.templatePathStart === undefined ||
      block.templatePathEnd === undefined
    ) {
      continue;
    }

    if (offset < block.templatePathStart || offset >= block.templatePathEnd) {
      continue;
    }

    return {
      templatePath: block.templatePath,
      start: block.templatePathStart,
      end: block.templatePathEnd,
    };
  }

  return undefined;
}

function extractNamedAssignmentsWithRanges(text, baseOffset = 0) {
  const assignments = new Map();
  const nameRanges = new Map();
  if (!text) {
    return {
      assignments,
      nameRanges,
    };
  }

  const pattern =
    /([A-Za-z_][A-Za-z0-9_]*)\s*=\s*("((?:[^"\\]|\\.)*)"|[^,\s{}]*)/g;
  let match;

  while ((match = pattern.exec(text))) {
    const name = match[1];
    assignments.set(name, match[3] ?? match[2]);
    nameRanges.set(name, {
      start: baseOffset + match.index,
      end: baseOffset + match.index + name.length,
    });
  }

  return {
    assignments,
    nameRanges,
  };
}

function extractNamedAssignments(text) {
  return extractNamedAssignmentsWithRanges(text).assignments;
}

function extractAssignedMacroNames(text) {
  if (!text) {
    return new Set();
  }

  const names = new Set();
  let segmentStart = 0;
  let escaped = false;

  const flushSegment = (segmentEnd) => {
    const segment = text.slice(segmentStart, segmentEnd);
    const match = segment.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=/);
    if (match?.[1]) {
      names.add(match[1]);
    }
  };

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (character === ",") {
      flushSegment(index);
      segmentStart = index + 1;
    }
  }

  flushSegment(text.length);
  return names;
}

function createPatternMacroAssignments(columns, valueSegment) {
  const assignments = new Map();
  const values = Array.isArray(valueSegment)
    ? valueSegment
    : tokenizeSubstitutionValues(valueSegment);

  for (let index = 0; index < columns.length; index += 1) {
    if (values[index] === undefined) {
      continue;
    }

    assignments.set(columns[index], values[index]);
  }

  return assignments;
}

function tokenizeSubstitutionValues(text) {
  return splitSubstitutionCommaSeparatedItems(text).map((value) => {
    const quotedMatch = value.match(/^"((?:[^"\\]|\\.)*)"$/);
    return quotedMatch ? quotedMatch[1] : value;
  });
}

function mergeMacroMaps(baseMacros, additionalMacros) {
  const merged = new Map(baseMacros || []);
  for (const [name, value] of additionalMacros || []) {
    merged.set(name, value);
  }

  return merged;
}

function stripOptionalQuotes(value) {
  if (
    typeof value === "string" &&
    value.length >= 2 &&
    value.startsWith("\"") &&
    value.endsWith("\"")
  ) {
    return value.slice(1, -1);
  }

  return value;
}

function extractStartupStatements(text) {
  const statements = [];
  let offset = 0;
  let lineNumber = 1;

  for (const line of text.split(/\r?\n/)) {
    const lineOffset = offset;
    const currentLineNumber = lineNumber;
    offset += line.length + 1;
    lineNumber += 1;

    if (/^\s*#/.test(line) || !line.trim()) {
      continue;
    }

    let match = line.match(/^\s*<\s*"?([^"\n]+)"?/);
    if (match) {
      const pathValue = match[1].trim();
      const pathStart = lineOffset + line.indexOf(pathValue);
      statements.push({
        kind: "include",
        command: "startupInclude",
        path: pathValue,
        pathStart,
        pathEnd: pathStart + pathValue.length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(
      /^\s*epicsEnvSet\(\s*"?([A-Za-z_][A-Za-z0-9_]*)"?\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)/,
    );
    if (match) {
      const nameStart = lineOffset + line.indexOf(match[1]);
      statements.push({
        kind: "envSet",
        name: match[1],
        value: match[2],
        nameStart,
        nameEnd: nameStart + match[1].length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(/^\s*cd(?:\s+|\(\s*)(?:"([^"\n]+)"|([^\s#\)]+))/);
    if (match) {
      const pathValue = (match[1] || match[2] || "").trim();
      const pathStart = lineOffset + line.indexOf(pathValue);
      statements.push({
        kind: "cd",
        command: "cd",
        path: pathValue,
        pathStart,
        pathEnd: pathStart + pathValue.length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(/^\s*dbLoadDatabase(?:\(\s*|\s+)"([^"\n]+)"/);
    if (match) {
      const pathValue = match[1];
      const pathStart = lineOffset + line.indexOf(pathValue);
      statements.push({
        kind: "load",
        command: "dbLoadDatabase",
        path: pathValue,
        pathStart,
        pathEnd: pathStart + pathValue.length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(
      /^\s*dbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"((?:[^"\\]|\\.)*)")?/,
    );
    if (match) {
      const pathValue = match[1];
      const pathStart = lineOffset + line.indexOf(pathValue);
      const macroStartSearchIndex = pathStart + pathValue.length + 1;
      const macroMatch = line
        .slice(Math.max(0, macroStartSearchIndex - lineOffset))
        .match(/^\s*,\s*"((?:[^"\\]|\\.)*)"/);
      const macroValueStart = macroMatch
        ? macroStartSearchIndex + macroMatch[0].indexOf("\"") + 1
        : undefined;
      const macroValueEnd =
        macroValueStart !== undefined ? macroValueStart + (match[2] || "").length : undefined;
      statements.push({
        kind: "load",
        command: "dbLoadRecords",
        path: pathValue,
        macros: match[2] || "",
        pathStart,
        pathEnd: pathStart + pathValue.length,
        macroValueStart,
        macroValueEnd,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(/^\s*dbLoadTemplate\(\s*"([^"\n]+)"/);
    if (match) {
      const pathValue = match[1];
      const pathStart = lineOffset + line.indexOf(pathValue);
      statements.push({
        kind: "load",
        command: "dbLoadTemplate",
        path: pathValue,
        pathStart,
        pathEnd: pathStart + pathValue.length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
      continue;
    }

    match = line.match(
      /^\s*([A-Za-z_][A-Za-z0-9_]*)_registerRecordDeviceDriver\s*\(/,
    );
    if (match) {
      const functionName = `${match[1]}_registerRecordDeviceDriver`;
      const nameStart = lineOffset + line.indexOf(functionName);
      statements.push({
        kind: "register",
        iocName: match[1],
        functionName,
        nameStart,
        nameEnd: nameStart + functionName.length,
        start: lineOffset,
        lineNumber: currentLineNumber,
      });
    }
  }

  return statements;
}

function getStartupStatementAtPosition(document, position) {
  const offset = document.offsetAt(position);

  for (const statement of extractStartupStatements(document.getText())) {
    switch (statement.kind) {
      case "include":
      case "cd":
      case "load":
        if (offset >= statement.pathStart && offset <= statement.pathEnd) {
          return statement;
        }
        break;

      default:
        break;
    }
  }

  return undefined;
}

function parseMakeAssignments(text) {
  const assignments = new Map();
  const lines = text.replace(/\\\n/g, " ").split(/\r?\n/);

  for (const rawLine of lines) {
    const line = stripMakeComments(rawLine).trim();
    if (!line) {
      continue;
    }

    const match = line.match(/^([A-Za-z0-9_.-]+)\s*(\+?=|:=|\?=)\s*(.*)$/);
    if (!match) {
      continue;
    }

    const variableName = match[1];
    const operator = match[2];
    const values = splitMakeValue(match[3]);

    if (operator === "?=") {
      if (!assignments.has(variableName)) {
        assignments.set(variableName, values);
      }
      continue;
    }

    if (operator === "=" || operator === ":=" || !assignments.has(variableName)) {
      assignments.set(variableName, values);
      continue;
    }

    assignments.set(variableName, [...assignments.get(variableName), ...values]);
  }

  return assignments;
}

function mergeProjectResourceMap(targetMap, sourceMap) {
  for (const [name, entry] of sourceMap.entries()) {
    if (!targetMap.has(name)) {
      targetMap.set(name, entry);
    }
  }
}

function resolveReleaseModuleRoots(rootPath, releaseVariables) {
  const resolvedPaths = new Map();
  const resolving = new Set();
  const roots = [];

  const resolveVariable = (variableName) => {
    if (resolvedPaths.has(variableName)) {
      return resolvedPaths.get(variableName);
    }

    if (resolving.has(variableName) || !releaseVariables.has(variableName)) {
      return undefined;
    }

    resolving.add(variableName);
    const rawValue = releaseVariables.get(variableName);
    const expandedValue = rawValue.replace(
      /\$\(([^)]+)\)|\$\{([^}]+)\}/g,
      (_, parenthesizedName, bracedName) => {
        const nestedName = parenthesizedName || bracedName;
        return (
          resolveVariable(nestedName) ||
          process.env[nestedName] ||
          ""
        );
      },
    );
    resolving.delete(variableName);

    if (expandedValue.includes("$(") || expandedValue.includes("${")) {
      return undefined;
    }

    const resolvedPath = path.isAbsolute(expandedValue)
      ? expandedValue
      : path.resolve(rootPath, expandedValue);
    resolvedPaths.set(variableName, resolvedPath);
    return resolvedPath;
  };

  for (const variableName of releaseVariables.keys()) {
    const resolvedPath = resolveVariable(variableName);
    if (!resolvedPath || !fs.existsSync(resolvedPath)) {
      continue;
    }

    let stats;
    try {
      stats = fs.statSync(resolvedPath);
    } catch (error) {
      continue;
    }

    if (!stats.isDirectory()) {
      continue;
    }

    roots.push({
      variableName,
      rootPath: normalizeFsPath(resolvedPath),
    });
  }

  return roots;
}

function buildProjectDbdEntries({ rootPath, runtimeArtifacts, releaseRoots, rootRelativePath }) {
  const entries = new Map();

  for (const artifact of runtimeArtifacts) {
    if (artifact.kind !== "dbd") {
      continue;
    }

    entries.set(artifact.runtimeFileName, {
      name: artifact.runtimeFileName,
      detail: artifact.detail,
      documentation: artifact.documentation,
      absolutePath: fs.existsSync(artifact.absoluteRuntimePath)
        ? artifact.absoluteRuntimePath
        : undefined,
    });
  }

  scanDbdDirectoryEntries(entries, path.join(rootPath, "dbd"), rootRelativePath || ".");
  scanLocalMakefileDirectoryDbdEntries(entries, rootPath, rootRelativePath || ".");

  for (const releaseRoot of releaseRoots) {
    scanDbdDirectoryEntries(
      entries,
      path.join(releaseRoot.rootPath, "dbd"),
      releaseRoot.variableName,
    );
  }

  return entries;
}

function buildProjectLibEntries({ rootPath, releaseRoots, rootRelativePath }) {
  const entries = new Map();

  scanLibraryDirectoryEntries(entries, path.join(rootPath, "lib"), rootRelativePath || ".");

  for (const releaseRoot of releaseRoots) {
    scanLibraryDirectoryEntries(
      entries,
      path.join(releaseRoot.rootPath, "lib"),
      releaseRoot.variableName,
    );
  }

  return entries;
}

function scanDbdDirectoryEntries(entries, dbdRootPath, sourceLabel) {
  const normalizedDbdRootPath = normalizeFsPath(dbdRootPath);
  if (!fs.existsSync(normalizedDbdRootPath)) {
    return;
  }

  let fileNames;
  try {
    fileNames = fs.readdirSync(normalizedDbdRootPath);
  } catch (error) {
    return;
  }

  for (const fileName of fileNames) {
    if (path.extname(fileName).toLowerCase() !== ".dbd" || entries.has(fileName)) {
      continue;
    }

    entries.set(fileName, {
      name: fileName,
      detail: `${sourceLabel} dbd/${fileName}`,
      documentation: `Found in ${sourceLabel}/dbd`,
      absolutePath: normalizeFsPath(path.join(normalizedDbdRootPath, fileName)),
    });
  }
}

function scanLocalMakefileDirectoryDbdEntries(entries, rootPath, sourceLabel) {
  const normalizedRootPath = normalizeFsPath(rootPath);
  if (!normalizedRootPath || !fs.existsSync(normalizedRootPath)) {
    return;
  }

  const pendingDirectories = [normalizedRootPath];
  while (pendingDirectories.length > 0) {
    const directoryPath = pendingDirectories.pop();
    let directoryEntries;
    try {
      directoryEntries = fs.readdirSync(directoryPath, { withFileTypes: true });
    } catch (error) {
      continue;
    }

    const hasMakefile = directoryEntries.some(
      (entry) => entry.isFile() && entry.name === "Makefile",
    );
    if (hasMakefile) {
      for (const entry of directoryEntries) {
        if (
          !entry.isFile() ||
          path.extname(entry.name).toLowerCase() !== ".dbd" ||
          entries.has(entry.name)
        ) {
          continue;
        }

        const absolutePath = normalizeFsPath(path.join(directoryPath, entry.name));
        const relativeDirectory = normalizePath(
          path.relative(normalizedRootPath, directoryPath),
        );
        const detailDirectory =
          relativeDirectory && relativeDirectory !== "."
            ? `${sourceLabel}/${relativeDirectory}`
            : sourceLabel;
        entries.set(entry.name, {
          name: entry.name,
          detail: `${detailDirectory}/${entry.name}`,
          documentation: `Found beside Makefile in ${detailDirectory}`,
          absolutePath,
        });
      }
    }

    for (const entry of directoryEntries) {
      if (!entry.isDirectory()) {
        continue;
      }

      if (
        entry.name === ".git" ||
        entry.name === ".hg" ||
        entry.name === ".svn" ||
        entry.name === "node_modules" ||
        entry.name === "out" ||
        entry.name === "dist" ||
        entry.name === "dbd" ||
        /^O(?:\.|$)/.test(entry.name)
      ) {
        continue;
      }

      pendingDirectories.push(normalizeFsPath(path.join(directoryPath, entry.name)));
    }
  }
}

function scanLibraryDirectoryEntries(entries, libRootPath, sourceLabel) {
  const normalizedLibRootPath = normalizeFsPath(libRootPath);
  if (!fs.existsSync(normalizedLibRootPath)) {
    return;
  }

  let architectureEntries;
  try {
    architectureEntries = fs.readdirSync(normalizedLibRootPath, {
      withFileTypes: true,
    });
  } catch (error) {
    return;
  }

  for (const architectureEntry of architectureEntries) {
    if (!architectureEntry.isDirectory()) {
      continue;
    }

    const architecturePath = path.join(normalizedLibRootPath, architectureEntry.name);
    let fileEntries;
    try {
      fileEntries = fs.readdirSync(architecturePath, { withFileTypes: true });
    } catch (error) {
      continue;
    }

    for (const fileEntry of fileEntries) {
      if (!fileEntry.isFile() && !fileEntry.isSymbolicLink()) {
        continue;
      }

      const libraryName = extractLibraryName(fileEntry.name);
      if (!libraryName) {
        continue;
      }

      const entry = {
        name: libraryName,
        detail: `${sourceLabel} lib/${architectureEntry.name}/${fileEntry.name}`,
        documentation: `Found in ${sourceLabel}/lib/${architectureEntry.name}`,
        absolutePath: normalizeFsPath(path.join(architecturePath, fileEntry.name)),
        priority: getLibraryFilePriority(fileEntry.name),
      };

      const existingEntry = entries.get(libraryName);
      if (!existingEntry || entry.priority < existingEntry.priority) {
        entries.set(libraryName, entry);
      }
    }
  }
}

function hasLocalMakefileDbdReference(document, referenceName) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return false;
  }

  const normalizedReferenceName = normalizePath(referenceName);
  if (
    !normalizedReferenceName ||
    normalizedReferenceName.includes("$(") ||
    normalizedReferenceName.includes("${")
  ) {
    return false;
  }

  const makefileDirectory = path.dirname(document.uri.fsPath);
  const directCandidate = normalizeFsPath(
    path.resolve(makefileDirectory, normalizedReferenceName),
  );
  return isExistingFile(directCandidate);
}

function extractLibraryName(fileName) {
  const staticOrSharedMatch = fileName.match(
    /^lib(.+?)(?:\.[0-9][^.]*?)*\.(a|so|dylib|dll)$/i,
  );
  if (staticOrSharedMatch) {
    return staticOrSharedMatch[1];
  }

  const importLibraryMatch = fileName.match(/^(.+)\.lib$/i);
  if (importLibraryMatch) {
    return importLibraryMatch[1];
  }

  return undefined;
}

function getLibraryFilePriority(fileName) {
  if (/^lib.+(?:\.[0-9][^.]*?)+\.(so|dylib|dll)$/i.test(fileName)) {
    return 2;
  }

  if (/^lib[^.]+\.(so|dylib|dll)$/i.test(fileName)) {
    return 0;
  }

  if (/^lib.+\.a$/i.test(fileName) || /^.+\.lib$/i.test(fileName)) {
    return 1;
  }

  return 3;
}

function findProjectArtifactByFileName(project, load) {
  const targetFileName = path.posix.basename(normalizePath(load.path));
  if (!targetFileName) {
    return undefined;
  }

  for (const artifact of project.runtimeArtifacts) {
    if (!doesArtifactMatchFileKind(artifact, load.command)) {
      continue;
    }

    if (path.posix.basename(artifact.runtimeRelativePath) === targetFileName) {
      return artifact;
    }
  }

  return undefined;
}

function getLocalMakefileNavigationTarget(document, reference) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  if (!["dbFile", "sourceFile"].includes(reference.kind)) {
    return undefined;
  }

  const makefileDirectory = normalizeFsPath(path.dirname(document.uri.fsPath));
  const directCandidate = normalizeFsPath(path.resolve(makefileDirectory, reference.name));
  if (isExistingFile(directCandidate)) {
    return { absolutePath: directCandidate };
  }

  if (reference.kind === "dbFile") {
    const substitutionsRelativePath = toSubstitutionsSiblingPath(reference.name);
    const substitutionsCandidate = substitutionsRelativePath
      ? normalizeFsPath(path.resolve(makefileDirectory, substitutionsRelativePath))
      : undefined;
    if (isExistingFile(substitutionsCandidate)) {
      return { absolutePath: substitutionsCandidate };
    }
  }

  if (reference.kind === "sourceFile") {
    const generatedCandidate = findGeneratedSourceFile(makefileDirectory, reference.name);
    if (generatedCandidate) {
      return { absolutePath: generatedCandidate };
    }
  }

  return undefined;
}

function toSubstitutionsSiblingPath(filePath) {
  const parsedPath = path.posix.parse(String(filePath || "").replace(/\\/g, "/"));
  if (!DATABASE_EXTENSIONS.has(parsedPath.ext.toLowerCase())) {
    return undefined;
  }

  return path.posix.join(parsedPath.dir, `${parsedPath.name}.substitutions`);
}

function getMakefileDatabaseReferenceTarget(snapshot, document, reference) {
  if (!reference || reference.kind !== "dbFile") {
    return undefined;
  }

  const localTarget = getLocalMakefileNavigationTarget(document, reference);
  if (localTarget?.absolutePath) {
    return localTarget;
  }

  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  if (project) {
    const projectSourceCandidate = findProjectSourceCandidate(snapshot, project, {
      command: "dbLoadRecords",
      path: reference.name,
    });
    if (projectSourceCandidate?.absolutePath) {
      return { absolutePath: projectSourceCandidate.absolutePath };
    }
  }

  const workspaceSourceCandidate = findWorkspaceDatabaseFileCandidate(
    snapshot,
    reference.name,
    document,
  );
  if (workspaceSourceCandidate?.absolutePath) {
    return { absolutePath: workspaceSourceCandidate.absolutePath };
  }

  if (!project) {
    return undefined;
  }

  for (const artifact of project.runtimeArtifacts) {
    if (!["database", "substitutions"].includes(artifact.kind)) {
      continue;
    }

    if (artifact.runtimeFileName !== reference.name) {
      continue;
    }

    const sourcePath = getSourceAbsolutePathForArtifact(artifact);
    if (sourcePath) {
      return { absolutePath: sourcePath };
    }

    const readablePath = getReadableAbsolutePathForArtifact(artifact);
    if (readablePath) {
      return { absolutePath: readablePath };
    }
  }

  return undefined;
}

function findGeneratedSourceFile(makefileDirectory, fileName) {
  let directoryEntries;
  try {
    directoryEntries = fs.readdirSync(makefileDirectory, { withFileTypes: true });
  } catch (error) {
    return undefined;
  }

  const candidateDirectories = directoryEntries
    .filter(
      (entry) =>
        entry.isDirectory() &&
        (/^O\./.test(entry.name) || entry.name === "O.Common"),
    )
    .map((entry) => normalizeFsPath(path.join(makefileDirectory, entry.name, fileName)));

  for (const candidate of candidateDirectories) {
    if (isExistingFile(candidate)) {
      return candidate;
    }
  }

  return undefined;
}

function findWorkspaceDatabaseFileCandidate(snapshot, fileName, document) {
  const targetFileName = path.posix.basename(normalizePath(fileName));
  if (!targetFileName) {
    return undefined;
  }

  const workspaceFolder = document?.uri
    ? vscode.workspace.getWorkspaceFolder(document.uri)
    : undefined;
  const workspaceRootPath =
    workspaceFolder?.uri?.scheme === "file"
      ? normalizeFsPath(workspaceFolder.uri.fsPath)
      : undefined;
  const matches = snapshot.workspaceFiles
    .filter((entry) => {
      if (!entry.absolutePath || !DATABASE_EXTENSIONS.has(entry.extension)) {
        return false;
      }

      if (path.posix.basename(normalizePath(entry.absolutePath)) !== targetFileName) {
        return false;
      }

      return !workspaceRootPath || isPathWithinRoot(entry.absolutePath, workspaceRootPath);
    })
    .sort((left, right) => compareLabels(left.relativePath, right.relativePath));

  return matches[0];
}

function extractMakefileReferences(text) {
  const references = [];
  let offset = 0;

  for (const line of text.split(/\r?\n/)) {
    const assignmentMatch = line.match(
      /^\s*([A-Za-z0-9_.-]+)\s*(?:\+?=|:=|\?=)\s*/,
    );
    if (!assignmentMatch) {
      offset += line.length + 1;
      continue;
    }

    const valueStartInLine = assignmentMatch[0].length;
    const commentIndex = line.indexOf("#", valueStartInLine);
    const valueSection =
      commentIndex >= 0
        ? line.slice(valueStartInLine, commentIndex)
        : line.slice(valueStartInLine);
    const tokenRegex = /[^\s]+/g;
    let tokenMatch;

    while ((tokenMatch = tokenRegex.exec(valueSection))) {
      const token = tokenMatch[0];
      const kind = getMakefileReferenceKind(assignmentMatch[1], token);
      if (!kind) {
        continue;
      }

      references.push({
        kind,
        name: token,
        variableName: assignmentMatch[1],
        start: offset + valueStartInLine + tokenMatch.index,
        end: offset + valueStartInLine + tokenMatch.index + token.length,
      });
    }

    offset += line.length + 1;
  }

  return references;
}

function getMakefileReferenceAtPosition(document, position) {
  const offset = document.offsetAt(position);

  for (const reference of extractMakefileReferences(document.getText())) {
    if (offset >= reference.start && offset <= reference.end) {
      return reference;
    }
  }

  return undefined;
}

function isConcreteMakefileReferenceToken(token, kind) {
  if (!token || containsMakeVariableReference(token) || token === "-nil-") {
    return false;
  }

  if (kind === "dbd") {
    return token.toLowerCase().endsWith(".dbd");
  }

  if (kind === "dbFile") {
    return /\.(db|template|sub|subs|substitutions)$/i.test(token);
  }

  if (kind === "sourceFile") {
    return /\.(c|cc|cpp|cxx|cp|C|h|hh|hpp|hxx)$/i.test(token);
  }

  return !token.startsWith("-");
}

function getMakefileReferenceKind(variableName, token) {
  if (containsMakeVariableReference(token) || token === "-nil-") {
    return undefined;
  }

  if (/^(?:[A-Za-z0-9_.-]+_)?DBD$/.test(variableName)) {
    return "dbd";
  }

  if (/^(?:[A-Za-z0-9_.-]+_)?LIBS$/.test(variableName)) {
    return "lib";
  }

  if (/^(?:[A-Za-z0-9_.-]+_)?DB$/.test(variableName)) {
    return "dbFile";
  }

  if (
    /^(?:[A-Za-z0-9_.-]+_)?SRCS(?:_[A-Za-z0-9_.-]+)?$/.test(variableName)
  ) {
    return "sourceFile";
  }

  return undefined;
}

function isProjectResourceMakefileReference(reference) {
  return reference && ["dbd", "lib"].includes(reference.kind);
}

function containsMakeVariableReference(value) {
  return value.includes("$(") || value.includes("${");
}

function findProjectSourceCandidate(snapshot, project, load) {
  if (!["dbLoadRecords", "dbLoadTemplate"].includes(load.command)) {
    return undefined;
  }

  const targetFileName = path.posix.basename(normalizePath(load.path));
  const allowedExtensions = getAllowedExtensionsForFileContext(load.command);
  if (!targetFileName) {
    return undefined;
  }

  for (const entry of snapshot.workspaceFiles) {
    if (!entry.absolutePath || !allowedExtensions.has(entry.extension)) {
      continue;
    }

    if (!isPathWithinRoot(entry.absolutePath, project.rootPath)) {
      continue;
    }

    if (path.posix.basename(normalizePath(entry.absolutePath)) !== targetFileName) {
      continue;
    }

    if (!/[\\/](?:Db|db)[\\/]/.test(entry.absolutePath)) {
      continue;
    }

    return entry;
  }

  return undefined;
}

function findProjectDbdHint(project, load) {
  if (load.command !== "dbLoadDatabase") {
    return undefined;
  }

  const fileName = path.posix.basename(normalizePath(load.path));
  if (!fileName.toLowerCase().endsWith(".dbd")) {
    return undefined;
  }

  const iocName = fileName.slice(0, -".dbd".length);
  const iocInfo = project.iocsByName.get(iocName);
  if (!iocInfo) {
    return undefined;
  }

  return `Startup path "${load.path}" does not resolve to a generated DBD in this application. Check DBD += in ${iocInfo.makefileRelativePath}.`;
}

function extractTemplateMappings(assignments) {
  const mappings = new Map();

  for (const [variableName, values] of assignments.entries()) {
    if (!variableName.endsWith("_template") || values.length === 0) {
      continue;
    }

    mappings.set(variableName.slice(0, -"_template".length), values[0]);
  }

  return mappings;
}

function resolveDatabaseSourceFileName(installedDbName, templateMappings) {
  const extension = path.extname(installedDbName).toLowerCase();
  const baseName = extension
    ? installedDbName.slice(0, -extension.length)
    : installedDbName;

  if (templateMappings.has(baseName)) {
    return templateMappings.get(baseName);
  }

  return installedDbName;
}

function getRuntimeArtifactKind(fileName) {
  const extension = path.extname(fileName).toLowerCase();

  if (DBD_EXTENSIONS.has(extension)) {
    return "dbd";
  }

  if (SUBSTITUTION_EXTENSIONS.has(extension)) {
    return "substitutions";
  }

  if (DATABASE_EXTENSIONS.has(extension)) {
    return "database";
  }

  return undefined;
}

function getSourceAbsolutePathForArtifact(artifact) {
  if (!artifact?.sourceRelativePath) {
    return undefined;
  }

  const sourcePath = normalizeFsPath(
    path.join(findRootPathForArtifact(artifact), artifact.sourceRelativePath),
  );
  return isExistingFile(sourcePath) ? sourcePath : undefined;
}

function createRuntimeArtifact({
  rootPath,
  appDirName,
  kind,
  runtimeRelativePath,
  sourceRelativePath,
  detail,
  documentation,
}) {
  return {
    appDirName,
    kind,
    runtimeRelativePath,
    runtimeFileName: path.posix.basename(runtimeRelativePath),
    absoluteRuntimePath: normalizeFsPath(path.join(rootPath, runtimeRelativePath)),
    sourceRelativePath,
    detail,
    documentation,
  };
}

function createWorkspaceFileEntry(uri) {
  const relativePath = normalizePath(vscode.workspace.asRelativePath(uri, false));
  return {
    uri,
    absolutePath: uri.scheme === "file" ? normalizeFsPath(uri.fsPath) : undefined,
    extension: path.extname(uri.fsPath).toLowerCase(),
    relativePath,
  };
}

function hasExtension(uri, extensions) {
  if (!uri || uri.scheme !== "file") {
    return false;
  }

  return extensions.has(path.extname(uri.fsPath).toLowerCase());
}

async function readWorkspaceFile(uri) {
  const openDocument = vscode.workspace.textDocuments.find(
    (document) => document.uri.toString() === uri.toString(),
  );
  if (openDocument) {
    return openDocument.getText();
  }

  try {
    const bytes = await vscode.workspace.fs.readFile(uri);
    return Buffer.from(bytes).toString("utf8");
  } catch (error) {
    return undefined;
  }
}

function readTextFile(filePath) {
  const normalizedPath = normalizeFsPath(filePath);
  const openDocument = vscode.workspace.textDocuments.find(
    (document) =>
      document.uri.scheme === "file" &&
      normalizeFsPath(document.uri.fsPath) === normalizedPath,
  );
  if (openDocument) {
    return openDocument.getText();
  }

  try {
    return fs.readFileSync(normalizedPath, "utf8");
  } catch (error) {
    return undefined;
  }
}

function readJsonFile(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function registerWatcher(disposables, glob, onChange) {
  const watcher = vscode.workspace.createFileSystemWatcher(glob);
  watcher.onDidCreate(onChange, undefined, disposables);
  watcher.onDidChange(onChange, undefined, disposables);
  watcher.onDidDelete(onChange, undefined, disposables);
  disposables.push(watcher);
}

async function collectWorkspaceUris() {
  const uriMap = new Map();

  for (const glob of [INDEX_GLOB, SOURCE_INDEX_GLOB, ...PROJECT_INDEX_GLOBS]) {
    const uris = await vscode.workspace.findFiles(glob, INDEX_EXCLUDE_GLOB);
    for (const uri of uris) {
      uriMap.set(uri.toString(), uri);
    }
  }

  return [...uriMap.values()];
}

function buildWorkspaceFileLookup(workspaceFiles) {
  const lookup = new Map();

  for (const entry of workspaceFiles) {
    if (entry.absolutePath) {
      lookup.set(entry.absolutePath, entry);
    }
  }

  return lookup;
}

function isIndexedContentFile(uri) {
  return getEpicsFileExtension(uri) !== undefined;
}

function isIndexedContentDocument(document) {
  return document && isIndexedContentFile(document.uri);
}

function isProjectModelUri(uri) {
  if (!uri || uri.scheme !== "file") {
    return false;
  }

  const baseName = path.basename(uri.fsPath);
  if (baseName === "Makefile") {
    return true;
  }

  return (
    baseName === "RELEASE" &&
    path.basename(path.dirname(uri.fsPath)) === "configure"
  ) || (
    baseName === "RELEASE.local" &&
    path.basename(path.dirname(uri.fsPath)) === "configure"
  );
}

function isProjectModelDocument(document) {
  return document && isProjectModelUri(document.uri);
}

function getEpicsLanguageIdForUri(uri) {
  if (!uri || uri.scheme !== "file") {
    return undefined;
  }

  if (hasExtension(uri, DATABASE_EXTENSIONS)) {
    return LANGUAGE_IDS.database;
  }

  if (hasExtension(uri, SUBSTITUTION_EXTENSIONS)) {
    return LANGUAGE_IDS.substitutions;
  }

  if (hasExtension(uri, STARTUP_EXTENSIONS)) {
    return LANGUAGE_IDS.startup;
  }

  if (hasExtension(uri, DBD_EXTENSIONS)) {
    return LANGUAGE_IDS.dbd;
  }

  if (hasExtension(uri, PVLIST_EXTENSIONS)) {
    return LANGUAGE_IDS.pvlist;
  }

  if (hasExtension(uri, PROBE_EXTENSIONS)) {
    return LANGUAGE_IDS.probe;
  }

  if (hasExtension(uri, PROTOCOL_EXTENSIONS)) {
    return LANGUAGE_IDS.proto;
  }

  if (hasExtension(uri, SOURCE_EXTENSIONS)) {
    return LANGUAGE_IDS.source;
  }

  return undefined;
}

function getEpicsFileExtension(uri) {
  if (!uri || uri.scheme !== "file") {
    return undefined;
  }

  const extension = path.extname(uri.fsPath).toLowerCase();
  if (
    DATABASE_EXTENSIONS.has(extension) ||
    SUBSTITUTION_EXTENSIONS.has(extension) ||
    STARTUP_EXTENSIONS.has(extension) ||
    DBD_EXTENSIONS.has(extension) ||
    PVLIST_EXTENSIONS.has(extension) ||
    PROBE_EXTENSIONS.has(extension) ||
    PROTOCOL_EXTENSIONS.has(extension) ||
    SOURCE_EXTENSIONS.has(extension)
  ) {
    return extension;
  }

  return undefined;
}

function addToMapOfSets(map, key, values) {
  if (!map.has(key)) {
    map.set(key, new Set());
  }

  const valueSet = map.get(key);
  for (const value of values) {
    if (value) {
      valueSet.add(value);
    }
  }
}

function addToMapOfArrays(map, key, value) {
  if (!map.has(key)) {
    map.set(key, []);
  }

  map.get(key).push(value);
}

function addToMapOfMaps(map, key, values) {
  if (!map.has(key)) {
    map.set(key, new Map());
  }

  const valueMap = map.get(key);
  for (const [nestedKey, nestedValue] of values.entries()) {
    if (!valueMap.has(nestedKey)) {
      valueMap.set(nestedKey, nestedValue);
    }
  }
}

function cloneMapOfSets(map) {
  const clone = new Map();

  for (const [key, values] of map.entries()) {
    clone.set(key, new Set(values));
  }

  return clone;
}

function cloneMapOfMaps(map) {
  const clone = new Map();

  for (const [key, values] of map.entries()) {
    clone.set(
      key,
      new Map(
        [...values.entries()].map(([nestedKey, nestedValue]) => [
          nestedKey,
          Array.isArray(nestedValue) ? [...nestedValue] : nestedValue,
        ]),
      ),
    );
  }

  return clone;
}

function cloneMapOfArrays(map) {
  const clone = new Map();

  for (const [key, values] of map.entries()) {
    clone.set(key, [...values]);
  }

  return clone;
}

function createFieldTypeMap(source) {
  const fieldTypeMap = new Map();

  for (const [recordType, fieldTypes] of Object.entries(source || {})) {
    addToMapOfMaps(
      fieldTypeMap,
      recordType,
      new Map(Object.entries(fieldTypes || {})),
    );
  }

  return fieldTypeMap;
}

function createFieldMenuChoiceMap(source) {
  const fieldMenuChoiceMap = new Map();

  for (const [recordType, fieldMenus] of Object.entries(source || {})) {
    addToMapOfMaps(
      fieldMenuChoiceMap,
      recordType,
      new Map(
        Object.entries(fieldMenus || {}).map(([fieldName, choices]) => [
          fieldName,
          Array.isArray(choices) ? [...choices] : [],
        ]),
      ),
    );
  }

  return fieldMenuChoiceMap;
}

function createFieldInitialValueMap(source) {
  const fieldInitialValueMap = new Map();

  for (const [recordType, fieldInitialValues] of Object.entries(source || {})) {
    addToMapOfMaps(
      fieldInitialValueMap,
      recordType,
      new Map(Object.entries(fieldInitialValues || {})),
    );
  }

  return fieldInitialValueMap;
}

function splitSymbolList(value) {
  return value
    .split(/[\s,]+/)
    .map((symbol) => symbol.trim())
    .filter(Boolean);
}

function splitMakeValue(value) {
  return value
    .split(/\s+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function stripMakeComments(value) {
  const hashIndex = value.indexOf("#");
  return hashIndex >= 0 ? value.slice(0, hashIndex) : value;
}

function countCharacters(text, character) {
  let count = 0;

  for (const current of text) {
    if (current === character) {
      count += 1;
    }
  }

  return count;
}

function getLineNumberAtOffset(text, offset) {
  if (offset <= 0) {
    return 1;
  }

  let lineNumber = 1;
  for (let index = 0; index < offset && index < text.length; index += 1) {
    if (text[index] === "\n") {
      lineNumber += 1;
    }
  }

  return lineNumber;
}


function findRecordBlockEnd(text, recordStart) {
  const openingBraceIndex = text.indexOf("{", recordStart);
  if (openingBraceIndex < 0) {
    return text.length;
  }

  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = openingBraceIndex; index < text.length; index += 1) {
    const character = text[index];

    if (escaped) {
      escaped = false;
      continue;
    }

    if (inString) {
      if (character === "\\") {
        escaped = true;
      } else if (character === "\"") {
        inString = false;
      }
      continue;
    }

    if (character === "\"") {
      inString = true;
      continue;
    }

    if (character === "{") {
      depth += 1;
      continue;
    }

    if (character === "}") {
      depth -= 1;
      if (depth === 0) {
        return index + 1;
      }
    }
  }

  return text.length;
}

function buildRecordPreview(text, declaration) {
  const preview = text
    .slice(declaration.recordStart, declaration.recordEnd)
    .trim();
  if (!preview) {
    return `record(${declaration.recordType}, "${declaration.name}")`;
  }

  const previewLines = preview.split(/\r?\n/);
  const truncatedLines =
    previewLines.length > RECORD_PREVIEW_MAX_LINES
      ? [...previewLines.slice(0, RECORD_PREVIEW_MAX_LINES), "..."]
      : previewLines;
  const truncatedPreview = truncatedLines.join("\n");

  return truncatedPreview.length > RECORD_PREVIEW_MAX_CHARACTERS
    ? `${truncatedPreview.slice(0, RECORD_PREVIEW_MAX_CHARACTERS - 3)}...`
    : truncatedPreview;
}

function escapeInlineCode(value) {
  return String(value).replace(/`/g, "\\`");
}

function escapeDoubleQuotedString(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function escapeMarkdownLinkLabel(value) {
  return String(value).replace(/([\\\[\]\(\)])/g, "\\$1");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function isSkippableNumericFieldValue(value) {
  const trimmedValue = String(value || "").trim();
  return !trimmedValue || containsEpicsMacroReference(trimmedValue);
}

function containsEpicsMacroReference(value) {
  return /\$\(|\$\{|\$[A-Za-z_]/.test(String(value || ""));
}

function isValidNumericFieldValue(value, dbfType) {
  const numericValue = Number(String(value).trim());
  if (!Number.isFinite(numericValue)) {
    return false;
  }

  if (INTEGER_DBF_TYPES.has(dbfType) && !Number.isInteger(numericValue)) {
    return false;
  }

  return true;
}

function getEscapedStringLength(value) {
  let length = 0;
  let escaped = false;

  for (const character of value) {
    if (escaped) {
      length += 1;
      escaped = false;
      continue;
    }

    if (character === "\\") {
      escaped = true;
      continue;
    }

    length += 1;
  }

  if (escaped) {
    length += 1;
  }

  return length;
}

function compareLabels(left, right) {
  return String(left).localeCompare(String(right));
}

function findBestMatchingLabel(labels, target) {
  const normalizedTarget = String(target || "").trim();
  if (!normalizedTarget) {
    return undefined;
  }

  const uniqueLabels = [...new Set((labels || []).filter(Boolean))];
  if (uniqueLabels.length === 0) {
    return undefined;
  }

  return uniqueLabels
    .map((label) => ({
      label,
      distance: computeLevenshteinDistance(
        String(label).toUpperCase(),
        normalizedTarget.toUpperCase(),
      ),
    }))
    .sort(
      (left, right) =>
        left.distance - right.distance ||
        compareLabels(left.label, right.label),
    )[0]?.label;
}

function computeLevenshteinDistance(left, right) {
  const source = String(left || "");
  const target = String(right || "");
  const matrix = Array.from({ length: source.length + 1 }, () =>
    new Array(target.length + 1).fill(0),
  );

  for (let row = 0; row <= source.length; row += 1) {
    matrix[row][0] = row;
  }
  for (let column = 0; column <= target.length; column += 1) {
    matrix[0][column] = column;
  }

  for (let row = 1; row <= source.length; row += 1) {
    for (let column = 1; column <= target.length; column += 1) {
      const substitutionCost = source[row - 1] === target[column - 1] ? 0 : 1;
      matrix[row][column] = Math.min(
        matrix[row - 1][column] + 1,
        matrix[row][column - 1] + 1,
        matrix[row - 1][column - 1] + substitutionCost,
      );
    }
  }

  return matrix[source.length][target.length];
}

function createDiagnostic(start, end, message, severity = vscode.DiagnosticSeverity.Error) {
  const diagnostic = new vscode.Diagnostic(
    new vscode.Range(start, end),
    message,
    severity,
  );
  diagnostic.source = "vscode-epics";
  return diagnostic;
}

function isDatabaseDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.database;
}

function isSequencerDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.sequencer;
}

function isStartupDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.startup;
}

function isSubstitutionsDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.substitutions;
}

function isProtocolDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.proto;
}

function isSourceDocument(document) {
  return document && hasExtension(document.uri, SOURCE_EXTENSIONS);
}

function isPvlistDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.pvlist;
}

function isProbeDocument(document) {
  return document && document.languageId === LANGUAGE_IDS.probe;
}

function getDocumentDisplayLabel(document) {
  if (!document?.uri) {
    return "EPICS";
  }

  return (
    path.basename(document.uri.fsPath || document.uri.path || document.fileName || "") ||
    document.fileName ||
    "EPICS"
  );
}

function isStartupStateDocument(document) {
  return (
    isStartupDocument(document) ||
    (
      document &&
      document.uri &&
      document.uri.scheme === "file" &&
      /^envPaths(?:\..+)?$/.test(path.basename(document.uri.fsPath))
    )
  );
}

function isMakefileDocument(document) {
  return (
    document &&
    document.uri &&
    document.uri.scheme === "file" &&
    path.basename(document.uri.fsPath) === "Makefile"
  );
}

function isSourceMakefileDocument(document) {
  return (
    isMakefileDocument(document) &&
    /[\\/][^\\/]+App[\\/]src[\\/]Makefile$/.test(document.uri.fsPath)
  );
}

function matchesCompletionQuery(candidate, partial) {
  if (!partial) {
    return true;
  }

  const candidateLower = String(candidate).toLowerCase();
  const partialLower = String(partial).toLowerCase();

  if (candidateLower.includes(partialLower)) {
    return true;
  }

  const normalizedCandidate = normalizeCompletionText(candidateLower);
  const normalizedPartial = normalizeCompletionText(partialLower);

  if (!normalizedPartial) {
    return true;
  }

  if (normalizedCandidate.includes(normalizedPartial)) {
    return true;
  }

  return isSubsequence(normalizedPartial, normalizedCandidate);
}

function buildFilterText(label) {
  const normalized = normalizeCompletionText(label);
  return normalized ? `${label} ${normalized}` : String(label);
}

function buildSortText(label, partial) {
  if (!partial) {
    return `1-${label}`;
  }

  const normalizedLabel = normalizeCompletionText(label);
  const normalizedPartial = normalizeCompletionText(partial);

  if (
    label.toLowerCase().startsWith(String(partial).toLowerCase()) ||
    (normalizedPartial && normalizedLabel.startsWith(normalizedPartial))
  ) {
    return `0-${label}`;
  }

  return `1-${label}`;
}

function inferDeviceChoiceName(supportName) {
  const text = String(supportName || "").trim();
  if (!text) {
    return undefined;
  }

  const stripped = text.replace(/^(?:dev|Dev|DSET_)/, "").replace(/^_+/, "");
  return stripped || text;
}

function normalizeCompletionText(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9:_-]/g, "");
}

function isSubsequence(needle, haystack) {
  let haystackIndex = 0;

  for (let needleIndex = 0; needleIndex < needle.length; needleIndex += 1) {
    const character = needle[needleIndex];
    haystackIndex = haystack.indexOf(character, haystackIndex);

    if (haystackIndex === -1) {
      return false;
    }

    haystackIndex += 1;
  }

  return true;
}

function normalizePath(value) {
  return value.split(path.sep).join(path.posix.sep);
}

function normalizeFsPath(value) {
  return normalizePath(path.resolve(value));
}

function getDocumentDirectoryPath(document) {
  if (!document?.uri || document.uri.scheme !== "file") {
    return undefined;
  }

  return normalizeFsPath(path.dirname(document.uri.fsPath));
}

function isDefinitionInDirectory(definition, directoryPath) {
  if (!directoryPath) {
    return true;
  }

  return (
    definition &&
    definition.absolutePath &&
    isPathWithinRoot(definition.absolutePath, directoryPath)
  );
}

function isPathWithinRoot(filePath, rootPath) {
  const normalizedFilePath = normalizeFsPath(filePath);
  const normalizedRootPath = normalizeFsPath(rootPath);
  return (
    normalizedFilePath === normalizedRootPath ||
    normalizedFilePath.startsWith(`${normalizedRootPath}/`)
  );
}

function isExistingFile(filePath) {
  if (!filePath || !fs.existsSync(filePath)) {
    return false;
  }

  try {
    return fs.statSync(filePath).isFile();
  } catch (error) {
    return false;
  }
}

function getRelativePathFromDocument(document, targetPath) {
  if (!document || !document.uri || document.uri.scheme !== "file") {
    return normalizePath(targetPath);
  }

  return normalizePath(
    path.relative(path.dirname(document.uri.fsPath), targetPath),
  ) || path.posix.basename(targetPath);
}

function getRelativePathFromBaseDirectory(baseDirectory, targetPath) {
  return normalizePath(path.relative(baseDirectory, targetPath)) || path.posix.basename(targetPath);
}

module.exports = {
  activate,
  deactivate,
};
