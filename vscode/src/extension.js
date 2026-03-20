const fs = require("fs");
const path = require("path");
const vscode = require("vscode");
const {
  formatDatabaseText,
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
const { registerRuntimeMonitor } = require("./runtimeMonitor");

const LANGUAGE_IDS = {
  database: "database",
  startup: "startup",
  substitutions: "substitutions",
  dbd: "database definition",
  sequencer: "sequencer",
};

const DATABASE_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const SUBSTITUTION_EXTENSIONS = new Set([".sub", ".subs", ".substitutions"]);
const STARTUP_EXTENSIONS = new Set([".cmd", ".iocsh"]);
const DBD_EXTENSIONS = new Set([".dbd"]);
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
const INDEX_GLOB = "**/*.{db,vdb,template,sub,subs,substitutions,cmd,iocsh,dbd}";
const SOURCE_INDEX_GLOB = "**/*.{c,cc,cpp,cxx,h,hh,hpp,hxx}";
const PROJECT_INDEX_GLOBS = [
  "**/Makefile",
  "**/configure/RELEASE",
  "**/configure/RELEASE.local",
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
const COPY_AS_MONITOR_FILE_COMMAND = "vscode-epics.copyAsMonitorFile";
const COPY_AS_EXPANDED_DB_COMMAND = "vscode-epics.copyAsExpandedDb";
const UPDATE_MENU_FIELD_VALUE_COMMAND = "vscode-epics.updateMenuFieldValue";
const STREAM_PROTOCOL_PATH_VARIABLE = "STREAM_PROTOCOL_PATH";
const DBD_DEVICE_LINK_TYPES = ["INST_IO"];
const TRIGGER_SUGGEST_COMMAND = "editor.action.triggerSuggest";
const RECORD_PREVIEW_MAX_LINES = 100;
const RECORD_PREVIEW_MAX_CHARACTERS = 12000;
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

function activate(context) {
  registerRuntimeMonitor(context, {
    extractDatabaseTocEntries,
    extractDatabaseTocMacroAssignments,
    extractRecordDeclarations,
  });
  const staticData = loadStaticData(context.extensionPath);
  recordTemplateFields = new Map(staticData.recordTemplateFields || []);
  recordTemplateStaticData = {
    fieldTypesByRecordType: staticData.fieldTypesByRecordType,
    fieldMenuChoicesByRecordType: staticData.fieldMenuChoicesByRecordType,
    fieldInitialValuesByRecordType: staticData.fieldInitialValuesByRecordType,
  };
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
      { language: LANGUAGE_IDS.sequencer },
      new EpicsReferenceProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.languages.registerHoverProvider(
      [
        { language: LANGUAGE_IDS.database },
        { language: LANGUAGE_IDS.startup },
        { language: LANGUAGE_IDS.substitutions },
        { language: LANGUAGE_IDS.sequencer },
        { language: "makefile" },
        { scheme: "file", pattern: "**/Makefile" },
        { scheme: "file", pattern: "**/envPaths*" },
      ],
      new EpicsHoverProvider(workspaceIndex),
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
      { language: LANGUAGE_IDS.sequencer },
      new EpicsSequencerFormattingProvider(),
    ),
  );
  context.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((document) => {
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
  provideReferences(document, position, context) {
    if (!isSequencerDocument(document)) {
      return [];
    }

    return getSequencerReferenceLocations(
      document,
      position,
      Boolean(context?.includeDeclaration),
    );
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

    snapshot.projectModel = buildProjectModel(projectFiles);
    snapshot.workspaceFilesByAbsolutePath = buildWorkspaceFileLookup(snapshot.workspaceFiles);
    snapshot.workspaceFiles.sort((left, right) =>
      compareLabels(left.relativePath, right.relativePath),
    );

    this.snapshot = snapshot;
    this.dirty = false;
    this.rebuildPromise = undefined;
  }

  dispose() {
    this.disposables.forEach((disposable) => disposable.dispose());
  }
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
      path: context.path,
    },
    state,
  );
  const resolvedFile = resolution
    ? getReadableStartupFileResolution(document, resolution)
    : undefined;
  if (!resolvedFile?.text) {
    return [];
  }

  const macroNames = extractMacroNames(maskDatabaseComments(resolvedFile.text));
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

  const linkedRecordName = resolveLinkedRecordName(
    snapshot,
    document,
    fieldDeclaration.value,
  );
  if (!linkedRecordName) {
    return undefined;
  }

  const definitions = getRecordDefinitionsForName(
    snapshot,
    document,
    linkedRecordName,
  );
  if (definitions.length === 0) {
    return undefined;
  }

  return createLinkedRecordHover(
    linkedRecordName,
    fieldDeclaration.fieldName,
    definitions,
    fieldDeclaration.range,
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

  const absolutePath = resolveSubstitutionTemplateAbsolutePathForDocument(
    snapshot,
    document,
    reference.templatePath,
  );
  if (!absolutePath) {
    return undefined;
  }

  return createSubstitutionTemplateHover(document, reference, absolutePath);
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

function createSubstitutionTemplateHover(document, reference, absolutePath) {
  const templateText = readTextFile(absolutePath);
  const markdown = new vscode.MarkdownString();
  markdown.isTrusted = true;
  markdown.appendMarkdown("**EPICS database/template file**");

  if (templateText === undefined) {
    markdown.appendMarkdown(
      `\n\nPath: ${createProtocolFileLink(absolutePath, absolutePath)}`,
    );
  } else {
    appendDatabaseFileHoverSummary(markdown, absolutePath, templateText);
  }

  return new vscode.Hover(
    markdown,
    new vscode.Range(
      document.positionAt(reference.start),
      document.positionAt(reference.end),
    ),
  );
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
    const rawValue =
      state.envVariables.get(reference.variableName) ||
      process.env[reference.variableName];
    if (!rawValue) {
      return undefined;
    }

    const resolvedValue = expandStartupValue(rawValue, state.envVariables);

    const absolutePath = computeAbsoluteVariablePath(
      resolvedValue,
      state.currentDirectory || path.dirname(document.uri.fsPath),
    );
    return createVariableHover(reference, resolvedValue, absolutePath);
  }

  return undefined;
}

function createVariableHover(reference, resolvedValue, absolutePath) {
  const lines = [
    `**${reference.variableName}**`,
    "",
    `Resolved value: \`${resolvedValue}\``,
  ];

  if (absolutePath && absolutePath !== resolvedValue) {
    lines.push("", `Absolute path: \`${absolutePath}\``);
  }

  return new vscode.Hover(lines.join("\n"), reference.range);
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
  const commandUri = buildOpenRecordCommandUri(
    definition.absolutePath,
    definition.line,
  );
  return `[${escapeMarkdownLinkLabel(label)}](${commandUri})`;
}

function createProtocolFileLink(absolutePath, label) {
  const commandUri = buildOpenRecordCommandUri(absolutePath, 1);
  return `[${escapeMarkdownLinkLabel(label)}](${commandUri})`;
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

  for (const startupFilePath of findProjectIocBootFilePaths(project.rootPath)) {
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
  for (const candidate of extractLinkedRecordCandidates(fieldValue)) {
    if (getRecordDefinitionsForName(snapshot, document, candidate).length > 0) {
      return candidate;
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
  const currentDefinitions = createRecordDefinitions(
    document.uri,
    document.getText(),
    extractRecordDeclarations(document.getText()).filter(
      (declaration) => declaration.name === recordName,
    ),
  );

  for (const definition of currentDefinitions) {
    const key = `${definition.absolutePath}:${definition.line}:${definition.name}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    definitions.push(definition);
  }

  for (const definition of snapshot.recordDefinitionsByName.get(recordName) || []) {
    if (currentPath && definition.absolutePath === currentPath) {
      continue;
    }

    const key = `${definition.absolutePath}:${definition.line}:${definition.name}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    definitions.push(definition);
  }

  return definitions;
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
  if (!document?.uri || document.uri.scheme !== "file" || !templatePath) {
    return undefined;
  }

  const project = findProjectForUri(snapshot.projectModel, document.uri);
  const releaseVariables = project ? project.releaseVariables : new Map();
  const expandedTemplatePath = expandEpicsValue(templatePath, [
    releaseVariables,
    process.env,
  ]);
  if (!expandedTemplatePath) {
    return undefined;
  }

  const localPath = normalizeFsPath(
    path.resolve(path.dirname(document.uri.fsPath), expandedTemplatePath),
  );
  if (readTextFile(localPath) !== undefined) {
    return localPath;
  }

  for (const releaseRoot of getProjectReleaseSearchRoots(project)) {
    for (const candidatePath of getReleaseTemplateCandidatePaths(
      releaseRoot,
      expandedTemplatePath,
    )) {
      if (readTextFile(candidatePath) !== undefined) {
        return candidatePath;
      }
    }
  }

  return undefined;
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
    addToMapOfArrays(
      snapshot.recordDefinitionsByName,
      recordDefinition.name,
      recordDefinition,
    );
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
    iocsByName: new Map(),
    releaseVariables: new Map(),
    availableDbds: new Map(),
    availableLibs: new Map(),
  };
}

function buildProjectModel(projectFiles) {
  const projectModel = createEmptyProjectModel();
  const filesByAbsolutePath = new Map();
  const rootPaths = new Set();

  for (const file of projectFiles) {
    filesByAbsolutePath.set(normalizeFsPath(file.uri.fsPath), file);

    if (
      path.basename(file.uri.fsPath) === "RELEASE" &&
      path.basename(path.dirname(file.uri.fsPath)) === "configure"
    ) {
      rootPaths.add(normalizeFsPath(path.dirname(path.dirname(file.uri.fsPath))));
    }
  }

  for (const rootPath of [...rootPaths].sort(compareLabels)) {
    const application = buildProjectApplication(rootPath, filesByAbsolutePath);
    if (!application) {
      continue;
    }

    projectModel.applications.push(application);

    for (const artifact of application.runtimeArtifacts) {
      projectModel.runtimeArtifacts.push(artifact);
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

function buildProjectApplication(rootPath, filesByAbsolutePath) {
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
        createDiagnostic(
          document.positionAt(fieldDeclaration.fieldNameStart),
          document.positionAt(fieldDeclaration.fieldNameEnd),
          `Field "${fieldDeclaration.fieldName}" is not valid for record type "${recordDeclaration.recordType}".`,
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
        createDiagnostic(
          document.positionAt(fieldDeclaration.valueStart),
          document.positionAt(fieldDeclaration.valueEnd),
          `Field "${fieldDeclaration.fieldName}" must be one of the menu choices for "${recordDeclaration.recordType}".`,
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
          createDiagnostic(
            document.positionAt(statement.nameStart),
            document.positionAt(statement.nameEnd),
            `Unknown IOC registration function "${statement.functionName}" for this EPICS application.`,
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

  const providedMacroNames = new Set(extractNamedAssignments(statement.macros).keys());
  const missingMacroNames = requiredMacroNames.filter(
    (macroName) => !providedMacroNames.has(macroName),
  );
  if (missingMacroNames.length === 0) {
    return [];
  }

  return [
    createDiagnostic(
      document.positionAt(statement.pathStart),
      document.positionAt(statement.pathEnd),
      `dbLoadRecords is missing macro assignments for "${path.posix.basename(
        normalizePath(statement.path),
      )}": ${missingMacroNames.join(", ")}.`,
    ),
  ];
}

function createSubstitutionDiagnostics(document, snapshot) {
  if (!isSubstitutionsDocument(document)) {
    return [];
  }

  const diagnostics = [];
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

    const message =
      reference.kind === "dbd"
        ? `Unknown DBD "${reference.name}". It was not found in this project's dbd outputs or the module roots from RELEASE.`
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
    `Copied ${recordLabel} and ${macroLabel} as a .monitor file.`,
  );
}

function buildMonitorFileText(recordNames, macroNames, eol = "\n") {
  const lines = [
    "# this is a monitor file for EPICS Workbench in VSCode",
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

function extractDbdFieldsByRecordType(text) {
  const fieldsByRecordType = new Map();
  const recordTypeRegex = /recordtype\(\s*([A-Za-z0-9_]+)\s*\)\s*\{/g;
  let match;

  while ((match = recordTypeRegex.exec(text))) {
    const block = readBalancedBlock(text, recordTypeRegex.lastIndex - 1);
    if (!block) {
      continue;
    }

    const recordType = match[1];
    const fieldRegex = /field\(\s*([A-Z0-9_]+)\s*,/g;
    let fieldMatch;

    while ((fieldMatch = fieldRegex.exec(block.body))) {
      addToMapOfSets(fieldsByRecordType, recordType, [fieldMatch[1]]);
    }

    recordTypeRegex.lastIndex = block.endIndex;
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

function createInitialStartupExecutionState(snapshot, document) {
  const project = findProjectForUri(snapshot.projectModel, document.uri);
  return {
    envVariables: new Map(project ? project.releaseVariables : []),
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
      state.envVariables.set(
        statement.name,
        expandStartupValue(statement.value, state.envVariables),
      );
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
        state.envVariables.set(
          statement.name,
          expandStartupValue(statement.value, state.envVariables),
        );
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

  for (const line of text.split(/\r?\n/)) {
    const lineOffset = offset;
    offset += line.length + 1;

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
      });
      continue;
    }

    match = line.match(
      /^\s*epicsEnvSet\(\s*"?([A-Za-z_][A-Za-z0-9_]*)"?\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)/,
    );
    if (match) {
      statements.push({
        kind: "envSet",
        name: match[1],
        value: match[2],
        start: lineOffset,
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
      });
      continue;
    }

    match = line.match(
      /^\s*dbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"((?:[^"\\]|\\.)*)")?/,
    );
    if (match) {
      const pathValue = match[1];
      const pathStart = lineOffset + line.indexOf(pathValue);
      statements.push({
        kind: "load",
        command: "dbLoadRecords",
        path: pathValue,
        macros: match[2] || "",
        pathStart,
        pathEnd: pathStart + pathValue.length,
        start: lineOffset,
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

  if (reference.kind === "dbFile" && reference.name.toLowerCase().endsWith(".db")) {
    const substitutionsCandidate = normalizeFsPath(
      path.resolve(
        makefileDirectory,
        `${reference.name.slice(0, -".db".length)}.substitutions`,
      ),
    );
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
  return /\$\(|\$\{/.test(String(value || ""));
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

function createDiagnostic(start, end, message) {
  const diagnostic = new vscode.Diagnostic(
    new vscode.Range(start, end),
    message,
    vscode.DiagnosticSeverity.Error,
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
