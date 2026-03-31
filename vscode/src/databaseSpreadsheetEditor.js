const os = require("os");
const path = require("path");
const vscode = require("vscode");
const {
  BACKGROUND_COLOR_OPTIONS,
  normalizeBackgroundToken,
} = require("./spreadsheetCellBackgrounds");

const XLSX_EDITOR_VIEW_TYPE = "epicsWorkbench.spreadsheetXlsxEditor";
const DATABASE_EDITOR_VIEW_TYPE = "epicsWorkbench.spreadsheetDatabaseEditor";
const OPEN_IN_SPREADSHEET_COMMAND = "vscode-epics.openInSpreadsheet";
const OPEN_SPREADSHEET_WIDGET_COMMAND = "vscode-epics.openSpreadsheetWidget";
const DATABASE_LANGUAGE_ID = "database";
const DATABASE_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const RECORD_ROW_KIND = "record";
const COMMENT_ROW_KIND = "comment";

function registerDatabaseSpreadsheetEditor(extensionContext, options) {
  const controller = new EpicsSpreadsheetEditorController(extensionContext, options);
  const provider = new EpicsSpreadsheetCustomEditorProvider(controller);

  extensionContext.subscriptions.push(controller, provider);
  extensionContext.subscriptions.push(
    vscode.window.registerCustomEditorProvider(XLSX_EDITOR_VIEW_TYPE, provider, {
      webviewOptions: {
        retainContextWhenHidden: true,
      },
      supportsMultipleEditorsPerDocument: true,
    }),
  );
  extensionContext.subscriptions.push(
    vscode.window.registerCustomEditorProvider(DATABASE_EDITOR_VIEW_TYPE, provider, {
      webviewOptions: {
        retainContextWhenHidden: true,
      },
      supportsMultipleEditorsPerDocument: true,
    }),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(OPEN_IN_SPREADSHEET_COMMAND, async (resourceUri) => {
      const targetUri = resolveSpreadsheetTargetUri(resourceUri);
      if (!targetUri) {
        return;
      }

      await vscode.commands.executeCommand(
        "vscode.openWith",
        targetUri,
        getSpreadsheetViewTypeForUri(targetUri),
      );
    }),
  );
  extensionContext.subscriptions.push(
    vscode.commands.registerCommand(OPEN_SPREADSHEET_WIDGET_COMMAND, async (resourceUri) => {
      await controller.openSpreadsheetWidget(resourceUri);
    }),
  );
}

class EpicsSpreadsheetCustomEditorProvider {
  constructor(controller) {
    this.controller = controller;
  }

  openCustomDocument(uri, openContext) {
    return this.controller.openCustomDocument(uri, openContext);
  }

  resolveCustomEditor(document, webviewPanel) {
    return this.controller.resolveCustomEditor(document, webviewPanel);
  }

  saveCustomDocument(document, cancellation) {
    return this.controller.saveCustomDocument(document, cancellation);
  }

  saveCustomDocumentAs(document, destination, cancellation) {
    return this.controller.saveCustomDocumentAs(document, destination, cancellation);
  }

  revertCustomDocument(document, cancellation) {
    return this.controller.revertCustomDocument(document, cancellation);
  }

  backupCustomDocument(document, context, cancellation) {
    return this.controller.backupCustomDocument(document, context, cancellation);
  }

  onDidChangeCustomDocument(listener, thisArg, disposables) {
    return this.controller.onDidChangeCustomDocument(listener, thisArg, disposables);
  }
}

class EpicsSpreadsheetDocument {
  constructor(uri, sourceKind, workbook, options = {}) {
    this.uri = uri;
    this.sourceKind = sourceKind;
    this.workbook = workbook;
    this.associatedExcelUri = options.associatedExcelUri;
    this.defaultExcelFileName =
      options.defaultExcelFileName || getDefaultSpreadsheetExcelFileName(uri, sourceKind);
    this.onDispose = options.onDispose;
  }

  dispose() {
    this.onDispose?.();
  }
}

class EpicsSpreadsheetEditorController {
  constructor(extensionContext, options) {
    this.extensionContext = extensionContext;
    this.options = options;
    this._onDidChangeCustomDocument = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;
    this.panelsByDocument = new Map();
    this.documentsByUri = new Map();
    this.databaseWatchersByUri = new Map();
    this.pendingDocumentOptions = new Map();
    this.staticStatePayload = undefined;

    extensionContext.subscriptions.push(
      vscode.workspace.onDidSaveTextDocument(async (textDocument) => {
        await this.handlePotentialExternalDatabaseChange(textDocument.uri);
      }),
    );
  }

  dispose() {
    this._onDidChangeCustomDocument.dispose();
    this.pendingDocumentOptions.clear();
    this.panelsByDocument.clear();
    this.documentsByUri.clear();
    this.databaseWatchersByUri.forEach((watcher) => watcher.dispose());
    this.databaseWatchersByUri.clear();
  }

  async openSpreadsheetWidget(resourceUri) {
    const sourceDocument = await resolveSpreadsheetWidgetSourceDocument(resourceUri);
    if (sourceDocument?.uri && isDatabaseUri(sourceDocument.uri)) {
      await vscode.commands.executeCommand(
        "vscode.openWith",
        sourceDocument.uri,
        DATABASE_EDITOR_VIEW_TYPE,
      );
      return;
    }

    const defaultExcelFileName = buildSpreadsheetWidgetExcelFileName();
    let workbook = createDefaultSpreadsheetWorkbook("Spreadsheet");

    if (
      sourceDocument &&
      (
        sourceDocument.languageId === DATABASE_LANGUAGE_ID ||
        isDatabaseUri(sourceDocument.uri)
      )
    ) {
      workbook = normalizeWorkbookModel({
        sheets: [
          this.options.buildSheetModelFromDatabaseText(
            sourceDocument.getText(),
            path.basename(sourceDocument.uri.fsPath || sourceDocument.uri.path || "Database")
              .replace(/\.[^./\\]+$/, ""),
          ),
        ],
      });
    }

    const targetUri = createUntitledSpreadsheetUri(defaultExcelFileName);
    this.pendingDocumentOptions.set(targetUri.toString(), {
      workbook,
      defaultExcelFileName,
    });
    await vscode.commands.executeCommand(
      "vscode.openWith",
      targetUri,
      XLSX_EDITOR_VIEW_TYPE,
    );
  }

  async openCustomDocument(uri, openContext) {
    const sourceKind = getSourceKindForUri(uri);
    if (!sourceKind) {
      throw new Error("Only .xlsx and EPICS database files can be opened in the spreadsheet editor.");
    }

    let workbook;
    const pendingOptions = this.pendingDocumentOptions.get(uri.toString());
    if (pendingOptions) {
      this.pendingDocumentOptions.delete(uri.toString());
    }
    if (openContext?.backupId) {
      const backupUri = vscode.Uri.parse(openContext.backupId);
      const backupBytes = Buffer.from(await vscode.workspace.fs.readFile(backupUri));
      const parsedBackup = JSON.parse(backupBytes.toString("utf8"));
      workbook = normalizeWorkbookModel(parsedBackup?.workbook);
    } else if (pendingOptions?.workbook) {
      workbook = normalizeWorkbookModel(pendingOptions.workbook);
    } else if (uri?.scheme === "untitled") {
      workbook = createDefaultSpreadsheetWorkbook(
        path.basename(uri.fsPath || uri.path || "Spreadsheet")
          .replace(/\.[^./\\]+$/, ""),
      );
    } else if (sourceKind === "xlsx") {
      const workbookBuffer = Buffer.from(await vscode.workspace.fs.readFile(uri));
      const parsedWorkbook = this.options.parseEpicsWorkbookBuffer(workbookBuffer);
      if (parsedWorkbook.unsupportedSheetNames?.length) {
        throw new Error(
          `This workbook contains unsupported sheets: ${parsedWorkbook.unsupportedSheetNames.join(", ")}. Only EPICS sheets whose header starts with Record and Type are supported.`,
        );
      }
      if (!parsedWorkbook.sheets?.length) {
        throw new Error("No EPICS workbook sheets were found. The first row must start with Record and Type.");
      }
      workbook = normalizeWorkbookModel({ sheets: parsedWorkbook.sheets });
    } else {
      const fileBytes = await vscode.workspace.fs.readFile(uri);
      const text = Buffer.from(fileBytes).toString("utf8");
      workbook = normalizeWorkbookModel({
        sheets: [
          this.options.buildSheetModelFromDatabaseText(
            text,
            path.basename(uri.fsPath || uri.path || "Database")
              .replace(/\.[^./\\]+$/, ""),
          ),
        ],
      });
    }

    const document = new EpicsSpreadsheetDocument(uri, sourceKind, workbook, {
      associatedExcelUri:
        sourceKind === "xlsx" && uri?.scheme !== "untitled"
          ? uri
          : undefined,
      defaultExcelFileName:
        pendingOptions?.defaultExcelFileName ||
        getDefaultSpreadsheetExcelFileName(uri, sourceKind),
      onDispose: () => {
        this.documentsByUri.delete(uri.toString());
        this.disposeDatabaseWatcher(uri);
      },
    });
    this.documentsByUri.set(uri.toString(), document);
    this.ensureDatabaseWatcher(document);
    return document;
  }

  ensureDatabaseWatcher(document) {
    if (
      document.sourceKind !== "database" ||
      document.uri?.scheme === "untitled" ||
      this.databaseWatchersByUri.has(document.uri.toString())
    ) {
      return;
    }
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(
        path.dirname(document.uri.fsPath),
        path.basename(document.uri.fsPath),
      ),
    );
    const reload = async () => {
      await this.handlePotentialExternalDatabaseChange(document.uri);
    };
    watcher.onDidChange(reload);
    watcher.onDidCreate(reload);
    watcher.onDidDelete(() => {
      this.disposeDatabaseWatcher(document.uri);
    });
    this.databaseWatchersByUri.set(document.uri.toString(), watcher);
    this.extensionContext.subscriptions.push(watcher);
  }

  disposeDatabaseWatcher(uri) {
    const uriKey = uri?.toString();
    if (!uriKey) {
      return;
    }
    const watcher = this.databaseWatchersByUri.get(uriKey);
    watcher?.dispose();
    this.databaseWatchersByUri.delete(uriKey);
  }

  async handlePotentialExternalDatabaseChange(uri) {
    const uriKey = uri?.toString();
    const document = uriKey ? this.documentsByUri.get(uriKey) : undefined;
    if (!document || document.sourceKind !== "database" || document.uri?.scheme === "untitled") {
      return;
    }
    try {
      const nextWorkbook = await this.readDatabaseWorkbookFromUri(document.uri);
      if (serializeForComparison(document.workbook) === serializeForComparison(nextWorkbook)) {
        return;
      }
      document.workbook = nextWorkbook;
      await this.postDocumentState(document);
    } catch (error) {
      // Ignore transient read errors while the file is being rewritten.
    }
  }

  async readDatabaseWorkbookFromUri(uri) {
    const fileBytes = await vscode.workspace.fs.readFile(uri);
    const text = Buffer.from(fileBytes).toString("utf8");
    return normalizeWorkbookModel({
      sheets: [
        this.options.buildSheetModelFromDatabaseText(
          text,
          path.basename(uri.fsPath || uri.path || "Database")
            .replace(/\.[^./\\]+$/, ""),
        ),
      ],
    });
  }

  async resolveCustomEditor(document, webviewPanel) {
    const documentKey = document.uri.toString();
    const initialState = this.buildDocumentStatePayload(document, { includeStaticData: true });
    webviewPanel.webview.options = {
      enableScripts: true,
    };
    webviewPanel.webview.html = buildSpreadsheetEditorHtml(webviewPanel.webview, initialState);

    let panels = this.panelsByDocument.get(documentKey);
    if (!panels) {
      panels = new Set();
      this.panelsByDocument.set(documentKey, panels);
    }
    panels.add(webviewPanel);

    webviewPanel.onDidDispose(() => {
      const currentPanels = this.panelsByDocument.get(documentKey);
      currentPanels?.delete(webviewPanel);
      if (!currentPanels || currentPanels.size === 0) {
        this.panelsByDocument.delete(documentKey);
      }
    });

    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      switch (message?.type) {
        case "applySpreadsheetWorkbook":
          await this.applySpreadsheetWorkbook(document, message.workbook, webviewPanel);
          break;

        case "showDatabaseFile":
          await this.showDatabaseFile(document);
          break;

        case "previewSpreadsheetAsDatabase":
          await this.previewSpreadsheetAsDatabase(document, message.workbook, {
            currentSheetOnly: !!message.currentSheetOnly,
            activeSheetIndex: message.activeSheetIndex,
          });
          break;

        case "saveSpreadsheet":
          await this.saveSpreadsheet(document, message.workbook);
          break;

        case "saveSpreadsheetAsExcel":
          await this.saveSpreadsheetAsExcel(document, message.workbook, {
            associateAfterSave: message.associateAfterSave !== false,
          });
          break;

        case "saveSpreadsheetAsDatabase":
          await this.saveSpreadsheetAsDatabase(document, message.workbook, {
            currentSheetOnly: !!message.currentSheetOnly,
            activeSheetIndex: message.activeSheetIndex,
          });
          break;

        default:
          break;
      }
    });

  }

  async saveCustomDocument(document) {
    await this.writeDocumentToUri(document, document.uri);
  }

  async saveCustomDocumentAs(document, destination) {
    await this.writeDocumentToUri(document, destination);
  }

  async revertCustomDocument(document) {
    const reloaded = await this.openCustomDocument(document.uri, {});
    document.workbook = reloaded.workbook;
    this._onDidChangeCustomDocument.fire({ document });
    await this.postDocumentState(document);
  }

  async backupCustomDocument(document, context) {
    const backupBuffer = Buffer.from(
      JSON.stringify(
        {
          sourceKind: document.sourceKind,
          workbook: document.workbook,
        },
        null,
        2,
      ),
      "utf8",
    );
    await vscode.workspace.fs.writeFile(context.destination, backupBuffer);
    return {
      id: context.destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(context.destination);
        } catch (error) {
          // Ignore missing temporary backups.
        }
      },
    };
  }

  async applySpreadsheetWorkbook(document, nextWorkbook, sourcePanel) {
    document.workbook = normalizeWorkbookModel(nextWorkbook);
    this._onDidChangeCustomDocument.fire({ document });
    await this.postDocumentState(document, { preserveEditingPanel: sourcePanel });
  }

  async postDocumentState(document, options = {}) {
    const targetPanels = options.targetPanel
      ? [options.targetPanel]
      : [...(this.panelsByDocument.get(document.uri.toString()) || [])];
    const payload = this.buildDocumentStatePayload(document);
    for (const panel of targetPanels) {
      await panel.webview.postMessage({
        ...payload,
        preserveEditing: panel === options.preserveEditingPanel,
      });
    }
  }

  getStaticStatePayload() {
    if (!this.staticStatePayload) {
      const staticData = this.options.getStaticData();
      this.staticStatePayload = {
        recordTypes: [...(staticData.recordTypes || [])].sort(compareLabels),
        fieldNames: [...(staticData.allFields || [])].sort(compareLabels),
        menuChoicesByRecordType: serializeFieldMenuChoicesByRecordType(
          staticData.fieldMenuChoicesByRecordType,
        ),
      };
    }
    return this.staticStatePayload;
  }

  buildDocumentStatePayload(document, options = {}) {
    const staticData = this.options.getStaticData();
    const validation = validateWorkbook(document.workbook, staticData);
    return {
      type: "spreadsheetState",
      fileName: path.basename(document.uri.fsPath || document.uri.path || "Spreadsheet"),
      displayFileName: getSpreadsheetDisplayFileName(document),
      sourceKind: document.sourceKind,
      workbook: document.workbook,
      validation: {
        issues: validation.issues,
        issueMap: { ...(validation.issueMap || {}) },
      },
      canExportToDatabase: validation.issues.length === 0,
      canSaveToCurrentFile: !!document.associatedExcelUri,
      ...(options.includeStaticData ? this.getStaticStatePayload() : {}),
    };
  }

  async writeDocumentToUri(document, targetUri) {
    const targetKind = getSourceKindForUri(targetUri) || document.sourceKind;
    if (targetKind === "xlsx") {
      const workbookBuffer = this.options.buildWorkbookBufferFromSheetModels(
        document.workbook.sheets,
      );
      await vscode.workspace.fs.writeFile(targetUri, workbookBuffer);
      return;
    }

    const staticData = this.options.getStaticData();
    const validation = validateWorkbook(document.workbook, staticData);
    if (validation.issues.length > 0) {
      throw new Error(validation.issues[0].message);
    }

    const databaseText = buildDatabaseTextFromWorkbook(
      document.workbook,
      this.options.formatDatabaseText,
    );
    await vscode.workspace.fs.writeFile(targetUri, Buffer.from(databaseText, "utf8"));
  }

  getNormalizedWorkbookForOperation(document, workbook) {
    return normalizeWorkbookModel(workbook || document.workbook);
  }

  getWorkbookForDatabaseOperation(document, workbook, options = {}) {
    const workingWorkbook = this.getNormalizedWorkbookForOperation(document, workbook);
    if (!options.currentSheetOnly) {
      return workingWorkbook;
    }

    const normalizedSheetIndex = Number.isInteger(options.activeSheetIndex)
      ? Math.max(0, Math.min(options.activeSheetIndex, workingWorkbook.sheets.length - 1))
      : 0;
    const sheet = workingWorkbook.sheets[normalizedSheetIndex] || workingWorkbook.sheets[0];
    return normalizeWorkbookModel({
      sheets: sheet ? [sheet] : [],
    });
  }

  async previewSpreadsheetAsDatabase(document, workbook, options = {}) {
    if (document.sourceKind === "database" && document.uri?.scheme !== "untitled") {
      await this.showDatabaseFile(document);
      return;
    }
    try {
      const staticData = this.options.getStaticData();
      const workingWorkbook = this.getWorkbookForDatabaseOperation(document, workbook, options);
      const validation = validateWorkbook(workingWorkbook, staticData);
      if (validation.issues.length > 0) {
        throw new Error(validation.issues[0].message);
      }

      const databaseText = buildDatabaseTextFromWorkbook(
        workingWorkbook,
        this.options.formatDatabaseText,
      );
      const previewDocument = await vscode.workspace.openTextDocument({
        language: DATABASE_LANGUAGE_ID,
        content: databaseText,
      });
      await vscode.window.showTextDocument(previewDocument, { preview: false });
    } catch (error) {
      vscode.window.showErrorMessage(`Cannot preview as database: ${getErrorMessage(error)}`);
    }
  }

  async showDatabaseFile(document) {
    if (document.sourceKind !== "database" || document.uri?.scheme === "untitled") {
      return;
    }
    try {
      const dbDocument = await vscode.workspace.openTextDocument(document.uri);
      await vscode.window.showTextDocument(dbDocument, { preview: false });
    } catch (error) {
      vscode.window.showErrorMessage(`Cannot show database file: ${getErrorMessage(error)}`);
    }
  }

  associateSpreadsheetWithExcelFile(document, targetUri) {
    document.associatedExcelUri = targetUri;
    document.defaultExcelFileName = path.basename(targetUri.fsPath || targetUri.path || "Spreadsheet.xlsx");
  }

  async saveSpreadsheet(document, workbook) {
    const workingWorkbook = this.getNormalizedWorkbookForOperation(document, workbook);
    document.workbook = workingWorkbook;

    if (this.shouldUseHostSaveCommand(document)) {
      await vscode.commands.executeCommand("workbench.action.files.save");
      return;
    }

    if (document.associatedExcelUri) {
      try {
        await this.writeDocumentToUri(document, document.associatedExcelUri);
        await this.postDocumentState(document);
        vscode.window.showInformationMessage(
          `Saved ${path.basename(document.associatedExcelUri.fsPath || document.associatedExcelUri.path)}.`,
        );
      } catch (error) {
        vscode.window.showErrorMessage(`Failed to save Excel workbook: ${getErrorMessage(error)}`);
      }
      return;
    }

    await this.saveSpreadsheetAsExcel(document, workingWorkbook);
  }

  async saveSpreadsheetAsExcel(document, workbook, options = {}) {
    const workingWorkbook = this.getNormalizedWorkbookForOperation(document, workbook);
    document.workbook = workingWorkbook;
    const associateAfterSave = options.associateAfterSave !== false;
    const defaultUri = vscode.Uri.file(
      document.associatedExcelUri?.fsPath ||
        path.join(
          vscode.workspace.workspaceFolders?.[0]?.uri.fsPath || os.tmpdir(),
          document.defaultExcelFileName || "Spreadsheet.xlsx",
        ),
    );
    const targetUri = await vscode.window.showSaveDialog({
      defaultUri,
      filters: {
        "Excel Workbook": ["xlsx"],
      },
      saveLabel:
        document.sourceKind === "database" && !associateAfterSave
          ? "Export As Spreadsheet"
          : "Save Spreadsheet As",
    });
    if (!targetUri) {
      return;
    }

    try {
      await this.writeDocumentToUri(document, targetUri);
      if (associateAfterSave) {
        this.associateSpreadsheetWithExcelFile(document, targetUri);
        await this.postDocumentState(document);
      }
      vscode.window.showInformationMessage(
        `Saved ${path.basename(targetUri.fsPath || targetUri.path)}.`,
      );
    } catch (error) {
      vscode.window.showErrorMessage(`Failed to save Excel workbook: ${getErrorMessage(error)}`);
    }
  }

  async saveSpreadsheetAsDatabase(document, workbook, options = {}) {
    const workingWorkbook = this.getWorkbookForDatabaseOperation(document, workbook, options);
    const defaultUri = vscode.Uri.file(
      replaceFileExtension(
        document.uri.fsPath || document.uri.path || "database",
        ".db",
      ),
    );
    const targetUri = await vscode.window.showSaveDialog({
      defaultUri,
      filters: {
        "EPICS Database": ["db", "vdb", "template"],
      },
      saveLabel:
        document.sourceKind === "database"
          ? "Save As Database"
          : "Export Sheet as Database",
    });
    if (!targetUri) {
      return;
    }

    try {
      await this.writeDocumentToUri(
        {
          ...document,
          workbook: workingWorkbook,
        },
        targetUri,
      );
      if (document.sourceKind === "database") {
        vscode.window.showInformationMessage(
          `Saved ${path.basename(targetUri.fsPath || targetUri.path)}.`,
        );
      } else {
        const openedDocument = await vscode.workspace.openTextDocument(targetUri);
        await vscode.window.showTextDocument(openedDocument, { preview: false });
      }
    } catch (error) {
      vscode.window.showErrorMessage(`Failed to save database file: ${getErrorMessage(error)}`);
    }
  }

  shouldUseHostSaveCommand(document) {
    if (document.sourceKind === "database") {
      return document.uri?.scheme !== "untitled";
    }
    return document.sourceKind === "xlsx" &&
      document.uri?.scheme !== "untitled" &&
      !!document.associatedExcelUri &&
      document.associatedExcelUri.toString() === document.uri.toString();
  }
}

function resolveSpreadsheetTargetUri(resourceUri) {
  if (resourceUri instanceof vscode.Uri) {
    return resourceUri;
  }
  if (Array.isArray(resourceUri) && resourceUri[0] instanceof vscode.Uri) {
    return resourceUri[0];
  }
  if (resourceUri?.fsPath) {
    return vscode.Uri.file(resourceUri.fsPath);
  }
  const activeUri = vscode.window.activeTextEditor?.document?.uri;
  return activeUri || undefined;
}

async function resolveSpreadsheetWidgetSourceDocument(resourceUri) {
  const targetUri = resolveSpreadsheetTargetUri(resourceUri);
  const activeDocument = vscode.window.activeTextEditor?.document;
  if (
    targetUri &&
    activeDocument &&
    activeDocument.uri.toString() === targetUri.toString() &&
    (
      activeDocument.languageId === DATABASE_LANGUAGE_ID ||
      isDatabaseUri(activeDocument.uri)
    )
  ) {
    return activeDocument;
  }

  if (targetUri) {
    if (!isDatabaseUri(targetUri)) {
      return undefined;
    }
    try {
      return await vscode.workspace.openTextDocument(targetUri);
    } catch (error) {
      return undefined;
    }
  }

  if (
    activeDocument &&
    (
      activeDocument.languageId === DATABASE_LANGUAGE_ID ||
      isDatabaseUri(activeDocument.uri)
    )
  ) {
    return activeDocument;
  }

  return undefined;
}

function createUntitledSpreadsheetUri(fileName = buildSpreadsheetWidgetExcelFileName()) {
  const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath || os.tmpdir();
  return vscode.Uri.file(path.join(workspaceRoot, fileName)).with({
    scheme: "untitled",
  });
}

function buildSpreadsheetWidgetExcelFileName() {
  const now = new Date();
  const year = String(now.getFullYear()).padStart(4, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const milliseconds = String(now.getMilliseconds()).padStart(3, "0");
  return `db-${year}-${month}-${day}_${hours}-${minutes}-${seconds}-${milliseconds}.xlsx`;
}

function getDefaultSpreadsheetExcelFileName(uri, sourceKind) {
  if (sourceKind === "xlsx") {
    return path.basename(uri?.fsPath || uri?.path || "Spreadsheet.xlsx");
  }
  const sourceName = path.basename(uri?.fsPath || uri?.path || "database");
  return replaceFileExtension(sourceName, ".xlsx");
}

function getSpreadsheetDisplayFileName(document) {
  if (document.sourceKind === "database") {
    return path.basename(document.uri?.fsPath || document.uri?.path || "database.db");
  }
  if (document.associatedExcelUri) {
    return path.basename(document.associatedExcelUri.fsPath || document.associatedExcelUri.path || "Spreadsheet.xlsx");
  }
  return document.defaultExcelFileName || getDefaultSpreadsheetExcelFileName(document.uri, document.sourceKind);
}

function serializeForComparison(value) {
  return JSON.stringify(value ?? null);
}

function getSpreadsheetViewTypeForUri(uri) {
  return isExcelUri(uri) ? XLSX_EDITOR_VIEW_TYPE : DATABASE_EDITOR_VIEW_TYPE;
}

function getSourceKindForUri(uri) {
  if (isExcelUri(uri)) {
    return "xlsx";
  }
  if (isDatabaseUri(uri)) {
    return "database";
  }
  return undefined;
}

function isExcelUri(uri) {
  const extname = path.extname(uri?.fsPath || uri?.path || "").toLowerCase();
  return extname === ".xlsx";
}

function isDatabaseUri(uri) {
  const extname = path.extname(uri?.fsPath || uri?.path || "").toLowerCase();
  return DATABASE_EXTENSIONS.has(extname);
}

function replaceFileExtension(filePath, nextExtension) {
  return String(filePath || "database").replace(/\.[^./\\]+$/, "") + nextExtension;
}

function createEmptyRecordRow(columnCount) {
  return {
    kind: RECORD_ROW_KIND,
    values: new Array(columnCount).fill(""),
    backgrounds: new Array(columnCount).fill(""),
  };
}

function createCommentRow(text = "", background = "") {
  return {
    kind: COMMENT_ROW_KIND,
    text: String(text || ""),
    background: normalizeBackgroundToken(background),
  };
}

function createDefaultSpreadsheetWorkbook(sheetName = "Spreadsheet") {
  return normalizeWorkbookModel({
    sheets: [
      {
        name: sheetName || "Spreadsheet",
        headers: ["Record", "Type", "INP"],
        rows: [createEmptyRecordRow(3)],
      },
    ],
  });
}

function isCommentRow(row) {
  return !!row && !Array.isArray(row) && row.kind === COMMENT_ROW_KIND;
}

function getRecordRowValues(row, columnCount) {
  const values = Array.isArray(row?.values)
    ? row.values
    : Array.isArray(row)
      ? row
      : [];
  const normalized = [];
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    normalized.push(String(values[columnIndex] || ""));
  }
  return normalized;
}

function getNormalizedBackgroundValues(backgrounds, columnCount) {
  const values = Array.isArray(backgrounds) ? backgrounds : [];
  const normalized = [];
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    normalized.push(normalizeBackgroundToken(values[columnIndex]));
  }
  return normalized;
}

function normalizeSheetRow(row, columnCount) {
  if (isCommentRow(row)) {
    return createCommentRow(row.text, row.background);
  }

  return {
    kind: RECORD_ROW_KIND,
    values: getRecordRowValues(row, columnCount),
    backgrounds: getNormalizedBackgroundValues(row?.backgrounds, columnCount),
  };
}

function getRecordRows(sheet) {
  return (sheet?.rows || []).filter((row) => !isCommentRow(row));
}

function normalizeWorkbookModel(workbook) {
  const sheets = Array.isArray(workbook?.sheets) && workbook.sheets.length > 0
    ? workbook.sheets
    : [{ name: "Database", headers: ["Record", "Type"], rows: [] }];
  const usedSheetNames = new Set();
  return {
    sheets: sheets.map((sheet, sheetIndex) => {
      const headers = Array.isArray(sheet?.headers) ? sheet.headers : [];
      const normalizedHeaders = ["Record", "Type"];
      for (const header of headers.slice(2)) {
        let nextHeader = String(header || "").trim();
        if (!nextHeader) {
          nextHeader = "";
        }
        normalizedHeaders.push(nextHeader);
      }

      const normalizedRows = Array.isArray(sheet?.rows)
        ? sheet.rows.map((row) => normalizeSheetRow(row, normalizedHeaders.length))
        : [];
      const normalizedHeaderBackgrounds = getNormalizedBackgroundValues(
        sheet?.headerBackgrounds,
        normalizedHeaders.length,
      );

      let sheetName = String(sheet?.name || `Sheet ${sheetIndex + 1}`).trim() || `Sheet ${sheetIndex + 1}`;
      if (usedSheetNames.has(sheetName)) {
        let suffix = 2;
        while (usedSheetNames.has(`${sheetName}_${suffix}`)) {
          suffix += 1;
        }
        sheetName = `${sheetName}_${suffix}`;
      }
      usedSheetNames.add(sheetName);

      return {
        name: sheetName,
        headers: normalizedHeaders,
        headerBackgrounds: normalizedHeaderBackgrounds,
        rows: normalizedRows,
      };
    }),
  };
}

function validateWorkbook(workbook, staticData) {
  const issues = [];
  const issueMap = {};
  const seenIssues = new Set();
  const allRecordTypes = staticData.recordTypes || new Set();
  const allFieldNames = staticData.allFields || new Set();
  const duplicateTracker = new Map();

  workbook.sheets.forEach((sheet, sheetIndex) => {
    const recordRows = getRecordRows(sheet);
    const columnHasValues = new Array(sheet.headers.length).fill(false);
    recordRows.forEach((row) => {
      for (let columnIndex = 2; columnIndex < sheet.headers.length; columnIndex += 1) {
        if (columnHasValues[columnIndex]) {
          continue;
        }
        if (String(row.values[columnIndex] || "").trim()) {
          columnHasValues[columnIndex] = true;
        }
      }
    });

    const headerNames = new Map();
    sheet.headers.slice(2).forEach((fieldName, columnOffset) => {
      const columnIndex = columnOffset + 2;
      const trimmedFieldName = String(fieldName || "").trim();
      if (!trimmedFieldName) {
        if (columnHasValues[columnIndex]) {
          addIssue(
            issues,
            seenIssues,
            issueMap,
            sheetIndex,
            -1,
            columnIndex,
            "Field name cannot be empty when the column has values.",
          );
        }
        return;
      }

      if (headerNames.has(trimmedFieldName)) {
        addIssue(
          issues,
          seenIssues,
          issueMap,
          sheetIndex,
          -1,
          columnIndex,
          `Duplicate field column "${trimmedFieldName}".`,
        );
      } else {
        headerNames.set(trimmedFieldName, columnIndex);
      }

      if (!allFieldNames.has(trimmedFieldName)) {
        addIssue(
          issues,
          seenIssues,
          issueMap,
          sheetIndex,
          -1,
          columnIndex,
          `Unknown EPICS field "${trimmedFieldName}".`,
        );
      }
    });

    sheet.rows.forEach((row, rowIndex) => {
      if (isCommentRow(row)) {
        return;
      }

      const values = row.values || [];
      const rowHasData = values.some((value) => String(value || "").trim());
      if (!rowHasData) {
        return;
      }

      const recordName = String(values[0] || "").trim();
      const recordType = String(values[1] || "").trim();
      if (!recordName) {
        addIssue(issues, seenIssues, issueMap, sheetIndex, rowIndex, 0, "Record name cannot be empty.");
      }
      if (!recordType) {
        addIssue(issues, seenIssues, issueMap, sheetIndex, rowIndex, 1, "Record type cannot be empty.");
      } else if (!allRecordTypes.has(recordType)) {
        addIssue(
          issues,
          seenIssues,
          issueMap,
          sheetIndex,
          rowIndex,
          1,
          `Unknown EPICS record type "${recordType}".`,
        );
      }

      if (recordName) {
        const occurrences = duplicateTracker.get(recordName) || [];
        occurrences.push({ sheetIndex, rowIndex });
        duplicateTracker.set(recordName, occurrences);
      }

      const allowedFields = staticData.fieldsByRecordType?.get(recordType);
      const fieldTypes = staticData.fieldTypesByRecordType?.get(recordType);
      for (let columnIndex = 2; columnIndex < sheet.headers.length; columnIndex += 1) {
        const fieldName = String(sheet.headers[columnIndex] || "").trim();
        const value = String(values[columnIndex] || "");
        const trimmedValue = value.trim();
        if (!trimmedValue) {
          continue;
        }
        if (!fieldName) {
          continue;
        }

        if (allowedFields && !allowedFields.has(fieldName)) {
          addIssue(
            issues,
            seenIssues,
            issueMap,
            sheetIndex,
            rowIndex,
            columnIndex,
            `Field "${fieldName}" is not valid for record type "${recordType}".`,
          );
          continue;
        }

        const dbfType = fieldTypes?.get(fieldName);
        if (NUMERIC_DBF_TYPES.has(dbfType) && !isSkippableNumericFieldValue(trimmedValue)) {
          if (!isValidNumericFieldValue(trimmedValue, dbfType)) {
            addIssue(
              issues,
              seenIssues,
              issueMap,
              sheetIndex,
              rowIndex,
              columnIndex,
              `Field "${fieldName}" expects a ${dbfType} numeric value.`,
            );
          }
        }

        if (dbfType === "DBF_MENU" && !containsEpicsMacroReference(trimmedValue)) {
          const choices = getMenuFieldChoices(staticData, recordType, fieldName);
          if (choices.length > 0 && !choices.includes(trimmedValue)) {
            addIssue(
              issues,
              seenIssues,
              issueMap,
              sheetIndex,
              rowIndex,
              columnIndex,
              `Field "${fieldName}" must be one of the menu choices for "${recordType}".`,
            );
          }
        }
      }
    });
  });

  for (const [recordName, locations] of duplicateTracker.entries()) {
    if (locations.length < 2) {
      continue;
    }
    locations.forEach((location) => {
      addIssue(
        issues,
        seenIssues,
        issueMap,
        location.sheetIndex,
        location.rowIndex,
        0,
        `Duplicate record name "${recordName}" in this spreadsheet.`,
      );
    });
  }

  return {
    issues,
    issueMap,
  };
}

function addIssue(issues, seenIssues, issueMap, sheetIndex, rowIndex, columnIndex, message) {
  const key = `${sheetIndex}:${rowIndex}:${columnIndex}:${message}`;
  if (seenIssues.has(key)) {
    return;
  }
  seenIssues.add(key);
  const cellKey = `${sheetIndex}:${rowIndex}:${columnIndex}`;
  if (!issueMap[cellKey]) {
    issueMap[cellKey] = [];
  }
  issueMap[cellKey].push(message);
  issues.push({
    sheetIndex,
    rowIndex,
    columnIndex,
    message,
  });
}

function buildDatabaseTextFromWorkbook(workbook, formatDatabaseText) {
  const sheetTexts = workbook.sheets
    .map((sheet, index) => buildDatabaseTextFromSheet(sheet, index))
    .filter((text) => text.trim());
  const combinedText = sheetTexts.join("\n\n");
  return typeof formatDatabaseText === "function"
    ? formatDatabaseText(`${combinedText}\n`)
    : `${combinedText}\n`;
}

function buildDatabaseTextFromSheet(sheet, sheetIndex) {
  const lines = [];
  if ((sheet.name || "").trim() && sheetIndex > 0) {
    lines.push(`# Sheet: ${sheet.name}`);
    lines.push("");
  }

  for (const row of sheet.rows || []) {
    if (isCommentRow(row)) {
      appendCommentRowText(lines, row.text);
      continue;
    }

    const values = (row.values || []).map((value) => String(value || ""));
    if (!values.some((value) => value.trim())) {
      continue;
    }
    const recordName = values[0].trim();
    const recordType = values[1].trim();
    if (!recordName || !recordType) {
      continue;
    }

    lines.push(`record(${recordType}, "${escapeDatabaseString(recordName)}") {`);
    for (let columnIndex = 2; columnIndex < sheet.headers.length; columnIndex += 1) {
      const fieldName = String(sheet.headers[columnIndex] || "").trim();
      const value = values[columnIndex].trim();
      if (!fieldName || !value) {
        continue;
      }
      lines.push(`    field(${fieldName}, "${escapeDatabaseString(value)}")`);
    }
    lines.push("}");
    lines.push("");
  }

  return lines.join("\n");
}

function appendCommentRowText(lines, text) {
  const commentLines = String(text || "")
    .split(/\r?\n/)
    .map((line) => `# ${line}`.trimEnd());
  if (!commentLines.some((line) => line.trim() !== "#")) {
    return;
  }

  lines.push(...commentLines);
  lines.push("");
}

function getMenuFieldChoices(staticData, recordType, fieldName) {
  if (!recordType || !fieldName) {
    return [];
  }
  const fieldMenus = staticData.fieldMenuChoicesByRecordType?.get(recordType);
  return fieldMenus?.get(fieldName) || [];
}

function serializeFieldMenuChoicesByRecordType(fieldMenuChoicesByRecordType) {
  const serialized = {};
  for (const [recordType, fieldMenus] of fieldMenuChoicesByRecordType || []) {
    const serializedFieldMenus = {};
    for (const [fieldName, choices] of fieldMenus || []) {
      serializedFieldMenus[fieldName] = Array.isArray(choices) ? choices.slice() : [];
    }
    serialized[recordType] = serializedFieldMenus;
  }
  return serialized;
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

function escapeDatabaseString(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n/g, "\\n")
    .replace(/"/g, '\\"');
}

function compareLabels(left, right) {
  return String(left || "").localeCompare(String(right || ""));
}

function getErrorMessage(error) {
  return error instanceof Error ? error.message : String(error || "Unknown error");
}

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

function buildSpreadsheetEditorHtml(webview, initialState) {
  const nonce = String(Date.now());
  const serializedInitialState = serializeForWebview(initialState);
  const serializedBackgroundOptions = serializeForWebview(BACKGROUND_COLOR_OPTIONS);
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta
    http-equiv="Content-Security-Policy"
    content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';"
  />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>EPICS Spreadsheet</title>
  <style>
    :root {
      color-scheme: light dark;
      --spreadsheet-font-size: 14px;
      --spreadsheet-row-min-height: 28px;
      --spreadsheet-cell-padding-y: 3px;
      --spreadsheet-cell-padding-x: 6px;
      --spreadsheet-row-index-width: 74px;
      --spreadsheet-choice-cell-min-width: 150px;
      --spreadsheet-comment-min-width: 200px;
      --spreadsheet-action-button-width: 22px;
      --spreadsheet-column-resizer-width: 6px;
      --spreadsheet-selection-accent: color-mix(
        in srgb,
        var(--vscode-focusBorder, #0078d4) 72%,
        var(--vscode-textLink-foreground, #0a84ff) 28%
      );
      --spreadsheet-selection-fill: color-mix(
        in srgb,
        var(--spreadsheet-selection-accent) 38%,
        var(--vscode-editor-background) 62%
      );
      --spreadsheet-selection-strong-fill: color-mix(
        in srgb,
        var(--spreadsheet-selection-accent) 56%,
        var(--vscode-editor-background) 44%
      );
      --spreadsheet-multi-selection-fill: color-mix(
        in srgb,
        var(--vscode-descriptionForeground, #808080) 18%,
        var(--vscode-editor-background) 82%
      );
      --spreadsheet-multi-selection-outline: color-mix(
        in srgb,
        var(--vscode-descriptionForeground, #808080) 45%,
        white 55%
      );
      --spreadsheet-selection-outline: color-mix(
        in srgb,
        var(--spreadsheet-selection-accent) 72%,
        white 28%
      );
      --spreadsheet-selection-foreground: var(
        --vscode-list-activeSelectionForeground,
        var(--vscode-editor-foreground)
      );
    }
    body {
      margin: 0;
      font-family: var(--vscode-font-family);
      color: var(--vscode-editor-foreground);
      background: var(--vscode-editor-background);
    }
    .app {
      display: grid;
      grid-template-rows: auto auto 1fr;
      min-height: 100vh;
    }
    .toolbar {
      display: flex;
      flex-direction: column;
      gap: 6px;
      align-items: flex-start;
      padding: 8px 10px;
      border-bottom: 1px solid var(--vscode-panel-border);
      background: color-mix(in srgb, var(--vscode-editor-background) 88%, var(--vscode-sideBar-background) 12%);
    }
    .toolbar-title {
      font-weight: 600;
    }
    .toolbar-actions {
      display: grid;
      gap: 6px;
    }
    .toolbar-row {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      align-items: center;
    }
    .toolbar-background-group {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      color: inherit;
    }
    .toolbar-background-group > span {
      color: var(--vscode-descriptionForeground);
      white-space: nowrap;
    }
    .toolbar-palette {
      display: inline-flex;
      flex-wrap: wrap;
      gap: 4px;
      align-items: center;
    }
    .toolbar-palette button {
      width: 22px;
      min-width: 22px;
      height: 22px;
      padding: 0;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 5px;
    }
    .toolbar-palette button.is-active {
      box-shadow: inset 0 0 0 1px var(--spreadsheet-selection-outline);
      background: var(--spreadsheet-selection-fill);
    }
    .toolbar-palette button[disabled] {
      opacity: 0.45;
      cursor: default;
    }
    .validation-badge {
      padding: 4px 10px;
      border-radius: 999px;
      border: 1px solid color-mix(in srgb, var(--vscode-panel-border) 75%, transparent);
      background: color-mix(in srgb, var(--vscode-editor-background) 92%, var(--vscode-sideBar-background) 8%);
      color: var(--vscode-descriptionForeground);
      font-size: 0.85rem;
      white-space: nowrap;
    }
    .validation-badge.has-issues {
      border-color: color-mix(in srgb, var(--vscode-errorForeground) 45%, transparent);
      background: color-mix(in srgb, var(--vscode-editor-background) 84%, var(--vscode-errorForeground) 16%);
      color: var(--vscode-errorForeground);
    }
    .toolbar button {
      border: 1px solid var(--vscode-button-border, transparent);
      background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
      color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
      padding: 5px 9px;
      border-radius: 6px;
      font: inherit;
      cursor: pointer;
    }
    .toolbar button.primary {
      background: var(--vscode-button-background);
      color: var(--vscode-button-foreground);
    }
    .toolbar button:hover {
      background: var(--vscode-button-secondaryHoverBackground, var(--vscode-button-hoverBackground));
    }
    .toolbar button.primary:hover {
      background: var(--vscode-button-hoverBackground);
    }
    .toolbar button.is-disabled,
    .toolbar button.is-disabled:hover,
    .toolbar button.is-disabled:focus-visible {
      opacity: 0.55;
      cursor: default;
      background: var(--vscode-button-secondaryBackground, var(--vscode-button-background));
      color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground));
      box-shadow: none;
      outline: none;
    }
    .toolbar-select-control {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font: inherit;
      color: inherit;
    }
    .toolbar-select-control select {
      border: 1px solid var(--vscode-dropdown-border, var(--vscode-panel-border));
      border-radius: 6px;
      padding: 4px 8px;
      font: inherit;
      color: inherit;
      background: var(--vscode-dropdown-background, var(--vscode-editor-background));
    }
    .sheet-tabs {
      display: flex;
      gap: 2px;
      padding: 8px 12px 0;
      border-bottom: 1px solid var(--vscode-panel-border);
      overflow-x: auto;
      font-size: var(--spreadsheet-font-size);
    }
    .sheet-tab {
      border: 1px solid var(--vscode-panel-border);
      border-bottom: none;
      padding: 8px 12px;
      border-radius: 8px 8px 0 0;
      cursor: pointer;
      white-space: nowrap;
      background: color-mix(in srgb, var(--vscode-editor-background) 92%, var(--vscode-sideBar-background) 8%);
    }
    .sheet-tab.active {
      background: var(--spreadsheet-selection-fill);
      color: var(--spreadsheet-selection-foreground);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-selection-outline);
    }
    .sheet-tab.draggable {
      cursor: grab;
    }
    body.reordering-sheet .sheet-tab.draggable {
      cursor: grabbing;
    }
    .sheet-tab.drop-before {
      box-shadow: inset 2px 0 0 var(--vscode-focusBorder);
    }
    .sheet-tab.drop-after {
      box-shadow: inset -2px 0 0 var(--vscode-focusBorder);
    }
    .sheet-tab.add-sheet {
      padding: 8px 10px;
      font-weight: 600;
    }
    .sheet-tab-input {
      min-width: 120px;
      box-sizing: border-box;
      border: 1px solid var(--vscode-focusBorder);
      border-bottom: none;
      padding: 8px 12px;
      border-radius: 8px 8px 0 0;
      font: inherit;
      color: inherit;
      background: var(--vscode-editor-background);
      outline: none;
    }
    .grid-panel {
      overflow: auto;
      padding: 8px 10px 10px;
      font-size: var(--spreadsheet-font-size);
    }
    body[data-density-mode="compact"] .grid-panel {
      padding: 4px 6px 6px;
    }
    table {
      border-collapse: collapse;
      width: max-content;
    }
    th,
    td {
      border: 1px solid var(--vscode-panel-border);
      padding: 0;
      vertical-align: top;
      background: var(--vscode-editor-background);
    }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: color-mix(in srgb, var(--vscode-editor-background) 84%, var(--vscode-sideBar-background) 16%);
    }
    .header-corner {
      top: 0;
      z-index: 5;
    }
    .column-index-cell {
      top: 0;
      z-index: 4;
      background: color-mix(in srgb, var(--vscode-editor-background) 88%, var(--vscode-sideBar-background) 12%);
    }
    .field-header-cell {
      top: var(--spreadsheet-row-min-height);
      z-index: 3;
    }
    .row-index {
      min-width: var(--spreadsheet-row-index-width);
      width: var(--spreadsheet-row-index-width);
      text-align: center;
      color: var(--vscode-descriptionForeground);
      background: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-sideBar-background) 10%);
      position: sticky;
      left: 0;
      z-index: 3;
    }
    .header-wrap,
    .column-index-wrap,
    .row-index-wrap,
    .grid-editor {
      display: flex;
      align-items: stretch;
      width: 100%;
      min-height: var(--spreadsheet-row-min-height);
    }
    .column-select-button {
      border: none;
      background: transparent;
      color: inherit;
      font: inherit;
      font-weight: 400;
      cursor: pointer;
      width: 100%;
      text-align: center;
      padding: 0 var(--spreadsheet-cell-padding-x);
      min-height: var(--spreadsheet-row-min-height);
    }
    .header-cell,
    .grid-cell,
    .comment-input {
      display: block;
      flex: 1 1 auto;
      width: 100%;
      min-width: 0;
      box-sizing: border-box;
      border: none;
      padding: var(--spreadsheet-cell-padding-y) var(--spreadsheet-cell-padding-x);
      font: inherit;
      color: inherit;
      background: transparent;
      outline: none;
      line-height: 1.25;
    }
    .comment-input {
      min-width: var(--spreadsheet-comment-min-width);
      resize: none;
      height: var(--spreadsheet-row-min-height);
      min-height: var(--spreadsheet-row-min-height);
      max-height: var(--spreadsheet-row-min-height);
      line-height: 1.25;
      overflow-y: auto;
    }
    body[data-density-mode="compact"] .header-cell,
    body[data-density-mode="compact"] .grid-cell,
    body[data-density-mode="compact"] .comment-input,
    body[data-density-mode="compact"] .row-select-button {
      line-height: 1.05;
    }
    .header-cell[readonly] {
      color: var(--vscode-disabledForeground);
      font-weight: 600;
    }
    .comment-cell {
      padding: 0;
      background: color-mix(in srgb, var(--vscode-editor-background) 95%, var(--vscode-sideBar-background) 5%);
    }
    .column-index-cell.has-custom-background,
    .field-header-cell.has-custom-background,
    .field-header-cell.has-custom-background .header-cell,
    td.has-custom-background,
    td.has-custom-background .grid-editor,
    .comment-cell.has-custom-background,
    .comment-cell.has-custom-background .comment-input {
      background: var(--spreadsheet-custom-background);
    }
    .cell-error,
    .cell-error .header-cell,
    .cell-error .header-cell[readonly],
    .cell-error .grid-editor,
    .cell-error .grid-cell,
    .cell-error .comment-input,
    .cell-error button {
      color: rgba(255, 0, 0, 1);
    }
    .cell-error .header-cell::placeholder,
    .cell-error .grid-cell::placeholder,
    .cell-error .comment-input::placeholder {
      color: rgba(255, 0, 0, 1);
      opacity: 1;
    }
    .column-index-cell.selected,
    .field-header-cell.selected,
    td.column-selected,
    td.column-selected .grid-editor {
      background: var(--spreadsheet-multi-selection-fill);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-multi-selection-outline);
    }
    .cell-selected,
    .cell-selected .grid-editor {
      background: var(--spreadsheet-selection-fill);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-selection-outline);
    }
    .cell-selected.cell-selected-multi,
    .cell-selected.cell-selected-multi .grid-editor {
      background: var(--spreadsheet-multi-selection-fill);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-multi-selection-outline);
    }
    .field-header-cell.has-custom-background.selected,
    .field-header-cell.has-custom-background.selected .header-cell,
    .column-index-cell.has-custom-background.selected,
    td.has-custom-background.column-selected,
    td.has-custom-background.column-selected .grid-editor,
    .row-selected td.has-custom-background,
    .row-selected td.has-custom-background .grid-editor,
    .row-selected .comment-cell.has-custom-background,
    .row-selected .comment-cell.has-custom-background .comment-input {
      background: color-mix(
        in srgb,
        var(--spreadsheet-custom-background) 76%,
        var(--spreadsheet-multi-selection-fill) 24%
      );
    }
    td.has-custom-background.cell-selected,
    td.has-custom-background.cell-selected .grid-editor {
      background: color-mix(
        in srgb,
        var(--spreadsheet-custom-background) 72%,
        var(--spreadsheet-selection-fill) 28%
      );
    }
    td.has-custom-background.cell-selected.cell-selected-multi,
    td.has-custom-background.cell-selected.cell-selected-multi .grid-editor {
      background: color-mix(
        in srgb,
        var(--spreadsheet-custom-background) 76%,
        var(--spreadsheet-multi-selection-fill) 24%
      );
    }
    .grid-editor.has-choice {
      min-width: var(--spreadsheet-choice-cell-min-width);
    }
    .menu-trigger,
    .choice-trigger {
      color: inherit;
      width: var(--spreadsheet-action-button-width);
      min-width: var(--spreadsheet-action-button-width);
      border-radius: 4px;
      cursor: pointer;
      line-height: 1;
      margin: 1px;
      flex: 0 0 auto;
    }
    .choice-trigger {
      border: 1px solid var(--vscode-panel-border);
      background: transparent;
      opacity: 0;
      pointer-events: none;
    }
    body[data-density-mode="compact"] .menu-trigger,
    body[data-density-mode="compact"] .choice-trigger {
      margin: 0;
      border-radius: 3px;
    }
    .menu-trigger {
      border: none;
      background: transparent;
      padding: 0;
      opacity: 0;
      pointer-events: none;
    }
    th:hover .menu-trigger,
    th:focus-within .menu-trigger,
    .row-index-cell:hover .menu-trigger,
    .row-index-cell:focus-within .menu-trigger,
    .menu-trigger:focus-visible {
      opacity: 1;
      pointer-events: auto;
    }
    td:hover .choice-trigger,
    td:focus-within .choice-trigger,
    .choice-trigger:focus-visible {
      opacity: 1;
      pointer-events: auto;
    }
    .column-resizer {
      width: var(--spreadsheet-column-resizer-width);
      min-width: var(--spreadsheet-column-resizer-width);
      cursor: col-resize;
      border-radius: 3px;
      margin: 1px 1px 1px 0;
      flex: 0 0 auto;
      background: transparent;
      opacity: 0;
    }
    .column-resizer:hover {
      background: transparent;
    }
    body[data-density-mode="compact"] .column-resizer {
      margin: 0;
      border-radius: 2px;
    }
    .row-select-button {
      appearance: none;
      border: none;
      background: transparent;
      color: inherit;
      display: flex;
      align-items: flex-start;
      justify-content: flex-start;
      font: inherit;
      line-height: 1.25;
      cursor: pointer;
      width: 100%;
      text-align: left;
      box-sizing: border-box;
      padding: var(--spreadsheet-cell-padding-y) var(--spreadsheet-cell-padding-x);
      min-height: var(--spreadsheet-row-min-height);
    }
    .row-select-button.drag-handle,
    .column-index-wrap.drag-handle,
    .column-index-wrap.drag-handle .column-select-button {
      cursor: grab;
    }
    body.reordering-row .row-select-button.drag-handle,
    body.reordering-column .column-index-wrap.drag-handle,
    body.reordering-column .column-index-wrap.drag-handle .column-select-button {
      cursor: grabbing;
    }
    .row-index-cell.selected,
    .row-selected td,
    .row-selected .comment-cell {
      background: var(--spreadsheet-multi-selection-fill);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-multi-selection-outline);
    }
    .row-index-cell.selected .row-select-button {
      color: var(--vscode-editor-foreground);
    }
    .column-index-cell.selected .column-select-button {
      color: var(--vscode-editor-foreground);
    }
    .row-index-cell.drop-before {
      box-shadow: inset 0 2px 0 var(--vscode-focusBorder);
    }
    .row-index-cell.drop-after {
      box-shadow: inset 0 -2px 0 var(--vscode-focusBorder);
    }
    th.drop-before {
      box-shadow: inset 2px 0 0 var(--vscode-focusBorder);
    }
    th.drop-after {
      box-shadow: inset -2px 0 0 var(--vscode-focusBorder);
    }
    .floating-menu {
      position: fixed;
      z-index: 2000;
      min-width: 220px;
      max-width: min(360px, calc(100vw - 24px));
      max-height: min(360px, calc(100vh - 24px));
      overflow: auto;
      border: 1px solid var(--vscode-panel-border);
      border-radius: 8px;
      background: color-mix(in srgb, var(--vscode-editor-background) 92%, var(--vscode-sideBar-background) 8%);
      box-shadow: 0 10px 28px rgba(0, 0, 0, 0.28);
      padding: 4px;
      font-size: var(--spreadsheet-font-size);
    }
    .floating-menu button {
      display: flex;
      align-items: center;
      gap: 10px;
      width: 100%;
      text-align: left;
      border: none;
      border-radius: 6px;
      background: transparent;
      color: inherit;
      font: inherit;
      padding: 7px 9px;
      cursor: pointer;
    }
    .menu-swatch {
      width: 14px;
      min-width: 14px;
      height: 14px;
      border-radius: 4px;
      border: 1px solid color-mix(in srgb, var(--vscode-panel-border) 85%, transparent);
      box-sizing: border-box;
      flex: 0 0 auto;
    }
    .menu-swatch.no-color {
      background:
        linear-gradient(
          135deg,
          transparent 0 44%,
          color-mix(in srgb, var(--vscode-errorForeground, #f14c4c) 70%, transparent) 44% 56%,
          transparent 56% 100%
        ),
        var(--vscode-editor-background);
    }
    .floating-menu button:hover:not([disabled]),
    .floating-menu button:focus-visible {
      background: var(--spreadsheet-selection-fill);
      color: var(--spreadsheet-selection-foreground);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-selection-outline);
      outline: none;
    }
    .floating-menu button.active {
      background: var(--spreadsheet-selection-strong-fill);
      color: var(--spreadsheet-selection-foreground);
      box-shadow: inset 0 0 0 1px var(--spreadsheet-selection-outline);
    }
    .floating-menu button[disabled] {
      opacity: 0.55;
      cursor: default;
    }
    .menu-hint {
      padding: 6px 9px;
      color: var(--vscode-descriptionForeground);
      font-size: 0.85rem;
      white-space: normal;
    }
    .menu-panel {
      display: grid;
      gap: 8px;
      padding: 6px;
    }
    .menu-warning {
      color: var(--vscode-errorForeground, #f14c4c);
      font-weight: 600;
      white-space: normal;
    }
    .menu-label {
      color: var(--vscode-descriptionForeground);
      font-size: 0.9rem;
      white-space: normal;
    }
    .menu-text-input {
      width: 100%;
      box-sizing: border-box;
      border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
      border-radius: 6px;
      padding: 6px 8px;
      font: inherit;
      color: inherit;
      background: var(--vscode-input-background, var(--vscode-editor-background));
      outline: none;
    }
    .menu-text-input:focus {
      border-color: var(--vscode-focusBorder);
    }
    .floating-menu button.menu-danger {
      background: color-mix(
        in srgb,
        var(--vscode-errorForeground, #f14c4c) 22%,
        var(--vscode-button-background) 78%
      );
      color: var(--vscode-button-foreground);
    }
    .floating-menu button.menu-danger:hover:not([disabled]),
    .floating-menu button.menu-danger:focus-visible {
      background: color-mix(
        in srgb,
        var(--vscode-errorForeground, #f14c4c) 34%,
        var(--vscode-button-hoverBackground, var(--vscode-button-background)) 66%
      );
      color: var(--vscode-button-foreground);
      box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--vscode-errorForeground, #f14c4c) 55%, transparent);
    }
    .validation-tooltip {
      position: fixed;
      z-index: 2100;
      max-width: min(420px, calc(100vw - 24px));
      padding: 6px 9px;
      border-radius: 6px;
      border: 1px solid color-mix(in srgb, var(--vscode-errorForeground, #f14c4c) 36%, transparent);
      background: color-mix(
        in srgb,
        var(--vscode-editorHoverWidget-background, var(--vscode-editor-background)) 90%,
        var(--vscode-errorForeground, #f14c4c) 10%
      );
      color: var(--vscode-editorHoverWidget-foreground, var(--vscode-editor-foreground));
      box-shadow: 0 10px 28px rgba(0, 0, 0, 0.28);
      white-space: pre-wrap;
      pointer-events: none;
      font-size: var(--spreadsheet-font-size);
    }
    body.column-resizing,
    body.column-resizing * {
      cursor: col-resize !important;
      user-select: none !important;
    }
    @media (max-width: 1000px) {
      .toolbar-actions,
      .toolbar-row {
        width: 100%;
      }
    }
  </style>
</head>
<body>
    <div class="app">
    <div class="toolbar">
      <div class="toolbar-title" id="toolbar-title">EPICS Spreadsheet</div>
      <div class="toolbar-actions">
        <div class="toolbar-row">
          <div id="validation-badge" class="validation-badge">Loading...</div>
          <label class="toolbar-select-control" for="font-size-select">
            <span>Font size</span>
            <select id="font-size-select"></select>
          </label>
          <label class="toolbar-select-control" for="density-select">
            <span>Layout</span>
            <select id="density-select">
              <option value="normal">Normal</option>
              <option value="compact">Compact</option>
            </select>
          </label>
          <button id="resize-columns-button" type="button">Auto resize spreadsheet</button>
        </div>
        <div class="toolbar-row">
          <button id="undo-button" type="button">Undo</button>
          <button id="redo-button" type="button">Redo</button>
          <button id="save-button" type="button">Save Spreadsheet</button>
          <button id="save-as-button" type="button">Save As Spreadsheet</button>
          <div class="toolbar-background-group">
            <span>Background</span>
            <div id="background-palette" class="toolbar-palette" role="group" aria-label="Set background color"></div>
          </div>
        </div>
        <div class="toolbar-row">
          <button id="preview-db-button" type="button">Preview Sheet As DB</button>
          <button id="save-db-button" type="button">Export Sheet As DB</button>
        </div>
      </div>
    </div>
    <div id="sheet-tabs" class="sheet-tabs"></div>
    <div class="grid-panel">
      <table>
        <colgroup id="grid-columns"></colgroup>
        <thead id="grid-head"></thead>
        <tbody id="grid-body"></tbody>
      </table>
    </div>
    <div id="floating-menu" class="floating-menu" hidden></div>
    <div id="suggestion-menu" class="floating-menu" hidden></div>
    <div id="validation-tooltip" class="validation-tooltip" hidden></div>
  </div>
  <script nonce="${nonce}">
    const initialState = ${serializedInitialState};
    const BACKGROUND_COLOR_OPTIONS = ${serializedBackgroundOptions};
    const startupValidationBadge = document.getElementById("validation-badge");
    function showStartupError(label, error) {
      if (startupValidationBadge) {
        startupValidationBadge.className = "validation-badge has-issues";
        startupValidationBadge.textContent = label;
        startupValidationBadge.title = error instanceof Error ? error.message : String(error || label);
      }
      console.error(label + ":", error);
    }
    window.addEventListener("error", (event) => {
      showStartupError("Startup error", event.error || event.message);
    });
    window.addEventListener("unhandledrejection", (event) => {
      showStartupError("Startup error", event.reason);
    });
    const vscode = typeof acquireVsCodeApi === "function"
      ? acquireVsCodeApi()
      : { postMessage() {} };
    if (typeof acquireVsCodeApi !== "function") {
      showStartupError("Webview API missing", "acquireVsCodeApi() is unavailable.");
    }
    try {
    const SAVE_DEBOUNCE_MS = 180;
    const VALIDATION_TOOLTIP_DELAY_MS = 1000;
    const BACKGROUND_OPTION_MAP = new Map(
      BACKGROUND_COLOR_OPTIONS.map((option) => [String(option.token || ""), option]),
    );
    const DEFAULT_FONT_SIZE_PX = Number.parseFloat(window.getComputedStyle(document.body).fontSize) || 14;
    const DEFAULT_DENSITY_MODE = "normal";
    const DENSITY_METRICS = {
      normal: {
        rowIndexColumnWidth: 74,
        minColumnWidth: 72,
        maxAutoColumnCharacters: 16,
        cellHorizontalPadding: 12,
        actionButtonWidth: 22,
        columnResizerWidth: 6,
        headerActionExtraWidth: 18,
        choiceActionWidth: 24,
        cellPaddingY: 3,
        cellPaddingX: 6,
        commentMinWidth: 200,
        choiceCellMinWidth: 150,
        rowMinHeightFloor: 24,
        rowMinHeightExtra: 12,
      },
      compact: {
        rowIndexColumnWidth: 52,
        minColumnWidth: 48,
        maxAutoColumnCharacters: 10,
        cellHorizontalPadding: 6,
        actionButtonWidth: 18,
        columnResizerWidth: 4,
        headerActionExtraWidth: 8,
        choiceActionWidth: 18,
        cellPaddingY: 0,
        cellPaddingX: 3,
        commentMinWidth: 120,
        choiceCellMinWidth: 92,
        rowMinHeightFloor: 18,
        rowMinHeightExtra: 3,
      },
    };
    const FIELD_NAME_SUGGESTION_LIMIT = 40;
    const toolbarTitle = document.getElementById("toolbar-title");
    const validationBadge = document.getElementById("validation-badge");
    const sheetTabs = document.getElementById("sheet-tabs");
    const gridColumns = document.getElementById("grid-columns");
    const gridHead = document.getElementById("grid-head");
    const gridBody = document.getElementById("grid-body");
    const floatingMenu = document.getElementById("floating-menu");
    const suggestionMenu = document.getElementById("suggestion-menu");
    const validationTooltip = document.getElementById("validation-tooltip");
    const fontSizeSelect = document.getElementById("font-size-select");
    const densitySelect = document.getElementById("density-select");
    const resizeColumnsButton = document.getElementById("resize-columns-button");
    const previewDbButton = document.getElementById("preview-db-button");
    const undoButton = document.getElementById("undo-button");
    const redoButton = document.getElementById("redo-button");
    const saveButton = document.getElementById("save-button");
    const saveAsButton = document.getElementById("save-as-button");
    const saveDbButton = document.getElementById("save-db-button");
    const backgroundPalette = document.getElementById("background-palette");

    const persistedUiState = typeof vscode.getState === "function"
      ? (vscode.getState() || {})
      : {};
    let state = initialState;
    let fontSizePx = Number(persistedUiState.fontSizePx) || DEFAULT_FONT_SIZE_PX;
    let densityMode = persistedUiState.densityMode === "compact"
      ? "compact"
      : DEFAULT_DENSITY_MODE;
    let activeSheetIndex = 0;
    let pendingSaveTimer;
    let workbookDirty = false;
    let rowSelectionAnchorIndex = undefined;
    let rowSelectionPivotIndex = undefined;
    let isSelectingRows = false;
    let selectedRowIndexes = [];
    let columnSelectionAnchorIndex = undefined;
    let columnSelectionPivotIndex = undefined;
    let isSelectingColumns = false;
    let selectedColumnIndexes = [];
    let cellSelectionAnchor;
    let isSelectingCells = false;
    let selectedCellRange;
    let cellSelectionEnhancementsEnabled = true;
    let renderWarningMessage = "";
    let rowClipboard;
    let columnClipboard;
    let activeSuggestionState;
    let textMeasureCanvas;
    let columnResizeState;
    let rowReorderDragState;
    let columnReorderDragState;
    let sheetReorderDragState;
    let reorderDropIndicatorState;
    let validationTooltipTimer;
    let validationTooltipHost;
    let pendingValidationTooltipHost;
    let validationTooltipPointer;
    let editingSheetIndex = undefined;
    let undoHistory = [];
    let redoHistory = [];
    const columnWidthStateBySheet = new Map();
    const HISTORY_LIMIT = 100;

    function getActiveSheet() {
      const sheets = state.workbook?.sheets || [];
      if (activeSheetIndex >= sheets.length) {
        activeSheetIndex = 0;
      }
      const sheet = sheets[activeSheetIndex] || { name: "Database", headers: ["Record", "Type"], rows: [] };
      ensureHeaderBackgroundLength(sheet);
      return sheet;
    }

    function getColumnWidthState(sheetIndex = activeSheetIndex) {
      let widthState = columnWidthStateBySheet.get(sheetIndex);
      if (!widthState) {
        widthState = {
          auto: [],
          manual: {},
        };
        columnWidthStateBySheet.set(sheetIndex, widthState);
      }
      return widthState;
    }

    function isCommentRow(row) {
      return !!row && row.kind === "comment";
    }

    function normalizeBackgroundToken(token) {
      const normalizedToken = String(token || "").trim().toLowerCase();
      return BACKGROUND_OPTION_MAP.has(normalizedToken)
        ? normalizedToken
        : "";
    }

    function getBackgroundCss(token) {
      return BACKGROUND_OPTION_MAP.get(normalizeBackgroundToken(token))?.css || "";
    }

    function getNormalizedBackgrounds(backgrounds, columnCount) {
      const values = Array.isArray(backgrounds) ? backgrounds : [];
      const normalized = [];
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        normalized.push(normalizeBackgroundToken(values[columnIndex]));
      }
      return normalized;
    }

    function cloneRow(row) {
      if (isCommentRow(row)) {
        return {
          kind: "comment",
          text: String(row.text || ""),
          background: normalizeBackgroundToken(row.background),
        };
      }

      const values = Array.isArray(row?.values) ? row.values : [];
      return {
        kind: "record",
        values: values.map((value) => String(value || "")),
        backgrounds: getNormalizedBackgrounds(row?.backgrounds, values.length),
      };
    }

    function createEmptyRecordRow(columnCount) {
      return {
        kind: "record",
        values: new Array(columnCount).fill(""),
        backgrounds: new Array(columnCount).fill(""),
      };
    }

    function createCommentRow(text, background = "") {
      return {
        kind: "comment",
        text: String(text || ""),
        background: normalizeBackgroundToken(background),
      };
    }

    function ensureHeaderBackgroundLength(sheet) {
      if (!sheet) {
        return;
      }
      if (!Array.isArray(sheet.headerBackgrounds)) {
        sheet.headerBackgrounds = [];
      }
      while (sheet.headerBackgrounds.length < sheet.headers.length) {
        sheet.headerBackgrounds.push("");
      }
      if (sheet.headerBackgrounds.length > sheet.headers.length) {
        sheet.headerBackgrounds.length = sheet.headers.length;
      }
      for (let columnIndex = 0; columnIndex < sheet.headerBackgrounds.length; columnIndex += 1) {
        sheet.headerBackgrounds[columnIndex] = normalizeBackgroundToken(sheet.headerBackgrounds[columnIndex]);
      }
    }

    function cloneSheetModel(sheet, sheetIndex = 0) {
      const sourceSheet = sheet || {
        name: "Sheet " + (sheetIndex + 1),
        headers: ["Record", "Type"],
        rows: [],
      };
      return {
        name: getSheetDisplayName(sourceSheet, sheetIndex),
        headers: Array.isArray(sourceSheet.headers)
          ? sourceSheet.headers.map((header) => String(header || ""))
          : ["Record", "Type"],
        headerBackgrounds: getNormalizedBackgrounds(
          sourceSheet.headerBackgrounds,
          Array.isArray(sourceSheet.headers) ? sourceSheet.headers.length : 2,
        ),
        rows: Array.isArray(sourceSheet.rows)
          ? sourceSheet.rows.map((row) => cloneRow(row))
          : [],
      };
    }

    function buildWorkbookSnapshot() {
      const sheets = Array.isArray(state.workbook?.sheets)
        ? state.workbook.sheets
        : [];
      return {
        sheets: sheets.map((sheet, sheetIndex) => cloneSheetModel(sheet, sheetIndex)),
      };
    }

    function cloneColumnWidthState(widthState) {
      return {
        auto: Array.isArray(widthState?.auto) ? widthState.auto.slice() : [],
        manual: { ...(widthState?.manual || {}) },
      };
    }

    function buildColumnWidthStateSnapshot() {
      const widthStates = [];
      const sheetCount = state.workbook?.sheets?.length || 0;
      for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex += 1) {
        widthStates.push(cloneColumnWidthState(columnWidthStateBySheet.get(sheetIndex)));
      }
      return widthStates;
    }

    function buildHistorySnapshot() {
      return {
        workbook: buildWorkbookSnapshot(),
        activeSheetIndex,
        widthStates: buildColumnWidthStateSnapshot(),
      };
    }

    function getHistorySnapshotKey(snapshot) {
      return JSON.stringify({
        workbook: snapshot?.workbook || { sheets: [] },
        activeSheetIndex: Number.isInteger(snapshot?.activeSheetIndex) ? snapshot.activeSheetIndex : 0,
        widthStates: Array.isArray(snapshot?.widthStates) ? snapshot.widthStates : [],
      });
    }

    function pushHistoryEntry(stack, snapshot) {
      if (!snapshot) {
        return;
      }
      const nextKey = getHistorySnapshotKey(snapshot);
      const previousSnapshot = stack[stack.length - 1];
      if (previousSnapshot && getHistorySnapshotKey(previousSnapshot) === nextKey) {
        return;
      }
      stack.push(snapshot);
      if (stack.length > HISTORY_LIMIT) {
        stack.splice(0, stack.length - HISTORY_LIMIT);
      }
    }

    function recordUndoSnapshot() {
      pushHistoryEntry(undoHistory, buildHistorySnapshot());
      redoHistory = [];
      refreshToolbarState();
    }

    function restoreHistorySnapshot(snapshot) {
      if (!snapshot) {
        return;
      }
      state.workbook = {
        sheets: Array.isArray(snapshot.workbook?.sheets)
          ? snapshot.workbook.sheets.map((sheet, sheetIndex) => cloneSheetModel(sheet, sheetIndex))
          : [],
      };
      rebuildColumnWidthStateMap(
        Array.isArray(snapshot.widthStates)
          ? snapshot.widthStates.map((widthState) => cloneColumnWidthState(widthState))
          : [],
      );
      const sheetCount = state.workbook.sheets.length;
      activeSheetIndex = sheetCount === 0
        ? 0
        : Math.max(0, Math.min(snapshot.activeSheetIndex || 0, sheetCount - 1));
      editingSheetIndex = undefined;
      clearSheetScopedSelectionState();
      queueWorkbookSave();
      safeRender();
    }

    function resetHistoryState() {
      undoHistory = [];
      redoHistory = [];
      refreshToolbarState();
    }

    function undoWorkbookChange() {
      if (undoHistory.length === 0) {
        return;
      }
      pushHistoryEntry(redoHistory, buildHistorySnapshot());
      restoreHistorySnapshot(undoHistory.pop());
      refreshToolbarState();
    }

    function redoWorkbookChange() {
      if (redoHistory.length === 0) {
        return;
      }
      pushHistoryEntry(undoHistory, buildHistorySnapshot());
      restoreHistorySnapshot(redoHistory.pop());
      refreshToolbarState();
    }

    function ensureRecordRowLength(row, columnCount) {
      if (!row || isCommentRow(row)) {
        return;
      }

      if (!Array.isArray(row.values)) {
        row.values = [];
      }

      while (row.values.length < columnCount) {
        row.values.push("");
      }
      if (row.values.length > columnCount) {
        row.values.length = columnCount;
      }

      if (!Array.isArray(row.backgrounds)) {
        row.backgrounds = [];
      }
      while (row.backgrounds.length < columnCount) {
        row.backgrounds.push("");
      }
      if (row.backgrounds.length > columnCount) {
        row.backgrounds.length = columnCount;
      }
      for (let columnIndex = 0; columnIndex < row.backgrounds.length; columnIndex += 1) {
        row.backgrounds[columnIndex] = normalizeBackgroundToken(row.backgrounds[columnIndex]);
      }
    }

    function queueWorkbookSave() {
      workbookDirty = true;
      if (pendingSaveTimer) {
        window.clearTimeout(pendingSaveTimer);
      }
      pendingSaveTimer = window.setTimeout(pushWorkbookState, SAVE_DEBOUNCE_MS);
    }

    function pushWorkbookState() {
      if (pendingSaveTimer) {
        window.clearTimeout(pendingSaveTimer);
        pendingSaveTimer = undefined;
      }
      if (!workbookDirty) {
        return;
      }
      workbookDirty = false;
      vscode.postMessage({
        type: "applySpreadsheetWorkbook",
        workbook: buildWorkbookSnapshot(),
      });
    }

    function flushWorkbookSave() {
      pushWorkbookState();
    }

    function render() {
      const activeSheet = getActiveSheet();
      closeFloatingMenu();
      closeSuggestionMenu();
      hideValidationTooltip();
      clearReorderDropIndicator();
      document.body.classList.remove("reordering-row", "reordering-column");
      refreshToolbarState();
      renderSheetTabs();
      renderGrid(activeSheet);
      renderValidationBadge();
    }

    function persistUiState() {
      if (typeof vscode.setState === "function") {
        vscode.setState({
          ...persistedUiState,
          fontSizePx,
          densityMode,
        });
      }
    }

    function getDensityMetrics(mode = densityMode) {
      return DENSITY_METRICS[mode] || DENSITY_METRICS[DEFAULT_DENSITY_MODE];
    }

    function getRowIndexColumnWidth() {
      return getDensityMetrics().rowIndexColumnWidth;
    }

    function getMinColumnWidth() {
      return getDensityMetrics().minColumnWidth;
    }

    function getHeaderActionWidth() {
      const metrics = getDensityMetrics();
      return metrics.actionButtonWidth + metrics.columnResizerWidth + metrics.headerActionExtraWidth;
    }

    function getChoiceActionWidth() {
      return getDensityMetrics().choiceActionWidth;
    }

    function syncLayoutMetrics() {
      const metrics = getDensityMetrics();
      document.body.dataset.densityMode = densityMode;
      document.documentElement.style.setProperty("--spreadsheet-font-size", fontSizePx + "px");
      document.documentElement.style.setProperty(
        "--spreadsheet-row-min-height",
        Math.max(metrics.rowMinHeightFloor, Math.round(fontSizePx + metrics.rowMinHeightExtra)) + "px",
      );
      document.documentElement.style.setProperty("--spreadsheet-cell-padding-y", metrics.cellPaddingY + "px");
      document.documentElement.style.setProperty("--spreadsheet-cell-padding-x", metrics.cellPaddingX + "px");
      document.documentElement.style.setProperty("--spreadsheet-row-index-width", metrics.rowIndexColumnWidth + "px");
      document.documentElement.style.setProperty("--spreadsheet-choice-cell-min-width", metrics.choiceCellMinWidth + "px");
      document.documentElement.style.setProperty("--spreadsheet-comment-min-width", metrics.commentMinWidth + "px");
      document.documentElement.style.setProperty("--spreadsheet-action-button-width", metrics.actionButtonWidth + "px");
      document.documentElement.style.setProperty("--spreadsheet-column-resizer-width", metrics.columnResizerWidth + "px");
    }

    function applyFontSize(nextFontSizePx, options = {}) {
      const parsedFontSize = Number(nextFontSizePx);
      fontSizePx = Number.isFinite(parsedFontSize) && parsedFontSize > 0
        ? parsedFontSize
        : DEFAULT_FONT_SIZE_PX;
      syncLayoutMetrics();
      if (fontSizeSelect instanceof HTMLSelectElement) {
        fontSizeSelect.value = String(fontSizePx);
      }
      if (options.persist !== false) {
        persistUiState();
      }
      if (options.rerender) {
        safeRender();
      }
    }

    function applyDensityMode(nextDensityMode, options = {}) {
      densityMode = nextDensityMode === "compact" ? "compact" : DEFAULT_DENSITY_MODE;
      syncLayoutMetrics();
      if (densitySelect instanceof HTMLSelectElement) {
        densitySelect.value = densityMode;
      }
      if (options.persist !== false) {
        persistUiState();
      }
      if (options.rerender) {
        safeRender();
      }
    }

    function initializeFontSizeOptions() {
      if (!(fontSizeSelect instanceof HTMLSelectElement)) {
        return;
      }
      fontSizeSelect.innerHTML = "";
      for (let fontSize = 8; fontSize <= 36; fontSize += 1) {
        const option = document.createElement("option");
        option.value = String(fontSize);
        option.textContent = String(fontSize);
        fontSizeSelect.appendChild(option);
      }
    }

    function getClientErrorMessage(error) {
      return error instanceof Error ? error.message : String(error || "Unknown error");
    }

    function applyRenderWarningBadge() {
      if (!renderWarningMessage) {
        return;
      }
      validationBadge.classList.add("has-issues");
      validationBadge.title = renderWarningMessage + (validationBadge.title ? "\\n\\n" + validationBadge.title : "");
    }

    function disableCellSelectionEnhancements(error) {
      if (!cellSelectionEnhancementsEnabled) {
        return;
      }
      console.error("Spreadsheet cell selection enhancements disabled:", error);
      cellSelectionEnhancementsEnabled = false;
      renderWarningMessage = "Cell selection features were disabled: " + getClientErrorMessage(error);
      selectedCellRange = undefined;
      cellSelectionAnchor = undefined;
      isSelectingCells = false;
    }

    function safeRender() {
      try {
        render();
        applyRenderWarningBadge();
        return;
      } catch (error) {
        if (cellSelectionEnhancementsEnabled) {
          disableCellSelectionEnhancements(error);
          try {
            render();
            applyRenderWarningBadge();
            return;
          } catch (recoveryError) {
            console.error("Spreadsheet recovery render failed:", recoveryError);
            error = recoveryError;
          }
        } else {
          console.error("Spreadsheet render failed:", error);
        }
      }

      validationBadge.className = "validation-badge has-issues";
      validationBadge.textContent = "Render error";
      validationBadge.title = renderWarningMessage || "The spreadsheet failed to render.";
    }

    function refreshToolbarState() {
      const isDatabaseSource = state.sourceKind === "database";
      const displayKind = isDatabaseSource ? "Database" : "Excel";
      const activeSheetIssueCount = (state.validation?.issues || []).filter(
        (issue) => issue.sheetIndex === activeSheetIndex,
      ).length;
      const cannotExportHint = activeSheetIssueCount > 0
        ? "There " + (activeSheetIssueCount === 1 ? "is 1 error" : "are " + activeSheetIssueCount + " errors") +
          " in the current sheet, cannot preview/export as DB."
        : "";
      toolbarTitle.textContent = state.displayFileName + " (" + displayKind + ")";
      previewDbButton.textContent = isDatabaseSource ? "Show DB File" : "Preview Sheet As DB";
      saveButton.textContent = isDatabaseSource ? "Save Database" : "Save Spreadsheet";
      saveAsButton.textContent = isDatabaseSource ? "Save As Database" : "Save As Spreadsheet";
      saveDbButton.textContent = isDatabaseSource ? "Export As Spreadsheet" : "Export Sheet As DB";
      previewDbButton.classList.toggle("is-disabled", !isDatabaseSource && activeSheetIssueCount > 0);
      previewDbButton.setAttribute("aria-disabled", !isDatabaseSource && activeSheetIssueCount > 0 ? "true" : "false");
      if (isDatabaseSource) {
        previewDbButton.title = "Show the associated database file.";
        delete previewDbButton.dataset.hoverHint;
      } else if (activeSheetIssueCount === 0) {
        previewDbButton.title = "Preview the current sheet as EPICS database text.";
        delete previewDbButton.dataset.hoverHint;
      } else {
        previewDbButton.removeAttribute("title");
        previewDbButton.dataset.hoverHint = cannotExportHint;
      }
      saveDbButton.classList.toggle("is-disabled", !isDatabaseSource && activeSheetIssueCount > 0);
      saveDbButton.setAttribute("aria-disabled", !isDatabaseSource && activeSheetIssueCount > 0 ? "true" : "false");
      if (isDatabaseSource || activeSheetIssueCount === 0) {
        saveDbButton.title = isDatabaseSource
          ? "Export the current database spreadsheet as an Excel workbook."
          : "Export the current sheet as an EPICS database file.";
        delete saveDbButton.dataset.hoverHint;
      } else {
        saveDbButton.removeAttribute("title");
        saveDbButton.dataset.hoverHint = cannotExportHint;
      }
      saveButton.title = isDatabaseSource
        ? "Save to the associated database file."
        : state.canSaveToCurrentFile
          ? "Save to the associated Excel file."
          : "Save to a new Excel file.";
      saveAsButton.title = isDatabaseSource
        ? "Save to a new database file."
        : "Save to a new Excel file and associate this spreadsheet.";
      undoButton.disabled = undoHistory.length === 0;
      redoButton.disabled = redoHistory.length === 0;
      undoButton.title = undoHistory.length === 0
        ? "Nothing to undo."
        : "Undo the last spreadsheet change.";
      redoButton.title = redoHistory.length === 0
        ? "Nothing to redo."
        : "Redo the last undone spreadsheet change.";
      refreshBackgroundPaletteState();
    }

    function getSheetDisplayName(sheet, sheetIndex) {
      return String(sheet?.name || "").trim() || ("Sheet " + (sheetIndex + 1));
    }

    function normalizeSheetNameInput(value, fallbackName = "Sheet") {
      const sanitized = String(value || "")
        .replace(/[\\\\/*?:\\[\\]]/g, " ")
        .replace(/\\s+/g, " ")
        .trim()
        .slice(0, 31);
      return sanitized || fallbackName;
    }

    function getUniqueSheetName(proposedName, ignoreSheetIndex = -1) {
      const sheets = state.workbook?.sheets || [];
      const existingNames = new Set(
        sheets
          .map((sheet, sheetIndex) => sheetIndex === ignoreSheetIndex ? "" : getSheetDisplayName(sheet, sheetIndex))
          .filter(Boolean),
      );
      const baseName = normalizeSheetNameInput(proposedName, "Sheet");
      if (!existingNames.has(baseName)) {
        return baseName;
      }

      let suffix = 2;
      while (true) {
        const suffixText = " " + suffix;
        const candidateBase = baseName.slice(0, Math.max(1, 31 - suffixText.length)).trim() || "Sheet";
        const candidate = candidateBase + suffixText;
        if (!existingNames.has(candidate)) {
          return candidate;
        }
        suffix += 1;
      }
    }

    function createDefaultSheetModel(sheetName) {
      return {
        name: getUniqueSheetName(sheetName),
        headers: ["Record", "Type", "INP"],
        rows: [createEmptyRecordRow(3)],
      };
    }

    function addSheet() {
      recordUndoSnapshot();
      if (!state.workbook || !Array.isArray(state.workbook.sheets)) {
        state.workbook = { sheets: [] };
      }
      const nextSheetName = getUniqueSheetName("Sheet " + (state.workbook.sheets.length + 1));
      state.workbook.sheets.push(createDefaultSheetModel(nextSheetName));
      activeSheetIndex = state.workbook.sheets.length - 1;
      queueWorkbookSave();
      safeRender();
    }

    function startSheetRename(sheetIndex) {
      const sheets = state.workbook?.sheets || [];
      if (!sheets[sheetIndex]) {
        return;
      }
      editingSheetIndex = sheetIndex;
      safeRender();
    }

    function finishSheetRename(sheetIndex, requestedName, cancel = false) {
      const sheets = state.workbook?.sheets || [];
      const sheet = sheets[sheetIndex];
      editingSheetIndex = undefined;
      if (!sheet) {
        safeRender();
        return;
      }
      if (cancel) {
        safeRender();
        return;
      }
      const currentName = getSheetDisplayName(sheet, sheetIndex);
      const nextName = getUniqueSheetName(
        normalizeSheetNameInput(requestedName, currentName),
        sheetIndex,
      );
      if (nextName === currentName) {
        safeRender();
        return;
      }
      recordUndoSnapshot();
      sheet.name = nextName;
      queueWorkbookSave();
      safeRender();
    }

    function clearSheetScopedSelectionState() {
      selectedRowIndexes = [];
      rowSelectionAnchorIndex = undefined;
      rowSelectionPivotIndex = undefined;
      isSelectingRows = false;
      selectedColumnIndexes = [];
      columnSelectionAnchorIndex = undefined;
      columnSelectionPivotIndex = undefined;
      isSelectingColumns = false;
      selectedCellRange = undefined;
      cellSelectionAnchor = undefined;
      isSelectingCells = false;
    }

    function rebuildColumnWidthStateMap(widthStates) {
      columnWidthStateBySheet.clear();
      widthStates.forEach((widthState, index) => {
        if (widthState) {
          columnWidthStateBySheet.set(index, widthState);
        }
      });
    }

    function removeSheetWidthState(sheetIndex, previousSheetCount) {
      const sheetCount = Number.isInteger(previousSheetCount)
        ? previousSheetCount
        : (state.workbook?.sheets?.length || 0);
      const widthStates = [];
      for (let index = 0; index < sheetCount; index += 1) {
        if (index === sheetIndex) {
          continue;
        }
        widthStates.push(columnWidthStateBySheet.get(index));
      }
      rebuildColumnWidthStateMap(widthStates);
    }

    function removeSheet(sheetIndex) {
      const sheets = state.workbook?.sheets || [];
      const previousSheetCount = sheets.length;
      if (sheetIndex < 0 || sheetIndex >= sheets.length) {
        return;
      }
      if (sheets.length <= 1) {
        recordUndoSnapshot();
        sheets.splice(0, sheets.length, createDefaultSheetModel("Sheet 1"));
        columnWidthStateBySheet.clear();
        activeSheetIndex = 0;
        editingSheetIndex = undefined;
        clearSheetScopedSelectionState();
        queueWorkbookSave();
        safeRender();
        return;
      }
      recordUndoSnapshot();
      sheets.splice(sheetIndex, 1);
      removeSheetWidthState(sheetIndex, previousSheetCount);
      if (activeSheetIndex > sheetIndex) {
        activeSheetIndex -= 1;
      } else if (activeSheetIndex >= sheets.length) {
        activeSheetIndex = sheets.length - 1;
      }
      if (typeof editingSheetIndex === "number") {
        if (editingSheetIndex === sheetIndex) {
          editingSheetIndex = undefined;
        } else if (editingSheetIndex > sheetIndex) {
          editingSheetIndex -= 1;
        }
      }
      clearSheetScopedSelectionState();
      queueWorkbookSave();
      safeRender();
    }

    function renderSheetTabs() {
      sheetTabs.innerHTML = "";
      if (state.sourceKind === "database") {
        sheetTabs.hidden = true;
        return;
      }
      sheetTabs.hidden = false;
      (state.workbook?.sheets || []).forEach((sheet, index) => {
        if (editingSheetIndex === index) {
          const input = document.createElement("input");
          input.type = "text";
          input.className = "sheet-tab-input";
          input.value = getSheetDisplayName(sheet, index);
          input.dataset.sheetIndex = String(index);
          input.addEventListener("click", (event) => {
            event.stopPropagation();
          });
          input.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              finishSheetRename(index, input.value);
            } else if (event.key === "Escape") {
              event.preventDefault();
              finishSheetRename(index, input.value, true);
            }
          });
          input.addEventListener("blur", () => {
            finishSheetRename(index, input.value);
          });
          sheetTabs.appendChild(input);
          window.requestAnimationFrame(() => {
            input.focus();
            input.select();
          });
          return;
        }

        const button = document.createElement("button");
        button.type = "button";
        button.className =
          "sheet-tab draggable" + (index === activeSheetIndex ? " active" : "");
        button.draggable = true;
        button.textContent = getSheetDisplayName(sheet, index);
        button.title = "Double-click to modify the sheet title.";
        button.addEventListener("click", () => {
          activeSheetIndex = index;
          safeRender();
        });
        button.addEventListener("dragstart", (event) => {
          startSheetReorderDrag(event, index);
        });
        button.addEventListener("dragover", (event) => {
          if (!sheetReorderDragState) {
            return;
          }
          event.preventDefault();
          if (event.dataTransfer) {
            event.dataTransfer.dropEffect = "move";
          }
          const insertIndex = getSheetDropInsertIndex(index, event.clientX, button);
          setReorderDropIndicator(button, insertIndex > index ? "after" : "before");
        });
        button.addEventListener("drop", (event) => {
          if (!sheetReorderDragState) {
            return;
          }
          event.preventDefault();
          const dragState = sheetReorderDragState;
          const insertIndex = getSheetDropInsertIndex(index, event.clientX, button);
          finishReorderDrag();
          moveSheetsToInsertIndex([dragState.sheetIndex], insertIndex);
        });
        button.addEventListener("dragend", () => {
          finishReorderDrag();
        });
        button.addEventListener("dblclick", (event) => {
          event.preventDefault();
          event.stopPropagation();
          startSheetRename(index);
        });
        button.addEventListener("contextmenu", (event) => {
          event.preventDefault();
          event.stopPropagation();
          openFloatingMenuAtPosition(event.clientX, event.clientY, [
            {
              label: "Rename sheet",
              action: () => startSheetRename(index),
            },
            {
              label: "Add sheet",
              action: () => addSheet(),
            },
            {
              label: "Remove sheet",
              disabled: (state.workbook?.sheets?.length || 0) <= 1,
              action: () => openSheetRemovalConfirmationAtPosition(event.clientX, event.clientY, index),
            },
          ]);
        });
        sheetTabs.appendChild(button);
      });

      const addSheetButton = document.createElement("button");
      addSheetButton.type = "button";
      addSheetButton.className = "sheet-tab add-sheet";
      addSheetButton.textContent = "+";
      addSheetButton.title = "Add sheet";
      addSheetButton.addEventListener("click", () => {
        addSheet();
      });
      sheetTabs.appendChild(addSheetButton);
    }

    function renderValidationBadge() {
      const activeIssues = (state.validation?.issues || []).filter(
        (issue) => issue.sheetIndex === activeSheetIndex,
      );
      validationBadge.className = "validation-badge" + (activeIssues.length > 0 ? " has-issues" : "");
      validationBadge.textContent = activeIssues.length === 0
        ? "No error"
        : activeIssues.length + " issue" + (activeIssues.length === 1 ? "" : "s");
      validationBadge.title = activeIssues.length === 0
        ? "No validation issues in this sheet."
        : activeIssues
          .map((issue) => {
            return issue.rowIndex < 0
              ? "Column " + getColumnLabel(issue.columnIndex) + ": " + issue.message
              : "Row " + getDisplayRowLabel(getActiveSheet(), issue.rowIndex) + ", Column " + getColumnLabel(issue.columnIndex) + ": " + issue.message;
          })
          .join("\\n");
    }

    function applyCustomBackgroundStyle(element, backgroundToken) {
      if (!(element instanceof HTMLElement)) {
        return;
      }
      const cssColor = getBackgroundCss(backgroundToken);
      element.classList.toggle("has-custom-background", !!cssColor);
      if (cssColor) {
        element.style.setProperty("--spreadsheet-custom-background", cssColor);
      } else {
        element.style.removeProperty("--spreadsheet-custom-background");
      }
    }

    function renderGrid(sheet) {
      gridHead.innerHTML = "";
      gridBody.innerHTML = "";
      ensureHeaderBackgroundLength(sheet);
      recomputeAutoColumnWidths(sheet);
      renderColumnGroup(sheet);

      const indexHeaderRow = document.createElement("tr");
      const fieldHeaderRow = document.createElement("tr");
      const topLeft = document.createElement("th");
      topLeft.className = "row-index header-corner";
      topLeft.rowSpan = 2;
      topLeft.textContent = "";
      indexHeaderRow.appendChild(topLeft);

      sheet.headers.forEach((header, columnIndex) => {
        const indexCell = document.createElement("th");
        indexCell.className = "column-index-cell";
        indexCell.dataset.columnIndex = String(columnIndex);
        indexCell.addEventListener("contextmenu", (event) => {
          event.preventDefault();
          event.stopPropagation();
          ensureColumnSelection(columnIndex);
          openColumnMenuAtPosition(event.clientX, event.clientY, columnIndex);
        });

        const indexWrapper = document.createElement("div");
        indexWrapper.className = "column-index-wrap";
        const columnButton = document.createElement("button");
        columnButton.type = "button";
        columnButton.className = "column-select-button";
        columnButton.textContent = getSpreadsheetColumnIndexLabel(columnIndex);
        if (isDraggableColumnIndex(columnIndex)) {
          indexWrapper.classList.add("drag-handle");
          columnButton.draggable = true;
          columnButton.addEventListener("dragstart", (event) => {
            startColumnReorderDrag(event, columnIndex);
          });
          columnButton.addEventListener("dragend", () => {
            finishReorderDrag();
          });
        }
        columnButton.addEventListener("mousedown", (event) => {
          if (event.button !== 0) {
            return;
          }
          closeFloatingMenu();
          const anchorIndex = event.shiftKey
            ? getColumnSelectionPivot(columnIndex)
            : columnIndex;
          columnSelectionAnchorIndex = anchorIndex;
          columnSelectionPivotIndex = anchorIndex;
          if (event.shiftKey) {
            isSelectingColumns = true;
            setSelectedColumnRange(anchorIndex, columnIndex);
            event.preventDefault();
            return;
          }
          if (selectedColumnIndexes.includes(columnIndex)) {
            isSelectingColumns = false;
            return;
          }
          isSelectingColumns = true;
          setSelectedColumnRange(columnIndex, columnIndex);
          event.preventDefault();
        });
        indexCell.addEventListener("mouseenter", () => {
          if (!isSelectingColumns || typeof columnSelectionAnchorIndex !== "number") {
            return;
          }
          setSelectedColumnRange(columnSelectionAnchorIndex, columnIndex);
        });
        indexWrapper.appendChild(columnButton);

        const menuButton = document.createElement("button");
        menuButton.type = "button";
        menuButton.className = "menu-trigger";
        menuButton.draggable = false;
        menuButton.textContent = "...";
        menuButton.title = "Column actions";
        menuButton.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          ensureColumnSelection(columnIndex);
          openColumnMenu(menuButton, columnIndex);
        });
        indexWrapper.appendChild(menuButton);

        const resizeHandle = document.createElement("div");
        resizeHandle.className = "column-resizer";
        resizeHandle.draggable = false;
        resizeHandle.title = "Resize column";
        resizeHandle.addEventListener("mousedown", (event) => {
          event.preventDefault();
          event.stopPropagation();
          startColumnResize(columnIndex, event.clientX);
        });
        indexWrapper.appendChild(resizeHandle);

        if (isDraggableColumnIndex(columnIndex)) {
          indexCell.addEventListener("dragover", (event) => {
            if (!columnReorderDragState || columnReorderDragState.sheetIndex !== activeSheetIndex) {
              return;
            }
            event.preventDefault();
            if (event.dataTransfer) {
              event.dataTransfer.dropEffect = "move";
            }
            const insertIndex = getColumnDropInsertIndex(columnIndex, event.clientX, indexCell);
            setReorderDropIndicator(
              indexCell,
              insertIndex > columnIndex ? "after" : "before",
            );
          });
          indexCell.addEventListener("drop", (event) => {
            if (!columnReorderDragState || columnReorderDragState.sheetIndex !== activeSheetIndex) {
              return;
            }
            event.preventDefault();
            const dragState = columnReorderDragState;
            const insertIndex = getColumnDropInsertIndex(columnIndex, event.clientX, indexCell);
            finishReorderDrag();
            moveColumnsToInsertIndex(dragState.columnIndexes, insertIndex);
          });
        }
        indexCell.appendChild(indexWrapper);
        applyCustomBackgroundStyle(indexCell, sheet.headerBackgrounds[columnIndex]);
        indexHeaderRow.appendChild(indexCell);

        const cell = document.createElement("th");
        cell.className = "field-header-cell";
        cell.dataset.columnIndex = String(columnIndex);
        cell.dataset.validationRowIndex = "-1";
        cell.dataset.validationColumnIndex = String(columnIndex);

        const wrapper = document.createElement("div");
        wrapper.className = "header-wrap";
        const input = document.createElement("input");
        input.className =
          "header-cell" +
          (columnIndex === 0 ? " record" : columnIndex === 1 ? " type" : "");
        input.dataset.scope = "header";
        input.dataset.columnIndex = String(columnIndex);
        input.draggable = false;
        input.value = header || "";
        input.addEventListener("mousedown", (event) => {
          if (event.button !== 0) {
            return;
          }
          clearColumnSelection();
        });
        if (columnIndex < 2) {
          input.readOnly = true;
        } else {
          input.autocomplete = "off";
          input.spellcheck = false;
          input.addEventListener("blur", () => {
            scheduleSuggestionMenuClose();
          });
          input.addEventListener("keydown", (event) => {
            handleSuggestionKeydown(event, input);
          });
          input.addEventListener("input", () => {
            updateHeaderFieldValue(columnIndex, input.value);
            openSuggestionMenu(input);
          });
        }
        input.addEventListener("keydown", handleSpreadsheetNavigationKeydown);
        wrapper.appendChild(input);

        cell.appendChild(wrapper);
        applyCustomBackgroundStyle(cell, sheet.headerBackgrounds[columnIndex]);
        applyIssueClass(cell, -1, columnIndex);
        fieldHeaderRow.appendChild(cell);
      });
      gridHead.append(indexHeaderRow, fieldHeaderRow);

      let recordDisplayNumber = 0;
      sheet.rows.forEach((row, rowIndex) => {
        const tr = document.createElement("tr");
        tr.className = "sheet-row";
        tr.dataset.rowIndex = String(rowIndex);
        const indexCell = document.createElement("td");
        indexCell.className = "row-index row-index-cell";
        indexCell.dataset.rowIndex = String(rowIndex);
        indexCell.addEventListener("contextmenu", (event) => {
          event.preventDefault();
          event.stopPropagation();
          ensureRowSelection(rowIndex);
          openRowMenuAtPosition(event.clientX, event.clientY, rowIndex);
        });

        const rowActions = document.createElement("div");
        rowActions.className = "row-index-wrap";
        const rowLabel = document.createElement("button");
        rowLabel.type = "button";
        rowLabel.className = "row-select-button";
        rowLabel.classList.add("drag-handle");
        rowLabel.draggable = true;
        rowLabel.textContent = getDisplayRowLabel(sheet, rowIndex, recordDisplayNumber);
        if (!isCommentRow(row)) {
          recordDisplayNumber += 1;
        }
        rowLabel.addEventListener("mousedown", (event) => {
          if (event.button !== 0) {
            return;
          }
          closeFloatingMenu();
          clearColumnSelection();
          const anchorIndex = event.shiftKey
            ? getRowSelectionPivot(rowIndex)
            : rowIndex;
          rowSelectionAnchorIndex = anchorIndex;
          rowSelectionPivotIndex = anchorIndex;
          if (event.shiftKey) {
            isSelectingRows = true;
            setSelectedRowRange(anchorIndex, rowIndex);
            event.preventDefault();
            return;
          }
          if (selectedRowIndexes.includes(rowIndex)) {
            isSelectingRows = false;
            return;
          }
          isSelectingRows = true;
          setSelectedRowRange(rowIndex, rowIndex);
          event.preventDefault();
        });
        rowLabel.addEventListener("dragstart", (event) => {
          startRowReorderDrag(event, rowIndex);
        });
        rowLabel.addEventListener("dragover", (event) => {
          if (!rowReorderDragState || rowReorderDragState.sheetIndex !== activeSheetIndex) {
            return;
          }
          event.preventDefault();
          if (event.dataTransfer) {
            event.dataTransfer.dropEffect = "move";
          }
          const insertIndex = getRowDropInsertIndex(rowIndex, event.clientY, rowLabel);
          setReorderDropIndicator(
            indexCell,
            insertIndex > rowIndex ? "after" : "before",
          );
        });
        rowLabel.addEventListener("drop", (event) => {
          if (!rowReorderDragState || rowReorderDragState.sheetIndex !== activeSheetIndex) {
            return;
          }
          event.preventDefault();
          const dragState = rowReorderDragState;
          const insertIndex = getRowDropInsertIndex(rowIndex, event.clientY, rowLabel);
          finishReorderDrag();
          moveRowsToInsertIndex(dragState.rowIndexes, insertIndex);
        });
        rowLabel.addEventListener("dragend", () => {
          finishReorderDrag();
        });
        rowLabel.addEventListener("mouseenter", () => {
          if (!isSelectingRows || typeof rowSelectionAnchorIndex !== "number") {
            return;
          }
          setSelectedRowRange(rowSelectionAnchorIndex, rowIndex);
        });
        const rowMenuButton = document.createElement("button");
        rowMenuButton.type = "button";
        rowMenuButton.className = "menu-trigger";
        rowMenuButton.textContent = "...";
        rowMenuButton.title = "Row actions";
        rowMenuButton.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          ensureRowSelection(rowIndex);
          openRowMenu(rowMenuButton, rowIndex);
        });
        rowActions.append(rowLabel, rowMenuButton);
        indexCell.appendChild(rowActions);
        tr.appendChild(indexCell);

        if (isCommentRow(row)) {
          const td = document.createElement("td");
          td.className = "comment-cell";
          td.colSpan = Math.max(1, sheet.headers.length);
          applyCustomBackgroundStyle(td, row.background);
          const textarea = document.createElement("textarea");
          textarea.className = "comment-input";
          textarea.rows = 1;
          textarea.dataset.scope = "comment";
          textarea.dataset.rowIndex = String(rowIndex);
          textarea.dataset.columnIndex = "0";
          textarea.value = row.text || "";
          textarea.addEventListener("mousedown", (event) => {
            if (event.button !== 0) {
              return;
            }
            clearRowSelection();
            clearColumnSelection();
            clearCellSelection();
          });
          textarea.addEventListener("input", () => {
            if (row.text !== textarea.value) {
              recordUndoSnapshot();
            }
            row.text = textarea.value;
            queueWorkbookSave();
          });
          textarea.addEventListener("keydown", handleSpreadsheetNavigationKeydown);
          td.appendChild(textarea);
          tr.appendChild(td);
        } else {
          ensureRecordRowLength(row, sheet.headers.length);
          sheet.headers.forEach((_, columnIndex) => {
            const td = document.createElement("td");
            td.dataset.rowIndex = String(rowIndex);
            td.dataset.columnIndex = String(columnIndex);
            td.dataset.validationRowIndex = String(rowIndex);
            td.dataset.validationColumnIndex = String(columnIndex);
            applyCustomBackgroundStyle(td, row.backgrounds?.[columnIndex]);
            td.addEventListener("mousedown", (event) => {
              handleCellMouseDown(event, rowIndex, columnIndex);
            });
            if (cellSelectionEnhancementsEnabled) {
              td.addEventListener("mouseenter", () => {
                handleCellMouseEnter(rowIndex, columnIndex);
              });
              td.addEventListener("contextmenu", (event) => {
                handleCellContextMenu(event, rowIndex, columnIndex);
              });
            }
            const editor = document.createElement("div");
            editor.className = "grid-editor";
            const input = document.createElement("input");
            input.className =
              "grid-cell" +
              (columnIndex === 0 ? " record" : columnIndex === 1 ? " type" : "");
            input.dataset.scope = "cell";
            input.dataset.rowIndex = String(rowIndex);
            input.dataset.columnIndex = String(columnIndex);
            input.value = row.values[columnIndex] || "";
            if (columnIndex === 1) {
              input.autocomplete = "off";
              input.spellcheck = false;
              input.addEventListener("click", () => {
                openSuggestionMenu(input);
              });
              input.addEventListener("blur", () => {
                scheduleSuggestionMenuClose();
              });
              input.addEventListener("keydown", (event) => {
                handleSuggestionKeydown(event, input);
              });
            }
            input.addEventListener("input", () => {
              if (row.values[columnIndex] !== input.value) {
                recordUndoSnapshot();
              }
              row.values[columnIndex] = input.value;
              if (columnIndex === 1) {
                refreshChoiceTriggersForRow(rowIndex);
                openSuggestionMenu(input);
              }
              refreshAutoColumnWidth(columnIndex);
              queueWorkbookSave();
            });
            input.addEventListener("keydown", handleSpreadsheetNavigationKeydown);
            editor.appendChild(input);

            const choices = getCellMenuChoices(sheet, rowIndex, columnIndex);
            if (choices.length > 0) {
              editor.classList.add("has-choice");
              editor.appendChild(createChoiceTriggerButton(rowIndex, columnIndex));
            }

            td.appendChild(editor);
            applyIssueClass(td, rowIndex, columnIndex);
            tr.appendChild(td);
          });
        }
        gridBody.appendChild(tr);
      });

      refreshColumnSelectionStyles();
      if (cellSelectionEnhancementsEnabled) {
        refreshCellSelectionStyles();
      }
      refreshRowSelectionStyles();
    }

    function getTextMeasureContext() {
      if (!textMeasureCanvas) {
        textMeasureCanvas = document.createElement("canvas");
      }
      const context = textMeasureCanvas.getContext("2d");
      if (!context) {
        return undefined;
      }
      const measureSource = gridBody.closest(".grid-panel") || gridBody || document.body;
      const measureStyle = window.getComputedStyle(measureSource);
      context.font = measureStyle.font || (measureStyle.fontSize + " " + measureStyle.fontFamily);
      return context;
    }

    function measureTextWidth(value, options = {}) {
      const context = getTextMeasureContext();
      if (!context) {
        return String(value || "").length * 8;
      }
      if (options.fontWeight) {
        const measureSource = gridBody.closest(".grid-panel") || gridBody || document.body;
        const measureStyle = window.getComputedStyle(measureSource);
        context.font = options.fontWeight + " " + (measureStyle.fontSize + " " + measureStyle.fontFamily);
      }
      return context.measureText(String(value || "")).width;
    }

    function getColumnMaxAutoWidth() {
      const metrics = getDensityMetrics();
      return Math.ceil(
        measureTextWidth("W".repeat(metrics.maxAutoColumnCharacters), { fontWeight: "600" }),
      ) + metrics.cellHorizontalPadding + getHeaderActionWidth();
    }

    function getColumnValues(sheet, columnIndex) {
      const values = [];
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        values.push(row.values[columnIndex] || "");
      });
      return values;
    }

    function getAutoColumnWidth(sheet, columnIndex) {
      const metrics = getDensityMetrics();
      const values = getColumnValues(sheet, columnIndex);
      let measuredWidth = Math.ceil(
        measureTextWidth(sheet.headers[columnIndex] || "", { fontWeight: "600" }),
      );
      values.forEach((value) => {
        measuredWidth = Math.max(measuredWidth, measureTextWidth(value));
      });

      let chromeWidth = metrics.cellHorizontalPadding + getHeaderActionWidth();
      if (columnIndex >= 2) {
        chromeWidth = Math.max(chromeWidth, metrics.cellHorizontalPadding + getChoiceActionWidth());
      }

      return Math.max(
        getMinColumnWidth(),
        Math.min(getColumnMaxAutoWidth(), Math.ceil(measuredWidth) + chromeWidth),
      );
    }

    function recomputeAutoColumnWidths(sheet) {
      const widthState = getColumnWidthState();
      widthState.auto = sheet.headers.map((_, columnIndex) => getAutoColumnWidth(sheet, columnIndex));
    }

    function getResolvedColumnWidth(columnIndex) {
      const widthState = getColumnWidthState();
      if (Object.prototype.hasOwnProperty.call(widthState.manual, columnIndex)) {
        return widthState.manual[columnIndex];
      }
      return widthState.auto[columnIndex] || getMinColumnWidth();
    }

    function renderColumnGroup(sheet) {
      gridColumns.innerHTML = "";

      const rowIndexColumn = document.createElement("col");
      rowIndexColumn.style.width = getRowIndexColumnWidth() + "px";
      gridColumns.appendChild(rowIndexColumn);

      sheet.headers.forEach((_, columnIndex) => {
        const col = document.createElement("col");
        col.dataset.columnIndex = String(columnIndex);
        col.style.width = getResolvedColumnWidth(columnIndex) + "px";
        gridColumns.appendChild(col);
      });
    }

    function applyColumnWidth(columnIndex) {
      const column = gridColumns.querySelector('col[data-column-index="' + columnIndex + '"]');
      if (!(column instanceof HTMLTableColElement)) {
        return;
      }
      column.style.width = getResolvedColumnWidth(columnIndex) + "px";
    }

    function refreshAutoColumnWidth(columnIndex) {
      const sheet = getActiveSheet();
      if (columnIndex < 0 || columnIndex >= sheet.headers.length) {
        return;
      }
      const widthState = getColumnWidthState();
      widthState.auto[columnIndex] = getAutoColumnWidth(sheet, columnIndex);
      if (!Object.prototype.hasOwnProperty.call(widthState.manual, columnIndex)) {
        applyColumnWidth(columnIndex);
      }
    }

    function resizeAllAutoColumns() {
      const sheet = getActiveSheet();
      const widthState = getColumnWidthState();
      widthState.manual = {};
      recomputeAutoColumnWidths(sheet);
      renderColumnGroup(sheet);
    }

    function shiftManualColumnWidthsForInsert(index, count) {
      const widthState = getColumnWidthState();
      const nextManualWidths = {};
      Object.entries(widthState.manual).forEach(([key, value]) => {
        const columnIndex = Number(key);
        nextManualWidths[columnIndex >= index ? columnIndex + count : columnIndex] = value;
      });
      widthState.manual = nextManualWidths;
    }

    function shiftManualColumnWidthsForDelete(index) {
      const widthState = getColumnWidthState();
      const nextManualWidths = {};
      Object.entries(widthState.manual).forEach(([key, value]) => {
        const columnIndex = Number(key);
        if (columnIndex === index) {
          return;
        }
        nextManualWidths[columnIndex > index ? columnIndex - 1 : columnIndex] = value;
      });
      widthState.manual = nextManualWidths;
    }

    function moveManualColumnWidth(columnIndex, targetIndex) {
      const widthState = getColumnWidthState();
      const orderedManualWidths = {};
      Object.entries(widthState.manual).forEach(([key, value]) => {
        let nextIndex = Number(key);
        if (nextIndex === columnIndex) {
          nextIndex = targetIndex;
        } else if (targetIndex < columnIndex && nextIndex >= targetIndex && nextIndex < columnIndex) {
          nextIndex += 1;
        } else if (targetIndex > columnIndex && nextIndex <= targetIndex && nextIndex > columnIndex) {
          nextIndex -= 1;
        }
        orderedManualWidths[nextIndex] = value;
      });
      widthState.manual = orderedManualWidths;
    }

    function startColumnResize(columnIndex, startClientX) {
      closeFloatingMenu();
      closeSuggestionMenu();
      columnResizeState = {
        sheetIndex: activeSheetIndex,
        columnIndex,
        startClientX,
        startWidth: getResolvedColumnWidth(columnIndex),
      };
      document.body.classList.add("column-resizing");
    }

    function updateColumnResize(clientX) {
      if (!columnResizeState || columnResizeState.sheetIndex !== activeSheetIndex) {
        return;
      }
      const widthState = getColumnWidthState();
      widthState.manual[columnResizeState.columnIndex] = Math.max(
        getMinColumnWidth(),
        Math.round(columnResizeState.startWidth + clientX - columnResizeState.startClientX),
      );
      applyColumnWidth(columnResizeState.columnIndex);
    }

    function stopColumnResize() {
      if (!columnResizeState) {
        return;
      }
      columnResizeState = undefined;
      document.body.classList.remove("column-resizing");
    }

    function clearReorderDropIndicator() {
      if (!reorderDropIndicatorState?.element) {
        reorderDropIndicatorState = undefined;
        return;
      }
      reorderDropIndicatorState.element.classList.remove("drop-before", "drop-after");
      reorderDropIndicatorState = undefined;
    }

    function setReorderDropIndicator(element, position) {
      if (!(element instanceof HTMLElement)) {
        clearReorderDropIndicator();
        return;
      }
      if (
        reorderDropIndicatorState?.element === element &&
        reorderDropIndicatorState.position === position
      ) {
        return;
      }
      clearReorderDropIndicator();
      element.classList.add(position === "after" ? "drop-after" : "drop-before");
      reorderDropIndicatorState = {
        element,
        position,
      };
    }

    function isDraggableColumnIndex(columnIndex) {
      return columnIndex >= 2;
    }

    function reorderIndexedBlock(items, indexes, insertIndex, minimumInsertIndex = 0) {
      const sortedIndexes = [...new Set((indexes || []).filter((index) =>
        Number.isInteger(index) && index >= minimumInsertIndex && index < items.length,
      ))].sort((left, right) => left - right);
      if (sortedIndexes.length === 0) {
        return undefined;
      }

      const boundedInsertIndex = Math.max(
        minimumInsertIndex,
        Math.min(Number.isInteger(insertIndex) ? insertIndex : items.length, items.length),
      );
      const blockSet = new Set(sortedIndexes);
      const block = sortedIndexes.map((index) => items[index]);
      const remaining = items.filter((_, index) => !blockSet.has(index));
      let nextInsertIndex = 0;
      for (let index = 0; index < Math.min(boundedInsertIndex, items.length); index += 1) {
        if (!blockSet.has(index)) {
          nextInsertIndex += 1;
        }
      }
      nextInsertIndex = Math.max(
        minimumInsertIndex,
        Math.min(nextInsertIndex, remaining.length),
      );
      remaining.splice(nextInsertIndex, 0, ...block);
      return {
        items: remaining,
        sortedIndexes,
        nextIndexes: block.map((_, offset) => nextInsertIndex + offset),
      };
    }

    function applyReorderedManualColumnWidths(columnIndexes, insertIndex, columnCount) {
      const widthState = getColumnWidthState();
      const manualWidths = new Array(columnCount).fill(undefined);
      Object.entries(widthState.manual).forEach(([key, value]) => {
        const columnIndex = Number(key);
        if (columnIndex >= 0 && columnIndex < columnCount) {
          manualWidths[columnIndex] = value;
        }
      });
      const reorderedWidths = reorderIndexedBlock(
        manualWidths,
        columnIndexes,
        insertIndex,
        2,
      );
      if (!reorderedWidths) {
        return;
      }
      const nextManualWidths = {};
      reorderedWidths.items.forEach((value, columnIndex) => {
        if (typeof value === "number") {
          nextManualWidths[columnIndex] = value;
        }
      });
      widthState.manual = nextManualWidths;
    }

    function moveSheetsToInsertIndex(sheetIndexes, insertIndex) {
      const sheets = state.workbook?.sheets || [];
      const reorderedSheets = reorderIndexedBlock(sheets, sheetIndexes, insertIndex, 0);
      if (!reorderedSheets) {
        return;
      }
      if (arraysEqual(reorderedSheets.sortedIndexes, reorderedSheets.nextIndexes)) {
        return;
      }

      recordUndoSnapshot();
      const widthStates = sheets.map((_, sheetIndex) => columnWidthStateBySheet.get(sheetIndex));
      const reorderedWidthStates = reorderIndexedBlock(widthStates, sheetIndexes, insertIndex, 0);
      sheets.splice(0, sheets.length, ...reorderedSheets.items);
      rebuildColumnWidthStateMap(reorderedWidthStates?.items || []);

      if (reorderedSheets.sortedIndexes.includes(activeSheetIndex)) {
        activeSheetIndex = reorderedSheets.nextIndexes[reorderedSheets.sortedIndexes.indexOf(activeSheetIndex)];
      } else {
        let nextActiveSheetIndex = activeSheetIndex;
        reorderedSheets.sortedIndexes.forEach((sheetIndex) => {
          if (sheetIndex < activeSheetIndex) {
            nextActiveSheetIndex -= 1;
          }
        });
        const insertAnchor = reorderedSheets.nextIndexes[0];
        if (insertAnchor <= nextActiveSheetIndex) {
          nextActiveSheetIndex += reorderedSheets.nextIndexes.length;
        }
        activeSheetIndex = Math.max(0, Math.min(nextActiveSheetIndex, sheets.length - 1));
      }

      if (typeof editingSheetIndex === "number") {
        const editingPosition = reorderedSheets.sortedIndexes.indexOf(editingSheetIndex);
        editingSheetIndex = editingPosition >= 0
          ? reorderedSheets.nextIndexes[editingPosition]
          : undefined;
      }
      clearSheetScopedSelectionState();
      queueWorkbookSave();
      safeRender();
    }

    function moveRowsToInsertIndex(rowIndexes, insertIndex) {
      const sheet = getActiveSheet();
      const reorderedRows = reorderIndexedBlock(sheet.rows, rowIndexes, insertIndex, 0);
      if (!reorderedRows) {
        return;
      }
      if (arraysEqual(reorderedRows.sortedIndexes, reorderedRows.nextIndexes)) {
        return;
      }
      recordUndoSnapshot();
      sheet.rows.splice(0, sheet.rows.length, ...reorderedRows.items);
      setSelectedRowIndexes(reorderedRows.nextIndexes);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function moveColumnsToInsertIndex(columnIndexes, insertIndex) {
      const sheet = getActiveSheet();
      const reorderedHeaders = reorderIndexedBlock(sheet.headers, columnIndexes, insertIndex, 2);
      if (!reorderedHeaders) {
        return;
      }
      if (arraysEqual(reorderedHeaders.sortedIndexes, reorderedHeaders.nextIndexes)) {
        return;
      }

      recordUndoSnapshot();
      const columnCount = sheet.headers.length;
      sheet.headers.splice(0, sheet.headers.length, ...reorderedHeaders.items);
      ensureHeaderBackgroundLength(sheet);
      const reorderedHeaderBackgrounds = reorderIndexedBlock(
        sheet.headerBackgrounds,
        columnIndexes,
        insertIndex,
        2,
      );
      if (reorderedHeaderBackgrounds) {
        sheet.headerBackgrounds.splice(0, sheet.headerBackgrounds.length, ...reorderedHeaderBackgrounds.items);
      }
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, columnCount);
        const reorderedValues = reorderIndexedBlock(row.values, columnIndexes, insertIndex, 2);
        if (reorderedValues) {
          row.values.splice(0, row.values.length, ...reorderedValues.items);
        }
        const reorderedBackgrounds = reorderIndexedBlock(row.backgrounds, columnIndexes, insertIndex, 2);
        if (reorderedBackgrounds) {
          row.backgrounds.splice(0, row.backgrounds.length, ...reorderedBackgrounds.items);
        }
        ensureRecordRowLength(row, sheet.headers.length);
      });
      applyReorderedManualColumnWidths(columnIndexes, insertIndex, columnCount);
      setSelectedColumnIndexes(reorderedHeaders.nextIndexes);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function getRowDropInsertIndex(rowIndex, clientY, element) {
      const rect = element.getBoundingClientRect();
      return clientY > rect.top + rect.height / 2
        ? rowIndex + 1
        : rowIndex;
    }

    function getColumnDropInsertIndex(columnIndex, clientX, element) {
      const rect = element.getBoundingClientRect();
      return clientX > rect.left + rect.width / 2
        ? columnIndex + 1
        : columnIndex;
    }

    function getSheetDropInsertIndex(sheetIndex, clientX, element) {
      const rect = element.getBoundingClientRect();
      return clientX > rect.left + rect.width / 2
        ? sheetIndex + 1
        : sheetIndex;
    }

    function startRowReorderDrag(event, rowIndex) {
      const rowIndexes = getEffectiveRowIndexes(rowIndex);
      rowReorderDragState = {
        sheetIndex: activeSheetIndex,
        rowIndexes,
      };
      columnReorderDragState = undefined;
      sheetReorderDragState = undefined;
      isSelectingRows = false;
      rowSelectionAnchorIndex = undefined;
      clearReorderDropIndicator();
      document.body.classList.add("reordering-row");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", "spreadsheet-row");
      }
    }

    function startColumnReorderDrag(event, columnIndex) {
      const columnIndexes = getEffectiveColumnIndexes(columnIndex);
      if (
        !isDraggableColumnIndex(columnIndex) ||
        columnIndexes.some((nextColumnIndex) => !isDraggableColumnIndex(nextColumnIndex))
      ) {
        event.preventDefault();
        return;
      }
      rowReorderDragState = undefined;
      sheetReorderDragState = undefined;
      columnReorderDragState = {
        sheetIndex: activeSheetIndex,
        columnIndexes,
      };
      isSelectingColumns = false;
      columnSelectionAnchorIndex = undefined;
      clearReorderDropIndicator();
      document.body.classList.add("reordering-column");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", "spreadsheet-column");
      }
    }

    function startSheetReorderDrag(event, sheetIndex) {
      if (typeof editingSheetIndex === "number") {
        event.preventDefault();
        return;
      }
      rowReorderDragState = undefined;
      columnReorderDragState = undefined;
      sheetReorderDragState = {
        sheetIndex,
      };
      clearReorderDropIndicator();
      document.body.classList.add("reordering-sheet");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", "spreadsheet-sheet");
      }
    }

    function finishReorderDrag() {
      rowReorderDragState = undefined;
      columnReorderDragState = undefined;
      sheetReorderDragState = undefined;
      clearReorderDropIndicator();
      document.body.classList.remove("reordering-row", "reordering-column", "reordering-sheet");
    }

    function createChoiceTriggerButton(rowIndex, columnIndex) {
      const choiceButton = document.createElement("button");
      choiceButton.type = "button";
      choiceButton.className = "choice-trigger";
      choiceButton.textContent = "⌄";
      choiceButton.title = "Select field value";
      choiceButton.dataset.rowIndex = String(rowIndex);
      choiceButton.dataset.columnIndex = String(columnIndex);
      choiceButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        const nextRowIndex = Number(choiceButton.dataset.rowIndex);
        const nextColumnIndex = Number(choiceButton.dataset.columnIndex);
        const choices = getCellMenuChoices(getActiveSheet(), nextRowIndex, nextColumnIndex);
        if (choices.length === 0) {
          return;
        }
        openChoiceMenu(choiceButton, nextRowIndex, nextColumnIndex, choices);
      });
      return choiceButton;
    }

    function syncCellChoiceTrigger(rowIndex, columnIndex) {
      const input = document.querySelector(
        'input[data-scope="cell"][data-row-index="' + rowIndex + '"][data-column-index="' + columnIndex + '"]',
      );
      if (!(input instanceof HTMLInputElement)) {
        return;
      }
      const editor = input.closest(".grid-editor");
      if (!(editor instanceof HTMLElement)) {
        return;
      }

      const choices = getCellMenuChoices(getActiveSheet(), rowIndex, columnIndex);
      const existingButton = editor.querySelector(".choice-trigger");
      if (choices.length === 0) {
        editor.classList.remove("has-choice");
        existingButton?.remove();
        return;
      }

      editor.classList.add("has-choice");
      if (existingButton instanceof HTMLButtonElement) {
        existingButton.dataset.rowIndex = String(rowIndex);
        existingButton.dataset.columnIndex = String(columnIndex);
        return;
      }

      editor.appendChild(createChoiceTriggerButton(rowIndex, columnIndex));
    }

    function refreshChoiceTriggersForRow(rowIndex) {
      const sheet = getActiveSheet();
      const row = sheet.rows[rowIndex];
      if (!row || isCommentRow(row)) {
        return;
      }
      for (let columnIndex = 2; columnIndex < sheet.headers.length; columnIndex += 1) {
        syncCellChoiceTrigger(rowIndex, columnIndex);
      }
    }

    function refreshChoiceTriggersForColumn(columnIndex) {
      if (columnIndex < 2) {
        return;
      }
      const sheet = getActiveSheet();
      sheet.rows.forEach((row, rowIndex) => {
        if (isCommentRow(row)) {
          return;
        }
        syncCellChoiceTrigger(rowIndex, columnIndex);
      });
    }

    function refreshValidationStyles() {
      document.querySelectorAll("[data-validation-row-index][data-validation-column-index]").forEach((element) => {
        applyIssueClass(
          element,
          Number(element.dataset.validationRowIndex),
          Number(element.dataset.validationColumnIndex),
        );
      });
    }

    function getActiveSelectedCellRange() {
      if (!selectedCellRange || selectedCellRange.sheetIndex !== activeSheetIndex) {
        return undefined;
      }
      return {
        sheetIndex: selectedCellRange.sheetIndex,
        startRowIndex: Math.min(selectedCellRange.startRowIndex, selectedCellRange.endRowIndex),
        endRowIndex: Math.max(selectedCellRange.startRowIndex, selectedCellRange.endRowIndex),
        startColumnIndex: Math.min(selectedCellRange.startColumnIndex, selectedCellRange.endColumnIndex),
        endColumnIndex: Math.max(selectedCellRange.startColumnIndex, selectedCellRange.endColumnIndex),
      };
    }

    function clearCellSelection() {
      if (!selectedCellRange && !cellSelectionAnchor) {
        return;
      }
      selectedCellRange = undefined;
      cellSelectionAnchor = undefined;
      isSelectingCells = false;
      refreshCellSelectionStyles();
    }

    function clearRowSelection() {
      if (selectedRowIndexes.length === 0 && !isSelectingRows && typeof rowSelectionAnchorIndex !== "number") {
        return;
      }
      selectedRowIndexes = [];
      rowSelectionAnchorIndex = undefined;
      rowSelectionPivotIndex = undefined;
      isSelectingRows = false;
      refreshRowSelectionStyles();
    }

    function clearColumnSelection() {
      if (selectedColumnIndexes.length === 0 && !isSelectingColumns && typeof columnSelectionAnchorIndex !== "number") {
        return;
      }
      selectedColumnIndexes = [];
      columnSelectionAnchorIndex = undefined;
      columnSelectionPivotIndex = undefined;
      isSelectingColumns = false;
      refreshColumnSelectionStyles();
    }

    function setSelectedCellRange(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex) {
      const sheet = getActiveSheet();
      if (isCommentRow(sheet.rows[startRowIndex]) || isCommentRow(sheet.rows[endRowIndex])) {
        return;
      }
      selectedCellRange = {
        sheetIndex: activeSheetIndex,
        startRowIndex,
        startColumnIndex,
        endRowIndex,
        endColumnIndex,
      };
      rowSelectionAnchorIndex = undefined;
      isSelectingRows = false;
      columnSelectionAnchorIndex = undefined;
      isSelectingColumns = false;
      if (selectedColumnIndexes.length > 0) {
        selectedColumnIndexes = [];
        refreshColumnSelectionStyles();
      }
      if (selectedRowIndexes.length > 0) {
        selectedRowIndexes = [];
        refreshRowSelectionStyles();
      }
      refreshCellSelectionStyles();
    }

    function isCellSelected(rowIndex, columnIndex) {
      const range = getActiveSelectedCellRange();
      if (!range || isCommentRow(getActiveSheet().rows[rowIndex])) {
        return false;
      }
      return rowIndex >= range.startRowIndex &&
        rowIndex <= range.endRowIndex &&
        columnIndex >= range.startColumnIndex &&
        columnIndex <= range.endColumnIndex;
    }

    function hasMultiCellSelection() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return false;
      }
      return range.startRowIndex !== range.endRowIndex ||
        range.startColumnIndex !== range.endColumnIndex;
    }

    function refreshCellSelectionStyles() {
      const isMultiSelection = hasMultiCellSelection();
      document.querySelectorAll('td[data-row-index][data-column-index]').forEach((cell) => {
        const selected = isCellSelected(Number(cell.dataset.rowIndex), Number(cell.dataset.columnIndex));
        cell.classList.toggle("cell-selected", selected);
        cell.classList.toggle("cell-selected-multi", selected && isMultiSelection);
      });
      refreshBackgroundPaletteState();
    }

    function handleCellMouseDown(event, rowIndex, columnIndex) {
      if (event.button !== 0) {
        return;
      }
      clearRowSelection();
      clearColumnSelection();
      if (!cellSelectionEnhancementsEnabled || !event.shiftKey) {
        clearCellSelection();
        return;
      }
      closeFloatingMenu();
      closeSuggestionMenu();
      const anchor = getCellSelectionAnchor(rowIndex, columnIndex);
      isSelectingCells = true;
      cellSelectionAnchor = {
        sheetIndex: activeSheetIndex,
        rowIndex: anchor.rowIndex,
        columnIndex: anchor.columnIndex,
      };
      setSelectedCellRange(anchor.rowIndex, anchor.columnIndex, rowIndex, columnIndex);
      event.preventDefault();
    }

    function handleCellMouseEnter(rowIndex, columnIndex) {
      if (!isSelectingCells || !cellSelectionAnchor || cellSelectionAnchor.sheetIndex !== activeSheetIndex) {
        return;
      }
      setSelectedCellRange(
        cellSelectionAnchor.rowIndex,
        cellSelectionAnchor.columnIndex,
        rowIndex,
        columnIndex,
      );
    }

    function buildSelectedCellClipboardText() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return "";
      }
      // Use TSV so rectangular selections stay portable across spreadsheet widgets and apps.
      const sheet = getActiveSheet();
      const lines = [];
      for (let rowIndex = range.startRowIndex; rowIndex <= range.endRowIndex; rowIndex += 1) {
        const row = sheet.rows[rowIndex];
        if (!row || isCommentRow(row)) {
          continue;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        const values = [];
        for (let columnIndex = range.startColumnIndex; columnIndex <= range.endColumnIndex; columnIndex += 1) {
          values.push(formatClipboardCellValue(row.values[columnIndex] || ""));
        }
        lines.push(values.join("\\t"));
      }
      return lines.join("\\n");
    }

    function formatClipboardCellValue(value) {
      const text = String(value || "");
      if (!/[\\t\\n\\r"]/.test(text)) {
        return text;
      }
      return '"' + text.replace(/"/g, '""') + '"';
    }

    async function writeTextToClipboard(text) {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return;
      }

      const helper = document.createElement("textarea");
      helper.value = text;
      helper.setAttribute("readonly", "true");
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      helper.style.pointerEvents = "none";
      document.body.appendChild(helper);
      helper.focus();
      helper.select();
      const didCopy = document.execCommand("copy");
      document.body.removeChild(helper);
      if (!didCopy) {
        throw new Error("Clipboard copy is not available in this webview.");
      }
    }

    async function readTextFromClipboard() {
      if (navigator.clipboard?.readText) {
        return navigator.clipboard.readText();
      }

      const helper = document.createElement("textarea");
      helper.value = "";
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      helper.style.pointerEvents = "none";
      document.body.appendChild(helper);
      helper.focus();
      const didPaste = document.execCommand("paste");
      const text = helper.value;
      document.body.removeChild(helper);
      if (!didPaste && !text) {
        throw new Error("Clipboard paste is not available in this webview.");
      }
      return text;
    }

    function parseClipboardTable(text) {
      const source = String(text || "");
      const rows = [];
      let currentRow = [];
      let currentCell = "";
      let inQuotes = false;

      for (let index = 0; index < source.length; index += 1) {
        const character = source[index];
        if (inQuotes) {
          if (character === '"') {
            if (source[index + 1] === '"') {
              currentCell += '"';
              index += 1;
            } else {
              inQuotes = false;
            }
          } else if (character === "\\r") {
            if (source[index + 1] === "\\n") {
              index += 1;
            }
            currentCell += "\\n";
          } else {
            currentCell += character;
          }
          continue;
        }

        if (character === '"') {
          inQuotes = true;
        } else if (character === "\t") {
          currentRow.push(currentCell);
          currentCell = "";
        } else if (character === "\\n" || character === "\\r") {
          if (character === "\\r" && source[index + 1] === "\\n") {
            index += 1;
          }
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = "";
        } else {
          currentCell += character;
        }
      }

      currentRow.push(currentCell);
      rows.push(currentRow);
      if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") {
        rows.pop();
      }
      return rows;
    }

    async function copySelectedCells() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return;
      }
      await writeTextToClipboard(buildSelectedCellClipboardText());
    }

    function clearSelectedCellValues() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return;
      }

      recordUndoSnapshot();
      const sheet = getActiveSheet();
      const affectedColumns = new Set();
      const affectedTypeRows = new Set();
      for (let rowIndex = range.startRowIndex; rowIndex <= range.endRowIndex; rowIndex += 1) {
        const row = sheet.rows[rowIndex];
        if (!row || isCommentRow(row)) {
          continue;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        for (let columnIndex = range.startColumnIndex; columnIndex <= range.endColumnIndex; columnIndex += 1) {
          row.values[columnIndex] = "";
          const input = getGridCellInput(rowIndex, columnIndex);
          if (input) {
            input.value = "";
          }
          affectedColumns.add(columnIndex);
          if (columnIndex === 1) {
            affectedTypeRows.add(rowIndex);
          }
        }
      }

      affectedTypeRows.forEach((rowIndex) => {
        refreshChoiceTriggersForRow(rowIndex);
      });
      affectedColumns.forEach((columnIndex) => {
        refreshAutoColumnWidth(columnIndex);
      });
      queueWorkbookSave();
      refreshCellSelectionStyles();
    }

    async function cutSelectedCells() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return;
      }
      await copySelectedCells();
      clearSelectedCellValues();
    }

    function ensureSpreadsheetColumnCount(columnCount) {
      const sheet = getActiveSheet();
      if (columnCount <= sheet.headers.length) {
        return;
      }
      while (sheet.headers.length < columnCount) {
        sheet.headers.push("");
      }
      ensureHeaderBackgroundLength(sheet);
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
      });
    }

    function getPasteTargetRowIndexes(startRowIndex, rowCount) {
      const sheet = getActiveSheet();
      const targetRowIndexes = [];
      let cursor = startRowIndex;
      while (targetRowIndexes.length < rowCount) {
        while (cursor < sheet.rows.length && isCommentRow(sheet.rows[cursor])) {
          cursor += 1;
        }
        if (cursor >= sheet.rows.length) {
          sheet.rows.push(createEmptyRecordRow(sheet.headers.length));
        }
        const row = sheet.rows[cursor];
        if (!row || isCommentRow(row)) {
          cursor += 1;
          continue;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        targetRowIndexes.push(cursor);
        cursor += 1;
      }
      return targetRowIndexes;
    }

    function pasteClipboardTableAt(rowIndex, columnIndex, clipboardTable) {
      const pastedRows = Array.isArray(clipboardTable)
        ? clipboardTable.filter((row) => Array.isArray(row))
        : [];
      if (pastedRows.length === 0) {
        return;
      }

      const maxColumnCount = pastedRows.reduce((max, row) => Math.max(max, row.length), 0);
      if (maxColumnCount <= 0) {
        return;
      }

      recordUndoSnapshot();
      ensureSpreadsheetColumnCount(columnIndex + maxColumnCount);
      const targetRowIndexes = getPasteTargetRowIndexes(rowIndex, pastedRows.length);
      pastedRows.forEach((values, pastedRowOffset) => {
        const targetRowIndex = targetRowIndexes[pastedRowOffset];
        const row = getActiveSheet().rows[targetRowIndex];
        ensureRecordRowLength(row, getActiveSheet().headers.length);
        values.forEach((value, pastedColumnOffset) => {
          row.values[columnIndex + pastedColumnOffset] = String(value || "");
        });
      });

      setSelectedCellRange(
        targetRowIndexes[0],
        columnIndex,
        targetRowIndexes[targetRowIndexes.length - 1],
        columnIndex + maxColumnCount - 1,
      );
      queueWorkbookSave();
      safeRender();
    }

    async function pasteIntoCellSelection(rowIndex, columnIndex) {
      const clipboardText = await readTextFromClipboard();
      const clipboardTable = parseClipboardTable(clipboardText);
      pasteClipboardTableAt(rowIndex, columnIndex, clipboardTable);
    }

    function handleCellContextMenu(event, rowIndex, columnIndex) {
      event.preventDefault();
      event.stopPropagation();
      const clientX = event.clientX;
      const clientY = event.clientY;
      if (!isCellSelected(rowIndex, columnIndex)) {
        setSelectedCellRange(rowIndex, columnIndex, rowIndex, columnIndex);
      }
      openFloatingMenuAtPosition(clientX, clientY, [
        {
          label: "Paste",
          action: async () => {
            await pasteIntoCellSelection(rowIndex, columnIndex);
          },
        },
        {
          label: "Copy",
          action: async () => {
            await copySelectedCells();
          },
        },
        {
          label: "Cut",
          action: async () => {
            await cutSelectedCells();
          },
        },
        {
          label: "Clear",
          action: async () => {
            clearSelectedCellValues();
          },
        },
      ]);
    }

    function getSelectedCellShortcutTarget() {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return undefined;
      }
      const activeElement = document.activeElement;
      if (
        activeElement instanceof HTMLInputElement &&
        activeElement.dataset.scope === "cell"
      ) {
        const rowIndex = Number(activeElement.dataset.rowIndex);
        const columnIndex = Number(activeElement.dataset.columnIndex);
        if (
          Number.isInteger(rowIndex) &&
          Number.isInteger(columnIndex) &&
          rowIndex >= range.startRowIndex &&
          rowIndex <= range.endRowIndex &&
          columnIndex >= range.startColumnIndex &&
          columnIndex <= range.endColumnIndex
        ) {
          return {
            rowIndex,
            columnIndex,
          };
        }
      }
      return {
        rowIndex: range.startRowIndex,
        columnIndex: range.startColumnIndex,
      };
    }

    function handleMultiCellShortcutKeydown(event) {
      if (event.defaultPrevented || !hasMultiCellSelection()) {
        return;
      }
      const target = getSelectedCellShortcutTarget();
      if (!target) {
        return;
      }

      const key = String(event.key || "");
      const normalizedKey = key.toLowerCase();
      const hasPrimaryModifier = (event.metaKey || event.ctrlKey) && !event.altKey;

      if (hasPrimaryModifier && normalizedKey === "c") {
        event.preventDefault();
        closeFloatingMenu();
        closeSuggestionMenu();
        void copySelectedCells().catch((error) => {
          console.error("Spreadsheet copy failed:", error);
        });
        return;
      }

      if (hasPrimaryModifier && normalizedKey === "x") {
        event.preventDefault();
        closeFloatingMenu();
        closeSuggestionMenu();
        void cutSelectedCells().catch((error) => {
          console.error("Spreadsheet cut failed:", error);
        });
        return;
      }

      if (hasPrimaryModifier && normalizedKey === "v") {
        event.preventDefault();
        closeFloatingMenu();
        closeSuggestionMenu();
        void pasteIntoCellSelection(target.rowIndex, target.columnIndex).catch((error) => {
          console.error("Spreadsheet paste failed:", error);
        });
        return;
      }

      if (
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        (key === "Delete" || key === "Backspace")
      ) {
        event.preventDefault();
        closeFloatingMenu();
        closeSuggestionMenu();
        clearSelectedCellValues();
      }
    }

    function isFieldNameHeaderInput(element) {
      return element instanceof HTMLInputElement &&
        element.dataset.scope === "header" &&
        Number(element.dataset.columnIndex) >= 2;
    }

    function isRecordTypeCellInput(element) {
      return element instanceof HTMLInputElement &&
        element.dataset.scope === "cell" &&
        Number(element.dataset.columnIndex) === 1;
    }

    function isSuggestionInput(element) {
      return isFieldNameHeaderInput(element) || isRecordTypeCellInput(element);
    }

    function getHeaderFieldInput(columnIndex) {
      const input = document.querySelector(
        'input[data-scope="header"][data-column-index="' + columnIndex + '"]',
      );
      return input instanceof HTMLInputElement ? input : undefined;
    }

    function getHeaderInput(columnIndex) {
      const input = document.querySelector(
        'input[data-scope="header"][data-column-index="' + columnIndex + '"]',
      );
      return input instanceof HTMLInputElement ? input : undefined;
    }

    function getRecordTypeCellInput(rowIndex) {
      const input = document.querySelector(
        'input[data-scope="cell"][data-row-index="' + rowIndex + '"][data-column-index="1"]',
      );
      return input instanceof HTMLInputElement ? input : undefined;
    }

    function updateHeaderFieldValue(columnIndex, value) {
      const sheet = getActiveSheet();
      if (sheet.headers[columnIndex] === value) {
        return;
      }
      recordUndoSnapshot();
      sheet.headers[columnIndex] = value;
      refreshChoiceTriggersForColumn(columnIndex);
      refreshAutoColumnWidth(columnIndex);
      queueWorkbookSave();
    }

    function updateRecordTypeValue(rowIndex, value) {
      const sheet = getActiveSheet();
      const row = sheet.rows[rowIndex];
      if (!row || isCommentRow(row)) {
        return;
      }
      ensureRecordRowLength(row, sheet.headers.length);
      const nextValue = String(value || "");
      if (row.values[1] === nextValue) {
        return;
      }
      recordUndoSnapshot();
      row.values[1] = nextValue;
      refreshChoiceTriggersForRow(rowIndex);
      refreshAutoColumnWidth(1);
      queueWorkbookSave();
    }

    function getFilteredSuggestions(items, query) {
      const normalizedQuery = String(query || "").trim().toUpperCase();
      const prefixMatches = [];
      const containsMatches = [];
      for (const item of items || []) {
        const normalizedItem = String(item || "").toUpperCase();
        if (!normalizedQuery || normalizedItem.startsWith(normalizedQuery)) {
          prefixMatches.push(item);
        } else if (normalizedItem.includes(normalizedQuery)) {
          containsMatches.push(item);
        }
      }
      const allMatches = prefixMatches.concat(containsMatches);
      return {
        items: allMatches.slice(0, FIELD_NAME_SUGGESTION_LIMIT),
      };
    }

    function getSuggestionDescriptor(input) {
      if (isFieldNameHeaderInput(input)) {
        return {
          kind: "fieldNameHeader",
          columnIndex: Number(input.dataset.columnIndex),
          suggestions: getFilteredSuggestions(state.fieldNames || [], input.value),
          emptyMessage: "No EPICS field names match this filter.",
          idleHint: "Type to filter EPICS field names.",
          apply: (value) => {
            input.value = value;
            updateHeaderFieldValue(Number(input.dataset.columnIndex), value);
          },
        };
      }

      if (isRecordTypeCellInput(input)) {
        return {
          kind: "recordTypeCell",
          rowIndex: Number(input.dataset.rowIndex),
          columnIndex: 1,
          suggestions: getFilteredSuggestions(state.recordTypes || [], input.value),
          emptyMessage: "No EPICS record types match this filter.",
          idleHint: "Type to filter EPICS record types.",
          apply: (value) => {
            input.value = value;
            updateRecordTypeValue(Number(input.dataset.rowIndex), value);
          },
        };
      }

      return undefined;
    }

    function positionMenu(menuElement, anchor) {
      const anchorRect = anchor.getBoundingClientRect();
      const menuRect = menuElement.getBoundingClientRect();
      let left = anchorRect.left;
      let top = anchorRect.bottom + 6;
      if (left + menuRect.width > window.innerWidth - 12) {
        left = Math.max(12, window.innerWidth - menuRect.width - 12);
      }
      if (top + menuRect.height > window.innerHeight - 12) {
        top = Math.max(12, anchorRect.top - menuRect.height - 6);
      }
      menuElement.style.left = left + "px";
      menuElement.style.top = top + "px";
    }

    function positionMenuAtPoint(menuElement, clientX, clientY) {
      const menuRect = menuElement.getBoundingClientRect();
      let left = clientX;
      let top = clientY;
      if (left + menuRect.width > window.innerWidth - 12) {
        left = Math.max(12, window.innerWidth - menuRect.width - 12);
      }
      if (top + menuRect.height > window.innerHeight - 12) {
        top = Math.max(12, window.innerHeight - menuRect.height - 12);
      }
      menuElement.style.left = left + "px";
      menuElement.style.top = top + "px";
    }

    function closeSuggestionMenu() {
      activeSuggestionState = undefined;
      suggestionMenu.hidden = true;
      suggestionMenu.innerHTML = "";
    }

    function getSuggestionInputForState() {
      if (!activeSuggestionState) {
        return undefined;
      }
      if (activeSuggestionState.kind === "fieldNameHeader") {
        return getHeaderFieldInput(activeSuggestionState.columnIndex);
      }
      if (activeSuggestionState.kind === "recordTypeCell") {
        return getRecordTypeCellInput(activeSuggestionState.rowIndex);
      }
      return undefined;
    }

    function isActiveSuggestionInput(input) {
      const descriptor = getSuggestionDescriptor(input);
      if (!descriptor || !activeSuggestionState) {
        return false;
      }
      return descriptor.kind === activeSuggestionState.kind &&
        descriptor.columnIndex === activeSuggestionState.columnIndex &&
        descriptor.rowIndex === activeSuggestionState.rowIndex;
    }

    function scheduleSuggestionMenuClose() {
      window.setTimeout(() => {
        const activeElement = document.activeElement;
        if (
          suggestionMenu.contains(activeElement) ||
          isSuggestionInput(activeElement)
        ) {
          return;
        }
        closeSuggestionMenu();
      }, 0);
    }

    function renderSuggestionMenu(input) {
      const descriptor = getSuggestionDescriptor(input);
      if (!descriptor) {
        closeSuggestionMenu();
        return;
      }

      const suggestions = descriptor.suggestions;
      const highlightedIndex = Math.min(
        activeSuggestionState?.highlightedIndex ?? 0,
        Math.max(0, suggestions.items.length - 1),
      );
      activeSuggestionState = {
        kind: descriptor.kind,
        rowIndex: descriptor.rowIndex,
        columnIndex: descriptor.columnIndex,
        highlightedIndex,
      };

      suggestionMenu.innerHTML = "";

      if (suggestions.items.length === 0) {
        const hint = document.createElement("div");
        hint.className = "menu-hint";
        hint.textContent = String(input.value || "").trim()
          ? descriptor.emptyMessage
          : descriptor.idleHint;
        suggestionMenu.appendChild(hint);
      } else {
        suggestions.items.forEach((item, index) => {
          const button = document.createElement("button");
          button.type = "button";
          button.textContent = item;
          if (index === highlightedIndex) {
            button.classList.add("active");
          }
          button.addEventListener("mousedown", (event) => {
            event.preventDefault();
          });
          button.addEventListener("click", () => {
            descriptor.apply(item);
            closeSuggestionMenu();
            input.focus();
            input.setSelectionRange(String(item).length, String(item).length);
          });
          suggestionMenu.appendChild(button);
        });
      }

      suggestionMenu.hidden = false;
      positionMenu(suggestionMenu, input);
      suggestionMenu.querySelector("button.active")?.scrollIntoView({ block: "nearest" });
    }

    function openSuggestionMenu(input) {
      const descriptor = getSuggestionDescriptor(input);
      if (!descriptor) {
        closeSuggestionMenu();
        return;
      }
      closeFloatingMenu();
      if (!isActiveSuggestionInput(input)) {
        activeSuggestionState = {
          kind: descriptor.kind,
          rowIndex: descriptor.rowIndex,
          columnIndex: descriptor.columnIndex,
          highlightedIndex: 0,
        };
      }
      renderSuggestionMenu(input);
    }

    function refreshSuggestionMenu() {
      if (!activeSuggestionState) {
        return;
      }
      const input = getSuggestionInputForState();
      if (!input) {
        closeSuggestionMenu();
        return;
      }
      renderSuggestionMenu(input);
    }

    function moveSuggestion(input, direction) {
      const descriptor = getSuggestionDescriptor(input);
      const suggestions = descriptor?.suggestions;
      if (suggestions.items.length === 0) {
        openSuggestionMenu(input);
        return;
      }
      const currentIndex = activeSuggestionState?.highlightedIndex ?? 0;
      activeSuggestionState = {
        kind: descriptor.kind,
        rowIndex: descriptor.rowIndex,
        columnIndex: descriptor.columnIndex,
        highlightedIndex: (currentIndex + direction + suggestions.items.length) % suggestions.items.length,
      };
      renderSuggestionMenu(input);
    }

    function applyActiveSuggestion(input) {
      const descriptor = getSuggestionDescriptor(input);
      const suggestions = descriptor?.suggestions;
      if (suggestions.items.length === 0) {
        return false;
      }
      const highlightedIndex = Math.min(
        activeSuggestionState?.highlightedIndex ?? 0,
        suggestions.items.length - 1,
      );
      const value = suggestions.items[highlightedIndex];
      descriptor.apply(value);
      closeSuggestionMenu();
      input.focus();
      input.setSelectionRange(String(value).length, String(value).length);
      return true;
    }

    function handleSuggestionKeydown(event, input) {
      if (!isSuggestionInput(input)) {
        return;
      }

      switch (event.key) {
        case "ArrowDown":
          if (suggestionMenu.hidden || !isActiveSuggestionInput(input)) {
            return;
          }
          event.preventDefault();
          moveSuggestion(input, 1);
          break;

        case "ArrowUp":
          if (suggestionMenu.hidden || !isActiveSuggestionInput(input)) {
            return;
          }
          event.preventDefault();
          moveSuggestion(input, -1);
          break;

        case "Enter":
          if (!suggestionMenu.hidden && isActiveSuggestionInput(input) && applyActiveSuggestion(input)) {
            event.preventDefault();
          }
          break;

        case "Escape":
          if (!suggestionMenu.hidden) {
            event.preventDefault();
            closeSuggestionMenu();
          }
          break;

        case "Tab":
          closeSuggestionMenu();
          break;

        default:
          break;
      }
    }

    function getGridCellInput(rowIndex, columnIndex) {
      const input = document.querySelector(
        'input[data-scope="cell"][data-row-index="' + rowIndex + '"][data-column-index="' + columnIndex + '"]',
      );
      return input instanceof HTMLInputElement ? input : undefined;
    }

    function getCommentInput(rowIndex) {
      const textarea = document.querySelector(
        'textarea[data-scope="comment"][data-row-index="' + rowIndex + '"]',
      );
      return textarea instanceof HTMLTextAreaElement ? textarea : undefined;
    }

    function getNavigationPosition(element) {
      if (element instanceof HTMLInputElement && element.dataset.scope === "header") {
        return {
          scope: "header",
          rowIndex: -1,
          columnIndex: Number(element.dataset.columnIndex),
        };
      }
      if (element instanceof HTMLInputElement && element.dataset.scope === "cell") {
        return {
          scope: "cell",
          rowIndex: Number(element.dataset.rowIndex),
          columnIndex: Number(element.dataset.columnIndex),
        };
      }
      if (element instanceof HTMLTextAreaElement && element.dataset.scope === "comment") {
        return {
          scope: "comment",
          rowIndex: Number(element.dataset.rowIndex),
          columnIndex: Number(element.dataset.columnIndex || 0),
        };
      }
      return undefined;
    }

    function getNavigationTarget(rowIndex, columnIndex) {
      const sheet = getActiveSheet();
      if (rowIndex < 0) {
        return getHeaderInput(columnIndex);
      }
      const row = sheet.rows[rowIndex];
      if (!row) {
        return undefined;
      }
      if (isCommentRow(row)) {
        return getCommentInput(rowIndex);
      }
      const boundedColumnIndex = Math.max(0, Math.min(columnIndex, sheet.headers.length - 1));
      return getGridCellInput(rowIndex, boundedColumnIndex);
    }

    function canMoveCaretWithinElement(element, directionKey) {
      if (
        !(element instanceof HTMLInputElement) &&
        !(element instanceof HTMLTextAreaElement)
      ) {
        return false;
      }
      if (
        typeof element.selectionStart !== "number" ||
        typeof element.selectionEnd !== "number"
      ) {
        return false;
      }
      if (element.selectionStart !== element.selectionEnd) {
        return true;
      }

      const value = String(element.value || "");
      const caretIndex = element.selectionStart;
      switch (directionKey) {
        case "ArrowLeft":
          return caretIndex > 0;

        case "ArrowRight":
          return caretIndex < value.length;

        case "ArrowUp":
          return element instanceof HTMLTextAreaElement &&
            value.slice(0, caretIndex).includes("\\n");

        case "ArrowDown":
          return element instanceof HTMLTextAreaElement &&
            value.slice(caretIndex).includes("\\n");

        default:
          return false;
      }
    }

    function moveEditingFocus(element, directionKey) {
      const position = getNavigationPosition(element);
      if (!position) {
        return false;
      }

      const sheet = getActiveSheet();
      let nextRowIndex = position.rowIndex;
      let nextColumnIndex = position.columnIndex;

      switch (directionKey) {
        case "ArrowLeft":
          if (position.scope === "comment" || nextColumnIndex <= 0) {
            return false;
          }
          nextColumnIndex -= 1;
          break;

        case "ArrowRight":
          if (position.scope === "comment" || nextColumnIndex >= sheet.headers.length - 1) {
            return false;
          }
          nextColumnIndex += 1;
          break;

        case "ArrowUp":
          nextRowIndex -= 1;
          break;

        case "ArrowDown":
          nextRowIndex += 1;
          break;

        default:
          return false;
      }

      const nextElement = getNavigationTarget(nextRowIndex, nextColumnIndex);
      if (
        !(nextElement instanceof HTMLInputElement) &&
        !(nextElement instanceof HTMLTextAreaElement)
      ) {
        return false;
      }

      closeSuggestionMenu();
      nextElement.focus();
      if (typeof nextElement.setSelectionRange === "function") {
        const caretIndex = directionKey === "ArrowLeft" || directionKey === "ArrowUp"
          ? String(nextElement.value || "").length
          : 0;
        nextElement.setSelectionRange(caretIndex, caretIndex);
      }
      return true;
    }

    function handleSpreadsheetNavigationKeydown(event) {
      if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
        return;
      }
      if (
        event.key !== "ArrowUp" &&
        event.key !== "ArrowDown" &&
        event.key !== "ArrowLeft" &&
        event.key !== "ArrowRight"
      ) {
        return;
      }
      const target = event.currentTarget;
      if (
        !(target instanceof HTMLInputElement) &&
        !(target instanceof HTMLTextAreaElement)
      ) {
        return;
      }
      if (canMoveCaretWithinElement(target, event.key)) {
        return;
      }
      if (moveEditingFocus(target, event.key)) {
        event.preventDefault();
      }
    }

    function arraysEqual(left, right) {
      if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
        return false;
      }
      for (let index = 0; index < left.length; index += 1) {
        if (left[index] !== right[index]) {
          return false;
        }
      }
      return true;
    }

    function applyIssueHoverTitle(element, title) {
      const nextTitle = title || "";
      const applyTitle = (target) => {
        if (!(target instanceof HTMLElement)) {
          return;
        }
        const baseTitle = target.dataset.baseTitle !== undefined
          ? target.dataset.baseTitle
          : target.title;
        target.dataset.baseTitle = baseTitle || "";
        if (nextTitle) {
          target.dataset.issueTitle = nextTitle;
        } else {
          delete target.dataset.issueTitle;
        }
        if (baseTitle) {
          target.title = baseTitle;
        } else {
          target.removeAttribute("title");
        }
      };
      applyTitle(element);
      element.querySelectorAll("input, textarea, .grid-editor, .header-wrap, button, .column-resizer").forEach((child) => {
        applyTitle(child);
      });
    }

    function clearValidationTooltipTimer() {
      if (validationTooltipTimer) {
        window.clearTimeout(validationTooltipTimer);
        validationTooltipTimer = undefined;
      }
    }

    function hideValidationTooltip() {
      clearValidationTooltipTimer();
      validationTooltipHost = undefined;
      pendingValidationTooltipHost = undefined;
      if (validationTooltip instanceof HTMLElement) {
        validationTooltip.hidden = true;
        validationTooltip.textContent = "";
      }
    }

    function getValidationTooltipHost(target) {
      if (!(target instanceof Element)) {
        return undefined;
      }
      const host = target.closest("[data-issue-title], [data-hover-hint]");
      if (!(host instanceof HTMLElement)) {
        return undefined;
      }
      return host.dataset.issueTitle || host.dataset.hoverHint
        ? host
        : undefined;
    }

    function positionValidationTooltip(host) {
      if (!(validationTooltip instanceof HTMLElement) || !(host instanceof HTMLElement)) {
        return;
      }
      const hostRect = host.getBoundingClientRect();
      const tooltipRect = validationTooltip.getBoundingClientRect();
      let left = typeof validationTooltipPointer?.clientX === "number"
        ? validationTooltipPointer.clientX + 12
        : hostRect.left + 8;
      let top = typeof validationTooltipPointer?.clientY === "number"
        ? validationTooltipPointer.clientY + 16
        : hostRect.bottom + 8;
      if (left + tooltipRect.width > window.innerWidth - 12) {
        left = Math.max(12, window.innerWidth - tooltipRect.width - 12);
      }
      if (top + tooltipRect.height > window.innerHeight - 12) {
        top = Math.max(12, hostRect.top - tooltipRect.height - 8);
      }
      validationTooltip.style.left = left + "px";
      validationTooltip.style.top = top + "px";
    }

    function showValidationTooltip(host) {
      if (!(validationTooltip instanceof HTMLElement) || !(host instanceof HTMLElement)) {
        return;
      }
      const message = host.dataset.issueTitle || host.dataset.hoverHint || "";
      if (!message) {
        hideValidationTooltip();
        return;
      }
      clearValidationTooltipTimer();
      pendingValidationTooltipHost = undefined;
      validationTooltipHost = host;
      validationTooltip.textContent = message;
      validationTooltip.hidden = false;
      positionValidationTooltip(host);
    }

    function scheduleValidationTooltip(host) {
      if (host === validationTooltipHost) {
        if (!validationTooltip?.hidden) {
          positionValidationTooltip(host);
        }
        return;
      }
      clearValidationTooltipTimer();
      pendingValidationTooltipHost = host;
      if (!host) {
        validationTooltipHost = undefined;
        if (validationTooltip instanceof HTMLElement) {
          validationTooltip.hidden = true;
          validationTooltip.textContent = "";
        }
        return;
      }
      validationTooltipHost = undefined;
      if (validationTooltip instanceof HTMLElement) {
        validationTooltip.hidden = true;
        validationTooltip.textContent = "";
      }
      validationTooltipTimer = window.setTimeout(() => {
        if (pendingValidationTooltipHost !== host) {
          return;
        }
        showValidationTooltip(host);
      }, VALIDATION_TOOLTIP_DELAY_MS);
    }

    function applyIssueClass(element, rowIndex, columnIndex) {
      element.classList.remove("cell-error");
      applyIssueHoverTitle(element, "");
      applyIssuePlaceholder(element, false);
      const cellKey = activeSheetIndex + ":" + rowIndex + ":" + columnIndex;
      if (state.validation?.issueMap?.[cellKey]) {
        element.classList.add("cell-error");
        applyIssueHoverTitle(element, state.validation.issueMap[cellKey].join("\\n"));
        applyIssuePlaceholder(element, true);
      }
    }

    function applyIssuePlaceholder(element, hasIssue) {
      if (!(element instanceof HTMLElement)) {
        return;
      }
      const editor = element.matches("input, textarea")
        ? element
        : element.querySelector("input, textarea");
      if (
        !(editor instanceof HTMLInputElement) &&
        !(editor instanceof HTMLTextAreaElement)
      ) {
        return;
      }
      const isEmpty = String(editor.value || "") === "";
      editor.placeholder = hasIssue && isEmpty ? "(empty)" : "";
    }

    function getDisplayRowLabel(sheet, rowIndex, recordDisplayNumberOverride) {
      const row = sheet.rows[rowIndex];
      if (isCommentRow(row)) {
        return "Comment";
      }
      if (typeof recordDisplayNumberOverride === "number") {
        return String(recordDisplayNumberOverride + 1);
      }

      let count = 0;
      for (let index = 0; index <= rowIndex; index += 1) {
        if (!isCommentRow(sheet.rows[index])) {
          count += 1;
        }
      }
      return String(count);
    }

    function getColumnLabel(columnIndex) {
      if (columnIndex === 0) {
        return "Record";
      }
      if (columnIndex === 1) {
        return "Type";
      }
      return getActiveSheet().headers[columnIndex] || ("Column " + (columnIndex + 1));
    }

    function getSpreadsheetColumnIndexLabel(columnIndex) {
      let value = Number(columnIndex);
      if (!Number.isInteger(value) || value < 0) {
        return "";
      }
      let label = "";
      do {
        label = String.fromCharCode(65 + (value % 26)) + label;
        value = Math.floor(value / 26) - 1;
      } while (value >= 0);
      return label;
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    function captureFocusState() {
      const activeElement = document.activeElement;
      if (
        !(activeElement instanceof HTMLInputElement) &&
        !(activeElement instanceof HTMLTextAreaElement)
      ) {
        return undefined;
      }
      if (!activeElement.dataset.scope || !activeElement.dataset.columnIndex) {
        return undefined;
      }
      return {
        scope: activeElement.dataset.scope,
        rowIndex: activeElement.dataset.rowIndex || "",
        columnIndex: activeElement.dataset.columnIndex || "",
        selectionStart: activeElement.selectionStart,
        selectionEnd: activeElement.selectionEnd,
      };
    }

    function restoreFocusState(focusState) {
      if (!focusState) {
        return;
      }
      const selector = focusState.scope === "header"
        ? 'input[data-scope="header"][data-column-index="' + focusState.columnIndex + '"]'
        : '[data-scope="' + focusState.scope + '"][data-row-index="' + focusState.rowIndex + '"][data-column-index="' + focusState.columnIndex + '"]';
      const input = document.querySelector(selector);
      if (
        !(input instanceof HTMLInputElement) &&
        !(input instanceof HTMLTextAreaElement)
      ) {
        return;
      }
      input.focus();
      if (
        typeof focusState.selectionStart === "number" &&
        typeof focusState.selectionEnd === "number"
      ) {
        input.setSelectionRange(focusState.selectionStart, focusState.selectionEnd);
      }
    }

    function renderWithFocusPreserved() {
      const focusState = captureFocusState();
      safeRender();
      restoreFocusState(focusState);
    }

    function getFocusedCellSelectionAnchor() {
      const activeElement = document.activeElement;
      if (!(activeElement instanceof HTMLInputElement) || activeElement.dataset.scope !== "cell") {
        return undefined;
      }
      const rowIndex = Number(activeElement.dataset.rowIndex);
      const columnIndex = Number(activeElement.dataset.columnIndex);
      if (!Number.isInteger(rowIndex) || !Number.isInteger(columnIndex)) {
        return undefined;
      }
      const row = getActiveSheet().rows[rowIndex];
      if (!row || isCommentRow(row)) {
        return undefined;
      }
      return {
        rowIndex,
        columnIndex,
      };
    }

    function getCellSelectionAnchor(defaultRowIndex, defaultColumnIndex) {
      return getFocusedCellSelectionAnchor() || {
        rowIndex: defaultRowIndex,
        columnIndex: defaultColumnIndex,
      };
    }

    function getRowSelectionPivot(rowIndex) {
      if (
        Number.isInteger(rowSelectionPivotIndex) &&
        selectedRowIndexes.includes(rowSelectionPivotIndex)
      ) {
        return rowSelectionPivotIndex;
      }
      if (selectedRowIndexes.length > 0) {
        return selectedRowIndexes[0];
      }
      return rowIndex;
    }

    function setSelectedRowIndexes(nextIndexes, pivotIndex = undefined) {
      clearCellSelection();
      clearColumnSelection();
      selectedRowIndexes = [...new Set((nextIndexes || []).filter((index) =>
        Number.isInteger(index) && index >= 0 && index < getActiveSheet().rows.length,
      ))].sort((left, right) => left - right);
      if (selectedRowIndexes.length === 0) {
        rowSelectionPivotIndex = undefined;
      } else if (
        Number.isInteger(pivotIndex) &&
        selectedRowIndexes.includes(pivotIndex)
      ) {
        rowSelectionPivotIndex = pivotIndex;
      } else if (!selectedRowIndexes.includes(rowSelectionPivotIndex)) {
        rowSelectionPivotIndex = selectedRowIndexes[0];
      }
      refreshRowSelectionStyles();
    }

    function setSelectedRowRange(start, end) {
      const min = Math.min(start, end);
      const max = Math.max(start, end);
      const nextIndexes = [];
      for (let index = min; index <= max; index += 1) {
        nextIndexes.push(index);
      }
      setSelectedRowIndexes(nextIndexes, start);
    }

    function ensureRowSelection(rowIndex) {
      if (!selectedRowIndexes.includes(rowIndex)) {
        setSelectedRowIndexes([rowIndex], rowIndex);
      } else if (!Number.isInteger(rowSelectionPivotIndex)) {
        rowSelectionPivotIndex = rowIndex;
      }
    }

    function getColumnSelectionPivot(columnIndex) {
      if (
        Number.isInteger(columnSelectionPivotIndex) &&
        selectedColumnIndexes.includes(columnSelectionPivotIndex)
      ) {
        return columnSelectionPivotIndex;
      }
      if (selectedColumnIndexes.length > 0) {
        return selectedColumnIndexes[0];
      }
      return columnIndex;
    }

    function setSelectedColumnIndexes(nextIndexes, pivotIndex = undefined) {
      clearRowSelection();
      clearCellSelection();
      selectedColumnIndexes = [...new Set((nextIndexes || []).filter((index) =>
        Number.isInteger(index) && index >= 0 && index < getActiveSheet().headers.length,
      ))].sort((left, right) => left - right);
      if (selectedColumnIndexes.length === 0) {
        columnSelectionPivotIndex = undefined;
      } else if (
        Number.isInteger(pivotIndex) &&
        selectedColumnIndexes.includes(pivotIndex)
      ) {
        columnSelectionPivotIndex = pivotIndex;
      } else if (!selectedColumnIndexes.includes(columnSelectionPivotIndex)) {
        columnSelectionPivotIndex = selectedColumnIndexes[0];
      }
      refreshColumnSelectionStyles();
    }

    function setSelectedColumnRange(start, end) {
      const min = Math.min(start, end);
      const max = Math.max(start, end);
      const nextIndexes = [];
      for (let index = min; index <= max; index += 1) {
        nextIndexes.push(index);
      }
      setSelectedColumnIndexes(nextIndexes, start);
    }

    function ensureColumnSelection(columnIndex) {
      if (!selectedColumnIndexes.includes(columnIndex)) {
        setSelectedColumnIndexes([columnIndex], columnIndex);
      } else if (!Number.isInteger(columnSelectionPivotIndex)) {
        columnSelectionPivotIndex = columnIndex;
      }
    }

    function refreshColumnSelectionStyles() {
      const selected = new Set(selectedColumnIndexes);
      document.querySelectorAll(".column-index-cell").forEach((cell) => {
        const columnIndex = Number(cell.dataset.columnIndex);
        cell.classList.toggle("selected", selected.has(columnIndex));
      });
      document.querySelectorAll(".field-header-cell").forEach((cell) => {
        const columnIndex = Number(cell.dataset.columnIndex);
        cell.classList.toggle("selected", selected.has(columnIndex));
      });
      document.querySelectorAll('td[data-column-index]').forEach((cell) => {
        const columnIndex = Number(cell.dataset.columnIndex);
        cell.classList.toggle("column-selected", selected.has(columnIndex));
      });
      refreshBackgroundPaletteState();
    }

    function refreshRowSelectionStyles() {
      const selected = new Set(selectedRowIndexes);
      document.querySelectorAll(".sheet-row").forEach((rowElement) => {
        const rowIndex = Number(rowElement.dataset.rowIndex);
        rowElement.classList.toggle("row-selected", selected.has(rowIndex));
      });
      document.querySelectorAll(".row-index-cell").forEach((cell) => {
        const rowIndex = Number(cell.dataset.rowIndex);
        cell.classList.toggle("selected", selected.has(rowIndex));
      });
      refreshBackgroundPaletteState();
    }

    function closeFloatingMenu() {
      floatingMenu.hidden = true;
      floatingMenu.innerHTML = "";
    }

    function populateFloatingMenu(items) {
      closeSuggestionMenu();
      closeFloatingMenu();
      items.forEach((item) => {
        const button = document.createElement("button");
        button.type = "button";
        if (item.swatchCss !== undefined || item.noColorSwatch) {
          const swatch = document.createElement("span");
          swatch.className = "menu-swatch" + (item.noColorSwatch ? " no-color" : "");
          if (item.swatchCss) {
            swatch.style.background = item.swatchCss;
          }
          button.appendChild(swatch);
        }
        const label = document.createElement("span");
        label.textContent = item.label;
        button.appendChild(label);
        button.disabled = !!item.disabled;
        button.addEventListener("click", async () => {
          if (item.disabled) {
            return;
          }
          closeFloatingMenu();
          try {
            await item.action();
          } catch (error) {
            console.error("Spreadsheet menu action failed:", error);
          }
        });
        floatingMenu.appendChild(button);
      });
      floatingMenu.hidden = false;
    }

    function openFloatingMenu(anchor, items) {
      populateFloatingMenu(items);
      positionMenu(floatingMenu, anchor);
    }

    function openFloatingMenuAtPosition(clientX, clientY, items) {
      populateFloatingMenu(items);
      positionMenuAtPoint(floatingMenu, clientX, clientY);
    }

    function openSheetRemovalConfirmationAtPosition(clientX, clientY, sheetIndex) {
      const sheets = state.workbook?.sheets || [];
      const sheet = sheets[sheetIndex];
      if (!sheet) {
        return;
      }

      closeSuggestionMenu();
      closeFloatingMenu();

      const sheetName = getSheetDisplayName(sheet, sheetIndex);
      const panel = document.createElement("div");
      panel.className = "menu-panel";

      const warning = document.createElement("div");
      warning.className = "menu-warning";
      warning.textContent = "Remove sheet";

      const label = document.createElement("div");
      label.className = "menu-label";
      label.textContent = 'Type "' + sheetName + '" to confirm removal.';

      const input = document.createElement("input");
      input.type = "text";
      input.className = "menu-text-input";
      input.placeholder = sheetName;
      input.setAttribute("aria-label", "Type the sheet name to confirm removal");

      const confirmButton = document.createElement("button");
      confirmButton.type = "button";
      confirmButton.className = "menu-danger";
      confirmButton.textContent = "Confirm remove";
      confirmButton.disabled = true;

      const cancelButton = document.createElement("button");
      cancelButton.type = "button";
      cancelButton.textContent = "Cancel";

      function updateConfirmationState() {
        confirmButton.disabled = input.value.trim() !== sheetName;
      }

      function closeConfirmation() {
        closeFloatingMenu();
      }

      function confirmRemoval() {
        if (confirmButton.disabled) {
          return;
        }
        closeFloatingMenu();
        removeSheet(sheetIndex);
      }

      input.addEventListener("input", updateConfirmationState);
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          confirmRemoval();
        } else if (event.key === "Escape") {
          event.preventDefault();
          closeConfirmation();
        }
      });
      confirmButton.addEventListener("click", confirmRemoval);
      cancelButton.addEventListener("click", closeConfirmation);

      panel.append(warning, label, input, confirmButton, cancelButton);
      floatingMenu.appendChild(panel);
      floatingMenu.hidden = false;
      positionMenuAtPoint(floatingMenu, clientX, clientY);
      window.requestAnimationFrame(() => {
        input.focus();
      });
    }

    function getEffectiveRowIndexes(rowIndex) {
      return selectedRowIndexes.includes(rowIndex)
        ? selectedRowIndexes.slice()
        : [rowIndex];
    }

    function getEffectiveColumnIndexes(columnIndex) {
      return selectedColumnIndexes.includes(columnIndex)
        ? selectedColumnIndexes.slice()
        : [columnIndex];
    }

    function applyBackgroundToSelectedCells(backgroundToken) {
      const range = getActiveSelectedCellRange();
      if (!range) {
        return;
      }
      recordUndoSnapshot();
      const sheet = getActiveSheet();
      const normalizedToken = normalizeBackgroundToken(backgroundToken);
      for (let rowIndex = range.startRowIndex; rowIndex <= range.endRowIndex; rowIndex += 1) {
        const row = sheet.rows[rowIndex];
        if (!row || isCommentRow(row)) {
          continue;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        for (let columnIndex = range.startColumnIndex; columnIndex <= range.endColumnIndex; columnIndex += 1) {
          row.backgrounds[columnIndex] = normalizedToken;
        }
      }
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function applyBackgroundToSingleCell(rowIndex, columnIndex, backgroundToken) {
      const sheet = getActiveSheet();
      const row = sheet.rows[rowIndex];
      if (!row || isCommentRow(row)) {
        return;
      }
      ensureRecordRowLength(row, sheet.headers.length);
      const normalizedToken = normalizeBackgroundToken(backgroundToken);
      if (normalizeBackgroundToken(row.backgrounds[columnIndex]) === normalizedToken) {
        return;
      }
      recordUndoSnapshot();
      row.backgrounds[columnIndex] = normalizedToken;
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function applyBackgroundToCommentRow(rowIndex, backgroundToken) {
      const sheet = getActiveSheet();
      const row = sheet.rows[rowIndex];
      if (!row || !isCommentRow(row)) {
        return;
      }
      const normalizedToken = normalizeBackgroundToken(backgroundToken);
      if (normalizeBackgroundToken(row.background) === normalizedToken) {
        return;
      }
      recordUndoSnapshot();
      row.background = normalizedToken;
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function applyBackgroundToRows(rowIndexes, backgroundToken) {
      if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) {
        return;
      }
      recordUndoSnapshot();
      const sheet = getActiveSheet();
      const normalizedToken = normalizeBackgroundToken(backgroundToken);
      rowIndexes.forEach((rowIndex) => {
        const row = sheet.rows[rowIndex];
        if (!row) {
          return;
        }
        if (isCommentRow(row)) {
          row.background = normalizedToken;
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        for (let columnIndex = 0; columnIndex < sheet.headers.length; columnIndex += 1) {
          row.backgrounds[columnIndex] = normalizedToken;
        }
      });
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function applyBackgroundToColumns(columnIndexes, backgroundToken) {
      if (!Array.isArray(columnIndexes) || columnIndexes.length === 0) {
        return;
      }
      recordUndoSnapshot();
      const sheet = getActiveSheet();
      const normalizedToken = normalizeBackgroundToken(backgroundToken);
      ensureHeaderBackgroundLength(sheet);
      columnIndexes.forEach((columnIndex) => {
        sheet.headerBackgrounds[columnIndex] = normalizedToken;
      });
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        columnIndexes.forEach((columnIndex) => {
          row.backgrounds[columnIndex] = normalizedToken;
        });
      });
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function getAvailableBackgroundOptions() {
      return BACKGROUND_COLOR_OPTIONS
        .filter((option) =>
          option.token !== "grey" &&
          option.token !== "pink" &&
          option.token !== "cyan" &&
          option.token !== "magenta"
        );
    }

    function getUniformToken(tokens) {
      const normalizedTokens = tokens.map((token) => normalizeBackgroundToken(token));
      if (normalizedTokens.length === 0) {
        return "";
      }
      return normalizedTokens.every((token) => token === normalizedTokens[0])
        ? normalizedTokens[0]
        : undefined;
    }

    function getFocusedBackgroundTarget() {
      const activeElement = document.activeElement;
      if (activeElement instanceof HTMLInputElement && activeElement.dataset.scope === "cell") {
        const rowIndex = Number(activeElement.dataset.rowIndex);
        const columnIndex = Number(activeElement.dataset.columnIndex);
        const sheet = getActiveSheet();
        const row = sheet.rows[rowIndex];
        if (!row || isCommentRow(row)) {
          return undefined;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        return {
          apply: (backgroundToken) => applyBackgroundToSingleCell(rowIndex, columnIndex, backgroundToken),
          token: normalizeBackgroundToken(row.backgrounds[columnIndex]),
        };
      }
      if (activeElement instanceof HTMLTextAreaElement && activeElement.dataset.scope === "comment") {
        const rowIndex = Number(activeElement.dataset.rowIndex);
        const row = getActiveSheet().rows[rowIndex];
        if (!row || !isCommentRow(row)) {
          return undefined;
        }
        return {
          apply: (backgroundToken) => applyBackgroundToCommentRow(rowIndex, backgroundToken),
          token: normalizeBackgroundToken(row.background),
        };
      }
      return undefined;
    }

    function getBackgroundToolbarTarget() {
      const cellRange = getActiveSelectedCellRange();
      if (cellRange) {
        const sheet = getActiveSheet();
        const tokens = [];
        for (let rowIndex = cellRange.startRowIndex; rowIndex <= cellRange.endRowIndex; rowIndex += 1) {
          const row = sheet.rows[rowIndex];
          if (!row || isCommentRow(row)) {
            continue;
          }
          ensureRecordRowLength(row, sheet.headers.length);
          for (let columnIndex = cellRange.startColumnIndex; columnIndex <= cellRange.endColumnIndex; columnIndex += 1) {
            tokens.push(row.backgrounds[columnIndex]);
          }
        }
        return {
          apply: (backgroundToken) => applyBackgroundToSelectedCells(backgroundToken),
          token: getUniformToken(tokens),
        };
      }
      if (selectedRowIndexes.length > 0) {
        const sheet = getActiveSheet();
        const tokens = [];
        selectedRowIndexes.forEach((rowIndex) => {
          const row = sheet.rows[rowIndex];
          if (!row) {
            return;
          }
          if (isCommentRow(row)) {
            tokens.push(row.background);
            return;
          }
          ensureRecordRowLength(row, sheet.headers.length);
          tokens.push(...row.backgrounds);
        });
        return {
          apply: (backgroundToken) => applyBackgroundToRows(selectedRowIndexes.slice(), backgroundToken),
          token: getUniformToken(tokens),
        };
      }
      if (selectedColumnIndexes.length > 0) {
        const sheet = getActiveSheet();
        ensureHeaderBackgroundLength(sheet);
        const tokens = [];
        selectedColumnIndexes.forEach((columnIndex) => {
          tokens.push(sheet.headerBackgrounds[columnIndex]);
        });
        sheet.rows.forEach((row) => {
          if (isCommentRow(row)) {
            return;
          }
          ensureRecordRowLength(row, sheet.headers.length);
          selectedColumnIndexes.forEach((columnIndex) => {
            tokens.push(row.backgrounds[columnIndex]);
          });
        });
        return {
          apply: (backgroundToken) => applyBackgroundToColumns(selectedColumnIndexes.slice(), backgroundToken),
          token: getUniformToken(tokens),
        };
      }
      return getFocusedBackgroundTarget();
    }

    function refreshBackgroundPaletteState() {
      if (!(backgroundPalette instanceof HTMLElement)) {
        return;
      }
      const target = getBackgroundToolbarTarget();
      const activeToken = target?.token;
      backgroundPalette.querySelectorAll("button[data-background-token]").forEach((button) => {
        if (!(button instanceof HTMLButtonElement)) {
          return;
        }
        const token = button.dataset.backgroundToken || "";
        button.disabled = !target;
        button.classList.toggle("is-active", typeof activeToken === "string" && activeToken === token);
      });
      backgroundPalette.title = target
        ? "Apply background color to the current selection."
        : "Select rows, columns, cells, or focus a cell/comment row to change background color.";
    }

    function renderBackgroundPalette() {
      if (!(backgroundPalette instanceof HTMLElement)) {
        return;
      }
      backgroundPalette.innerHTML = "";
      getAvailableBackgroundOptions().forEach((option) => {
        const button = document.createElement("button");
        button.type = "button";
        button.dataset.backgroundToken = option.token;
        button.title = option.label;
        const swatch = document.createElement("span");
        swatch.className = "menu-swatch" + (option.token ? "" : " no-color");
        if (option.css) {
          swatch.style.background = option.css;
        }
        button.appendChild(swatch);
        button.addEventListener("mousedown", (event) => {
          // Keep focus on the active cell/comment editor so focused-cell coloring works.
          event.preventDefault();
        });
        button.addEventListener("click", () => {
          const target = getBackgroundToolbarTarget();
          if (!target) {
            return;
          }
          target.apply(option.token);
        });
        backgroundPalette.appendChild(button);
      });
      refreshBackgroundPaletteState();
    }

    function insertRowsAt(index, rowsToInsert) {
      const sheet = getActiveSheet();
      const normalizedIndex = Math.max(0, Math.min(index, sheet.rows.length));
      const insertedRows = rowsToInsert.map((row) => {
        const nextRow = cloneRow(row);
        if (!isCommentRow(nextRow)) {
          ensureRecordRowLength(nextRow, sheet.headers.length);
        }
        return nextRow;
      });
      if (insertedRows.length === 0) {
        return;
      }
      recordUndoSnapshot();
      sheet.rows.splice(normalizedIndex, 0, ...insertedRows);
      setSelectedRowIndexes(insertedRows.map((_, offset) => normalizedIndex + offset));
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function insertEmptyRows(index, count) {
      const sheet = getActiveSheet();
      insertRowsAt(
        index,
        new Array(count).fill(undefined).map(() => createEmptyRecordRow(sheet.headers.length)),
      );
    }

    function insertCommentAt(index, text) {
      insertRowsAt(index, [createCommentRow(text)]);
    }

    function deleteRows(rowIndexes) {
      const sheet = getActiveSheet();
      const sorted = rowIndexes.slice().sort((left, right) => right - left);
      if (sorted.length === 0) {
        return;
      }
      recordUndoSnapshot();
      sorted.forEach((rowIndex) => {
        sheet.rows.splice(rowIndex, 1);
      });
      if (!sheet.rows.some((row) => !isCommentRow(row))) {
        sheet.rows.push(createEmptyRecordRow(sheet.headers.length));
      }
      const nextIndex = Math.min(sorted[sorted.length - 1] || 0, sheet.rows.length - 1);
      setSelectedRowIndexes(nextIndex >= 0 ? [nextIndex] : []);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function copyRows(rowIndexes) {
      const sheet = getActiveSheet();
      rowClipboard = rowIndexes
        .slice()
        .sort((left, right) => left - right)
        .map((rowIndex) => cloneRow(sheet.rows[rowIndex]));
    }

    function cutRows(rowIndexes) {
      copyRows(rowIndexes);
      deleteRows(rowIndexes);
    }

    function moveRows(rowIndexes, direction) {
      const sheet = getActiveSheet();
      const sorted = rowIndexes.slice().sort((left, right) => left - right);
      if (sorted.length === 0) {
        return;
      }

      const first = sorted[0];
      const last = sorted[sorted.length - 1];
      if (direction < 0 && first === 0) {
        return;
      }
      if (direction > 0 && last >= sheet.rows.length - 1) {
        return;
      }

      recordUndoSnapshot();
      const block = sheet.rows.splice(first, sorted.length);
      const nextIndex = direction < 0 ? first - 1 : first + 1;
      sheet.rows.splice(nextIndex, 0, ...block);
      setSelectedRowIndexes(block.map((_, offset) => nextIndex + offset));
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function clearRows(rowIndexes) {
      const sorted = [...new Set((rowIndexes || []).filter((rowIndex) =>
        Number.isInteger(rowIndex) && rowIndex >= 0 && rowIndex < getActiveSheet().rows.length,
      ))].sort((left, right) => left - right);
      if (sorted.length === 0) {
        return;
      }

      recordUndoSnapshot();
      const sheet = getActiveSheet();
      const affectedColumns = new Set();
      const affectedTypeRows = new Set();
      sorted.forEach((rowIndex) => {
        const row = sheet.rows[rowIndex];
        if (!row) {
          return;
        }
        if (isCommentRow(row)) {
          row.text = "";
          const commentInput = getCommentInput(rowIndex);
          if (commentInput) {
            commentInput.value = "";
          }
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        for (let columnIndex = 0; columnIndex < sheet.headers.length; columnIndex += 1) {
          row.values[columnIndex] = "";
          const input = getGridCellInput(rowIndex, columnIndex);
          if (input) {
            input.value = "";
          }
          affectedColumns.add(columnIndex);
          if (columnIndex === 1) {
            affectedTypeRows.add(rowIndex);
          }
        }
      });

      affectedTypeRows.forEach((rowIndex) => {
        refreshChoiceTriggersForRow(rowIndex);
      });
      affectedColumns.forEach((columnIndex) => {
        refreshAutoColumnWidth(columnIndex);
      });
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function getRowMenuItems(rowIndex) {
      const rowIndexes = getEffectiveRowIndexes(rowIndex);
      const first = rowIndexes[0];
      const last = rowIndexes[rowIndexes.length - 1];

      return [
        {
          label: "Add row below",
          action: () => insertEmptyRows(last + 1, 1),
        },
        {
          label: "Add 5 new rows below",
          action: () => insertEmptyRows(last + 1, 5),
        },
        {
          label: "Add 20 rows below",
          action: () => insertEmptyRows(last + 1, 20),
        },
        {
          label: "Add 100 rows below",
          action: () => insertEmptyRows(last + 1, 100),
        },
        {
          label: "Add 500 rows below",
          action: () => insertEmptyRows(last + 1, 500),
        },
        {
          label: "Add row above",
          action: () => insertEmptyRows(first, 1),
        },
        {
          label: "Add comment below",
          action: () => insertCommentAt(last + 1, ""),
        },
        {
          label: "Add comment above",
          action: () => insertCommentAt(first, ""),
        },
        {
          label: rowIndexes.length > 1 ? "Delete rows" : "Delete this row",
          action: () => deleteRows(rowIndexes),
        },
        {
          label: rowIndexes.length > 1 ? "Duplicate rows below" : "Duplicate this row",
          action: () => insertRowsAt(last + 1, rowIndexes.map((nextIndex) => cloneRow(getActiveSheet().rows[nextIndex]))),
        },
        {
          label: rowIndexes.length > 1 ? "Copy rows" : "Copy this row",
          action: () => copyRows(rowIndexes),
        },
        {
          label: rowIndexes.length > 1 ? "Cut rows" : "Cut this row",
          action: () => cutRows(rowIndexes),
        },
        {
          label: "Clear rows",
          action: () => clearRows(rowIndexes),
        },
        {
          label: "Insert copied rows above",
          disabled: !rowClipboard || rowClipboard.length === 0,
          action: () => insertRowsAt(first, rowClipboard || []),
        },
        {
          label: "Insert copied rows below",
          disabled: !rowClipboard || rowClipboard.length === 0,
          action: () => insertRowsAt(last + 1, rowClipboard || []),
        },
        {
          label: rowIndexes.length > 1 ? "Insert empty rows above" : "Insert empty row above",
          action: () => insertEmptyRows(first, rowIndexes.length),
        },
        {
          label: rowIndexes.length > 1 ? "Insert empty rows below" : "Insert empty row below",
          action: () => insertEmptyRows(last + 1, rowIndexes.length),
        },
        {
          label: rowIndexes.length > 1 ? "Move rows up" : "Move this row up",
          disabled: first === 0,
          action: () => moveRows(rowIndexes, -1),
        },
        {
          label: rowIndexes.length > 1 ? "Move rows down" : "Move this row down",
          disabled: last >= getActiveSheet().rows.length - 1,
          action: () => moveRows(rowIndexes, 1),
        },
      ];
    }

    function openRowMenu(anchor, rowIndex) {
      openFloatingMenu(
        anchor,
        getRowMenuItems(rowIndex),
      );
    }

    function openRowMenuAtPosition(clientX, clientY, rowIndex) {
      openFloatingMenuAtPosition(
        clientX,
        clientY,
        getRowMenuItems(rowIndex),
      );
    }

    function cloneColumn(sheet, columnIndex) {
      return {
        header: String(sheet.headers[columnIndex] || ""),
        headerBackground: normalizeBackgroundToken(sheet.headerBackgrounds?.[columnIndex]),
        values: sheet.rows.map((row) => {
          if (isCommentRow(row)) {
            return undefined;
          }
          ensureRecordRowLength(row, sheet.headers.length);
          return String(row.values[columnIndex] || "");
        }),
        backgrounds: sheet.rows.map((row) => {
          if (isCommentRow(row)) {
            return undefined;
          }
          ensureRecordRowLength(row, sheet.headers.length);
          return normalizeBackgroundToken(row.backgrounds[columnIndex]);
        }),
      };
    }

    function insertColumnsAt(index, columnsToInsert) {
      const sheet = getActiveSheet();
      const previousColumnCount = sheet.headers.length;
      const normalizedIndex = Math.max(2, Math.min(index, previousColumnCount));
      const columns = (columnsToInsert || []).map((column) => ({
        header: String(column?.header || ""),
        headerBackground: normalizeBackgroundToken(column?.headerBackground),
        values: Array.isArray(column?.values) ? column.values.slice() : [],
        backgrounds: Array.isArray(column?.backgrounds) ? column.backgrounds.slice() : [],
      }));
      if (columns.length === 0) {
        return;
      }

      recordUndoSnapshot();
      sheet.headers.splice(
        normalizedIndex,
        0,
        ...columns.map((column) => column.header),
      );
      ensureHeaderBackgroundLength(sheet);
      sheet.headerBackgrounds.splice(
        normalizedIndex,
        0,
        ...columns.map((column) => column.headerBackground),
      );

      sheet.rows.forEach((row, rowIndex) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, previousColumnCount);
        row.values.splice(
          normalizedIndex,
          0,
          ...columns.map((column) => String(column.values[rowIndex] || "")),
        );
        row.backgrounds.splice(
          normalizedIndex,
          0,
          ...columns.map((column) => normalizeBackgroundToken(column.backgrounds[rowIndex])),
        );
        ensureRecordRowLength(row, sheet.headers.length);
      });

      shiftManualColumnWidthsForInsert(normalizedIndex, columns.length);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function deleteColumn(columnIndex) {
      if (columnIndex < 2) {
        return;
      }
      const sheet = getActiveSheet();
      recordUndoSnapshot();
      sheet.headers.splice(columnIndex, 1);
      ensureHeaderBackgroundLength(sheet);
      sheet.headerBackgrounds.splice(columnIndex, 1);
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length + 1);
        row.values.splice(columnIndex, 1);
        row.backgrounds.splice(columnIndex, 1);
      });
      shiftManualColumnWidthsForDelete(columnIndex);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function moveColumn(columnIndex, direction) {
      if (columnIndex < 2) {
        return;
      }
      const sheet = getActiveSheet();
      const targetIndex = columnIndex + direction;
      if (targetIndex < 2 || targetIndex >= sheet.headers.length) {
        return;
      }

      recordUndoSnapshot();
      const header = sheet.headers.splice(columnIndex, 1)[0];
      sheet.headers.splice(targetIndex, 0, header);
      ensureHeaderBackgroundLength(sheet);
      const headerBackground = sheet.headerBackgrounds.splice(columnIndex, 1)[0];
      sheet.headerBackgrounds.splice(targetIndex, 0, headerBackground);
      sheet.rows.forEach((row) => {
        if (isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        const value = row.values.splice(columnIndex, 1)[0];
        row.values.splice(targetIndex, 0, value);
        const background = row.backgrounds.splice(columnIndex, 1)[0];
        row.backgrounds.splice(targetIndex, 0, background);
      });
      moveManualColumnWidth(columnIndex, targetIndex);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function clearColumns(columnIndexes) {
      const sheet = getActiveSheet();
      const normalizedIndexes = [...new Set((columnIndexes || []).filter((columnIndex) =>
        Number.isInteger(columnIndex) && columnIndex >= 0 && columnIndex < sheet.headers.length,
      ))].sort((left, right) => left - right);
      if (normalizedIndexes.length === 0) {
        return;
      }

      recordUndoSnapshot();
      normalizedIndexes.forEach((columnIndex) => {
        if (columnIndex >= 2) {
          sheet.headers[columnIndex] = "";
          const headerInput = getHeaderFieldInput(columnIndex);
          if (headerInput) {
            headerInput.value = "";
          }
        }
      });

      const affectedTypeRows = new Set();
      sheet.rows.forEach((row, rowIndex) => {
        if (!row || isCommentRow(row)) {
          return;
        }
        ensureRecordRowLength(row, sheet.headers.length);
        normalizedIndexes.forEach((columnIndex) => {
          row.values[columnIndex] = "";
          const input = getGridCellInput(rowIndex, columnIndex);
          if (input) {
            input.value = "";
          }
          if (columnIndex === 1) {
            affectedTypeRows.add(rowIndex);
          }
        });
      });

      normalizedIndexes.forEach((columnIndex) => {
        refreshChoiceTriggersForColumn(columnIndex);
        refreshAutoColumnWidth(columnIndex);
      });
      affectedTypeRows.forEach((rowIndex) => {
        refreshChoiceTriggersForRow(rowIndex);
      });
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function getColumnMenuItems(columnIndex) {
      const sheet = getActiveSheet();
      const columnIndexes = getEffectiveColumnIndexes(columnIndex);
      const canAddColumnLeft = columnIndex >= 2;
      const canAddColumnRight = columnIndex >= 1;
      return [
        {
          label: "Add column left",
          disabled: !canAddColumnLeft,
          action: () => insertColumnsAt(Math.max(2, columnIndex), [{ header: "", values: [] }]),
        },
        {
          label: "Add column right",
          disabled: !canAddColumnRight,
          action: () => insertColumnsAt(columnIndex >= 2 ? columnIndex + 1 : 2, [{ header: "", values: [] }]),
        },
        {
          label: "Delete this column",
          disabled: columnIndex < 2,
          action: () => deleteColumn(columnIndex),
        },
        {
          label: "Duplicate this column",
          disabled: columnIndex < 2,
          action: () => insertColumnsAt(columnIndex + 1, [cloneColumn(sheet, columnIndex)]),
        },
        {
          label: "Copy this column",
          disabled: columnIndex < 2,
          action: () => {
            columnClipboard = [cloneColumn(sheet, columnIndex)];
          },
        },
        {
          label: "Cut this column",
          disabled: columnIndex < 2,
          action: () => {
            columnClipboard = [cloneColumn(sheet, columnIndex)];
            deleteColumn(columnIndex);
          },
        },
        {
          label: "Clear columns",
          action: () => clearColumns(columnIndexes),
        },
        {
          label: "Insert copied columns left",
          disabled: !columnClipboard || columnClipboard.length === 0,
          action: () => insertColumnsAt(Math.max(2, columnIndex), columnClipboard || []),
        },
        {
          label: "Insert copied columns right",
          disabled: !columnClipboard || columnClipboard.length === 0,
          action: () => insertColumnsAt(Math.max(2, columnIndex + 1), columnClipboard || []),
        },
        {
          label: "Move this column left",
          disabled: columnIndex <= 2,
          action: () => moveColumn(columnIndex, -1),
        },
        {
          label: "Move this column right",
          disabled: columnIndex < 2 || columnIndex >= sheet.headers.length - 1,
          action: () => moveColumn(columnIndex, 1),
        },
      ];
    }

    function openColumnMenu(anchor, columnIndex) {
      openFloatingMenu(
        anchor,
        getColumnMenuItems(columnIndex),
      );
    }

    function openColumnMenuAtPosition(clientX, clientY, columnIndex) {
      openFloatingMenuAtPosition(
        clientX,
        clientY,
        getColumnMenuItems(columnIndex),
      );
    }

    function getCellMenuChoices(sheet, rowIndex, columnIndex) {
      if (columnIndex < 2) {
        return [];
      }
      const row = sheet.rows[rowIndex];
      if (isCommentRow(row)) {
        return [];
      }
      const recordType = String(row.values[1] || "").trim();
      const fieldName = String(sheet.headers[columnIndex] || "").trim();
      return state.menuChoicesByRecordType?.[recordType]?.[fieldName] || [];
    }

    function setCellValue(rowIndex, columnIndex, value) {
      const sheet = getActiveSheet();
      const row = sheet.rows[rowIndex];
      if (!row || isCommentRow(row)) {
        return;
      }
      ensureRecordRowLength(row, sheet.headers.length);
      const nextValue = String(value || "");
      if (row.values[columnIndex] === nextValue) {
        return;
      }
      recordUndoSnapshot();
      row.values[columnIndex] = nextValue;
      refreshAutoColumnWidth(columnIndex);
      queueWorkbookSave();
      renderWithFocusPreserved();
    }

    function openChoiceMenu(anchor, rowIndex, columnIndex, choices) {
      openFloatingMenu(
        anchor,
        choices.map((choice) => ({
          label: choice,
          action: () => setCellValue(rowIndex, columnIndex, choice),
        })),
      );
    }

    resizeColumnsButton.addEventListener("click", () => {
      resizeAllAutoColumns();
    });

    fontSizeSelect?.addEventListener("change", () => {
      applyFontSize(fontSizeSelect.value, { rerender: true });
    });

    densitySelect?.addEventListener("change", () => {
      applyDensityMode(densitySelect.value, { rerender: true });
    });

    function postWorkbookAction(type, options = {}) {
      flushWorkbookSave();
      vscode.postMessage({
        type,
        workbook: buildWorkbookSnapshot(),
        currentSheetOnly: !!options.currentSheetOnly,
        activeSheetIndex,
        associateAfterSave: options.associateAfterSave !== false,
      });
    }

    previewDbButton.addEventListener("click", () => {
      if (previewDbButton.getAttribute("aria-disabled") === "true") {
        return;
      }
      if (state.sourceKind === "database") {
        vscode.postMessage({ type: "showDatabaseFile" });
        return;
      }
      postWorkbookAction("previewSpreadsheetAsDatabase", { currentSheetOnly: true });
    });

    undoButton.addEventListener("click", () => {
      undoWorkbookChange();
    });

    redoButton.addEventListener("click", () => {
      redoWorkbookChange();
    });

    saveButton.addEventListener("click", () => {
      postWorkbookAction("saveSpreadsheet");
    });

    saveAsButton.addEventListener("click", () => {
      postWorkbookAction(
        state.sourceKind === "database"
          ? "saveSpreadsheetAsDatabase"
          : "saveSpreadsheetAsExcel",
      );
    });

    saveDbButton.addEventListener("click", () => {
      if (saveDbButton.getAttribute("aria-disabled") === "true") {
        return;
      }
      if (state.sourceKind === "database") {
        postWorkbookAction("saveSpreadsheetAsExcel", { associateAfterSave: false });
        return;
      }
      postWorkbookAction("saveSpreadsheetAsDatabase", { currentSheetOnly: true });
    });

    document.addEventListener("mouseup", () => {
      isSelectingRows = false;
      isSelectingColumns = false;
      columnSelectionAnchorIndex = undefined;
      isSelectingCells = false;
      cellSelectionAnchor = undefined;
      finishReorderDrag();
    });

    window.addEventListener("mousemove", (event) => {
      updateColumnResize(event.clientX);
    });

    window.addEventListener("mouseup", () => {
      stopColumnResize();
      isSelectingColumns = false;
      columnSelectionAnchorIndex = undefined;
      isSelectingCells = false;
      cellSelectionAnchor = undefined;
      finishReorderDrag();
    });

    document.addEventListener("mousedown", (event) => {
      hideValidationTooltip();
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        closeFloatingMenu();
        closeSuggestionMenu();
        return;
      }
      if (!(floatingMenu.contains(target) || target.closest(".menu-trigger") || target.closest(".choice-trigger"))) {
        closeFloatingMenu();
      }
      if (!(suggestionMenu.contains(target) || isSuggestionInput(target.closest("input")))) {
        closeSuggestionMenu();
      }
    });

    document.addEventListener("focusin", () => {
      refreshBackgroundPaletteState();
    });

    document.addEventListener("keydown", (event) => {
      handleMultiCellShortcutKeydown(event);
    });

    document.addEventListener("mousemove", (event) => {
      validationTooltipPointer = {
        clientX: event.clientX,
        clientY: event.clientY,
      };
      const host = getValidationTooltipHost(event.target);
      if (host === validationTooltipHost) {
        positionValidationTooltip(host);
        return;
      }
      if (host !== pendingValidationTooltipHost) {
        scheduleValidationTooltip(host);
      }
    });

    document.addEventListener("mouseout", (event) => {
      const currentHost = getValidationTooltipHost(event.target);
      const nextHost = getValidationTooltipHost(event.relatedTarget);
      if (currentHost && currentHost !== nextHost) {
        hideValidationTooltip();
      }
      if (event.relatedTarget) {
        return;
      }
      hideValidationTooltip();
    });

    window.addEventListener("resize", () => {
      stopColumnResize();
      closeFloatingMenu();
      closeSuggestionMenu();
      hideValidationTooltip();
    });

    document.addEventListener("scroll", (event) => {
      const target = event.target;
      if (
        target instanceof HTMLElement &&
        (
          target === floatingMenu ||
          target === suggestionMenu ||
          floatingMenu.contains(target) ||
          suggestionMenu.contains(target)
        )
      ) {
        return;
      }
      closeFloatingMenu();
      closeSuggestionMenu();
      hideValidationTooltip();
    }, true);

    window.addEventListener("message", (event) => {
      const message = event.data;
      if (message?.type !== "spreadsheetState") {
        return;
      }
      const focusState = message.preserveEditing ? captureFocusState() : undefined;
      const nextMessageState = message.preserveEditing
        ? {
            ...message,
            workbook: state.workbook,
          }
        : message;
      state = {
        ...state,
        ...nextMessageState,
      };
      if (!message.preserveEditing) {
        resetHistoryState();
      }
      if (message.preserveEditing) {
        closeFloatingMenu();
        refreshToolbarState();
        renderValidationBadge();
        refreshValidationStyles();
        if (cellSelectionEnhancementsEnabled) {
          try {
            refreshCellSelectionStyles();
            applyRenderWarningBadge();
          } catch (error) {
            disableCellSelectionEnhancements(error);
            safeRender();
            restoreFocusState(focusState);
            return;
          }
        }
        restoreFocusState(focusState);
        refreshSuggestionMenu();
        return;
      }
      safeRender();
      restoreFocusState(focusState);
    });

    initializeFontSizeOptions();
    renderBackgroundPalette();
    applyDensityMode(densityMode, { persist: false });
    applyFontSize(fontSizePx, { persist: false });
    safeRender();
    } catch (error) {
      showStartupError("Startup error", error);
    }
  </script>
</body>
</html>`;
}

function serializeForWebview(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

module.exports = {
  registerDatabaseSpreadsheetEditor,
};
