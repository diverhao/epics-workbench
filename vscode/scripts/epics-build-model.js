const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const DATABASE_EXTENSIONS = new Set([".db", ".vdb", ".template"]);
const SUBSTITUTION_EXTENSIONS = new Set([".sub", ".subs", ".substitutions"]);
const DBD_EXTENSIONS = new Set([".dbd"]);
const STARTUP_EXTENSIONS = new Set([".cmd", ".iocsh"]);
const IGNORED_DIRECTORY_NAMES = new Set([
  ".git",
  ".hg",
  ".svn",
  "node_modules",
  "out",
  "dist",
]);
const DEFAULT_MAKE_TIMEOUT_MS = 5000;

function normalizeFsPath(value) {
  if (!value) {
    return "";
  }

  return path.resolve(String(value)).replace(/\\/g, "/");
}

function normalizePath(value) {
  return String(value || "").replace(/\\/g, "/");
}

function compareLabels(left, right) {
  return String(left || "").localeCompare(String(right || ""), undefined, {
    numeric: true,
    sensitivity: "base",
  });
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
  try {
    return fs.statSync(filePath).isFile();
  } catch (error) {
    return false;
  }
}

function isExistingDirectory(directoryPath) {
  try {
    return fs.statSync(directoryPath).isDirectory();
  } catch (error) {
    return false;
  }
}

function readTextFile(filePath) {
  try {
    return fs.readFileSync(filePath, "utf8");
  } catch (error) {
    return undefined;
  }
}

function stripOptionalQuotes(value) {
  const text = String(value || "").trim();
  if (
    (text.startsWith('"') && text.endsWith('"')) ||
    (text.startsWith("'") && text.endsWith("'"))
  ) {
    return text.slice(1, -1);
  }
  return text;
}

function stripMakeComments(line) {
  let inSingleQuote = false;
  let inDoubleQuote = false;
  let escaped = false;

  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (escaped) {
      escaped = false;
      continue;
    }

    if (character === "\\") {
      escaped = true;
      continue;
    }

    if (character === "'" && !inDoubleQuote) {
      inSingleQuote = !inSingleQuote;
      continue;
    }

    if (character === '"' && !inSingleQuote) {
      inDoubleQuote = !inDoubleQuote;
      continue;
    }

    if (character === "#" && !inSingleQuote && !inDoubleQuote) {
      return line.slice(0, index);
    }
  }

  return line;
}

function splitMakeValue(value) {
  const text = String(value || "").trim();
  if (!text) {
    return [];
  }

  const values = [];
  let current = "";
  let inSingleQuote = false;
  let inDoubleQuote = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (escaped) {
      current += character;
      escaped = false;
      continue;
    }

    if (character === "\\") {
      current += character;
      escaped = true;
      continue;
    }

    if (character === "'" && !inDoubleQuote) {
      inSingleQuote = !inSingleQuote;
      current += character;
      continue;
    }

    if (character === '"' && !inSingleQuote) {
      inDoubleQuote = !inDoubleQuote;
      current += character;
      continue;
    }

    if (/\s/.test(character) && !inSingleQuote && !inDoubleQuote) {
      if (current) {
        values.push(stripOptionalQuotes(current));
        current = "";
      }
      continue;
    }

    current += character;
  }

  if (current) {
    values.push(stripOptionalQuotes(current));
  }

  return values;
}

function parseMakeAssignments(text) {
  const assignments = new Map();
  const lines = String(text || "").replace(/\\\r?\n/g, " ").split(/\r?\n/);

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
    const existingValues = assignments.get(variableName) || [];

    if (operator === "?=" && existingValues.length > 0) {
      continue;
    }

    if (operator === "+=") {
      assignments.set(variableName, [...existingValues, ...values]);
      continue;
    }

    assignments.set(variableName, values);
  }

  return assignments;
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
    const rawValue = String(releaseVariables.get(variableName) || "");
    const expandedValue = rawValue.replace(
      /\$\(([^)]+)\)|\$\{([^}]+)\}/g,
      (_, parenthesizedName, bracedName) => {
        const nestedName = parenthesizedName || bracedName;
        return resolveVariable(nestedName) || process.env[nestedName] || "";
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
    if (!resolvedPath || !isExistingDirectory(resolvedPath)) {
      continue;
    }

    roots.push({
      variableName,
      rootPath: normalizeFsPath(resolvedPath),
    });
  }

  return roots;
}

function computeAbsoluteVariablePath(value, rootPath) {
  const normalizedValue = stripOptionalQuotes(value);
  if (!normalizedValue) {
    return undefined;
  }

  return normalizeFsPath(
    path.isAbsolute(normalizedValue)
      ? normalizedValue
      : path.resolve(rootPath, normalizedValue),
  );
}

function loadReleaseVariablesWithSources(rootPath) {
  const normalizedRootPath = normalizeFsPath(rootPath);
  const rawValues = new Map([["TOP", normalizedRootPath]]);
  const sources = new Map();
  const visitedFiles = new Set();
  const resolvedCache = new Map();
  const releaseEntryPaths = [
    normalizeFsPath(path.join(normalizedRootPath, "configure", "RELEASE")),
    normalizeFsPath(path.join(normalizedRootPath, "configure", "RELEASE.local")),
  ];

  const normalizeResolvedReleaseValue = (value) => {
    const normalizedValue = stripOptionalQuotes(String(value || "").trim());
    if (!normalizedValue || normalizedValue.includes("$(") || normalizedValue.includes("${")) {
      return normalizedValue;
    }
    return computeAbsoluteVariablePath(normalizedValue, normalizedRootPath) || normalizedValue;
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

  const expandReleaseInlineValue = (rawValue) =>
    String(rawValue || "").replace(
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

function scanDbdDirectoryEntries(entries, dbdRootPath, sourceLabel) {
  const normalizedDbdRootPath = normalizeFsPath(dbdRootPath);
  if (!isExistingDirectory(normalizedDbdRootPath)) {
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

function scanLibraryDirectoryEntries(entries, libRootPath, sourceLabel) {
  const normalizedLibRootPath = normalizeFsPath(libRootPath);
  if (!isExistingDirectory(normalizedLibRootPath)) {
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
      absolutePath: isExistingFile(artifact.absoluteRuntimePath)
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

function scanLocalMakefileDirectoryDbdEntries(entries, rootPath, sourceLabel) {
  const normalizedRootPath = normalizeFsPath(rootPath);
  if (!normalizedRootPath || !isExistingDirectory(normalizedRootPath)) {
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
        IGNORED_DIRECTORY_NAMES.has(entry.name) ||
        entry.name === "dbd" ||
        /^O(?:\.|$)/.test(entry.name)
      ) {
        continue;
      }
      pendingDirectories.push(normalizeFsPath(path.join(directoryPath, entry.name)));
    }
  }
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

function findProjectIocBootFilePaths(rootPath) {
  const iocBootPath = normalizeFsPath(path.join(rootPath, "iocBoot"));
  if (!isExistingDirectory(iocBootPath)) {
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

function isConcreteMakefileReferenceToken(token, kind) {
  if (!token || token.includes("$(") || token.includes("${") || token === "-nil-") {
    return false;
  }

  if (kind === "dbd") {
    return token.toLowerCase().endsWith(".dbd");
  }

  if (kind === "dbFile") {
    return /\.(db|vdb|template|sub|subs|substitutions)$/i.test(token);
  }

  if (kind === "iocName") {
    return !token.startsWith("-");
  }

  return true;
}

function parseConcreteVariableItems(variableText, kind) {
  return splitMakeValue(variableText).filter((token) =>
    isConcreteMakefileReferenceToken(token, kind),
  );
}

function findRelevantMakefiles(rootPath) {
  const normalizedRootPath = normalizeFsPath(rootPath);
  const results = [];
  const pendingDirectories = [normalizedRootPath];

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
        if (
          IGNORED_DIRECTORY_NAMES.has(entry.name) ||
          /^O(?:\.|$)/.test(entry.name)
        ) {
          continue;
        }
        pendingDirectories.push(absolutePath);
        continue;
      }

      if (!entry.isFile() || entry.name !== "Makefile") {
        continue;
      }

      const relativePath = normalizePath(path.relative(normalizedRootPath, absolutePath));
      if (/^[^/]+App\/src\/Makefile$/i.test(relativePath)) {
        results.push({ absolutePath, relativePath, kind: "src" });
        continue;
      }

      if (/^[^/]+App\/(?:Db|db)\/Makefile$/i.test(relativePath)) {
        results.push({ absolutePath, relativePath, kind: "db" });
      }
    }
  }

  return results.sort((left, right) => compareLabels(left.relativePath, right.relativePath));
}

function parseMakePrintDatabaseVariables(stdout) {
  const variables = new Map();
  for (const line of String(stdout || "").split(/\r?\n/)) {
    const match = line.match(/^([A-Za-z0-9_.-]+)\s*(?::=|\+=|\?=|=)\s*(.*)$/);
    if (!match) {
      continue;
    }
    variables.set(match[1], match[2]);
  }
  return variables;
}

function runMakePrintDatabase(makefilePath, timeoutMs = DEFAULT_MAKE_TIMEOUT_MS) {
  return new Promise((resolve) => {
    const workingDirectory = path.dirname(makefilePath);
    const makefileName = path.basename(makefilePath);
    let stdout = "";
    let stderr = "";
    let settled = false;

    const child = spawn("make", ["-pn", "-C", workingDirectory, "-f", makefileName], {
      env: process.env,
    });

    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      child.kill("SIGTERM");
      resolve({
        stdout,
        stderr,
        exitCode: undefined,
        timedOut: true,
        error: `Timed out after ${timeoutMs}ms.`,
      });
    }, timeoutMs);

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    child.on("error", (error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      resolve({
        stdout,
        stderr,
        exitCode: undefined,
        timedOut: false,
        error: error.message,
      });
    });
    child.on("close", (exitCode) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      resolve({
        stdout,
        stderr,
        exitCode,
        timedOut: false,
        error: undefined,
      });
    });
  });
}

function mapToSortedObject(map) {
  const entries = [...(map instanceof Map ? map.entries() : Object.entries(map || {}))].sort(
    (left, right) => compareLabels(left[0], right[0]),
  );
  return Object.fromEntries(entries);
}

function collectApplicationRuntimeArtifacts(rootPath, interrogatedMakefiles, buildInfo) {
  const iocsByName = new Map();
  const runtimeArtifacts = [];
  const seenArtifacts = new Set();

  const addArtifact = (artifact) => {
    const key = `${artifact.kind}:${artifact.runtimeRelativePath}`;
    if (seenArtifacts.has(key)) {
      return;
    }
    seenArtifacts.add(key);
    runtimeArtifacts.push(artifact);
  };

  return Promise.all(
    interrogatedMakefiles.map(async (makefile) => {
      const makeResult = await runMakePrintDatabase(makefile.absolutePath);
      buildInfo.interrogatedMakefiles.push({
        relativePath: makefile.relativePath,
        exitCode: makeResult.exitCode,
        timedOut: makeResult.timedOut,
        error: makeResult.error,
      });

      if (!makeResult.stdout) {
        return;
      }

      const variables = parseMakePrintDatabaseVariables(makeResult.stdout);
      const appDirName = makefile.relativePath.split("/")[0];

      if (makefile.kind === "src") {
        const iocNames = parseConcreteVariableItems(variables.get("PROD_IOC"), "iocName");
        const dbdNames = parseConcreteVariableItems(variables.get("DBD"), "dbd");

        for (const iocName of iocNames) {
          if (iocsByName.has(iocName)) {
            continue;
          }
          iocsByName.set(iocName, {
            name: iocName,
            appDirName,
            makefileRelativePath: makefile.relativePath,
            registerFunctionName: `${iocName}_registerRecordDeviceDriver`,
          });
        }

        for (const dbdName of dbdNames) {
          addArtifact(
            createRuntimeArtifact({
              rootPath,
              appDirName,
              kind: "dbd",
              runtimeRelativePath: normalizePath(path.posix.join("dbd", dbdName)),
              sourceRelativePath: makefile.relativePath,
              detail: `Generated DBD from ${makefile.relativePath}`,
              documentation: `Interrogated from ${makefile.relativePath} with make -pn`,
            }),
          );
        }
        return;
      }

      if (makefile.kind === "db") {
        const rawText = readTextFile(makefile.absolutePath) || "";
        const installedDbNames = parseConcreteVariableItems(variables.get("DB"), "dbFile");
        const templateMappings = extractTemplateMappings(parseMakeAssignments(rawText));
        const dbDirRelativePath = normalizePath(path.posix.join(appDirName, path.basename(path.dirname(makefile.relativePath))));

        for (const installedDbName of installedDbNames) {
          const artifactKind = getRuntimeArtifactKind(installedDbName);
          if (!artifactKind) {
            continue;
          }

          const sourceRelativePath = normalizePath(
            path.posix.join(
              dbDirRelativePath,
              resolveDatabaseSourceFileName(installedDbName, templateMappings),
            ),
          );
          addArtifact(
            createRuntimeArtifact({
              rootPath,
              appDirName,
              kind: artifactKind,
              runtimeRelativePath: normalizePath(path.posix.join("db", installedDbName)),
              sourceRelativePath,
              detail: `Installed ${artifactKind} from ${sourceRelativePath}`,
              documentation: `Interrogated from ${makefile.relativePath} with make -pn`,
            }),
          );
        }
      }
    }),
  ).then(() => ({
    iocs: [...iocsByName.values()].sort((left, right) => compareLabels(left.name, right.name)),
    runtimeArtifacts: runtimeArtifacts.sort((left, right) =>
      compareLabels(
        `${left.kind}:${left.runtimeRelativePath}`,
        `${right.kind}:${right.runtimeRelativePath}`,
      ),
    ),
  }));
}

async function collectEpicsBuildApplication(rootPath, options = {}) {
  const normalizedRootPath = normalizeFsPath(rootPath);
  if (!normalizedRootPath || !isExistingDirectory(normalizedRootPath)) {
    return undefined;
  }

  const releaseData = loadReleaseVariablesWithSources(normalizedRootPath);
  const interrogatedMakefiles = findRelevantMakefiles(normalizedRootPath);
  const buildInfo = {
    generator: "make -pn",
    makeTimeoutMs: Number(options.timeoutMs) || DEFAULT_MAKE_TIMEOUT_MS,
    interrogatedMakefiles: [],
    errors: [],
  };

  const { iocs, runtimeArtifacts } = await collectApplicationRuntimeArtifacts(
    normalizedRootPath,
    interrogatedMakefiles,
    buildInfo,
  );

  for (const makefileInfo of buildInfo.interrogatedMakefiles) {
    if (!makefileInfo.error && !makefileInfo.timedOut) {
      continue;
    }
    buildInfo.errors.push(
      `${makefileInfo.relativePath}: ${
        makefileInfo.error || `make exited with code ${makefileInfo.exitCode}`
      }`,
    );
  }

  const releaseRoots = resolveReleaseModuleRoots(normalizedRootPath, releaseData.values);
  const relativeRootPath = normalizePath(path.basename(normalizedRootPath));
  const availableDbds = [
    ...buildProjectDbdEntries({
      rootPath: normalizedRootPath,
      runtimeArtifacts,
      releaseRoots,
      rootRelativePath: relativeRootPath,
    }).values(),
  ].sort((left, right) => compareLabels(left.name, right.name));
  const availableLibs = [
    ...buildProjectLibEntries({
      rootPath: normalizedRootPath,
      releaseRoots,
      rootRelativePath: relativeRootPath,
    }).values(),
  ]
    .sort((left, right) => compareLabels(left.name, right.name))
    .map((entry) => {
      const clone = { ...entry };
      delete clone.priority;
      return clone;
    });

  return {
    rootPath: normalizedRootPath,
    releaseVariables: mapToSortedObject(releaseData.values),
    iocs,
    runtimeArtifacts,
    availableDbds,
    availableLibs,
    startupEntryPoints: findProjectIocBootFilePaths(normalizedRootPath),
    buildInfo,
  };
}

function parseCliArguments(argv) {
  const options = {};

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--root") {
      options.rootPath = argv[index + 1];
      index += 1;
      continue;
    }

    if (argument === "--timeout-ms") {
      options.timeoutMs = Number(argv[index + 1]) || DEFAULT_MAKE_TIMEOUT_MS;
      index += 1;
    }
  }

  return options;
}

async function runCli() {
  const options = parseCliArguments(process.argv.slice(2));
  if (!options.rootPath) {
    process.stderr.write("Usage: node epics-build-model.js --root <epics-root>\n");
    process.exitCode = 1;
    return;
  }

  const result = await collectEpicsBuildApplication(options.rootPath, options);
  process.stdout.write(`${JSON.stringify(result || null, null, 2)}\n`);
}

module.exports = {
  collectEpicsBuildApplication,
};

if (require.main === module) {
  void runCli();
}
