#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const dbdRoot = process.argv[2] || process.env.EPICS_DBD_ROOT;

if (!dbdRoot) {
  console.error(
    "Usage: node scripts/generate-embedded-record-field-types.js /path/to/epics-base/dbd",
  );
  process.exit(1);
}

const normalizedDbdRoot = path.resolve(dbdRoot);
const dbdFilesByName = new Map();

for (const fileName of fs.readdirSync(normalizedDbdRoot)) {
  if (path.extname(fileName).toLowerCase() !== ".dbd") {
    continue;
  }

  dbdFilesByName.set(fileName, path.join(normalizedDbdRoot, fileName));
}

const output = {};

for (const [fileName, filePath] of [...dbdFilesByName.entries()].sort()) {
  const recordTypeMatch = fileName.match(/^([A-Za-z0-9_]+)Record\.dbd$/);
  if (!recordTypeMatch) {
    continue;
  }

  const recordType = recordTypeMatch[1];
  const fieldTypes = extractFieldTypesForRecordTypeFromDbdFile(
    filePath,
    recordType,
    new Set(),
  );
  const sortedEntries = [...fieldTypes.entries()].sort(([left], [right]) =>
    left.localeCompare(right),
  );
  if (sortedEntries.length > 0) {
    output[recordType] = Object.fromEntries(sortedEntries);
  }
}

process.stdout.write(`${JSON.stringify(output, null, 2)}\n`);

function extractFieldTypesForRecordTypeFromDbdFile(filePath, recordType, visitedPaths) {
  const normalizedPath = path.resolve(filePath);
  if (visitedPaths.has(normalizedPath)) {
    return new Map();
  }

  visitedPaths.add(normalizedPath);

  let text;
  try {
    text = fs.readFileSync(normalizedPath, "utf8");
  } catch (error) {
    return new Map();
  }

  const fieldTypes = new Map();
  const recordTypeRegex = new RegExp(
    `recordtype\\(\\s*${escapeRegex(recordType)}\\s*\\)\\s*\\{`,
    "g",
  );
  let match;

  while ((match = recordTypeRegex.exec(text))) {
    const block = readBalancedBlock(text, recordTypeRegex.lastIndex - 1);
    if (!block) {
      continue;
    }

    addStandaloneDbdFieldTypesAndIncludes(
      fieldTypes,
      block.body,
      normalizedPath,
      visitedPaths,
    );
    recordTypeRegex.lastIndex = block.endIndex;
  }

  return fieldTypes;
}

function addStandaloneDbdFieldTypesAndIncludes(
  fieldTypes,
  text,
  sourcePath,
  visitedPaths,
) {
  const fieldRegex = /field\(\s*([A-Z0-9_]+)\s*,\s*(DBF_[A-Z0-9_]+)\s*\)/g;
  let fieldMatch;
  while ((fieldMatch = fieldRegex.exec(text))) {
    if (!fieldTypes.has(fieldMatch[1])) {
      fieldTypes.set(fieldMatch[1], fieldMatch[2]);
    }
  }

  for (const includePath of extractDbdIncludePaths(text)) {
    const resolvedPath = resolveDbdIncludePath(sourcePath, includePath);
    if (!resolvedPath) {
      continue;
    }

    const normalizedResolvedPath = path.resolve(resolvedPath);
    if (visitedPaths.has(normalizedResolvedPath)) {
      continue;
    }

    visitedPaths.add(normalizedResolvedPath);

    let includedText;
    try {
      includedText = fs.readFileSync(normalizedResolvedPath, "utf8");
    } catch (error) {
      continue;
    }

    addStandaloneDbdFieldTypesAndIncludes(
      fieldTypes,
      includedText,
      normalizedResolvedPath,
      visitedPaths,
    );
  }
}

function extractDbdIncludePaths(text) {
  const includePaths = [];
  const includeRegex = /^\s*include\s+"([^"\n]+)"/gm;
  let match;

  while ((match = includeRegex.exec(text))) {
    includePaths.push(match[1]);
  }

  return includePaths;
}

function resolveDbdIncludePath(sourcePath, includePath) {
  const directPath = path.resolve(path.dirname(sourcePath), includePath);
  if (isExistingFile(directPath)) {
    return directPath;
  }

  const includeFileName = path.basename(includePath);
  return dbdFilesByName.get(includeFileName);
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

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function isExistingFile(filePath) {
  try {
    return fs.statSync(filePath).isFile();
  } catch (error) {
    return false;
  }
}
