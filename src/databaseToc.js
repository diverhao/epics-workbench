const DATABASE_TOC_BEGIN_MARKER = "# EPICS TOC BEGIN";
const DATABASE_TOC_END_MARKER = "# EPICS TOC END";
const DATABASE_TOC_VALUE_COLUMN_INDEX = 1;
const DATABASE_TOC_VALUE_COLUMN_WIDTH = 18;

function createDatabaseTocTools({
  openRecordLocationCommand,
  getRecordTemplateStaticData,
  getDefaultFieldValue,
  extractRecordDeclarations,
  extractFieldDeclarationsInRecord,
  maskDatabaseComments,
  extractMacroNames,
  compareLabels,
  escapeRegExp,
}) {
  function upsertDatabaseTocText(text, eol) {
    const contentWithoutToc = removeDatabaseTocBlock(text).replace(
      /^(?:[ \t]*\r?\n)+/,
      "",
    );
    const tocBlock = buildDatabaseTocBlock(text, eol);
    if (!contentWithoutToc) {
      return `${tocBlock}${eol}`;
    }

    return `${tocBlock}${eol}${eol}${contentWithoutToc}`;
  }

  function buildDatabaseTocBlock(text, eol) {
    const extraFieldNames = getDatabaseTocExtraFieldNames(text);
    const macroNames = getDatabaseTocMacroNames(text);
    const macroAssignments = extractDatabaseTocMacroAssignments(text);
    const headerSuffix =
      extraFieldNames.length > 0 ? ` ${extraFieldNames.join(" ")}` : "";
    const lines = [DATABASE_TOC_BEGIN_MARKER];
    if (macroNames.length > 0) {
      lines.push("# Macros:");
      for (const macroName of macroNames) {
        const assignment = macroAssignments.get(macroName);
        lines.push(
          `#  - ${macroName}${assignment?.hasAssignment ? ` = ${assignment.value}` : ""}`,
        );
      }
    }
    lines.push(`# Table of Contents${headerSuffix}`);
    const declarations = extractRecordDeclarations(text);
    const headerRow = ["Record", "Value", "Type", ...extraFieldNames];
    const rows = declarations.map((declaration) =>
      buildDatabaseTocRowValues(text, declaration, extraFieldNames),
    );
    const columnWidths = getDatabaseTocColumnWidths([headerRow, ...rows]);

    lines.push(`# ${formatDatabaseTocMarkdownRow(headerRow, columnWidths)}`);
    lines.push(`# ${formatDatabaseTocMarkdownSeparatorRow(columnWidths)}`);

    for (const row of rows) {
      lines.push(`# ${formatDatabaseTocMarkdownRow(row, columnWidths)}`);
    }

    lines.push(DATABASE_TOC_END_MARKER);
    return lines.join(eol);
  }

  function getDatabaseTocMacroNames(text) {
    return extractMacroNames(maskDatabaseComments(text)).sort(compareLabels);
  }

  function extractDatabaseTocMacroAssignments(text) {
    const range = findDatabaseTocBlockRange(text);
    if (!range) {
      return new Map();
    }

    const assignments = new Map();
    const tocText = text.slice(range.start, range.end);
    let inMacroSection = false;

    for (const lineText of tocText.split(/\r?\n/)) {
      if (/^#\s*Macros:\s*$/.test(lineText)) {
        inMacroSection = true;
        continue;
      }

      if (!inMacroSection) {
        continue;
      }

      if (/^#\s*Table of Contents(?:\s+.*)?\s*$/.test(lineText)) {
        break;
      }

      const match = lineText.match(
        /^#\s*-\s*([A-Za-z_][A-Za-z0-9_]*)(?:\s*=\s*(.*))?\s*$/,
      );
      if (!match) {
        continue;
      }

      assignments.set(match[1], {
        hasAssignment: match[2] !== undefined,
        value: match[2] || "",
      });
    }

    return assignments;
  }

  function getDatabaseTocExtraFieldNames(text) {
    const range = findDatabaseTocBlockRange(text);
    const searchableText = range ? text.slice(range.start, range.end) : text;
    const headerMatch = searchableText.match(
      /^#\s*Table of Contents(?:[ \t]+(.*?))?\s*$/m,
    );
    if (!headerMatch || !headerMatch[1]) {
      return [];
    }

    return [...new Set(
      headerMatch[1]
        .split(/\s+/)
        .map((fieldName) => fieldName.trim().toUpperCase())
        .filter((fieldName) => /^[A-Z0-9_]+$/.test(fieldName)),
    )];
  }

  function buildDatabaseTocRowValues(text, declaration, extraFieldNames) {
    const row = [declaration.name, undefined, declaration.recordType];

    const fieldDeclarations = extractFieldDeclarationsInRecord(text, declaration);
    for (const fieldName of extraFieldNames) {
      const fieldDeclaration = fieldDeclarations.find(
        (candidate) => candidate.fieldName === fieldName,
      );
      row.push(
        fieldDeclaration
          ? fieldDeclaration.value
          : getDatabaseTocMissingFieldValue(declaration.recordType, fieldName),
      );
    }

    return row;
  }

  function getDatabaseTocMissingFieldValue(recordType, fieldName) {
    const recordTemplateStaticData = getRecordTemplateStaticData();
    if (
      !recordTemplateStaticData?.fieldTypesByRecordType
        .get(recordType)
        ?.has(fieldName)
    ) {
      return "NA";
    }

    return getDefaultFieldValue(recordTemplateStaticData, recordType, fieldName);
  }

  function formatDatabaseTocValue(value) {
    if (value === undefined || value === null) {
      return "";
    }

    return value === "" ? '""' : String(value);
  }

  function getDatabaseTocColumnWidths(rows) {
    const widths = [];
    for (const row of rows) {
      row.forEach((value, index) => {
        widths[index] = Math.max(
          widths[index] || 0,
          formatDatabaseTocValue(value).length,
        );
      });
    }

    widths[DATABASE_TOC_VALUE_COLUMN_INDEX] = DATABASE_TOC_VALUE_COLUMN_WIDTH;
    return widths;
  }

  function formatDatabaseTocMarkdownRow(row, columnWidths) {
    return `| ${row
      .map((value, index) => formatDatabaseTocValue(value).padEnd(columnWidths[index]))
      .join(" | ")} |`;
  }

  function formatDatabaseTocMarkdownSeparatorRow(columnWidths) {
    return `| ${columnWidths.map((width) => "-".repeat(Math.max(3, width))).join(" | ")} |`;
  }

  function removeDatabaseTocBlock(text) {
    const range = findDatabaseTocBlockRange(text);
    if (!range) {
      return text;
    }

    return `${text.slice(0, range.start)}${text.slice(range.end)}`;
  }

  function findDatabaseTocBlockRange(text) {
    const start = text.indexOf(DATABASE_TOC_BEGIN_MARKER);
    if (start < 0) {
      return undefined;
    }

    const endMarkerStart = text.indexOf(DATABASE_TOC_END_MARKER, start);
    if (endMarkerStart < 0) {
      return undefined;
    }

    let end = endMarkerStart + DATABASE_TOC_END_MARKER.length;
    if (text.slice(end, end + 2) === "\r\n") {
      end += 2;
    } else if (text[end] === "\n") {
      end += 1;
    }

    while (text.slice(end, end + 2) === "\r\n") {
      end += 2;
    }
    while (text[end] === "\n") {
      end += 1;
    }

    return { start, end };
  }

  function extractDatabaseTocEntries(text) {
    const range = findDatabaseTocBlockRange(text);
    if (!range) {
      return [];
    }

    const tocText = text.slice(range.start, range.end);
    const recordDeclarations = extractRecordDeclarations(text);
    const entries = [];

    for (const lineMatch of tocText.matchAll(/^.*$/gm)) {
      const lineText = lineMatch[0];
      const lineOffset = range.start + lineMatch.index;
      const legacyEntry = parseLegacyDatabaseTocEntry(lineText, lineOffset);
      if (legacyEntry) {
        entries.push(legacyEntry);
        continue;
      }

      const markdownEntry = parseMarkdownDatabaseTocEntry(
        lineText,
        lineOffset,
        recordDeclarations,
      );
      if (markdownEntry) {
        entries.push(markdownEntry);
        continue;
      }

      const plainEntry = parsePlainDatabaseTocEntry(
        lineText,
        lineOffset,
        recordDeclarations,
      );
      if (plainEntry) {
        entries.push(plainEntry);
      }
    }

    return entries;
  }

  function parseLegacyDatabaseTocEntry(lineText, lineOffset) {
    const match = lineText.match(
      /^#\s*-\s*\[([A-Za-z0-9_]+)\s+([^\]]+)\]\(#([^)]+)\)\s*$/,
    );
    if (!match) {
      return undefined;
    }

    const linkStart = lineOffset + lineText.indexOf("[");
    const linkEnd = lineOffset + lineText.lastIndexOf(")") + 1;
    const typeStart = linkStart + 1;
    const typeEnd = typeStart + match[1].length;
    const nameStart = linkStart + 1 + match[1].length + 1;
    const nameEnd = nameStart + match[2].length;
    return {
      recordType: match[1],
      recordName: match[2],
      anchor: match[3],
      linkStart,
      linkEnd,
      hoverStart: typeStart,
      hoverEnd: typeEnd,
      nameStart,
      nameEnd,
    };
  }

  function parseMarkdownDatabaseTocEntry(lineText, lineOffset, declarations) {
    const cellEntries = extractCommentTableCells(lineText, lineOffset);
    if (cellEntries.length < 2) {
      return undefined;
    }

    const cellValues = cellEntries.map((cell) => cell.value);
    if (
      cellValues.every((value) => /^:?-{3,}:?$/.test(value)) ||
      /^record$/i.test(cellValues[0])
    ) {
      return undefined;
    }

    for (const declaration of declarations) {
      if (cellValues[0] !== declaration.name) {
        continue;
      }

      const firstCell = cellEntries[0];
      const hasValueColumn =
        cellEntries.length >= 3 && cellValues[2] === declaration.recordType;
      const typeCell = hasValueColumn ? cellEntries[2] : cellEntries[1];
      if (typeCell?.value !== declaration.recordType) {
        continue;
      }

      const valueCell = hasValueColumn ? cellEntries[1] : undefined;
      const lastCell = cellEntries[cellEntries.length - 1];
      return {
        recordType: declaration.recordType,
        recordName: declaration.name,
        linkStart: firstCell.start,
        linkEnd: lastCell.end,
        hoverStart: typeCell.start,
        hoverEnd: typeCell.end,
        nameStart: firstCell.start,
        nameEnd: firstCell.end,
        valueStart: valueCell?.displayStart,
        valueEnd: valueCell?.displayEnd,
      };
    }

    return undefined;
  }

  function parsePlainDatabaseTocEntry(lineText, lineOffset, declarations) {
    const prefixMatch = lineText.match(/^#\s+(.*?)\s*$/);
    if (!prefixMatch) {
      return undefined;
    }

    const content = prefixMatch[1];
    const contentStartInLine = lineText.indexOf(content);
    if (contentStartInLine < 0) {
      return undefined;
    }

    for (const declaration of declarations) {
      const newLayoutRegex = new RegExp(
        `^${escapeRegExp(declaration.name)}\\s+${escapeRegExp(declaration.recordType)}(?:\\s+.*)?$`,
      );
      const oldLayoutRegex = new RegExp(
        `^${escapeRegExp(declaration.recordType)}\\s+${escapeRegExp(declaration.name)}(?:\\s+.*)?$`,
      );
      const isNewLayout = newLayoutRegex.test(content);
      const isOldLayout = !isNewLayout && oldLayoutRegex.test(content);
      if (!isNewLayout && !isOldLayout) {
        continue;
      }

      const nameStartInContent = content.indexOf(declaration.name);
      const typeStartInContent = isNewLayout
        ? content.indexOf(
            declaration.recordType,
            nameStartInContent + declaration.name.length,
          )
        : content.indexOf(declaration.recordType);
      if (nameStartInContent < 0 || typeStartInContent < 0) {
        continue;
      }

      const contentStart = lineOffset + contentStartInLine;
      const nameStart = contentStart + nameStartInContent;
      const nameEnd = nameStart + declaration.name.length;
      const typeStart = contentStart + typeStartInContent;
      const typeEnd = typeStart + declaration.recordType.length;
      return {
        recordType: declaration.recordType,
        recordName: declaration.name,
        linkStart: contentStart,
        linkEnd: contentStart + content.length,
        hoverStart: typeStart,
        hoverEnd: typeEnd,
        nameStart,
        nameEnd,
      };
    }

    return undefined;
  }

  function extractCommentTableCells(lineText, lineOffset) {
    const prefixMatch = lineText.match(/^#\s*(\|.*)$/);
    if (!prefixMatch) {
      return [];
    }

    const tableText = prefixMatch[1];
    const tableOffset = lineOffset + lineText.indexOf(tableText);
    const cells = [];
    let cellStart = tableOffset + 1;

    for (let index = 1; index < tableText.length; index += 1) {
      if (tableText[index] !== "|") {
        continue;
      }

      const rawCell = tableText.slice(cellStart - tableOffset, index);
      const leadingWhitespaceLength = rawCell.match(/^\s*/)[0].length;
      const trailingWhitespaceLength = rawCell.match(/\s*$/)[0].length;
      const trimmedStart = cellStart + leadingWhitespaceLength;
      const trimmedEnd = tableOffset + index - trailingWhitespaceLength;
      const displayStart = cellStart + Math.min(leadingWhitespaceLength, 1);
      const displayEnd = tableOffset + index - Math.min(trailingWhitespaceLength, 1);
      const value = rawCell.trim();
      cells.push({
        value,
        start: trimmedStart,
        end: Math.max(trimmedStart, trimmedEnd),
        displayStart,
        displayEnd: Math.max(displayStart, displayEnd),
      });
      cellStart = tableOffset + index + 1;
    }

    return cells;
  }

  function findRecordDeclarationByTypeAndName(text, recordType, recordName) {
    return extractRecordDeclarations(text).find(
      (declaration) =>
        declaration.recordType === recordType && declaration.name === recordName,
    );
  }

  function buildOpenRecordCommandUri(absolutePath, line) {
    const commandArguments = encodeURIComponent(
      JSON.stringify([
        {
          absolutePath,
          line,
        },
      ]),
    );
    return `command:${openRecordLocationCommand}?${commandArguments}`;
  }

  return {
    buildOpenRecordCommandUri,
    extractDatabaseTocEntries,
    extractDatabaseTocMacroAssignments,
    findRecordDeclarationByTypeAndName,
    removeDatabaseTocBlock,
    upsertDatabaseTocText,
  };
}

module.exports = {
  createDatabaseTocTools,
};
