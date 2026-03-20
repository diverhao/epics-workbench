function formatDatabaseText(text, options) {
  const normalizedText = String(text || "").replace(/\r\n/g, "\n");
  const hadTrailingNewline = normalizedText.endsWith("\n");
  const contentText = hadTrailingNewline
    ? normalizedText.slice(0, -1)
    : normalizedText;
  const lines = contentText ? contentText.split("\n") : [];
  const formattedLines = [];
  const indentUnit = getDatabaseIndentUnit(options);
  let indentLevel = 0;

  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) {
      formattedLines.push("");
      continue;
    }

    const effectiveIndentLevel = trimmedLine.startsWith("}")
      ? Math.max(indentLevel - 1, 0)
      : indentLevel;
    const formattedLine = formatDatabaseLine(trimmedLine);
    formattedLines.push(`${indentUnit.repeat(effectiveIndentLevel)}${formattedLine}`);
    indentLevel = Math.max(
      0,
      indentLevel + getBraceDeltaOutsideStrings(trimmedLine),
    );
  }

  let formattedText = formattedLines.join("\n");
  if (hadTrailingNewline) {
    formattedText += "\n";
  }

  return formattedText;
}

function formatSubstitutionText(text, options) {
  const normalizedText = String(text || "").replace(/\r\n/g, "\n");
  const hadTrailingNewline = normalizedText.endsWith("\n");
  const contentText = hadTrailingNewline
    ? normalizedText.slice(0, -1)
    : normalizedText;
  const lines = contentText ? contentText.split("\n") : [];
  const formattedLines = [];
  const indentUnit = getDatabaseIndentUnit(options);
  const state = {
    assignmentOrder: undefined,
  };
  let indentLevel = 0;

  for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
    const line = lines[lineIndex];
    const trimmedLine = line.trim();
    if (!trimmedLine) {
      formattedLines.push("");
      continue;
    }

    const { code, comment } = splitSubstitutionLineComment(trimmedLine);
    const trimmedCode = code.trim();
    const trimmedComment = comment ? comment.trim() : "";

    const startsWithClosingBrace =
      trimmedCode.length > 0 && trimmedCode.startsWith("}");
    const effectiveIndentLevel = startsWithClosingBrace
      ? Math.max(indentLevel - 1, 0)
      : indentLevel;
    const indentation = indentUnit.repeat(effectiveIndentLevel);

    if (!trimmedCode) {
      formattedLines.push(`${indentation}${trimmedComment}`);
      continue;
    }

    if (isSubstitutionBlockLine(trimmedCode)) {
      state.assignmentOrder = undefined;
    }

    const patternBlock = formatAlignedSubstitutionPatternBlock(
      lines,
      lineIndex,
      indentUnit,
      effectiveIndentLevel,
    );
    if (patternBlock) {
      formattedLines.push(...patternBlock.lines);
      lineIndex = patternBlock.nextLineIndex;
      continue;
    }

    const formattedCode = formatSubstitutionLine(trimmedCode, state);
    formattedLines.push(
      `${indentation}${formattedCode}${trimmedComment ? ` ${trimmedComment}` : ""}`,
    );
    indentLevel = Math.max(
      0,
      indentLevel + getBraceDeltaOutsideStrings(trimmedCode),
    );
    if (indentLevel === 0) {
      state.assignmentOrder = undefined;
    }
  }

  let formattedText = formattedLines.join("\n");
  if (hadTrailingNewline) {
    formattedText += "\n";
  }

  return formattedText;
}

function formatMonitorText(text) {
  const normalizedText = String(text || "").replace(/\r\n/g, "\n");
  const hadTrailingNewline = normalizedText.endsWith("\n");
  const contentText = hadTrailingNewline
    ? normalizedText.slice(0, -1)
    : normalizedText;
  const lines = contentText ? contentText.split("\n") : [];
  const formattedLines = [];

  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) {
      formattedLines.push("");
      continue;
    }

    if (trimmedLine.startsWith("#")) {
      const commentText = trimmedLine.slice(1).trimStart();
      formattedLines.push(commentText ? `# ${commentText}` : "#");
      continue;
    }

    const macroMatch = trimmedLine.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\S+)\s*$/);
    if (macroMatch) {
      formattedLines.push(`${macroMatch[1]} = ${macroMatch[2]}`);
      continue;
    }

    if (!/\s/.test(trimmedLine) && !trimmedLine.includes("=")) {
      formattedLines.push(trimmedLine);
      continue;
    }

    formattedLines.push(trimmedLine);
  }

  let formattedText = formattedLines.join("\n");
  if (hadTrailingNewline) {
    formattedText += "\n";
  }

  return formattedText;
}

function formatAlignedSubstitutionPatternBlock(
  lines,
  startIndex,
  indentUnit,
  effectiveIndentLevel,
) {
  const trimmedHeaderLine = lines[startIndex].trim();
  const { code: headerCode, comment: headerComment } =
    splitSubstitutionLineComment(trimmedHeaderLine);
  const headerValues = getSubstitutionPatternHeaderValues(headerCode.trim());
  if (!headerValues) {
    return undefined;
  }

  const rowEntries = [];
  let lineIndex = startIndex + 1;
  while (lineIndex < lines.length) {
    const trimmedLine = lines[lineIndex].trim();
    if (!trimmedLine) {
      break;
    }

    const { code, comment } = splitSubstitutionLineComment(trimmedLine);
    const trimmedCode = code.trim();
    if (!trimmedCode) {
      break;
    }

    const rowValues = getSubstitutionPatternRowValues(trimmedCode);
    if (!rowValues) {
      break;
    }

    rowEntries.push({
      values: rowValues,
      comment: comment ? comment.trim() : "",
    });
    lineIndex += 1;
  }

  if (rowEntries.length === 0) {
    return undefined;
  }

  const normalizedHeaderValues = headerValues.map((value) => String(value).trim());
  const normalizedRowValues = rowEntries.map((rowEntry) =>
    rowEntry.values.map((value) => formatSubstitutionScalarValue(value)),
  );
  const columnCount = Math.max(
    normalizedHeaderValues.length,
    ...normalizedRowValues.map((values) => values.length),
  );
  const columnWidths = Array.from({ length: columnCount }, (_, columnIndex) =>
    Math.max(
      normalizedHeaderValues[columnIndex]?.length || 0,
      ...normalizedRowValues.map((values) => values[columnIndex]?.length || 0),
    ),
  );
  const indentation = indentUnit.repeat(effectiveIndentLevel);
  const patternRowIndentation = `${indentation}${" ".repeat("pattern ".length)}`;
  const formattedLines = [
    `${indentation}pattern { ${formatAlignedSubstitutionPatternCells(
      normalizedHeaderValues,
      columnWidths,
    )} }${headerComment ? ` ${headerComment.trim()}` : ""}`,
  ];

  for (let rowIndex = 0; rowIndex < rowEntries.length; rowIndex += 1) {
    formattedLines.push(
      `${patternRowIndentation}{ ${formatAlignedSubstitutionPatternCells(
        normalizedRowValues[rowIndex],
        columnWidths,
      )} }${rowEntries[rowIndex].comment ? ` ${rowEntries[rowIndex].comment}` : ""}`,
    );
  }

  return {
    lines: formattedLines,
    nextLineIndex: lineIndex - 1,
  };
}

function getSubstitutionPatternHeaderValues(trimmedLine) {
  const match = trimmedLine.match(/^pattern\s*\{(.*)\}\s*$/);
  if (!match) {
    return undefined;
  }

  return splitSubstitutionCommaSeparatedItems(match[1]);
}

function getSubstitutionPatternRowValues(trimmedLine) {
  const match = trimmedLine.match(/^\{(.*)\}\s*$/);
  if (!match || parseSubstitutionAssignmentEntries(match[1])) {
    return undefined;
  }

  return splitSubstitutionCommaSeparatedItems(match[1]);
}

function formatAlignedSubstitutionPatternCells(values, columnWidths) {
  return values
    .map((value, columnIndex) => {
      const text = String(value ?? "").trim();
      if (columnIndex >= values.length - 1) {
        return text;
      }

      const paddingWidth = Math.max(
        0,
        (columnWidths[columnIndex] || 0) - text.length,
      );
      return `${text}, ${" ".repeat(paddingWidth)}`;
    })
    .join("");
}

function getDatabaseIndentUnit(options) {
  const tabSize = Number(options?.tabSize) > 0 ? Number(options.tabSize) : 4;
  if (options?.insertSpaces === false) {
    return "\t";
  }

  return " ".repeat(tabSize);
}

function formatDatabaseLine(trimmedLine) {
  if (trimmedLine.startsWith("#")) {
    return trimmedLine;
  }

  const normalizedRecordLine = normalizeDatabaseRecordLine(trimmedLine);
  if (normalizedRecordLine) {
    return normalizedRecordLine;
  }

  const normalizedFieldLine = normalizeDatabaseFieldLine(trimmedLine);
  if (normalizedFieldLine) {
    return normalizedFieldLine;
  }

  const normalizedClosingBraceLine = normalizeClosingBraceLine(trimmedLine);
  if (normalizedClosingBraceLine) {
    return normalizedClosingBraceLine;
  }

  return trimmedLine;
}

function formatSubstitutionLine(trimmedLine, state) {
  if (trimmedLine.startsWith("#")) {
    return trimmedLine;
  }

  const normalizedBlockLine = normalizeSubstitutionBlockLine(trimmedLine);
  if (normalizedBlockLine) {
    return normalizedBlockLine;
  }

  const normalizedPatternLine = normalizeSubstitutionPatternLine(trimmedLine);
  if (normalizedPatternLine) {
    return normalizedPatternLine;
  }

  const normalizedRowLine = normalizeSubstitutionRowLine(trimmedLine, state);
  if (normalizedRowLine) {
    return normalizedRowLine;
  }

  const normalizedClosingBraceLine = normalizeClosingBraceLine(trimmedLine);
  if (normalizedClosingBraceLine) {
    return normalizedClosingBraceLine;
  }

  return trimmedLine;
}

function isSubstitutionBlockLine(trimmedLine) {
  return /^(?:file\s+(?:"(?:[^"\\]|\\.)*"|[^\s{]+)|global)\s*\{\s*$/.test(trimmedLine) ||
    /^global\s*\{\s*$/.test(trimmedLine);
}

function normalizeDatabaseRecordLine(trimmedLine) {
  const match = trimmedLine.match(
    /^record\(\s*([A-Za-z0-9_]+)\s*,\s*("(?:(?:[^"\\]|\\.)*)")\s*\)\s*(\{)?\s*(#.*)?$/,
  );
  if (!match) {
    return undefined;
  }

  return `record(${match[1]}, ${match[2]})${match[3] ? " {" : ""}${match[4] ? ` ${match[4].trim()}` : ""}`;
}

function normalizeSubstitutionBlockLine(trimmedLine) {
  const fileMatch = trimmedLine.match(
    /^file\s+("(?:(?:[^"\\]|\\.)*)"|[^\s{]+)\s*\{\s*$/,
  );
  if (fileMatch) {
    return `file ${fileMatch[1]} {`;
  }

  if (/^global\s*\{\s*$/.test(trimmedLine)) {
    return "global {";
  }

  return undefined;
}

function normalizeDatabaseFieldLine(trimmedLine) {
  const match = trimmedLine.match(
    /^field\(\s*((?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+))\s*,\s*(.+?)\s*\)\s*(#.*)?$/,
  );
  if (!match) {
    return undefined;
  }

  return `field(${match[1]}, ${match[2].trim()})${match[3] ? ` ${match[3].trim()}` : ""}`;
}

function normalizeSubstitutionPatternLine(trimmedLine) {
  const match = trimmedLine.match(/^pattern\s*\{(.*)\}\s*$/);
  if (!match) {
    return undefined;
  }

  const values = splitSubstitutionCommaSeparatedItems(match[1]);
  return `pattern { ${values
    .map((value) => formatSubstitutionScalarValue(value))
    .join(", ")} }`;
}

function normalizeSubstitutionRowLine(trimmedLine, state) {
  const match = trimmedLine.match(/^\{(.*)\}\s*$/);
  if (!match) {
    return undefined;
  }

  const innerText = match[1];
  if (!innerText.trim()) {
    return "{ }";
  }

  const assignmentEntries = parseSubstitutionAssignmentEntries(innerText);
  if (assignmentEntries) {
    return `{ ${orderSubstitutionAssignmentEntries(assignmentEntries, state)
      .map(([name, value]) => `${name}=${formatSubstitutionAssignmentValue(value)}`)
      .join(", ")} }`;
  }

  const values = splitSubstitutionCommaSeparatedItems(innerText);
  return `{ ${values.map((value) => formatSubstitutionScalarValue(value)).join(", ")} }`;
}

function normalizeClosingBraceLine(trimmedLine) {
  const match = trimmedLine.match(/^\}(.*)$/);
  if (!match) {
    return undefined;
  }

  return `}${match[1] ? ` ${match[1].trim()}` : ""}`;
}

function getBraceDeltaOutsideStrings(text) {
  let delta = 0;
  let inString = false;
  let escaped = false;

  for (const character of text) {
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

    if (character === "#") {
      break;
    }

    if (character === "{") {
      delta += 1;
    } else if (character === "}") {
      delta -= 1;
    }
  }

  return delta;
}

function splitSubstitutionLineComment(text) {
  let inString = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
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

    if (character === "#") {
      return {
        code: text.slice(0, index).trimEnd(),
        comment: text.slice(index),
      };
    }
  }

  return {
    code: text,
    comment: "",
  };
}

function splitSubstitutionCommaSeparatedItems(text) {
  const items = [];
  let segmentStart = 0;
  let inString = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
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

    if (character === ",") {
      items.push(text.slice(segmentStart, index).trim());
      segmentStart = index + 1;
    }
  }

  items.push(text.slice(segmentStart).trim());
  return items;
}

function parseSubstitutionAssignmentEntries(text) {
  const items = splitSubstitutionCommaSeparatedItems(text);
  const entries = [];

  for (const item of items) {
    const match = item.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!match) {
      return undefined;
    }

    entries.push([match[1], match[2].trim()]);
  }

  return entries;
}

function orderSubstitutionAssignmentEntries(entries, state) {
  if (!state) {
    return entries;
  }

  if (!state.assignmentOrder) {
    state.assignmentOrder = entries.map(([name]) => name);
    return entries;
  }

  const remainingEntries = entries.map((entry, index) => ({
    entry,
    index,
  }));
  const orderedEntries = [];

  for (const name of state.assignmentOrder) {
    for (let index = 0; index < remainingEntries.length; index += 1) {
      if (remainingEntries[index].entry[0] !== name) {
        continue;
      }

      orderedEntries.push(remainingEntries[index].entry);
      remainingEntries.splice(index, 1);
      index -= 1;
    }
  }

  for (const { entry } of remainingEntries) {
    if (!state.assignmentOrder.includes(entry[0])) {
      state.assignmentOrder.push(entry[0]);
    }
    orderedEntries.push(entry);
  }

  return orderedEntries;
}

function formatSubstitutionScalarValue(rawValue) {
  const trimmedValue = String(rawValue ?? "").trim();
  if (!trimmedValue) {
    return "\"\"";
  }

  const quotedMatch = trimmedValue.match(/^"((?:[^"\\]|\\.)*)"$/);
  if (quotedMatch) {
    return `"${quotedMatch[1]}"`;
  }

  return /[\s,{}#"]/.test(trimmedValue)
    ? `"${escapeDoubleQuotedString(trimmedValue)}"`
    : trimmedValue;
}

function formatSubstitutionAssignmentValue(rawValue) {
  const trimmedValue = String(rawValue ?? "").trim();
  return trimmedValue ? formatSubstitutionScalarValue(trimmedValue) : "";
}

function escapeDoubleQuotedString(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

module.exports = {
  formatDatabaseText,
  formatMonitorText,
  formatSubstitutionText,
  splitSubstitutionCommaSeparatedItems,
};
