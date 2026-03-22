const path = require("path");
const zlib = require("zlib");

function createDatabaseExcelImportTools() {
  function importDatabaseWorkbookBuffer(workbookBuffer, metadata = {}) {
    const workbook = parseWorkbookBuffer(workbookBuffer);
    const importedAt = normalizeDate(metadata.importedAt) || new Date();
    return workbook.sheets
      .filter(isEpicsDatabaseSheet)
      .map((sheet, index) => ({
        sheetName: sheet.name,
        suggestedFileName: buildSuggestedFileName(
          metadata.sourceFileName || "import",
          sheet.name,
          index,
        ),
        text: buildImportedDatabaseText(sheet, {
          sourceFileName: metadata.sourceFileName || "unknown.xlsx",
          sourceSheetName: sheet.name,
          sourceCreatedAt: normalizeDate(metadata.sourceCreatedAt),
          sourceModifiedAt: normalizeDate(metadata.sourceModifiedAt),
          importedAt,
        }),
      }));
  }

  function parseWorkbookBuffer(workbookBuffer) {
    const zipEntries = readZipEntries(workbookBuffer);
    const workbookXml = requireEntryText(zipEntries, "xl/workbook.xml");
    const workbookRelsXml = requireEntryText(zipEntries, "xl/_rels/workbook.xml.rels");
    const relationships = parseWorkbookRelationships(workbookRelsXml);
    const sharedStrings = parseSharedStrings(zipEntries.get("xl/sharedStrings.xml"));

    const sheets = [];
    for (const sheetDefinition of parseWorkbookSheets(workbookXml)) {
      const target = relationships.get(sheetDefinition.relationshipId);
      if (!target) {
        continue;
      }

      const worksheetPath = resolveWorkbookTargetPath(target);
      const worksheetXml = zipEntries.get(worksheetPath);
      if (!worksheetXml) {
        continue;
      }

      sheets.push({
        name: sheetDefinition.name,
        rows: parseWorksheet(worksheetXml.toString("utf8"), sharedStrings),
      });
    }

    return { sheets };
  }

  function isEpicsDatabaseSheet(sheet) {
    const headerRow = sheet.rows.get(1);
    if (!headerRow) {
      return false;
    }

    return (
      (headerRow.get(0) || "").trim() === "Record" &&
      (headerRow.get(1) || "").trim() === "Type"
    );
  }

  function buildImportedDatabaseText(sheet, metadata) {
    const lines = [
      "# Imported from Excel by EPICS Workbench",
      `# Source file: ${metadata.sourceFileName}`,
      `# Source sheet: ${metadata.sourceSheetName}`,
      `# Source created time: ${formatTimestamp(metadata.sourceCreatedAt)}`,
      `# Source modified time: ${formatTimestamp(metadata.sourceModifiedAt)}`,
      `# Imported at: ${formatTimestamp(metadata.importedAt)}`,
      "",
    ];

    const headerRow = sheet.rows.get(1) || new Map();
    const headers = [];
    let maxColumnIndex = -1;
    for (const columnIndex of headerRow.keys()) {
      if (columnIndex > maxColumnIndex) {
        maxColumnIndex = columnIndex;
      }
    }
    for (let columnIndex = 0; columnIndex <= maxColumnIndex; columnIndex += 1) {
      headers.push((headerRow.get(columnIndex) || "").trim());
    }

    const sortedRowNumbers = [...sheet.rows.keys()].filter((rowNumber) => rowNumber > 1).sort((a, b) => a - b);
    for (const rowNumber of sortedRowNumbers) {
      const row = sheet.rows.get(rowNumber) || new Map();
      const recordName = String(row.get(0) || "").trim();
      const recordType = String(row.get(1) || "").trim();
      if (!recordName || !recordType) {
        continue;
      }

      lines.push(`record(${recordType}, "${escapeDatabaseString(recordName)}") {`);
      for (let columnIndex = 2; columnIndex < headers.length; columnIndex += 1) {
        const fieldName = headers[columnIndex];
        const rawValue = row.get(columnIndex);
        if (!fieldName || rawValue == null || !String(rawValue).trim()) {
          continue;
        }
        lines.push(`    field(${fieldName}, "${escapeDatabaseString(String(rawValue))}")`);
      }
      lines.push("}");
      lines.push("");
    }

    if (lines[lines.length - 1] === "") {
      return `${lines.join("\n")}\n`;
    }
    return `${lines.join("\n")}\n`;
  }

  function readZipEntries(buffer) {
    const eocdOffset = findEndOfCentralDirectoryOffset(buffer);
    if (eocdOffset < 0) {
      throw new Error("The selected file is not a valid XLSX workbook.");
    }

    const centralDirectorySize = buffer.readUInt32LE(eocdOffset + 12);
    const centralDirectoryOffset = buffer.readUInt32LE(eocdOffset + 16);
    const zipEntries = new Map();
    let offset = centralDirectoryOffset;
    const endOffset = centralDirectoryOffset + centralDirectorySize;

    while (offset < endOffset) {
      if (buffer.readUInt32LE(offset) !== 0x02014b50) {
        throw new Error("The selected file has an invalid XLSX central directory.");
      }

      const compressionMethod = buffer.readUInt16LE(offset + 10);
      const compressedSize = buffer.readUInt32LE(offset + 20);
      const fileNameLength = buffer.readUInt16LE(offset + 28);
      const extraFieldLength = buffer.readUInt16LE(offset + 30);
      const fileCommentLength = buffer.readUInt16LE(offset + 32);
      const localHeaderOffset = buffer.readUInt32LE(offset + 42);
      const fileName = buffer.toString(
        "utf8",
        offset + 46,
        offset + 46 + fileNameLength,
      );

      const localFileNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
      const localExtraFieldLength = buffer.readUInt16LE(localHeaderOffset + 28);
      const dataOffset = localHeaderOffset + 30 + localFileNameLength + localExtraFieldLength;
      const compressedData = buffer.subarray(dataOffset, dataOffset + compressedSize);
      let data;
      if (compressionMethod === 0) {
        data = Buffer.from(compressedData);
      } else if (compressionMethod === 8) {
        data = zlib.inflateRawSync(compressedData);
      } else {
        throw new Error(`Unsupported XLSX compression method ${compressionMethod}.`);
      }

      zipEntries.set(fileName, data);
      offset += 46 + fileNameLength + extraFieldLength + fileCommentLength;
    }

    return zipEntries;
  }

  function findEndOfCentralDirectoryOffset(buffer) {
    const minOffset = Math.max(0, buffer.length - 0xffff - 22);
    for (let offset = buffer.length - 22; offset >= minOffset; offset -= 1) {
      if (buffer.readUInt32LE(offset) === 0x06054b50) {
        return offset;
      }
    }
    return -1;
  }

  function requireEntryText(zipEntries, entryName) {
    const entry = zipEntries.get(entryName);
    if (!entry) {
      throw new Error(`The selected workbook is missing ${entryName}.`);
    }
    return entry.toString("utf8");
  }

  function parseWorkbookRelationships(xml) {
    const relationships = new Map();
    const relationshipPattern = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
    let match;
    while ((match = relationshipPattern.exec(xml))) {
      relationships.set(match[1], match[2]);
    }
    return relationships;
  }

  function parseWorkbookSheets(xml) {
    const sheets = [];
    const sheetPattern = /<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]+)"[^>]*\/>/g;
    let match;
    while ((match = sheetPattern.exec(xml))) {
      sheets.push({
        name: decodeXmlText(match[1]),
        relationshipId: match[2],
      });
    }
    return sheets;
  }

  function resolveWorkbookTargetPath(target) {
    if (!target) {
      return "";
    }
    if (target.startsWith("/")) {
      return target.replace(/^\/+/, "");
    }
    return path.posix.normalize(path.posix.join("xl", target));
  }

  function parseSharedStrings(buffer) {
    if (!buffer) {
      return [];
    }

    const xml = buffer.toString("utf8");
    const strings = [];
    const stringPattern = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let match;
    while ((match = stringPattern.exec(xml))) {
      strings.push(extractInlineString(match[1]));
    }
    return strings;
  }

  function parseWorksheet(xml, sharedStrings) {
    const rows = new Map();
    const cellPattern = /<c\b([^>]*?)(?:>([\s\S]*?)<\/c>|\/>)/g;
    let match;
    while ((match = cellPattern.exec(xml))) {
      const attributes = match[1];
      const body = match[2] || "";
      const referenceMatch = /\br="([A-Z]+)(\d+)"/.exec(attributes);
      if (!referenceMatch) {
        continue;
      }

      const columnIndex = columnNameToIndex(referenceMatch[1]);
      const rowNumber = Number(referenceMatch[2]);
      const cellType = (/\bt="([^"]+)"/.exec(attributes) || [])[1] || "";
      const cellValue = parseWorksheetCellValue(cellType, body, sharedStrings);
      const row = rows.get(rowNumber) || new Map();
      row.set(columnIndex, cellValue);
      rows.set(rowNumber, row);
    }
    return rows;
  }

  function parseWorksheetCellValue(cellType, body, sharedStrings) {
    if (cellType === "inlineStr") {
      return extractInlineString(body);
    }

    const valueMatch = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body);
    const rawValue = valueMatch ? decodeXmlText(valueMatch[1]) : "";

    if (cellType === "s") {
      const sharedStringIndex = Number(rawValue);
      return Number.isInteger(sharedStringIndex)
        ? sharedStrings[sharedStringIndex] || ""
        : "";
    }

    if (cellType === "str" || cellType === "b" || cellType === "e") {
      return rawValue;
    }

    return rawValue;
  }

  function extractInlineString(xml) {
    const parts = [];
    const textPattern = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let match;
    while ((match = textPattern.exec(xml))) {
      parts.push(decodeXmlText(match[1]));
    }
    return parts.join("");
  }

  function columnNameToIndex(columnName) {
    let value = 0;
    for (const character of columnName) {
      value = value * 26 + (character.charCodeAt(0) - 64);
    }
    return value - 1;
  }

  function buildSuggestedFileName(sourceFileName, sheetName, index) {
    const baseName = path.basename(sourceFileName, path.extname(sourceFileName)) || "import";
    const safeSheetName = sanitizeFileNameFragment(sheetName || `sheet-${index + 1}`);
    return `${sanitizeFileNameFragment(baseName)}-${safeSheetName}.db`;
  }

  function sanitizeFileNameFragment(value) {
    return String(value || "sheet")
      .replace(/[<>:"/\\|?*\u0000-\u001f]+/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "") || "sheet";
  }

  function escapeDatabaseString(value) {
    return String(value)
      .replace(/\\/g, "\\\\")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .replace(/\n/g, "\\n")
      .replace(/"/g, '\\"');
  }

  function decodeXmlText(value) {
    return String(value)
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  function formatTimestamp(value) {
    return value instanceof Date && !Number.isNaN(value.getTime())
      ? value.toISOString()
      : "unavailable";
  }

  function normalizeDate(value) {
    if (!value) {
      return undefined;
    }
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? undefined : value;
    }
    const nextValue = new Date(value);
    return Number.isNaN(nextValue.getTime()) ? undefined : nextValue;
  }

  return {
    importDatabaseWorkbookBuffer,
  };
}

module.exports = {
  createDatabaseExcelImportTools,
};
