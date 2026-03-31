function createDatabaseExcelTools({
  extractRecordDeclarations,
  extractFieldDeclarationsInRecord,
}) {
  const COMMENT_ROW_KIND = "comment";
  const RECORD_ROW_KIND = "record";
  const {
    BACKGROUND_COLOR_OPTIONS,
    getBackgroundArgb,
    normalizeBackgroundToken,
  } = require("./spreadsheetCellBackgrounds");
  const FILLED_BACKGROUND_OPTIONS = BACKGROUND_COLOR_OPTIONS.filter((option) => option.token);

  function buildDatabaseWorkbookBuffer(text) {
    return buildWorkbookBufferFromSheetModels([
      buildSheetModelFromDatabaseText(text, "Database"),
    ]);
  }

  function buildSheetModelFromDatabaseText(text, name = "Database") {
    const entries = buildEntries(text);
    const headers = buildHeaders(entries);
    return {
      name,
      headers,
      rows: entries.map((entry) => {
        if (entry.kind === COMMENT_ROW_KIND) {
          return {
            kind: COMMENT_ROW_KIND,
            text: entry.text,
          };
        }

        return {
          kind: RECORD_ROW_KIND,
          values: [
            entry.recordName,
            entry.recordType,
            ...headers
              .slice(2)
              .map((fieldName) => entry.explicitFieldValues.get(fieldName) || ""),
          ],
        };
      }),
    };
  }

  function buildWorkbookBufferFromSheetModels(sheetModels) {
    const normalizedSheets = normalizeSheetModels(sheetModels);
    const parts = new Map([
      ["[Content_Types].xml", Buffer.from(buildContentTypesXml(normalizedSheets), "utf8")],
      ["_rels/.rels", Buffer.from(buildRootRelsXml(), "utf8")],
      ["xl/workbook.xml", Buffer.from(buildWorkbookXml(normalizedSheets), "utf8")],
      ["xl/_rels/workbook.xml.rels", Buffer.from(buildWorkbookRelsXml(normalizedSheets), "utf8")],
      ["xl/styles.xml", Buffer.from(buildStylesXml(), "utf8")],
    ]);
    normalizedSheets.forEach((sheet, index) => {
      parts.set(
        `xl/worksheets/sheet${index + 1}.xml`,
        Buffer.from(buildSheetXml(sheet), "utf8"),
      );
    });
    return buildZipBuffer(parts);
  }

  function normalizeSheetModels(sheetModels) {
    const usedNames = new Set();
    const fallbackSheets =
      Array.isArray(sheetModels) && sheetModels.length > 0
        ? sheetModels
        : [{ name: "Database", headers: ["Record", "Type"], rows: [] }];

    return fallbackSheets.map((sheet, index) => {
      const normalizedHeaders = normalizeHeaders(sheet?.headers);
      const normalizedName = getUniqueSheetName(
        normalizeSheetName(sheet?.name || `Sheet ${index + 1}`),
        usedNames,
      );
      return {
        name: normalizedName,
        headers: normalizedHeaders,
        headerBackgrounds: normalizeBackgrounds(sheet?.headerBackgrounds, normalizedHeaders.length),
        rows: normalizeRows(sheet?.rows, normalizedHeaders.length),
      };
    });
  }

  function normalizeHeaders(headers) {
    const normalizedHeaders = Array.isArray(headers)
      ? headers.map((header) => String(header || "").trim())
      : [];
    if (normalizedHeaders[0] !== "Record") {
      normalizedHeaders[0] = "Record";
    }
    if (normalizedHeaders[1] !== "Type") {
      normalizedHeaders[1] = "Type";
    }
    return normalizedHeaders.slice(0, 2).concat(
      normalizedHeaders
        .slice(2)
        .map((header, index) => header || `Field${index + 1}`),
    );
  }

  function normalizeRows(rows, columnCount) {
    if (!Array.isArray(rows)) {
      return [];
    }
    return rows.map((row) => {
      if (isCommentRow(row)) {
        return {
          kind: COMMENT_ROW_KIND,
          text: String(row.text || ""),
          background: normalizeBackgroundToken(row.background),
        };
      }

      const values = Array.isArray(row?.values)
        ? row.values
        : Array.isArray(row)
          ? row
          : [];
      const normalized = [];
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        normalized.push(String(values[columnIndex] || ""));
      }
      return {
        kind: RECORD_ROW_KIND,
        values: normalized,
        backgrounds: normalizeBackgrounds(row?.backgrounds, columnCount),
      };
    });
  }

  function normalizeBackgrounds(backgrounds, columnCount) {
    const values = Array.isArray(backgrounds) ? backgrounds : [];
    const normalized = [];
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      normalized.push(normalizeBackgroundToken(values[columnIndex]));
    }
    return normalized;
  }

  function getUniqueSheetName(baseName, usedNames) {
    let candidate = baseName || "Sheet";
    let suffix = 2;
    while (usedNames.has(candidate)) {
      const trimmedBase = baseName.slice(0, Math.max(1, 31 - String(suffix).length - 1));
      candidate = `${trimmedBase}_${suffix}`;
      suffix += 1;
    }
    usedNames.add(candidate);
    return candidate;
  }

  function normalizeSheetName(name) {
    const sanitized = String(name || "Sheet")
      .replace(/[\[\]:*?/\\]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const fallback = sanitized || "Sheet";
    return fallback.slice(0, 31);
  }

  function buildEntries(text) {
    const declarations = extractRecordDeclarations(text)
      .slice()
      .sort((left, right) => left.recordStart - right.recordStart);
    const entries = [];
    let cursor = 0;

    declarations.forEach((declaration) => {
      entries.push(
        ...collectCommentEntriesFromSegment(
          text.slice(cursor, declaration.recordStart),
          cursor,
        ),
      );

      const explicitFieldValues = new Map();
      for (const fieldDeclaration of extractFieldDeclarationsInRecord(text, declaration)) {
        explicitFieldValues.set(
          fieldDeclaration.fieldName,
          fieldDeclaration.value,
        );
      }

      entries.push(
        ...collectCommentEntriesFromRecord(
          text.slice(declaration.recordStart, declaration.recordEnd),
        ),
      );

      entries.push({
        kind: RECORD_ROW_KIND,
        recordName: declaration.name,
        recordType: declaration.recordType,
        explicitFieldValues,
      });
      cursor = Math.max(cursor, declaration.recordEnd);
    });

    entries.push(
      ...collectCommentEntriesFromSegment(text.slice(cursor), cursor),
    );
    return entries;
  }

  function collectCommentEntriesFromSegment(segment, segmentOffset = 0) {
    const entries = [];
    let lineOffset = Number(segmentOffset) || 0;

    for (const rawLineWithBreak of String(segment || "").match(/[^\n]*(?:\n|$)/g) || []) {
      if (!rawLineWithBreak) {
        continue;
      }

      const rawLine = rawLineWithBreak.endsWith("\n")
        ? rawLineWithBreak.slice(0, -1).replace(/\r$/, "")
        : rawLineWithBreak.replace(/\r$/, "");
      const commentMatch = rawLine.match(/^\s*#\s?(.*)$/);

      if (commentMatch) {
        entries.push({
          kind: COMMENT_ROW_KIND,
          position: lineOffset,
          text: commentMatch[1] || "",
        });
      }

      lineOffset += rawLineWithBreak.length;
    }
    return entries;
  }

  function collectCommentEntriesFromRecord(segment) {
    return collectCommentEntriesFromSegment(segment, 0);
  }

  function buildHeaders(rows) {
    const headers = ["Record", "Type"];
    const seen = new Set(headers);
    for (const row of rows) {
      if (row.kind === COMMENT_ROW_KIND) {
        continue;
      }

      for (const fieldName of row.explicitFieldValues.keys()) {
        if (!seen.has(fieldName)) {
          seen.add(fieldName);
          headers.push(fieldName);
        }
      }
    }
    return headers;
  }

  function getCellStyleIndex(isHeaderRow, backgroundToken) {
    const normalizedToken = normalizeBackgroundToken(backgroundToken);
    if (!normalizedToken) {
      return isHeaderRow ? 1 : 0;
    }
    const fillOffset = FILLED_BACKGROUND_OPTIONS.findIndex((option) => option.token === normalizedToken);
    if (fillOffset < 0) {
      return isHeaderRow ? 1 : 0;
    }
    return 2 + fillOffset * 2 + (isHeaderRow ? 1 : 0);
  }

  function buildSheetXml(sheet) {
    const headers = Array.isArray(sheet?.headers) ? sheet.headers : ["Record", "Type"];
    const allRows = [
      {
        values: headers,
        backgrounds: normalizeBackgrounds(sheet?.headerBackgrounds, headers.length),
        isHeaderRow: true,
      },
      ...((sheet?.rows || []).map((row) => serializeSheetRow(row, headers.length))),
    ];

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    xml += '<sheetViews><sheetView workbookViewId="0">';
    xml += '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>';
    xml += '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
    xml += "</sheetView></sheetViews><sheetData>";
    allRows.forEach((row, rowIndex) => {
      xml += `<row r="${rowIndex + 1}">`;
      row.values.forEach((value, columnIndex) => {
        const styleIndex = getCellStyleIndex(row.isHeaderRow, row.backgrounds[columnIndex]);
        if (!value && styleIndex === (row.isHeaderRow ? 1 : 0)) {
          return;
        }
        if (value) {
          xml += `<c r="${getCellReference(columnIndex, rowIndex)}" t="inlineStr" s="${styleIndex}"><is><t xml:space="preserve">${escapeXml(
            value,
          )}</t></is></c>`;
          return;
        }
        xml += `<c r="${getCellReference(columnIndex, rowIndex)}" s="${styleIndex}"/>`;
      });
      xml += "</row>";
    });
    xml += "</sheetData></worksheet>";
    return xml;
  }

  function serializeSheetRow(row, columnCount) {
    if (isCommentRow(row)) {
      const values = new Array(Math.max(2, columnCount)).fill("");
      values[0] = "Comment";
      values[1] = String(row.text || "");
      const backgrounds = new Array(values.length).fill("");
      const rowBackground = normalizeBackgroundToken(row.background);
      backgrounds[0] = rowBackground;
      backgrounds[1] = rowBackground;
      return {
        values,
        backgrounds,
        isHeaderRow: false,
      };
    }

    if (Array.isArray(row?.values)) {
      return {
        values: row.values.map((value) => String(value || "")),
        backgrounds: normalizeBackgrounds(row?.backgrounds, columnCount),
        isHeaderRow: false,
      };
    }

    return {
      values: Array.isArray(row) ? row.map((value) => String(value || "")) : new Array(columnCount).fill(""),
      backgrounds: normalizeBackgrounds(undefined, columnCount),
      isHeaderRow: false,
    };
  }

  function isCommentRow(row) {
    return !!row && !Array.isArray(row) && row.kind === COMMENT_ROW_KIND;
  }

  function buildContentTypesXml(sheets) {
    const worksheetOverrides = sheets
      .map(
        (_, index) =>
          `  <Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
      )
      .join("\n");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${worksheetOverrides}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }

  function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  function buildWorkbookXml(sheets) {
    const sheetEntries = sheets
      .map(
        (sheet, index) =>
          `    <sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
      )
      .join("\n");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${sheetEntries}
  </sheets>
</workbook>`;
  }

  function buildWorkbookRelsXml(sheets) {
    const worksheetRelationships = sheets
      .map(
        (_, index) =>
          `  <Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
      )
      .join("\n");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${worksheetRelationships}
  <Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  function buildStylesXml() {
    const fillEntries = FILLED_BACKGROUND_OPTIONS
      .map((option) => {
        return `    <fill><patternFill patternType="solid"><fgColor rgb="${getBackgroundArgb(option.token)}"/><bgColor indexed="64"/></patternFill></fill>`;
      })
      .join("\n");
    const cellXfEntries = [
      '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    ].concat(
      FILLED_BACKGROUND_OPTIONS.flatMap((_, index) => {
        const fillId = index + 2;
        return [
          `    <xf numFmtId="0" fontId="0" fillId="${fillId}" borderId="0" xfId="0" applyFill="1"/>`,
          `    <xf numFmtId="0" fontId="1" fillId="${fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`,
        ];
      }),
    ).join("\n");
    const fillCount = 2 + FILLED_BACKGROUND_OPTIONS.length;
    const cellXfCount = 2 + FILLED_BACKGROUND_OPTIONS.length * 2;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="${fillCount}">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
${fillEntries}
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="${cellXfCount}">
${cellXfEntries}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
  }

  function buildZipBuffer(parts) {
    const localFileChunks = [];
    const centralDirectoryChunks = [];
    let localOffset = 0;

    for (const [name, data] of parts.entries()) {
      const fileNameBuffer = Buffer.from(name, "utf8");
      const crc = crc32(data);
      const localHeader = Buffer.alloc(30);
      localHeader.writeUInt32LE(0x04034b50, 0);
      localHeader.writeUInt16LE(20, 4);
      localHeader.writeUInt16LE(0, 6);
      localHeader.writeUInt16LE(0, 8);
      localHeader.writeUInt16LE(0, 10);
      localHeader.writeUInt16LE(0, 12);
      localHeader.writeUInt32LE(crc >>> 0, 14);
      localHeader.writeUInt32LE(data.length, 18);
      localHeader.writeUInt32LE(data.length, 22);
      localHeader.writeUInt16LE(fileNameBuffer.length, 26);
      localHeader.writeUInt16LE(0, 28);
      localFileChunks.push(localHeader, fileNameBuffer, data);

      const centralHeader = Buffer.alloc(46);
      centralHeader.writeUInt32LE(0x02014b50, 0);
      centralHeader.writeUInt16LE(20, 4);
      centralHeader.writeUInt16LE(20, 6);
      centralHeader.writeUInt16LE(0, 8);
      centralHeader.writeUInt16LE(0, 10);
      centralHeader.writeUInt16LE(0, 12);
      centralHeader.writeUInt16LE(0, 14);
      centralHeader.writeUInt32LE(crc >>> 0, 16);
      centralHeader.writeUInt32LE(data.length, 20);
      centralHeader.writeUInt32LE(data.length, 24);
      centralHeader.writeUInt16LE(fileNameBuffer.length, 28);
      centralHeader.writeUInt16LE(0, 30);
      centralHeader.writeUInt16LE(0, 32);
      centralHeader.writeUInt16LE(0, 34);
      centralHeader.writeUInt16LE(0, 36);
      centralHeader.writeUInt32LE(0, 38);
      centralHeader.writeUInt32LE(localOffset, 42);
      centralDirectoryChunks.push(centralHeader, fileNameBuffer);

      localOffset += localHeader.length + fileNameBuffer.length + data.length;
    }

    const centralDirectory = Buffer.concat(centralDirectoryChunks);
    const localFiles = Buffer.concat(localFileChunks);
    const endOfCentralDirectory = Buffer.alloc(22);
    endOfCentralDirectory.writeUInt32LE(0x06054b50, 0);
    endOfCentralDirectory.writeUInt16LE(0, 4);
    endOfCentralDirectory.writeUInt16LE(0, 6);
    endOfCentralDirectory.writeUInt16LE(parts.size, 8);
    endOfCentralDirectory.writeUInt16LE(parts.size, 10);
    endOfCentralDirectory.writeUInt32LE(centralDirectory.length, 12);
    endOfCentralDirectory.writeUInt32LE(localFiles.length, 16);
    endOfCentralDirectory.writeUInt16LE(0, 20);

    return Buffer.concat([localFiles, centralDirectory, endOfCentralDirectory]);
  }

  function getCellReference(columnIndex, rowIndex) {
    return `${getColumnName(columnIndex)}${rowIndex + 1}`;
  }

  function getColumnName(index) {
    let value = index;
    let result = "";
    do {
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26) - 1;
    } while (value >= 0);
    return result;
  }

  function escapeXml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) {
        value = (value & 1) !== 0 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
      }
      table[index] = value >>> 0;
    }
    return table;
  })();

  function crc32(buffer) {
    let crc = 0xffffffff;
    for (const byte of buffer) {
      crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  return {
    buildDatabaseWorkbookBuffer,
    buildSheetModelFromDatabaseText,
    buildWorkbookBufferFromSheetModels,
  };
}

module.exports = {
  createDatabaseExcelTools,
};
