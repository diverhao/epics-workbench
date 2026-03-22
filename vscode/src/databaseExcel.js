function createDatabaseExcelTools({
  extractRecordDeclarations,
  extractFieldDeclarationsInRecord,
}) {
  function buildDatabaseWorkbookBuffer(text) {
    const rows = buildRows(text);
    const headers = buildHeaders(rows);
    const parts = new Map([
      ["[Content_Types].xml", Buffer.from(buildContentTypesXml(), "utf8")],
      ["_rels/.rels", Buffer.from(buildRootRelsXml(), "utf8")],
      ["xl/workbook.xml", Buffer.from(buildWorkbookXml(), "utf8")],
      ["xl/_rels/workbook.xml.rels", Buffer.from(buildWorkbookRelsXml(), "utf8")],
      ["xl/styles.xml", Buffer.from(buildStylesXml(), "utf8")],
      ["xl/worksheets/sheet1.xml", Buffer.from(buildSheetXml(headers, rows), "utf8")],
    ]);
    return buildZipBuffer(parts);
  }

  function buildRows(text) {
    return extractRecordDeclarations(text).map((declaration) => {
      const explicitFieldValues = new Map();
      for (const fieldDeclaration of extractFieldDeclarationsInRecord(text, declaration)) {
        explicitFieldValues.set(
          fieldDeclaration.fieldName,
          fieldDeclaration.value,
        );
      }
      return {
        recordName: declaration.name,
        recordType: declaration.recordType,
        explicitFieldValues,
      };
    });
  }

  function buildHeaders(rows) {
    const headers = ["Record", "Type"];
    const seen = new Set(headers);
    for (const row of rows) {
      for (const fieldName of row.explicitFieldValues.keys()) {
        if (!seen.has(fieldName)) {
          seen.add(fieldName);
          headers.push(fieldName);
        }
      }
    }
    return headers;
  }

  function buildSheetXml(headers, rows) {
    const allRows = [
      headers,
      ...rows.map((row) => [
        row.recordName,
        row.recordType,
        ...headers.slice(2).map((fieldName) => row.explicitFieldValues.get(fieldName) || ""),
      ]),
    ];

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    xml += '<sheetViews><sheetView workbookViewId="0">';
    xml += '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>';
    xml += '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
    xml += "</sheetView></sheetViews><sheetData>";
    allRows.forEach((row, rowIndex) => {
      xml += `<row r="${rowIndex + 1}">`;
      row.forEach((value, columnIndex) => {
        if (!value) {
          return;
        }
        const styleIndex = rowIndex === 0 ? 1 : 0;
        xml += `<c r="${getCellReference(columnIndex, rowIndex)}" t="inlineStr" s="${styleIndex}"><is><t xml:space="preserve">${escapeXml(
          value,
        )}</t></is></c>`;
      });
      xml += "</row>";
    });
    xml += "</sheetData></worksheet>";
    return xml;
  }

  function buildContentTypesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }

  function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  function buildWorkbookXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Database" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
  }

  function buildWorkbookRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
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
  };
}

module.exports = {
  createDatabaseExcelTools,
};
