package org.epics.workbench.export

import org.epics.workbench.completion.EpicsRecordCompletionSupport
import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

internal object EpicsDatabaseExcelExporter {
  internal data class ExportRow(
    val recordName: String,
    val recordType: String,
    val explicitFieldValues: Map<String, String>,
  )

  fun buildWorkbook(text: String): ByteArray {
    val rows = buildRows(text)
    val headers = buildHeaders(rows)

    val parts = linkedMapOf<String, ByteArray>()
    parts["[Content_Types].xml"] = buildContentTypesXml().toByteArray(StandardCharsets.UTF_8)
    parts["_rels/.rels"] = buildRootRelsXml().toByteArray(StandardCharsets.UTF_8)
    parts["xl/workbook.xml"] = buildWorkbookXml().toByteArray(StandardCharsets.UTF_8)
    parts["xl/_rels/workbook.xml.rels"] = buildWorkbookRelsXml().toByteArray(StandardCharsets.UTF_8)
    parts["xl/styles.xml"] = buildStylesXml().toByteArray(StandardCharsets.UTF_8)
    parts["xl/worksheets/sheet1.xml"] = buildSheetXml(headers, rows).toByteArray(StandardCharsets.UTF_8)

    return buildZip(parts)
  }

  private fun buildRows(text: String): List<ExportRow> {
    return EpicsRecordCompletionSupport.extractRecordDeclarations(text).map { declaration ->
      val explicitFields = linkedMapOf<String, String>()
      EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, declaration).forEach { field ->
        explicitFields[field.fieldName] = field.value
      }
      ExportRow(
        recordName = declaration.name,
        recordType = declaration.recordType,
        explicitFieldValues = explicitFields,
      )
    }
  }

  private fun buildHeaders(rows: List<ExportRow>): List<String> {
    val headers = linkedSetOf("Record", "Type")
    rows.forEach { row -> headers.addAll(row.explicitFieldValues.keys) }
    return headers.toList()
  }

  private fun buildSheetXml(headers: List<String>, rows: List<ExportRow>): String {
    val allRows = buildList {
      add(headers)
      rows.forEach { row ->
        add(
          buildList {
            add(row.recordName)
            add(row.recordType)
            headers.drop(2).forEach { fieldName ->
              add(row.explicitFieldValues[fieldName].orEmpty())
            }
          },
        )
      }
    }

    return buildString {
      append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
      append("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""")
      append("""<sheetViews><sheetView workbookViewId="0">""")
      append("""<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>""")
      append("""<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>""")
      append("""<sheetData>""")
      allRows.forEachIndexed { rowIndex, row ->
        append("""<row r="${rowIndex + 1}">""")
        row.forEachIndexed rowLoop@{ columnIndex, value ->
          if (value.isEmpty()) {
            return@rowLoop
          }
          val cellReference = getCellReference(columnIndex, rowIndex)
          val styleIndex = if (rowIndex == 0) 1 else 0
          append("""<c r="$cellReference" t="inlineStr" s="$styleIndex"><is><t xml:space="preserve">""")
          append(escapeXml(value))
          append("""</t></is></c>""")
        }
        append("</row>")
      }
      append("""</sheetData></worksheet>""")
    }
  }

  private fun buildContentTypesXml(): String = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>
  """.trimIndent()

  private fun buildRootRelsXml(): String = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>
  """.trimIndent()

  private fun buildWorkbookXml(): String = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
        <sheet name="Database" sheetId="1" r:id="rId1"/>
      </sheets>
    </workbook>
  """.trimIndent()

  private fun buildWorkbookRelsXml(): String = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>
  """.trimIndent()

  private fun buildStylesXml(): String = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    </styleSheet>
  """.trimIndent()

  private fun buildZip(parts: Map<String, ByteArray>): ByteArray {
    val output = ByteArrayOutputStream()
    ZipOutputStream(output).use { zipOutput ->
      parts.forEach { (path, data) ->
        val entry = ZipEntry(path)
        zipOutput.putNextEntry(entry)
        zipOutput.write(data)
        zipOutput.closeEntry()
      }
    }
    return output.toByteArray()
  }

  private fun getCellReference(columnIndex: Int, rowIndex: Int): String {
    return columnName(columnIndex) + (rowIndex + 1)
  }

  private fun columnName(index: Int): String {
    var value = index
    val builder = StringBuilder()
    do {
      builder.append(('A'.code + (value % 26)).toChar())
      value = value / 26 - 1
    } while (value >= 0)
    return builder.reverse().toString()
  }

  private fun escapeXml(value: String): String {
    return buildString(value.length) {
      value.forEach { character ->
        when (character) {
          '&' -> append("&amp;")
          '<' -> append("&lt;")
          '>' -> append("&gt;")
          '"' -> append("&quot;")
          '\'' -> append("&apos;")
          else -> append(character)
        }
      }
    }
  }
}
