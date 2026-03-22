package org.epics.workbench.export

import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.testFramework.LightVirtualFile
import org.epics.workbench.filetypes.EpicsDatabaseFileType
import org.w3c.dom.Element
import org.w3c.dom.Node
import java.io.StringReader
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.attribute.BasicFileAttributes
import java.time.Instant
import java.time.format.DateTimeFormatter
import java.util.zip.ZipFile
import javax.xml.parsers.DocumentBuilderFactory
import org.xml.sax.InputSource

internal object EpicsDatabaseExcelImporter {
  internal data class ImportedSheet(
    val sheetName: String,
    val suggestedFileName: String,
    val text: String,
  )

  private data class WorkbookSheet(
    val name: String,
    val rows: Map<Int, Map<Int, String>>,
  )

  fun importWorkbook(sourcePath: Path): List<ImportedSheet> {
    val attributes = runCatching {
      Files.readAttributes(sourcePath, BasicFileAttributes::class.java)
    }.getOrNull()
    return ZipFile(sourcePath.toFile()).use { zipFile ->
      importWorkbook(
        zipFile = zipFile,
        sourceFileName = sourcePath.fileName.toString(),
        sourceCreatedAt = attributes?.creationTime()?.toInstant(),
        sourceModifiedAt = attributes?.lastModifiedTime()?.toInstant(),
        importedAt = Instant.now(),
      )
    }
  }

  fun openImportedSheets(project: com.intellij.openapi.project.Project, sheets: List<ImportedSheet>) {
    val fileEditorManager = FileEditorManager.getInstance(project)
    sheets.forEachIndexed { index, sheet ->
      val virtualFile = LightVirtualFile(sheet.suggestedFileName, EpicsDatabaseFileType.INSTANCE, sheet.text)
      fileEditorManager.openFile(virtualFile, index == sheets.lastIndex, true)
    }
  }

  private fun importWorkbook(
    zipFile: ZipFile,
    sourceFileName: String,
    sourceCreatedAt: Instant?,
    sourceModifiedAt: Instant?,
    importedAt: Instant,
  ): List<ImportedSheet> {
    val workbookXml = readRequiredEntry(zipFile, "xl/workbook.xml")
    val workbookRelsXml = readRequiredEntry(zipFile, "xl/_rels/workbook.xml.rels")
    val relationships = parseWorkbookRelationships(workbookRelsXml)
    val sharedStrings = zipFile.getEntry("xl/sharedStrings.xml")?.let { entry ->
      zipFile.getInputStream(entry).bufferedReader().use { parseSharedStrings(it.readText()) }
    }.orEmpty()

    return parseWorkbookSheets(workbookXml)
      .mapNotNull { sheet ->
        val targetPath = relationships[sheet.second] ?: return@mapNotNull null
        val worksheetPath = resolveWorkbookTargetPath(targetPath)
        val worksheetEntry = zipFile.getEntry(worksheetPath) ?: return@mapNotNull null
        val worksheetXml = zipFile.getInputStream(worksheetEntry).bufferedReader().use { it.readText() }
        WorkbookSheet(
          name = sheet.first,
          rows = parseWorksheet(worksheetXml, sharedStrings),
        )
      }
      .filter(::isEpicsDatabaseSheet)
      .mapIndexed { index, sheet ->
        ImportedSheet(
          sheetName = sheet.name,
          suggestedFileName = buildSuggestedFileName(sourceFileName, sheet.name, index),
          text = buildImportedDatabaseText(
            sheet = sheet,
            sourceFileName = sourceFileName,
            sourceCreatedAt = sourceCreatedAt,
            sourceModifiedAt = sourceModifiedAt,
            importedAt = importedAt,
          ),
        )
      }
  }

  private fun readRequiredEntry(zipFile: ZipFile, entryName: String): String {
    val entry = zipFile.getEntry(entryName)
      ?: throw IllegalArgumentException("The selected workbook is missing $entryName.")
    return zipFile.getInputStream(entry).bufferedReader().use { it.readText() }
  }

  private fun parseWorkbookRelationships(xml: String): Map<String, String> {
    val document = parseXml(xml)
    val relationships = linkedMapOf<String, String>()
    val nodes = document.getElementsByTagName("Relationship")
    for (index in 0 until nodes.length) {
      val element = nodes.item(index) as? Element ?: continue
      relationships[element.getAttribute("Id")] = element.getAttribute("Target")
    }
    return relationships
  }

  private fun parseWorkbookSheets(xml: String): List<Pair<String, String>> {
    val document = parseXml(xml)
    val sheets = mutableListOf<Pair<String, String>>()
    val nodes = document.getElementsByTagName("sheet")
    for (index in 0 until nodes.length) {
      val element = nodes.item(index) as? Element ?: continue
      sheets += element.getAttribute("name") to element.getAttribute("r:id")
    }
    return sheets
  }

  private fun parseSharedStrings(xml: String): List<String> {
    val document = parseXml(xml)
    val strings = mutableListOf<String>()
    val nodes = document.getElementsByTagName("si")
    for (index in 0 until nodes.length) {
      val element = nodes.item(index) as? Element ?: continue
      strings += extractInlineString(element)
    }
    return strings
  }

  private fun parseWorksheet(xml: String, sharedStrings: List<String>): Map<Int, Map<Int, String>> {
    val document = parseXml(xml)
    val rows = linkedMapOf<Int, MutableMap<Int, String>>()
    val cells = document.getElementsByTagName("c")
    for (index in 0 until cells.length) {
      val cell = cells.item(index) as? Element ?: continue
      val reference = cell.getAttribute("r")
      val parsedReference = parseCellReference(reference) ?: continue
      val row = rows.getOrPut(parsedReference.second) { linkedMapOf() }
      row[parsedReference.first] = parseCellValue(cell, sharedStrings)
    }
    return rows
  }

  private fun parseCellReference(reference: String): Pair<Int, Int>? {
    val match = Regex("""^([A-Z]+)(\d+)$""").matchEntire(reference) ?: return null
    return columnNameToIndex(match.groupValues[1]) to match.groupValues[2].toInt()
  }

  private fun parseCellValue(cell: Element, sharedStrings: List<String>): String {
    return when (cell.getAttribute("t")) {
      "inlineStr" -> extractInlineString(cell)
      "s" -> {
        val index = childText(cell, "v").toIntOrNull() ?: return ""
        sharedStrings.getOrElse(index) { "" }
      }
      else -> childText(cell, "v")
    }
  }

  private fun extractInlineString(element: Element): String {
    val parts = mutableListOf<String>()
    collectDescendantTextNodes(element, "t", parts)
    return parts.joinToString(separator = "")
  }

  private fun collectDescendantTextNodes(node: Node, tagName: String, destination: MutableList<String>) {
    if (node is Element && node.tagName == tagName) {
      destination += node.textContent.orEmpty()
    }
    val children = node.childNodes
    for (index in 0 until children.length) {
      collectDescendantTextNodes(children.item(index), tagName, destination)
    }
  }

  private fun childText(element: Element, tagName: String): String {
    val children = element.getElementsByTagName(tagName)
    return if (children.length > 0) children.item(0).textContent.orEmpty() else ""
  }

  private fun isEpicsDatabaseSheet(sheet: WorkbookSheet): Boolean {
    val headerRow = sheet.rows[1] ?: return false
    return headerRow[0]?.trim() == "Record" && headerRow[1]?.trim() == "Type"
  }

  private fun buildImportedDatabaseText(
    sheet: WorkbookSheet,
    sourceFileName: String,
    sourceCreatedAt: Instant?,
    sourceModifiedAt: Instant?,
    importedAt: Instant,
  ): String {
    val lines = mutableListOf(
      "# Imported from Excel by EPICS Workbench",
      "# Source file: $sourceFileName",
      "# Source sheet: ${sheet.name}",
      "# Source created time: ${formatTimestamp(sourceCreatedAt)}",
      "# Source modified time: ${formatTimestamp(sourceModifiedAt)}",
      "# Imported at: ${formatTimestamp(importedAt)}",
      "",
    )

    val headerRow = sheet.rows[1].orEmpty()
    val maxColumnIndex = headerRow.keys.maxOrNull() ?: -1
    val headers = (0..maxColumnIndex).map { headerRow[it].orEmpty().trim() }

    sheet.rows.keys.filter { it > 1 }.sorted().forEach { rowNumber ->
      val row = sheet.rows[rowNumber].orEmpty()
      val recordName = row[0].orEmpty().trim()
      val recordType = row[1].orEmpty().trim()
      if (recordName.isEmpty() || recordType.isEmpty()) {
        return@forEach
      }

      lines += """record($recordType, "${escapeDatabaseString(recordName)}") {"""
      headers.drop(2).forEachIndexed { fieldOffset, fieldName ->
        val columnIndex = fieldOffset + 2
        val value = row[columnIndex]
        if (fieldName.isBlank() || value.isNullOrBlank()) {
          return@forEachIndexed
        }
        lines += """    field($fieldName, "${escapeDatabaseString(value)}")"""
      }
      lines += "}"
      lines += ""
    }

    return lines.joinToString(separator = "\n", postfix = "\n")
  }

  private fun buildSuggestedFileName(sourceFileName: String, sheetName: String, index: Int): String {
    val baseName = sourceFileName.substringBeforeLast('.', sourceFileName)
    val safeSheetName = sanitizeFileNameFragment(sheetName.ifBlank { "sheet-${index + 1}" })
    return "${sanitizeFileNameFragment(baseName)}-$safeSheetName.db"
  }

  private fun sanitizeFileNameFragment(value: String): String {
    return value
      .replace(Regex("""[<>:"/\\|?*\u0000-\u001f]+"""), "_")
      .replace(Regex("""\s+"""), "_")
      .replace(Regex("""_+"""), "_")
      .trim('_')
      .ifEmpty { "sheet" }
  }

  private fun resolveWorkbookTargetPath(target: String): String {
    return if (target.startsWith("/")) {
      target.trimStart('/')
    } else {
      Path.of("xl").resolve(target).normalize().toString().replace('\\', '/')
    }
  }

  private fun columnNameToIndex(columnName: String): Int {
    var value = 0
    columnName.forEach { character ->
      value = value * 26 + (character.code - 'A'.code + 1)
    }
    return value - 1
  }

  private fun escapeDatabaseString(value: String): String {
    return value
      .replace("\\", "\\\\")
      .replace("\r\n", "\n")
      .replace("\r", "\n")
      .replace("\n", "\\n")
      .replace("\"", "\\\"")
  }

  private fun formatTimestamp(value: Instant?): String {
    return value?.let(DateTimeFormatter.ISO_INSTANT::format) ?: "unavailable"
  }

  private fun parseXml(xml: String) =
    documentBuilderFactory.newDocumentBuilder().parse(InputSource(StringReader(xml)))

  private val documentBuilderFactory = DocumentBuilderFactory.newInstance().apply {
    isNamespaceAware = false
  }
}
