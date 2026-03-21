package org.epics.workbench.toc

import org.epics.workbench.completion.EpicsRecordCompletionSupport

internal object EpicsDatabaseToc {
  private const val TOC_BEGIN_MARKER = "# EPICS TOC BEGIN"
  private const val TOC_END_MARKER = "# EPICS TOC END"
  private const val VALUE_COLUMN_INDEX = 1
  private const val VALUE_COLUMN_WIDTH = 18
  private val TOC_HEADER_REGEX = Regex("""^#\s*Table of Contents(?:[ \t]+(.*?))?\s*$""", RegexOption.MULTILINE)
  private val TOC_MACRO_ASSIGNMENT_REGEX = Regex(
    """^#\s*-\s*([A-Za-z_][A-Za-z0-9_]*)(?:\s*=\s*(.*))?\s*$""",
  )
  private val EPICS_VARIABLE_REGEX = Regex(
    """\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""",
  )
  private val LINE_REGEX = Regex("^.*$", RegexOption.MULTILINE)

  internal data class TocRecordReference(
    val recordType: String,
    val recordName: String,
  )

  internal data class TocRuntimeEntry(
    val recordType: String,
    val recordName: String,
    val valueStart: Int,
    val valueEnd: Int,
  )

  fun upsert(text: String, eol: String): String {
    val contentWithoutToc = removeBlock(text).replace(Regex("""^(?:[ \t]*\r?\n)+"""), "")
    val tocBlock = buildBlock(text, eol)
    return if (contentWithoutToc.isEmpty()) {
      "$tocBlock$eol"
    } else {
      "$tocBlock$eol$eol$contentWithoutToc"
    }
  }

  internal fun findRecordReferenceAtTypeOffset(text: String, offset: Int): TocRecordReference? {
    val range = findBlockRange(text) ?: return null
    if (offset !in range.start until range.endExclusive) {
      return null
    }

    val declarations = extractRecordDeclarations(text)
    val tocText = text.substring(range.start, range.endExclusive)
    for (lineMatch in LINE_REGEX.findAll(tocText)) {
      val lineText = lineMatch.value
      val lineOffset = range.start + lineMatch.range.first
      val entry = parseMarkdownTocEntry(lineText, lineOffset, declarations) ?: continue
      if (offset in entry.typeStart..entry.typeEnd) {
        return TocRecordReference(
          recordType = entry.recordType,
          recordName = entry.recordName,
        )
      }
    }

    return null
  }

  internal fun extractRuntimeEntries(text: String): List<TocRuntimeEntry> {
    val range = findBlockRange(text) ?: return emptyList()
    val declarations = extractRecordDeclarations(text)
    val tocText = text.substring(range.start, range.endExclusive)
    val entries = mutableListOf<TocRuntimeEntry>()

    for (lineMatch in LINE_REGEX.findAll(tocText)) {
      val lineText = lineMatch.value
      val lineOffset = range.start + lineMatch.range.first
      val entry = parseMarkdownTocEntry(lineText, lineOffset, declarations) ?: continue
      val valueStart = entry.valueStart ?: continue
      val valueEnd = entry.valueEnd ?: continue
      entries += TocRuntimeEntry(
        recordType = entry.recordType,
        recordName = entry.recordName,
        valueStart = valueStart,
        valueEnd = valueEnd,
      )
    }

    return entries
  }

  internal fun extractRuntimeMacroAssignments(text: String): Map<String, MacroAssignment> {
    return extractMacroAssignments(text)
  }

  private fun buildBlock(text: String, eol: String): String {
    val extraFieldNames = getExtraFieldNames(text)
    val macroNames = getMacroNames(text)
    val macroAssignments = extractMacroAssignments(text)
    val lines = mutableListOf<String>()
    val headerSuffix = if (extraFieldNames.isEmpty()) "" else " ${extraFieldNames.joinToString(" ")}"

    lines += TOC_BEGIN_MARKER
    if (macroNames.isNotEmpty()) {
      lines += "# Macros:"
      for (macroName in macroNames) {
        val assignment = macroAssignments[macroName]
        lines += if (assignment?.hasAssignment == true) {
          "#  - $macroName = ${assignment.value}"
        } else {
          "#  - $macroName"
        }
      }
    }

    lines += "# Table of Contents$headerSuffix"

    val headerRow = mutableListOf("Record", "Value", "Type").apply {
      addAll(extraFieldNames)
    }
    val rows = extractRecordDeclarations(text).map { declaration ->
      buildRowValues(text, declaration, extraFieldNames)
    }
    val columnWidths = getColumnWidths(listOf(headerRow) + rows)

    lines += "# ${formatMarkdownRow(headerRow, columnWidths)}"
    lines += "# ${formatSeparatorRow(columnWidths)}"
    for (row in rows) {
      lines += "# ${formatMarkdownRow(row, columnWidths)}"
    }
    lines += TOC_END_MARKER

    return lines.joinToString(eol)
  }

  private fun getMacroNames(text: String): List<String> {
    val names = linkedSetOf<String>()
    EPICS_VARIABLE_REGEX.findAll(maskComments(text)).forEach { match ->
      val name = match.groups[1]?.value
        ?: match.groups[3]?.value
        ?: match.groups[5]?.value
        ?: return@forEach
      if (name.isNotBlank()) {
        names += name
      }
    }
    return names.toList().sorted()
  }

  private fun extractMacroAssignments(text: String): Map<String, MacroAssignment> {
    val range = findBlockRange(text) ?: return emptyMap()
    val assignments = linkedMapOf<String, MacroAssignment>()
    var inMacroSection = false

    for (line in text.substring(range.start, range.endExclusive).split(Regex("""\r?\n"""))) {
      if (line.matches(Regex("""^#\s*Macros:\s*$"""))) {
        inMacroSection = true
        continue
      }
      if (!inMacroSection) {
        continue
      }
      if (line.matches(Regex("""^#\s*Table of Contents(?:[ \t]+.*)?\s*$"""))) {
        break
      }

      val match = TOC_MACRO_ASSIGNMENT_REGEX.matchEntire(line) ?: continue
      val name = match.groups[1]?.value ?: continue
      assignments[name] = MacroAssignment(
        hasAssignment = match.groups[2] != null,
        value = match.groups[2]?.value.orEmpty(),
      )
    }

    return assignments
  }

  private fun getExtraFieldNames(text: String): List<String> {
    val searchableText = findBlockRange(text)
      ?.let { text.substring(it.start, it.endExclusive) }
      ?: text
    val headerSuffix = TOC_HEADER_REGEX.find(searchableText)?.groups?.get(1)?.value ?: return emptyList()
    if (headerSuffix.isBlank()) {
      return emptyList()
    }

    val names = linkedSetOf<String>()
    headerSuffix.split(Regex("""\s+"""))
      .map { it.trim().uppercase() }
      .filter { it.matches(Regex("""^[A-Z0-9_]+$""")) }
      .forEach(names::add)
    return names.toList()
  }

  private fun buildRowValues(
    text: String,
    declaration: RecordDeclaration,
    extraFieldNames: List<String>,
  ): List<String?> {
    val fieldsByName = extractFieldDeclarationsInRecord(text, declaration)
      .associateBy { it.fieldName.uppercase() }
    return buildList {
      add(declaration.name)
      add(null)
      add(declaration.recordType)
      for (fieldName in extraFieldNames) {
        add(
          fieldsByName[fieldName]?.value
            ?: getMissingFieldValue(declaration.recordType, fieldName),
        )
      }
    }
  }

  private fun getMissingFieldValue(recordType: String, fieldName: String): String {
    val knownFields = EpicsRecordCompletionSupport.getDeclaredFieldNamesForRecordType(recordType)
    if (knownFields?.contains(fieldName.uppercase()) != true) {
      return "NA"
    }
    return EpicsRecordCompletionSupport.getDefaultFieldValue(recordType, fieldName)
  }

  private fun getColumnWidths(rows: List<List<String?>>): MutableList<Int> {
    val widths = mutableListOf<Int>()
    rows.forEach { row ->
      row.forEachIndexed { index, value ->
        val valueLength = formatValue(value).length
        if (index >= widths.size) {
          widths += valueLength
        } else {
          widths[index] = maxOf(widths[index], valueLength)
        }
      }
    }
    while (widths.size <= VALUE_COLUMN_INDEX) {
      widths += 0
    }
    widths[VALUE_COLUMN_INDEX] = VALUE_COLUMN_WIDTH
    return widths
  }

  private fun formatMarkdownRow(row: List<String?>, widths: List<Int>): String {
    return row.mapIndexed { index, value ->
      formatValue(value).padEnd(widths[index])
    }.joinToString(prefix = "| ", separator = " | ", postfix = " |")
  }

  private fun formatSeparatorRow(widths: List<Int>): String {
    return widths.joinToString(prefix = "| ", separator = " | ", postfix = " |") {
      "-".repeat(maxOf(3, it))
    }
  }

  private fun formatValue(value: String?): String {
    return when {
      value == null -> ""
      value.isEmpty() -> "\"\""
      else -> value
    }
  }

  private fun parseMarkdownTocEntry(
    lineText: String,
    lineOffset: Int,
    declarations: List<RecordDeclaration>,
  ): TocEntry? {
    val cellEntries = extractCommentTableCells(lineText, lineOffset)
    if (cellEntries.size < 2) {
      return null
    }

    val cellValues = cellEntries.map { it.value }
    if (
      cellValues.all { value -> value.matches(Regex("""^:?-{3,}:?$""")) } ||
      cellValues[0].equals("record", ignoreCase = true)
    ) {
      return null
    }

    for (declaration in declarations) {
      if (cellValues[0] != declaration.name) {
        continue
      }

      val hasValueColumn = cellEntries.size >= 3 && cellValues[2] == declaration.recordType
      val typeCell = if (hasValueColumn) cellEntries[2] else cellEntries[1]
      if (typeCell.value != declaration.recordType) {
        continue
      }

      return TocEntry(
        recordType = declaration.recordType,
        recordName = declaration.name,
        typeStart = typeCell.start,
        typeEnd = typeCell.end,
        valueStart = if (hasValueColumn) cellEntries[1].displayStart else null,
        valueEnd = if (hasValueColumn) cellEntries[1].displayEnd else null,
      )
    }

    return null
  }

  private fun extractCommentTableCells(lineText: String, lineOffset: Int): List<CommentTableCell> {
    val prefixMatch = Regex("""^#\s*(\|.*)$""").find(lineText) ?: return emptyList()
    val tableText = prefixMatch.groups[1]?.value ?: return emptyList()
    val tableOffset = lineOffset + lineText.indexOf(tableText)
    val cells = mutableListOf<CommentTableCell>()
    var cellStart = tableOffset + 1

    for (index in 1 until tableText.length) {
      if (tableText[index] != '|') {
        continue
      }

      val rawCell = tableText.substring(cellStart - tableOffset, index)
      val leadingWhitespaceLength = rawCell.takeWhile { it.isWhitespace() }.length
      val trailingWhitespaceLength = rawCell.reversed().takeWhile { it.isWhitespace() }.length
      val trimmedStart = cellStart + leadingWhitespaceLength
      val trimmedEnd = tableOffset + index - trailingWhitespaceLength
      val displayStart = cellStart + minOf(leadingWhitespaceLength, 1)
      val displayEnd = tableOffset + index - minOf(trailingWhitespaceLength, 1)
      cells += CommentTableCell(
        value = rawCell.trim(),
        start = trimmedStart,
        end = maxOf(trimmedStart, trimmedEnd),
        displayStart = displayStart,
        displayEnd = maxOf(displayStart, displayEnd),
      )
      cellStart = tableOffset + index + 1
    }

    return cells
  }

  private fun removeBlock(text: String): String {
    val range = findBlockRange(text) ?: return text
    return buildString(text.length - (range.endExclusive - range.start)) {
      append(text.substring(0, range.start))
      append(text.substring(range.endExclusive))
    }
  }

  private fun findBlockRange(text: String): TocRange? {
    val start = text.indexOf(TOC_BEGIN_MARKER)
    if (start < 0) {
      return null
    }

    val endMarkerStart = text.indexOf(TOC_END_MARKER, start)
    if (endMarkerStart < 0) {
      return null
    }

    var end = endMarkerStart + TOC_END_MARKER.length
    if (text.startsWith("\r\n", end)) {
      end += 2
    } else if (end < text.length && text[end] == '\n') {
      end += 1
    }

    while (text.startsWith("\r\n", end)) {
      end += 2
    }
    while (end < text.length && text[end] == '\n') {
      end += 1
    }

    return TocRange(start = start, endExclusive = end)
  }

  private fun extractRecordDeclarations(text: String): List<RecordDeclaration> {
    val declarations = mutableListOf<RecordDeclaration>()
    val sanitizedText = maskComments(text)
    val regex = Regex("""\b(?:g?record)\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"""")
    val prefixRegex = Regex("""(?:g?record)\(\s*[A-Za-z0-9_]+\s*,\s*"""")

    for (match in regex.findAll(sanitizedText)) {
      val recordType = match.groups[1]?.value.orEmpty()
      val recordName = match.groups[2]?.value.orEmpty()
      val prefixLength = prefixRegex.find(match.value)?.value?.length ?: continue
      val recordStart = match.range.first
      val recordEnd = findRecordBlockEnd(sanitizedText, recordStart)
      declarations += RecordDeclaration(
        recordType = recordType,
        name = recordName,
        nameStart = match.range.first + prefixLength,
        nameEnd = match.range.first + prefixLength + recordName.length,
        recordStart = recordStart,
        recordEnd = recordEnd,
      )
    }

    return declarations
  }

  private fun extractFieldDeclarationsInRecord(
    text: String,
    recordDeclaration: RecordDeclaration,
  ): List<FieldDeclaration> {
    val declarations = mutableListOf<FieldDeclaration>()
    val sanitizedText = maskComments(text)
    val recordText = sanitizedText.substring(recordDeclaration.recordStart, recordDeclaration.recordEnd)
    val regex = Regex("""field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"""")
    val valuePrefixRegex = Regex("""field\(\s*(?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+)\s*,\s*"""")

    for (match in regex.findAll(recordText)) {
      val fieldName = (match.groups[1]?.value ?: match.groups[2]?.value ?: continue).uppercase()
      val valuePrefixLength = valuePrefixRegex.find(match.value)?.value?.length ?: continue
      val value = match.groups[3]?.value.orEmpty()
      declarations += FieldDeclaration(
        fieldName = fieldName,
        value = value,
        valueStart = recordDeclaration.recordStart + match.range.first + valuePrefixLength,
        valueEnd = recordDeclaration.recordStart + match.range.first + valuePrefixLength + value.length,
      )
    }

    return declarations
  }

  private fun maskComments(text: String): String {
    val sanitized = StringBuilder(text.length)
    var inString = false
    var escaped = false
    var inComment = false

    for (character in text) {
      if (inComment) {
        if (character == '\n') {
          inComment = false
          sanitized.append(character)
        } else {
          sanitized.append(' ')
        }
        continue
      }

      if (escaped) {
        sanitized.append(character)
        escaped = false
        continue
      }

      when {
        character == '\\' -> {
          sanitized.append(character)
          escaped = true
        }

        character == '"' -> {
          inString = !inString
          sanitized.append(character)
        }

        !inString && character == '#' -> {
          inComment = true
          sanitized.append(' ')
        }

        else -> sanitized.append(character)
      }
    }

    return sanitized.toString()
  }

  private fun findRecordBlockEnd(text: String, recordStart: Int): Int {
    val openingBraceIndex = text.indexOf('{', recordStart)
    if (openingBraceIndex < 0) {
      return recordStart
    }
    return findBalancedBlockEnd(text, openingBraceIndex)
  }

  private fun findBalancedBlockEnd(text: String, openingBraceIndex: Int): Int {
    var depth = 0
    var inString = false
    var escaped = false
    var inComment = false

    for (index in openingBraceIndex until text.length) {
      val character = text[index]

      if (inComment) {
        if (character == '\n') {
          inComment = false
        }
        continue
      }

      if (escaped) {
        escaped = false
        continue
      }

      when {
        character == '\\' -> escaped = true
        character == '"' -> inString = !inString
        inString -> Unit
        character == '#' -> inComment = true
        character == '{' -> depth += 1
        character == '}' -> {
          depth -= 1
          if (depth == 0) {
            return index + 1
          }
        }
      }
    }

    return text.length
  }

  internal data class MacroAssignment(
    val hasAssignment: Boolean,
    val value: String,
  )

  private data class TocEntry(
    val recordType: String,
    val recordName: String,
    val typeStart: Int,
    val typeEnd: Int,
    val valueStart: Int?,
    val valueEnd: Int?,
  )

  private data class CommentTableCell(
    val value: String,
    val start: Int,
    val end: Int,
    val displayStart: Int,
    val displayEnd: Int,
  )

  private data class TocRange(
    val start: Int,
    val endExclusive: Int,
  )

  private data class RecordDeclaration(
    val recordType: String,
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
    val recordStart: Int,
    val recordEnd: Int,
  )

  private data class FieldDeclaration(
    val fieldName: String,
    val value: String,
    val valueStart: Int,
    val valueEnd: Int,
  )
}
