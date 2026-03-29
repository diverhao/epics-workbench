package org.epics.workbench.formatting

internal object EpicsTextFormatter {
  fun formatDatabaseText(text: String, indentUnit: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()
    var indentLevel = 0

    for (line in lines) {
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        continue
      }

      val effectiveIndentLevel = if (trimmedLine.startsWith("}")) {
        maxOf(indentLevel - 1, 0)
      } else {
        indentLevel
      }
      val formattedLine = formatDatabaseLine(trimmedLine)
      formattedLines += indentUnit.repeat(effectiveIndentLevel) + formattedLine
      indentLevel = maxOf(0, indentLevel + getBraceDeltaOutsideStrings(trimmedLine))
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatSubstitutionText(text: String, indentUnit: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()
    val state = SubstitutionFormattingState()
    var indentLevel = 0
    var lineIndex = 0

    while (lineIndex < lines.size) {
      val line = lines[lineIndex]
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        lineIndex += 1
        continue
      }

      val lineParts = splitSubstitutionLineComment(trimmedLine)
      val trimmedCode = lineParts.code.trim()
      val trimmedComment = lineParts.comment.trim()
      val startsWithClosingBrace = trimmedCode.isNotEmpty() && trimmedCode.startsWith("}")
      val effectiveIndentLevel = if (startsWithClosingBrace) {
        maxOf(indentLevel - 1, 0)
      } else {
        indentLevel
      }
      val indentation = indentUnit.repeat(effectiveIndentLevel)

      if (trimmedCode.isEmpty()) {
        formattedLines += indentation + trimmedComment
        lineIndex += 1
        continue
      }

      if (isSubstitutionBlockLine(trimmedCode)) {
        state.assignmentOrder = null
      }

      val patternBlock = formatAlignedSubstitutionPatternBlock(
        lines = lines,
        startIndex = lineIndex,
        indentUnit = indentUnit,
        effectiveIndentLevel = effectiveIndentLevel,
      )
      if (patternBlock != null) {
        formattedLines += patternBlock.lines
        lineIndex = patternBlock.nextLineIndex + 1
        continue
      }

      val formattedCode = formatSubstitutionLine(trimmedCode, state)
      formattedLines += buildString {
        append(indentation)
        append(formattedCode)
        if (trimmedComment.isNotEmpty()) {
          append(' ')
          append(trimmedComment)
        }
      }
      indentLevel = maxOf(0, indentLevel + getBraceDeltaOutsideStrings(trimmedCode))
      if (indentLevel == 0) {
        state.assignmentOrder = null
      }
      lineIndex += 1
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatMonitorText(text: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()

    for (line in lines) {
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        continue
      }

      if (trimmedLine.startsWith("#")) {
        val commentText = trimmedLine.drop(1).trimStart()
        formattedLines += if (commentText.isEmpty()) "#" else "# $commentText"
        continue
      }

      val macroMatch = Regex("""^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(\S+)\s*$""").matchEntire(trimmedLine)
      if (macroMatch != null) {
        formattedLines += "${macroMatch.groupValues[1]} = ${macroMatch.groupValues[2]}"
        continue
      }

      formattedLines += trimmedLine
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatStartupText(text: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()

    for ((lineIndex, line) in lines.withIndex()) {
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        continue
      }

      if (lineIndex == 0 && trimmedLine.startsWith("#!")) {
        formattedLines += trimmedLine
        continue
      }

      val lineParts = splitStartupLineComment(trimmedLine)
      val trimmedCode = lineParts.code.trim()
      val trimmedComment = lineParts.comment.trim()
      if (trimmedCode.isEmpty()) {
        formattedLines += trimmedComment
        continue
      }

      val formattedCode = formatStartupLine(trimmedCode)
      formattedLines += buildString {
        append(formattedCode)
        if (trimmedComment.isNotEmpty()) {
          append(' ')
          append(trimmedComment)
        }
      }
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatMakefileText(text: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()

    for (line in lines) {
      val withoutTrailingWhitespace = line.replace(Regex("""[ \t]+$"""), "")
      if (withoutTrailingWhitespace.isBlank()) {
        formattedLines += ""
        continue
      }

      if (withoutTrailingWhitespace.startsWith('\t')) {
        formattedLines += withoutTrailingWhitespace
        continue
      }

      val trimmedLine = withoutTrailingWhitespace.trim()
      if (trimmedLine.startsWith("#")) {
        val commentText = trimmedLine.drop(1).trimStart()
        formattedLines += if (commentText.isEmpty()) "#" else "# $commentText"
        continue
      }

      val normalizedIncludeLine = normalizeMakefileIncludeLine(trimmedLine)
      if (normalizedIncludeLine != null) {
        formattedLines += normalizedIncludeLine
        continue
      }

      val normalizedAssignmentLine = normalizeMakefileAssignmentLine(trimmedLine)
      if (normalizedAssignmentLine != null) {
        formattedLines += normalizedAssignmentLine
        continue
      }

      formattedLines += trimmedLine
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatProtocolText(text: String, indentUnit: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()
    var indentLevel = 0

    for (line in lines) {
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        continue
      }

      val lineParts = splitProtocolLineComment(trimmedLine)
      val trimmedCode = lineParts.code.trim()
      val trimmedComment = lineParts.comment.trim()
      val effectiveIndentLevel = if (trimmedCode.startsWith("}")) {
        maxOf(indentLevel - 1, 0)
      } else {
        indentLevel
      }
      val indentation = indentUnit.repeat(effectiveIndentLevel)

      if (trimmedCode.isEmpty()) {
        formattedLines += indentation + trimmedComment
        continue
      }

      val formattedCode = formatProtocolLine(trimmedCode)
      formattedLines += buildString {
        append(indentation)
        append(formattedCode)
        if (trimmedComment.isNotEmpty()) {
          append(' ')
          append(trimmedComment)
        }
      }
      indentLevel = maxOf(0, indentLevel + getBraceDeltaOutsideProtocolStrings(trimmedCode))
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  fun formatSequencerText(text: String): String {
    val normalizedText = text.replace("\r\n", "\n")
    val hadTrailingNewline = normalizedText.endsWith("\n")
    val contentText = if (hadTrailingNewline) normalizedText.dropLast(1) else normalizedText
    val lines = if (contentText.isEmpty()) emptyList() else contentText.split('\n')
    val formattedLines = mutableListOf<String>()
    val indentUnit = " ".repeat(4)
    var indentLevel = 0
    var inBlockComment = false

    for (line in lines) {
      val trimmedLine = line.trim()
      if (trimmedLine.isEmpty()) {
        formattedLines += ""
        continue
      }

      val logicalLines = expandSequencerLogicalLines(trimmedLine, inBlockComment)
      for (logicalLine in logicalLines) {
        val braceInfo = getSequencerBraceInfo(logicalLine, inBlockComment)
        val effectiveIndentLevel = if (braceInfo.startsWithClosingBrace) {
          maxOf(indentLevel - 1, 0)
        } else {
          indentLevel
        }
        val indentation = if (isSequencerColumnZeroLine(logicalLine)) "" else indentUnit.repeat(effectiveIndentLevel)
        formattedLines += indentation + logicalLine
        indentLevel = maxOf(0, indentLevel + braceInfo.delta)
        inBlockComment = braceInfo.inBlockComment
      }
    }

    return formattedLines.joinToString("\n").let { if (hadTrailingNewline) "$it\n" else it }
  }

  private fun formatAlignedSubstitutionPatternBlock(
    lines: List<String>,
    startIndex: Int,
    indentUnit: String,
    effectiveIndentLevel: Int,
  ): FormattedPatternBlock? {
    val trimmedHeaderLine = lines[startIndex].trim()
    val headerParts = splitSubstitutionLineComment(trimmedHeaderLine)
    val headerValues = getSubstitutionPatternHeaderValues(headerParts.code.trim()) ?: return null

    val rowEntries = mutableListOf<PatternRowEntry>()
    var lineIndex = startIndex + 1
    while (lineIndex < lines.size) {
      val trimmedLine = lines[lineIndex].trim()
      if (trimmedLine.isEmpty()) {
        break
      }

      val rowParts = splitSubstitutionLineComment(trimmedLine)
      val trimmedCode = rowParts.code.trim()
      if (trimmedCode.isEmpty()) {
        break
      }

      val rowValues = getSubstitutionPatternRowValues(trimmedCode) ?: break
      rowEntries += PatternRowEntry(rowValues, rowParts.comment.trim())
      lineIndex += 1
    }

    if (rowEntries.isEmpty()) {
      return null
    }

    val normalizedHeaderValues = headerValues.map { it.trim() }
    val normalizedRowValues = rowEntries.map { row -> row.values.map(::formatSubstitutionScalarValue) }
    val columnCount = maxOf(
      normalizedHeaderValues.size,
      normalizedRowValues.maxOfOrNull { it.size } ?: 0,
    )
    val columnWidths = (0 until columnCount).map { columnIndex ->
      maxOf(
        normalizedHeaderValues.getOrNull(columnIndex)?.length ?: 0,
        normalizedRowValues.maxOfOrNull { values -> values.getOrNull(columnIndex)?.length ?: 0 } ?: 0,
      )
    }

    val indentation = indentUnit.repeat(effectiveIndentLevel)
    val patternRowIndentation = indentation + " ".repeat("pattern ".length)
    val formattedLines = mutableListOf<String>()
    formattedLines += buildString {
      append(indentation)
      append("pattern { ")
      append(formatAlignedSubstitutionPatternCells(normalizedHeaderValues, columnWidths))
      append(" }")
      if (headerParts.comment.isNotBlank()) {
        append(' ')
        append(headerParts.comment.trim())
      }
    }

    rowEntries.forEachIndexed { rowIndex, rowEntry ->
      formattedLines += buildString {
        append(patternRowIndentation)
        append("{ ")
        append(formatAlignedSubstitutionPatternCells(normalizedRowValues[rowIndex], columnWidths))
        append(" }")
        if (rowEntry.comment.isNotBlank()) {
          append(' ')
          append(rowEntry.comment)
        }
      }
    }

    return FormattedPatternBlock(formattedLines, lineIndex - 1)
  }

  private fun getSubstitutionPatternHeaderValues(trimmedLine: String): List<String>? {
    val match = Regex("""^pattern\s*\{(.*)\}\s*$""").matchEntire(trimmedLine) ?: return null
    return splitSubstitutionCommaSeparatedItems(match.groupValues[1])
  }

  private fun getSubstitutionPatternRowValues(trimmedLine: String): List<String>? {
    val match = Regex("""^\{(.*)\}\s*$""").matchEntire(trimmedLine) ?: return null
    if (parseSubstitutionAssignmentEntries(match.groupValues[1]) != null) {
      return null
    }
    return splitSubstitutionCommaSeparatedItems(match.groupValues[1])
  }

  private fun formatAlignedSubstitutionPatternCells(values: List<String>, columnWidths: List<Int>): String {
    return buildString {
      values.forEachIndexed { columnIndex, rawValue ->
        val text = rawValue.trim()
        if (columnIndex >= values.lastIndex) {
          append(text)
        } else {
          val paddingWidth = maxOf(0, (columnWidths.getOrNull(columnIndex) ?: 0) - text.length)
          append(text)
          append(", ")
          append(" ".repeat(paddingWidth))
        }
      }
    }
  }

  private fun formatDatabaseLine(trimmedLine: String): String {
    if (trimmedLine.startsWith("#")) {
      return trimmedLine
    }

    normalizeDatabaseRecordLine(trimmedLine)?.let { return it }
    normalizeDatabaseFieldLine(trimmedLine)?.let { return it }
    normalizeClosingBraceLine(trimmedLine)?.let { return it }
    return trimmedLine
  }

  private fun formatSubstitutionLine(trimmedLine: String, state: SubstitutionFormattingState): String {
    if (trimmedLine.startsWith("#")) {
      return trimmedLine
    }

    normalizeSubstitutionBlockLine(trimmedLine)?.let { return it }
    normalizeSubstitutionPatternLine(trimmedLine)?.let { return it }
    normalizeSubstitutionRowLine(trimmedLine, state)?.let { return it }
    normalizeClosingBraceLine(trimmedLine)?.let { return it }
    return trimmedLine
  }

  private fun formatProtocolLine(trimmedLine: String): String {
    if (trimmedLine.startsWith("#")) {
      return trimmedLine
    }

    normalizeProtocolBlockLine(trimmedLine)?.let { return it }
    normalizeProtocolAssignmentLine(trimmedLine)?.let { return it }
    normalizeClosingBraceLine(trimmedLine)?.let { return it }
    return trimmedLine
  }

  private fun formatStartupLine(trimmedLine: String): String {
    if (trimmedLine.startsWith("#")) {
      if (trimmedLine.startsWith("#!")) {
        return trimmedLine
      }
      val commentText = trimmedLine.drop(1).trimStart()
      return if (commentText.isEmpty()) "#" else "# $commentText"
    }

    normalizeStartupFunctionCallLine(trimmedLine)?.let { return it }
    normalizeStartupIncludeLine(trimmedLine)?.let { return it }
    normalizeStartupCommandLine(trimmedLine)?.let { return it }
    return trimmedLine
  }

  private fun isSubstitutionBlockLine(trimmedLine: String): Boolean {
    return Regex("""^(?:file\s+(?:"(?:[^"\\]|\\.)*"|[^\s{]+)|global)\s*\{\s*$""").matches(trimmedLine) ||
      Regex("""^global\s*\{\s*$""").matches(trimmedLine)
  }

  private fun normalizeDatabaseRecordLine(trimmedLine: String): String? {
    val match = Regex(
      """^record\(\s*([A-Za-z0-9_]+)\s*,\s*("(?:(?:[^"\\]|\\.)*)")\s*\)\s*(\{)?\s*(#.*)?$""",
    ).matchEntire(trimmedLine) ?: return null
    return buildString {
      append("record(")
      append(match.groupValues[1])
      append(", ")
      append(match.groupValues[2])
      append(")")
      if (match.groupValues[3].isNotEmpty()) {
        append(" {")
      }
      if (match.groupValues[4].isNotEmpty()) {
        append(' ')
        append(match.groupValues[4].trim())
      }
    }
  }

  private fun normalizeSubstitutionBlockLine(trimmedLine: String): String? {
    val fileMatch = Regex("""^file\s+("(?:(?:[^"\\]|\\.)*)"|[^\s{]+)\s*\{\s*$""")
      .matchEntire(trimmedLine)
    if (fileMatch != null) {
      return "file ${fileMatch.groupValues[1]} {"
    }
    if (Regex("""^global\s*\{\s*$""").matches(trimmedLine)) {
      return "global {"
    }
    return null
  }

  private fun normalizeProtocolBlockLine(trimmedLine: String): String? {
    val match = Regex("""^(@?[A-Za-z_][A-Za-z0-9_-]*)\s*\{\s*$""").matchEntire(trimmedLine) ?: return null
    return "${match.groupValues[1]} {"
  }

  private fun normalizeDatabaseFieldLine(trimmedLine: String): String? {
    val match = Regex(
      """^field\(\s*((?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+))\s*,\s*(.+?)\s*\)\s*(#.*)?$""",
    ).matchEntire(trimmedLine) ?: return null
    return buildString {
      append("field(")
      append(match.groupValues[1])
      append(", ")
      append(match.groupValues[2].trim())
      append(")")
      if (match.groupValues[3].isNotEmpty()) {
        append(' ')
        append(match.groupValues[3].trim())
      }
    }
  }

  private fun normalizeSubstitutionPatternLine(trimmedLine: String): String? {
    val match = Regex("""^pattern\s*\{(.*)\}\s*$""").matchEntire(trimmedLine) ?: return null
    val values = splitSubstitutionCommaSeparatedItems(match.groupValues[1])
    return "pattern { ${values.joinToString(", ") { formatSubstitutionScalarValue(it) }} }"
  }

  private fun normalizeSubstitutionRowLine(
    trimmedLine: String,
    state: SubstitutionFormattingState,
  ): String? {
    val match = Regex("""^\{(.*)\}\s*$""").matchEntire(trimmedLine) ?: return null
    val innerText = match.groupValues[1]
    if (innerText.trim().isEmpty()) {
      return "{ }"
    }

    val assignmentEntries = parseSubstitutionAssignmentEntries(innerText)
    if (assignmentEntries != null) {
      val orderedEntries = orderSubstitutionAssignmentEntries(assignmentEntries, state)
      return "{ " + orderedEntries.joinToString(", ") { (name, value) ->
        "$name=${formatSubstitutionAssignmentValue(value)}"
      } + " }"
    }

    val values = splitSubstitutionCommaSeparatedItems(innerText)
    return "{ " + values.joinToString(", ") { formatSubstitutionScalarValue(it) } + " }"
  }

  private fun normalizeProtocolAssignmentLine(trimmedLine: String): String? {
    val match = Regex("""^([A-Za-z_][A-Za-z0-9_-]*)\s*=\s*(.+?)\s*;\s*$""").matchEntire(trimmedLine) ?: return null
    return "${match.groupValues[1]} = ${match.groupValues[2].trim()};"
  }

  private fun normalizeStartupFunctionCallLine(trimmedLine: String): String? {
    val match = Regex("""^([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)\s*$""").matchEntire(trimmedLine) ?: return null
    val innerText = match.groupValues[2].trim()
    if (innerText.isEmpty()) {
      return "${match.groupValues[1]}()"
    }

    val values = splitSubstitutionCommaSeparatedItems(innerText)
    return "${match.groupValues[1]}(${values.joinToString(", ") { it.trim() }})"
  }

  private fun normalizeStartupIncludeLine(trimmedLine: String): String? {
    val match = Regex("""^<\s*(.+)$""").matchEntire(trimmedLine) ?: return null
    return "< ${match.groupValues[1].trim()}"
  }

  private fun normalizeStartupCommandLine(trimmedLine: String): String? {
    val match = Regex("""^([./A-Za-z_][A-Za-z0-9_./-]*)\s+(.+)$""").matchEntire(trimmedLine) ?: return null
    return "${match.groupValues[1]} ${match.groupValues[2].trim()}"
  }

  private fun normalizeMakefileIncludeLine(trimmedLine: String): String? {
    val match = Regex("""^(-?include)\s+(.+?)\s*$""").matchEntire(trimmedLine) ?: return null
    return "${match.groupValues[1]} ${match.groupValues[2].trim()}"
  }

  private fun normalizeMakefileAssignmentLine(trimmedLine: String): String? {
    val match = Regex("""^([A-Za-z0-9_.$(){}\-/]+)\s*(\+?[:?]?=)\s*(.*?)\s*$""")
      .matchEntire(trimmedLine) ?: return null
    val valueText = match.groupValues[3].trim()
    return if (valueText.isEmpty()) {
      "${match.groupValues[1]} ${match.groupValues[2]}"
    } else {
      "${match.groupValues[1]} ${match.groupValues[2]} $valueText"
    }
  }

  private fun normalizeClosingBraceLine(trimmedLine: String): String? {
    val match = Regex("""^\}(.*)$""").matchEntire(trimmedLine) ?: return null
    return if (match.groupValues[1].isEmpty()) "}" else "} ${match.groupValues[1].trim()}"
  }

  private fun getBraceDeltaOutsideProtocolStrings(text: String): Int {
    var delta = 0
    var inDoubleQuote = false
    var inSingleQuote = false
    var escaped = false

    for (character in text) {
      if (escaped) {
        escaped = false
        continue
      }
      if (inDoubleQuote) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inDoubleQuote = false
        }
        continue
      }
      if (inSingleQuote) {
        when (character) {
          '\\' -> escaped = true
          '\'' -> inSingleQuote = false
        }
        continue
      }
      when (character) {
        '"' -> inDoubleQuote = true
        '\'' -> inSingleQuote = true
        '#' -> break
        '{' -> delta += 1
        '}' -> delta -= 1
      }
    }
    return delta
  }

  private fun getBraceDeltaOutsideStrings(text: String): Int {
    var delta = 0
    var inString = false
    var escaped = false

    for (character in text) {
      if (escaped) {
        escaped = false
        continue
      }
      if (inString) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inString = false
        }
        continue
      }
      when (character) {
        '"' -> inString = true
        '#' -> break
        '{' -> delta += 1
        '}' -> delta -= 1
      }
    }
    return delta
  }

  private fun splitSubstitutionLineComment(text: String): LineCommentParts {
    var inString = false
    var escaped = false

    for (index in text.indices) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }
      if (inString) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inString = false
        }
        continue
      }
      when (character) {
        '"' -> inString = true
        '#' -> return LineCommentParts(
          code = text.substring(0, index).trimEnd(),
          comment = text.substring(index),
        )
      }
    }

    return LineCommentParts(text, "")
  }

  private fun splitProtocolLineComment(text: String): LineCommentParts {
    var inDoubleQuote = false
    var inSingleQuote = false
    var escaped = false

    for (index in text.indices) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }
      if (inDoubleQuote) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inDoubleQuote = false
        }
        continue
      }
      if (inSingleQuote) {
        when (character) {
          '\\' -> escaped = true
          '\'' -> inSingleQuote = false
        }
        continue
      }
      when (character) {
        '"' -> inDoubleQuote = true
        '\'' -> inSingleQuote = true
        '#' -> return LineCommentParts(
          code = text.substring(0, index).trimEnd(),
          comment = text.substring(index),
        )
      }
    }

    return LineCommentParts(text, "")
  }

  private fun splitStartupLineComment(text: String): LineCommentParts {
    var inString = false
    var escaped = false

    for (index in text.indices) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }
      if (inString) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inString = false
        }
        continue
      }
      when (character) {
        '"' -> inString = true
        '#' -> {
          if (text.getOrNull(index + 1) == '!') {
            continue
          }
          return LineCommentParts(
            code = text.substring(0, index).trimEnd(),
            comment = text.substring(index).trim(),
          )
        }
      }
    }

    return LineCommentParts(text, "")
  }

  private fun splitSubstitutionCommaSeparatedItems(text: String): List<String> {
    val items = mutableListOf<String>()
    var segmentStart = 0
    var inString = false
    var escaped = false

    for (index in text.indices) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }
      if (inString) {
        when (character) {
          '\\' -> escaped = true
          '"' -> inString = false
        }
        continue
      }
      when (character) {
        '"' -> inString = true
        ',' -> {
          items += text.substring(segmentStart, index).trim()
          segmentStart = index + 1
        }
      }
    }
    items += text.substring(segmentStart).trim()
    return items
  }

  private fun parseSubstitutionAssignmentEntries(text: String): List<Pair<String, String>>? {
    val items = splitSubstitutionCommaSeparatedItems(text)
    val entries = mutableListOf<Pair<String, String>>()

    for (item in items) {
      val match = Regex("""^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$""").matchEntire(item) ?: return null
      entries += match.groupValues[1] to match.groupValues[2].trim()
    }

    return entries
  }

  private fun orderSubstitutionAssignmentEntries(
    entries: List<Pair<String, String>>,
    state: SubstitutionFormattingState,
  ): List<Pair<String, String>> {
    val assignmentOrder = state.assignmentOrder
    if (assignmentOrder == null) {
      state.assignmentOrder = entries.map { it.first }.toMutableList()
      return entries
    }

    val remainingEntries = entries.mapIndexed { index, entry -> IndexedEntry(entry, index) }.toMutableList()
    val orderedEntries = mutableListOf<Pair<String, String>>()

    for (name in assignmentOrder) {
      var index = 0
      while (index < remainingEntries.size) {
        if (remainingEntries[index].entry.first == name) {
          orderedEntries += remainingEntries[index].entry
          remainingEntries.removeAt(index)
        } else {
          index += 1
        }
      }
    }

    for (remainingEntry in remainingEntries) {
      if (!assignmentOrder.contains(remainingEntry.entry.first)) {
        assignmentOrder += remainingEntry.entry.first
      }
      orderedEntries += remainingEntry.entry
    }

    return orderedEntries
  }

  private fun formatSubstitutionScalarValue(rawValue: String): String {
    val trimmedValue = rawValue.trim()
    if (trimmedValue.isEmpty()) {
      return "\"\""
    }

    val quotedMatch = Regex("""^"((?:[^"\\]|\\.)*)"$""").matchEntire(trimmedValue)
    if (quotedMatch != null) {
      return "\"${quotedMatch.groupValues[1]}\""
    }

    return if (Regex("""[\s,{}#"]""").containsMatchIn(trimmedValue)) {
      "\"${escapeDoubleQuotedString(trimmedValue)}\""
    } else {
      trimmedValue
    }
  }

  private fun formatSubstitutionAssignmentValue(rawValue: String): String {
    val trimmedValue = rawValue.trim()
    return if (trimmedValue.isEmpty()) "" else formatSubstitutionScalarValue(trimmedValue)
  }

  private fun escapeDoubleQuotedString(value: String): String {
    return value.replace("\\", "\\\\").replace("\"", "\\\"")
  }

  private fun expandSequencerLogicalLines(text: String, initialInBlockComment: Boolean): List<String> {
    val splitLines = mutableListOf<String>()
    var inBlockComment = initialInBlockComment
    val compoundLines = splitSequencerCompoundLine(text, inBlockComment)

    for (compoundLine in compoundLines) {
      val logicalLines = splitSequencerTrailingClosingBraces(compoundLine, inBlockComment)
      for (logicalLine in logicalLines) {
        splitLines += logicalLine
        inBlockComment = getSequencerBraceInfo(logicalLine, inBlockComment).inBlockComment
      }
    }
    return splitLines
  }

  private fun splitSequencerCompoundLine(text: String, initialInBlockComment: Boolean): List<String> {
    if (text.isEmpty() || isSequencerColumnZeroLine(text)) {
      return listOf(text)
    }

    val parts = mutableListOf<String>()
    var segmentStart = 0
    var inBlockComment = initialInBlockComment
    var inDoubleQuote = false
    var inSingleQuote = false
    var escaped = false
    var index = 0

    while (index < text.length) {
      val character = text[index]
      val nextCharacter = text.getOrNull(index + 1)

      when {
        inBlockComment -> {
          if (character == '*' && nextCharacter == '/') {
            inBlockComment = false
            index += 1
          }
        }

        inDoubleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '"' -> inDoubleQuote = false
          }
        }

        inSingleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '\'' -> inSingleQuote = false
          }
        }

        character == '/' && nextCharacter == '*' -> {
          inBlockComment = true
          index += 1
        }

        character == '/' && nextCharacter == '/' -> break

        character == '%' && nextCharacter == '{' -> index += 1
        character == '}' && nextCharacter == '%' -> index += 1
        character == '"' -> inDoubleQuote = true
        character == '\'' -> inSingleQuote = true
        character == '{' -> {
          val nextNonWhitespaceIndex = getNextNonWhitespaceIndex(text, index + 1)
          if (nextNonWhitespaceIndex >= 0) {
            val head = text.substring(segmentStart, index + 1).trim()
            if (head.isNotEmpty()) {
              parts += head
            }
            segmentStart = nextNonWhitespaceIndex
          }
        }

        character == '}' -> {
          val currentSegment = text.substring(segmentStart, index).trim()
          val nextNonWhitespaceIndex = getNextNonWhitespaceIndex(text, index + 1)
          if (currentSegment.isNotEmpty() && nextNonWhitespaceIndex >= 0) {
            parts += currentSegment
            segmentStart = index
          }
        }
      }

      index += 1
    }

    val tail = text.substring(segmentStart).trim()
    if (tail.isNotEmpty()) {
      parts += tail
    }

    return if (parts.isEmpty()) listOf(text) else parts
  }

  private fun splitSequencerTrailingClosingBraces(
    text: String,
    initialInBlockComment: Boolean,
  ): List<String> {
    if (text.isEmpty() || isSequencerColumnZeroLine(text)) {
      return listOf(text)
    }

    var inBlockComment = initialInBlockComment
    var inDoubleQuote = false
    var inSingleQuote = false
    var escaped = false
    var firstCodeTokenIndex = -1
    val trailingBracePositions = mutableListOf<Int>()
    var index = 0

    while (index < text.length) {
      val character = text[index]
      val nextCharacter = text.getOrNull(index + 1)

      when {
        inBlockComment -> {
          if (character == '*' && nextCharacter == '/') {
            inBlockComment = false
            index += 1
          }
        }

        inDoubleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '"' -> inDoubleQuote = false
          }
        }

        inSingleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '\'' -> inSingleQuote = false
          }
        }

        character == '/' && nextCharacter == '*' -> {
          inBlockComment = true
          index += 1
        }

        character == '/' && nextCharacter == '/' -> break

        character == '%' && nextCharacter == '{' -> {
          if (firstCodeTokenIndex < 0) {
            firstCodeTokenIndex = index
          }
          trailingBracePositions.clear()
          index += 1
        }

        character == '}' && nextCharacter == '%' -> {
          if (firstCodeTokenIndex < 0) {
            firstCodeTokenIndex = index
          }
          trailingBracePositions.clear()
          index += 1
        }

        character == '"' -> inDoubleQuote = true
        character == '\'' -> inSingleQuote = true
        character.isWhitespace() -> Unit
        else -> {
          if (firstCodeTokenIndex < 0) {
            firstCodeTokenIndex = index
          }
          if (character == '}') {
            trailingBracePositions += index
          } else {
            trailingBracePositions.clear()
          }
        }
      }

      index += 1
    }

    if (trailingBracePositions.isEmpty()) {
      return listOf(text)
    }

    val firstTrailingBraceIndex = trailingBracePositions.first()
    if (firstCodeTokenIndex < 0 || firstCodeTokenIndex >= firstTrailingBraceIndex) {
      return listOf(text)
    }

    val head = text.substring(0, firstTrailingBraceIndex).trimEnd()
    if (head.isEmpty()) {
      return listOf(text)
    }

    return buildList {
      add(head)
      repeat(trailingBracePositions.size) { add("}") }
    }
  }

  private fun getNextNonWhitespaceIndex(text: String, startIndex: Int): Int {
    for (index in startIndex until text.length) {
      if (!text[index].isWhitespace()) {
        return index
      }
    }
    return -1
  }

  private fun isSequencerColumnZeroLine(trimmedLine: String): Boolean {
    return trimmedLine.startsWith("#") || trimmedLine.startsWith("%%")
  }

  private fun getSequencerBraceInfo(text: String, initialInBlockComment: Boolean): SequencerBraceInfo {
    var delta = 0
    var inBlockComment = initialInBlockComment
    var inDoubleQuote = false
    var inSingleQuote = false
    var escaped = false
    var sawCodeToken = false
    var startsWithClosingBrace = false
    var index = 0

    while (index < text.length) {
      val character = text[index]
      val nextCharacter = text.getOrNull(index + 1)

      when {
        inBlockComment -> {
          if (character == '*' && nextCharacter == '/') {
            inBlockComment = false
            index += 1
          }
        }

        inDoubleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '"' -> inDoubleQuote = false
          }
        }

        inSingleQuote -> {
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '\'' -> inSingleQuote = false
          }
        }

        character == '/' && nextCharacter == '*' -> {
          inBlockComment = true
          index += 1
        }

        character == '/' && nextCharacter == '/' -> break
        character == '%' && nextCharacter == '{' -> index += 1
        character == '}' && nextCharacter == '%' -> {
          sawCodeToken = true
          index += 1
        }

        character == '"' -> inDoubleQuote = true
        character == '\'' -> inSingleQuote = true
        character.isWhitespace() -> Unit
        else -> {
          if (!sawCodeToken && character == '}') {
            startsWithClosingBrace = true
          }
          sawCodeToken = true
          if (character == '{') {
            delta += 1
          } else if (character == '}') {
            delta -= 1
          }
        }
      }

      index += 1
    }

    return SequencerBraceInfo(
      delta = delta,
      inBlockComment = inBlockComment,
      startsWithClosingBrace = startsWithClosingBrace,
    )
  }

  private data class LineCommentParts(
    val code: String,
    val comment: String,
  )

  private data class SubstitutionFormattingState(
    var assignmentOrder: MutableList<String>? = null,
  )

  private data class FormattedPatternBlock(
    val lines: List<String>,
    val nextLineIndex: Int,
  )

  private data class PatternRowEntry(
    val values: List<String>,
    val comment: String,
  )

  private data class IndexedEntry(
    val entry: Pair<String, String>,
    val index: Int,
  )

  private data class SequencerBraceInfo(
    val delta: Int,
    val inBlockComment: Boolean,
    val startsWithClosingBrace: Boolean,
  )
}
