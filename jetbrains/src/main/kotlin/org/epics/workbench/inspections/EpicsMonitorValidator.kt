package org.epics.workbench.inspections

internal object EpicsMonitorValidator {
  fun collectIssues(text: String): List<EpicsDatabaseValueValidator.ValidationIssue> {
    val issues = mutableListOf<EpicsDatabaseValueValidator.ValidationIssue>()
    val macroDefinitions = linkedMapOf<String, String>()
    val recordLines = mutableListOf<RecordLine>()
    var lineStart = 0

    while (lineStart <= text.length) {
      val lineEnd = text.indexOf('\n', lineStart).let { if (it >= 0) it else text.length }
      val rawLineEnd = if (lineEnd > lineStart && text[lineEnd - 1] == '\r') lineEnd - 1 else lineEnd
      val lineText = text.substring(lineStart, rawLineEnd)
      val trimmed = lineText.trim()

      when {
        trimmed.isEmpty() || trimmed.startsWith("#") -> Unit
        parseMacroAssignment(lineText)?.let { assignment ->
          macroDefinitions[assignment.name] = assignment.value
          true
        } == true -> Unit
        else -> {
          collectMultipleChannelIssue(lineText, lineStart)?.let(issues::add)
          recordLines += RecordLine(lineText, lineStart)
        }
      }

      if (lineEnd >= text.length) {
        break
      }
      lineStart = lineEnd + 1
    }

    for (recordLine in recordLines) {
      for (reference in findMacroReferences(recordLine.text, recordLine.startOffset)) {
        val unresolved = collectUndefinedMacros(
          reference.name,
          reference.defaultValue,
          macroDefinitions,
          linkedSetOf(),
        )
        if (unresolved.isEmpty()) {
          continue
        }

        val message = if (unresolved.size == 1) {
          """Undefined monitor macro "${unresolved.first()}"."""
        } else {
          """Undefined monitor macros: ${unresolved.joinToString(", ") { "\"$it\"" }}."""
        }
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = reference.startOffset,
          endOffset = reference.endOffset,
          message = message,
        )
      }
    }

    return issues
  }

  private fun collectMultipleChannelIssue(
    lineText: String,
    lineStartOffset: Int,
  ): EpicsDatabaseValueValidator.ValidationIssue? {
    val trimmed = lineText.trim()
    if (trimmed.isEmpty() || trimmed.startsWith("#")) {
      return null
    }

    val tokens = Regex("""\S+""").findAll(lineText).toList()
    if (tokens.size <= 1) {
      return null
    }

    val first = tokens.first()
    val last = tokens.last()
    return EpicsDatabaseValueValidator.ValidationIssue(
      startOffset = lineStartOffset + first.range.first,
      endOffset = lineStartOffset + last.range.last + 1,
      message = "Monitor lines must contain only one channel name.",
    )
  }

  private fun parseMacroAssignment(lineText: String): MacroAssignment? {
    val match = MACRO_ASSIGNMENT_REGEX.matchEntire(lineText) ?: return null
    return MacroAssignment(
      name = match.groups[1]?.value.orEmpty(),
      value = match.groups[2]?.value.orEmpty(),
    )
  }

  private fun findMacroReferences(
    lineText: String,
    lineStartOffset: Int,
  ): List<MacroReference> {
    return MACRO_REFERENCE_REGEX.findAll(lineText).map { match ->
      MacroReference(
        name = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty(),
        defaultValue = match.groups[2]?.value,
        startOffset = lineStartOffset + match.range.first,
        endOffset = lineStartOffset + match.range.last + 1,
      )
    }.toList()
  }

  private fun collectUndefinedMacros(
    macroName: String,
    defaultValue: String?,
    macroDefinitions: Map<String, String>,
    stack: LinkedHashSet<String>,
  ): Set<String> {
    if (macroName in stack) {
      return emptySet()
    }

    val definition = macroDefinitions[macroName]
    if (definition == null) {
      return if (defaultValue != null) {
        collectUndefinedMacrosInText(defaultValue, macroDefinitions, stack)
      } else {
        setOf(macroName)
      }
    }

    val nextStack = LinkedHashSet(stack)
    nextStack += macroName
    return collectUndefinedMacrosInText(definition, macroDefinitions, nextStack)
  }

  private fun collectUndefinedMacrosInText(
    text: String,
    macroDefinitions: Map<String, String>,
    stack: LinkedHashSet<String>,
  ): Set<String> {
    val unresolved = linkedSetOf<String>()
    for (match in MACRO_REFERENCE_REGEX.findAll(text)) {
      val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
      val defaultValue = match.groups[2]?.value
      unresolved += collectUndefinedMacros(macroName, defaultValue, macroDefinitions, stack)
    }
    return unresolved
  }

  private data class MacroAssignment(
    val name: String,
    val value: String,
  )

  private data class RecordLine(
    val text: String,
    val startOffset: Int,
  )

  private data class MacroReference(
    val name: String,
    val defaultValue: String?,
    val startOffset: Int,
    val endOffset: Int,
  )

  private val MACRO_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$""")
  private val MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
}
