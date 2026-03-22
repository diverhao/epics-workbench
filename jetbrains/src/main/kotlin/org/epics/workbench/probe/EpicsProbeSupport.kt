package org.epics.workbench.probe

import org.epics.workbench.inspections.EpicsDatabaseValueValidator

internal data class EpicsProbeDocumentAnalysis(
  val recordName: String?,
  val overlayOffset: Int?,
  val issues: List<EpicsDatabaseValueValidator.ValidationIssue>,
)

internal object EpicsProbeSupport {
  private val macroRegex = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")

  fun analyzeText(text: String): EpicsProbeDocumentAnalysis {
    val issues = mutableListOf<EpicsDatabaseValueValidator.ValidationIssue>()
    val recordLines = mutableListOf<RecordLine>()
    var lineStart = 0

    while (lineStart <= text.length) {
      val lineEnd = text.indexOf('\n', lineStart).let { if (it >= 0) it else text.length }
      val rawLineEnd = if (lineEnd > lineStart && text[lineEnd - 1] == '\r') lineEnd - 1 else lineEnd
      val lineText = text.substring(lineStart, rawLineEnd)
      val trimmed = lineText.trim()

      if (trimmed.isNotEmpty()) {
        if (!trimmed.startsWith("#")) {
          val startCharacter = lineText.indexOf(trimmed)
          recordLines += RecordLine(
            value = trimmed,
            startOffset = lineStart + startCharacter,
            endOffset = lineStart + startCharacter + trimmed.length,
          )
        }
      }

      if (lineEnd >= text.length) {
        break
      }
      lineStart = lineEnd + 1
    }

    if (recordLines.size > 1) {
      recordLines.forEach { line ->
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = line.startOffset,
          endOffset = line.endOffset,
          message = "Probe files must contain exactly one non-empty record-name line.",
        )
      }
    }

    val recordLine = recordLines.firstOrNull()
    if (recordLine != null) {
      if (macroRegex.containsMatchIn(recordLine.value)) {
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = recordLine.startOffset,
          endOffset = recordLine.endOffset,
          message = "Probe record names cannot contain EPICS macros.",
        )
      }

      if (recordLine.value.any(Char::isWhitespace)) {
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = recordLine.startOffset,
          endOffset = recordLine.endOffset,
          message = "Probe files allow only one record name with no extra text.",
        )
      }
    }

    return EpicsProbeDocumentAnalysis(
      recordName = if (issues.isEmpty()) recordLine?.value else null,
      overlayOffset = recordLine?.endOffset,
      issues = issues,
    )
  }

  fun buildProbeFileText(recordName: String): String {
    val lines = mutableListOf("# EPICS probe file for EPICS Workbench")
    lines += recordName
    return lines.joinToString(separator = "\n", postfix = "\n")
  }

  fun buildProbeFileName(recordName: String): String = "${sanitizeFileNameFragment(recordName)}.probe"

  private fun sanitizeFileNameFragment(value: String): String {
    return value
      .replace(Regex("""[<>:"/\\|?*\u0000-\u001f]+"""), "_")
      .replace(Regex("""\s+"""), "_")
      .replace(Regex("""_+"""), "_")
      .trim('_')
      .ifEmpty { "probe" }
  }

  private data class RecordLine(
    val value: String,
    val startOffset: Int,
    val endOffset: Int,
  )
}
