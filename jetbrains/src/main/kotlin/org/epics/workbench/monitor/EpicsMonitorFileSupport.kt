package org.epics.workbench.monitor

import org.epics.workbench.completion.EpicsRecordCompletionSupport

internal object EpicsMonitorFileSupport {
  private val macroRegex = Regex("""\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}""")

  fun buildRecordNamesClipboardText(
    recordNames: List<String>,
    eol: String = "\n",
  ): String {
    val lines = mutableListOf<String>()
    val macroNames = extractRecordNameMacroNamesInAppearanceOrder(recordNames)

    if (macroNames.isNotEmpty()) {
      macroNames.forEach { macroName ->
        lines += "$macroName = "
      }
      lines += ""
    }

    lines += recordNames
    return lines.joinToString(eol)
  }

  fun buildMonitorFileText(
    recordNames: List<String>,
    macroNames: List<String>,
    eol: String = "\n",
  ): String {
    val lines = mutableListOf(
      "# this is a pvlist file for EPICS Workbench",
      "# Fill in the macro values below, then open this file and click the EPICS play button in the status bar.",
      "# Each non-comment line after the macro block monitors one EPICS record or PV.",
      "",
    )

    if (macroNames.isNotEmpty()) {
      macroNames.forEach { macroName ->
        lines += "$macroName = "
      }
      lines += ""
    }

    lines += recordNames
    return lines.joinToString(eol)
  }

  fun buildMonitorFileText(
    recordNames: List<String>,
    macroAssignments: Map<String, String>,
    eol: String = "\n",
  ): String {
    val lines = mutableListOf(
      "# this is a pvlist file for EPICS Workbench",
      "# Fill in the macro values below, then open this file and click the EPICS play button in the status bar.",
      "# Each non-comment line after the macro block monitors one EPICS record or PV.",
      "",
    )

    if (macroAssignments.isNotEmpty()) {
      macroAssignments.forEach { (macroName, value) ->
        lines += "$macroName = $value"
      }
      lines += ""
    }

    lines += recordNames
    return lines.joinToString(eol)
  }

  fun extractUniqueRecordNames(text: String): List<String> {
    val names = mutableListOf<String>()
    val seen = linkedSetOf<String>()
    for (declaration in EpicsRecordCompletionSupport.extractRecordDeclarations(text)) {
      val recordName = declaration.name
      if (recordName.isBlank() || !seen.add(recordName)) {
        continue
      }
      names += recordName
    }
    return names
  }

  fun extractRecordNameMacroNames(recordNames: List<String>): List<String> {
    val names = linkedSetOf<String>()
    val text = recordNames.joinToString("\n")
    for (match in macroRegex.findAll(text)) {
      names += match.groups[1]?.value ?: match.groups[2]?.value.orEmpty()
    }
    return names
      .filter(String::isNotBlank)
      .sortedWith(String.CASE_INSENSITIVE_ORDER)
  }

  private fun extractRecordNameMacroNamesInAppearanceOrder(recordNames: List<String>): List<String> {
    val names = linkedSetOf<String>()
    val text = recordNames.joinToString("\n")
    for (match in macroRegex.findAll(text)) {
      names += match.groups[1]?.value ?: match.groups[2]?.value.orEmpty()
    }
    return names.filter(String::isNotBlank)
  }
}
