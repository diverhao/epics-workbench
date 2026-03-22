package org.epics.workbench.pvlist

import org.epics.workbench.monitor.EpicsMonitorFileSupport
import org.epics.workbench.runtime.MonitorProtocol
import org.epics.workbench.toc.EpicsDatabaseToc
import org.epics.workbench.completion.EpicsRecordCompletionSupport

internal enum class EpicsPvlistWidgetSourceKind {
  DATABASE,
  PVLIST,
}

internal data class EpicsPvlistWidgetModel(
  var sourceLabel: String,
  var sourcePath: String? = null,
  var sourceKind: EpicsPvlistWidgetSourceKind,
  val rawPvNames: MutableList<String>,
  val macroNames: MutableList<String>,
  val macroValues: LinkedHashMap<String, String>,
)

internal data class EpicsPvlistWidgetBuildResult(
  val model: EpicsPvlistWidgetModel?,
  val issues: List<String>,
)

internal data class EpicsPvlistWidgetDefinition(
  val key: String,
  val protocol: MonitorProtocol,
  val pvName: String,
)

internal data class EpicsPvlistWidgetRowPlan(
  val channelName: String,
  val definitionKey: String?,
  val unresolvedValue: String? = null,
)

internal data class EpicsPvlistWidgetPlan(
  val rows: List<EpicsPvlistWidgetRowPlan>,
  val definitions: List<EpicsPvlistWidgetDefinition>,
)

internal object EpicsPvlistWidgetSupport {
  fun buildFromDatabaseText(
    text: String,
    sourceLabel: String,
    sourcePath: String? = null,
  ): EpicsPvlistWidgetBuildResult {
    val rawPvNames = linkedSetOf<String>()
    EpicsRecordCompletionSupport.extractRecordDeclarations(text).forEach { declaration ->
      val name = declaration.name.trim()
      if (name.isNotBlank()) {
        rawPvNames += name
      }
    }

    val macroNames = extractOrderedMacroNames(listOf(text))
    val macroAssignments = LinkedHashMap<String, String>()
    val tocAssignments = EpicsDatabaseToc.extractRuntimeMacroAssignments(text)
    macroNames.forEach { macroName ->
      macroAssignments[macroName] = tocAssignments[macroName]
        ?.takeIf { it.hasAssignment }
        ?.value
        .orEmpty()
    }

    return EpicsPvlistWidgetBuildResult(
      model = EpicsPvlistWidgetModel(
        sourceLabel = sourceLabel,
        sourcePath = sourcePath,
        sourceKind = EpicsPvlistWidgetSourceKind.DATABASE,
        rawPvNames = rawPvNames.toMutableList(),
        macroNames = macroNames.toMutableList(),
        macroValues = macroAssignments,
      ),
      issues = emptyList(),
    )
  }

  fun buildFromPvlistText(
    text: String,
    sourceLabel: String,
    sourcePath: String? = null,
  ): EpicsPvlistWidgetBuildResult {
    val issues = mutableListOf<String>()
    val rawPvNames = mutableListOf<String>()
    val macroNames = mutableListOf<String>()
    val macroValues = linkedMapOf<String, String>()
    val seenMacros = linkedSetOf<String>()
    val seenPvs = linkedSetOf<String>()

    text.split(Regex("\\r?\\n")).forEach { rawLine ->
      val trimmed = rawLine.trim()
      when {
        trimmed.isEmpty() || trimmed.startsWith("#") -> Unit
        MACRO_ASSIGNMENT_REGEX.matchEntire(rawLine) != null -> {
          val match = MACRO_ASSIGNMENT_REGEX.matchEntire(rawLine) ?: return@forEach
          val macroName = match.groups[1]?.value.orEmpty()
          if (macroValues.containsKey(macroName)) {
            issues += """Duplicate pvlist macro "$macroName"."""
          } else {
            macroValues[macroName] = match.groups[2]?.value.orEmpty()
          }
          if (seenMacros.add(macroName)) {
            macroNames.add(macroName)
          }
        }
        trimmed.contains("=") -> {
          issues += """PV list macro definitions must be exactly "NAME = value" with no extra text."""
        }
        trimmed.any(Char::isWhitespace) -> {
          issues += "PV list lines must contain exactly one record name with no extra text."
        }
        else -> {
          if (seenPvs.add(trimmed)) {
            rawPvNames.add(trimmed)
          }
          extractOrderedMacroNames(listOf(trimmed)).forEach { macroName ->
            if (seenMacros.add(macroName)) {
              macroNames.add(macroName)
              macroValues.putIfAbsent(macroName, "")
            }
          }
        }
      }
    }

    return if (issues.isEmpty()) {
      EpicsPvlistWidgetBuildResult(
        model = EpicsPvlistWidgetModel(
          sourceLabel = sourceLabel,
          sourcePath = sourcePath,
          sourceKind = EpicsPvlistWidgetSourceKind.PVLIST,
          rawPvNames = rawPvNames,
          macroNames = macroNames,
          macroValues = LinkedHashMap(macroValues),
        ),
        issues = emptyList(),
      )
    } else {
      EpicsPvlistWidgetBuildResult(model = null, issues = issues)
    }
  }

  fun addChannels(model: EpicsPvlistWidgetModel, text: String): Boolean {
    var changed = false
    val seen = linkedSetOf<String>().apply {
      addAll(model.rawPvNames.map(String::trim).filter(String::isNotBlank))
    }
    parseAddedChannelLines(text).forEach { channelName ->
      if (!seen.add(channelName)) {
        return@forEach
      }
      model.rawPvNames.add(channelName)
      changed = true
      extractOrderedMacroNames(listOf(channelName)).forEach { macroName ->
        if (macroName !in model.macroNames) {
          model.macroNames.add(macroName)
          model.macroValues.putIfAbsent(macroName, "")
        }
      }
    }
    return changed
  }

  fun addMacros(model: EpicsPvlistWidgetModel, text: String): Boolean {
    var changed = false
    parseAddedMacroLines(text).forEach { macroName ->
      if (macroName in model.macroNames) {
        return@forEach
      }
      model.macroNames.add(macroName)
      model.macroValues.putIfAbsent(macroName, "")
      changed = true
    }
    return changed
  }

  fun buildMonitorPlan(
    model: EpicsPvlistWidgetModel,
    defaultProtocol: MonitorProtocol,
  ): EpicsPvlistWidgetPlan {
    val macroDefinitions = linkedMapOf<String, String>()
    model.macroNames.forEach { macroName ->
      val value = model.macroValues[macroName].orEmpty()
      if (value.isNotBlank()) {
        macroDefinitions[macroName] = value
      }
    }

    val rows = mutableListOf<EpicsPvlistWidgetRowPlan>()
    val definitions = mutableListOf<EpicsPvlistWidgetDefinition>()
    val seenDefinitionKeys = linkedSetOf<String>()

    model.rawPvNames.forEach { rawPvName ->
      val expanded = expandMonitorValue(rawPvName, macroDefinitions, linkedSetOf())
      if (expanded.isNullOrBlank() || expanded.any(Char::isWhitespace)) {
        rows += EpicsPvlistWidgetRowPlan(
          channelName = rawPvName,
          definitionKey = null,
          unresolvedValue = "(set macros)",
        )
        return@forEach
      }

      val (protocol, pvName) = splitMonitorProtocol(expanded, defaultProtocol)
      val definitionKey = buildDefinitionKey(protocol, pvName)
      rows += EpicsPvlistWidgetRowPlan(
        channelName = pvName,
        definitionKey = definitionKey,
      )
      if (seenDefinitionKeys.add(definitionKey)) {
        definitions += EpicsPvlistWidgetDefinition(
          key = definitionKey,
          protocol = protocol,
          pvName = pvName,
        )
      }
    }

    return EpicsPvlistWidgetPlan(rows = rows, definitions = definitions)
  }

  fun buildFileText(model: EpicsPvlistWidgetModel, eol: String = "\n"): String {
    val macroAssignments = linkedMapOf<String, String>()
    model.macroNames.forEach { macroName ->
      macroAssignments[macroName] = model.macroValues[macroName].orEmpty()
    }
    return EpicsMonitorFileSupport.buildMonitorFileText(
      recordNames = model.rawPvNames,
      macroAssignments = macroAssignments,
      eol = eol,
    )
  }

  private fun parseAddedChannelLines(text: String): List<String> {
    val results = mutableListOf<String>()
    val seen = linkedSetOf<String>()
    text.split(Regex("\\r?\\n")).forEach { rawLine ->
      val trimmed = rawLine.trim()
      if (trimmed.isEmpty() || trimmed.startsWith("#") || trimmed.any(Char::isWhitespace)) {
        return@forEach
      }
      if (seen.add(trimmed)) {
        results.add(trimmed)
      }
    }
    return results
  }

  private fun parseAddedMacroLines(text: String): List<String> {
    val results = mutableListOf<String>()
    val seen = linkedSetOf<String>()
    text.split(Regex("\\r?\\n")).forEach { rawLine ->
      val trimmed = rawLine.trim()
      if (trimmed.isEmpty() || trimmed.startsWith("#") || !MACRO_NAME_REGEX.matches(trimmed)) {
        return@forEach
      }
      if (seen.add(trimmed)) {
        results.add(trimmed)
      }
    }
    return results
  }

  private fun extractOrderedMacroNames(texts: List<String>): List<String> {
    val names = mutableListOf<String>()
    val seen = linkedSetOf<String>()
    texts.forEach { text ->
      MACRO_REFERENCE_REGEX.findAll(text).forEach { match ->
        val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
        if (macroName.isNotBlank() && seen.add(macroName)) {
          names.add(macroName)
        }
      }
    }
    return names
  }

  private fun expandMonitorValue(
    text: String,
    macroDefinitions: Map<String, String>,
    stack: LinkedHashSet<String>,
  ): String? {
    var unresolved = false
    val expanded = MACRO_REFERENCE_REGEX.replace(text) { match ->
      val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
      val defaultValue = match.groups[2]?.value
      val resolved = resolveMonitorMacro(macroName, defaultValue, macroDefinitions, stack)
      if (resolved == null) {
        unresolved = true
        ""
      } else {
        resolved
      }
    }
    return if (unresolved) null else expanded
  }

  private fun resolveMonitorMacro(
    macroName: String,
    defaultValue: String?,
    macroDefinitions: Map<String, String>,
    stack: LinkedHashSet<String>,
  ): String? {
    if (macroName in stack) {
      return null
    }
    val value = macroDefinitions[macroName] ?: return defaultValue
    val nextStack = LinkedHashSet(stack)
    nextStack += macroName
    return expandMonitorValue(value, macroDefinitions, nextStack)
  }

  private fun splitMonitorProtocol(
    value: String,
    defaultProtocol: MonitorProtocol,
  ): Pair<MonitorProtocol, String> {
    return when {
      value.startsWith("pva://", ignoreCase = true) ->
        MonitorProtocol.PVA to value.removePrefix("pva://").removePrefix("PVA://")
      value.startsWith("ca://", ignoreCase = true) ->
        MonitorProtocol.CA to value.removePrefix("ca://").removePrefix("CA://")
      else -> defaultProtocol to value
    }
  }

  private fun buildDefinitionKey(protocol: MonitorProtocol, pvName: String): String {
    return "${protocol.name.lowercase()}:$pvName"
  }

  private val MACRO_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$""")
  private val MACRO_NAME_REGEX = Regex("""^[A-Za-z_][A-Za-z0-9_]*$""")
  private val MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
}
