package org.epics.workbench.probe

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.toc.EpicsDatabaseToc
import org.epics.workbench.widget.openEpicsWidget

class OpenInProbeAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val editor = event.getData(CommonDataKeys.EDITOR)
    val enabled = project != null &&
      file != null &&
      editor != null &&
      (isDatabaseFile(file) || isStartupFile(file))
    event.presentation.isEnabledAndVisible = enabled
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    val editor = event.getData(CommonDataKeys.EDITOR) ?: return
    val target = resolveProbeTarget(project, file, editor.document.text, editor.caretModel.offset)
    openEpicsWidget(project, target?.recordName.orEmpty())
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun resolveProbeTarget(
    project: com.intellij.openapi.project.Project,
    file: VirtualFile,
    text: String,
    offset: Int,
  ): ProbeTarget? {
    return when {
      isDatabaseFile(file) -> resolveDatabaseProbeTarget(project, file, text, offset)
      isStartupFile(file) -> EpicsRecordResolver.resolveRecordDefinition(project, file, offset)?.let { definition ->
        ProbeTarget(definition.recordName)
      }
      else -> null
    }
  }

  private fun resolveDatabaseProbeTarget(
    project: com.intellij.openapi.project.Project,
    file: VirtualFile,
    text: String,
    offset: Int,
  ): ProbeTarget? {
    EpicsDatabaseToc.findRuntimeEntryAtOffset(text, offset)?.let { tocEntry ->
      val macroAssignments = EpicsDatabaseToc.extractRuntimeMacroAssignments(text)
      return ProbeTarget(
        recordName = expandProbeRecordName(tocEntry.recordName, macroAssignments),
      )
    }

    EpicsRecordCompletionSupport.extractRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.nameStart until declaration.nameEnd }
      ?.let { declaration ->
        val macroAssignments = EpicsDatabaseToc.extractRuntimeMacroAssignments(text)
        return ProbeTarget(
          recordName = expandProbeRecordName(declaration.name, macroAssignments),
        )
      }

    EpicsRecordResolver.resolveRecordDefinition(project, file, offset)?.let { definition ->
      return ProbeTarget(definition.recordName)
    }

    return null
  }

  private fun expandProbeRecordName(
    recordName: String,
    macroAssignments: Map<String, EpicsDatabaseToc.MacroAssignment>,
  ): String {
    val resolvedMacros = linkedMapOf<String, String>()
    macroAssignments.forEach { (name, assignment) ->
      if (assignment.hasAssignment) {
        resolvedMacros[name] = assignment.value
      }
    }
    val cache = mutableMapOf<String, String>()
    return expandMacroText(recordName, resolvedMacros, cache, emptyList()).trim()
      .ifEmpty { recordName.trim() }
  }

  private fun expandMacroText(
    text: String,
    macroValues: Map<String, String>,
    cache: MutableMap<String, String>,
    stack: List<String>,
  ): String {
    val source = text
    val builder = StringBuilder()
    var cursor = 0

    for (match in MACRO_REFERENCE_REGEX.findAll(source)) {
      builder.append(source, cursor, match.range.first)
      val originalText = match.value
      val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
      val defaultValue = match.groups[2]?.value
      builder.append(
        resolveMacroValue(
          macroName = macroName,
          defaultValue = defaultValue,
          originalText = originalText,
          macroValues = macroValues,
          cache = cache,
          stack = stack,
        ),
      )
      cursor = match.range.last + 1
    }

    builder.append(source.substring(cursor))
    return builder.toString()
  }

  private fun resolveMacroValue(
    macroName: String,
    defaultValue: String?,
    originalText: String,
    macroValues: Map<String, String>,
    cache: MutableMap<String, String>,
    stack: List<String>,
  ): String {
    val cacheKey = "$macroName\u0000${defaultValue.orEmpty()}\u0000$originalText"
    cache[cacheKey]?.let { return it }

    val resolved = when {
      macroName in stack -> originalText
      macroValues.containsKey(macroName) -> expandMacroText(
        macroValues[macroName].orEmpty(),
        macroValues,
        cache,
        stack + macroName,
      )
      defaultValue != null -> expandMacroText(defaultValue, macroValues, cache, stack)
      else -> originalText
    }

    cache[cacheKey] = resolved
    return resolved
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("db", "vdb", "template")
  }

  private fun isStartupFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("cmd", "iocsh") || file.name == "st.cmd"
  }

  private data class ProbeTarget(
    val recordName: String,
  )

  private companion object {
    private val MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
  }
}
