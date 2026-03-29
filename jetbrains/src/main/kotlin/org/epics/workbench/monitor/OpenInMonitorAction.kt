package org.epics.workbench.monitor

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.projectHasEpicsRoot
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.probe.EpicsProbeSupport
import org.epics.workbench.substitutions.EpicsSubstitutionsExpansionSupport
import org.epics.workbench.toc.EpicsDatabaseToc
import org.epics.workbench.widget.EpicsMonitorWidgetVirtualFile

class OpenInMonitorAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible =
      project != null && file != null && (projectHasEpicsRoot(project) || isSupportedFile(file))
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isSupportedFile(file)) {
      FileEditorManager.getInstance(project).openFile(
        EpicsMonitorWidgetVirtualFile(emptyList()),
        true,
        true,
      )
      return
    }
    val editor = event.getData(CommonDataKeys.EDITOR)
    val initialChannels = if (EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file)) {
      val expandedResult = EpicsSubstitutionsExpansionSupport.expandToDatabaseText(project, file)
      val expandedText = expandedResult.expandedText
      if (expandedText == null) {
        Messages.showErrorDialog(project, expandedResult.issues.joinToString("\n"), TITLE)
        return
      }
      EpicsMonitorFileSupport.extractUniqueRecordNames(expandedText)
    } else {
      val channelName = if (editor != null) {
        resolveMonitorTarget(project, file, editor.document.text, editor.caretModel.offset)
      } else {
        null
      }
      channelName?.let(::listOf).orEmpty()
    }
    FileEditorManager.getInstance(project).openFile(
      EpicsMonitorWidgetVirtualFile(initialChannels),
      true,
      true,
    )
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun isSupportedFile(file: VirtualFile): Boolean {
    return isDatabaseFile(file) ||
      isStartupFile(file) ||
      isDbdFile(file) ||
      isProtocolFile(file) ||
      isPvlistFile(file) ||
      isProbeFile(file) ||
      EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file)
  }

  private fun resolveMonitorTarget(
    project: com.intellij.openapi.project.Project,
    file: VirtualFile,
    text: String,
    offset: Int,
  ): String? {
    return when {
      isDatabaseFile(file) -> resolveDatabaseTarget(project, file, text, offset)
      isStartupFile(file) -> EpicsRecordResolver.resolveRecordDefinition(project, file, offset)?.recordName
      isDbdFile(file) -> null
      isProtocolFile(file) -> null
      isPvlistFile(file) -> resolvePvlistTarget(text, offset)
      isProbeFile(file) -> EpicsProbeSupport.analyzeText(text).recordName
      else -> null
    }?.trim()?.takeIf(String::isNotBlank)
  }

  private fun resolveDatabaseTarget(
    project: com.intellij.openapi.project.Project,
    file: VirtualFile,
    text: String,
    offset: Int,
  ): String? {
    EpicsDatabaseToc.findRuntimeEntryAtOffset(text, offset)?.let { tocEntry ->
      val macroAssignments = EpicsDatabaseToc.extractRuntimeMacroAssignments(text)
      return expandMacroText(tocEntry.recordName, macroAssignments)
    }

    EpicsRecordCompletionSupport.extractRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.nameStart until declaration.nameEnd }
      ?.let { declaration ->
        val macroAssignments = EpicsDatabaseToc.extractRuntimeMacroAssignments(text)
        return expandMacroText(declaration.name, macroAssignments)
      }

    return EpicsRecordResolver.resolveRecordDefinition(project, file, offset)?.recordName
  }

  private fun resolvePvlistTarget(text: String, offset: Int): String? {
    val lineStart = text.lastIndexOf('\n', (offset - 1).coerceAtLeast(0)).let { if (it >= 0) it + 1 else 0 }
    val lineEnd = text.indexOf('\n', offset).let { if (it >= 0) it else text.length }
    val trimmed = text.substring(lineStart, lineEnd).trim()
    return when {
      trimmed.isEmpty() -> null
      trimmed.startsWith("#") -> null
      MACRO_ASSIGNMENT_REGEX.matches(trimmed) -> null
      trimmed.any(Char::isWhitespace) -> null
      else -> trimmed
    }
  }

  private fun expandMacroText(
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
    return expandMacroText(recordName, resolvedMacros, cache, emptyList()).trim().ifEmpty { recordName.trim() }
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

  private fun isPvlistFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "pvlist"
  }

  private fun isDbdFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "dbd"
  }

  private fun isProtocolFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "proto"
  }

  private fun isProbeFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "probe"
  }

  private companion object {
    private const val TITLE = "Open PV Monitor Widget"
    private val MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
    private val MACRO_ASSIGNMENT_REGEX = Regex("""^[A-Za-z_][A-Za-z0-9_]*\s*=.*$""")
  }
}
