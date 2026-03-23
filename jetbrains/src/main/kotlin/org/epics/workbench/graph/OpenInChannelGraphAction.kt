package org.epics.workbench.graph

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.probe.EpicsProbeSupport
import org.epics.workbench.toc.EpicsDatabaseToc
import org.epics.workbench.widget.openEpicsChannelGraphWidget

class OpenInChannelGraphAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val editor = event.getData(CommonDataKeys.EDITOR)
    event.presentation.isEnabledAndVisible = project != null &&
      file != null &&
      (isDatabaseFile(file) || (editor != null && isSupportedEditorFile(file)))
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    val editor = event.getData(CommonDataKeys.EDITOR)
    val target = resolveTarget(project, file, editor?.document?.text, editor?.caretModel?.offset)
      ?: GraphTarget(
        sourceFile = file,
        sourceText = if (isDatabaseFile(file)) readText(file).orEmpty() else "",
        seedRecordName = null,
        message = null,
      )
    openEpicsChannelGraphWidget(
      project = project,
      sourceLabel = target.sourceFile.name,
      sourceText = target.sourceText,
      seedRecordName = target.seedRecordName,
      message = target.message,
      sourcePath = target.sourceFile.path,
    )
  }

  private fun resolveTarget(
    project: Project,
    file: VirtualFile,
    editorText: String?,
    offset: Int?,
  ): GraphTarget? {
    return when {
      isDatabaseFile(file) -> resolveDatabaseTarget(project, file, editorText, offset)
      isStartupFile(file) -> resolveResolvedRecordTarget(project, file, offset)
      isPvlistFile(file) -> resolveRecordNameSearchTarget(project, resolvePvlistTarget(editorText.orEmpty(), offset))
      isProbeFile(file) -> resolveRecordNameSearchTarget(project, EpicsProbeSupport.analyzeText(editorText.orEmpty()).recordName)
      else -> null
    }
  }

  private fun resolveDatabaseTarget(
    project: Project,
    file: VirtualFile,
    editorText: String?,
    offset: Int?,
  ): GraphTarget? {
    val sourceText = readText(file) ?: return null
    if (editorText == null || offset == null) {
      return GraphTarget(file, sourceText, null, null)
    }

    EpicsDatabaseToc.findRuntimeEntryAtOffset(editorText, offset)?.let { entry ->
      return GraphTarget(file, sourceText, entry.recordName, null)
    }

    EpicsRecordCompletionSupport.extractRecordDeclarations(editorText)
      .firstOrNull { declaration -> offset in declaration.nameStart until declaration.nameEnd }
      ?.let { declaration ->
        return GraphTarget(file, sourceText, declaration.name, null)
      }

    val resolved = EpicsRecordResolver.resolveRecordDefinition(project, file, offset)
    return if (resolved != null) {
      buildResolvedTarget(resolved)
    } else {
      GraphTarget(file, sourceText, null, null)
    }
  }

  private fun resolveResolvedRecordTarget(
    project: Project,
    file: VirtualFile,
    offset: Int?,
  ): GraphTarget? {
    if (offset == null) {
      return null
    }
    val resolved = EpicsRecordResolver.resolveRecordDefinition(project, file, offset) ?: return null
    return buildResolvedTarget(resolved)
  }

  private fun resolveRecordNameSearchTarget(project: Project, recordName: String?): GraphTarget? {
    val normalized = recordName?.trim().orEmpty()
    if (normalized.isBlank()) {
      return null
    }
    val resolved = findRecordDefinitionInProject(project, normalized) ?: return null
    return buildResolvedTarget(resolved)
  }

  private fun buildResolvedTarget(resolved: org.epics.workbench.navigation.EpicsResolvedRecordDefinition): GraphTarget? {
    val sourceText = readText(resolved.targetFile) ?: return null
    val seedRecordName = EpicsRecordCompletionSupport.extractRecordDeclarations(sourceText)
      .firstOrNull { declaration -> declaration.recordStart == resolved.recordStartOffset }
      ?.name
      ?: resolved.recordName
    return GraphTarget(resolved.targetFile, sourceText, seedRecordName, null)
  }

  private fun findRecordDefinitionInProject(
    project: Project,
    recordName: String,
  ): org.epics.workbench.navigation.EpicsResolvedRecordDefinition? {
    val roots = ProjectRootManager.getInstance(project).contentRoots
    val stack = ArrayDeque<VirtualFile>()
    roots.forEach(stack::addLast)

    while (stack.isNotEmpty()) {
      val current = stack.removeLast()
      if (current.isDirectory) {
        if (current.name in EXCLUDED_DIRECTORIES) {
          continue
        }
        current.children.forEach(stack::addLast)
        continue
      }

      if (!isDatabaseFile(current)) {
        continue
      }

      EpicsRecordResolver.resolveRecordDefinitionInFile(current, recordName)?.let { return it }
    }

    return null
  }

  private fun resolvePvlistTarget(text: String, offset: Int?): String? {
    if (offset == null) {
      return null
    }
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

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun readText(file: VirtualFile): String? {
    return runCatching { String(file.contentsToByteArray(), file.charset) }.getOrNull()
  }

  private fun isSupportedEditorFile(file: VirtualFile): Boolean {
    return isDatabaseFile(file) || isStartupFile(file) || isPvlistFile(file) || isProbeFile(file)
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

  private fun isProbeFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "probe"
  }

  private data class GraphTarget(
    val sourceFile: VirtualFile,
    val sourceText: String,
    val seedRecordName: String?,
    val message: String?,
  )

  private companion object {
    private val MACRO_ASSIGNMENT_REGEX = Regex("""^[A-Za-z_][A-Za-z0-9_]*\s*=.*$""")
    private val EXCLUDED_DIRECTORIES = setOf(".git", "build", "dist", "node_modules", "out")
  }
}
