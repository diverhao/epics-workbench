package org.epics.workbench.pvlist

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.projectHasEpicsRoot
import org.epics.workbench.substitutions.EpicsSubstitutionsExpansionSupport
import org.epics.workbench.widget.EpicsPvlistWidgetVirtualFile
import org.epics.workbench.widget.openEpicsPvlistWidget
import java.nio.file.Files
import java.nio.file.Path

class OpenInPvlistAction : DumbAwareAction() {
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
      openEpicsPvlistWidget(
        project,
        EpicsPvlistWidgetModel(
          sourceLabel = file.name,
          sourcePath = null,
          sourceKind = EpicsPvlistWidgetSourceKind.PVLIST,
          rawPvNames = mutableListOf(),
          macroNames = mutableListOf(),
          macroValues = linkedMapOf(),
        ),
      )
      return
    }

    val path = Path.of(file.path)
    val text = readSourceText(file) ?: run {
      val errorMessage = runCatching { Files.readString(path) }
        .exceptionOrNull()
        ?.message
        ?: "Failed to read ${file.name}."
      Messages.showErrorDialog(project, errorMessage, TITLE)
      return
    }

    val result = when {
      isPvlistFile(file) -> EpicsPvlistWidgetSupport.buildFromPvlistText(text, file.name, file.path)
      isDbdFile(file) || isProtocolFile(file) -> EpicsPvlistWidgetBuildResult(
        model = EpicsPvlistWidgetModel(
          sourceLabel = file.name,
          sourcePath = file.path,
          sourceKind = EpicsPvlistWidgetSourceKind.PVLIST,
          rawPvNames = mutableListOf(),
          macroNames = mutableListOf(),
          macroValues = linkedMapOf(),
        ),
        issues = emptyList(),
      )
      EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file) -> {
        val expandedResult = EpicsSubstitutionsExpansionSupport.expandToDatabaseText(project, file)
        val expandedText = expandedResult.expandedText
        if (expandedText == null) {
          Messages.showErrorDialog(project, expandedResult.issues.joinToString("\n"), TITLE)
          return
        }
        EpicsPvlistWidgetSupport.buildFromDatabaseText(expandedText, file.name, file.path)
      }
      else -> EpicsPvlistWidgetSupport.buildFromDatabaseText(text, file.name, file.path)
    }

    val model = result.model
    if (model == null) {
      Messages.showErrorDialog(project, result.issues.joinToString("\n"), TITLE)
      return
    }

    FileEditorManager.getInstance(project).openFile(EpicsPvlistWidgetVirtualFile(model), true, true)
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun readSourceText(file: VirtualFile): String? {
    FileDocumentManager.getInstance().getCachedDocument(file)?.let { document ->
      return document.text
    }
    return runCatching { Files.readString(Path.of(file.path)) }.getOrNull()
  }

  private fun isSupportedFile(file: VirtualFile): Boolean {
    return isPvlistFile(file) ||
      isDbdFile(file) ||
      isProtocolFile(file) ||
      file.extension?.lowercase() in DATABASE_EXTENSIONS ||
      EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file)
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

  private companion object {
    private const val TITLE = "Open PV List Widget"
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  }
}
