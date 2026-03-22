package org.epics.workbench.pvlist

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.widget.openEpicsPvlistWidget
import java.nio.file.Files
import java.nio.file.Path

class OpenInPvlistAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible = project != null && file != null && isSupportedFile(file)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isSupportedFile(file)) {
      return
    }

    val path = Path.of(file.path)
    val text = runCatching { Files.readString(path) }.getOrElse { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to read ${file.name}.", TITLE)
      return
    }

    val result = if (isPvlistFile(file)) {
      EpicsPvlistWidgetSupport.buildFromPvlistText(text, file.name, file.path)
    } else {
      EpicsPvlistWidgetSupport.buildFromDatabaseText(text, file.name, file.path)
    }

    val model = result.model
    if (model == null) {
      Messages.showErrorDialog(project, result.issues.joinToString("\n"), TITLE)
      return
    }

    openEpicsPvlistWidget(project, model)
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun isSupportedFile(file: VirtualFile): Boolean {
    return isPvlistFile(file) || file.extension?.lowercase() in DATABASE_EXTENSIONS
  }

  private fun isPvlistFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() == "pvlist"
  }

  private companion object {
    private const val TITLE = "Open PV List Widget"
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  }
}
