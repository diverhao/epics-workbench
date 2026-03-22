package org.epics.workbench.export

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import java.nio.file.Path

class ImportDatabaseFromExcelAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible =
      event.project != null && file != null && file.extension.equals("xlsx", ignoreCase = true)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!file.extension.equals("xlsx", ignoreCase = true)) {
      return
    }

    val importedSheets = runCatching {
      EpicsDatabaseExcelImporter.importWorkbook(Path.of(file.path))
    }.getOrElse { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to import ${file.name}.", TITLE)
      return
    }

    if (importedSheets.isEmpty()) {
      Messages.showWarningDialog(project, "No EPICS-style sheets were found in ${file.name}.", TITLE)
      return
    }

    EpicsDatabaseExcelImporter.openImportedSheets(project, importedSheets)
  }

  private fun getTargetFile(event: AnActionEvent) = event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile

  private companion object {
    private const val TITLE = "Import as EPICS Database"
  }
}
