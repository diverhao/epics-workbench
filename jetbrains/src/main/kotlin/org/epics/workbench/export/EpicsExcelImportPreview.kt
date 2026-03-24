package org.epics.workbench.export

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import java.nio.file.Path

class OpenExcelImportPreviewAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    promptImportEpicsExcelWorkbook(project)
  }

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null
  }
}

internal fun promptImportEpicsExcelWorkbook(project: Project) {
  val descriptor = com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
    .createSingleFileDescriptor("xlsx")
    .withTitle("Import Excel as EPICS DB")
  com.intellij.openapi.fileChooser.FileChooser.chooseFile(descriptor, project, null) { file ->
    importWorkbookPath(project, Path.of(file.path))
  }
}

internal fun importWorkbookPath(project: Project, path: Path) {
  val importedSheets = runCatching {
    EpicsDatabaseExcelImporter.importWorkbook(path)
  }.getOrElse { error ->
    Messages.showErrorDialog(
      project,
      error.message ?: "Failed to import ${path.fileName}.",
      TITLE,
    )
    return
  }

  if (importedSheets.isEmpty()) {
    Messages.showWarningDialog(
      project,
      "No EPICS-style sheets were found in ${path.fileName}.",
      TITLE,
    )
    return
  }

  EpicsDatabaseExcelImporter.openImportedSheets(project, importedSheets)
}

private const val TITLE = "Import Excel as EPICS DB"
