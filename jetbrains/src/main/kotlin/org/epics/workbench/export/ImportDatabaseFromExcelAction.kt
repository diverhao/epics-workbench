package org.epics.workbench.export

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.fileChooser.FileChooser
import com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import java.nio.file.Path

class ImportDatabaseFromExcelAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible =
      project != null && (
        isExcelFile(file) ||
          isEpicsFile(file) ||
          isEpicsProject(project)
      )
    event.presentation.text = if (file?.extension?.lowercase() == "pvlist") {
      "Import Excel"
    } else {
      "Import Excel as EPICS DB"
    }
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event)
    if (file == null || !isExcelFile(file)) {
      promptImportWorkbook(project)
      return
    }
    importWorkbook(project, Path.of(file.path))
  }

  private fun getTargetFile(event: AnActionEvent) = event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile

  private fun isExcelFile(file: VirtualFile?): Boolean {
    return file?.extension.equals("xlsx", ignoreCase = true)
  }

  private fun isEpicsFile(file: VirtualFile?): Boolean {
    if (file == null || file.isDirectory) {
      return false
    }
    return file.extension?.lowercase() in EPICS_EXTENSIONS || file.name == "st.cmd"
  }

  private fun isEpicsProject(project: Project): Boolean {
    return ProjectRootManager.getInstance(project).contentRoots.any { root ->
      root.findFileByRelativePath("configure/RELEASE") != null ||
        root.findFileByRelativePath("configure/RELEASE.local") != null
    }
  }

  private fun promptImportWorkbook(project: Project) {
    val descriptor = FileChooserDescriptorFactory
      .createSingleFileDescriptor("xlsx")
      .withTitle("Import Excel as EPICS DB")
    FileChooser.chooseFile(descriptor, project, null) { file ->
      importWorkbook(project, Path.of(file.path))
    }
  }

  private fun importWorkbook(project: Project, path: Path) {
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

  private companion object {
    private const val TITLE = "Import Excel as EPICS DB"
    private val EPICS_EXTENSIONS = setOf(
      "db",
      "vdb",
      "template",
      "substitutions",
      "sub",
      "subs",
      "cmd",
      "iocsh",
      "probe",
      "pvlist",
      "st",
      "proto",
    )
  }
}
