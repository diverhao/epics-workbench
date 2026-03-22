package org.epics.workbench.export

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.fileChooser.FileChooserFactory
import com.intellij.openapi.fileChooser.FileSaverDescriptor
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.LocalFileSystem
import java.nio.file.Files
import java.nio.file.Path

class ExportDatabaseToExcelAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible = event.project != null && file != null && isDatabaseFile(file.extension)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file.extension)) {
      return
    }

    val sourcePath = Path.of(file.path)
    val sourceText = runCatching { Files.readString(sourcePath) }.getOrElse { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to read ${file.name}.", TITLE)
      return
    }

    val descriptor = FileSaverDescriptor(TITLE, "Export the database as an Excel workbook.", "xlsx")
    val saver = FileChooserFactory.getInstance().createSaveFileDialog(descriptor, project)
    val defaultName = "${sourcePath.fileName.toString().substringBeforeLast('.')}.xlsx"
    val targetWrapper = saver.save(sourcePath.parent, defaultName) ?: return
    val targetPath = targetWrapper.file.toPath()

    val workbookBytes = EpicsDatabaseExcelExporter.buildWorkbook(sourceText)
    runCatching {
      Files.write(targetPath, workbookBytes)
      LocalFileSystem.getInstance().refreshNioFiles(listOf(targetPath))
    }.onFailure { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to write ${targetPath.fileName}.", TITLE)
      return
    }
  }

  private fun getTargetFile(event: AnActionEvent) = event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile

  private fun isDatabaseFile(extension: String?): Boolean = extension?.lowercase() in DATABASE_EXTENSIONS

  private companion object {
    private const val TITLE = "Export to Excel"
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  }
}
