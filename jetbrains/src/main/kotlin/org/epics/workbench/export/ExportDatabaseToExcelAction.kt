package org.epics.workbench.export

import com.intellij.notification.NotificationAction
import com.intellij.notification.NotificationGroupManager
import com.intellij.notification.NotificationType
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.fileChooser.FileChooserFactory
import com.intellij.openapi.fileChooser.FileSaverDescriptor
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.LocalFileSystem
import org.epics.workbench.substitutions.EpicsSubstitutionsExpansionSupport
import java.awt.Desktop
import java.nio.file.Files
import java.nio.file.Path

class ExportDatabaseToExcelAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible =
      event.project != null &&
      file != null &&
      canExportDatabaseFile(file)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!canExportDatabaseFile(file)) {
      return
    }

    exportDatabaseFileToExcel(project, file)
  }

  private fun getTargetFile(event: AnActionEvent) = event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile

  private companion object {
    private const val TITLE = EXPORT_TO_EXCEL_TITLE
  }
}

internal fun canExportDatabaseFile(file: com.intellij.openapi.vfs.VirtualFile?): Boolean {
  return file != null &&
    (file.extension?.lowercase() in DATABASE_EXTENSIONS || EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file))
}

internal fun exportDatabaseFileToExcel(
  project: com.intellij.openapi.project.Project,
  file: com.intellij.openapi.vfs.VirtualFile,
) {
  if (!canExportDatabaseFile(file)) {
    return
  }

  val sourcePath = Path.of(file.path)
  val sourceText = if (EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file)) {
    val expandedResult = EpicsSubstitutionsExpansionSupport.expandToDatabaseText(project, file)
    expandedResult.expandedText ?: run {
      Messages.showErrorDialog(project, expandedResult.issues.joinToString("\n"), EXPORT_TO_EXCEL_TITLE)
      return
    }
  } else {
    runCatching { Files.readString(sourcePath) }.getOrElse { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to read ${file.name}.", EXPORT_TO_EXCEL_TITLE)
      return
    }
  }

  val descriptor = FileSaverDescriptor(EXPORT_TO_EXCEL_TITLE, "Export the database as an Excel workbook.", "xlsx")
  val saver = FileChooserFactory.getInstance().createSaveFileDialog(descriptor, project)
  val defaultName = "${sourcePath.fileName.toString().substringBeforeLast('.')}.xlsx"
  val targetWrapper = saver.save(sourcePath.parent, defaultName) ?: return
  val targetPath = targetWrapper.file.toPath()

  val workbookBytes = EpicsDatabaseExcelExporter.buildWorkbook(sourceText)
  runCatching {
    Files.write(targetPath, workbookBytes)
    LocalFileSystem.getInstance().refreshNioFiles(listOf(targetPath))
  }.onFailure { error ->
    Messages.showErrorDialog(project, error.message ?: "Failed to write ${targetPath.fileName}.", EXPORT_TO_EXCEL_TITLE)
    return
  }

  NotificationGroupManager.getInstance()
    .getNotificationGroup(NOTIFICATION_GROUP_ID)
    .createNotification(
      "Saved to ${targetPath.fileName}.",
      targetPath.toString(),
      NotificationType.INFORMATION,
    )
    .addAction(
      NotificationAction.createSimpleExpiring("Open") {
        ApplicationManager.getApplication().executeOnPooledThread {
          runCatching {
            if (Desktop.isDesktopSupported()) {
              Desktop.getDesktop().open(targetPath.toFile())
            } else {
              throw IllegalStateException("Desktop open is not supported on this platform.")
            }
          }.onFailure { error ->
            ApplicationManager.getApplication().invokeLater {
              Messages.showErrorDialog(
                project,
                error.message ?: "Failed to open ${targetPath.fileName}.",
                EXPORT_TO_EXCEL_TITLE,
              )
            }
          }
        }
      },
    )
    .notify(project)
}

private const val EXPORT_TO_EXCEL_TITLE = "Export Database to Excel"
private const val NOTIFICATION_GROUP_ID = "EPICS Workbench Notifications"
private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
