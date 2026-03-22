package org.epics.workbench.monitor

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.ide.CopyPasteManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import java.awt.datatransfer.StringSelection

class CopyDatabaseAsMonitorFileAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val enabled = project != null && file != null && isDatabaseFile(file)
    event.presentation.isEnabledAndVisible = enabled
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file)) {
      return
    }

    val text = getTargetText(file) ?: return
    val recordNames = EpicsMonitorFileSupport.extractUniqueRecordNames(text)
    if (recordNames.isEmpty()) {
      Messages.showWarningDialog(project, "No EPICS record names were found in the selected database file.", TITLE)
      return
    }

    val macroNames = EpicsMonitorFileSupport.extractRecordNameMacroNames(recordNames)
    val monitorText = EpicsMonitorFileSupport.buildMonitorFileText(
      recordNames = recordNames,
      macroNames = macroNames,
      eol = detectEol(text),
    )
    CopyPasteManager.getInstance().setContents(StringSelection(monitorText))

    val recordLabel = "${recordNames.size} record${if (recordNames.size == 1) "" else "s"}"
    val macroLabel = "${macroNames.size} macro${if (macroNames.size == 1) "" else "s"}"
    Messages.showInfoMessage(project, "Copied $recordLabel and $macroLabel as a .pvlist file.", TITLE)
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun getTargetText(file: VirtualFile): String? {
    val document = FileDocumentManager.getInstance().getDocument(file)
    if (document != null) {
      return document.text
    }
    return runCatching { String(file.contentsToByteArray(), file.charset) }.getOrNull()
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in DATABASE_EXTENSIONS
  }

  private fun detectEol(text: String): String {
    return if (text.contains("\r\n")) "\r\n" else "\n"
  }

  private companion object {
    private const val TITLE = "Copy as PV List File"
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  }
}
