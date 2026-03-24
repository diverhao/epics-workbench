package org.epics.workbench.database

import com.intellij.application.options.CodeStyle
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.command.WriteCommandAction
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.Document
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.ide.CopyPasteManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiDocumentManager
import com.intellij.psi.PsiFile
import org.epics.workbench.formatting.EpicsTextFormatter
import org.epics.workbench.monitor.EpicsMonitorFileSupport
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.substitutions.EpicsSubstitutionsExpansionSupport
import org.epics.workbench.toc.EpicsDatabaseToc
import java.awt.datatransfer.StringSelection

class FormatDatabaseFileAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible =
      event.project != null && (
        file?.let(::isDatabaseFile) == true ||
          isDbdFile(file) ||
          isProtocolFile(file) ||
          EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file) ||
          isPvlistFile(file)
      )
    event.presentation.text = if (
      isDbdFile(file) ||
      isProtocolFile(file) ||
      EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file) ||
      isPvlistFile(file)
    ) {
      "Format File"
    } else {
      "Format DB File"
    }
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file) &&
      !isDbdFile(file) &&
      !isProtocolFile(file) &&
      !EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file) &&
      !isPvlistFile(file)
    ) {
      return
    }

    val document = getTargetDocument(file) ?: return
    val psiFile = PsiDocumentManager.getInstance(project).getPsiFile(document) ?: return
    val originalText = document.text
    val formattedText = when {
      isDatabaseFile(file) -> EpicsTextFormatter.formatDatabaseText(originalText, getIndentUnit(psiFile))
      isDbdFile(file) -> EpicsTextFormatter.formatDatabaseText(originalText, getIndentUnit(psiFile))
      isProtocolFile(file) -> EpicsTextFormatter.formatProtocolText(originalText, getIndentUnit(psiFile))
      isPvlistFile(file) -> EpicsTextFormatter.formatMonitorText(originalText)
      EpicsSubstitutionsExpansionSupport.isSubstitutionsFile(file) ->
        EpicsTextFormatter.formatSubstitutionText(originalText, getIndentUnit(psiFile))
      else -> return
    }
    if (formattedText == originalText) {
      return
    }

    WriteCommandAction.runWriteCommandAction(project, COMMAND_NAME, null, Runnable {
      document.setText(formattedText)
      PsiDocumentManager.getInstance(project).commitDocument(document)
    }, psiFile)
  }

  private fun getIndentUnit(file: PsiFile): String {
    val indentOptions = CodeStyle.getIndentOptions(file)
    return if (indentOptions.USE_TAB_CHARACTER) {
      "\t"
    } else {
      " ".repeat(indentOptions.INDENT_SIZE.coerceAtLeast(DEFAULT_INDENT_SIZE))
    }
  }

  private companion object {
    private const val COMMAND_NAME = "Format DB File"
    private const val DEFAULT_INDENT_SIZE = 4
  }
}

class CopyAllRecordNamesAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null && getTargetFile(event)?.let(::isDatabaseFile) == true
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file)) {
      return
    }

    val text = readCurrentText(file) ?: return
    val recordNames = EpicsMonitorFileSupport.extractUniqueRecordNames(text)
    if (recordNames.isEmpty()) {
      Messages.showWarningDialog(project, "No EPICS record names were found in the selected database file.", TITLE)
      return
    }

    CopyPasteManager.getInstance().setContents(
      StringSelection(EpicsMonitorFileSupport.buildRecordNamesClipboardText(recordNames)),
    )
  }

  private companion object {
    private const val TITLE = "Copy All Record Names"
  }
}

class ToggleMonitorChannelsAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null && file != null && isDatabaseFile(file) && hasTableOfContents(file)
    event.presentation.isEnabledAndVisible = visible
    if (project != null) {
      val runtimeService = project.service<EpicsMonitorRuntimeService>()
      event.presentation.text = if (runtimeService.isMonitoringActive()) {
        "Stop Monitor Channels"
      } else {
        "Start Monitor Channels"
      }
    }
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file) || !hasTableOfContents(file)) {
      return
    }

    val runtimeService = project.service<EpicsMonitorRuntimeService>()
    if (runtimeService.isMonitoringActive()) {
      runtimeService.stopMonitoring()
    } else {
      runtimeService.startMonitoring()
    }
  }
}

private fun getTargetFile(event: AnActionEvent): VirtualFile? {
  return event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
}

private fun isDatabaseFile(file: VirtualFile): Boolean {
  return file.extension?.lowercase() in setOf("db", "vdb", "template")
}

private fun isPvlistFile(file: VirtualFile?): Boolean {
  return file?.extension?.lowercase() == "pvlist"
}

private fun isDbdFile(file: VirtualFile?): Boolean {
  return file?.extension?.lowercase() == "dbd"
}

private fun isProtocolFile(file: VirtualFile?): Boolean {
  return file?.extension?.lowercase() == "proto"
}

private fun getTargetDocument(file: VirtualFile): Document? {
  return FileDocumentManager.getInstance().getDocument(file)
}

private fun readCurrentText(file: VirtualFile): String? {
  val document = getTargetDocument(file)
  if (document != null) {
    return document.text
  }
  return runCatching { String(file.contentsToByteArray(), file.charset) }.getOrNull()
}

private fun hasTableOfContents(file: VirtualFile): Boolean {
  val text = readCurrentText(file) ?: return false
  return EpicsDatabaseToc.extractRuntimeEntries(text).isNotEmpty()
}
