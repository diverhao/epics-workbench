package org.epics.workbench.toc

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.command.WriteCommandAction
import com.intellij.openapi.editor.Document
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiDocumentManager

class UpdateDatabaseTocAction : DumbAwareAction() {
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

    val document = getTargetDocument(file) ?: return
    val psiFile = PsiDocumentManager.getInstance(project).getPsiFile(document) ?: return
    val originalText = document.text
    val updatedText = EpicsDatabaseToc.upsert(originalText, detectEol(originalText))
    if (updatedText == originalText) {
      return
    }

    WriteCommandAction.runWriteCommandAction(project, COMMAND_NAME, null, Runnable {
      document.setText(updatedText)
      PsiDocumentManager.getInstance(project).commitDocument(document)
    }, psiFile)
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun getTargetDocument(file: VirtualFile): Document? {
    return FileDocumentManager.getInstance().getDocument(file)
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in DATABASE_EXTENSIONS
  }

  private fun detectEol(text: String): String {
    return if (text.contains("\r\n")) "\r\n" else "\n"
  }

  companion object {
    private const val COMMAND_NAME = "Update Table of Contents"
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  }
}
