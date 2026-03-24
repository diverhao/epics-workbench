package org.epics.workbench.folding

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.OpenFileDescriptor
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport

internal class CollapseAllRecordsAction : BaseDatabaseRecordFoldingAction(expand = false)

internal class ExpandAllRecordsAction : BaseDatabaseRecordFoldingAction(expand = true)

internal abstract class BaseDatabaseRecordFoldingAction(
  private val expand: Boolean,
) : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    event.presentation.isEnabledAndVisible = project != null && file != null && isDatabaseFile(file)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!isDatabaseFile(file)) {
      return
    }

    val editor = getTargetEditor(event, project, file) ?: return
    applyRecordFolding(editor, expand)
  }

  private fun getTargetFile(event: AnActionEvent): VirtualFile? {
    return event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
  }

  private fun getTargetEditor(event: AnActionEvent, project: Project, file: VirtualFile): Editor? {
    return event.getData(CommonDataKeys.EDITOR)
      ?: FileEditorManager.getInstance(project).openTextEditor(OpenFileDescriptor(project, file), true)
  }

  private fun applyRecordFolding(editor: Editor, expand: Boolean) {
    val text = editor.document.text
    val ranges = EpicsRecordCompletionSupport.extractRecordDeclarations(text)
      .mapNotNull { declaration ->
        val openingBrace = text.indexOf('{', declaration.recordStart)
        if (openingBrace < 0 || declaration.recordEnd <= openingBrace + 1) {
          return@mapNotNull null
        }
        openingBrace to declaration.recordEnd
      }

    editor.foldingModel.runBatchFoldingOperation {
      ranges.forEach { (startOffset, endOffset) ->
        val region = editor.foldingModel.getFoldRegion(startOffset, endOffset)
          ?: editor.foldingModel.addFoldRegion(startOffset, endOffset, PLACEHOLDER_TEXT)
        region?.isExpanded = expand
      }
    }
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in DATABASE_EXTENSIONS
  }

  private companion object {
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
    private const val PLACEHOLDER_TEXT = "{...}"
  }
}
