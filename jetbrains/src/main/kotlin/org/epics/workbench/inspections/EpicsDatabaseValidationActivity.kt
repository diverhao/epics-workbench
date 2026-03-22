package org.epics.workbench.inspections

import com.intellij.openapi.editor.Document
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.editor.EditorFactory
import com.intellij.openapi.editor.colors.CodeInsightColors
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.editor.event.DocumentEvent
import com.intellij.openapi.editor.event.DocumentListener
import com.intellij.openapi.editor.event.EditorFactoryEvent
import com.intellij.openapi.editor.event.EditorFactoryListener
import com.intellij.openapi.editor.markup.HighlighterLayer
import com.intellij.openapi.editor.markup.HighlighterTargetArea
import com.intellij.openapi.editor.markup.RangeHighlighter
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.startup.ProjectActivity
import com.intellij.openapi.util.Key

class EpicsDatabaseValidationActivity : ProjectActivity {
  override suspend fun execute(project: Project) {
    val listener = EpicsDatabaseValidationListener(project)
    val multicaster = EditorFactory.getInstance().eventMulticaster
    multicaster.addDocumentListener(listener, project)
    EditorFactory.getInstance().addEditorFactoryListener(listener, project)
    listener.refreshOpenEditors()
  }
}

internal class EpicsDatabaseValidationListener(
  private val project: Project,
) : DocumentListener, EditorFactoryListener {
  override fun documentChanged(event: DocumentEvent) {
    updateDocument(event.document)
  }

  override fun editorCreated(event: EditorFactoryEvent) {
    updateDocument(event.editor.document)
  }

  fun refreshOpenEditors() {
    for (editor in EditorFactory.getInstance().allEditors) {
      if (editor.project == project) {
        updateDocument(editor.document)
      }
    }
  }

  private fun updateDocument(document: Document) {
    val file = FileDocumentManager.getInstance().getFile(document)
    val issues = collectIssues(file, document)

    for (editor in EditorFactory.getInstance().getEditors(document, project)) {
      if (file == null || !isSupportedFile(file.name)) {
        clearHighlighters(editor)
      } else {
        applyHighlighters(editor, issues)
      }
    }
  }

  private fun collectIssues(
    file: com.intellij.openapi.vfs.VirtualFile?,
    document: Document,
  ): List<EpicsDatabaseValueValidator.ValidationIssue> {
    if (file == null) {
      return emptyList()
    }

    return when {
      isDatabaseFile(file.name) -> EpicsDatabaseValueValidator.collectIssues(document.text)
      isStartupFile(file.name) -> EpicsStartupMacroValidator.collectIssues(project, file, document.text)
      isMonitorFile(file.name) -> EpicsMonitorValidator.collectIssues(document.text)
      isProbeFile(file.name) -> EpicsProbeValidator.collectIssues(document.text)
      else -> emptyList()
    }
  }

  private fun applyHighlighters(
    editor: Editor,
    issues: List<EpicsDatabaseValueValidator.ValidationIssue>,
  ) {
    clearHighlighters(editor)

    val errorAttributes = EditorColorsManager.getInstance()
      .globalScheme
      .getAttributes(CodeInsightColors.ERRORS_ATTRIBUTES)

    val highlighters = mutableListOf<RangeHighlighter>()
    for (issue in issues) {
      val startOffset = issue.startOffset.coerceIn(0, editor.document.textLength)
      val endOffset = issue.endOffset.coerceAtLeast(startOffset).coerceIn(0, editor.document.textLength)
      val rangeEnd = if (endOffset > startOffset) endOffset else (startOffset + 1).coerceAtMost(editor.document.textLength)
      if (rangeEnd < startOffset) {
        continue
      }

      val highlighter = editor.markupModel.addRangeHighlighter(
        startOffset,
        rangeEnd,
        HighlighterLayer.ERROR,
        errorAttributes,
        HighlighterTargetArea.EXACT_RANGE,
      )
      highlighter.errorStripeTooltip = issue.message
      highlighters += highlighter
    }

    editor.putUserData(ISSUES_KEY, issues)
    editor.putUserData(HIGHLIGHTERS_KEY, highlighters)
  }

  private fun clearHighlighters(editor: Editor) {
    editor.getUserData(HIGHLIGHTERS_KEY)?.forEach(RangeHighlighter::dispose)
    editor.putUserData(ISSUES_KEY, emptyList())
    editor.putUserData(HIGHLIGHTERS_KEY, mutableListOf())
  }

  private fun isDatabaseFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("db", "vdb", "template")
  }

  private fun isStartupFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("cmd", "iocsh") || fileName == "st.cmd"
  }

  private fun isSupportedFile(fileName: String): Boolean {
    return isDatabaseFile(fileName) || isStartupFile(fileName) || isMonitorFile(fileName) || isProbeFile(fileName)
  }

  private fun isMonitorFile(fileName: String): Boolean {
    return fileName.substringAfterLast('.', "").lowercase() == "pvlist"
  }

  private fun isProbeFile(fileName: String): Boolean {
    return fileName.substringAfterLast('.', "").lowercase() == "probe"
  }

  companion object {
    private val HIGHLIGHTERS_KEY =
      Key.create<MutableList<RangeHighlighter>>("org.epics.workbench.inspections.databaseValueHighlighters")
    private val ISSUES_KEY =
      Key.create<List<EpicsDatabaseValueValidator.ValidationIssue>>("org.epics.workbench.inspections.databaseValidationIssues")

    internal fun findIssueAt(editor: Editor, offset: Int): EpicsDatabaseValueValidator.ValidationIssue? {
      return editor.getUserData(ISSUES_KEY)
        ?.firstOrNull { issue ->
          val rangeEnd = if (issue.endOffset > issue.startOffset) issue.endOffset else issue.startOffset + 1
          offset in issue.startOffset until rangeEnd
        }
    }
  }
}
