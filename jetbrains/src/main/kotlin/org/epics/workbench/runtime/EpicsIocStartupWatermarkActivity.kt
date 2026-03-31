package org.epics.workbench.runtime

import com.intellij.openapi.editor.Editor
import com.intellij.openapi.editor.EditorFactory
import com.intellij.openapi.editor.colors.EditorFontType
import com.intellij.openapi.editor.event.EditorFactoryEvent
import com.intellij.openapi.editor.event.EditorFactoryListener
import com.intellij.openapi.editor.markup.CustomHighlighterRenderer
import com.intellij.openapi.editor.markup.HighlighterLayer
import com.intellij.openapi.editor.markup.HighlighterTargetArea
import com.intellij.openapi.editor.markup.RangeHighlighter
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.components.service
import com.intellij.openapi.project.Project
import com.intellij.openapi.startup.ProjectActivity
import com.intellij.openapi.util.Key
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.openapi.vfs.VirtualFileManager
import com.intellij.openapi.vfs.newvfs.BulkFileListener
import com.intellij.openapi.vfs.newvfs.events.VFileEvent
import com.intellij.ui.JBColor
import java.awt.Color
import java.awt.Font
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.RenderingHints

class EpicsIocStartupWatermarkActivity : ProjectActivity {
  override suspend fun execute(project: Project) {
    val listener = EpicsIocStartupWatermarkListener(project)
    EditorFactory.getInstance().addEditorFactoryListener(listener, project)
    project.messageBus.connect(project).subscribe(EpicsIocRuntimeStateListener.TOPIC, listener)
    project.messageBus.connect(project).subscribe(
      VirtualFileManager.VFS_CHANGES,
      object : BulkFileListener {
        override fun after(events: MutableList<out VFileEvent>) {
          if (events.isEmpty()) {
            return
          }
          listener.refreshOpenEditors()
        }
      },
    )
    listener.refreshOpenEditors()
  }
}

private class EpicsIocStartupWatermarkListener(
  private val project: Project,
) : EditorFactoryListener, EpicsIocRuntimeStateListener {
  private val runtimeService = project.service<EpicsIocRuntimeService>()

  override fun editorCreated(event: EditorFactoryEvent) {
    if (event.editor.project == project) {
      updateEditor(event.editor)
    }
  }

  override fun editorReleased(event: EditorFactoryEvent) {
    clearRunningWatermark(event.editor)
    clearReadOnlyDatabaseWatermark(event.editor)
  }

  override fun startupStateChanged(startupPath: String, running: Boolean) {
    refreshOpenEditors()
  }

  fun refreshOpenEditors() {
    for (editor in EditorFactory.getInstance().allEditors) {
      if (editor.project == project) {
        updateEditor(editor)
      }
    }
  }

  private fun updateEditor(editor: Editor) {
    val file = FileDocumentManager.getInstance().getFile(editor.document)
    if (file == null) {
      clearRunningWatermark(editor)
      clearReadOnlyDatabaseWatermark(editor)
      return
    }

    updateRunningWatermark(editor, file)
    updateReadOnlyDatabaseWatermark(editor, file)
  }

  private fun updateRunningWatermark(editor: Editor, file: VirtualFile) {
    if (!EpicsIocRuntimeService.isIocBootStartupFile(file) || !runtimeService.isRunning(file)) {
      clearRunningWatermark(editor)
      return
    }

    if (editor.getUserData(RUNNING_WATERMARK_HIGHLIGHTER_KEY)?.isValid == true) {
      return
    }

    val highlighter = editor.markupModel.addRangeHighlighter(
      0,
      editor.document.textLength,
      HighlighterLayer.CARET_ROW + 1,
      null,
      HighlighterTargetArea.LINES_IN_RANGE,
    )
    highlighter.setCustomRenderer(
      RepeatedWatermarkRenderer(
        text = "Running ...",
        lightColor = Color(210, 70, 70, 58),
        darkColor = Color(255, 96, 96, 54),
      ),
    )
    editor.putUserData(RUNNING_WATERMARK_HIGHLIGHTER_KEY, highlighter)
  }

  private fun clearRunningWatermark(editor: Editor) {
    editor.getUserData(RUNNING_WATERMARK_HIGHLIGHTER_KEY)?.let { highlighter ->
      if (highlighter.isValid) {
        editor.markupModel.removeHighlighter(highlighter)
      }
    }
    editor.putUserData(RUNNING_WATERMARK_HIGHLIGHTER_KEY, null)
  }

  private fun updateReadOnlyDatabaseWatermark(editor: Editor, file: VirtualFile) {
    if (!isDatabaseFile(file) || file.isWritable) {
      clearReadOnlyDatabaseWatermark(editor)
      return
    }

    if (editor.getUserData(READ_ONLY_DATABASE_WATERMARK_HIGHLIGHTER_KEY)?.isValid == true) {
      return
    }

    val highlighter = editor.markupModel.addRangeHighlighter(
      0,
      editor.document.textLength,
      HighlighterLayer.CARET_ROW + 1,
      null,
      HighlighterTargetArea.LINES_IN_RANGE,
    )
    highlighter.setCustomRenderer(
      RepeatedWatermarkRenderer(
        text = "Read Only",
        lightColor = Color(120, 130, 150, 16),
        darkColor = Color(160, 170, 190, 14),
      ),
    )
    editor.putUserData(READ_ONLY_DATABASE_WATERMARK_HIGHLIGHTER_KEY, highlighter)
  }

  private fun clearReadOnlyDatabaseWatermark(editor: Editor) {
    editor.getUserData(READ_ONLY_DATABASE_WATERMARK_HIGHLIGHTER_KEY)?.let { highlighter ->
      if (highlighter.isValid) {
        editor.markupModel.removeHighlighter(highlighter)
      }
    }
    editor.putUserData(READ_ONLY_DATABASE_WATERMARK_HIGHLIGHTER_KEY, null)
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("db", "vdb", "template")
  }

  companion object {
    private val RUNNING_WATERMARK_HIGHLIGHTER_KEY = Key.create<RangeHighlighter>(
      "org.epics.workbench.runtime.runningWatermarkHighlighter",
    )
    private val READ_ONLY_DATABASE_WATERMARK_HIGHLIGHTER_KEY = Key.create<RangeHighlighter>(
      "org.epics.workbench.runtime.readOnlyDatabaseWatermarkHighlighter",
    )
  }
}

private class RepeatedWatermarkRenderer(
  private val text: String,
  private val lightColor: Color,
  private val darkColor: Color,
) : CustomHighlighterRenderer {
  override fun paint(editor: Editor, highlighter: RangeHighlighter, graphics: Graphics) {
    if (editor.isDisposed || !highlighter.isValid) {
      return
    }

    val g2 = graphics.create() as Graphics2D
    try {
      g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON)
      g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
      val visibleArea = editor.scrollingModel.visibleArea
      val font = getWatermarkFont(editor)
      g2.font = font
      g2.color = JBColor(lightColor, darkColor)
      val metrics = editor.contentComponent.getFontMetrics(font)
      val textWidth = metrics.stringWidth(text)
      val centerX = visibleArea.x + (visibleArea.width - textWidth) / 2
      var y = visibleArea.y + TOP_PADDING
      while (y < visibleArea.y + visibleArea.height + metrics.height) {
        g2.drawString(text, centerX, y)
        y += ROW_SPACING
      }
    } finally {
      g2.dispose()
    }
  }

  private fun getWatermarkFont(editor: Editor): Font {
    return editor.colorsScheme
      .getFont(EditorFontType.BOLD)
      .deriveFont(Font.BOLD, 42f)
  }

  companion object {
    private const val TOP_PADDING = 110
    private const val ROW_SPACING = 150
  }
}
