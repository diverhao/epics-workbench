package org.epics.workbench.documentation

import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.application.ModalityState
import com.intellij.openapi.application.ReadAction
import com.intellij.openapi.command.WriteCommandAction
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.editor.EditorFactory
import com.intellij.openapi.editor.event.EditorMouseEvent
import com.intellij.openapi.editor.event.EditorMouseEventArea
import com.intellij.openapi.editor.event.EditorMouseListener
import com.intellij.openapi.editor.event.EditorMouseMotionListener
import com.intellij.openapi.fileEditor.OpenFileDescriptor
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.startup.ProjectActivity
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.ui.JBColor
import com.intellij.ui.components.JBScrollPane
import com.intellij.util.Alarm
import com.intellij.util.ui.JBUI
import org.epics.workbench.inspections.EpicsDatabaseValidationListener
import java.awt.BorderLayout
import java.awt.Component
import java.awt.Container
import java.awt.Cursor
import java.awt.Dimension
import java.awt.Point
import java.awt.Rectangle
import java.awt.Toolkit
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.net.URLDecoder
import java.net.URI
import java.nio.charset.StandardCharsets
import java.nio.file.Paths
import javax.swing.BorderFactory
import javax.swing.JComponent
import javax.swing.JEditorPane
import javax.swing.JLayeredPane
import javax.swing.JPanel
import javax.swing.JMenuItem
import javax.swing.JPopupMenu
import javax.swing.KeyStroke
import javax.swing.SwingUtilities
import javax.swing.border.EmptyBorder
import javax.swing.event.PopupMenuEvent
import javax.swing.event.PopupMenuListener
import javax.swing.event.HyperlinkEvent
import javax.swing.text.DefaultCaret
import javax.swing.text.DefaultEditorKit
import javax.swing.text.JTextComponent

class EpicsHoverDocumentationActivity : ProjectActivity {
  override suspend fun execute(project: Project) {
    val listener = EpicsHoverDocumentationListener(project)
    val multicaster = EditorFactory.getInstance().eventMulticaster
    multicaster.addEditorMouseMotionListener(listener, project)
    multicaster.addEditorMouseListener(listener, project)
  }
}

private class EpicsHoverDocumentationListener(
  private val project: Project,
) : EditorMouseMotionListener, EditorMouseListener {
  private val hoverAlarm = Alarm(Alarm.ThreadToUse.POOLED_THREAD, project)
  private val closeAlarm = Alarm(Alarm.ThreadToUse.SWING_THREAD, project)
  private var activePopupLayeredPane: JLayeredPane? = null
  private var activePopupContent: JPanel? = null
  private var activeReferenceKey: String? = null
  private var activeEditor: Editor? = null
  private var hoverRequestId: Long = 0
  private var isPointerInsidePopup = false
  private var isContextMenuVisible = false

  override fun mouseMoved(event: EditorMouseEvent) {
    if (event.area != EditorMouseEventArea.EDITING_AREA || !event.isOverText) {
      schedulePopupClose()
      return
    }

    cancelScheduledClose()
    val editor = event.editor
    val hostFile = FileDocumentManager.getInstance().getFile(editor.document)
    if (hostFile == null) {
      schedulePopupClose()
      return
    }
    schedulePopup(editor, hostFile, event.offset, captureAnchor(event.mouseEvent))
  }

  override fun mouseEntered(event: EditorMouseEvent) {
    if (activeEditor != null && activeEditor !== event.editor) {
      hidePopup()
    }
  }

  override fun mouseExited(event: EditorMouseEvent) {
    cancelPendingHover()
    schedulePopupClose()
  }

  override fun mousePressed(event: EditorMouseEvent) {
    hidePopup()
  }

  override fun mouseClicked(event: EditorMouseEvent) {
    hidePopup()
  }

  private fun schedulePopup(
    editor: Editor,
    hostFile: VirtualFile,
    offset: Int,
    anchor: PopupAnchor,
  ) {
    EpicsDatabaseValidationListener.findIssueAt(editor, offset)?.let { issue ->
      val requestId = ++hoverRequestId
      hoverAlarm.cancelAllRequests()
      hoverAlarm.addRequest(
        {
          if (project.isDisposed || requestId != hoverRequestId) {
            return@addRequest
          }

          val validationPreview = createValidationPreview(issue.message, issue.startOffset, issue.endOffset)
          ApplicationManager.getApplication().invokeLater(
            {
              if (project.isDisposed || requestId != hoverRequestId) {
                return@invokeLater
              }
              applyHoverResult(editor, anchor, validationPreview)
            },
            ModalityState.any(),
          )
        },
        HOVER_SHOW_DELAY_MS,
      )
      return
    }

    val requestId = ++hoverRequestId
    hoverAlarm.cancelAllRequests()
    hoverAlarm.addRequest(
      {
        if (project.isDisposed || requestId != hoverRequestId) {
          return@addRequest
        }

        val documentationPreview = ReadAction.compute<EpicsDocumentationPreview?, Throwable> {
          EpicsDocumentationProvider.createDocumentationPreview(project, hostFile, offset)
        }

        ApplicationManager.getApplication().invokeLater(
          {
            if (project.isDisposed || requestId != hoverRequestId) {
              return@invokeLater
            }
            applyHoverResult(editor, anchor, documentationPreview)
          },
          ModalityState.any(),
        )
      },
      HOVER_SHOW_DELAY_MS,
    )
  }

  private fun applyHoverResult(
    editor: Editor,
    anchor: PopupAnchor,
    documentationPreview: EpicsDocumentationPreview?,
  ) {
    if (documentationPreview == null) {
      clearPopup()
      return
    }

    if (documentationPreview.referenceKey == activeReferenceKey && activeEditor === editor) {
      return
    }

    clearPopup()
    showPopup(editor, anchor, documentationPreview)
  }

  private fun showPopup(
    editor: Editor,
    anchor: PopupAnchor,
    documentationPreview: EpicsDocumentationPreview,
  ) {
    val content = createPopupContent(documentationPreview.html)
    val popupContent = JPanel(BorderLayout()).apply {
      background = JBColor.PanelBackground
      border = BorderFactory.createLineBorder(JBColor.GRAY)
      isFocusable = true
      add(content, BorderLayout.CENTER)
    }
    popupContent.size = popupContent.preferredSize

    val layeredPane = SwingUtilities.getRootPane(editor.contentComponent)?.layeredPane ?: return
    installPopupHoverTracking(popupContent)
    val popupLocation = computePopupLocation(popupContent.preferredSize, anchor, layeredPane)
    popupContent.setBounds(
      popupLocation.x,
      popupLocation.y,
      popupContent.preferredSize.width,
      popupContent.preferredSize.height,
    )
    layeredPane.add(popupContent, JLayeredPane.POPUP_LAYER)
    layeredPane.revalidate()
    layeredPane.repaint()

    activePopupLayeredPane = layeredPane
    activePopupContent = popupContent
    activeReferenceKey = documentationPreview.referenceKey
    activeEditor = editor
    isPointerInsidePopup = false
  }

  private fun createPopupContent(html: String): JBScrollPane {
    val editorPane = JEditorPane("text/html", html).apply {
      isEditable = false
      isFocusable = true
      isOpaque = true
      background = JBColor.PanelBackground
      foreground = JBColor.foreground()
      cursor = Cursor.getDefaultCursor()
      putClientProperty(JEditorPane.HONOR_DISPLAY_PROPERTIES, true)
      border = EmptyBorder(
        JBUI.scale(8),
        JBUI.scale(10),
        JBUI.scale(8),
        JBUI.scale(10),
      )
      caretPosition = 0
      addMouseListener(object : MouseAdapter() {
        override fun mousePressed(event: MouseEvent) {
          requestFocus()
          requestFocusInWindow()
        }
      })
      addHyperlinkListener { event ->
        if (event.eventType == HyperlinkEvent.EventType.ACTIVATED) {
          openLinkedFile(event.description ?: event.url?.toExternalForm())
        }
      }
    }
    (editorPane.caret as? DefaultCaret)?.apply {
      isVisible = false
      isSelectionVisible = true
    }
    installTextInteractions(editorPane)

    val preferredWidth = JBUI.scale(520)
    editorPane.setSize(preferredWidth, Int.MAX_VALUE)
    val preferred = editorPane.preferredSize
    val width = preferred.width.coerceIn(JBUI.scale(280), JBUI.scale(560))
    editorPane.setSize(width, Int.MAX_VALUE)
    val adjustedPreferred = editorPane.preferredSize
    val height = adjustedPreferred.height.coerceAtMost(JBUI.scale(420))

    return JBScrollPane(editorPane).apply {
      border = EmptyBorder(0, 0, 0, 0)
      preferredSize = Dimension(width + JBUI.scale(6), height)
    }
  }

  private fun computePopupLocation(
    popupSize: Dimension,
    anchor: PopupAnchor,
    layeredPane: JLayeredPane,
  ): Point {
    val gap = JBUI.scale(6)
    val cursorXOffset = JBUI.scale(6)
    val belowY = anchor.screenPoint.y + gap
    val aboveY = anchor.screenPoint.y - popupSize.height - gap
    val y = if (belowY + popupSize.height <= anchor.screenBounds.y + anchor.screenBounds.height - gap) {
      belowY
    } else {
      aboveY.coerceAtLeast(anchor.screenBounds.y + gap)
    }

    val maxX = anchor.screenBounds.x + anchor.screenBounds.width - popupSize.width - gap
    val minX = anchor.screenBounds.x + gap
    val x = (anchor.screenPoint.x + cursorXOffset).coerceIn(minX, maxX)
    val screenPoint = Point(x, y)
    SwingUtilities.convertPointFromScreen(screenPoint, layeredPane)
    return screenPoint
  }

  private fun captureAnchor(mouseEvent: MouseEvent): PopupAnchor {
    val screenBounds = mouseEvent.component.graphicsConfiguration?.bounds
      ?: Rectangle(0, 0, java.awt.Toolkit.getDefaultToolkit().screenSize.width, java.awt.Toolkit.getDefaultToolkit().screenSize.height)
    return PopupAnchor(mouseEvent.locationOnScreen, screenBounds)
  }

  private fun cancelPendingHover() {
    hoverRequestId += 1
    hoverAlarm.cancelAllRequests()
  }

  private fun cancelScheduledClose() {
    closeAlarm.cancelAllRequests()
  }

  private fun schedulePopupClose() {
    cancelScheduledClose()
    closeAlarm.addRequest(
      {
        if (!isPointerInsidePopup && !isContextMenuVisible) {
          clearPopup()
        }
      },
      POPUP_CLOSE_DELAY_MS,
    )
  }

  private fun hidePopup() {
    cancelPendingHover()
    cancelScheduledClose()
    clearPopup()
  }

  private fun clearPopup() {
    isPointerInsidePopup = false
    isContextMenuVisible = false
    activePopupContent?.let { popupContent ->
      (popupContent.parent as? JComponent)?.remove(popupContent)
      activePopupLayeredPane?.revalidate()
      activePopupLayeredPane?.repaint()
    }
    activePopupLayeredPane = null
    activePopupContent = null
    activeReferenceKey = null
    activeEditor = null
  }

  private fun installPopupHoverTracking(component: Component) {
    component.addMouseListener(object : MouseAdapter() {
      override fun mouseEntered(event: MouseEvent) {
        isPointerInsidePopup = true
        cancelScheduledClose()
      }

      override fun mouseExited(event: MouseEvent) {
        isPointerInsidePopup = false
        schedulePopupClose()
      }
    })

    if (component is Container) {
      component.components.forEach { child ->
        installPopupHoverTracking(child)
      }
    }
  }

  private fun installTextInteractions(textComponent: JTextComponent) {
    val menuShortcutMask = Toolkit.getDefaultToolkit().menuShortcutKeyMaskEx
    textComponent.inputMap.put(
      KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_C, menuShortcutMask),
      DefaultEditorKit.copyAction,
    )
    textComponent.inputMap.put(
      KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_A, menuShortcutMask),
      DefaultEditorKit.selectAllAction,
    )
    textComponent.componentPopupMenu = JPopupMenu().apply {
      addPopupMenuListener(object : PopupMenuListener {
        override fun popupMenuWillBecomeVisible(event: PopupMenuEvent?) {
          isContextMenuVisible = true
          cancelScheduledClose()
        }

        override fun popupMenuWillBecomeInvisible(event: PopupMenuEvent?) {
          isContextMenuVisible = false
          if (!isPointerInsidePopup) {
            schedulePopupClose()
          }
        }

        override fun popupMenuCanceled(event: PopupMenuEvent?) {
          isContextMenuVisible = false
          if (!isPointerInsidePopup) {
            schedulePopupClose()
          }
        }
      })
      add(JMenuItem("Copy").apply {
        accelerator = KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_C, menuShortcutMask)
        addActionListener { textComponent.copy() }
      })
      add(JMenuItem("Select All").apply {
        accelerator = KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_A, menuShortcutMask)
        addActionListener {
          textComponent.requestFocusInWindow()
          textComponent.selectAll()
        }
      })
    }
  }

  private fun createValidationPreview(
    message: String,
    startOffset: Int,
    endOffset: Int,
  ): EpicsDocumentationPreview {
    val escapedMessage = escapeHtml(message)
    return EpicsDocumentationPreview(
      referenceKey = "validation:$startOffset:$endOffset:$message",
      html = """
        <html>
          <body style="margin:0; padding:0; background:#${colorHex(JBColor.PanelBackground)}; color:#${colorHex(JBColor.foreground())};">
            <div style="padding:8px 10px; font-family: sans-serif;">
              <div style="font-size:13px; font-weight:bold; margin-bottom:6px;">EPICS database error</div>
              <div style="font-size:12px; line-height:1.4;">$escapedMessage</div>
            </div>
          </body>
        </html>
      """.trimIndent(),
    )
  }

  private fun escapeHtml(text: String): String {
    return buildString(text.length) {
      text.forEach { character ->
        when (character) {
          '&' -> append("&amp;")
          '<' -> append("&lt;")
          '>' -> append("&gt;")
          '"' -> append("&quot;")
          else -> append(character)
        }
      }
    }
  }

  private fun colorHex(color: java.awt.Color): String {
    return "%02x%02x%02x".format(color.red, color.green, color.blue)
  }

  private fun openLinkedFile(linkTarget: String?) {
    if (linkTarget.isNullOrBlank()) {
      return
    }

    if (openMenuChoiceLink(linkTarget)) {
      hidePopup()
      return
    }

    val destination = resolveLinkDestination(linkTarget) ?: return
    ApplicationManager.getApplication().invokeLater(
      {
        if (!project.isDisposed) {
          if (destination.offset != null && destination.offset >= 0) {
            OpenFileDescriptor(project, destination.file, destination.offset).navigate(true)
          } else {
            FileEditorManager.getInstance(project).openFile(destination.file, true)
          }
        }
      },
      ModalityState.any(),
    )
  }

  private fun openMenuChoiceLink(linkTarget: String): Boolean {
    val menuChoiceLink = parseMenuChoiceLink(linkTarget) ?: return false
    val document = FileDocumentManager.getInstance().getDocument(menuChoiceLink.file) ?: return true

    WriteCommandAction.runWriteCommandAction(project, "Set EPICS Menu Field Value", null, Runnable {
      if (document.isWritable) {
        val safeStart = menuChoiceLink.start.coerceIn(0, document.textLength)
        val safeEnd = menuChoiceLink.end.coerceIn(safeStart, document.textLength)
        document.replaceString(safeStart, safeEnd, menuChoiceLink.value)
      }
    })

    return true
  }

  private fun resolveLinkDestination(linkTarget: String): LinkDestination? {
    return try {
      val uri = URI(linkTarget)
      if (uri.scheme != "file") {
        return null
      }
      val cleanUri = URI(uri.scheme, uri.authority, uri.path, null, null)
      val file = LocalFileSystem.getInstance().findFileByNioFile(Paths.get(cleanUri)) ?: return null
      val offset = uri.fragment
        ?.takeIf { it.startsWith("offset=") }
        ?.substringAfter("offset=")
        ?.toIntOrNull()
      LinkDestination(file, offset)
    } catch (_: Exception) {
      null
    }
  }

  private fun parseMenuChoiceLink(linkTarget: String): MenuChoiceLink? {
    return try {
      val uri = URI(linkTarget)
      if (uri.scheme != "epics-menu" || uri.host != "replace") {
        return null
      }
      val params = parseQueryParameters(uri.rawQuery)
      val fileTarget = params["file"] ?: return null
      val file = resolveLinkDestination(fileTarget)?.file ?: return null
      val start = params["start"]?.toIntOrNull() ?: return null
      val end = params["end"]?.toIntOrNull() ?: return null
      val value = params["value"] ?: return null
      MenuChoiceLink(file, start, end, value)
    } catch (_: Exception) {
      null
    }
  }

  private fun parseQueryParameters(rawQuery: String?): Map<String, String> {
    if (rawQuery.isNullOrBlank()) {
      return emptyMap()
    }
    return rawQuery
      .split('&')
      .mapNotNull { entry ->
        val separator = entry.indexOf('=')
        if (separator <= 0) {
          return@mapNotNull null
        }
        val key = URLDecoder.decode(entry.substring(0, separator), StandardCharsets.UTF_8)
        val value = URLDecoder.decode(entry.substring(separator + 1), StandardCharsets.UTF_8)
        key to value
      }
      .toMap()
  }

}

private data class LinkDestination(
  val file: VirtualFile,
  val offset: Int?,
)

private data class MenuChoiceLink(
  val file: VirtualFile,
  val start: Int,
  val end: Int,
  val value: String,
)

private data class PopupAnchor(
  val screenPoint: Point,
  val screenBounds: Rectangle,
)

private const val HOVER_SHOW_DELAY_MS = 300
private const val POPUP_CLOSE_DELAY_MS = 150
