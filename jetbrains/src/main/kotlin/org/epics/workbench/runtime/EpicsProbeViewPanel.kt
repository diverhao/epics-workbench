package org.epics.workbench.runtime

import com.intellij.openapi.Disposable
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.editor.colors.EditorFontType
import com.intellij.openapi.editor.ex.EditorEx
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.components.JBTextField
import java.awt.BorderLayout
import java.awt.Component
import java.awt.Cursor
import java.awt.Dimension
import java.awt.Font
import java.util.Locale
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import javax.swing.BorderFactory
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JEditorPane
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.Timer
import javax.swing.event.DocumentEvent
import javax.swing.event.DocumentListener

internal class EpicsProbeViewPanel(
  private val stateProvider: () -> EpicsProbeViewState?,
  private val putHandler: (String) -> Unit,
  private val openLinkedProbeHandler: (String) -> Unit,
  private val isMonitoringActive: () -> Boolean,
  private val startHandler: () -> Unit,
  private val stopHandler: () -> Unit,
  private val processHandler: () -> Unit,
  private val showStartStopControls: Boolean = true,
) : JPanel(BorderLayout()), Disposable {
  private val cardPanel = JPanel(BorderLayout())
  private val emptyPanel = createEmptyPanel()
  private val contentPanel = JPanel(BorderLayout())
  private val bodyPanel = JPanel(BorderLayout(0, 8))
  private val toolbarPanel = JPanel(BorderLayout())
  private val metaPanel = JPanel().apply {
    layout = BoxLayout(this, BoxLayout.Y_AXIS)
    alignmentX = Component.LEFT_ALIGNMENT
  }
  private val infoPanel = JPanel(BorderLayout(0, 8))
  private val topPanel = JPanel(BorderLayout(0, 8))
  private val buttonPanel = JPanel().apply {
    layout = java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 8, 0)
    isOpaque = false
  }
  private val startButton = JButton("Start Probe")
  private val stopButton = JButton("Stop Probe")
  private val processButton = JButton("Process")
  private val statusLabel = JBLabel()
  private val titleLabel = JLabel("EPICS Probe").apply {
    font = font.deriveFont(font.style or Font.BOLD, font.size2D)
  }
  private val messageLabel = JLabel().apply {
    border = BorderFactory.createEmptyBorder(0, 0, 8, 0)
  }
  private val valueRowLabel = createInteractiveValueLabel()
  private val recordTypeRowLabel = JLabel()
  private val lastUpdateRowLabel = JLabel()
  private val accessRowLabel = JLabel()
  private val fieldsTitleLabel = JLabel("Fields")
  private val fieldFilterField = JBTextField().apply {
    emptyText.text = "Filter fields"
    toolTipText = "Case-insensitive field-name filter. Text after the first space is ignored."
    maximumSize = Dimension(220, preferredSize.height)
  }
  private val fieldControlsPanel = JPanel().apply {
    layout = BoxLayout(this, BoxLayout.Y_AXIS)
    isOpaque = false
    fieldsTitleLabel.alignmentX = Component.LEFT_ALIGNMENT
    fieldFilterField.alignmentX = Component.LEFT_ALIGNMENT
    add(fieldsTitleLabel)
    add(fieldFilterField)
  }
  private val fieldRowsPanel = JPanel().apply {
    layout = BoxLayout(this, BoxLayout.Y_AXIS)
    alignmentX = Component.LEFT_ALIGNMENT
  }
  private val fieldListPanel = JPanel(BorderLayout(0, 4)).apply {
    isOpaque = false
    add(fieldControlsPanel, BorderLayout.NORTH)
    add(fieldRowsPanel, BorderLayout.CENTER)
  }
  private val fieldScrollPane = JBScrollPane(fieldListPanel)
  private val refreshTimer = Timer(1000) { refreshFromService() }
  private var currentState: EpicsProbeViewState? = null
  private var editorBackground = background
  private var editorForeground = foreground
  private var plainFont = font
  private var boldFont = font.deriveFont(Font.BOLD, font.size2D)

  init {
    border = BorderFactory.createEmptyBorder(4, 4, 4, 4)

    buttonPanel.add(startButton)
    buttonPanel.add(stopButton)
    buttonPanel.add(processButton)
    buttonPanel.add(statusLabel)
    toolbarPanel.isOpaque = false
    toolbarPanel.border = BorderFactory.createEmptyBorder(0, 0, 8, 0)
    toolbarPanel.add(buttonPanel, BorderLayout.WEST)

    startButton.addActionListener { startHandler() }
    stopButton.addActionListener { stopHandler() }
    processButton.addActionListener { processHandler() }

    startButton.isVisible = showStartStopControls
    stopButton.isVisible = showStartStopControls
    fieldFilterField.document.addDocumentListener(
      object : DocumentListener {
        override fun insertUpdate(event: DocumentEvent?) {
          refreshFieldLines(currentState?.fields.orEmpty())
        }

        override fun removeUpdate(event: DocumentEvent?) {
          refreshFieldLines(currentState?.fields.orEmpty())
        }

        override fun changedUpdate(event: DocumentEvent?) {
          refreshFieldLines(currentState?.fields.orEmpty())
        }
      },
    )

    metaPanel.isOpaque = false
    listOf(valueRowLabel, recordTypeRowLabel, lastUpdateRowLabel, accessRowLabel).forEach { label ->
      label.alignmentX = Component.LEFT_ALIGNMENT
      metaPanel.add(label)
    }

    infoPanel.isOpaque = false
    infoPanel.add(titleLabel, BorderLayout.NORTH)
    infoPanel.add(messageLabel, BorderLayout.CENTER)
    infoPanel.add(metaPanel, BorderLayout.SOUTH)

    topPanel.isOpaque = false
    topPanel.add(toolbarPanel, BorderLayout.NORTH)
    topPanel.add(infoPanel, BorderLayout.CENTER)

    bodyPanel.isOpaque = false
    bodyPanel.add(topPanel, BorderLayout.NORTH)
    bodyPanel.add(fieldListPanel, BorderLayout.CENTER)

    fieldScrollPane.setViewportView(bodyPanel)
    contentPanel.add(fieldScrollPane, BorderLayout.CENTER)
    contentPanel.isOpaque = false
    fieldScrollPane.border = BorderFactory.createEmptyBorder()
    fieldScrollPane.viewport.isOpaque = false

    add(cardPanel, BorderLayout.CENTER)
    applyGlobalEditorStyle()

    refreshTimer.initialDelay = 0
    refreshTimer.start()
    refreshFromService()
  }

  override fun dispose() {
    refreshTimer.stop()
  }

  fun updatePreferredSize(editor: EditorEx) {
    applyEditorStyle(editor)
    val visibleArea = editor.scrollingModel.visibleArea
    val targetWidth = (visibleArea.width - 24).coerceAtLeast(520)
    val totalTargetHeight = (visibleArea.height * 0.86).toInt().coerceIn(420, 1200)

    fieldScrollPane.preferredSize = Dimension(targetWidth, totalTargetHeight)
    preferredSize = Dimension(targetWidth, totalTargetHeight)
    revalidate()
  }

  fun refreshFromService() {
    val state = stateProvider()
    currentState = state
    if (state == null) {
      showEmptyPanel()
      return
    }

    titleLabel.text = state.recordName.ifBlank { "EPICS Probe" }
    messageLabel.text = state.message.orEmpty()
    messageLabel.isVisible = !state.message.isNullOrBlank()
    valueRowLabel.text = formatMetaRow("Value", state.value)
    valueRowLabel.cursor = if (state.valueCanPut) Cursor.getPredefinedCursor(Cursor.HAND_CURSOR) else Cursor.getDefaultCursor()
    valueRowLabel.toolTipText = if (state.valueCanPut) "Double-click to change the value" else null
    recordTypeRowLabel.text = formatMetaRow("Record Type", state.recordType)
    lastUpdateRowLabel.text = formatMetaRow("Last Update", state.lastUpdated)
    accessRowLabel.text = formatMetaRow("Permission", state.access)
    refreshFieldLines(state.fields)
    val active = isMonitoringActive()
    startButton.isEnabled = !active
    stopButton.isEnabled = active
    processButton.isEnabled = active && state.recordName.isNotBlank()
    statusLabel.text = if (state.isConnected) "Connected" else "Not connected"
    showContentPanel()
  }

  private fun showEmptyPanel() {
    cardPanel.removeAll()
    cardPanel.add(emptyPanel, BorderLayout.CENTER)
    cardPanel.revalidate()
    cardPanel.repaint()
  }

  private fun showContentPanel() {
    cardPanel.removeAll()
    cardPanel.add(contentPanel, BorderLayout.CENTER)
    cardPanel.revalidate()
    cardPanel.repaint()
  }

  private fun createInteractiveValueLabel(): JLabel {
    return JLabel().apply {
      cursor = Cursor.getDefaultCursor()
      addMouseListener(
        object : MouseAdapter() {
          override fun mouseClicked(event: MouseEvent) {
            if (event.clickCount != 2 || event.button != MouseEvent.BUTTON1) {
              return
            }
            val state = currentState ?: return
            val key = state.valueKey ?: return
            if (state.valueCanPut) {
              putHandler(key)
            }
          }
        },
      )
    }
  }

  private fun createEmptyPanel(): JPanel {
    val text = JEditorPane(
      "text/html",
      """
      <html>
        <body style="font-family: sans-serif; font-size: 12px;">
          <p>Open a <b>.probe</b> file or use <b>Open in Probe</b> from a database or startup file.</p>
          <p>The selected probe file will show a live EPICS record page here while monitoring is running.</p>
        </body>
      </html>
      """.trimIndent(),
    ).apply {
      isEditable = false
      isOpaque = true
      border = BorderFactory.createEmptyBorder()
      cursor = Cursor.getDefaultCursor()
    }
    return JPanel(BorderLayout()).apply {
      isOpaque = true
      add(text, BorderLayout.NORTH)
    }
  }

  private fun refreshFieldLines(fields: List<EpicsProbeFieldViewState>) {
    fieldRowsPanel.removeAll()
    val filteredFields = filterFields(fields)
    if (filteredFields.isEmpty()) {
      fieldRowsPanel.add(
        JLabel(
          if (fields.isNotEmpty() && normalizedFieldFilterTerm().isNotEmpty()) {
            "No fields match filter."
          } else {
            currentState?.message ?: "No fields loaded."
          },
        ).apply {
          alignmentX = Component.LEFT_ALIGNMENT
          applyRowLabelStyle(this, bold = false)
        },
      )
      fieldRowsPanel.revalidate()
      fieldRowsPanel.repaint()
      return
    }

    val fieldColumnWidth = maxOf(
      "Field".length,
      filteredFields.maxOfOrNull { it.fieldName.length } ?: 0,
    ) + 1
    fieldRowsPanel.add(createFieldHeaderRow(fieldColumnWidth))

    filteredFields.forEach { field ->
      fieldRowsPanel.add(createFieldRow(field, fieldColumnWidth))
    }
    fieldRowsPanel.revalidate()
    fieldRowsPanel.repaint()
  }

  private fun normalizedFieldFilterTerm(): String {
    return fieldFilterField.text
      .orEmpty()
      .trim()
      .split(FIELD_FILTER_SPLIT_REGEX, limit = 2)
      .firstOrNull()
      .orEmpty()
      .lowercase(Locale.ROOT)
  }

  private fun filterFields(fields: List<EpicsProbeFieldViewState>): List<EpicsProbeFieldViewState> {
    val filterTerm = normalizedFieldFilterTerm()
    if (filterTerm.isEmpty()) {
      return fields
    }

    return fields.filter { field ->
      field.fieldName.lowercase(Locale.ROOT).contains(filterTerm)
    }
  }

  private fun formatMetaRow(label: String, value: String): String {
    return "${label.padEnd(META_LABEL_WIDTH)}  $value"
  }

  private fun createFieldHeaderRow(fieldColumnWidth: Int): JPanel {
    return JPanel(BorderLayout(FIELD_ROW_GAP, 0)).apply {
      alignmentX = Component.LEFT_ALIGNMENT
      isOpaque = false
      border = BorderFactory.createEmptyBorder(1, 0, 1, 0)
      add(
        createFieldTextLabel("Field", bold = true).apply {
          applyFieldNameColumnWidth(this, fieldColumnWidth)
        },
        BorderLayout.WEST,
      )
      add(createFieldTextLabel("Value", bold = true), BorderLayout.CENTER)
    }
  }

  private fun createFieldRow(field: EpicsProbeFieldViewState, fieldColumnWidth: Int): JPanel {
    val fieldNameLabel =
      createFieldTextLabel(field.fieldName).apply {
        applyFieldNameColumnWidth(this, fieldColumnWidth)
        val probeTargetRecordName = field.probeTargetRecordName
        cursor =
          if (!probeTargetRecordName.isNullOrBlank()) {
            Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
          } else {
            Cursor.getDefaultCursor()
          }
        toolTipText =
          probeTargetRecordName?.let { "Open Probe for $it" }
        if (!probeTargetRecordName.isNullOrBlank()) {
          addMouseListener(
            object : MouseAdapter() {
              override fun mouseClicked(event: MouseEvent) {
                if (event.clickCount != 1 || event.button != MouseEvent.BUTTON1) {
                  return
                }
                openLinkedProbeHandler(probeTargetRecordName)
              }
            },
          )
        }
      }
    val valueLabel =
      createFieldTextLabel(field.value).apply {
        cursor =
          if (field.canPut) Cursor.getPredefinedCursor(Cursor.HAND_CURSOR) else Cursor.getDefaultCursor()
        toolTipText = if (field.canPut) "Double-click to change the value" else null
        addMouseListener(
          object : MouseAdapter() {
            override fun mouseClicked(event: MouseEvent) {
              if (event.clickCount != 2 || event.button != MouseEvent.BUTTON1 || !field.canPut) {
                return
              }
              putHandler(field.key)
            }
          },
        )
      }
    return JPanel(BorderLayout(FIELD_ROW_GAP, 0)).apply {
      alignmentX = Component.LEFT_ALIGNMENT
      isOpaque = false
      border = BorderFactory.createEmptyBorder(1, 0, 1, 0)
      add(fieldNameLabel, BorderLayout.WEST)
      add(valueLabel, BorderLayout.CENTER)
    }
  }

  private fun createFieldTextLabel(text: String, bold: Boolean = false): JLabel {
    return JLabel(text).apply {
      alignmentX = Component.LEFT_ALIGNMENT
      applyRowLabelStyle(this, bold = bold)
    }
  }

  private fun applyFieldNameColumnWidth(label: JLabel, fieldColumnWidth: Int) {
    val targetWidth = label.getFontMetrics(label.font).stringWidth("W".repeat(fieldColumnWidth.coerceAtLeast(1)))
    val targetHeight = label.preferredSize.height
    val size = Dimension(targetWidth, targetHeight)
    label.minimumSize = size
    label.preferredSize = size
    label.maximumSize = Dimension(targetWidth, Int.MAX_VALUE)
  }

  private fun applyRowLabelStyle(label: JLabel, bold: Boolean) {
    label.background = editorBackground
    label.foreground = editorForeground
    label.font = if (bold) boldFont else plainFont
    label.isOpaque = false
  }

  private fun applyEditorStyle(editor: EditorEx) {
    applySchemeStyle(
      schemeBackground = editor.colorsScheme.defaultBackground,
      schemeForeground = editor.colorsScheme.defaultForeground,
      schemeFont = buildMonospaceEditorFont(editor.colorsScheme.editorFontName, editor.colorsScheme.editorFontSize),
    )
  }

  private fun applyGlobalEditorStyle() {
    val scheme = EditorColorsManager.getInstance().globalScheme
    applySchemeStyle(
      schemeBackground = scheme.defaultBackground,
      schemeForeground = scheme.defaultForeground,
      schemeFont = buildMonospaceEditorFont(scheme.editorFontName, scheme.editorFontSize),
    )
  }

  private fun applySchemeStyle(schemeBackground: java.awt.Color, schemeForeground: java.awt.Color, schemeFont: Font) {
    editorBackground = schemeBackground
    editorForeground = schemeForeground
    plainFont = schemeFont
    boldFont = plainFont.deriveFont(Font.BOLD, plainFont.size2D)

    this.background = editorBackground
    this.foreground = editorForeground
    this.font = plainFont
    isOpaque = true

    listOf(cardPanel, buttonPanel, emptyPanel, fieldRowsPanel, fieldListPanel, fieldControlsPanel, bodyPanel).forEach { component ->
      component.background = editorBackground
      component.foreground = editorForeground
      component.font = plainFont
    }

    listOf(contentPanel, toolbarPanel, metaPanel, infoPanel, topPanel).forEach { component ->
      component.background = editorBackground
      component.foreground = editorForeground
      component.font = plainFont
      component.isOpaque = false
    }

    titleLabel.background = editorBackground
    titleLabel.foreground = editorForeground
    titleLabel.font = boldFont

    messageLabel.background = editorBackground
    messageLabel.foreground = editorForeground
    messageLabel.font = plainFont

    statusLabel.background = editorBackground
    statusLabel.foreground = editorForeground
    statusLabel.font = plainFont

    listOf(startButton, stopButton, processButton).forEach { button ->
      button.font = plainFont
    }

    fieldsTitleLabel.background = editorBackground
    fieldsTitleLabel.foreground = editorForeground
    fieldsTitleLabel.font = boldFont

    listOf(valueRowLabel, recordTypeRowLabel, lastUpdateRowLabel, accessRowLabel).forEach { label ->
      applyRowLabelStyle(label, bold = false)
    }

    fieldScrollPane.background = editorBackground
    fieldScrollPane.foreground = editorForeground
    fieldScrollPane.viewport.background = editorBackground

    refreshFieldLines(currentState?.fields.orEmpty())

    revalidate()
    repaint()
  }

  private fun buildMonospaceEditorFont(fontName: String, fontSize: Int): Font {
    val candidate = Font(fontName, Font.PLAIN, fontSize)
    val family = candidate.family.lowercase(Locale.ROOT)
    return if (family.contains("mono") || family == Font.MONOSPACED.lowercase(Locale.ROOT)) {
      candidate
    } else {
      Font(Font.MONOSPACED, Font.PLAIN, fontSize)
    }
  }

  private companion object {
    private const val FIELD_ROW_GAP = 8
    private const val META_LABEL_WIDTH = 12
    private val FIELD_FILTER_SPLIT_REGEX = Regex("\\s+")
  }
}
