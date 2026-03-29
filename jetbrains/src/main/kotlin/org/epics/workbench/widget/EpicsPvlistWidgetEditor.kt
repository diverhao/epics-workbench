package org.epics.workbench.widget

import com.intellij.openapi.Disposable
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.fileEditor.FileEditor
import com.intellij.openapi.fileEditor.FileEditorLocation
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorPolicy
import com.intellij.openapi.fileEditor.FileEditorProvider
import com.intellij.openapi.fileEditor.FileEditorState
import com.intellij.openapi.fileEditor.FileEditorStateLevel
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import com.intellij.ui.ColorUtil
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.components.JBTextField
import com.intellij.ui.table.JBTable
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSupport
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsPvlistWidgetRowViewState
import java.awt.BorderLayout
import java.awt.CardLayout
import java.awt.Color
import java.awt.Component
import java.awt.Cursor
import java.awt.Dimension
import java.awt.FlowLayout
import java.awt.Font
import java.awt.Graphics
import java.beans.PropertyChangeListener
import java.util.LinkedHashMap
import java.util.Locale
import java.util.UUID
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.awt.event.MouseMotionAdapter
import javax.swing.AbstractCellEditor
import javax.swing.BorderFactory
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JComponent
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.JTable
import javax.swing.JTextArea
import javax.swing.JTextField
import javax.swing.Timer
import javax.swing.table.DefaultTableCellRenderer
import javax.swing.table.AbstractTableModel
import javax.swing.table.TableCellEditor
import javax.swing.event.DocumentEvent
import javax.swing.event.DocumentListener

internal class EpicsPvlistWidgetVirtualFile(
  initialModel: EpicsPvlistWidgetModel,
) : LightVirtualFile(TAB_TITLE) {
  val widgetId: String = UUID.randomUUID().toString()
  val model: EpicsPvlistWidgetModel = initialModel

  companion object {
    const val TAB_TITLE: String = "EPICS PvList"
  }
}

internal fun openEpicsPvlistWidget(project: Project, model: EpicsPvlistWidgetModel) {
  FileEditorManager.getInstance(project).openFile(EpicsPvlistWidgetVirtualFile(model), true, true)
}

class EpicsPvlistWidgetFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: VirtualFile): Boolean {
    return file is EpicsPvlistWidgetVirtualFile
  }

  override fun createEditor(project: Project, file: VirtualFile): FileEditor {
    return EpicsPvlistWidgetFileEditor(project, file as EpicsPvlistWidgetVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-pvlist-widget-editor"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsPvlistWidgetFileEditor(
  private val project: Project,
  private val file: EpicsPvlistWidgetVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val runtimeService = project.service<EpicsMonitorRuntimeService>()
  private val cardLayout = CardLayout()
  private val cardPanel = JPanel(cardLayout)
  private val tableModel = PvlistTableModel()
  private val channelTable = JBTable(tableModel)
  private val channelEditor = PvlistValueCellEditor()
  private val configureChannelsButton = JButton("Configure Channels")
  private val sourceLabel = JBLabel()
  private val messageLabel = JBLabel()
  private val macroFields = LinkedHashMap<String, JTextField>()
  private val macrosTitleLabel = JLabel("Macros")
  private val channelsTitleLabel = JLabel("Channels")
  private val channelFilterField = JBTextField().apply {
    emptyText.text = "Filter records"
    toolTipText = "Case-insensitive record-name filter. Spaces mean AND."
    maximumSize = Dimension(320, preferredSize.height)
  }
  private val overlayTitleLabel = JLabel("Configure Channels")
  private val addChannelsArea = PromptTextArea("One channel per line", 16, 28)
  private val addChannelsButton = JButton("OK")
  private val refreshTimer = Timer(1000) { refreshViewState() }
  private val component = JPanel(BorderLayout(0, 12))
  private val macrosFieldsContainer = JPanel()

  init {
    runtimeService.initialize()
    runtimeService.startMonitoring()
    buildUi()
    applyEditorStyle()
    rebuildMacroPanel()
    installEpicsWidgetPopupMenu(
      project = project,
      component = component,
      channelsProvider = {
        file.model.rawPvNames.mapNotNull { rawValue ->
          val trimmed = rawValue.trim()
          trimmed.takeIf(String::isNotBlank)
        }
      },
      primaryChannelProvider = {
        val plan = EpicsPvlistWidgetSupport.buildMonitorPlan(
          file.model,
          runtimeService.defaultProtocol,
        )
        plan.rows.firstOrNull { row -> row.definitionKey != null }?.channelName
          ?: file.model.rawPvNames.firstOrNull()?.trim()?.takeIf(String::isNotBlank)
      },
      sourceLabelProvider = { file.model.sourceLabel.ifBlank { EpicsPvlistWidgetVirtualFile.TAB_TITLE } },
      exportFileProvider = {
        file.model.sourcePath
          ?.takeIf(String::isNotBlank)
          ?.let(com.intellij.openapi.vfs.LocalFileSystem.getInstance()::findFileByPath)
      },
    )
    refreshTimer.initialDelay = 0
    refreshTimer.start()
    refreshViewState()
  }

  override fun getComponent(): JComponent = component

  override fun getPreferredFocusedComponent(): JComponent? = addChannelsArea

  override fun getFile(): VirtualFile = file

  override fun getName(): String = EpicsPvlistWidgetVirtualFile.TAB_TITLE

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() {
    refreshTimer.stop()
    if (channelTable.isEditing) {
      channelTable.cellEditor?.cancelCellEditing()
    }
    runtimeService.releaseWidgetPvlistSession(file.widgetId)
  }

  private fun buildUi() {
    val controlRow = JPanel(FlowLayout(FlowLayout.LEFT, 8, 0)).apply {
      isOpaque = false
      add(configureChannelsButton)
    }

    configureChannelsButton.addActionListener { showOverlay() }
    channelFilterField.document.addDocumentListener(
      object : DocumentListener {
        override fun insertUpdate(event: DocumentEvent?) {
          applyChannelFilter()
        }

        override fun removeUpdate(event: DocumentEvent?) {
          applyChannelFilter()
        }

        override fun changedUpdate(event: DocumentEvent?) {
          applyChannelFilter()
        }
      },
    )

    sourceLabel.border = BorderFactory.createEmptyBorder(0, 0, 2, 0)
    messageLabel.border = BorderFactory.createEmptyBorder(0, 0, 10, 0)

    macrosFieldsContainer.layout = BoxLayout(macrosFieldsContainer, BoxLayout.Y_AXIS)
    macrosFieldsContainer.isOpaque = false

    val macrosSection = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      add(macrosTitleLabel)
      add(Box.createVerticalStrut(8))
      add(macrosFieldsContainer)
    }

    val topSection = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      add(controlRow)
      add(Box.createVerticalStrut(8))
      add(sourceLabel)
      add(messageLabel)
    }

    channelTable.setShowGrid(false)
    channelTable.fillsViewportHeight = true
    channelTable.rowHeight = (channelTable.font.size * 1.55).toInt().coerceAtLeast(20)
    channelTable.rowMargin = 0
    channelTable.intercellSpacing = Dimension(0, 0)
    channelTable.columnModel.getColumn(0).preferredWidth = 280
    channelTable.columnModel.getColumn(0).cellRenderer = PvlistChannelCellRenderer()
    channelTable.columnModel.getColumn(1).preferredWidth = 420
    channelTable.columnModel.getColumn(1).cellRenderer = PvlistValueCellRenderer()
    channelTable.columnModel.getColumn(1).cellEditor = channelEditor
    channelTable.putClientProperty("terminateEditOnFocusLost", false)
    channelTable.addMouseMotionListener(
      object : MouseMotionAdapter() {
        override fun mouseMoved(event: MouseEvent) {
          val row = channelTable.rowAtPoint(event.point)
          val column = channelTable.columnAtPoint(event.point)
          val channelName = row.takeIf { it >= 0 }
            ?.let(channelTable::convertRowIndexToModel)
            ?.let(tableModel::getRow)
            ?.channelName
            ?.trim()
            .orEmpty()
          channelTable.cursor =
            if (column == 0 && channelName.isNotBlank()) {
              Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
            } else {
              Cursor.getDefaultCursor()
            }
        }
      },
    )
    channelTable.addMouseListener(
      object : MouseAdapter() {
        override fun mouseExited(event: MouseEvent) {
          channelTable.cursor = Cursor.getDefaultCursor()
        }

        override fun mouseClicked(event: MouseEvent) {
          if (event.button != MouseEvent.BUTTON1 || event.clickCount != 1) {
            return
          }
          val row = channelTable.rowAtPoint(event.point)
          val column = channelTable.columnAtPoint(event.point)
          if (row < 0 || column != 0) {
            return
          }
          val rowState = tableModel.getRow(channelTable.convertRowIndexToModel(row))
          val recordName = rowState.channelName.trim()
          if (recordName.isBlank()) {
            return
          }
          openEpicsWidget(project, recordName)
        }
      },
    )

    addChannelsButton.addActionListener {
      val previousMacros = file.model.macroNames.toList()
      if (EpicsPvlistWidgetSupport.replaceChannels(file.model, addChannelsArea.text)) {
        if (previousMacros != file.model.macroNames) {
          rebuildMacroPanel()
        }
        refreshViewState()
      }
      showMainContent()
    }

    val addSection = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      add(createInputBox(addChannelsArea))
      add(Box.createVerticalStrut(6))
      add(
        JPanel(FlowLayout(FlowLayout.LEFT, 0, 0)).apply {
          isOpaque = false
          add(addChannelsButton)
          alignmentX = Component.LEFT_ALIGNMENT
        },
      )
    }

    topSection.add(macrosSection)

    val channelsSection = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      channelsTitleLabel.alignmentX = Component.LEFT_ALIGNMENT
      channelFilterField.alignmentX = Component.LEFT_ALIGNMENT
      add(channelsTitleLabel)
      add(Box.createVerticalStrut(6))
      add(channelFilterField)
      add(Box.createVerticalStrut(8))
      add(JBScrollPane(channelTable).apply {
        border = BorderFactory.createEmptyBorder()
        alignmentX = Component.LEFT_ALIGNMENT
      })
    }

    val mainContent = JPanel(BorderLayout(0, 16)).apply {
      border = BorderFactory.createEmptyBorder(8, 8, 8, 8)
      isOpaque = false
      add(topSection, BorderLayout.NORTH)
      add(channelsSection, BorderLayout.CENTER)
    }

    val overlayHeader = JPanel(BorderLayout()).apply {
      isOpaque = false
      add(overlayTitleLabel, BorderLayout.WEST)
    }

    val overlayContent = JPanel(BorderLayout(0, 16)).apply {
      border = BorderFactory.createEmptyBorder(8, 8, 8, 8)
      isOpaque = false
      add(overlayHeader, BorderLayout.NORTH)
      add(addSection, BorderLayout.CENTER)
    }

    component.border = BorderFactory.createEmptyBorder(4, 4, 4, 4)
    cardPanel.isOpaque = false
    cardPanel.add(mainContent, MAIN_CARD)
    cardPanel.add(overlayContent, OVERLAY_CARD)
    component.add(cardPanel, BorderLayout.CENTER)
    showMainContent()
  }

  private fun applyEditorStyle() {
    val scheme = EditorColorsManager.getInstance().globalScheme
    val background = scheme.defaultBackground
    val foreground = scheme.defaultForeground
    val font = buildMonospaceEditorFont(scheme.editorFontName, scheme.editorFontSize)

    component.background = background
    component.foreground = foreground

    listOf(
      sourceLabel,
      messageLabel,
      addChannelsArea,
      channelTable,
      channelFilterField,
    ).forEach { component ->
      component.background = background
      component.foreground = foreground
      component.font = font
    }

    listOf(configureChannelsButton, addChannelsButton).forEach { button ->
      button.font = font
    }

    macrosFieldsContainer.background = background
    macrosFieldsContainer.foreground = foreground
    macrosFieldsContainer.font = font

    listOf(
      macrosTitleLabel,
      channelsTitleLabel,
      overlayTitleLabel,
    ).forEach { label ->
      label.background = background
      label.foreground = foreground
      label.font = font.deriveFont(Font.BOLD, font.size2D)
    }

    listOf(addChannelsArea).forEach { area ->
      area.isOpaque = true
      area.background = ColorUtil.mix(background, foreground, 0.06)
      area.border = BorderFactory.createCompoundBorder(
        BorderFactory.createLineBorder(ColorUtil.mix(background, foreground, 0.18)),
        BorderFactory.createEmptyBorder(4, 4, 4, 4),
      )
      area.caretColor = foreground
      area.selectionColor = Color(foreground.red, foreground.green, foreground.blue, 48)
      area.selectedTextColor = foreground
    }
  }

  private fun rebuildMacroPanel() {
    val font = channelTable.font
    macrosFieldsContainer.removeAll()
    macroFields.clear()

    if (file.model.macroNames.isEmpty()) {
      macrosFieldsContainer.add(JLabel("No macros detected.").apply {
        this.font = channelTable.font
        foreground = channelTable.foreground
      })
    } else {
      file.model.macroNames.forEach { macroName ->
        val field = JTextField(file.model.macroValues[macroName].orEmpty(), 20).apply {
          maximumSize = Dimension(Int.MAX_VALUE, preferredSize.height)
          this.font = font
          foreground = channelTable.foreground
          caretColor = channelTable.foreground
          isOpaque = false
          border = BorderFactory.createEmptyBorder(2, 0, 2, 0)
          addActionListener {
            file.model.macroValues[macroName] = text
            refreshViewState()
          }
        }
        macroFields[macroName] = field
        macrosFieldsContainer.add(
          JPanel(BorderLayout(8, 0)).apply {
            isOpaque = false
            alignmentX = Component.LEFT_ALIGNMENT
            add(
              JLabel("$macroName =").apply {
                this.font = font
                foreground = channelTable.foreground
              },
              BorderLayout.WEST,
            )
            add(field, BorderLayout.CENTER)
          },
        )
        macrosFieldsContainer.add(Box.createVerticalStrut(4))
      }
    }

    macrosFieldsContainer.revalidate()
    macrosFieldsContainer.repaint()
  }

  private fun refreshViewState() {
    if (!runtimeService.isMonitoringActive()) {
      runtimeService.startMonitoring()
    }
    val viewState = runtimeService.getWidgetPvlistViewState(file.widgetId, file.model)

    sourceLabel.text = "Source: ${file.model.sourceLabel}"
    messageLabel.text = viewState.message.orEmpty()
    messageLabel.isVisible = !viewState.message.isNullOrBlank()

    if (!channelTable.isEditing) {
      tableModel.setRows(viewState.rows)
    }
    updateChannelEmptyText()
  }

  private fun applyChannelFilter() {
    if (channelTable.isEditing) {
      channelTable.cellEditor?.cancelCellEditing()
    }
    tableModel.setFilterQuery(channelFilterField.text)
    updateChannelEmptyText()
  }

  private fun updateChannelEmptyText() {
    channelTable.emptyText.text =
      if (tableModel.hasRows() && tableModel.isFilterActive()) {
        "No records match filter."
      } else {
        "No channel rows are available."
      }
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

  private inner class PvlistChannelCellRenderer : DefaultTableCellRenderer() {
    override fun getTableCellRendererComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      hasFocus: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val component = super.getTableCellRendererComponent(
        table,
        value,
        isSelected,
        hasFocus,
        row,
        column,
      )
      val channelName = tableModel.getRow(table.convertRowIndexToModel(row)).channelName.trim()
      border = BorderFactory.createEmptyBorder(1, 6, 1, 6)
      foreground =
        if (isSelected) {
          table.selectionForeground
        } else if (channelName.isNotBlank()) {
          ColorUtil.mix(table.foreground, Color(0, 102, 204), 0.55)
        } else {
          table.foreground
        }
      toolTipText = if (channelName.isNotBlank()) "Open in Probe" else null
      return component
    }
  }

  private inner class PvlistValueCellRenderer : DefaultTableCellRenderer() {
    override fun getTableCellRendererComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      hasFocus: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val component = super.getTableCellRendererComponent(
        table,
        value,
        isSelected,
        hasFocus,
        row,
        column,
      )
      border = BorderFactory.createEmptyBorder(1, 6, 1, 6)
      return component
    }
  }

  private inner class PvlistValueCellEditor : AbstractCellEditor(), TableCellEditor {
    private val textField = JTextField()
    private var editingRowKey: String? = null

    init {
      textField.font = channelTable.font
      textField.border = BorderFactory.createEmptyBorder(0, 6, 0, 6)
      textField.addActionListener { stopCellEditing() }
    }

    override fun getTableCellEditorComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val rowState = tableModel.getRow(row)
      editingRowKey = rowState.definitionKey
      textField.text = rowState.value
      return textField
    }

    override fun getCellEditorValue(): Any = textField.text

    override fun stopCellEditing(): Boolean {
      val rowKey = editingRowKey
      val input = textField.text
      val stopped = super.stopCellEditing()
      editingRowKey = null
      if (stopped && rowKey != null) {
        runtimeService.requestPutWidgetPvlistValue(file.widgetId, file.model, rowKey, input)
        refreshViewState()
      }
      return stopped
    }

    override fun cancelCellEditing() {
      editingRowKey = null
      super.cancelCellEditing()
    }
  }

  private class PvlistTableModel : AbstractTableModel() {
    private val allRows = mutableListOf<EpicsPvlistWidgetRowViewState>()
    private val rows = mutableListOf<EpicsPvlistWidgetRowViewState>()
    private var filterTerms = emptyList<String>()

    override fun getRowCount(): Int = rows.size

    override fun getColumnCount(): Int = 2

    override fun getColumnName(column: Int): String {
      return when (column) {
        0 -> "Channel"
        else -> "Value"
      }
    }

    override fun getValueAt(rowIndex: Int, columnIndex: Int): Any {
      val row = rows[rowIndex]
      return when (columnIndex) {
        0 -> row.channelName
        else -> row.value
      }
    }

    override fun isCellEditable(rowIndex: Int, columnIndex: Int): Boolean {
      return columnIndex == 1 && rows.getOrNull(rowIndex)?.canPut == true
    }

    fun setFilterQuery(value: String?) {
      filterTerms = value.orEmpty()
        .trim()
        .lowercase(Locale.ROOT)
        .split(Regex("\\s+"))
        .filter(String::isNotBlank)
      applyFilter()
    }

    fun setRows(newRows: List<EpicsPvlistWidgetRowViewState>) {
      allRows.clear()
      allRows.addAll(newRows)
      applyFilter()
    }

    fun getRow(rowIndex: Int): EpicsPvlistWidgetRowViewState = rows[rowIndex]

    fun hasRows(): Boolean = allRows.isNotEmpty()

    fun isFilterActive(): Boolean = filterTerms.isNotEmpty()

    private fun applyFilter() {
      rows.clear()
      if (filterTerms.isEmpty()) {
        rows.addAll(allRows)
      } else {
        rows.addAll(
          allRows.filter { row ->
            val channelName = row.channelName.lowercase(Locale.ROOT)
            filterTerms.all(channelName::contains)
          },
        )
      }
      fireTableDataChanged()
    }
  }

  private companion object {
    private const val MAIN_CARD = "main"
    private const val OVERLAY_CARD = "overlay"
  }

  private fun showMainContent() {
    cardLayout.show(cardPanel, MAIN_CARD)
  }

  private fun showOverlay() {
    addChannelsArea.text = file.model.rawPvNames.joinToString("\n")
    cardLayout.show(cardPanel, OVERLAY_CARD)
    addChannelsArea.requestFocusInWindow()
  }
}

private fun createInputBox(textArea: JTextArea): JComponent {
  return JBScrollPane(textArea).apply {
    border = BorderFactory.createEmptyBorder()
    viewport.isOpaque = false
    isOpaque = false
    alignmentX = Component.LEFT_ALIGNMENT
    preferredSize = Dimension(720, textArea.preferredSize.height.coerceAtLeast(420))
  }
}

private class PromptTextArea(
  private val prompt: String,
  rows: Int,
  columns: Int,
) : JTextArea(rows, columns) {
  override fun paintComponent(graphics: Graphics) {
    super.paintComponent(graphics)
    if (text.isNotEmpty() || hasFocus()) {
      return
    }
    graphics.font = font
    graphics.color = Color(foreground.red, foreground.green, foreground.blue, 110)
    val metrics = graphics.getFontMetrics(font)
    graphics.drawString(prompt, insets.left + 2, insets.top + metrics.ascent)
  }
}
