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
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.components.JBTextField
import com.intellij.ui.table.JBTable
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSupport
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsPvlistWidgetRowViewState
import org.epics.workbench.ui.EpicsWidgetPalette
import org.epics.workbench.ui.applyEpicsWidgetButtonStyle
import org.epics.workbench.ui.applyEpicsWidgetTextAreaStyle
import org.epics.workbench.ui.applyEpicsWidgetTextFieldStyle
import org.epics.workbench.ui.buildEpicsWidgetPalette
import java.awt.BasicStroke
import java.awt.BorderLayout
import java.awt.CardLayout
import java.awt.Color
import java.awt.Component
import java.awt.Cursor
import java.awt.Dimension
import java.awt.FlowLayout
import java.awt.Font
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.RenderingHints
import java.beans.PropertyChangeListener
import java.util.EventObject
import java.util.LinkedHashMap
import java.util.Locale
import java.util.UUID
import java.awt.event.FocusAdapter
import java.awt.event.FocusEvent
import java.awt.event.KeyEvent
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.awt.event.MouseMotionAdapter
import javax.swing.AbstractAction
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
import javax.swing.KeyStroke
import javax.swing.SwingConstants
import javax.swing.Timer
import javax.swing.table.AbstractTableModel
import javax.swing.table.DefaultTableCellRenderer
import javax.swing.table.TableCellEditor
import javax.swing.table.TableCellRenderer
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
  private val fieldNameInput = JBTextField().apply {
    emptyText.text = "Add field to list, e.g. INP"
    toolTipText = "Press Enter to add a record field column."
    maximumSize = Dimension(240, preferredSize.height)
  }
  private val channelTableScrollPane = JBScrollPane(channelTable)
  private val overlayTitleLabel = JLabel("Configure Channels")
  private val addChannelsArea = PromptTextArea("One channel per line", 16, 28)
  private val addChannelsButton = JButton("OK")
  private val refreshTimer = Timer(1000) { refreshViewState() }
  private val component = JPanel(BorderLayout(0, 12))
  private val macrosFieldsContainer = JPanel()
  private val valueCellRenderer = PvlistValueCellRenderer()
  private val processCellRenderer = PvlistProcessCellRenderer()
  private val headerRenderer = PvlistHeaderRenderer()
  private var palette = buildEpicsWidgetPalette(Color.WHITE, Color.BLACK)
  private var hoveredFieldHeaderViewColumn: Int = -1
  private var hoveredChannelHeaderVisible: Boolean = false
  private var hoveredChannelCellViewRow: Int = -1

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
    fieldNameInput.addActionListener { addFieldColumnFromInput() }

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

    channelTable.setShowGrid(true)
    channelTable.fillsViewportHeight = true
    channelTable.rowHeight = (channelTable.font.size * 1.35).toInt().coerceAtLeast(18)
    channelTable.rowMargin = 0
    channelTable.intercellSpacing = Dimension(1, 0)
    channelTable.tableHeader.reorderingAllowed = false
    channelTable.tableHeader.resizingAllowed = true
    channelTable.putClientProperty("terminateEditOnFocusLost", false)
    configureChannelTableColumns()
    channelTable.addMouseMotionListener(
      object : MouseMotionAdapter() {
        override fun mouseMoved(event: MouseEvent) {
          updateHoveredChannelCell(viewRow = channelTable.rowAtPoint(event.point), viewColumn = channelTable.columnAtPoint(event.point))
          updateTableCursor(event)
        }
      },
    )
    channelTable.addMouseListener(
      object : MouseAdapter() {
        override fun mouseExited(event: MouseEvent) {
          updateHoveredChannelCell(-1, -1)
          channelTable.cursor = Cursor.getDefaultCursor()
        }

        override fun mouseClicked(event: MouseEvent) {
          if (event.button != MouseEvent.BUTTON1 || event.clickCount != 1) {
            return
          }
          val row = channelTable.rowAtPoint(event.point)
          val column = channelTable.columnAtPoint(event.point)
          if (row < 0 || column < 0) {
            return
          }
          val rowState = tableModel.getRow(channelTable.convertRowIndexToModel(row))
          when {
            tableModel.isProcessColumn(column) && rowState.canProcess -> {
              runtimeService.requestProcessWidgetPvlistRecord(
                rowState.recordName.orEmpty(),
                rowState.protocol,
              )
            }
            tableModel.isChannelColumn(column) -> {
              when (getChannelCellAction(row, column, event.x)) {
                PvlistChannelAction.AddBelow -> {
                  showOverlayWithInsertedChannel(rowState.sourceIndex + 1)
                  return
                }
                PvlistChannelAction.Remove -> {
                  removeChannel(rowState.sourceIndex)
                  return
                }
                null -> Unit
              }
              val recordName = rowState.probeTargetRecordName?.trim().orEmpty()
              if (recordName.isBlank()) {
                return
              }
              openEpicsWidget(project, recordName)
            }
          }
        }
      },
    )
    channelTable.tableHeader.addMouseMotionListener(
      object : MouseMotionAdapter() {
        override fun mouseMoved(event: MouseEvent) {
          val viewColumn = channelTable.tableHeader.columnAtPoint(event.point)
          updateHoveredChannelHeader(viewColumn)
          updateHoveredFieldHeader(viewColumn)
          channelTable.tableHeader.cursor =
            if (
              isChannelHeaderAddHotspot(viewColumn, event.x) ||
              isHeaderRemoveHotspot(viewColumn, event.x)
            ) {
              Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
            } else {
              Cursor.getDefaultCursor()
            }
        }
      },
    )
    channelTable.tableHeader.addMouseListener(
      object : MouseAdapter() {
        override fun mouseExited(event: MouseEvent) {
          updateHoveredChannelHeader(-1)
          updateHoveredFieldHeader(-1)
          channelTable.tableHeader.cursor = Cursor.getDefaultCursor()
        }

        override fun mouseClicked(event: MouseEvent) {
          if (event.button != MouseEvent.BUTTON1 || event.clickCount != 1) {
            return
          }
          val viewColumn = channelTable.tableHeader.columnAtPoint(event.point)
          if (isChannelHeaderAddHotspot(viewColumn, event.x)) {
            showOverlayWithInsertedChannel(0)
            return
          }
          val fieldName = tableModel.getFieldNameForColumn(viewColumn) ?: return
          if (!isHeaderRemoveHotspot(viewColumn, event.x)) {
            return
          }
          removeFieldColumn(fieldName)
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
      fieldNameInput.alignmentX = Component.LEFT_ALIGNMENT
      add(channelsTitleLabel)
      add(Box.createVerticalStrut(6))
      add(
        JPanel(FlowLayout(FlowLayout.LEFT, 8, 0)).apply {
          isOpaque = false
          alignmentX = Component.LEFT_ALIGNMENT
          add(channelFilterField)
          add(fieldNameInput)
        },
      )
      add(Box.createVerticalStrut(8))
      add(channelTableScrollPane.apply {
        border = BorderFactory.createLineBorder(palette.borderColor)
        alignmentX = Component.LEFT_ALIGNMENT
        viewport.isOpaque = true
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
    palette = buildEpicsWidgetPalette(background, foreground)

    component.background = background
    component.foreground = foreground
    component.isOpaque = true

    listOf(
      sourceLabel,
      messageLabel,
      addChannelsArea,
      channelTable,
    ).forEach { component ->
      component.background = background
      component.foreground = foreground
      component.font = font
    }

    sourceLabel.foreground = palette.mutedForeground
    messageLabel.foreground = palette.mutedForeground

    listOf(channelFilterField, fieldNameInput).forEach { field ->
      applyEpicsWidgetTextFieldStyle(field, palette)
      field.font = font
    }

    listOf(configureChannelsButton, addChannelsButton).forEach { button ->
      button.font = font
      applyEpicsWidgetButtonStyle(button, palette)
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

    applyEpicsWidgetTextAreaStyle(addChannelsArea, palette)
    addChannelsArea.font = font

    channelTable.background = background
    channelTable.foreground = foreground
    channelTable.selectionBackground = palette.selectionBackground
    channelTable.selectionForeground = palette.selectionForeground
    channelTable.rowHeight = (font.size * 1.35).toInt().coerceAtLeast(18)
    channelTable.gridColor = palette.separatorColor
    channelTable.tableHeader.background = palette.headerBackground
    channelTable.tableHeader.foreground = foreground
    channelTable.tableHeader.font = font.deriveFont(Font.BOLD, font.size2D)
    channelTable.tableHeader.border = BorderFactory.createMatteBorder(0, 0, 1, 0, palette.borderColor)
    channelTableScrollPane.background = background
    channelTableScrollPane.viewport.background = background
    channelTableScrollPane.border = BorderFactory.createLineBorder(palette.borderColor)
    channelEditor.textField.font = font
    channelEditor.applyStyle()
    processCellRenderer.applyStyle()
    configureChannelTableColumns()
  }

  private fun rebuildMacroPanel() {
    val font = channelTable.font
    macrosFieldsContainer.removeAll()
    macroFields.clear()

    if (file.model.macroNames.isEmpty()) {
      macrosFieldsContainer.add(JLabel("No macros detected.").apply {
        this.font = channelTable.font
        foreground = palette.mutedForeground
      })
    } else {
      file.model.macroNames.forEach { macroName ->
        val field = JTextField(file.model.macroValues[macroName].orEmpty(), 20).apply {
          maximumSize = Dimension(Int.MAX_VALUE, preferredSize.height)
          this.font = font
          applyEpicsWidgetTextFieldStyle(this, palette)
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
      val structureChanged = tableModel.setViewState(viewState.rows, viewState.fieldColumns)
      if (structureChanged) {
        configureChannelTableColumns()
      }
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

  private fun addFieldColumnFromInput() {
    val normalizedFieldName = EpicsPvlistWidgetSupport.normalizeFieldName(fieldNameInput.text)
    when {
      normalizedFieldName.isBlank() -> {
        com.intellij.openapi.ui.Messages.showWarningDialog(
          project,
          "Field names must match [A-Za-z_][A-Za-z0-9_]*.",
          EpicsPvlistWidgetVirtualFile.TAB_TITLE,
        )
      }
      normalizedFieldName == "RTYP" -> {
        com.intellij.openapi.ui.Messages.showInfoMessage(
          project,
          "Type is already shown in the fixed RTYP column.",
          EpicsPvlistWidgetVirtualFile.TAB_TITLE,
        )
      }
      EpicsPvlistWidgetSupport.getFieldNames(file.model).contains(normalizedFieldName) -> {
        com.intellij.openapi.ui.Messages.showInfoMessage(
          project,
          "Field $normalizedFieldName is already shown.",
          EpicsPvlistWidgetVirtualFile.TAB_TITLE,
        )
      }
      else -> {
        if (channelTable.isEditing) {
          channelTable.cellEditor?.cancelCellEditing()
        }
        file.model.fieldNames.add(normalizedFieldName)
        fieldNameInput.text = ""
        refreshViewState()
      }
    }
  }

  private fun removeFieldColumn(fieldName: String) {
    if (channelTable.isEditing) {
      channelTable.cellEditor?.cancelCellEditing()
    }
    val normalizedFieldName = EpicsPvlistWidgetSupport.normalizeFieldName(fieldName)
    val nextFieldNames = file.model.fieldNames
      .map(EpicsPvlistWidgetSupport::normalizeFieldName)
      .filter(String::isNotBlank)
      .filterNot { it == normalizedFieldName }
    file.model.fieldNames.clear()
    file.model.fieldNames.addAll(nextFieldNames)
    updateHoveredFieldHeader(-1)
    refreshViewState()
  }

  private fun removeChannel(sourceIndex: Int) {
    if (channelTable.isEditing) {
      channelTable.cellEditor?.cancelCellEditing()
    }
    if (!EpicsPvlistWidgetSupport.removeChannelAt(file.model, sourceIndex)) {
      return
    }
    rebuildMacroPanel()
    refreshViewState()
  }

  private fun showOverlayWithInsertedChannel(insertIndex: Int) {
    val safeInsertIndex = insertIndex.coerceIn(0, file.model.rawPvNames.size)
    val draftLines = file.model.rawPvNames.toMutableList().apply {
      add(safeInsertIndex, "")
    }
    var caretOffset = 0
    repeat(safeInsertIndex) { index ->
      caretOffset += draftLines[index].length + 1
    }
    showOverlay(
      channelDraftText = draftLines.joinToString("\n"),
      caretOffset = caretOffset,
    )
  }

  private fun configureChannelTableColumns() {
    for (columnIndex in 0 until channelTable.columnModel.columnCount) {
      val descriptor = tableModel.getColumnDescriptor(columnIndex)
      val column = channelTable.columnModel.getColumn(columnIndex)
      column.headerRenderer = headerRenderer
      when (descriptor) {
        PvlistColumnDescriptor.Channel -> {
          column.preferredWidth = 280
          column.cellRenderer = PvlistChannelCellRenderer()
          column.cellEditor = null
        }
        PvlistColumnDescriptor.Type -> {
          column.preferredWidth = 92
          column.cellRenderer = valueCellRenderer
          column.cellEditor = null
        }
        PvlistColumnDescriptor.Value -> {
          column.preferredWidth = 160
          column.cellRenderer = valueCellRenderer
          column.cellEditor = channelEditor
        }
        PvlistColumnDescriptor.Process -> {
          column.preferredWidth = 108
          column.cellRenderer = processCellRenderer
          column.cellEditor = null
        }
        is PvlistColumnDescriptor.Field -> {
          column.preferredWidth = 140
          column.cellRenderer = valueCellRenderer
          column.cellEditor = null
        }
      }
    }
    channelTable.tableHeader.revalidate()
    channelTable.tableHeader.repaint()
  }

  private fun getChannelCellAction(viewRow: Int, viewColumn: Int, mouseX: Int): PvlistChannelAction? {
    if (!tableModel.isChannelColumn(viewColumn)) {
      return null
    }
    val cellRect = channelTable.getCellRect(viewRow, viewColumn, false)
    val actionsRight = cellRect.x + cellRect.width - CHANNEL_ACTION_RIGHT_PADDING
    val removeLeft = actionsRight - CHANNEL_ACTION_BUTTON_SIZE
    val addLeft = removeLeft - CHANNEL_ACTION_GAP - CHANNEL_ACTION_BUTTON_SIZE
    return when {
      mouseX in addLeft until (addLeft + CHANNEL_ACTION_BUTTON_SIZE) -> PvlistChannelAction.AddBelow
      mouseX in removeLeft until (removeLeft + CHANNEL_ACTION_BUTTON_SIZE) -> PvlistChannelAction.Remove
      else -> null
    }
  }

  private fun updateTableCursor(event: MouseEvent) {
    val viewRow = channelTable.rowAtPoint(event.point)
    val viewColumn = channelTable.columnAtPoint(event.point)
    if (viewRow < 0 || viewColumn < 0) {
      channelTable.cursor = Cursor.getDefaultCursor()
      return
    }
    val row = tableModel.getRow(channelTable.convertRowIndexToModel(viewRow))
    channelTable.cursor =
      when {
        getChannelCellAction(viewRow, viewColumn, event.x) != null ->
          Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        tableModel.isChannelColumn(viewColumn) && !row.probeTargetRecordName.isNullOrBlank() ->
          Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        tableModel.isProcessColumn(viewColumn) && row.canProcess ->
          Cursor.getPredefinedCursor(Cursor.HAND_CURSOR)
        else -> Cursor.getDefaultCursor()
      }
  }

  private fun updateHoveredChannelCell(viewRow: Int, viewColumn: Int) {
    val normalizedRow = if (tableModel.isChannelColumn(viewColumn)) viewRow else -1
    if (hoveredChannelCellViewRow == normalizedRow) {
      return
    }
    hoveredChannelCellViewRow = normalizedRow
    channelTable.repaint()
  }

  private fun updateHoveredChannelHeader(viewColumn: Int) {
    val visible = tableModel.isChannelColumn(viewColumn)
    if (hoveredChannelHeaderVisible == visible) {
      return
    }
    hoveredChannelHeaderVisible = visible
    channelTable.tableHeader.repaint()
  }

  private fun updateHoveredFieldHeader(viewColumn: Int) {
    val normalizedColumn = if (tableModel.getFieldNameForColumn(viewColumn) != null) viewColumn else -1
    if (hoveredFieldHeaderViewColumn == normalizedColumn) {
      return
    }
    hoveredFieldHeaderViewColumn = normalizedColumn
    channelTable.tableHeader.repaint()
  }

  private fun isChannelHeaderAddHotspot(viewColumn: Int, mouseX: Int): Boolean {
    if (!tableModel.isChannelColumn(viewColumn)) {
      return false
    }
    val headerRect = channelTable.tableHeader.getHeaderRect(viewColumn)
    val left = headerRect.x + headerRect.width - HEADER_ACTION_HOTSPOT_WIDTH
    return mouseX in left..(left + CHANNEL_ACTION_BUTTON_SIZE)
  }

  private fun isHeaderRemoveHotspot(viewColumn: Int, mouseX: Int): Boolean {
    if (tableModel.getFieldNameForColumn(viewColumn) == null) {
      return false
    }
    val headerRect = channelTable.tableHeader.getHeaderRect(viewColumn)
    return mouseX >= headerRect.x + headerRect.width - HEADER_REMOVE_HOTSPOT_WIDTH &&
      mouseX <= headerRect.x + headerRect.width - HEADER_REMOVE_HOTSPOT_PADDING
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

  private fun rowBackgroundColor(isSelected: Boolean, row: Int): Color {
    return if (isSelected) {
      palette.selectionBackground
    } else if (row % 2 == 0) {
      palette.background
    } else {
      palette.panelBackground
    }
  }

  private fun DefaultTableCellRenderer.applyRowCellStyle(isSelected: Boolean, row: Int) {
    border = BorderFactory.createEmptyBorder(1, 8, 1, 8)
    background = rowBackgroundColor(isSelected, row)
    isOpaque = true
  }

  private inner class PvlistChannelCellRenderer : TableCellRenderer {
    override fun getTableCellRendererComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      hasFocus: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val rowState = tableModel.getRow(table.convertRowIndexToModel(row))
      val foreground =
        if (isSelected) {
          palette.selectionForeground
        } else if (!rowState.probeTargetRecordName.isNullOrBlank()) {
          palette.linkForeground
        } else {
          table.foreground
        }
      val titleLabel = JLabel(rowState.channelName, SwingConstants.CENTER).apply {
        this.foreground = foreground
        this.font = table.font
        toolTipText = if (!rowState.probeTargetRecordName.isNullOrBlank()) "Open in Probe" else null
      }
      val actionsPanel = JPanel(FlowLayout(FlowLayout.RIGHT, CHANNEL_ACTION_GAP, 0)).apply {
        isOpaque = false
        add(
          ActionGlyph(
            kind = PvlistChannelAction.AddBelow,
            glyphColor = if (isSelected) palette.selectionForeground else palette.mutedForeground,
            visible = hoveredChannelCellViewRow == row,
          ),
        )
        add(
          ActionGlyph(
            kind = PvlistChannelAction.Remove,
            glyphColor = if (isSelected) palette.selectionForeground else palette.mutedForeground,
            visible = hoveredChannelCellViewRow == row,
          ),
        )
      }
      return JPanel(BorderLayout(8, 0)).apply {
        isOpaque = true
        background = rowBackgroundColor(isSelected, row)
        border = BorderFactory.createEmptyBorder(1, 8, 1, 8)
        add(Box.createHorizontalStrut(CHANNEL_ACTION_GROUP_WIDTH), BorderLayout.WEST)
        add(titleLabel, BorderLayout.CENTER)
        add(actionsPanel, BorderLayout.EAST)
      }
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
      super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column)
      horizontalAlignment = SwingConstants.CENTER
      applyRowCellStyle(isSelected, row)
      foreground = if (isSelected) palette.selectionForeground else palette.foreground
      return this
    }
  }

  private inner class PvlistProcessCellRenderer : TableCellRenderer {
    private val button = JButton("Process")

    fun applyStyle() {
      button.font = channelTable.font
      applyEpicsWidgetButtonStyle(button, palette)
      button.border = BorderFactory.createCompoundBorder(
        BorderFactory.createLineBorder(palette.borderColor),
        BorderFactory.createEmptyBorder(1, 8, 1, 8),
      )
    }

    override fun getTableCellRendererComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      hasFocus: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val rowState = tableModel.getRow(table.convertRowIndexToModel(row))
      button.isEnabled = rowState.canProcess
      button.background = if (isSelected) palette.selectionBackground else palette.panelBackground
      button.foreground =
        if (rowState.canProcess) {
          if (isSelected) palette.selectionForeground else palette.foreground
        } else {
          palette.mutedForeground
        }
      return JPanel(BorderLayout()).apply {
        isOpaque = true
        background = rowBackgroundColor(isSelected, row)
        border = BorderFactory.createEmptyBorder(1, 8, 1, 8)
        add(button, BorderLayout.CENTER)
      }
    }
  }

  private inner class PvlistHeaderRenderer : TableCellRenderer {
    override fun getTableCellRendererComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      hasFocus: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val descriptor = tableModel.getColumnDescriptor(column)
      val removableField = descriptor as? PvlistColumnDescriptor.Field
      val isChannelColumn = descriptor == PvlistColumnDescriptor.Channel
      val showRemoveButton = removableField != null && hoveredFieldHeaderViewColumn == column
      val titleLabel = JLabel(descriptor.headerTitle, SwingConstants.CENTER).apply {
        foreground = palette.foreground
        font = table.font.deriveFont(Font.BOLD, table.font.size2D)
      }
      val actionIcon: JComponent =
        when {
          isChannelColumn ->
            ActionGlyph(
              kind = PvlistChannelAction.AddBelow,
              glyphColor = palette.mutedForeground,
              visible = hoveredChannelHeaderVisible,
            )

          else ->
            HeaderCloseGlyph(
              visible = showRemoveButton,
              glyphColor = if (showRemoveButton) palette.mutedForeground else palette.headerBackground,
            )
        }
      return JPanel(BorderLayout(8, 0)).apply {
        isOpaque = true
        background = palette.headerBackground
        border = BorderFactory.createEmptyBorder(4, 10, 4, 10)
        add(Box.createHorizontalStrut(HEADER_ICON_BOX_SIZE), BorderLayout.WEST)
        add(titleLabel, BorderLayout.CENTER)
        add(actionIcon, BorderLayout.EAST)
      }
    }
  }

  private inner class PvlistValueCellEditor : AbstractCellEditor(), TableCellEditor {
    val textField = JTextField()
    private var editingRowKey: String? = null

    init {
      textField.horizontalAlignment = JTextField.CENTER
      textField.addActionListener { stopCellEditing() }
      textField.addFocusListener(
        object : FocusAdapter() {
          override fun focusLost(event: FocusEvent?) {
            if (channelTable.isEditing) {
              cancelCellEditing()
            }
          }
        },
      )
      textField.inputMap.put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), "cancel-edit")
      textField.actionMap.put(
        "cancel-edit",
        object : AbstractAction() {
          override fun actionPerformed(event: java.awt.event.ActionEvent?) {
            cancelCellEditing()
          }
        },
      )
    }

    fun applyStyle() {
      textField.font = channelTable.font
      textField.horizontalAlignment = JTextField.CENTER
      applyEpicsWidgetTextFieldStyle(textField, palette)
      textField.border = BorderFactory.createCompoundBorder(
        BorderFactory.createLineBorder(palette.borderColor),
        BorderFactory.createEmptyBorder(1, 8, 1, 8),
      )
    }

    override fun isCellEditable(event: EventObject?): Boolean {
      return when (event) {
        is MouseEvent -> event.clickCount >= 2
        else -> true
      }
    }

    override fun getTableCellEditorComponent(
      table: JTable,
      value: Any?,
      isSelected: Boolean,
      row: Int,
      column: Int,
    ): Component {
      val rowState = tableModel.getRow(table.convertRowIndexToModel(row))
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
    private val columns = mutableListOf<PvlistColumnDescriptor>(
      PvlistColumnDescriptor.Channel,
      PvlistColumnDescriptor.Type,
      PvlistColumnDescriptor.Value,
      PvlistColumnDescriptor.Process,
    )
    private var filterTerms = emptyList<String>()

    override fun getRowCount(): Int = rows.size

    override fun getColumnCount(): Int = columns.size

    override fun getColumnName(column: Int): String = columns[column].headerTitle

    override fun getValueAt(rowIndex: Int, columnIndex: Int): Any {
      val row = rows[rowIndex]
      return when (val descriptor = columns[columnIndex]) {
        PvlistColumnDescriptor.Channel -> row.channelName
        PvlistColumnDescriptor.Type -> row.recordType
        PvlistColumnDescriptor.Value -> row.value
        PvlistColumnDescriptor.Process -> row.canProcess
        is PvlistColumnDescriptor.Field ->
          row.fieldCells.firstOrNull { it.name == descriptor.fieldName }?.value.orEmpty()
      }
    }

    override fun isCellEditable(rowIndex: Int, columnIndex: Int): Boolean {
      return columns.getOrNull(columnIndex) == PvlistColumnDescriptor.Value &&
        rows.getOrNull(rowIndex)?.canPut == true
    }

    fun getRow(rowIndex: Int): EpicsPvlistWidgetRowViewState = rows[rowIndex]

    fun getColumnDescriptor(columnIndex: Int): PvlistColumnDescriptor = columns[columnIndex]

    fun isChannelColumn(columnIndex: Int): Boolean = columns.getOrNull(columnIndex) == PvlistColumnDescriptor.Channel

    fun isProcessColumn(columnIndex: Int): Boolean = columns.getOrNull(columnIndex) == PvlistColumnDescriptor.Process

    fun getFieldNameForColumn(columnIndex: Int): String? {
      return (columns.getOrNull(columnIndex) as? PvlistColumnDescriptor.Field)?.fieldName
    }

    fun setFilterQuery(value: String?) {
      filterTerms = value.orEmpty()
        .trim()
        .lowercase(Locale.ROOT)
        .split(Regex("\\s+"))
        .filter(String::isNotBlank)
      applyFilter(structureChanged = false)
    }

    fun setViewState(
      newRows: List<EpicsPvlistWidgetRowViewState>,
      fieldColumns: List<String>,
    ): Boolean {
      val nextColumns = mutableListOf<PvlistColumnDescriptor>(
        PvlistColumnDescriptor.Channel,
        PvlistColumnDescriptor.Type,
        PvlistColumnDescriptor.Value,
        PvlistColumnDescriptor.Process,
      ).apply {
        fieldColumns.forEach { fieldName ->
          add(PvlistColumnDescriptor.Field(fieldName))
        }
      }
      val structureChanged = columns != nextColumns
      if (structureChanged) {
        columns.clear()
        columns.addAll(nextColumns)
      }
      allRows.clear()
      allRows.addAll(newRows)
      applyFilter(structureChanged)
      return structureChanged
    }

    fun hasRows(): Boolean = allRows.isNotEmpty()

    fun isFilterActive(): Boolean = filterTerms.isNotEmpty()

    private fun applyFilter(structureChanged: Boolean) {
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
      if (structureChanged) {
        fireTableStructureChanged()
      } else {
        fireTableDataChanged()
      }
    }
  }

  private companion object {
    private const val MAIN_CARD = "main"
    private const val OVERLAY_CARD = "overlay"
    private const val HEADER_ICON_BOX_SIZE = 16
    private const val HEADER_ACTION_HOTSPOT_WIDTH = 20
    private const val HEADER_REMOVE_HOTSPOT_WIDTH = 20
    private const val HEADER_REMOVE_HOTSPOT_PADDING = 4
    private const val CHANNEL_ACTION_BUTTON_SIZE = 16
    private const val CHANNEL_ACTION_GAP = 4
    private const val CHANNEL_ACTION_GROUP_WIDTH = CHANNEL_ACTION_BUTTON_SIZE * 2 + CHANNEL_ACTION_GAP
    private const val CHANNEL_ACTION_RIGHT_PADDING = 12
  }

  private fun showMainContent() {
    cardLayout.show(cardPanel, MAIN_CARD)
  }

  private fun showOverlay(
    channelDraftText: String = file.model.rawPvNames.joinToString("\n"),
    caretOffset: Int? = null,
  ) {
    addChannelsArea.text = channelDraftText
    cardLayout.show(cardPanel, OVERLAY_CARD)
    addChannelsArea.requestFocusInWindow()
    val targetCaret = (caretOffset ?: addChannelsArea.document.length)
      .coerceIn(0, addChannelsArea.document.length)
    addChannelsArea.caretPosition = targetCaret
  }
}

private sealed class PvlistColumnDescriptor(val headerTitle: String) {
  object Channel : PvlistColumnDescriptor("Channel")
  object Type : PvlistColumnDescriptor("Type")
  object Value : PvlistColumnDescriptor("Value")
  object Process : PvlistColumnDescriptor("Process")
  data class Field(val fieldName: String) : PvlistColumnDescriptor(fieldName)
}

private enum class PvlistChannelAction {
  AddBelow,
  Remove,
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

private class HeaderCloseGlyph(
  private val visible: Boolean,
  private val glyphColor: Color,
) : JComponent() {
  init {
    preferredSize = Dimension(16, 16)
    minimumSize = preferredSize
    maximumSize = preferredSize
    isOpaque = false
    toolTipText = null
  }

  override fun paintComponent(graphics: Graphics) {
    if (!visible) {
      return
    }
    val graphics2d = graphics.create() as Graphics2D
    graphics2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
    graphics2d.color = glyphColor
    graphics2d.stroke = BasicStroke(1.6f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND)
    val inset = 4
    graphics2d.drawLine(inset, inset, width - inset, height - inset)
    graphics2d.drawLine(width - inset, inset, inset, height - inset)
    graphics2d.dispose()
  }
}

private class ActionGlyph(
  private val kind: PvlistChannelAction,
  private val glyphColor: Color,
  private val visible: Boolean = true,
) : JComponent() {
  init {
    preferredSize = Dimension(16, 16)
    minimumSize = preferredSize
    maximumSize = preferredSize
    isOpaque = false
    toolTipText = null
  }

  override fun paintComponent(graphics: Graphics) {
    if (!visible) {
      return
    }
    val graphics2d = graphics.create() as Graphics2D
    graphics2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
    graphics2d.color = glyphColor
    graphics2d.stroke = BasicStroke(1.6f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND)
    val inset = 4
    when (kind) {
      PvlistChannelAction.AddBelow -> {
        graphics2d.drawLine(inset, height / 2, width - inset, height / 2)
        graphics2d.drawLine(width / 2, inset, width / 2, height - inset)
      }
      PvlistChannelAction.Remove -> {
        graphics2d.drawLine(inset, inset, width - inset, height - inset)
        graphics2d.drawLine(width - inset, inset, inset, height - inset)
      }
    }
    graphics2d.dispose()
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
