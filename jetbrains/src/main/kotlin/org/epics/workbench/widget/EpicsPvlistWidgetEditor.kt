package org.epics.workbench.widget

import com.intellij.openapi.Disposable
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.fileChooser.FileChooserFactory
import com.intellij.openapi.fileChooser.FileSaverDescriptor
import com.intellij.openapi.fileEditor.FileEditor
import com.intellij.openapi.fileEditor.FileEditorLocation
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorPolicy
import com.intellij.openapi.fileEditor.FileEditorProvider
import com.intellij.openapi.fileEditor.FileEditorState
import com.intellij.openapi.fileEditor.FileEditorStateLevel
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import com.intellij.ui.ColorUtil
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.table.JBTable
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSourceKind
import org.epics.workbench.pvlist.EpicsPvlistWidgetSupport
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsPvlistWidgetRowViewState
import org.epics.workbench.runtime.EpicsPvlistWidgetViewState
import java.awt.BorderLayout
import java.awt.CardLayout
import java.awt.Color
import java.awt.Component
import java.awt.Dimension
import java.awt.FlowLayout
import java.awt.Font
import java.awt.Graphics
import java.beans.PropertyChangeListener
import java.nio.file.Files
import java.nio.file.Path
import java.util.LinkedHashMap
import java.util.Locale
import java.util.UUID
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
import javax.swing.table.AbstractTableModel
import javax.swing.table.TableCellEditor

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
  private val saveButton = JButton("Save")
  private val addOverlayButton = JButton("Add Channels")
  private val closeOverlayButton = JButton("Done")
  private val sourceLabel = JBLabel()
  private val messageLabel = JBLabel()
  private val macroFields = LinkedHashMap<String, JTextField>()
  private val macrosTitleLabel = JLabel("Macros")
  private val channelsTitleLabel = JLabel("Channels")
  private val overlayTitleLabel = JLabel("Add Channels")
  private val addChannelsArea = PromptTextArea("Add one channel per line", 4, 28)
  private val addChannelsButton = JButton("Add Channels")
  private val refreshTimer = Timer(1000) { refreshViewState() }
  private val component = JPanel(BorderLayout(0, 12))
  private val macrosFieldsContainer = JPanel()

  init {
    runtimeService.initialize()
    runtimeService.startMonitoring()
    buildUi()
    applyEditorStyle()
    rebuildMacroPanel()
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
      add(saveButton)
      add(addOverlayButton)
    }

    saveButton.addActionListener { saveWidgetToFile() }
    addOverlayButton.addActionListener { showOverlay() }
    closeOverlayButton.addActionListener { showMainContent() }

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
    channelTable.rowHeight = (channelTable.font.size * 1.8).toInt().coerceAtLeast(24)
    channelTable.columnModel.getColumn(0).preferredWidth = 320
    channelTable.columnModel.getColumn(1).preferredWidth = 480
    channelTable.columnModel.getColumn(1).cellEditor = channelEditor
    channelTable.putClientProperty("terminateEditOnFocusLost", false)

    addChannelsButton.addActionListener {
      val previousMacroCount = file.model.macroNames.size
      if (EpicsPvlistWidgetSupport.addChannels(file.model, addChannelsArea.text)) {
        if (file.model.macroNames.size != previousMacroCount) {
          rebuildMacroPanel()
        }
        refreshViewState()
      }
      addChannelsArea.text = ""
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
      add(channelsTitleLabel)
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
      add(
        JPanel(FlowLayout(FlowLayout.RIGHT, 0, 0)).apply {
          isOpaque = false
          add(closeOverlayButton)
        },
        BorderLayout.EAST,
      )
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
    ).forEach { component ->
      component.background = background
      component.foreground = foreground
      component.font = font
    }

    listOf(saveButton, addOverlayButton, closeOverlayButton, addChannelsButton).forEach { button ->
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
  }

  private fun saveWidgetToFile() {
    val targetPath = if (
      file.model.sourceKind == EpicsPvlistWidgetSourceKind.PVLIST &&
      !file.model.sourcePath.isNullOrBlank()
    ) {
      Path.of(file.model.sourcePath!!)
    } else {
      chooseSavePath() ?: return
    }

    val fileText = EpicsPvlistWidgetSupport.buildFileText(file.model)
    runCatching {
      Files.writeString(targetPath, fileText)
      LocalFileSystem.getInstance().refreshNioFiles(listOf(targetPath))
    }.onFailure { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to write ${targetPath.fileName}.", TITLE)
      return
    }

    file.model.sourceKind = EpicsPvlistWidgetSourceKind.PVLIST
    file.model.sourcePath = targetPath.toString()
    file.model.sourceLabel = targetPath.fileName.toString()
    refreshViewState()
  }

  private fun chooseSavePath(): Path? {
    val descriptor = FileSaverDescriptor(TITLE, "Save the current PV list widget as a .pvlist file.", "pvlist")
    val saver = FileChooserFactory.getInstance().createSaveFileDialog(descriptor, project)
    val sourcePath = file.model.sourcePath?.let(Path::of)
    val defaultDirectory = sourcePath?.parent ?: project.basePath?.let(Path::of)
    val defaultName = file.model.sourceLabel.substringBeforeLast('.').ifBlank { "epics" } + ".pvlist"
    return saver.save(defaultDirectory, defaultName)?.file?.toPath()
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

  private inner class PvlistValueCellEditor : AbstractCellEditor(), TableCellEditor {
    private val textField = JTextField()
    private var editingRowKey: String? = null

    init {
      textField.font = channelTable.font
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
    private val rows = mutableListOf<EpicsPvlistWidgetRowViewState>()

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

    fun setRows(newRows: List<EpicsPvlistWidgetRowViewState>) {
      rows.clear()
      rows.addAll(newRows)
      fireTableDataChanged()
    }

    fun getRow(rowIndex: Int): EpicsPvlistWidgetRowViewState = rows[rowIndex]
  }

  private companion object {
    private const val TITLE = "Save PV List File"
    private const val MAIN_CARD = "main"
    private const val OVERLAY_CARD = "overlay"
  }

  private fun showMainContent() {
    cardLayout.show(cardPanel, MAIN_CARD)
  }

  private fun showOverlay() {
    cardLayout.show(cardPanel, OVERLAY_CARD)
  }
}

private fun createInputBox(textArea: JTextArea): JComponent {
  return JBScrollPane(textArea).apply {
    border = BorderFactory.createEmptyBorder()
    viewport.isOpaque = false
    isOpaque = false
    alignmentX = Component.LEFT_ALIGNMENT
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
