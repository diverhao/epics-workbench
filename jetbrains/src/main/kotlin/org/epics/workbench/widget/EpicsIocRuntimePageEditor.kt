package org.epics.workbench.widget

import com.intellij.openapi.Disposable
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.components.service
import com.intellij.openapi.fileEditor.FileEditor
import com.intellij.openapi.fileEditor.FileEditorLocation
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorPolicy
import com.intellij.openapi.fileEditor.FileEditorProvider
import com.intellij.openapi.fileEditor.FileEditorState
import com.intellij.openapi.fileEditor.FileEditorStateLevel
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import com.intellij.ui.JBColor
import com.intellij.ui.components.JBCheckBox
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.components.JBTextArea
import com.intellij.ui.components.JBTextField
import com.intellij.util.ui.JBUI
import org.epics.workbench.runtime.EpicsIocRuntimeEnvironmentEntry
import org.epics.workbench.runtime.EpicsIocRuntimeService
import org.epics.workbench.runtime.EpicsIocRuntimeStateListener
import org.epics.workbench.runtime.EpicsIocRuntimeVariable
import java.awt.BorderLayout
import java.awt.Component
import java.awt.Cursor
import java.awt.Dimension
import java.awt.FlowLayout
import java.awt.Font
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.GridBagConstraints
import java.awt.GridBagLayout
import java.awt.RenderingHints
import java.beans.PropertyChangeListener
import java.util.concurrent.atomic.AtomicInteger
import javax.swing.BorderFactory
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JComponent
import javax.swing.JLabel
import javax.swing.JLayer
import javax.swing.JPanel
import javax.swing.JScrollPane
import javax.swing.event.DocumentEvent
import javax.swing.event.DocumentListener
import javax.swing.plaf.LayerUI

internal enum class EpicsIocRuntimePageType(
  val tabTitle: String,
) {
  COMMANDS("IOC Runtime Commands"),
  VARIABLES("IOC Runtime Variables"),
  ENVIRONMENT("IOC Runtime Environment"),
  ;
}

internal class EpicsIocRuntimePageVirtualFile(
  val pageType: EpicsIocRuntimePageType,
  val startupPath: String,
  val startupName: String,
) : LightVirtualFile(pageType.tabTitle)

internal fun openEpicsIocRuntimePage(
  project: Project,
  pageType: EpicsIocRuntimePageType,
  startupFile: VirtualFile,
) {
  val startupPath = startupFile.path
  val manager = FileEditorManager.getInstance(project)
  val existing = manager.openFiles
    .filterIsInstance<EpicsIocRuntimePageVirtualFile>()
    .firstOrNull { file ->
      file.pageType == pageType && file.startupPath == startupPath
    }
  if (existing != null) {
    manager.openFile(existing, true, true)
    return
  }
  manager.openFile(EpicsIocRuntimePageVirtualFile(pageType, startupPath, startupFile.name), true, true)
}

class EpicsIocRuntimePageFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: VirtualFile): Boolean {
    return file is EpicsIocRuntimePageVirtualFile
  }

  override fun createEditor(project: Project, file: VirtualFile): FileEditor {
    return EpicsIocRuntimePageFileEditor(project, file as EpicsIocRuntimePageVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-ioc-runtime-page-editor"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsIocRuntimePageFileEditor(
  private val project: Project,
  private val file: EpicsIocRuntimePageVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val startupFile = LocalFileSystem.getInstance().findFileByPath(file.startupPath)
  private val pagePanel = when (file.pageType) {
    EpicsIocRuntimePageType.COMMANDS -> EpicsIocRuntimeCommandsPanel(project, startupFile)
    EpicsIocRuntimePageType.VARIABLES -> EpicsIocRuntimeVariablesPanel(project, startupFile)
    EpicsIocRuntimePageType.ENVIRONMENT -> EpicsIocRuntimeEnvironmentPanel(project, startupFile)
  }.apply {
    initializePanel()
  }

  override fun getComponent(): JComponent = pagePanel

  override fun getPreferredFocusedComponent(): JComponent? = pagePanel.resolvePreferredFocusTarget()

  override fun getFile(): VirtualFile = file

  override fun getName(): String = file.pageType.tabTitle

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() {
    pagePanel.dispose()
  }
}

private abstract class AbstractEpicsIocRuntimePagePanel(
  protected val project: Project,
  protected val startupFile: VirtualFile?,
  pageTitle: String,
) : JPanel(BorderLayout()), Disposable {
  protected val runtimeService = project.service<EpicsIocRuntimeService>()
  private val titleLabel = JBLabel(pageTitle).apply {
    font = font.deriveFont(Font.BOLD, font.size2D + 6f)
  }
  private val showConsoleButton = JButton("Show Running Terminal")
  private val startIocButton = JButton("Start IOC")
  private val stopIocButton = JButton("Stop IOC")
  protected val bodyPanel = JPanel(BorderLayout())
  private val overlayLayerUi = RuntimePageStoppedWatermarkLayerUi()
  private val overlayContainer = JLayer(bodyPanel, overlayLayerUi)
  private var lastRunningState: Boolean? = null

  init {
    border = JBUI.Borders.empty(18)

    val actionRow = JPanel(FlowLayout(FlowLayout.LEFT, 10, 0)).apply {
      isOpaque = false
      add(showConsoleButton)
      add(startIocButton)
      add(stopIocButton)
    }

    val headerPanel = JPanel(BorderLayout(12, 0)).apply {
      isOpaque = false
      add(titleLabel, BorderLayout.WEST)
      add(actionRow, BorderLayout.EAST)
    }

    add(headerPanel, BorderLayout.NORTH)
    add(overlayContainer, BorderLayout.CENTER)

    bodyPanel.isOpaque = false

    showConsoleButton.addActionListener {
      val file = startupFile ?: return@addActionListener
      runtimeService.showRunningConsole(file)
    }
    startIocButton.addActionListener {
      val file = startupFile ?: return@addActionListener
      val result = runtimeService.startIoc(file)
      result.exceptionOrNull()?.let { error ->
        Messages.showErrorDialog(project, error.message ?: "Failed to start ${file.name}.", pageTitle)
        return@addActionListener
      }
      refreshLifecycleState(forceReload = true)
    }
    stopIocButton.addActionListener {
      val file = startupFile ?: return@addActionListener
      runtimeService.stopIoc(file)
      refreshLifecycleState(forceReload = false)
    }

    val connection = project.messageBus.connect(this)
    connection.subscribe(
      EpicsIocRuntimeStateListener.TOPIC,
      object : EpicsIocRuntimeStateListener {
        override fun startupStateChanged(startupPath: String, running: Boolean) {
          if (startupFile?.path == startupPath) {
            ApplicationManager.getApplication().invokeLater {
              refreshLifecycleState(forceReload = running)
            }
          }
        }
      },
    )

  }

  override fun dispose() = Unit

  fun initializePanel() {
    buildBody()
    refreshLifecycleState(forceReload = true)
  }

  protected abstract fun buildBody()

  protected abstract fun reloadPageDataAsync()

  open fun resolvePreferredFocusTarget(): JComponent? = null

  private fun refreshLifecycleState(forceReload: Boolean) {
    val running = startupFile != null && runtimeService.isRunning(startupFile)
    showConsoleButton.isVisible = running
    startIocButton.isVisible = !running
    stopIocButton.isVisible = running
    overlayLayerUi.showStoppedWatermark = !running
    setEnabledRecursively(bodyPanel, running)
    overlayContainer.repaint()
    val previousState = lastRunningState
    lastRunningState = running
    if (running && (forceReload || previousState != true)) {
      reloadPageDataAsync()
    }
  }

  protected fun showErrorMessage(title: String, message: String) {
    ApplicationManager.getApplication().invokeLater {
      Messages.showErrorDialog(project, message, title)
    }
  }

  protected fun isRunning(): Boolean {
    return startupFile != null && runtimeService.isRunning(startupFile)
  }

  private fun setEnabledRecursively(component: Component, enabled: Boolean) {
    component.isEnabled = enabled
    if (component is JComponent) {
      component.components.forEach { child ->
        setEnabledRecursively(child, enabled)
      }
    }
  }
}

private class EpicsIocRuntimeCommandsPanel(
  project: Project,
  startupFile: VirtualFile?,
) : AbstractEpicsIocRuntimePagePanel(project, startupFile, "IOC Runtime Commands") {
  private val filterField = JBTextField().apply {
    emptyText.text = "Filter commands and parameters (space-separated AND terms)"
  }
  private val freeFormField = JBTextField()
  private val freeFormCaptureCheck = JBCheckBox("Capture output")
  private val freeFormSendButton = JButton("Send")
  private val loadingLabel = JBLabel("Loading commands...")
  private val rowsPanel = JPanel()
  private val rowsScrollPane = JBScrollPane(rowsPanel)
  private val loadGeneration = AtomicInteger(0)
  private var commandRows: List<CommandRow> = emptyList()
  private var maxCommandNameWidth = 0

  init {
    rowsPanel.layout = GridBagLayout()
    rowsPanel.isOpaque = false
    rowsPanel.border = JBUI.Borders.empty(12)
    rowsScrollPane.border = BorderFactory.createEmptyBorder()
    rowsScrollPane.verticalScrollBar.unitIncrement = 18
  }

  override fun buildBody() {
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.emptyTop(16)
    }

    val filterLabel = JBLabel("Filter commands").apply {
      preferredSize = Dimension(COMMAND_LABEL_WIDTH, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }
    filterField.maximumSize = Dimension(Int.MAX_VALUE, filterField.preferredSize.height)
    filterField.putClientProperty(
      "JTextField.placeholderText",
      "Space-separated AND terms across command names, arguments, and help",
    )
    installDocumentListener(filterField) { applyFilter() }

    val filterRow = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.X_AXIS)
      isOpaque = false
      alignmentX = Component.LEFT_ALIGNMENT
      add(filterLabel)
      add(Box.createHorizontalStrut(12))
      add(filterField)
      add(Box.createHorizontalGlue())
    }

    val freeFormLabel = JBLabel("Run any command").apply {
      preferredSize = Dimension(COMMAND_LABEL_WIDTH, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }

    val freeFormRow = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.X_AXIS)
      isOpaque = false
      alignmentX = Component.LEFT_ALIGNMENT
      add(freeFormLabel)
      add(Box.createHorizontalStrut(12))
      freeFormField.columns = 32
      freeFormField.maximumSize = Dimension(420, freeFormField.preferredSize.height)
      add(freeFormField)
      add(Box.createHorizontalStrut(10))
      add(freeFormCaptureCheck)
      add(Box.createHorizontalStrut(10))
      add(freeFormSendButton)
      add(Box.createHorizontalGlue())
    }

    fun sendFreeFormCommand() {
      startupFile ?: return
      val commandText = freeFormField.text.trim()
      if (commandText.isEmpty()) {
        return
      }
      runCommand(
        titlePrefix = "ioc-command",
        commandText = commandText,
        captureOutput = freeFormCaptureCheck.isSelected,
        onSuccess = { },
      )
    }

    freeFormSendButton.addActionListener { sendFreeFormCommand() }
    freeFormField.addActionListener { sendFreeFormCommand() }

    topPanel.add(filterRow)
    topPanel.add(Box.createVerticalStrut(10))
    topPanel.add(freeFormRow)
    topPanel.add(Box.createVerticalStrut(12))
    topPanel.add(JPanel(BorderLayout()).apply {
      isOpaque = false
      border = JBUI.Borders.customLineBottom(JBColor.border())
    })

    bodyPanel.add(topPanel, BorderLayout.NORTH)
    bodyPanel.add(rowsScrollPane, BorderLayout.CENTER)

    loadingLabel.border = JBUI.Borders.empty(12)
    rowsPanel.add(loadingLabel)
  }

  override fun reloadPageDataAsync() {
    val file = startupFile ?: return
    if (!isRunning()) {
      return
    }
    val generation = loadGeneration.incrementAndGet()
    ApplicationManager.getApplication().executeOnPooledThread {
      val commands = runCatching { runtimeService.getCommandNames(file) }.getOrElse { error ->
        showErrorMessage("IOC Runtime Commands", error.message ?: "Failed to load IOC commands.")
        return@executeOnPooledThread
      }
      val helpByCommand = runCatching { runtimeService.getCommandHelp(file, commands) }.getOrElse { error ->
        showErrorMessage("IOC Runtime Commands", error.message ?: "Failed to load IOC command help.")
        return@executeOnPooledThread
      }
      ApplicationManager.getApplication().invokeLater {
        if (loadGeneration.get() != generation) {
          return@invokeLater
        }
        rebuildRows(commands, helpByCommand)
      }
    }
  }

  override fun resolvePreferredFocusTarget(): JComponent = filterField

  private fun rebuildRows(commands: List<String>, helpByCommand: Map<String, String>) {
    rowsPanel.removeAll()
    maxCommandNameWidth = commands
      .map { commandName ->
        JBLabel(commandName).preferredSize.width
      }
      .maxOrNull()
      ?.coerceAtLeast(COMMAND_LABEL_WIDTH)
      ?: COMMAND_LABEL_WIDTH
    commandRows = commands.map { commandName ->
      CommandRow(commandName, helpByCommand[commandName].orEmpty())
    }
    commandRows.forEachIndexed { index, row ->
      rowsPanel.add(
        row.container,
        GridBagConstraints().apply {
          gridx = 0
          gridy = index
          weightx = 1.0
          fill = GridBagConstraints.HORIZONTAL
          anchor = GridBagConstraints.NORTHWEST
          insets = JBUI.insetsBottom(12)
        },
      )
    }
    rowsPanel.add(
      JPanel().apply { isOpaque = false },
      GridBagConstraints().apply {
        gridx = 0
        gridy = commandRows.size
        weightx = 1.0
        weighty = 1.0
        fill = GridBagConstraints.BOTH
      },
    )
    applyFilter()
    rowsPanel.revalidate()
    rowsPanel.repaint()
  }

  private fun applyFilter() {
    val terms = filterField.text
      .trim()
      .lowercase()
      .split(Regex("\\s+"))
      .filter(String::isNotEmpty)
    commandRows.forEach { row ->
      row.container.isVisible = terms.all(row.searchableText::contains)
    }
    rowsPanel.revalidate()
    rowsPanel.repaint()
  }

  private fun runCommand(
    titlePrefix: String,
    commandText: String,
    captureOutput: Boolean,
    onSuccess: () -> Unit,
  ) {
    val file = startupFile ?: return
    ApplicationManager.getApplication().executeOnPooledThread {
      runCatching {
        if (captureOutput) {
          val output = runtimeService.captureCommandOutput(file, commandText)
          runtimeService.openCapturedOutput(file, titlePrefix, commandText, output)
        } else {
          runtimeService.sendCommandText(file, commandText)
        }
      }.onFailure { error ->
        showErrorMessage("IOC Runtime Commands", error.message ?: "Failed to run IOC command.")
      }.onSuccess {
        onSuccess()
      }
    }
  }

  private inner class CommandRow(
    private val commandName: String,
    helpText: String,
  ) {
    val container = JPanel()
    val searchableText: String
      get() = buildString {
        append(commandName.lowercase())
        append(' ')
        append(argumentField.text.lowercase())
        append(' ')
        append(parameterHints.joinToString(" ").lowercase())
        append(' ')
        append(helpArea.text.lowercase())
      }

    private val parameterHints = parseParameterHints(helpText)
    private val nameLabel = JBLabel(commandName).apply {
      font = font.deriveFont(Font.BOLD)
      preferredSize = Dimension(maxCommandNameWidth, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }
    private val argumentField = JBTextField()
    private val captureOutputCheck = JBCheckBox("Capture output")
    private val sendButton = JButton("Send")
    private val helpArea = JBTextArea(helpText.ifBlank { "No help available." }).apply {
      isEditable = false
      isFocusable = false
      isRequestFocusEnabled = false
      lineWrap = true
      wrapStyleWord = true
      border = JBUI.Borders.empty(0, 0, 0, 0)
      isOpaque = false
      font = font.deriveFont(font.size2D - 1f)
      foreground = JBColor.GRAY
      columns = 96
      cursor = Cursor.getDefaultCursor()
      caretPosition = 0
    }

    init {
      container.layout = BorderLayout(0, 8)
      container.isOpaque = false
      container.alignmentX = Component.LEFT_ALIGNMENT
      container.border = JBUI.Borders.compound(
        JBUI.Borders.customLine(JBColor.border(), 1),
        JBUI.Borders.empty(10),
      )

      val row = JPanel().apply {
        layout = BoxLayout(this, BoxLayout.X_AXIS)
        isOpaque = false
        add(nameLabel)
        add(Box.createHorizontalStrut(12))
        argumentField.columns = 28
        argumentField.maximumSize = Dimension(420, argumentField.preferredSize.height)
        updatePlaceholder()
        add(argumentField)
        add(Box.createHorizontalStrut(10))
        add(captureOutputCheck)
        add(Box.createHorizontalStrut(10))
        add(sendButton)
        add(Box.createHorizontalGlue())
      }

      fun sendCommand() {
        val commandText = buildString {
          append(commandName)
          append('(')
          append(argumentField.text.trim())
          append(')')
        }
        runCommand(
          titlePrefix = "ioc-command",
          commandText = commandText,
          captureOutput = captureOutputCheck.isSelected,
          onSuccess = { },
        )
      }

      sendButton.addActionListener { sendCommand() }
      argumentField.addActionListener { sendCommand() }
      installDocumentListener(argumentField) {
        updatePlaceholder()
        applyFilter()
      }

      container.add(row, BorderLayout.NORTH)
      container.add(helpArea, BorderLayout.CENTER)
    }

    private fun updatePlaceholder() {
      val filledCount = parseFilledArgumentCount(argumentField.text)
      val remainingHints = parameterHints.drop(filledCount)
      argumentField.putClientProperty("JTextField.placeholderText", remainingHints.joinToString(", "))
      argumentField.repaint()
    }
  }

  companion object {
    private const val COMMAND_LABEL_WIDTH = 140

    private fun parseParameterHints(helpText: String): List<String> {
      val firstLine = helpText.lineSequence().firstOrNull()?.trim().orEmpty()
      if (firstLine.isEmpty()) {
        return emptyList()
      }
      val signature = firstLine.substringAfter(' ', "")
      if (signature.isBlank()) {
        return emptyList()
      }
      val hints = mutableListOf<String>()
      var index = 0
      while (index < signature.length) {
        while (index < signature.length && signature[index].isWhitespace()) {
          index += 1
        }
        if (index >= signature.length) {
          break
        }
        val quote = signature[index]
        if (quote == '\'' || quote == '"') {
          val end = signature.indexOf(quote, index + 1)
          if (end > index) {
            hints += signature.substring(index + 1, end)
            index = end + 1
          } else {
            hints += signature.substring(index + 1)
            break
          }
        } else {
          val start = index
          while (index < signature.length && !signature[index].isWhitespace()) {
            index += 1
          }
          hints += signature.substring(start, index)
        }
      }
      return hints
    }

    private fun parseFilledArgumentCount(argumentText: String): Int {
      var filled = 0
      var inSingleQuote = false
      var inDoubleQuote = false
      var currentHasText = false
      argumentText.forEach { character ->
        when {
          character == '\'' && !inDoubleQuote -> {
            inSingleQuote = !inSingleQuote
            currentHasText = true
          }
          character == '"' && !inSingleQuote -> {
            inDoubleQuote = !inDoubleQuote
            currentHasText = true
          }
          character == ',' && !inSingleQuote && !inDoubleQuote -> {
            if (currentHasText) {
              filled += 1
            }
            currentHasText = false
          }
          !character.isWhitespace() -> currentHasText = true
        }
      }
      if (currentHasText) {
        filled += 1
      }
      return filled
    }
  }
}

private class RuntimePageStoppedWatermarkLayerUi : LayerUI<JComponent>() {
  var showStoppedWatermark: Boolean = false
    set(value) {
      if (field != value) {
        field = value
      }
    }

  override fun paint(graphics: Graphics, component: JComponent) {
    super.paint(graphics, component)
    if (!showStoppedWatermark) {
      return
    }
    val g2 = graphics.create() as Graphics2D
    try {
      g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON)
      g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
      val font = component.font.deriveFont(Font.BOLD, 52f)
      g2.font = font
      g2.color = JBColor(0xC65353, 0xE06C75)
      val metrics = g2.fontMetrics
      val textWidth = metrics.stringWidth(WATERMARK_TEXT)
      val x = (component.width - textWidth) / 2
      var y = TOP_PADDING
      while (y < component.height + metrics.height) {
        g2.drawString(WATERMARK_TEXT, x, y)
        y += ROW_SPACING
      }
    } finally {
      g2.dispose()
    }
  }

  companion object {
    private const val WATERMARK_TEXT = "Stopped"
    private const val TOP_PADDING = 110
    private const val ROW_SPACING = 150
  }
}

private class EpicsIocRuntimeVariablesPanel(
  project: Project,
  startupFile: VirtualFile?,
) : AbstractEpicsIocRuntimePagePanel(project, startupFile, "IOC Runtime Variables") {
  private val warningLabel = JBLabel("Note: the current value may not reflect the real value.").apply {
    foreground = JBColor(0xC53A3A, 0xE06C75)
  }
  private val loadingLabel = JBLabel("Loading IOC runtime variables...")
  private val rowsPanel = JPanel()
  private var maxNameWidth = 0
  private var maxTypeWidth = 0

  init {
    rowsPanel.layout = GridBagLayout()
    rowsPanel.isOpaque = false
    rowsPanel.border = JBUI.Borders.empty(12)
  }

  override fun buildBody() {
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.emptyTop(16)
      add(warningLabel)
      add(Box.createVerticalStrut(12))
    }
    bodyPanel.add(topPanel, BorderLayout.NORTH)
    bodyPanel.add(JBScrollPane(rowsPanel).apply {
      border = BorderFactory.createEmptyBorder()
    }, BorderLayout.CENTER)
    rowsPanel.add(loadingLabel)
  }

  override fun reloadPageDataAsync() {
    val file = startupFile ?: return
    if (!isRunning()) {
      return
    }
    ApplicationManager.getApplication().executeOnPooledThread {
      val variables = runCatching { runtimeService.listRuntimeVariables(file) }.getOrElse { error ->
        showErrorMessage("IOC Runtime Variables", error.message ?: "Failed to load IOC runtime variables.")
        return@executeOnPooledThread
      }
      ApplicationManager.getApplication().invokeLater {
        rebuildRows(variables)
      }
    }
  }

  private fun rebuildRows(variables: List<EpicsIocRuntimeVariable>) {
    rowsPanel.removeAll()
    if (variables.isEmpty()) {
      rowsPanel.add(JBLabel("No IOC runtime variables reported by `var`."))
    } else {
      maxNameWidth = variables
        .map { variable -> JLabel(variable.name).preferredSize.width }
        .maxOrNull()
        ?.coerceAtLeast(160)
        ?: 160
      maxTypeWidth = variables
        .map { variable -> JLabel(variable.type).preferredSize.width }
        .maxOrNull()
        ?.coerceAtLeast(48)
        ?: 48
      variables.forEachIndexed { index, variable ->
        rowsPanel.add(
          buildVariableRow(variable),
          GridBagConstraints().apply {
            gridx = 0
            gridy = index
            weightx = 1.0
            fill = GridBagConstraints.HORIZONTAL
            anchor = GridBagConstraints.NORTHWEST
            insets = JBUI.insetsBottom(8)
          },
        )
      }
      rowsPanel.add(
        JPanel().apply { isOpaque = false },
        GridBagConstraints().apply {
          gridx = 0
          gridy = variables.size
          weightx = 1.0
          weighty = 1.0
          fill = GridBagConstraints.BOTH
        },
      )
    }
    rowsPanel.revalidate()
    rowsPanel.repaint()
  }

  private fun buildVariableRow(variable: EpicsIocRuntimeVariable): JComponent {
    val field = JBTextField(variable.value, 28)
    val setButton = JButton("Set")
    val nameLabel = JBLabel(variable.name).apply {
      font = font.deriveFont(Font.BOLD)
      preferredSize = Dimension(maxNameWidth, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }
    val typeLabel = JBLabel(variable.type).apply {
      foreground = JBColor.GRAY
      preferredSize = Dimension(maxTypeWidth, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }
    setButton.preferredSize = Dimension(96, setButton.preferredSize.height)
    setButton.minimumSize = setButton.preferredSize
    setButton.maximumSize = setButton.preferredSize

    fun applyValue() {
      val file = startupFile ?: return
      val value = field.text
      ApplicationManager.getApplication().executeOnPooledThread {
        runCatching {
          runtimeService.setRuntimeVariable(file, variable.name, value)
        }.onFailure { error ->
          showErrorMessage("IOC Runtime Variables", error.message ?: "Failed to set ${variable.name}.")
        }
      }
    }

    setButton.addActionListener { applyValue() }
    field.addActionListener { applyValue() }

    return JPanel().apply {
      layout = GridBagLayout()
      isOpaque = false
      add(
        nameLabel,
        GridBagConstraints().apply {
          gridx = 0
          gridy = 0
          anchor = GridBagConstraints.WEST
          insets = JBUI.insetsRight(10)
        },
      )
      add(
        typeLabel,
        GridBagConstraints().apply {
          gridx = 1
          gridy = 0
          anchor = GridBagConstraints.WEST
          insets = JBUI.insetsRight(12)
        },
      )
      add(
        field,
        GridBagConstraints().apply {
          gridx = 2
          gridy = 0
          weightx = 1.0
          fill = GridBagConstraints.HORIZONTAL
          insets = JBUI.insetsRight(12)
        },
      )
      add(
        setButton,
        GridBagConstraints().apply {
          gridx = 3
          gridy = 0
          anchor = GridBagConstraints.EAST
        },
      )
    }
  }
}

private class EpicsIocRuntimeEnvironmentPanel(
  project: Project,
  startupFile: VirtualFile?,
) : AbstractEpicsIocRuntimePagePanel(project, startupFile, "IOC Runtime Environment") {
  private val warningLabel = JBLabel("Note: the current value may not reflect the real value.").apply {
    foreground = JBColor(0xC53A3A, 0xE06C75)
  }
  private val rowsPanel = JPanel()
  private var maxNameWidth = 0

  init {
    rowsPanel.layout = GridBagLayout()
    rowsPanel.isOpaque = false
    rowsPanel.border = JBUI.Borders.empty(12)
  }

  override fun buildBody() {
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.emptyTop(16)
      add(warningLabel)
      add(Box.createVerticalStrut(12))
    }
    bodyPanel.add(topPanel, BorderLayout.NORTH)
    bodyPanel.add(JBScrollPane(rowsPanel).apply {
      border = BorderFactory.createEmptyBorder()
    }, BorderLayout.CENTER)
    rowsPanel.add(JBLabel("Loading IOC environment variables..."))
  }

  override fun reloadPageDataAsync() {
    val file = startupFile ?: return
    if (!isRunning()) {
      return
    }
    ApplicationManager.getApplication().executeOnPooledThread {
      val environmentEntries = runCatching { runtimeService.listRuntimeEnvironment(file) }.getOrElse { error ->
        showErrorMessage("IOC Runtime Environment", error.message ?: "Failed to load IOC environment.")
        return@executeOnPooledThread
      }
      ApplicationManager.getApplication().invokeLater {
        rebuildRows(environmentEntries)
      }
    }
  }

  private fun rebuildRows(entries: List<EpicsIocRuntimeEnvironmentEntry>) {
    rowsPanel.removeAll()
    if (entries.isEmpty()) {
      rowsPanel.add(JBLabel("No IOC environment variables reported by `epicsEnvShow`."))
    } else {
      maxNameWidth = entries
        .map { entry -> JLabel(entry.name).preferredSize.width }
        .maxOrNull()
        ?.coerceAtLeast(180)
        ?: 180
      entries.forEachIndexed { index, entry ->
        rowsPanel.add(
          buildEnvironmentRow(entry),
          GridBagConstraints().apply {
            gridx = 0
            gridy = index
            weightx = 1.0
            fill = GridBagConstraints.HORIZONTAL
            anchor = GridBagConstraints.NORTHWEST
            insets = JBUI.insetsBottom(8)
          },
        )
      }
      rowsPanel.add(
        JPanel().apply { isOpaque = false },
        GridBagConstraints().apply {
          gridx = 0
          gridy = entries.size
          weightx = 1.0
          weighty = 1.0
          fill = GridBagConstraints.BOTH
        },
      )
    }
    rowsPanel.revalidate()
    rowsPanel.repaint()
  }

  private fun buildEnvironmentRow(entry: EpicsIocRuntimeEnvironmentEntry): JComponent {
    val field = JBTextField(entry.value, 36)
    val setButton = JButton("Set")
    val nameLabel = JBLabel(entry.name).apply {
      font = font.deriveFont(Font.BOLD)
      preferredSize = Dimension(maxNameWidth, preferredSize.height)
      minimumSize = preferredSize
      maximumSize = preferredSize
    }
    setButton.preferredSize = Dimension(96, setButton.preferredSize.height)
    setButton.minimumSize = setButton.preferredSize
    setButton.maximumSize = setButton.preferredSize

    fun applyValue() {
      val file = startupFile ?: return
      val value = field.text
      ApplicationManager.getApplication().executeOnPooledThread {
        runCatching {
          runtimeService.setRuntimeEnvironment(file, entry.name, value)
        }.onFailure { error ->
          showErrorMessage("IOC Runtime Environment", error.message ?: "Failed to set ${entry.name}.")
        }
      }
    }

    setButton.addActionListener { applyValue() }
    field.addActionListener { applyValue() }

    return JPanel().apply {
      layout = GridBagLayout()
      isOpaque = false
      add(
        nameLabel,
        GridBagConstraints().apply {
          gridx = 0
          gridy = 0
          anchor = GridBagConstraints.WEST
          insets = JBUI.insetsRight(12)
        },
      )
      add(
        field,
        GridBagConstraints().apply {
          gridx = 1
          gridy = 0
          weightx = 1.0
          fill = GridBagConstraints.HORIZONTAL
          insets = JBUI.insetsRight(12)
        },
      )
      add(
        setButton,
        GridBagConstraints().apply {
          gridx = 2
          gridy = 0
          anchor = GridBagConstraints.EAST
        },
      )
    }
  }
}

private fun installDocumentListener(field: JBTextField, callback: () -> Unit) {
  field.document.addDocumentListener(
    object : DocumentListener {
      override fun insertUpdate(event: DocumentEvent) = callback()

      override fun removeUpdate(event: DocumentEvent) = callback()

      override fun changedUpdate(event: DocumentEvent) = callback()
    },
  )
}

class OpenEpicsIocRuntimeCommandsPageAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val startupFile = event.getData(com.intellij.openapi.actionSystem.CommonDataKeys.VIRTUAL_FILE) ?: return
    openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.COMMANDS, startupFile)
  }
}
