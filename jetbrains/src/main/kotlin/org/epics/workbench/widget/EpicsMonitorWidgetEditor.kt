package org.epics.workbench.widget

import com.intellij.openapi.Disposable
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.editor.colors.EditorFontType
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
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import gov.aps.jca.Channel
import gov.aps.jca.Context
import gov.aps.jca.Monitor
import gov.aps.jca.TimeoutException
import gov.aps.jca.dbr.DBR
import gov.aps.jca.dbr.DBRType
import gov.aps.jca.dbr.DBR_Enum
import gov.aps.jca.dbr.LABELS
import gov.aps.jca.dbr.STS
import gov.aps.jca.dbr.TIME
import gov.aps.jca.dbr.TimeStamp
import gov.aps.jca.event.ConnectionEvent
import gov.aps.jca.event.ConnectionListener
import gov.aps.jca.event.MonitorEvent
import gov.aps.jca.event.MonitorListener
import org.epics.pva.client.ClientChannelState
import org.epics.pva.client.MonitorListener as PvaMonitorListener
import org.epics.pva.client.PVAChannel
import org.epics.pva.client.PVAClient
import org.epics.pva.data.PVAArray
import org.epics.pva.data.PVAData
import org.epics.pva.data.PVADouble
import org.epics.pva.data.PVAFloat
import org.epics.pva.data.PVAInt
import org.epics.pva.data.PVALong
import org.epics.pva.data.PVAShort
import org.epics.pva.data.PVAString
import org.epics.pva.data.PVAStringArray
import org.epics.pva.data.PVAStructure
import org.epics.pva.data.PVAValue
import org.epics.pva.data.nt.PVAEnum
import org.epics.pva.data.nt.PVATimeStamp
import org.epics.workbench.runtime.EpicsClientLibraries
import org.epics.workbench.runtime.EpicsRuntimeProjectConfigurationService
import org.epics.workbench.runtime.MonitorProtocol
import org.epics.workbench.ui.EpicsWidgetPalette
import org.epics.workbench.ui.applyEpicsWidgetButtonStyle
import org.epics.workbench.ui.applyEpicsWidgetTextFieldStyle
import org.epics.workbench.ui.buildEpicsWidgetPalette
import java.awt.BorderLayout
import java.awt.CardLayout
import java.awt.Component
import java.awt.Color
import java.awt.Dimension
import java.awt.FlowLayout
import java.awt.Font
import java.beans.PropertyChangeListener
import java.nio.file.Files
import java.nio.file.Path
import java.time.Instant
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.util.ArrayDeque
import java.util.Locale
import java.util.UUID
import java.util.concurrent.CountDownLatch
import java.util.concurrent.TimeUnit
import java.util.concurrent.atomic.AtomicBoolean
import javax.swing.BorderFactory
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JComponent
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.JScrollPane
import javax.swing.JTextArea
import javax.swing.JTextField
import javax.swing.text.DefaultCaret

internal class EpicsMonitorWidgetVirtualFile(
  initialChannels: List<String> = emptyList(),
) : LightVirtualFile(TAB_TITLE) {
  val widgetId: String = UUID.randomUUID().toString()
  val initialChannels: MutableList<String> = initialChannels.toMutableList()

  companion object {
    const val TAB_TITLE: String = "EPICS Monitor"
  }
}

class OpenEpicsMonitorWidgetAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    openEpicsMonitorWidget(project)
  }

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null
  }
}

internal fun openEpicsMonitorWidget(project: Project, initialChannels: List<String> = emptyList()) {
  FileEditorManager.getInstance(project).openFile(EpicsMonitorWidgetVirtualFile(initialChannels), true, true)
}

class EpicsMonitorWidgetFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: VirtualFile): Boolean {
    return file is EpicsMonitorWidgetVirtualFile
  }

  override fun createEditor(project: Project, file: VirtualFile): FileEditor {
    return EpicsMonitorWidgetFileEditor(project, file as EpicsMonitorWidgetVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-monitor-widget-editor"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsMonitorWidgetFileEditor(
  private val project: Project,
  private val file: EpicsMonitorWidgetVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val configurationService = project.service<EpicsRuntimeProjectConfigurationService>()
  private val cardLayout = CardLayout()
  private val cardPanel = JPanel(cardLayout)
  private val component = JPanel(BorderLayout())
  private val channelsPanel = JPanel()
  private val historyArea = JTextArea()
  private val historyScrollPane = JScrollPane(historyArea)
  private val addChannelButton = JButton("Add Channel")
  private val configureButton = JButton("Configure")
  private val exportButton = JButton("Export Data")
  private val bufferSizeField = JTextField(DEFAULT_BUFFER_SIZE.toString(), 10)
  private val bufferHintLabel = JLabel("Buffer size controls how many monitor lines are kept in the widget.")
  private val closeConfigButton = JButton("Done")
  private val historyLines = ArrayDeque<String>()
  private val channelRows = mutableListOf<MonitorChannelRow>()
  private val sessionsByRowId = linkedMapOf<String, MonitorChannelSession>()
  private var palette = buildEpicsWidgetPalette(Color.WHITE, Color.BLACK)

  private var monitoringActive = false
  private var bufferSize = DEFAULT_BUFFER_SIZE
  private var caContext: Context? = null
  private var pvaClient: PVAClient? = null
  private var defaultProtocol: MonitorProtocol = MonitorProtocol.CA

  init {
    buildUi()
    loadInitialChannels()
    applyEditorStyle()
    installEpicsWidgetPopupMenu(
      project = project,
      component = component,
      channelsProvider = {
        channelRows.mapNotNull { row ->
          row.currentChannelName.trim().takeIf(String::isNotBlank)
        }
      },
      primaryChannelProvider = {
        channelRows.firstNotNullOfOrNull { row ->
          row.currentChannelName.trim().takeIf(String::isNotBlank)
        }
      },
      sourceLabelProvider = { EpicsMonitorWidgetVirtualFile.TAB_TITLE },
    )
    startMonitoring()
  }

  override fun getComponent(): JComponent = component

  override fun getPreferredFocusedComponent(): JComponent? = channelRows.firstOrNull()?.textField ?: addChannelButton

  override fun getFile(): VirtualFile = file

  override fun getName(): String = EpicsMonitorWidgetVirtualFile.TAB_TITLE

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() {
    stopMonitoring()
  }

  private fun buildUi() {
    channelsPanel.layout = BoxLayout(channelsPanel, BoxLayout.Y_AXIS)
    channelsPanel.isOpaque = false

    val controlRow = JPanel(FlowLayout(FlowLayout.LEFT, 8, 0)).apply {
      isOpaque = false
      add(addChannelButton)
      add(configureButton)
      add(exportButton)
    }

    val stickyHeader = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = BorderFactory.createEmptyBorder(8, 8, 8, 8)
      add(controlRow)
      add(Box.createVerticalStrut(12))
      add(channelsPanel)
    }

    historyArea.isEditable = false
    historyArea.lineWrap = false
    historyArea.wrapStyleWord = false
    historyArea.border = BorderFactory.createEmptyBorder(8, 8, 8, 8)
    (historyArea.caret as? DefaultCaret)?.updatePolicy = DefaultCaret.NEVER_UPDATE

    historyScrollPane.border = BorderFactory.createEmptyBorder()

    val mainPanel = JPanel(BorderLayout()).apply {
      add(stickyHeader, BorderLayout.NORTH)
      add(historyScrollPane, BorderLayout.CENTER)
    }

    val configPanel = buildConfigPanel()

    cardPanel.add(mainPanel, MAIN_CARD)
    cardPanel.add(configPanel, CONFIG_CARD)
    component.add(cardPanel, BorderLayout.CENTER)

    addChannelButton.addActionListener { addChannelRow("") }
    configureButton.addActionListener { showConfig() }
    exportButton.addActionListener { exportMonitorFile() }
    closeConfigButton.addActionListener {
      applyBufferSize()
      showMain()
    }
    bufferSizeField.addActionListener { applyBufferSize() }

    if (channelRows.isEmpty()) {
      addChannelRow("")
    }
    showMain()
  }

  private fun buildConfigPanel(): JPanel {
    val fieldRow = JPanel(FlowLayout(FlowLayout.LEFT, 8, 0)).apply {
      isOpaque = false
      add(JLabel("Buffer size:"))
      bufferSizeField.maximumSize = Dimension(160, bufferSizeField.preferredSize.height)
      add(bufferSizeField)
    }

    return JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      border = BorderFactory.createEmptyBorder(20, 24, 20, 24)
      isOpaque = false
      add(JLabel("Monitor Configuration").apply {
        font = font.deriveFont(Font.BOLD, font.size2D + 2f)
        alignmentX = Component.LEFT_ALIGNMENT
      })
      add(Box.createVerticalStrut(12))
      add(bufferHintLabel.apply { alignmentX = Component.LEFT_ALIGNMENT })
      add(Box.createVerticalStrut(12))
      add(fieldRow.apply { alignmentX = Component.LEFT_ALIGNMENT })
      add(Box.createVerticalStrut(16))
      add(
        JPanel(FlowLayout(FlowLayout.LEFT, 0, 0)).apply {
          isOpaque = false
          add(closeConfigButton)
          alignmentX = Component.LEFT_ALIGNMENT
        },
      )
      add(Box.createVerticalGlue())
    }
  }

  private fun loadInitialChannels() {
    val initialChannels = file.initialChannels.filter { it.isNotBlank() }
    if (initialChannels.isEmpty()) {
      return
    }
    channelsPanel.removeAll()
    channelRows.clear()
    initialChannels.forEach(::addChannelRow)
  }

  private fun addChannelRow(initialValue: String) {
    val rowId = UUID.randomUUID().toString()
    val field = JTextField(initialValue, 32)
    val row = MonitorChannelRow(
      id = rowId,
      textField = field,
      currentChannelName = initialValue.trim(),
    )
    field.maximumSize = Dimension(Int.MAX_VALUE, field.preferredSize.height)
    field.addActionListener { applyChannelName(row) }
    row.panel.add(JLabel("Channel:"))
    row.panel.add(Box.createHorizontalStrut(8))
    row.panel.add(field)
    row.panel.add(Box.createHorizontalGlue())
    channelsPanel.add(row.panel)
    channelsPanel.add(Box.createVerticalStrut(6))
    channelRows += row
    channelsPanel.revalidate()
    channelsPanel.repaint()
  }

  private fun applyChannelName(row: MonitorChannelRow) {
    val nextName = row.textField.text.trim()
    if (row.currentChannelName == nextName) {
      return
    }
    sessionsByRowId.remove(row.id)?.close()
    row.currentChannelName = nextName
    updateInitialChannels()
    if (nextName.isNotBlank()) {
      startSessionForRow(row)
    }
  }

  private fun updateInitialChannels() {
    file.initialChannels.clear()
    file.initialChannels += channelRows.mapNotNull { row ->
      row.currentChannelName.takeIf(String::isNotBlank)
    }
  }

  private fun startMonitoring() {
    if (monitoringActive) {
      return
    }
    val configuration = configurationService.loadConfiguration()
    defaultProtocol = when (configuration.protocol) {
      org.epics.workbench.runtime.EpicsRuntimeProtocol.CA -> MonitorProtocol.CA
      org.epics.workbench.runtime.EpicsRuntimeProtocol.PVA -> MonitorProtocol.PVA
    }
    caContext = EpicsClientLibraries.createCaContext(configuration)
    pvaClient = EpicsClientLibraries.createPvaClient(configuration)
    monitoringActive = true
    channelRows.forEach(::startSessionForRow)
  }

  private fun stopMonitoring() {
    sessionsByRowId.values.toList().forEach(MonitorChannelSession::close)
    sessionsByRowId.clear()
    monitoringActive = false
    caContext?.let { context -> runCatching { context.destroy() } }
    pvaClient?.let { client -> runCatching { client.close() } }
    caContext = null
    pvaClient = null
  }

  private fun startSessionForRow(row: MonitorChannelRow) {
    val channelName = row.currentChannelName.trim()
    if (!monitoringActive || channelName.isBlank()) {
      return
    }
    if (sessionsByRowId.containsKey(row.id)) {
      return
    }
    val activeCaContext = caContext ?: return
    val activePvaClient = pvaClient ?: return
    val (protocol, pvName) = splitMonitorProtocol(channelName, defaultProtocol)
    val session = when (protocol) {
      MonitorProtocol.CA -> CaMonitorChannelSession(activeCaContext, pvName, ::appendMonitorLine)
      MonitorProtocol.PVA -> PvaMonitorChannelSession(activePvaClient, pvName, ::appendMonitorLine)
    }
    sessionsByRowId[row.id] = session
    session.start()
  }

  private fun appendMonitorLine(line: String) {
    com.intellij.openapi.application.ApplicationManager.getApplication().invokeLater {
      if (!component.isDisplayable) {
        return@invokeLater
      }
      val scrollBar = historyScrollPane.verticalScrollBar
      val previousValue = scrollBar.value
      val pinnedToBottom = scrollBar.value + scrollBar.visibleAmount >= scrollBar.maximum - 16
      historyLines += line
      while (historyLines.size > bufferSize) {
        if (historyLines.isNotEmpty()) {
          historyLines.removeFirst()
        }
      }
      historyArea.text = historyLines.joinToString("\n")
      if (pinnedToBottom) {
        scrollBar.value = scrollBar.maximum
      } else {
        scrollBar.value = previousValue.coerceAtMost((scrollBar.maximum - scrollBar.visibleAmount).coerceAtLeast(0))
      }
    }
  }

  private fun applyBufferSize() {
    val parsed = bufferSizeField.text.trim().toIntOrNull()
    if (parsed == null || parsed <= 0) {
      Messages.showErrorDialog(project, "Buffer size must be a positive integer.", TITLE)
      bufferSizeField.text = bufferSize.toString()
      return
    }
    bufferSize = parsed
    while (historyLines.size > bufferSize) {
      if (historyLines.isNotEmpty()) {
        historyLines.removeFirst()
      }
    }
    historyArea.text = historyLines.joinToString("\n")
  }

  private fun exportMonitorFile() {
    val path = chooseSavePath() ?: return
    val channels = channelRows.mapNotNull { row -> row.currentChannelName.trim().takeIf(String::isNotBlank) }
    val lines = buildList {
      add(
        when (channels.size) {
          1 -> "# monitor data for channel ${channels.first()}"
          0 -> "# monitor data exported from EPICS Monitor widget"
          else -> "# monitor data for channels ${channels.joinToString(", ")}"
        },
      )
      add("")
      addAll(historyLines)
    }
    runCatching {
      Files.writeString(path, lines.joinToString("\n"))
      LocalFileSystem.getInstance().refreshNioFiles(listOf(path))
    }.onFailure { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to write ${path.fileName}.", TITLE)
    }
  }

  private fun chooseSavePath(): Path? {
    val descriptor = FileSaverDescriptor(TITLE, "Export the current monitor history as a plain-text data file.", "txt")
    val saver = FileChooserFactory.getInstance().createSaveFileDialog(descriptor, project)
    val defaultDirectory = project.basePath?.let(Path::of)
    return saver.save(defaultDirectory, "epics-monitor-data.txt")?.file?.toPath()
  }

  private fun applyEditorStyle() {
    val scheme = EditorColorsManager.getInstance().globalScheme
    val background = scheme.defaultBackground
    val foreground = scheme.defaultForeground
    val editorFont = scheme.getFont(EditorFontType.PLAIN)
    palette = buildEpicsWidgetPalette(background, foreground)

    component.background = background
    component.foreground = foreground
    cardPanel.background = background
    cardPanel.foreground = foreground
    historyArea.background = palette.panelBackground
    historyArea.foreground = foreground
    historyArea.font = editorFont
    historyScrollPane.background = background
    historyScrollPane.viewport.background = palette.panelBackground
    historyScrollPane.border = BorderFactory.createLineBorder(palette.borderColor)
    channelsPanel.background = background
    channelsPanel.foreground = foreground
    bufferHintLabel.foreground = palette.mutedForeground
    bufferHintLabel.font = editorFont
    bufferSizeField.font = editorFont
    applyEpicsWidgetTextFieldStyle(bufferSizeField, palette)
    channelRows.forEach { row ->
      row.panel.background = background
      row.panel.foreground = foreground
      row.textField.font = editorFont
      applyEpicsWidgetTextFieldStyle(row.textField, palette)
    }
    listOf(addChannelButton, configureButton, exportButton, closeConfigButton).forEach { button ->
      button.font = editorFont
      applyEpicsWidgetButtonStyle(button, palette)
    }
  }

  private fun showMain() {
    cardLayout.show(cardPanel, MAIN_CARD)
  }

  private fun showConfig() {
    bufferSizeField.text = bufferSize.toString()
    cardLayout.show(cardPanel, CONFIG_CARD)
  }

  private data class MonitorChannelRow(
    val id: String,
    val textField: JTextField,
    var currentChannelName: String,
    val panel: JPanel = JPanel(),
  ) {
    init {
      panel.layout = BoxLayout(panel, BoxLayout.X_AXIS)
      panel.isOpaque = false
      panel.alignmentX = Component.LEFT_ALIGNMENT
      panel.maximumSize = Dimension(Int.MAX_VALUE, textField.preferredSize.height)
    }
  }

  private companion object {
    private const val TITLE = "Export Monitor Data"
    private const val MAIN_CARD = "main"
    private const val CONFIG_CARD = "config"
    private const val DEFAULT_BUFFER_SIZE = 500
  }
}

private interface MonitorChannelSession : AutoCloseable {
  fun start()
}

private class CaMonitorChannelSession(
  private val caContext: Context,
  private val pvName: String,
  private val lineConsumer: (String) -> Unit,
) : MonitorChannelSession {
  private val disposed = AtomicBoolean(false)
  private var channel: Channel? = null
  private var monitor: Monitor? = null

  override fun start() {
    ApplicationManager.getApplication().executeOnPooledThread {
      runCatching {
        caContext.attachCurrentThread()
        val latch = CountDownLatch(1)
        var connected = false
        val nextChannel = caContext.createChannel(
          pvName,
          ConnectionListener { event: ConnectionEvent ->
            if (event.isConnected) {
              connected = true
              latch.countDown()
            }
          },
        )
        caContext.flushIO()
        if (!latch.await(CA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS) || !connected) {
          runCatching { nextChannel.destroy() }
          throw TimeoutException("Timed out connecting to $pvName")
        }
        val fieldType = nextChannel.fieldType
        val labels = if (fieldType == DBRType.ENUM) fetchCaEnumLabels(caContext, nextChannel) else emptyList()
        val nextMonitor = nextChannel.addMonitor(
          caMonitorType(fieldType),
          nextChannel.elementCount.coerceAtLeast(1),
          Monitor.VALUE or Monitor.ALARM,
          MonitorListener { event: MonitorEvent ->
            if (!disposed.get() && event.status.isSuccessful) {
              lineConsumer(
                buildMonitorLine(
                  pvName = pvName,
                  timestamp = formatCaTimestamp(event.dbr),
                  value = formatCaValue(event.dbr, labels),
                  alarmText = formatCaAlarm(event.dbr),
                ),
              )
            }
          },
        )
        caContext.flushIO()
        if (disposed.get()) {
          runCatching { nextMonitor.clear() }
          runCatching { nextChannel.destroy() }
          return@runCatching
        }
        channel = nextChannel
        monitor = nextMonitor
      }
    }
  }

  override fun close() {
    if (!disposed.compareAndSet(false, true)) {
      return
    }
    runCatching { monitor?.clear() }
    runCatching { channel?.destroy() }
    monitor = null
    channel = null
  }
}

private class PvaMonitorChannelSession(
  private val pvaClient: PVAClient,
  private val pvName: String,
  private val lineConsumer: (String) -> Unit,
) : MonitorChannelSession {
  private val disposed = AtomicBoolean(false)
  private var channel: PVAChannel? = null
  private var subscription: AutoCloseable? = null

  override fun start() {
    ApplicationManager.getApplication().executeOnPooledThread {
      runCatching {
        val nextChannel = pvaClient.getChannel(pvName) { _: PVAChannel, _: ClientChannelState -> }
        nextChannel.connect().get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
        val nextSubscription = nextChannel.subscribe(
          "",
          PvaMonitorListener { _: PVAChannel, _, _, structure: PVAStructure ->
            if (!disposed.get()) {
              lineConsumer(
                buildMonitorLine(
                  pvName = pvName,
                  timestamp = formatPvaTimestamp(structure),
                  value = formatPvaStructure(structure),
                  alarmText = "",
                ),
              )
            }
          },
        )
        if (disposed.get()) {
          runCatching { nextSubscription.close() }
          runCatching { nextChannel.close() }
          return@runCatching
        }
        channel = nextChannel
        subscription = nextSubscription
      }
    }
  }

  override fun close() {
    if (!disposed.compareAndSet(false, true)) {
      return
    }
    runCatching { subscription?.close() }
    runCatching { channel?.close() }
    subscription = null
    channel = null
  }
}

private fun splitMonitorProtocol(
  value: String,
  defaultProtocol: MonitorProtocol,
): Pair<MonitorProtocol, String> {
  return when {
    value.startsWith("pva://", ignoreCase = true) -> MonitorProtocol.PVA to value.removePrefix("pva://").removePrefix("PVA://")
    value.startsWith("ca://", ignoreCase = true) -> MonitorProtocol.CA to value.removePrefix("ca://").removePrefix("CA://")
    else -> defaultProtocol to value
  }
}

private fun buildMonitorLine(
  pvName: String,
  timestamp: String,
  value: String,
  alarmText: String,
): String {
  return buildString {
    append(pvName.padEnd(30))
    append(' ')
    append(timestamp)
    append(' ')
    append(value)
    if (alarmText.isNotBlank()) {
      append(' ')
      append(alarmText)
    }
  }
}

private fun caMonitorType(fieldType: DBRType): DBRType {
  return when (fieldType) {
    DBRType.STRING -> DBRType.TIME_STRING
    DBRType.SHORT -> DBRType.TIME_SHORT
    DBRType.FLOAT -> DBRType.TIME_FLOAT
    DBRType.ENUM -> DBRType.LABELS_ENUM
    DBRType.BYTE -> DBRType.TIME_BYTE
    DBRType.INT -> DBRType.TIME_INT
    DBRType.DOUBLE -> DBRType.TIME_DOUBLE
    else -> fieldType
  }
}

private fun fetchCaEnumLabels(caContext: Context, channel: Channel): List<String> {
  return try {
    caContext.attachCurrentThread()
    val labelsDbr = channel.get(DBRType.LABELS_ENUM, channel.elementCount.coerceAtLeast(1))
    when (labelsDbr) {
      is LABELS -> labelsDbr.labels?.map { it ?: "" }.orEmpty()
      else -> emptyList()
    }
  } catch (_: Exception) {
    emptyList()
  }
}

private fun formatCaValue(dbr: DBR?, fallbackChoices: List<String>): String {
  if (dbr == null) {
    return ""
  }

  if (dbr is DBR_Enum) {
    val index = dbr.enumValue.firstOrNull()?.toInt() ?: 0
    val labels = when (dbr) {
      is LABELS -> dbr.labels?.map { it ?: "" }.orEmpty()
      else -> fallbackChoices
    }
    return formatEnumValue(index, labels)
  }

  val value = dbr.value ?: return ""
  return formatRuntimeObject(value)
}

private fun formatCaAlarm(dbr: DBR?): String {
  val sts = dbr as? STS ?: return ""
  val status = sts.status?.name.orEmpty()
  val severity = sts.severity?.name.orEmpty()
  return listOf(status, severity).filter(String::isNotBlank).joinToString(" ")
}

private fun formatCaTimestamp(dbr: DBR?): String {
  val timeStamp = (dbr as? TIME)?.timeStamp ?: return ""
  return formatEpicsTimestamp(timeStamp)
}

private fun formatPvaTimestamp(structure: PVAStructure): String {
  val timeStamp = runCatching { PVATimeStamp.getTimeStamp(structure) }.getOrNull() ?: return ""
  val instant = runCatching { timeStamp.instant() }.getOrNull() ?: return ""
  return formatInstantTimestamp(instant)
}

private fun formatRuntimeObject(value: Any): String {
  return when (value) {
    is String -> value
    is Array<*> -> formatArray(value.map { it?.toString().orEmpty() })
    is ByteArray -> formatArray(value.map(Byte::toString))
    is ShortArray -> formatArray(value.map(Short::toString))
    is IntArray -> formatArray(value.map(Int::toString))
    is LongArray -> formatArray(value.map(Long::toString))
    is FloatArray -> formatArray(value.map(Float::toString))
    is DoubleArray -> formatArray(value.map(Double::toString))
    else -> value.toString()
  }
}

private fun formatArray(values: List<String>): String {
  if (values.isEmpty()) {
    return "[]"
  }
  return if (values.size == 1) {
    values.first()
  } else {
    val preview = values.take(5).joinToString(", ")
    if (values.size > 5) "[$preview, ...]" else "[$preview]"
  }
}

private fun formatPvaStructure(structure: PVAStructure): String {
  val valueField = structure.get<PVAData>("value")
  if (valueField == null) {
    return if (structure.get().isNotEmpty()) "Has data, but no value" else ""
  }

  if (valueField is PVAStructure) {
    val pvaEnum = runCatching { PVAEnum.fromStructure(valueField) }.getOrNull()
    if (pvaEnum != null) {
      val index = pvaEnum.get<PVAInt>("index")?.get() ?: 0
      val choices = pvaEnum.get<PVAStringArray>("choices")?.get()?.map { it ?: "" }.orEmpty()
      return formatEnumValue(index, choices)
    }
  }

  if (valueField is PVAArray) {
    return valueField.toString()
  }

  if (valueField is PVAValue) {
    return valueField.formatDisplayValue()
  }

  return valueField.toString()
}

private fun formatEnumValue(index: Int, choices: List<String>): String {
  val choice = choices.getOrNull(index).orEmpty()
  return "[$index] $choice".trimEnd()
}

private fun PVAValue.formatDisplayValue(): String {
  return when (this) {
    is PVAString -> get().orEmpty()
    is PVAShort -> get().toString()
    is PVAInt -> get().toString()
    is PVALong -> get().toString()
    is PVAFloat -> get().toString()
    is PVADouble -> get().toString()
    else -> toString()
  }
}

private fun formatEpicsTimestamp(timeStamp: TimeStamp): String {
  val epochSeconds = EPICS_EPOCH_UNIX_OFFSET_SECONDS + timeStamp.secPastEpoch()
  val instant = Instant.ofEpochSecond(epochSeconds, timeStamp.nsec())
  return formatInstantTimestamp(instant)
}

private fun formatInstantTimestamp(instant: Instant): String {
  return LocalDateTime.ofInstant(instant, ZoneId.systemDefault()).format(MONITOR_TIMESTAMP_FORMATTER)
}

private val MONITOR_TIMESTAMP_FORMATTER: DateTimeFormatter =
  DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SSSSSS")

private const val EPICS_EPOCH_UNIX_OFFSET_SECONDS: Long = 631_152_000L
private const val CA_CONNECT_TIMEOUT_MS: Long = 4000
private const val PVA_CONNECT_TIMEOUT_MS: Long = 4000
