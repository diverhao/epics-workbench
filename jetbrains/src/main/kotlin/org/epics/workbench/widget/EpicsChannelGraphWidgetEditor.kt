package org.epics.workbench.widget

import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.Disposable
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.components.service
import com.intellij.openapi.fileChooser.FileChooser
import com.intellij.openapi.fileChooser.FileChooserDescriptor
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
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import com.intellij.ui.ColorUtil
import com.intellij.ui.JBColor
import gov.aps.jca.Channel
import gov.aps.jca.Context
import gov.aps.jca.Monitor
import gov.aps.jca.dbr.DBR
import gov.aps.jca.dbr.DBRType
import gov.aps.jca.dbr.DBR_Enum
import gov.aps.jca.dbr.LABELS
import gov.aps.jca.event.ConnectionEvent
import gov.aps.jca.event.ConnectionListener
import gov.aps.jca.event.MonitorEvent
import gov.aps.jca.event.MonitorListener
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.runtime.EpicsClientLibraries
import org.epics.workbench.runtime.EpicsRuntimeProjectConfigurationService
import org.epics.workbench.runtime.MonitorProtocol
import org.epics.pva.client.ClientChannelState
import org.epics.pva.client.MonitorListener as PvaMonitorListener
import org.epics.pva.client.PVAChannel
import org.epics.pva.client.PVAClient
import org.epics.pva.data.PVAArray
import org.epics.pva.data.PVAByte
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
import java.awt.BasicStroke
import java.awt.BorderLayout
import java.awt.Color
import java.awt.Dimension
import java.awt.FontMetrics
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.Point
import java.awt.Rectangle
import java.awt.RenderingHints
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.awt.geom.CubicCurve2D
import java.awt.geom.Line2D
import java.awt.geom.Path2D
import java.beans.PropertyChangeListener
import java.util.Locale
import java.util.LinkedHashMap
import java.util.UUID
import java.util.concurrent.CountDownLatch
import java.util.concurrent.TimeUnit
import java.util.concurrent.TimeoutException
import java.util.concurrent.atomic.AtomicBoolean
import javax.swing.BorderFactory
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JComponent
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.JScrollPane
import javax.swing.SwingConstants
import javax.swing.JTextField
import kotlin.math.atan2
import kotlin.math.cos
import kotlin.math.max
import kotlin.math.min
import kotlin.math.sin

internal class EpicsChannelGraphVirtualFile(
  val initialSources: List<ChannelGraphSource>,
  val seedRecordName: String? = null,
  initialMessage: String? = null,
) : LightVirtualFile(TAB_TITLE) {
  val widgetId: String = UUID.randomUUID().toString()
  var message: String? = initialMessage

  companion object {
    const val TAB_TITLE: String = "EPICS Channel Graph"
  }
}

private enum class ChannelGraphResolutionMode(
  private val displayLabel: String,
) {
  STATIC("Static"),
  DYNAMIC("Dynamic"),
  ;

  override fun toString(): String = displayLabel
}

class OpenEpicsChannelGraphAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    openEpicsChannelGraphWidget(project, sourceLabel = "Channel Graph", sourceText = "")
  }

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null
  }
}

internal fun openEpicsChannelGraphWidget(
  project: Project,
  sourceLabel: String,
  sourceText: String,
  seedRecordName: String? = null,
  message: String? = null,
  sourcePath: String? = null,
) {
  val initialSources = if (sourceText.isBlank()) {
    emptyList()
  } else {
    listOf(ChannelGraphSource(label = sourceLabel, text = sourceText, path = sourcePath))
  }
  FileEditorManager.getInstance(project).openFile(
    EpicsChannelGraphVirtualFile(
      initialSources = initialSources,
      seedRecordName = seedRecordName,
      initialMessage = message,
    ),
    true,
    true,
  )
}

class EpicsChannelGraphFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: VirtualFile): Boolean {
    return file is EpicsChannelGraphVirtualFile
  }

  override fun createEditor(project: Project, file: VirtualFile): FileEditor {
    return EpicsChannelGraphFileEditor(project, file as EpicsChannelGraphVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-channel-graph-editor"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsChannelGraphFileEditor(
  private val project: Project,
  private val file: EpicsChannelGraphVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val sourceEntries = LinkedHashMap<String, ChannelGraphSource>()
  private val initialSeedRecordName = file.seedRecordName?.trim()?.takeIf(String::isNotBlank)
  private val originNodeIds = linkedSetOf<String>()
  private var graphModel = ChannelGraphSupport.build("", "EPICS Channel Graph", originNodeIds.firstOrNull())
  private var currentMode = ChannelGraphResolutionMode.DYNAMIC
  private var runtimeSession: ChannelGraphRuntimeSession? = null
  private var runtimeSessionSerial = 0
  private val pendingDynamicRevealNodeIds = linkedSetOf<String>()
  private val addDatabaseFilesButton = JButton("Add Database Files")
  private val addChannelField = JTextField()
  private val addChannelButton = JButton("Add Channel")
  private val clearGraphButton = JButton("Clear Graph")
  private val sourceLabel = JLabel("", SwingConstants.LEFT)
  private val sourceFilesPanel = JPanel()
  private val hintLabel = JLabel("", SwingConstants.LEFT)
  private val messageLabel = JLabel("", SwingConstants.LEFT)
  private val canvas = ChannelGraphCanvas(graphModel, originNodeIds) { nodeId -> handleExpandRequest(nodeId) }
  private val component = JPanel(BorderLayout())

  init {
    file.initialSources.forEach { source ->
      sourceEntries[getChannelGraphSourceKey(source)] = source
    }
    if (sourceEntries.isEmpty() && initialSeedRecordName != null) {
      originNodeIds += initialSeedRecordName
    }

    val header = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      border = BorderFactory.createEmptyBorder(12, 12, 8, 12)
      isOpaque = false
      val toolbar = JPanel().apply {
        layout = BoxLayout(this, BoxLayout.X_AXIS)
        isOpaque = false
        addDatabaseFilesButton.addActionListener { addDatabaseFiles() }
        add(addDatabaseFilesButton)
        add(Box.createHorizontalStrut(12))
        addChannelField.columns = 18
        addChannelField.toolTipText = "Enter channel name"
        addChannelField.addActionListener { addOriginNodeFromInput() }
        add(addChannelField)
        add(Box.createHorizontalStrut(6))
        addChannelButton.addActionListener { addOriginNodeFromInput() }
        add(addChannelButton)
        add(Box.createHorizontalStrut(6))
        clearGraphButton.addActionListener { clearGraph() }
        add(clearGraphButton)
        add(Box.createHorizontalGlue())
      }
      add(toolbar)
      add(Box.createVerticalStrut(8))
      add(sourceLabel)
      add(Box.createVerticalStrut(6))
      sourceFilesPanel.layout = BoxLayout(sourceFilesPanel, BoxLayout.Y_AXIS)
      sourceFilesPanel.isOpaque = false
      add(sourceFilesPanel)
      add(Box.createVerticalStrut(6))
      add(hintLabel)
      add(Box.createVerticalStrut(6))
      add(messageLabel)
    }

    val scrollPane = JScrollPane(canvas).apply {
      border = BorderFactory.createEmptyBorder()
      viewport.isOpaque = true
      canvas.background = viewport.background
    }

    component.add(header, BorderLayout.NORTH)
    component.add(scrollPane, BorderLayout.CENTER)
    refreshSourceFilesPanel()
    currentMode = determineAutomaticMode()
    applyMode(currentMode, file.message)
  }

  private fun addDatabaseFiles() {
    val descriptor = FileChooserDescriptor(true, false, false, false, false, true).apply {
      title = "Add Database Files"
      description = "Select EPICS database files to include in the Channel Graph."
    }
    val selectedFiles = FileChooser.chooseFiles(descriptor, project, null)
      .filter(::isDatabaseGraphSource)
    if (selectedFiles.isEmpty()) {
      return
    }

    var changed = false
    selectedFiles.forEach { selectedFile ->
      val sourceText = runCatching { String(selectedFile.contentsToByteArray(), selectedFile.charset) }.getOrNull()
      if (sourceText == null) {
        Messages.showErrorDialog(project, "Failed to read ${selectedFile.name}.", "Add Database Files")
        return@forEach
      }
      val source = ChannelGraphSource(label = selectedFile.name, text = sourceText, path = selectedFile.path)
      val key = getChannelGraphSourceKey(source)
      if (key in sourceEntries) {
        return@forEach
      }
      sourceEntries[key] = source
      changed = true
    }

    if (changed) {
      refreshSourceFilesPanel()
      applyMode(determineAutomaticMode())
    }
  }

  private fun refreshSourceFilesPanel() {
    sourceFilesPanel.removeAll()
    sourceEntries.values.forEach { source ->
      val row = JPanel().apply {
        layout = BoxLayout(this, BoxLayout.X_AXIS)
        isOpaque = false
      }
      val removeButton = JButton("×").apply {
        toolTipText = "Remove database file"
        addActionListener { removeDatabaseSource(source) }
      }
      val pathLabel = JLabel(source.path ?: source.label).apply {
        toolTipText = source.path ?: source.label
      }
      row.add(removeButton)
      row.add(Box.createHorizontalStrut(6))
      row.add(pathLabel)
      row.alignmentX = JComponent.LEFT_ALIGNMENT
      sourceFilesPanel.add(row)
      sourceFilesPanel.add(Box.createVerticalStrut(4))
    }
    sourceFilesPanel.isVisible = sourceEntries.isNotEmpty()
    sourceFilesPanel.revalidate()
    sourceFilesPanel.repaint()
  }

  private fun removeDatabaseSource(source: ChannelGraphSource) {
    sourceEntries.remove(getChannelGraphSourceKey(source))
    refreshSourceFilesPanel()
    applyMode(determineAutomaticMode())
  }

  private fun addOriginNodeFromInput() {
    val nodeId = addChannelField.text.trim()
    if (nodeId.isBlank()) {
      return
    }
    addChannelField.text = ""
    addOriginNode(nodeId)
  }

  private fun addOriginNode(nodeId: String) {
    val normalizedNodeId = nodeId.trim()
    if (normalizedNodeId.isBlank()) {
      return
    }
    originNodeIds += normalizedNodeId
    when (currentMode) {
      ChannelGraphResolutionMode.STATIC -> {
        canvas.addOrigin(normalizedNodeId)
        messageLabel.text = ""
        messageLabel.isVisible = false
      }

      ChannelGraphResolutionMode.DYNAMIC -> {
        pendingDynamicRevealNodeIds += normalizedNodeId
        ensureDynamicRuntimeSession().addOriginNode(normalizedNodeId)
      }
    }
  }

  private fun clearGraph() {
    originNodeIds.clear()
    pendingDynamicRevealNodeIds.clear()
    when (currentMode) {
      ChannelGraphResolutionMode.STATIC -> {
        canvas.clearGraphView()
        messageLabel.text = ""
        messageLabel.isVisible = false
      }

      ChannelGraphResolutionMode.DYNAMIC -> {
        runtimeSession?.clearGraph()
      }
    }
  }

  private fun rebuildGraphModel(forcedMessage: String? = null, resetState: Boolean = true) {
    val combinedText = sourceEntries.values.joinToString("\n\n") { it.text }
    val combinedLabel = buildChannelGraphSourceLabel(sourceEntries.values.toList())
    graphModel = ChannelGraphSupport.build(combinedText, combinedLabel, originNodeIds.firstOrNull())
    canvas.updateModel(graphModel, originNodeIds, resetState = resetState)
    sourceLabel.text = "Source: ${graphModel.sourceLabel}"
    messageLabel.text = forcedMessage
      ?: if (originNodeIds.isEmpty()) "Enter a channel name to start the Channel Graph." else graphModel.message.orEmpty()
    messageLabel.isVisible = messageLabel.text.isNotBlank()
    component.revalidate()
    component.repaint()
  }

  private fun applyMode(
    mode: ChannelGraphResolutionMode,
    forcedMessage: String? = null,
  ) {
    currentMode = mode
    runtimeSessionSerial += 1
    runtimeSession?.close()
    runtimeSession = null
    pendingDynamicRevealNodeIds.clear()

    when (mode) {
      ChannelGraphResolutionMode.STATIC -> {
        hintLabel.text = "Drag nodes to reposition. Double-click a node to expand one more hop."
        rebuildGraphModel(forcedMessage = forcedMessage, resetState = true)
      }

      ChannelGraphResolutionMode.DYNAMIC -> {
        hintLabel.text = "Drag nodes to reposition. Double-click a node to expand one more hop."
        val dynamicOriginNodeIds = originNodeIds.toList()
        val configuration = project.service<EpicsRuntimeProjectConfigurationService>().loadConfiguration()
        val protocol = when (configuration.protocol) {
          org.epics.workbench.runtime.EpicsRuntimeProtocol.PVA -> MonitorProtocol.PVA
          else -> MonitorProtocol.CA
        }
        val sourceName = sourceEntries.values.firstOrNull()?.label ?: "EPICS Channel Graph"
        val initialModel = buildInitialDynamicModel(dynamicOriginNodeIds, protocol, sourceName)
        graphModel = initialModel
        canvas.updateModel(initialModel, dynamicOriginNodeIds.toSet(), resetState = true)
        sourceLabel.text = "Source: ${initialModel.sourceLabel}"
        messageLabel.text = initialModel.message.orEmpty()
        messageLabel.isVisible = messageLabel.text.isNotBlank()
        pendingDynamicRevealNodeIds += dynamicOriginNodeIds
        ensureDynamicRuntimeSession(dynamicOriginNodeIds).start()
      }
    }
  }

  private fun ensureDynamicRuntimeSession(initialOriginNodeIds: List<String> = originNodeIds.toList()): ChannelGraphRuntimeSession {
    runtimeSession?.let { return it }

    val configuration = project.service<EpicsRuntimeProjectConfigurationService>().loadConfiguration()
    val protocol = when (configuration.protocol) {
      org.epics.workbench.runtime.EpicsRuntimeProtocol.PVA -> MonitorProtocol.PVA
      else -> MonitorProtocol.CA
    }
    val sourceName = sourceEntries.values.firstOrNull()?.label ?: "EPICS Channel Graph"
    val sessionSerial = runtimeSessionSerial
    var createdSession: ChannelGraphRuntimeSession? = null
    val session = ChannelGraphRuntimeSession(
      project = project,
      initialOriginNodeIds = initialOriginNodeIds,
      protocol = protocol,
      sourceName = sourceName,
    ) { updatedModel ->
      ApplicationManager.getApplication().invokeLater(
        {
          if (project.isDisposed || runtimeSessionSerial != sessionSerial) {
            return@invokeLater
          }
          graphModel = updatedModel
          canvas.updateModel(updatedModel, createdSession?.getOriginNodeIds().orEmpty(), resetState = false)
          pendingDynamicRevealNodeIds.toList().forEach { pendingNodeId ->
            if (pendingNodeId in updatedModel.nodeById) {
              canvas.addOrigin(pendingNodeId)
              pendingDynamicRevealNodeIds.remove(pendingNodeId)
            }
          }
          sourceLabel.text = "Source: ${updatedModel.sourceLabel}"
          messageLabel.text = updatedModel.message.orEmpty()
          messageLabel.isVisible = messageLabel.text.isNotBlank()
          component.revalidate()
          component.repaint()
        },
        { project.isDisposed || runtimeSessionSerial != sessionSerial },
      )
    }
    createdSession = session
    runtimeSession = session
    return session
  }

  private fun handleExpandRequest(nodeId: String) {
    when (currentMode) {
      ChannelGraphResolutionMode.STATIC -> canvas.revealNeighbors(nodeId)
      ChannelGraphResolutionMode.DYNAMIC -> {
        canvas.markExpanded(nodeId)
        pendingDynamicRevealNodeIds += nodeId
        runtimeSession?.requestExpand(nodeId)
      }
    }
  }

  override fun getComponent(): JComponent = component

  override fun getPreferredFocusedComponent(): JComponent? = canvas

  override fun getFile(): VirtualFile = file

  override fun getName(): String = EpicsChannelGraphVirtualFile.TAB_TITLE

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() {
    runtimeSession?.close()
    runtimeSession = null
  }

  private fun determineAutomaticMode(): ChannelGraphResolutionMode {
    return if (sourceEntries.values.any(::isStaticDatabaseSource)) {
      ChannelGraphResolutionMode.STATIC
    } else {
      ChannelGraphResolutionMode.DYNAMIC
    }
  }
}

private class ChannelGraphCanvas(
  private var model: ChannelGraphModel,
  initialOriginNodeIds: Set<String>,
  private val expandRequestHandler: (String) -> Unit,
) : JPanel() {
  private val originNodeIds = linkedSetOf<String>().apply { addAll(initialOriginNodeIds) }
  private val visibleNodeIds = linkedSetOf<String>()
  private val expandedNodeIds = linkedSetOf<String>()
  private val nodePositions = linkedMapOf<String, Point>()
  private val nodeBounds = linkedMapOf<String, Rectangle>()
  private var dragNodeId: String? = null
  private var dragOffset: Point = Point()

  init {
    background = JBColor.PanelBackground
    isOpaque = true
    border = BorderFactory.createEmptyBorder(16, 16, 16, 16)
    layoutInitialState()
    installInteraction()
  }

  fun updateModel(model: ChannelGraphModel, originNodeIds: Set<String>, resetState: Boolean = true) {
    this.model = model
    this.originNodeIds.clear()
    this.originNodeIds.addAll(originNodeIds)
    if (resetState) {
      visibleNodeIds.clear()
      expandedNodeIds.clear()
      nodePositions.clear()
      nodeBounds.clear()
      layoutInitialState()
    } else {
      val validIds = model.nodeById.keys
      visibleNodeIds.retainAll(validIds)
      expandedNodeIds.retainAll(validIds)
      nodePositions.keys.retainAll(validIds)
      nodeBounds.keys.retainAll(validIds)
      if (visibleNodeIds.isEmpty()) {
        layoutInitialState()
      } else {
        revalidate()
        repaint()
      }
    }
  }

  override fun getPreferredSize(): Dimension {
    val bounds = computeCanvasBounds()
    return Dimension(max(bounds.width + 64, 900), max(bounds.height + 64, 700))
  }

  override fun paintComponent(graphics: Graphics) {
    super.paintComponent(graphics)
    val g = graphics as Graphics2D
    g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
    g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON)
    g.stroke = BasicStroke(2f)

    val metrics = g.fontMetrics
    val visibleNodes = model.nodes.filter { visibleNodeIds.contains(it.id) }
    nodeBounds.clear()
    visibleNodes.forEach { node ->
      nodeBounds[node.id] = createNodeBounds(node, metrics)
    }

    paintEdges(g, metrics)
    visibleNodes.forEach { node ->
      paintNode(g, node, nodeBounds[node.id]!!, metrics)
    }
  }

  private fun layoutInitialState() {
    if (originNodeIds.isEmpty()) {
      revalidate()
      repaint()
      return
    }
    originNodeIds.forEach(::addOrigin)
    revalidate()
    repaint()
  }

  private fun layoutAllNodes() {
    val spacingX = 260
    val spacingY = 170
    model.nodes.forEachIndexed { index, node ->
      val column = index % 4
      val row = index / 4
      nodePositions.putIfAbsent(node.id, Point(120 + column * spacingX, 120 + row * spacingY))
    }
  }

  private fun layoutSeedComponent(seedNodeId: String, component: Set<String>) {
    val levels = linkedMapOf(seedNodeId to 0)
    val queue = ArrayDeque<String>()
    queue += seedNodeId

    while (queue.isNotEmpty()) {
      val current = queue.removeFirst()
      val nextLevel = (levels[current] ?: 0) + 1
      model.adjacency[current].orEmpty().sorted().forEach { neighborId ->
        if (neighborId !in component || neighborId in levels) {
          return@forEach
        }
        levels[neighborId] = nextLevel
        queue += neighborId
      }
    }

    levels.entries
      .groupBy({ it.value }, { it.key })
      .toSortedMap()
      .forEach { (level, nodeIds) ->
        nodeIds.sorted().forEachIndexed { index, nodeId ->
          nodePositions.putIfAbsent(
            nodeId,
            Point(
              140 + level * 280,
              120 + index * 170,
            ),
          )
        }
      }
  }

  private fun installInteraction() {
    val mouseAdapter = object : MouseAdapter() {
      override fun mousePressed(event: MouseEvent) {
        val nodeId = findNodeAt(event.point) ?: return
        val bounds = nodeBounds[nodeId] ?: return
        dragNodeId = nodeId
        dragOffset = Point(event.x - bounds.x, event.y - bounds.y)
      }

      override fun mouseDragged(event: MouseEvent) {
        val nodeId = dragNodeId ?: return
        nodePositions[nodeId] = Point(event.x - dragOffset.x, event.y - dragOffset.y)
        revalidate()
        repaint()
      }

      override fun mouseReleased(event: MouseEvent) {
        dragNodeId = null
      }

      override fun mouseClicked(event: MouseEvent) {
        if (event.clickCount != 2 || event.button != MouseEvent.BUTTON1) {
          return
        }
        val nodeId = findNodeAt(event.point) ?: return
        expandRequestHandler(nodeId)
      }
    }

    addMouseListener(mouseAdapter)
    addMouseMotionListener(mouseAdapter)
  }

  fun markExpanded(nodeId: String) {
    expandedNodeIds += nodeId
    repaint()
  }

  fun revealNeighbors(nodeId: String) {
    val neighborIds = model.adjacency[nodeId].orEmpty()
      .filter { candidate -> !visibleNodeIds.contains(candidate) }
    if (neighborIds.isEmpty()) {
      expandedNodeIds += nodeId
      repaint()
      return
    }

    val basePoint = nodePositions[nodeId] ?: Point(420, 240)
    val radius = 220
    neighborIds.forEachIndexed { index, neighborId ->
      val angle = (Math.PI * 2.0 * index) / neighborIds.size.toDouble().coerceAtLeast(1.0)
      nodePositions.putIfAbsent(
        neighborId,
        Point(
          (basePoint.x + cos(angle) * radius).toInt(),
          (basePoint.y + sin(angle) * radius).toInt(),
        ),
      )
      visibleNodeIds += neighborId
    }
    expandedNodeIds += nodeId
    revalidate()
    repaint()
  }

  fun addOrigin(nodeId: String) {
    if (nodeId !in model.nodeById) {
      return
    }
    originNodeIds += nodeId
    if (nodeId !in visibleNodeIds) {
      visibleNodeIds += nodeId
      if (nodeId !in nodePositions) {
        val index = originNodeIds.toList().indexOf(nodeId).coerceAtLeast(0)
        nodePositions[nodeId] = Point(140 + index * 280, 120 + index * 80)
      }
    }
    revealNeighbors(nodeId)
  }

  fun clearGraphView() {
    originNodeIds.clear()
    visibleNodeIds.clear()
    expandedNodeIds.clear()
    nodePositions.clear()
    nodeBounds.clear()
    revalidate()
    repaint()
  }

  private fun findNodeAt(point: Point): String? {
    return nodeBounds.entries.lastOrNull { (_, bounds) -> bounds.contains(point) }?.key
  }

  private fun createNodeBounds(node: ChannelGraphNode, metrics: FontMetrics): Rectangle {
    val position = nodePositions[node.id] ?: Point(120, 120)
    val lines = buildNodeLines(node, isExpandedNode(node.id))
    val textWidth = lines.maxOfOrNull { line -> metrics.stringWidth(line) } ?: 60
    val collapsed = node.external || !isExpandedNode(node.id)
    val width = max(textWidth + 32, if (collapsed) 120 else 170)
    val height = max(lines.size * metrics.height + 24, if (collapsed) 44 else 72)
    return Rectangle(position.x, position.y, width, height)
  }

  private fun paintNode(
    graphics: Graphics2D,
    node: ChannelGraphNode,
    bounds: Rectangle,
    metrics: FontMetrics,
  ) {
    val expanded = isExpandedNode(node.id)
    val collapsed = node.external || !expanded
    val fill = if (node.external) {
      ColorUtil.mix(background, JBColor.BLUE, 0.18)
    } else {
      ColorUtil.mix(background, JBColor(0x9ED2FF, 0x365A7A), 0.55)
    }
    val border = if (node.external) {
      JBColor(0x6AAEFF, 0x7FBFFF)
    } else {
      JBColor(0xE500FF, 0xF066FF)
    }
    val foregroundColor = JBColor.foreground()
    graphics.color = fill
    if (collapsed) {
      graphics.fillOval(bounds.x, bounds.y, bounds.width, bounds.height)
    } else {
      graphics.fillRoundRect(bounds.x, bounds.y, bounds.width, bounds.height, 24, 24)
    }
    graphics.color = border
    if (collapsed) {
      graphics.drawOval(bounds.x, bounds.y, bounds.width, bounds.height)
    } else {
      graphics.drawRoundRect(bounds.x, bounds.y, bounds.width, bounds.height, 24, 24)
    }

    graphics.color = foregroundColor
    val lines = buildNodeLines(node, expanded)
    val totalHeight = lines.size * metrics.height
    var textY = bounds.y + (bounds.height - totalHeight) / 2 + metrics.ascent
    for (line in lines) {
      val textX = bounds.x + (bounds.width - metrics.stringWidth(line)) / 2
      graphics.drawString(line, textX, textY)
      textY += metrics.height
    }
  }

  private fun paintEdges(graphics: Graphics2D, metrics: FontMetrics) {
    val visibleEdges = model.edges.filter { edge ->
      visibleNodeIds.contains(edge.fromId) && visibleNodeIds.contains(edge.toId)
    }
    val edgeColor = JBColor.BLUE
    val labelBackground = ColorUtil.mix(background, JBColor.WHITE, 0.75)
    visibleEdges.forEach { edge ->
      val fromBounds = nodeBounds[edge.fromId] ?: return@forEach
      val toBounds = nodeBounds[edge.toId] ?: return@forEach
      if (edge.fromId == edge.toId) {
        paintSelfEdge(graphics, edge, fromBounds, metrics, labelBackground, edgeColor)
        return@forEach
      }
      val line = createConnectionLine(fromBounds, toBounds)
      graphics.color = edgeColor
      graphics.draw(line)
      paintArrow(graphics, line)

      val labelX = ((line.x1 + line.x2) / 2.0).toInt()
      val labelY = ((line.y1 + line.y2) / 2.0).toInt()
      val labelWidth = metrics.stringWidth(edge.label)
      graphics.color = labelBackground
      graphics.fillRoundRect(labelX - labelWidth / 2 - 6, labelY - metrics.ascent, labelWidth + 12, metrics.height, 10, 10)
      graphics.color = JBColor.foreground()
      graphics.drawString(edge.label, labelX - labelWidth / 2, labelY)
    }
  }

  private fun paintSelfEdge(
    graphics: Graphics2D,
    edge: ChannelGraphEdge,
    bounds: Rectangle,
    metrics: FontMetrics,
    labelBackground: Color,
    edgeColor: Color,
  ) {
    val startX = bounds.x + bounds.width * 0.62
    val startY = bounds.y + 10.0
    val endX = bounds.x + bounds.width * 0.38
    val endY = bounds.y + 10.0
    val controlOffsetY = max(bounds.height.toDouble(), 56.0)
    val curve = CubicCurve2D.Double(
      startX,
      startY,
      bounds.x + bounds.width + 36.0,
      bounds.y - controlOffsetY * 0.35,
      bounds.x - 36.0,
      bounds.y - controlOffsetY * 0.35,
      endX,
      endY,
    )
    graphics.color = edgeColor
    graphics.draw(curve)
    val arrowLine = Line2D.Float(
      (endX + 10.0).toFloat(),
      (endY - 12.0).toFloat(),
      endX.toFloat(),
      endY.toFloat(),
    )
    paintArrow(graphics, arrowLine)

    val labelX = bounds.x + bounds.width / 2
    val labelY = (bounds.y - controlOffsetY * 0.35).toInt()
    val labelWidth = metrics.stringWidth(edge.label)
    graphics.color = labelBackground
    graphics.fillRoundRect(labelX - labelWidth / 2 - 6, labelY - metrics.ascent, labelWidth + 12, metrics.height, 10, 10)
    graphics.color = JBColor.foreground()
    graphics.drawString(edge.label, labelX - labelWidth / 2, labelY)
  }

  private fun createConnectionLine(fromBounds: Rectangle, toBounds: Rectangle): Line2D.Float {
    val fromCenterX = fromBounds.centerX.toFloat()
    val fromCenterY = fromBounds.centerY.toFloat()
    val toCenterX = toBounds.centerX.toFloat()
    val toCenterY = toBounds.centerY.toFloat()
    val angle = atan2(toCenterY - fromCenterY, toCenterX - fromCenterX)
    val fromRadiusX = fromBounds.width / 2f
    val fromRadiusY = fromBounds.height / 2f
    val toRadiusX = toBounds.width / 2f
    val toRadiusY = toBounds.height / 2f
    val startX = fromCenterX + cos(angle).toFloat() * min(fromRadiusX, fromRadiusY)
    val startY = fromCenterY + sin(angle).toFloat() * min(fromRadiusX, fromRadiusY)
    val endX = toCenterX - cos(angle).toFloat() * min(toRadiusX, toRadiusY)
    val endY = toCenterY - sin(angle).toFloat() * min(toRadiusX, toRadiusY)
    return Line2D.Float(startX, startY, endX, endY)
  }

  private fun paintArrow(graphics: Graphics2D, line: Line2D.Float) {
    val arrowLength = 12.0
    val arrowWidth = 7.0
    val angle = atan2((line.y2 - line.y1).toDouble(), (line.x2 - line.x1).toDouble())
    val x2 = line.x2.toDouble()
    val y2 = line.y2.toDouble()
    val path = Path2D.Double().apply {
      moveTo(x2, y2)
      lineTo(
        x2 - arrowLength * cos(angle) + arrowWidth * sin(angle),
        y2 - arrowLength * sin(angle) - arrowWidth * cos(angle),
      )
      lineTo(
        x2 - arrowLength * cos(angle) - arrowWidth * sin(angle),
        y2 - arrowLength * sin(angle) + arrowWidth * cos(angle),
      )
      closePath()
    }
    graphics.fill(path)
  }

  private fun computeCanvasBounds(): Rectangle {
    if (nodeBounds.isEmpty()) {
      val metrics = getFontMetrics(font)
      model.nodes.filter { visibleNodeIds.contains(it.id) }.forEach { node ->
        nodeBounds[node.id] = createNodeBounds(node, metrics)
      }
    }
    if (nodeBounds.isEmpty()) {
      return Rectangle(0, 0, 900, 700)
    }
    val minX = nodeBounds.values.minOf { it.x }
    val minY = nodeBounds.values.minOf { it.y }
    val maxX = nodeBounds.values.maxOf { it.x + it.width }
    val maxY = nodeBounds.values.maxOf { it.y + it.height }
    return Rectangle(minX, minY, maxX - minX, maxY - minY)
  }

  private fun isExpandedNode(nodeId: String): Boolean = expandedNodeIds.contains(nodeId)

  private fun buildNodeLines(node: ChannelGraphNode, expanded: Boolean): List<String> {
    if (node.external || !expanded) {
      return listOfNotNull(
        node.name,
        node.runtimeValue?.takeIf { it.isNotBlank() },
      )
    }
    val detailLine = buildString {
      node.recordType.takeIf { it.isNotBlank() }?.let {
        append("($it)")
      }
      node.scanValue?.takeIf { it.isNotBlank() }?.let {
        if (isNotBlank()) {
          append(" ")
        }
        append("($it)")
      }
    }
    return listOfNotNull(
      node.name,
      detailLine.takeIf { it.isNotBlank() },
      node.runtimeValue?.takeIf { it.isNotBlank() },
    )
  }
}

private data class ChannelGraphNode(
  val id: String,
  val name: String,
  val recordType: String = "",
  val scanValue: String? = null,
  val external: Boolean = false,
  val channelName: String? = null,
  val runtimeValue: String? = null,
)

private data class ChannelGraphEdge(
  val fromId: String,
  val toId: String,
  val label: String,
)

private data class ChannelGraphModel(
  val sourceLabel: String,
  val seedRecordName: String? = null,
  val message: String? = null,
  val nodes: List<ChannelGraphNode>,
  val edges: List<ChannelGraphEdge>,
  val nodeById: Map<String, ChannelGraphNode>,
  val adjacency: Map<String, Set<String>>,
)

internal data class ChannelGraphSource(
  val label: String,
  val text: String,
  val path: String? = null,
)

private fun buildChannelGraphSourceLabel(sources: List<ChannelGraphSource>): String {
  if (sources.isEmpty()) {
    return "EPICS Channel Graph"
  }
  if (sources.size == 1) {
    return sources.first().label
  }
  return "${sources.first().label} + ${sources.size - 1} more"
}

private fun getChannelGraphSourceKey(source: ChannelGraphSource): String {
  return source.path ?: "${source.label}\u0000${source.text}"
}

private fun isStaticDatabaseSource(source: ChannelGraphSource): Boolean {
  val extension = source.path?.substringAfterLast('.', "")?.lowercase()
  return extension in setOf("db", "vdb", "template") || source.text.isNotBlank()
}

private fun isDatabaseGraphSource(file: VirtualFile): Boolean {
  return file.extension?.lowercase() in setOf("db", "vdb", "template")
}

private val CHANNEL_GRAPH_LINK_FIELD_TYPES = setOf("DBF_INLINK", "DBF_OUTLINK", "DBF_FWDLINK")

private fun buildEmptyDynamicModel(message: String): ChannelGraphModel {
  return ChannelGraphModel(
    sourceLabel = "Runtime",
    seedRecordName = null,
    message = message,
    nodes = emptyList(),
    edges = emptyList(),
    nodeById = emptyMap(),
    adjacency = emptyMap(),
  )
}

private fun buildInitialDynamicModel(
  originNodeIds: List<String>,
  protocol: MonitorProtocol,
  sourceName: String,
): ChannelGraphModel {
  val nodes = originNodeIds.map { originNodeId ->
    ChannelGraphNode(
      id = originNodeId,
      name = originNodeId,
      external = false,
      channelName = originNodeId,
      runtimeValue = "(connecting...)",
    )
  }
  val nodeById = linkedMapOf<String, ChannelGraphNode>().apply {
    nodes.forEach { put(it.id, it) }
  }
  return ChannelGraphModel(
    sourceLabel = "$sourceName [runtime ${protocol.name.lowercase(Locale.US)}]",
    seedRecordName = originNodeIds.firstOrNull(),
    message = if (originNodeIds.isEmpty()) "Enter a channel name to start the Channel Graph." else null,
    nodes = nodes,
    edges = emptyList(),
    nodeById = nodeById,
    adjacency = linkedMapOf<String, MutableSet<String>>().apply {
      originNodeIds.forEach { put(it, linkedSetOf()) }
    },
  )
}

private fun parseLinkTarget(value: String): ParsedLinkTarget? {
  val trimmed = value.trim()
  if (trimmed.isBlank()) {
    return null
  }
  if (trimmed.startsWith("@")) {
    return ParsedLinkTarget(rawValue = trimmed.trim('"', '\''))
  }
  val token = trimmed.split(Regex("""[\s,]+""")).firstOrNull()?.trim()?.trim('"', '\'').orEmpty()
  if (token.isBlank()) {
    return null
  }
  if (token == "0" || token == "1") {
    return ParsedLinkTarget(rawValue = token)
  }
  val recordToken = token.substringBefore('.')
  if (recordToken.isBlank()) {
    return ParsedLinkTarget(rawValue = trimmed.trim('"', '\''))
  }
  return ParsedLinkTarget(
    recordName = recordToken,
    targetField = token.substringAfter('.', "").takeIf { it.isNotBlank() }?.uppercase(),
    rawValue = trimmed.trim('"', '\''),
  )
}

private fun looksLikeLinkField(fieldName: String): Boolean {
  return isInputLikeField(fieldName) ||
    fieldName == "OUT" ||
    fieldName == "FLNK" ||
    fieldName == "FWDLINK" ||
    fieldName == "SELL" ||
    fieldName == "DOL" ||
    fieldName == "SDIS" ||
    fieldName == "SIOL" ||
    fieldName == "TSEL" ||
    Regex("""^OUT[A-U]$""").matches(fieldName) ||
    Regex("""^LNK[0-9A-F]$""").matches(fieldName)
}

private fun isInputLikeField(fieldName: String): Boolean {
  return fieldName == "INP" ||
    Regex("""^INP[A-U]$""").matches(fieldName) ||
    Regex("""^DOL[0-9A-F]$""").matches(fieldName)
}

private fun isInputLikeLink(dbfType: String?, fieldName: String): Boolean {
  return dbfType == "DBF_INLINK" || (dbfType !in CHANNEL_GRAPH_LINK_FIELD_TYPES && isInputLikeField(fieldName))
}

private data class ParsedLinkTarget(
  val recordName: String? = null,
  val targetField: String? = null,
  val rawValue: String? = null,
)

private object ChannelGraphSupport {
  fun build(sourceText: String, sourceLabel: String, seedRecordName: String?): ChannelGraphModel {
    if (sourceText.isBlank()) {
      return ChannelGraphModel(
        sourceLabel = sourceLabel,
        seedRecordName = seedRecordName,
        message = "No database source is available for Channel Graph.",
        nodes = emptyList(),
        edges = emptyList(),
        nodeById = emptyMap(),
        adjacency = emptyMap(),
      )
    }

    val declarations = EpicsRecordCompletionSupport.extractRecordDeclarations(sourceText)
    if (declarations.isEmpty()) {
      return ChannelGraphModel(
        sourceLabel = sourceLabel,
        seedRecordName = seedRecordName,
        message = "No record declarations were found in $sourceLabel.",
        nodes = emptyList(),
        edges = emptyList(),
        nodeById = emptyMap(),
        adjacency = emptyMap(),
      )
    }

    val scanByRecord = mutableMapOf<String, String>()
    declarations.forEach { declaration ->
      val scanField = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(sourceText, declaration)
        .firstOrNull { field -> field.fieldName == "SCAN" }
      scanByRecord[declaration.name] = scanField?.value.orEmpty()
    }

    val nodesById = LinkedHashMap<String, ChannelGraphNode>()
    declarations.forEach { declaration ->
      nodesById[declaration.name] = ChannelGraphNode(
        id = declaration.name,
        name = declaration.name,
        recordType = declaration.recordType,
        scanValue = scanByRecord[declaration.name],
        external = false,
        channelName = declaration.name,
      )
    }

    val edges = linkedSetOf<ChannelGraphEdge>()
    declarations.forEach { declaration ->
      val fields = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(sourceText, declaration)
      fields.forEach fieldLoop@{ field ->
        val dbfType = EpicsRecordCompletionSupport.getFieldType(declaration.recordType, field.fieldName)
        if (dbfType !in CHANNEL_GRAPH_LINK_FIELD_TYPES && !looksLikeLinkField(field.fieldName)) {
          return@fieldLoop
        }
        val target = parseLinkTarget(field.value) ?: return@fieldLoop
        val targetNodeId = target.recordName ?: "__value__:${declaration.name}:${field.fieldName}:${target.rawValue}"
        val targetNodeName = target.recordName ?: target.rawValue ?: return@fieldLoop
        if (!nodesById.containsKey(targetNodeId)) {
          nodesById[targetNodeId] = ChannelGraphNode(
            id = targetNodeId,
            name = targetNodeName,
            external = true,
            channelName = target.recordName,
          )
        }
        val label = if (target.targetField != null) "${field.fieldName}:${target.targetField}" else field.fieldName
        val edge = when (dbfType) {
          "DBF_INLINK" -> ChannelGraphEdge(fromId = targetNodeId, toId = declaration.name, label = label)
          "DBF_OUTLINK", "DBF_FWDLINK" -> ChannelGraphEdge(fromId = declaration.name, toId = targetNodeId, label = label)
          else -> if (isInputLikeField(field.fieldName)) {
            ChannelGraphEdge(fromId = targetNodeId, toId = declaration.name, label = label)
          } else {
            ChannelGraphEdge(fromId = declaration.name, toId = targetNodeId, label = label)
          }
        }
        edges += edge
      }
    }

    val adjacency = linkedMapOf<String, MutableSet<String>>()
    edges.forEach { edge ->
      adjacency.getOrPut(edge.fromId) { linkedSetOf() } += edge.toId
      adjacency.getOrPut(edge.toId) { linkedSetOf() } += edge.fromId
    }

    val message = when {
      seedRecordName != null && seedRecordName !in nodesById -> "Seed record \"$seedRecordName\" was not found in $sourceLabel. Showing the full graph."
      edges.isEmpty() -> "No link relationships were found in $sourceLabel."
      else -> null
    }

    return ChannelGraphModel(
      sourceLabel = sourceLabel,
      seedRecordName = seedRecordName,
      message = message,
      nodes = nodesById.values.toList(),
      edges = edges.toList(),
      nodeById = nodesById,
      adjacency = adjacency,
    )
  }
}

private class ChannelGraphRuntimeSession(
  project: Project,
  initialOriginNodeIds: List<String>,
  private val protocol: MonitorProtocol,
  private val sourceName: String,
  private val onModelChanged: (ChannelGraphModel) -> Unit,
) : AutoCloseable {
  private val disposed = AtomicBoolean(false)
  private val configuration = project.service<EpicsRuntimeProjectConfigurationService>().loadConfiguration()
  private val caContext: Context? = if (protocol == MonitorProtocol.CA) EpicsClientLibraries.createCaContext(configuration) else null
  private val pvaClient: PVAClient? = if (protocol == MonitorProtocol.PVA) EpicsClientLibraries.createPvaClient(configuration) else null
  private val nodesById = LinkedHashMap<String, ChannelGraphNode>()
  private val edges = linkedSetOf<ChannelGraphEdge>()
  private val expandedNodes = linkedSetOf<String>()
  private val originNodeIds = linkedSetOf<String>().apply {
    initialOriginNodeIds.map(String::trim).filter(String::isNotBlank).forEach(::add)
  }
  private val nodeSessions = LinkedHashMap<String, ChannelGraphNodeRuntimeSession>()
  private val lock = Any()

  init {
    originNodeIds.forEach { originNodeId ->
      nodesById[originNodeId] = ChannelGraphNode(
        id = originNodeId,
        name = originNodeId,
        external = false,
        channelName = originNodeId,
        runtimeValue = CONNECTING_DISPLAY,
      )
    }
  }

  fun start() {
    publish()
    originNodeIds.forEach { originNodeId ->
      ensureNodeSession(originNodeId)
      requestExpand(originNodeId)
    }
  }

  fun getOriginNodeIds(): Set<String> = synchronized(lock) { LinkedHashSet(originNodeIds) }

  fun addOriginNode(nodeId: String) {
    val normalizedNodeId = nodeId.trim()
    if (normalizedNodeId.isBlank()) {
      return
    }
    synchronized(lock) {
      originNodeIds += normalizedNodeId
      nodesById.putIfAbsent(
        normalizedNodeId,
        ChannelGraphNode(
          id = normalizedNodeId,
          name = normalizedNodeId,
          external = false,
          channelName = normalizedNodeId,
          runtimeValue = CONNECTING_DISPLAY,
        ),
      )
    }
    publish()
    ensureNodeSession(normalizedNodeId)
    requestExpand(normalizedNodeId)
  }

  fun clearGraph() {
    synchronized(lock) {
      nodeSessions.values.toList().forEach(ChannelGraphNodeRuntimeSession::close)
      nodeSessions.clear()
      edges.clear()
      expandedNodes.clear()
      originNodeIds.clear()
      nodesById.clear()
    }
    publish()
  }

  fun requestExpand(nodeId: String) {
    val channelName = synchronized(lock) {
      expandedNodes += nodeId
      nodesById[nodeId]?.channelName
    } ?: run {
      publish()
      return
    }

    publish()
    ApplicationManager.getApplication().executeOnPooledThread {
      val expansion = runCatching { resolveRuntimeExpansion(nodeId, channelName) }.getOrNull() ?: return@executeOnPooledThread
      val channelsToConnect = mutableListOf<String>()
      synchronized(lock) {
        val currentNode = nodesById[nodeId]
        if (currentNode != null) {
          nodesById[nodeId] = currentNode.copy(
            recordType = expansion.recordType ?: currentNode.recordType,
            scanValue = expansion.scanValue ?: currentNode.scanValue,
          )
        }
        expansion.targets.forEach { target ->
          val existingNode = nodesById[target.node.id]
          if (existingNode == null) {
            nodesById[target.node.id] = target.node
          } else {
            nodesById[target.node.id] = existingNode.copy(
              external = existingNode.external && target.node.external,
              channelName = existingNode.channelName ?: target.node.channelName,
            )
          }
          edges += target.edge
          target.node.channelName?.let(channelsToConnect::add)
        }
      }
      channelsToConnect.forEach(::ensureNodeSession)
      publish()
    }
  }

  override fun close() {
    if (!disposed.compareAndSet(false, true)) {
      return
    }
    synchronized(lock) {
      nodeSessions.values.toList().forEach(ChannelGraphNodeRuntimeSession::close)
      nodeSessions.clear()
    }
    caContext?.let { context -> runCatching { context.destroy() } }
    pvaClient?.let { client -> runCatching { client.close() } }
  }

  private fun ensureNodeSession(nodeId: String) {
    val node = synchronized(lock) {
      if (nodeId in nodeSessions) {
        return
      }
      nodesById[nodeId]
    } ?: return
    val channelName = node.channelName ?: return

    val session = when (protocol) {
      MonitorProtocol.CA -> ChannelGraphCaNodeRuntimeSession(
        caContext = caContext ?: return,
        channelName = channelName,
        onValue = { value -> updateNodeRuntime(nodeId, value, connected = true) },
        onConnectFailure = { updateNodeRuntime(nodeId, null, connected = false) },
        onDisconnect = { updateNodeRuntime(nodeId, CONNECTING_DISPLAY, connected = true) },
      )

      MonitorProtocol.PVA -> ChannelGraphPvaNodeRuntimeSession(
        pvaClient = pvaClient ?: return,
        channelName = channelName,
        onValue = { value -> updateNodeRuntime(nodeId, value, connected = true) },
        onConnectFailure = { updateNodeRuntime(nodeId, null, connected = false) },
        onDisconnect = { updateNodeRuntime(nodeId, CONNECTING_DISPLAY, connected = true) },
      )
    }

    synchronized(lock) {
      if (disposed.get()) {
        session.close()
        return
      }
      nodeSessions[nodeId] = session
    }
    session.start()
  }

  private fun updateNodeRuntime(
    nodeId: String,
    value: String?,
    connected: Boolean,
  ) {
    synchronized(lock) {
      val existingNode = nodesById[nodeId] ?: return
      nodesById[nodeId] = existingNode.copy(
        external = !connected,
        runtimeValue = value,
      )
    }
    publish()
  }

  private fun resolveRuntimeExpansion(
    nodeId: String,
    channelName: String,
  ): RuntimeExpansion {
    val recordType = runCatching { readRuntimeText("$channelName.RTYP") }.getOrNull().orEmpty().trim()
    val normalizedRecordType = recordType.takeIf { it.isNotBlank() }
    val scanValue = runCatching { readRuntimeText("$channelName.SCAN") }.getOrNull()
      ?.takeIf { it.isNotBlank() }
      ?.let(::simplifyRuntimeScanValue)
    val fieldNames = EpicsRecordCompletionSupport.getFieldNamesForRecordType(normalizedRecordType)
      .filter { fieldName ->
        val dbfType = normalizedRecordType?.let { EpicsRecordCompletionSupport.getFieldType(it, fieldName) }
        dbfType in CHANNEL_GRAPH_LINK_FIELD_TYPES || looksLikeLinkField(fieldName)
      }

    val targets = mutableListOf<RuntimeExpansionTarget>()
    fieldNames.forEach { fieldName ->
      val rawValue = runCatching { readRuntimeText("$channelName.$fieldName") }.getOrNull().orEmpty().trim()
      if (rawValue.isBlank()) {
        return@forEach
      }
      val dbfType = normalizedRecordType?.let { EpicsRecordCompletionSupport.getFieldType(it, fieldName) }
      val parsedTarget = parseLinkTarget(rawValue) ?: return@forEach
      val node = if (parsedTarget.recordName != null) {
        ChannelGraphNode(
          id = parsedTarget.recordName,
          name = parsedTarget.recordName,
          external = false,
          channelName = parsedTarget.recordName,
          runtimeValue = CONNECTING_DISPLAY,
        )
      } else {
        ChannelGraphNode(
          id = "__value__:$nodeId:$fieldName:${parsedTarget.rawValue}",
          name = parsedTarget.rawValue ?: rawValue,
          external = true,
          channelName = null,
        )
      }
      val label = parsedTarget.targetField?.let { "$fieldName:$it" } ?: fieldName
      val edge = if (isInputLikeLink(dbfType, fieldName)) {
        ChannelGraphEdge(fromId = node.id, toId = nodeId, label = label)
      } else {
        ChannelGraphEdge(fromId = nodeId, toId = node.id, label = label)
      }
      targets += RuntimeExpansionTarget(node, edge)
    }

    return RuntimeExpansion(
      recordType = normalizedRecordType,
      scanValue = scanValue,
      targets = targets,
    )
  }

  private fun readRuntimeText(pvName: String): String? {
    return when (protocol) {
      MonitorProtocol.CA -> caContext?.let { readCaText(it, pvName) }
      MonitorProtocol.PVA -> pvaClient?.let { readPvaText(it, pvName) }
    }
  }

  private fun publish() {
    if (disposed.get()) {
      return
    }
    val model = synchronized(lock) {
      val edgeList = edges.toList()
      val adjacency = linkedMapOf<String, MutableSet<String>>()
      edgeList.forEach { edge ->
        adjacency.getOrPut(edge.fromId) { linkedSetOf() } += edge.toId
        adjacency.getOrPut(edge.toId) { linkedSetOf() } += edge.fromId
      }
      ChannelGraphModel(
        sourceLabel = "$sourceName [runtime ${protocol.name.lowercase(Locale.US)}]",
        seedRecordName = originNodeIds.firstOrNull(),
        message = if (originNodeIds.isEmpty()) "Enter a channel name to start the Channel Graph." else null,
        nodes = nodesById.values.toList(),
        edges = edgeList,
        nodeById = LinkedHashMap(nodesById),
        adjacency = adjacency,
      )
    }
    onModelChanged(model)
  }

  private data class RuntimeExpansion(
    val recordType: String?,
    val scanValue: String?,
    val targets: List<RuntimeExpansionTarget>,
  )

  private data class RuntimeExpansionTarget(
    val node: ChannelGraphNode,
    val edge: ChannelGraphEdge,
  )

  private companion object {
    private const val CONNECTING_DISPLAY = "(connecting...)"
  }
}

private interface ChannelGraphNodeRuntimeSession : AutoCloseable {
  fun start()
}

private class ChannelGraphCaNodeRuntimeSession(
  private val caContext: Context,
  private val channelName: String,
  private val onValue: (String) -> Unit,
  private val onConnectFailure: () -> Unit,
  private val onDisconnect: () -> Unit,
) : ChannelGraphNodeRuntimeSession {
  private val disposed = AtomicBoolean(false)
  private var channel: Channel? = null
  private var monitor: Monitor? = null

  override fun start() {
    ApplicationManager.getApplication().executeOnPooledThread {
      runCatching {
        caContext.attachCurrentThread()
        val nextChannel = connectCaChannel(caContext, channelName) {
          if (!disposed.get()) {
            onDisconnect()
          }
        }
        val fieldType = nextChannel.fieldType
        val labels = if (fieldType == DBRType.ENUM) fetchCaEnumLabels(caContext, nextChannel) else emptyList()
        val initialDbr = nextChannel.get(getCaGraphReadType(fieldType), nextChannel.elementCount.coerceAtLeast(1))
        onValue(formatCaGraphValue(initialDbr, labels))
        val nextMonitor = nextChannel.addMonitor(
          getCaGraphReadType(fieldType),
          nextChannel.elementCount.coerceAtLeast(1),
          Monitor.VALUE or Monitor.ALARM,
          MonitorListener { event: MonitorEvent ->
            if (disposed.get()) {
              return@MonitorListener
            }
            if (!event.status.isSuccessful) {
              onDisconnect()
              return@MonitorListener
            }
            onValue(formatCaGraphValue(event.dbr, labels))
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
      }.onFailure {
        if (!disposed.get()) {
          onConnectFailure()
        }
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

private class ChannelGraphPvaNodeRuntimeSession(
  private val pvaClient: PVAClient,
  private val channelName: String,
  private val onValue: (String) -> Unit,
  private val onConnectFailure: () -> Unit,
  private val onDisconnect: () -> Unit,
) : ChannelGraphNodeRuntimeSession {
  private val disposed = AtomicBoolean(false)
  private var channel: PVAChannel? = null
  private var subscription: AutoCloseable? = null

  override fun start() {
    ApplicationManager.getApplication().executeOnPooledThread {
      runCatching {
        val nextChannel = pvaClient.getChannel(channelName) { _: PVAChannel, channelState: ClientChannelState ->
          if (!disposed.get() && channelState != ClientChannelState.CONNECTED) {
            onDisconnect()
          }
        }
        nextChannel.connect().get(PVA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
        val initial = nextChannel.read("").get(PVA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
        onValue(formatPvaGraphStructure(initial))
        val nextSubscription = nextChannel.subscribe(
          "",
          PvaMonitorListener { _: PVAChannel, _, _, structure: PVAStructure ->
            if (!disposed.get()) {
              onValue(formatPvaGraphStructure(structure))
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
      }.onFailure {
        if (!disposed.get()) {
          onConnectFailure()
        }
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

private fun simplifyRuntimeScanValue(value: String): String {
  return Regex("""^\[(\d+)]\s*(.+)$""").matchEntire(value)?.groups?.get(2)?.value ?: value
}

private fun connectCaChannel(
  caContext: Context,
  pvName: String,
  onDisconnect: () -> Unit = {},
): Channel {
  caContext.attachCurrentThread()
  val latch = CountDownLatch(1)
  var connected = false
  val channel = caContext.createChannel(
    pvName,
    ConnectionListener { event: ConnectionEvent ->
      if (event.isConnected) {
        connected = true
        latch.countDown()
      } else {
        onDisconnect()
      }
    },
  )
  caContext.flushIO()
  if (!latch.await(CA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS) || !connected) {
    runCatching { channel.destroy() }
    throw TimeoutException("Timed out connecting to $pvName")
  }
  return channel
}

private fun readCaText(
  caContext: Context,
  pvName: String,
): String? {
  val channel = connectCaChannel(caContext, pvName)
  return try {
    val fieldType = channel.fieldType
    val labels = if (fieldType == DBRType.ENUM) fetchCaEnumLabels(caContext, channel) else emptyList()
    val dbr = channel.get(getCaGraphReadType(fieldType), channel.elementCount.coerceAtLeast(1))
    formatCaGraphValue(dbr, labels)
  } finally {
    runCatching { channel.destroy() }
  }
}

private fun readPvaText(
  pvaClient: PVAClient,
  pvName: String,
): String? {
  val channel = pvaClient.getChannel(pvName)
  return try {
    channel.connect().get(PVA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
    val structure = channel.read("").get(PVA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
    formatPvaGraphStructure(structure)
  } finally {
    runCatching { channel.close() }
  }
}

private fun getCaGraphReadType(fieldType: DBRType): DBRType {
  return when (fieldType) {
    DBRType.STRING -> DBRType.STRING
    DBRType.SHORT -> DBRType.SHORT
    DBRType.FLOAT -> DBRType.FLOAT
    DBRType.ENUM -> DBRType.LABELS_ENUM
    DBRType.BYTE -> DBRType.BYTE
    DBRType.INT -> DBRType.INT
    DBRType.DOUBLE -> DBRType.DOUBLE
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

private fun formatCaGraphValue(dbr: DBR?, fallbackLabels: List<String>): String {
  if (dbr == null) {
    return ""
  }
  if (dbr is DBR_Enum) {
    val index = dbr.enumValue.firstOrNull()?.toInt() ?: 0
    val labels = when (dbr) {
      is LABELS -> dbr.labels?.map { it ?: "" }.orEmpty()
      else -> fallbackLabels
    }
    return formatEnumValue(index, labels)
  }
  val value = dbr.value ?: return ""
  return formatRuntimeGraphObject(value)
}

private fun formatRuntimeGraphObject(value: Any): String {
  return when (value) {
    is String -> value
    is Array<*> -> formatRuntimeGraphArray(value.map { it?.toString().orEmpty() })
    is ByteArray -> formatRuntimeGraphArray(value.map(Byte::toString))
    is ShortArray -> formatRuntimeGraphArray(value.map(Short::toString))
    is IntArray -> formatRuntimeGraphArray(value.map(Int::toString))
    is LongArray -> formatRuntimeGraphArray(value.map(Long::toString))
    is FloatArray -> formatRuntimeGraphArray(value.map(Float::toString))
    is DoubleArray -> formatRuntimeGraphArray(value.map(Double::toString))
    else -> value.toString()
  }
}

private fun formatRuntimeGraphArray(values: List<String>): String {
  if (values.isEmpty()) {
    return "[]"
  }
  return if (values.size == 1) values.first() else values.joinToString(", ", prefix = "[", postfix = "]")
}

private fun formatPvaGraphStructure(structure: PVAStructure?): String {
  if (structure == null) {
    return ""
  }
  val valueField = structure.get<PVAData>("value") ?: return if (structure.get().isNotEmpty()) "Has data, but no value" else ""
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
  return if (valueField is PVAValue) {
    valueField.formatGraphDisplayValue()
  } else {
    valueField.toString()
  }
}

private fun formatEnumValue(index: Int, choices: List<String>): String {
  val choice = choices.getOrNull(index).orEmpty()
  return "[$index] $choice".trimEnd()
}

private fun PVAValue.formatGraphDisplayValue(): String {
  return when (this) {
    is PVAString -> get().orEmpty()
    is PVAShort -> get().toString()
    is PVAInt -> get().toString()
    is PVALong -> get().toString()
    is PVAFloat -> get().toString()
    is PVADouble -> get().toString()
    is PVAByte -> get().toString()
    else -> toString()
  }
}

private const val CA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS: Long = 3000
private const val PVA_CHANNEL_GRAPH_CONNECT_TIMEOUT_MS: Long = 3000
