package org.epics.workbench.runtime

import com.intellij.openapi.components.Service
import com.intellij.openapi.components.service
import com.intellij.openapi.Disposable
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.editor.Document
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.editor.EditorCustomElementRenderer
import com.intellij.openapi.editor.EditorFactory
import com.intellij.openapi.editor.colors.EditorFontType
import com.intellij.openapi.editor.ex.EditorEx
import com.intellij.openapi.editor.Inlay
import com.intellij.openapi.editor.event.DocumentEvent
import com.intellij.openapi.editor.event.DocumentListener
import com.intellij.openapi.editor.event.EditorFactoryEvent
import com.intellij.openapi.editor.event.EditorFactoryListener
import com.intellij.openapi.editor.event.EditorMouseEvent
import com.intellij.openapi.editor.event.EditorMouseEventArea
import com.intellij.openapi.editor.event.EditorMouseListener
import com.intellij.openapi.editor.impl.EditorEmbeddedComponentManager
import com.intellij.openapi.editor.markup.CustomHighlighterRenderer
import com.intellij.openapi.editor.markup.HighlighterLayer
import com.intellij.openapi.editor.markup.HighlighterTargetArea
import com.intellij.openapi.editor.markup.RangeHighlighter
import com.intellij.openapi.editor.markup.TextAttributes
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorManagerEvent
import com.intellij.openapi.fileEditor.FileEditorManagerListener
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.startup.ProjectActivity
import com.intellij.openapi.util.Key
import com.intellij.ui.JBColor
import com.intellij.util.concurrency.AppExecutorUtil
import com.intellij.util.messages.Topic
import gov.aps.jca.Channel
import gov.aps.jca.Context
import gov.aps.jca.Monitor
import gov.aps.jca.TimeoutException
import gov.aps.jca.dbr.DBR
import gov.aps.jca.dbr.DBRType
import gov.aps.jca.dbr.DBR_Enum
import gov.aps.jca.dbr.LABELS
import gov.aps.jca.event.ConnectionEvent
import gov.aps.jca.event.ConnectionListener
import gov.aps.jca.event.MonitorEvent
import gov.aps.jca.event.MonitorListener
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
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetPlan
import org.epics.workbench.pvlist.EpicsPvlistWidgetSupport
import org.epics.workbench.probe.EpicsProbeDocumentAnalysis
import org.epics.workbench.toc.EpicsDatabaseToc
import java.awt.Font
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.Rectangle
import java.awt.event.ComponentAdapter
import java.awt.event.ComponentEvent
import java.awt.geom.Rectangle2D
import java.util.concurrent.ConcurrentHashMap
import java.util.concurrent.CountDownLatch
import java.util.concurrent.TimeUnit
import java.util.concurrent.atomic.AtomicBoolean

class EpicsMonitorRuntimeActivity : ProjectActivity {
  override suspend fun execute(project: Project) {
    project.service<EpicsMonitorRuntimeService>().initialize()
  }
}

private class ProbeEditorOverlay(
  private val editor: EditorEx,
  private val panel: EpicsProbeViewPanel,
  private val inlay: Disposable,
  private val resizeListener: ComponentAdapter,
) : Disposable {
  override fun dispose() {
    editor.contentComponent.removeComponentListener(resizeListener)
    panel.dispose()
    inlay.dispose()
  }
}

interface EpicsMonitorRuntimeStateListener {
  fun monitoringStateChanged(active: Boolean)

  companion object {
    val TOPIC: Topic<EpicsMonitorRuntimeStateListener> = Topic.create(
      "EPICS monitor runtime state",
      EpicsMonitorRuntimeStateListener::class.java,
    )
  }
}

@Service(Service.Level.PROJECT)
class EpicsMonitorRuntimeService(
  internal val project: Project,
) : DocumentListener, EditorFactoryListener, EditorMouseListener, Disposable {
  @Volatile
  private var initialized = false

  @Volatile
  private var monitoringActive = false

  @Volatile
  private var caContext: Context? = null

  @Volatile
  private var pvaClient: PVAClient? = null

  @Volatile
  internal var defaultProtocol: MonitorProtocol = MonitorProtocol.CA

  @Volatile
  private var activeProbeSession: EpicsProbeRuntimeSession? = null

  private val widgetProbeSessions = ConcurrentHashMap<String, EpicsProbeRuntimeSession>()
  private val widgetPvlistSessions = ConcurrentHashMap<String, EpicsPvlistWidgetSession>()

  fun initialize() {
    if (initialized) {
      return
    }
    initialized = true
    val multicaster = EditorFactory.getInstance().eventMulticaster
    multicaster.addDocumentListener(this, this)
    EditorFactory.getInstance().addEditorFactoryListener(this, this)
    multicaster.addEditorMouseListener(this, this)
    project.messageBus.connect(this).subscribe(
      FileEditorManagerListener.FILE_EDITOR_MANAGER,
      object : FileEditorManagerListener {
        override fun selectionChanged(event: FileEditorManagerEvent) {
          refreshProbeOverlaysForSelection()
        }
      },
    )
    refreshOpenEditors()
    refreshProbeOverlaysForSelection()
  }

  fun isMonitoringActive(): Boolean = monitoringActive

  fun startMonitoring() {
    if (project.isDisposed || monitoringActive) {
      return
    }
    val runtimeConfiguration = project.service<EpicsRuntimeProjectConfigurationService>().loadConfiguration()
    defaultProtocol = runtimeConfiguration.protocol.toMonitorProtocol()
    caContext = EpicsClientLibraries.createCaContext(runtimeConfiguration)
    pvaClient = EpicsClientLibraries.createPvaClient(runtimeConfiguration)
    monitoringActive = true
    refreshOpenEditors()
    refreshProbeOverlaysForSelection()
    publishStateChanged()
  }

  fun stopMonitoring() {
    if (!monitoringActive && caContext == null && pvaClient == null) {
      return
    }
    monitoringActive = false
    disposeActiveProbeSession()
    disposeAllWidgetProbeSessions()
    disposeAllWidgetPvlistSessions()
    disposeAllSessions()
    caContext?.let { context -> runCatching { context.destroy() } }
    pvaClient?.let { client -> runCatching { client.close() } }
    caContext = null
    pvaClient = null
    refreshProbeOverlaysForSelection()
    publishStateChanged()
  }

  fun restartMonitoringIfActive() {
    if (!monitoringActive) {
      return
    }
    stopMonitoring()
    startMonitoring()
  }

  fun toggleMonitoring() {
    if (monitoringActive) {
      stopMonitoring()
    } else {
      startMonitoring()
    }
  }

  override fun documentChanged(event: DocumentEvent) {
    updateDocument(event.document)
  }

  override fun editorCreated(event: EditorFactoryEvent) {
    updateDocument(event.editor.document)
  }

  override fun editorReleased(event: EditorFactoryEvent) {
    disposeSession(event.editor)
  }

  override fun mouseClicked(event: EditorMouseEvent) {
    if (
      !monitoringActive ||
      event.isConsumed ||
      event.mouseEvent.clickCount != 2 ||
      event.area != EditorMouseEventArea.EDITING_AREA
    ) {
      return
    }

    val session = event.editor.getUserData(SESSION_KEY) as? RuntimeEditorSession ?: return
    if (session.handleDoubleClick(event)) {
      event.consume()
    }
  }

  private fun refreshOpenEditors() {
    for (editor in EditorFactory.getInstance().allEditors) {
      if (editor.project == project) {
        updateDocument(editor.document)
      }
    }
  }

  private fun updateDocument(document: Document) {
    val file = FileDocumentManager.getInstance().getFile(document)
    for (editor in EditorFactory.getInstance().getEditors(document, project)) {
      when {
        file == null -> {
          disposeSession(editor)
          disposeProbeOverlay(editor)
        }
        isProbeFileName(file.name) -> {
          disposeSession(editor)
        }
        !isRuntimeFile(file.name) || !monitoringActive -> {
          disposeSession(editor)
          disposeProbeOverlay(editor)
        }
        else -> {
          disposeProbeOverlay(editor)
          rebuildSession(editor, file.name, document.text)
        }
      }
    }
    if (file == null || isProbeFileName(file.name)) {
      refreshProbeOverlaysForSelection()
    }
  }

  private fun rebuildSession(editor: Editor, fileName: String, text: String) {
    disposeSession(editor)
    if (!monitoringActive) {
      return
    }
    val states = when {
      isMonitorFile(fileName) -> {
        val definition = parseMonitorDocument(text, defaultProtocol)
        definition.entries.map { entry ->
          MonitorLineState(
            editor = editor,
            entry = entry,
            alignColumn = definition.maxDisplayWidth,
          )
        }
      }
      isDatabaseFile(fileName) -> parseDatabaseTocStates(editor, text, defaultProtocol)
      else -> emptyList()
    }
    if (states.isEmpty()) {
      return
    }

    val session = RuntimeEditorSession(
      project = project,
      caContext = caContext ?: return,
      pvaClient = pvaClient ?: return,
      states = states,
    )
    editor.putUserData(SESSION_KEY, session)
    session.start()
  }

  private fun disposeSession(editor: Editor) {
    editor.getUserData(SESSION_KEY)?.dispose()
    editor.putUserData(SESSION_KEY, null)
  }

  private fun rebuildProbeOverlay(editor: Editor) {
    disposeProbeOverlay(editor)
    val context = getProbeContext(editor) ?: return
    val offset = context.analysis.overlayOffset ?: return
    val editorEx = editor as? EditorEx ?: return
    val panel = EpicsProbeViewPanel(
      stateProvider = { getProbeViewState(editor) },
      putHandler = { key -> requestPutProbeValue(editor, key) },
      isMonitoringActive = { isMonitoringActive() },
      startHandler = { startMonitoring() },
      stopHandler = { stopMonitoring() },
      processHandler = { requestProcessProbe(editor) },
    )
    panel.updatePreferredSize(editorEx)
    val properties = EditorEmbeddedComponentManager.Properties(
      EditorEmbeddedComponentManager.ResizePolicy.any(),
      null,
      false,
      false,
      true,
      true,
      0,
      offset,
    )
    val inlay = EditorEmbeddedComponentManager.getInstance().addComponent(editorEx, panel, properties)
      ?: run {
        panel.dispose()
        return
      }
    val resizeListener =
      object : ComponentAdapter() {
        override fun componentResized(event: ComponentEvent) {
          panel.updatePreferredSize(editorEx)
        }
      }
    editorEx.contentComponent.addComponentListener(resizeListener)
    editor.putUserData(PROBE_OVERLAY_KEY, ProbeEditorOverlay(editorEx, panel, inlay, resizeListener))
    panel.refreshFromService()
  }

  private fun disposeProbeOverlay(editor: Editor) {
    editor.getUserData(PROBE_OVERLAY_KEY)?.dispose()
    editor.putUserData(PROBE_OVERLAY_KEY, null)
  }

  private fun refreshProbeOverlaysForSelection() {
    val selectedEditor = FileEditorManager.getInstance(project).selectedTextEditor
    for (editor in EditorFactory.getInstance().allEditors) {
      if (editor.project != project) {
        continue
      }
      val file = FileDocumentManager.getInstance().getFile(editor.document)
      val shouldShow =
        editor == selectedEditor &&
          file != null &&
          isProbeFileName(file.name)
      if (shouldShow) {
        rebuildProbeOverlay(editor)
      } else {
        disposeProbeOverlay(editor)
      }
    }
  }

  private fun disposeAllSessions() {
    for (editor in EditorFactory.getInstance().allEditors) {
      if (editor.project == project) {
        disposeSession(editor)
        disposeProbeOverlay(editor)
      }
    }
  }

  internal fun getActiveProbeViewState(): EpicsProbeViewState? {
    val editor = FileEditorManager.getInstance(project).selectedTextEditor ?: return null
    return getProbeViewState(editor)
  }

  internal fun getProbeViewState(editor: Editor): EpicsProbeViewState? {
    val context = getProbeContext(editor)
    if (context == null) {
      disposeActiveProbeSession()
      return null
    }

    if (context.analysis.issues.isNotEmpty()) {
      disposeActiveProbeSession()
      return EpicsProbeViewState(
        recordName = context.analysis.recordName.orEmpty(),
        recordType = "",
        value = "",
        valueKey = null,
        valueCanPut = false,
        lastUpdated = "",
        access = "",
        fields = emptyList(),
        message = context.analysis.issues.first().message,
      )
    }

    val recordName = context.analysis.recordName
    if (recordName.isNullOrBlank()) {
      disposeActiveProbeSession()
      return EpicsProbeViewState(
        recordName = "",
        recordType = "",
        value = "",
        valueKey = null,
        valueCanPut = false,
        lastUpdated = "",
        access = "",
        fields = emptyList(),
        message = "Probe files must contain exactly one record name.",
      )
    }

    if (!monitoringActive) {
      disposeActiveProbeSession()
      return EpicsProbeViewState(
        recordName = recordName,
        recordType = "(stopped)",
        value = "(stopped)",
        valueKey = null,
        valueCanPut = false,
        lastUpdated = "Waiting for monitoring",
        access = "(stopped)",
        fields = emptyList(),
        message = "Start Monitoring to view probe runtime values.",
      )
    }

    val session = ensureActiveProbeSession(context) ?: return null
    return session.buildViewState()
  }

  internal fun requestPutActiveProbeValue(key: String) {
    val editor = FileEditorManager.getInstance(project).selectedTextEditor ?: return
    requestPutProbeValue(editor, key)
  }

  internal fun requestPutProbeValue(editor: Editor, key: String) {
    ensureActiveProbeSession(getProbeContext(editor) ?: return)?.requestPut(key)
  }

  internal fun getWidgetProbeViewState(widgetId: String, recordNameInput: String): EpicsProbeViewState? {
    val recordName = recordNameInput.trim()
    if (recordName.isBlank()) {
      releaseWidgetProbeSession(widgetId)
      return EpicsProbeViewState(
        recordName = "",
        recordType = "",
        value = "",
        valueKey = null,
        valueCanPut = false,
        lastUpdated = "",
        access = "",
        fields = emptyList(),
        message = "Enter a channel name and press Enter to start the probe.",
      )
    }

    if (!monitoringActive) {
      releaseWidgetProbeSession(widgetId)
      return EpicsProbeViewState(
        recordName = recordName,
        recordType = "(stopped)",
        value = "(stopped)",
        valueKey = null,
        valueCanPut = false,
        lastUpdated = "Waiting for monitoring",
        access = "(stopped)",
        fields = emptyList(),
        message = "Press Enter after changing the channel name to start the probe.",
      )
    }

    return ensureWidgetProbeSession(widgetId, recordName)?.buildViewState()
  }

  internal fun requestPutWidgetValue(widgetId: String, recordNameInput: String, key: String) {
    val recordName = recordNameInput.trim()
    if (recordName.isBlank()) {
      return
    }
    ensureWidgetProbeSession(widgetId, recordName)?.requestPut(key)
  }

  internal fun requestProcessProbe(editor: Editor) {
    if (!monitoringActive) {
      return
    }
    val context = getProbeContext(editor) ?: return
    val recordName = context.analysis.recordName?.takeIf { it.isNotBlank() } ?: return
    requestProcessRecord(recordName)
  }

  internal fun requestProcessWidget(widgetId: String, recordNameInput: String) {
    if (!monitoringActive) {
      return
    }
    val recordName = recordNameInput.trim()
    if (recordName.isBlank()) {
      return
    }
    ensureWidgetProbeSession(widgetId, recordName)
    requestProcessRecord(recordName)
  }

  internal fun releaseWidgetProbeSession(widgetId: String) {
    widgetProbeSessions.remove(widgetId)?.close()
  }

  internal fun getWidgetPvlistViewState(
    widgetId: String,
    model: EpicsPvlistWidgetModel,
  ): EpicsPvlistWidgetViewState {
    val plan = EpicsPvlistWidgetSupport.buildMonitorPlan(model, defaultProtocol)
    if (model.rawPvNames.isEmpty()) {
      releaseWidgetPvlistSession(widgetId)
      return EpicsPvlistWidgetViewState(
        rows = emptyList(),
        message = "Add channels to start monitoring.",
      )
    }

    if (!monitoringActive) {
      releaseWidgetPvlistSession(widgetId)
      return buildStoppedPvlistViewState(plan)
    }

    val session = ensureWidgetPvlistSession(widgetId, plan)
    return session?.buildViewState(plan, monitoringActive = true)
      ?: EpicsPvlistWidgetViewState(
        rows = plan.rows.map { row ->
          EpicsPvlistWidgetRowViewState(
            channelName = row.channelName,
            value = row.unresolvedValue ?: "(connecting...)",
            definitionKey = row.definitionKey,
          )
        },
        message = "Connecting PV list channels.",
      )
  }

  internal fun requestPutWidgetPvlistValue(
    widgetId: String,
    model: EpicsPvlistWidgetModel,
    definitionKey: String,
    input: String,
  ) {
    if (!monitoringActive) {
      return
    }
    val plan = EpicsPvlistWidgetSupport.buildMonitorPlan(model, defaultProtocol)
    val definition = plan.definitions.firstOrNull { it.key == definitionKey } ?: return
    ensureWidgetPvlistSession(widgetId, plan)
    val activeCaContext = caContext ?: return
    val activePvaClient = pvaClient ?: return
    ApplicationManager.getApplication().executeOnPooledThread {
      try {
        RuntimeEditorSession.putValueNow(
          caContext = activeCaContext,
          pvaClient = activePvaClient,
          protocol = definition.protocol,
          pvName = definition.pvName,
          input = input,
        )
      } catch (error: Exception) {
        showRuntimePutError(project, "${definition.pvName}: ${error.message ?: error.javaClass.simpleName}")
      }
    }
  }

  internal fun releaseWidgetPvlistSession(widgetId: String) {
    widgetPvlistSessions.remove(widgetId)?.close()
  }

  private fun isRuntimeFile(fileName: String): Boolean = isMonitorFile(fileName) || isDatabaseFile(fileName)

  private fun isMonitorFile(fileName: String): Boolean = fileName.substringAfterLast('.', "").lowercase() == "pvlist"

  internal fun isProbeFileName(fileName: String): Boolean = fileName.substringAfterLast('.', "").lowercase() == "probe"

  private fun isDatabaseFile(fileName: String): Boolean {
    return fileName.substringAfterLast('.', "").lowercase() in DATABASE_FILE_EXTENSIONS
  }

  private fun ensureActiveProbeSession(context: ActiveProbeContext): EpicsProbeRuntimeSession? {
    val existingSession = activeProbeSession
    if (existingSession != null && existingSession.matches(context.sourceKey, context.analysis, defaultProtocol)) {
      return existingSession
    }

    disposeActiveProbeSession()
    val activeCaContext = caContext ?: return null
    val activePvaClient = pvaClient ?: return null
    return EpicsProbeRuntimeSession(
      project = project,
      caContext = activeCaContext,
      pvaClient = activePvaClient,
      protocol = defaultProtocol,
      sourceKey = context.sourceKey,
      analysis = context.analysis,
    ).also { session ->
      activeProbeSession = session
      session.start()
    }
  }

  private fun ensureWidgetProbeSession(widgetId: String, recordName: String): EpicsProbeRuntimeSession? {
    val analysis = EpicsProbeDocumentAnalysis(
      recordName = recordName,
      overlayOffset = null,
      issues = emptyList(),
    )
    val sourceKey = buildWidgetProbeSourceKey(widgetId, recordName)
    val existingSession = widgetProbeSessions[widgetId]
    if (existingSession != null && existingSession.matches(sourceKey, analysis, defaultProtocol)) {
      return existingSession
    }

    existingSession?.close()
    val activeCaContext = caContext ?: return null
    val activePvaClient = pvaClient ?: return null
    return EpicsProbeRuntimeSession(
      project = project,
      caContext = activeCaContext,
      pvaClient = activePvaClient,
      protocol = defaultProtocol,
      sourceKey = sourceKey,
      analysis = analysis,
    ).also { session ->
      widgetProbeSessions[widgetId] = session
      session.start()
    }
  }

  private fun disposeActiveProbeSession() {
    activeProbeSession?.close()
    activeProbeSession = null
  }

  private fun disposeAllWidgetProbeSessions() {
    widgetProbeSessions.values.forEach(EpicsProbeRuntimeSession::close)
    widgetProbeSessions.clear()
  }

  private fun ensureWidgetPvlistSession(
    widgetId: String,
    plan: EpicsPvlistWidgetPlan,
  ): EpicsPvlistWidgetSession? {
    val existingSession = widgetPvlistSessions[widgetId]
    if (existingSession != null) {
      existingSession.updatePlan(plan)
      return existingSession
    }

    val activeCaContext = caContext ?: return null
    val activePvaClient = pvaClient ?: return null
    val createdSession = EpicsPvlistWidgetSession(project, activeCaContext, activePvaClient)
    val previousSession = widgetPvlistSessions.putIfAbsent(widgetId, createdSession)
    val session = if (previousSession != null) {
      createdSession.close()
      previousSession
    } else {
      createdSession
    }
    session.updatePlan(plan)
    return session
  }

  private fun disposeAllWidgetPvlistSessions() {
    widgetPvlistSessions.values.forEach(EpicsPvlistWidgetSession::close)
    widgetPvlistSessions.clear()
  }

  private fun buildWidgetProbeSourceKey(widgetId: String, recordName: String): String {
    return "__widget__:$widgetId:$recordName:${defaultProtocol.name}"
  }

  private fun buildStoppedPvlistViewState(plan: EpicsPvlistWidgetPlan): EpicsPvlistWidgetViewState {
    return EpicsPvlistWidgetViewState(
      rows = plan.rows.map { row ->
        EpicsPvlistWidgetRowViewState(
          channelName = row.channelName,
          value = row.unresolvedValue ?: "(stopped)",
          definitionKey = row.definitionKey,
        )
      },
      message = "Start Monitoring to view PV list runtime values.",
    )
  }

  private fun requestProcessRecord(recordName: String) {
    val activeCaContext = caContext ?: return
    val activePvaClient = pvaClient ?: return
    val protocol = defaultProtocol
    ApplicationManager.getApplication().executeOnPooledThread {
      try {
        RuntimeEditorSession.putValueNow(
          caContext = activeCaContext,
          pvaClient = activePvaClient,
          protocol = protocol,
          pvName = "$recordName.PROC",
          input = "1",
        )
      } catch (error: Exception) {
        showRuntimePutError(project, "$recordName.PROC: ${error.message ?: error.javaClass.simpleName}")
      }
    }
  }

  private fun parseDatabaseTocStates(
    editor: Editor,
    text: String,
    defaultProtocol: MonitorProtocol,
  ): List<RuntimeValueState> {
    val tocEntries = EpicsDatabaseToc.extractRuntimeEntries(text)
    if (tocEntries.isEmpty()) {
      return emptyList()
    }

    val macroDefinitions = createDatabaseMonitorMacroDefinitions(
      EpicsDatabaseToc.extractRuntimeMacroAssignments(text),
    )
    val macroExpansionCache = mutableMapOf<String, String>()
    return tocEntries.map { tocEntry ->
      val pvName = normalizeDatabaseMonitorPvName(
        expandDatabaseMonitorValue(
          tocEntry.recordName,
          macroDefinitions,
          macroExpansionCache,
          emptyList(),
        ),
        tocEntry.recordName,
      )
      DatabaseTocLineState(
        editor = editor,
        protocol = defaultProtocol,
        pvName = pvName,
        valueStartOffset = tocEntry.valueStart,
        valueEndOffset = tocEntry.valueEnd,
      )
    }
  }

  override fun dispose() {
    stopMonitoring()
  }

  private fun publishStateChanged() {
    project.messageBus.syncPublisher(EpicsMonitorRuntimeStateListener.TOPIC)
      .monitoringStateChanged(monitoringActive)
  }

  private companion object {
    val SESSION_KEY = Key.create<Disposable>("org.epics.workbench.runtime.monitorSession")
    val PROBE_OVERLAY_KEY = Key.create<ProbeEditorOverlay>("org.epics.workbench.runtime.probeOverlay")
  }
}

internal class RuntimeEditorSession(
  private val project: Project,
  private val caContext: Context,
  private val pvaClient: PVAClient,
  private val states: List<RuntimeValueState>,
) : Disposable {
  private val disposed = AtomicBoolean(false)
  private val channelHandles = ConcurrentHashMap<RuntimeValueState, RuntimeChannelHandle>()
  private val cleanupActions = mutableListOf<() -> Unit>()

  fun start() {
    states.forEach(RuntimeValueState::install)
    for (state in states) {
      ApplicationManager.getApplication().executeOnPooledThread {
        connectAndMonitor(state)
      }
    }
  }

  fun handleDoubleClick(event: EditorMouseEvent): Boolean {
    val state = states.firstOrNull { it.matchesDoubleClick(event) } ?: return false
    promptForPutValue(state)
    return true
  }

  fun requestPutValue(state: RuntimeValueState) {
    if (!states.contains(state)) {
      return
    }
    promptForPutValue(state)
  }

  private fun connectAndMonitor(state: RuntimeValueState) {
    if (disposed.get() || project.isDisposed) {
      return
    }

    state.setConnecting()
    when (state.protocol) {
      MonitorProtocol.CA -> connectCa(state)
      MonitorProtocol.PVA -> connectPva(state)
    }
  }

  private fun connectCa(state: RuntimeValueState) {
    var channel: Channel? = null
    var monitor: Monitor? = null
    try {
      caContext.attachCurrentThread()
      channel = connectCaChannel(state.pvName, state)
      val fieldType = channel.fieldType
      val labels = if (fieldType == DBRType.ENUM) fetchCaEnumLabels(channel) else emptyList()
      state.setEnumChoices(labels)
      state.setAccess(if (channel.getWriteAccess()) "Read/Write" else "Read only", channel.getWriteAccess())
      channelHandles[state] = RuntimeChannelHandle.Ca(channel, fieldType)

      val monitorType = getCaMonitorType(fieldType)
      monitor = channel.addMonitor(
        monitorType,
        channel.elementCount.coerceAtLeast(1),
        Monitor.VALUE or Monitor.ALARM,
        MonitorListener { event ->
          if (disposed.get()) {
            return@MonitorListener
          }
          handleCaMonitorEvent(state, event, labels)
        },
      )
      caContext.flushIO()
      cleanupActions += {
        runCatching { monitor?.clear() }
        runCatching { channel.destroy() }
        channelHandles.remove(state)
      }
    } catch (error: Exception) {
      channelHandles.remove(state)
      state.setConnecting()
      runCatching { monitor?.clear() }
      runCatching { channel?.destroy() }
    }
  }

  private fun connectCaChannel(pvName: String, state: RuntimeValueState): Channel {
    val latch = CountDownLatch(1)
    var connected = false
    val channel = caContext.createChannel(
      pvName,
      ConnectionListener { event: ConnectionEvent ->
        if (event.isConnected) {
          connected = true
          latch.countDown()
        } else {
          state.setConnecting()
        }
      },
    )
    caContext.flushIO()
    if (!latch.await(CA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS) || !connected) {
      runCatching { channel.destroy() }
      throw TimeoutException("Timed out connecting to $pvName")
    }
    return channel
  }

  private fun fetchCaEnumLabels(channel: Channel): List<String> {
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

  private fun handleCaMonitorEvent(
    state: RuntimeValueState,
    event: MonitorEvent,
    fallbackLabels: List<String>,
  ) {
    if (!event.status.isSuccessful) {
      state.setConnecting()
      return
    }
    val display = formatCaValue(event.dbr, fallbackLabels)
    state.setValue(display)
  }

  private fun connectPva(state: RuntimeValueState) {
    var channel: PVAChannel? = null
    var subscription: AutoCloseable? = null
    try {
      channel = pvaClient.getChannel(state.pvName) { _: PVAChannel, channelState: ClientChannelState ->
        if (!disposed.get() && channelState != ClientChannelState.CONNECTED) {
          state.setConnecting()
        }
      }
      channel.connect().get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      channelHandles[state] = RuntimeChannelHandle.Pva(channel)
      val initialValue = channel.read("").get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      val (accessLabel, canPut) = determinePvaAccess(initialValue)
      state.setAccess(accessLabel, canPut)
      handlePvaStructure(state, initialValue)
      subscription = channel.subscribe("", PvaMonitorListener { _: PVAChannel, _, _, structure: PVAStructure ->
        if (disposed.get()) {
          return@PvaMonitorListener
        }
        handlePvaStructure(state, structure)
      })
      cleanupActions += {
        runCatching { subscription?.close() }
        runCatching { channel?.close() }
        channelHandles.remove(state)
      }
    } catch (error: Exception) {
      channelHandles.remove(state)
      state.setConnecting()
      runCatching { subscription?.close() }
      runCatching { channel?.close() }
    }
  }

  private fun handlePvaStructure(state: RuntimeValueState, structure: PVAStructure?) {
    if (structure == null) {
      state.setValue("")
      return
    }
    state.setValue(formatPvaStructure(structure, state))
  }

  override fun dispose() {
    if (!disposed.compareAndSet(false, true)) {
      return
    }
    channelHandles.clear()
    cleanupActions.asReversed().forEach { action -> runCatching(action) }
    cleanupActions.clear()
    states.forEach(RuntimeValueState::dispose)
  }

  private fun promptForPutValue(state: RuntimeValueState) {
    val input = Messages.showInputDialog(
      project,
      "Put value for ${state.pvName}",
      "EPICS Put Runtime Value",
      null,
      state.getPutInitialValue(),
      null,
    ) ?: return

    ApplicationManager.getApplication().executeOnPooledThread {
      try {
        putRuntimeValue(state, input)
      } catch (error: Exception) {
        showPutError("${state.pvName}: ${error.message ?: error.javaClass.simpleName}")
      }
    }
  }

  private fun putRuntimeValue(state: RuntimeValueState, input: String) {
    when (val handle = channelHandles[state] ?: throw IllegalStateException("Channel is not connected.")) {
      is RuntimeChannelHandle.Ca -> putCaValue(state, handle, input)
      is RuntimeChannelHandle.Pva -> putPvaValue(state, handle, input)
    }
  }

  private fun putCaValue(state: RuntimeValueState, handle: RuntimeChannelHandle.Ca, input: String) {
    val channel = handle.channel
    if (!channel.getWriteAccess()) {
      throw IllegalStateException("Channel is not writable.")
    }
    if (channel.elementCount > 1) {
      throw IllegalStateException("Array values cannot be changed from runtime editing.")
    }

    when (handle.fieldType) {
      DBRType.STRING -> channel.put(input)
      DBRType.ENUM -> channel.put(resolveEnumInput(input, state.getEnumChoices()))
      DBRType.BYTE -> channel.put(parseIntegerInput(input).toByte())
      DBRType.SHORT -> channel.put(parseIntegerInput(input).toShort())
      DBRType.INT -> channel.put(parseIntegerInput(input))
      DBRType.FLOAT -> channel.put(parseDecimalInput(input).toFloat())
      DBRType.DOUBLE -> channel.put(parseDecimalInput(input))
      else -> throw IllegalStateException("Unsupported CA field type: ${handle.fieldType}.")
    }
    caContext.flushIO()
  }

  private fun putPvaValue(state: RuntimeValueState, handle: RuntimeChannelHandle.Pva, input: String) {
    val channel = handle.channel
    if (!channel.isConnected) {
      throw IllegalStateException("Channel is not connected.")
    }

    val structure = channel.read("").get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
    val valueField = structure?.get<PVAData>("value")
      ?: throw IllegalStateException("PV has data, but no value field.")

    when {
      valueField is PVAStructure -> {
        val pvaEnum = runCatching { PVAEnum.fromStructure(valueField) }.getOrNull()
          ?: throw IllegalStateException("Unsupported structured PVA value.")
        val choices = pvaEnum.get<PVAStringArray>("choices")?.get()?.map { it ?: "" }
          ?: state.getEnumChoices()
        channel.write("value.index", resolveEnumInput(input, choices))
          .get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      }
      valueField is PVAArray -> throw IllegalStateException("Array values cannot be changed from runtime editing.")
      valueField is PVAString -> channel.write("value", input).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVAByte -> channel.write("value", parseIntegerInput(input).toByte()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVAShort -> channel.write("value", parseIntegerInput(input).toShort()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVAInt -> channel.write("value", parseIntegerInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVALong -> channel.write("value", parseLongInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVAFloat -> channel.write("value", parseDecimalInput(input).toFloat()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      valueField is PVADouble -> channel.write("value", parseDecimalInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
      else -> throw IllegalStateException("Unsupported PVA value type.")
    }
  }

  private fun showPutError(message: String) {
    showRuntimePutError(project, message)
  }

  private fun getCaMonitorType(fieldType: DBRType): DBRType {
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

  companion object {
    fun putValueNow(
      caContext: Context,
      pvaClient: PVAClient,
      protocol: MonitorProtocol,
      pvName: String,
      input: String,
    ) {
      when (protocol) {
        MonitorProtocol.CA -> putCaValueNow(caContext, pvName, input)
        MonitorProtocol.PVA -> putPvaValueNow(pvaClient, pvName, input)
      }
    }

    private fun putCaValueNow(
      caContext: Context,
      pvName: String,
      input: String,
    ) {
      caContext.attachCurrentThread()
      val channel = connectCaChannelNow(caContext, pvName)
      try {
        if (!channel.getWriteAccess()) {
          throw IllegalStateException("Channel is not writable.")
        }
        if (channel.elementCount > 1) {
          throw IllegalStateException("Array values cannot be changed from runtime editing.")
        }
        val fieldType = channel.fieldType
        val labels = if (fieldType == DBRType.ENUM) fetchCaEnumLabelsNow(caContext, channel) else emptyList()
        when (fieldType) {
          DBRType.STRING -> channel.put(input)
          DBRType.ENUM -> channel.put(resolveEnumInput(input, labels))
          DBRType.BYTE -> channel.put(parseIntegerInput(input).toByte())
          DBRType.SHORT -> channel.put(parseIntegerInput(input).toShort())
          DBRType.INT -> channel.put(parseIntegerInput(input))
          DBRType.FLOAT -> channel.put(parseDecimalInput(input).toFloat())
          DBRType.DOUBLE -> channel.put(parseDecimalInput(input))
          else -> throw IllegalStateException("Unsupported CA field type: $fieldType.")
        }
        caContext.flushIO()
      } finally {
        runCatching { channel.destroy() }
      }
    }

    private fun putPvaValueNow(
      pvaClient: PVAClient,
      pvName: String,
      input: String,
    ) {
      val channel = pvaClient.getChannel(pvName)
      try {
        channel.connect().get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
        if (!channel.isConnected) {
          throw IllegalStateException("Channel is not connected.")
        }
        val structure = channel.read("").get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
        val valueField = structure?.get<PVAData>("value")
          ?: throw IllegalStateException("PV has data, but no value field.")

        when {
          valueField is PVAStructure -> {
            val pvaEnum = runCatching { PVAEnum.fromStructure(valueField) }.getOrNull()
              ?: throw IllegalStateException("Unsupported structured PVA value.")
            val choices = pvaEnum.get<PVAStringArray>("choices")?.get()?.map { it ?: "" }.orEmpty()
            channel.write("value.index", resolveEnumInput(input, choices))
              .get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          }
          valueField is PVAArray -> throw IllegalStateException("Array values cannot be changed from runtime editing.")
          valueField is PVAString -> channel.write("value", input).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVAByte -> channel.write("value", parseIntegerInput(input).toByte()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVAShort -> channel.write("value", parseIntegerInput(input).toShort()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVAInt -> channel.write("value", parseIntegerInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVALong -> channel.write("value", parseLongInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVAFloat -> channel.write("value", parseDecimalInput(input).toFloat()).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          valueField is PVADouble -> channel.write("value", parseDecimalInput(input)).get(PVA_CONNECT_TIMEOUT_MS, TimeUnit.MILLISECONDS)
          else -> throw IllegalStateException("Unsupported PVA value type.")
        }
      } finally {
        runCatching { channel.close() }
      }
    }

    private fun connectCaChannelNow(caContext: Context, pvName: String): Channel {
      val latch = CountDownLatch(1)
      var connected = false
      val channel = caContext.createChannel(
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
        runCatching { channel.destroy() }
        throw TimeoutException("Timed out connecting to $pvName")
      }
      return channel
    }

    private fun fetchCaEnumLabelsNow(caContext: Context, channel: Channel): List<String> {
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
  }
}

private fun showRuntimePutError(project: Project, message: String) {
  ApplicationManager.getApplication().invokeLater {
    if (!project.isDisposed) {
      Messages.showErrorDialog(project, message, "EPICS Put Runtime Value")
    }
  }
}

internal interface RuntimeValueState : Disposable {
  val protocol: MonitorProtocol
  val pvName: String

  fun install()
  fun matchesDoubleClick(event: EditorMouseEvent): Boolean
  fun setEnumChoices(choices: List<String>)
  fun getEnumChoices(): List<String>
  fun getPutInitialValue(): String
  fun setValue(value: String)
  fun setConnecting()
  fun setAccess(accessLabel: String, canPut: Boolean) = Unit
}

private class MonitorLineState(
  private val editor: Editor,
  val entry: MonitorEntry,
  private val alignColumn: Int,
) : RuntimeValueState {
  private val renderer = MonitorValueRenderer()
  private var inlay: Inlay<MonitorValueRenderer>? = null
  private var enumChoices: List<String> = emptyList()
  private var putInitialValue: String = ""

  override val protocol: MonitorProtocol
    get() = entry.protocol

  override val pvName: String
    get() = entry.pvName

  override fun install() {
    val lineEndOffset = entry.lineStartOffset + entry.lineText.length
    inlay = editor.inlayModel.addAfterLineEndElement(lineEndOffset, false, renderer)
  }

  override fun matchesDoubleClick(event: EditorMouseEvent): Boolean = event.inlay == inlay

  override fun setEnumChoices(choices: List<String>) {
    enumChoices = choices
  }

  override fun getPutInitialValue(): String = putInitialValue

  override fun setValue(value: String) {
    putInitialValue = value
    val padding = (alignColumn - entry.displayText.length).coerceAtLeast(0) + 2
    val padded = " ".repeat(padding) + value
    ApplicationManager.getApplication().invokeLater(
      {
        if (!editor.isDisposed) {
          renderer.text = padded
          inlay?.update()
        }
      },
      { editor.isDisposed },
    )
  }

  override fun setConnecting() {
    setValue(CONNECTING_DISPLAY)
  }

  override fun getEnumChoices(): List<String> = enumChoices

  override fun dispose() {
    inlay?.dispose()
    inlay = null
  }

  private companion object {
    private const val CONNECTING_DISPLAY = "(connecting...)"
  }
}

private class DatabaseTocLineState(
  private val editor: Editor,
  override val protocol: MonitorProtocol,
  override val pvName: String,
  private val valueStartOffset: Int,
  private val valueEndOffset: Int,
) : RuntimeValueState {
  private val renderer = DatabaseTocValueRenderer()
  private val updateLock = Any()
  private var highlighter: RangeHighlighter? = null
  private var enumChoices: List<String> = emptyList()
  private var pendingText: String = ""
  private var lastFlushAt: Long = 0L
  private var flushScheduled = false
  private var putInitialValue: String = ""

  override fun install() {
    val markupModel = editor.markupModel
    highlighter = markupModel.addRangeHighlighter(
      valueStartOffset,
      valueEndOffset,
      HighlighterLayer.ELEMENT_UNDER_CARET,
      null,
      HighlighterTargetArea.EXACT_RANGE,
    ).also { createdHighlighter ->
      createdHighlighter.setCustomRenderer(renderer)
    }
  }

  override fun matchesDoubleClick(event: EditorMouseEvent): Boolean {
    if (event.inlay != null) {
      return false
    }
    return event.offset in valueStartOffset until valueEndOffset
  }

  override fun setEnumChoices(choices: List<String>) {
    enumChoices = choices
  }

  override fun getEnumChoices(): List<String> = enumChoices

  override fun getPutInitialValue(): String = putInitialValue

  override fun setValue(value: String) {
    putInitialValue = value
    val displayText = formatDatabaseTocDisplayValue(value)
    scheduleOverlayUpdate(displayText)
  }

  override fun setConnecting() {
    setValue(CONNECTING_DISPLAY)
  }

  override fun dispose() {
    highlighter?.let(editor.markupModel::removeHighlighter)
    highlighter = null
  }

  private fun scheduleOverlayUpdate(displayText: String) {
    val delayMs = synchronized(updateLock) {
      pendingText = displayText
      val now = System.currentTimeMillis()
      val delay = (lastFlushAt + TOC_REFRESH_INTERVAL_MS - now).coerceAtLeast(0L)
      if (flushScheduled) {
        return@synchronized null
      }
      flushScheduled = true
      delay
    } ?: return

    if (delayMs == 0L) {
      flushOverlayUpdate()
    } else {
      AppExecutorUtil.getAppScheduledExecutorService().schedule(
        { flushOverlayUpdate() },
        delayMs,
        TimeUnit.MILLISECONDS,
      )
    }
  }

  private fun flushOverlayUpdate() {
    val displayText = synchronized(updateLock) {
      flushScheduled = false
      lastFlushAt = System.currentTimeMillis()
      pendingText
    }

    ApplicationManager.getApplication().invokeLater(
      {
        if (!editor.isDisposed) {
          renderer.text = displayText
          editor.contentComponent.repaint()
        }
      },
      { editor.isDisposed },
    )
  }

  private companion object {
    private const val CONNECTING_DISPLAY = "(connecting...)"
    private const val TOC_REFRESH_INTERVAL_MS = 1000L
  }
}

private class MonitorValueRenderer : EditorCustomElementRenderer {
  var text: String = " "

  override fun calcWidthInPixels(inlay: Inlay<*>): Int {
    val font = getEditorFont(inlay.editor)
    val width = inlay.editor.contentComponent.getFontMetrics(font).stringWidth(text)
    return width.coerceAtLeast(1)
  }

  override fun paint(
    inlay: Inlay<*>,
    graphics: Graphics2D,
    targetRegion: Rectangle2D,
    textAttributes: TextAttributes,
  ) {
    paintString(inlay, graphics, targetRegion.x.toInt(), targetRegion.y.toInt())
  }

  @Deprecated("Old renderer signature")
  override fun paint(
    inlay: Inlay<*>,
    graphics: java.awt.Graphics,
    targetRegion: Rectangle,
    textAttributes: TextAttributes,
  ) {
    paintString(inlay, graphics as Graphics2D, targetRegion.x, targetRegion.y)
  }

  private fun paintString(
    inlay: Inlay<*>,
    graphics: Graphics2D,
    x: Int,
    y: Int,
  ) {
    if (text.isEmpty()) {
      return
    }
    val font = getEditorFont(inlay.editor)
    graphics.font = font
    graphics.color = JBColor.GRAY
    val metrics = inlay.editor.contentComponent.getFontMetrics(font)
    graphics.drawString(text, x, y + metrics.ascent)
  }

  private fun getEditorFont(editor: Editor): Font {
    return editor.colorsScheme.getFont(EditorFontType.PLAIN)
  }
}

private class DatabaseTocValueRenderer : CustomHighlighterRenderer {
  var text: String = ""

  override fun paint(editor: Editor, highlighter: RangeHighlighter, graphics: Graphics) {
    if (text.isEmpty() || editor.isDisposed || !highlighter.isValid) {
      return
    }

    val startPoint = editor.offsetToXY(highlighter.startOffset)
    val endPoint = editor.offsetToXY(highlighter.endOffset)
    val clipWidth = (endPoint.x - startPoint.x).coerceAtLeast(1)
    val clipHeight = editor.lineHeight
    val oldClip = graphics.clip

    try {
      graphics.clipRect(startPoint.x, startPoint.y, clipWidth, clipHeight)
      graphics.font = editor.colorsScheme.getFont(EditorFontType.PLAIN)
      graphics.color = JBColor.GRAY
      graphics.drawString(text, startPoint.x, startPoint.y + editor.ascent)
    } finally {
      graphics.clip = oldClip
    }
  }
}

private data class MonitorDocumentDefinition(
  val entries: List<MonitorEntry>,
  val maxDisplayWidth: Int,
)

private data class MonitorEntry(
  val protocol: MonitorProtocol,
  val pvName: String,
  val lineText: String,
  val displayText: String,
  val lineStartOffset: Int,
)

internal enum class MonitorProtocol {
  CA,
  PVA,
}

private fun parseMonitorDocument(
  text: String,
  defaultProtocol: MonitorProtocol,
): MonitorDocumentDefinition {
  val entries = mutableListOf<MonitorEntry>()
  val macroDefinitions = linkedMapOf<String, String>()
  var maxDisplayWidth = 0
  var lineStart = 0

  while (lineStart <= text.length) {
    val lineEnd = text.indexOf('\n', lineStart).let { if (it >= 0) it else text.length }
    val rawLineEnd = if (lineEnd > lineStart && text[lineEnd - 1] == '\r') lineEnd - 1 else lineEnd
    val lineText = text.substring(lineStart, rawLineEnd)
    val trimmed = lineText.trim()

    when {
      trimmed.isEmpty() || trimmed.startsWith("#") -> Unit
      parseMonitorMacroAssignment(lineText)?.let { assignment ->
        macroDefinitions[assignment.name] = assignment.value
        true
      } == true -> Unit
      !trimmed.contains(' ') && !trimmed.contains('\t') -> {
        val expanded = expandMonitorValue(trimmed, macroDefinitions, linkedSetOf()) ?: ""
        if (expanded.isNotBlank() && !expanded.any(Char::isWhitespace)) {
          val (protocol, pvName) = splitMonitorProtocol(expanded, defaultProtocol)
          entries += MonitorEntry(
            protocol = protocol,
            pvName = pvName,
            lineText = lineText,
            displayText = trimmed,
            lineStartOffset = lineStart,
          )
          maxDisplayWidth = maxOf(maxDisplayWidth, trimmed.length)
        }
      }
    }

    if (lineEnd >= text.length) {
      break
    }
    lineStart = lineEnd + 1
  }

  return MonitorDocumentDefinition(entries, maxDisplayWidth)
}

private fun parseMonitorMacroAssignment(lineText: String): MonitorMacroAssignment? {
  val match = MONITOR_MACRO_ASSIGNMENT_REGEX.matchEntire(lineText) ?: return null
  return MonitorMacroAssignment(
    name = match.groups[1]?.value.orEmpty(),
    value = match.groups[2]?.value.orEmpty(),
  )
}

private fun expandMonitorValue(
  text: String,
  macroDefinitions: Map<String, String>,
  stack: LinkedHashSet<String>,
): String? {
  var unresolved = false
  val expanded = MONITOR_MACRO_REFERENCE_REGEX.replace(text) { match ->
    val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
    val defaultValue = match.groups[2]?.value
    val resolved = resolveMonitorMacro(macroName, defaultValue, macroDefinitions, stack)
    if (resolved == null) {
      unresolved = true
      ""
    } else {
      resolved
    }
  }
  return if (unresolved) null else expanded
}

private fun resolveMonitorMacro(
  macroName: String,
  defaultValue: String?,
  macroDefinitions: Map<String, String>,
  stack: LinkedHashSet<String>,
): String? {
  if (macroName in stack) {
    return null
  }

  val value = macroDefinitions[macroName] ?: return defaultValue
  val nextStack = LinkedHashSet(stack)
  nextStack += macroName
  return expandMonitorValue(value, macroDefinitions, nextStack)
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

private fun EpicsRuntimeProtocol.toMonitorProtocol(): MonitorProtocol {
  return when (this) {
    EpicsRuntimeProtocol.CA -> MonitorProtocol.CA
    EpicsRuntimeProtocol.PVA -> MonitorProtocol.PVA
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

private fun formatPvaStructure(structure: PVAStructure, state: RuntimeValueState): String {
  val valueField = structure.get<PVAData>("value")
  if (valueField == null) {
    return if (structure.get().isNotEmpty()) "Has data, but no value" else ""
  }

  if (valueField is PVAStructure) {
    val pvaEnum = runCatching { PVAEnum.fromStructure(valueField) }.getOrNull()
    if (pvaEnum != null) {
      val index = pvaEnum.get<PVAInt>("index")?.get() ?: 0
      val choices = pvaEnum.get<PVAStringArray>("choices")?.get()?.map { it ?: "" }
        ?: state.getEnumChoices()
      state.setEnumChoices(choices)
      return formatEnumValue(index, choices)
    }
  }

  if (valueField is PVAArray) {
    return valueField.toString()
  }

  if (valueField is PVAValue) {
    return valueField.formatValue()
  }

  return valueField.toString()
}

private fun determinePvaAccess(structure: PVAStructure?): Pair<String, Boolean> {
  val valueField = structure?.get<PVAData>("value") ?: return "Read only" to false
  return when (valueField) {
    is PVAStructure -> {
      val pvaEnum = runCatching { PVAEnum.fromStructure(valueField) }.getOrNull()
      if (pvaEnum != null) {
        "Read/Write" to true
      } else {
        "Read only" to false
      }
    }
    is PVAArray -> "Read only" to false
    is PVAString,
    is PVAByte,
    is PVAShort,
    is PVAInt,
    is PVALong,
    is PVAFloat,
    is PVADouble,
    -> "Read/Write" to true
    else -> "Read only" to false
  }
}

private fun createDatabaseMonitorMacroDefinitions(
  assignments: Map<String, EpicsDatabaseToc.MacroAssignment>,
): Map<String, String> {
  val macroDefinitions = linkedMapOf<String, String>()
  for ((macroName, assignment) in assignments) {
    if (!assignment.hasAssignment) {
      continue
    }
    macroDefinitions[macroName] = assignment.value
  }
  return macroDefinitions
}

private fun expandDatabaseMonitorValue(
  text: String,
  macroDefinitions: Map<String, String>,
  cache: MutableMap<String, String>,
  stack: List<String>,
): String {
  val sourceText = text
  val builder = StringBuilder()
  var cursor = 0

  for (match in DATABASE_MONITOR_MACRO_REFERENCE_REGEX.findAll(sourceText)) {
    val matchText = match.value
    val matchIndex = match.range.first
    builder.append(sourceText, cursor, matchIndex)

    val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
    val defaultValue = match.groups[2]?.value
    builder.append(
      resolveDatabaseMonitorMacro(
        macroName = macroName,
        defaultValue = defaultValue,
        originalText = matchText,
        macroDefinitions = macroDefinitions,
        cache = cache,
        stack = stack,
      ),
    )
    cursor = match.range.last + 1
  }

  builder.append(sourceText.substring(cursor))
  return builder.toString()
}

private fun resolveDatabaseMonitorMacro(
  macroName: String,
  defaultValue: String?,
  originalText: String,
  macroDefinitions: Map<String, String>,
  cache: MutableMap<String, String>,
  stack: List<String>,
): String {
  val cacheKey = "$macroName\u0000${defaultValue ?: ""}\u0000$originalText"
  cache[cacheKey]?.let { return it }

  val resolvedValue = when {
    macroName in stack -> originalText
    macroDefinitions.containsKey(macroName) -> expandDatabaseMonitorValue(
      macroDefinitions[macroName].orEmpty(),
      macroDefinitions,
      cache,
      stack + macroName,
    )
    defaultValue != null -> expandDatabaseMonitorValue(
      defaultValue,
      macroDefinitions,
      cache,
      stack,
    )
    else -> originalText
  }

  cache[cacheKey] = resolvedValue
  return resolvedValue
}

private fun normalizeDatabaseMonitorPvName(expandedPvName: String, fallbackPvName: String): String {
  val normalizedExpandedPvName = expandedPvName.trim()
  return if (normalizedExpandedPvName.isEmpty() || normalizedExpandedPvName.any(Char::isWhitespace)) {
    fallbackPvName.trim()
  } else {
    normalizedExpandedPvName
  }
}

private fun formatDatabaseTocDisplayValue(value: String): String {
  val normalized = if (value.isEmpty()) "\"\"" else value
  return truncateText(normalized, DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH)
}

private fun truncateText(text: String, maxLength: Int): String {
  if (text.length <= maxLength) {
    return text
  }
  if (maxLength <= 3) {
    return ".".repeat(maxLength)
  }
  return text.take(maxLength - 3) + "..."
}

private fun parseIntegerInput(input: String): Int {
  return normalizePutInput(input).toIntOrNull()
    ?: throw IllegalArgumentException("Expected an integer value.")
}

private fun parseLongInput(input: String): Long {
  return normalizePutInput(input).toLongOrNull()
    ?: throw IllegalArgumentException("Expected an integer value.")
}

private fun parseDecimalInput(input: String): Double {
  return normalizePutInput(input).toDoubleOrNull()
    ?: throw IllegalArgumentException("Expected a numeric value.")
}

private fun resolveEnumInput(input: String, choices: List<String>): Short {
  val normalizedInput = normalizePutInput(input)
  ENUM_INDEX_INPUT_REGEX.matchEntire(normalizedInput)?.groups?.get(1)?.value?.toShortOrNull()?.let { index ->
    return index
  }

  normalizedInput.toShortOrNull()?.let { index ->
    return index
  }

  val unquoted = normalizedInput.removeSurrounding("\"")
  val choiceIndex = choices.indexOf(unquoted)
  if (choiceIndex >= 0) {
    return choiceIndex.toShort()
  }

  throw IllegalArgumentException("Expected an enum index or one of: ${choices.joinToString(", ")}")
}

private fun normalizePutInput(input: String): String {
  return input.trim()
}

private fun formatEnumValue(index: Int, choices: List<String>): String {
  val choice = choices.getOrNull(index)
  return when {
    choice == null -> "[$index] Illegal_Value"
    choice.isEmpty() -> """[$index] """""
    else -> "[$index] $choice"
  }
}

private data class MonitorMacroAssignment(
  val name: String,
  val value: String,
)

private sealed interface RuntimeChannelHandle {
  data class Ca(
    val channel: Channel,
    val fieldType: DBRType,
  ) : RuntimeChannelHandle

  data class Pva(
    val channel: PVAChannel,
  ) : RuntimeChannelHandle
}

private val MONITOR_MACRO_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$""")
private val MONITOR_MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
private val DATABASE_MONITOR_MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
private val ENUM_INDEX_INPUT_REGEX = Regex("""^\[(\d+)](?:\s+.*)?$""")

private const val CA_CONNECT_TIMEOUT_MS = 5000L
private const val PVA_CONNECT_TIMEOUT_MS = 5000L
private const val DATABASE_TOC_VALUE_DISPLAY_MAX_LENGTH = 18
private val DATABASE_FILE_EXTENSIONS = setOf("db", "vdb", "template")
