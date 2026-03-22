package org.epics.workbench.runtime

import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.Project
import gov.aps.jca.Context
import org.epics.pva.client.PVAClient
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.probe.EpicsProbeDocumentAnalysis
import org.epics.workbench.probe.EpicsProbeSupport
import java.time.LocalTime
import java.time.format.DateTimeFormatter

internal data class EpicsProbeFieldViewState(
  val key: String,
  val fieldName: String,
  val value: String,
  val updated: String,
  val canPut: Boolean,
)

internal data class EpicsProbeViewState(
  val recordName: String,
  val recordType: String,
  val value: String,
  val valueKey: String?,
  val valueCanPut: Boolean,
  val lastUpdated: String,
  val access: String,
  val fields: List<EpicsProbeFieldViewState>,
  val message: String? = null,
)

internal class EpicsProbeRuntimeSession(
  private val project: Project,
  private val caContext: Context,
  private val pvaClient: PVAClient,
  private val protocol: MonitorProtocol,
  val sourceKey: String,
  private val analysis: EpicsProbeDocumentAnalysis,
) : AutoCloseable {
  private val mainState = ProbeValueState(protocol, analysis.recordName.orEmpty(), MAIN_KEY)
  private val recordTypeState = ProbeValueState(protocol, "${analysis.recordName.orEmpty()}.RTYP", RECORD_TYPE_KEY)
  private val coreSession = RuntimeEditorSession(project, caContext, pvaClient, listOf(mainState, recordTypeState))
  private var fieldSession: RuntimeEditorSession? = null
  private var fieldStates: List<ProbeValueState> = emptyList()
  private var fieldRecordType: String? = null

  fun start() {
    coreSession.start()
    ensureFieldSession()
  }

  fun matches(sourceKey: String, analysis: EpicsProbeDocumentAnalysis, protocol: MonitorProtocol): Boolean {
    return this.sourceKey == sourceKey &&
      this.protocol == protocol &&
      analysis.recordName == this.analysis.recordName
  }

  fun buildViewState(): EpicsProbeViewState {
    ensureFieldSession()
    val resolvedRecordType = getResolvedRecordType()
    return EpicsProbeViewState(
      recordName = analysis.recordName.orEmpty(),
      recordType = resolvedRecordType ?: "(connecting...)",
      value = mainState.displayValue,
      valueKey = mainState.key,
      valueCanPut = mainState.canPut,
      lastUpdated = mainState.lastUpdated ?: "Waiting for data",
      access = mainState.accessLabel,
      fields = fieldStates.map { state ->
        EpicsProbeFieldViewState(
          key = state.key,
          fieldName = state.fieldName,
          value = state.displayValue,
          updated = state.lastUpdated.orEmpty(),
          canPut = state.canPut,
        )
      },
      message = null,
    )
  }

  fun requestPut(key: String) {
    when (key) {
      MAIN_KEY -> coreSession.requestPutValue(mainState)
      else -> fieldStates.firstOrNull { it.key == key }?.let { state ->
        fieldSession?.requestPutValue(state)
      }
    }
  }

  override fun close() {
    fieldSession?.dispose()
    fieldSession = null
    fieldStates = emptyList()
    coreSession.dispose()
  }

  private fun ensureFieldSession() {
    val resolvedRecordType = getResolvedRecordType() ?: return
    if (fieldRecordType == resolvedRecordType && fieldSession != null) {
      return
    }

    fieldSession?.dispose()
    val fieldNames = EpicsRecordCompletionSupport.getFieldNamesForRecordType(resolvedRecordType)
      .filterNot { it.equals("RTYP", ignoreCase = true) }
      .filterNot { fieldName ->
        EpicsRecordCompletionSupport.getFieldType(resolvedRecordType, fieldName)
          ?.equals("DBF_NOACCESS", ignoreCase = true) == true
      }
    val states = fieldNames.map { fieldName ->
      ProbeValueState(
        protocol = protocol,
        pvName = "${analysis.recordName.orEmpty()}.$fieldName",
        key = "field:$fieldName",
        fieldName = fieldName,
      )
    }
    fieldStates = states
    fieldRecordType = resolvedRecordType
    fieldSession = RuntimeEditorSession(project, caContext, pvaClient, states).also(RuntimeEditorSession::start)
  }

  private fun getResolvedRecordType(): String? {
    return recordTypeState.displayValue.takeIf { it.isNotBlank() && it != CONNECTING_DISPLAY }
  }

  private companion object {
    private const val MAIN_KEY = "main"
    private const val RECORD_TYPE_KEY = "recordType"
    private const val CONNECTING_DISPLAY = "(connecting...)"
  }
}

internal class ProbeValueState(
  override val protocol: MonitorProtocol,
  override val pvName: String,
  val key: String,
  val fieldName: String = pvName.substringAfterLast('.', pvName),
) : RuntimeValueState {
  private var enumChoices: List<String> = emptyList()
  private var putInitialValue: String = ""

  @Volatile
  var displayValue: String = CONNECTING_DISPLAY
    private set

  @Volatile
  var lastUpdated: String? = null
    private set

  @Volatile
  var accessLabel: String = CONNECTING_DISPLAY
    private set

  @Volatile
  var canPut: Boolean = false
    private set

  override fun install() = Unit

  override fun matchesDoubleClick(event: com.intellij.openapi.editor.event.EditorMouseEvent): Boolean = false

  override fun setEnumChoices(choices: List<String>) {
    enumChoices = choices
  }

  override fun getEnumChoices(): List<String> = enumChoices

  override fun getPutInitialValue(): String = putInitialValue

  override fun setValue(value: String) {
    displayValue = value
    putInitialValue = value
    lastUpdated = DATE_FORMATTER.format(LocalTime.now())
  }

  override fun setConnecting() {
    displayValue = CONNECTING_DISPLAY
    accessLabel = CONNECTING_DISPLAY
    canPut = false
  }

  override fun setAccess(accessLabel: String, canPut: Boolean) {
    this.accessLabel = accessLabel
    this.canPut = canPut
  }

  override fun dispose() = Unit

  private companion object {
    private const val CONNECTING_DISPLAY = "(connecting...)"
    private val DATE_FORMATTER: DateTimeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss")
  }
}

internal data class ActiveProbeContext(
  val sourceKey: String,
  val analysis: EpicsProbeDocumentAnalysis,
)

internal fun EpicsMonitorRuntimeService.getProbeContext(editor: Editor): ActiveProbeContext? {
  val file = FileDocumentManager.getInstance().getFile(editor.document) ?: return null
  if (!isProbeFileName(file.name)) {
    return null
  }

  val analysis = EpicsProbeSupport.analyzeText(editor.document.text)
  return ActiveProbeContext(
    sourceKey = buildProbeSourceKey(file.path, analysis),
    analysis = analysis,
  )
}

internal fun EpicsMonitorRuntimeService.buildProbeSourceKey(
  filePath: String,
  analysis: EpicsProbeDocumentAnalysis,
): String {
  return "$filePath:${analysis.recordName.orEmpty()}:${defaultProtocol.name}"
}
