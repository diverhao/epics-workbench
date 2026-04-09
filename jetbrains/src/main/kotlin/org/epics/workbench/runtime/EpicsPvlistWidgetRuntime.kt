package org.epics.workbench.runtime

import gov.aps.jca.Context
import org.epics.pva.client.PVAClient
import org.epics.workbench.pvlist.EpicsPvlistWidgetDefinition
import org.epics.workbench.pvlist.EpicsPvlistWidgetPlan
import org.epics.workbench.pvlist.EpicsPvlistWidgetRowPlan
import org.epics.workbench.pvlist.EpicsPvlistWidgetSupport
import com.intellij.openapi.project.Project
import java.util.concurrent.ConcurrentHashMap

internal data class EpicsPvlistWidgetRowViewState(
  val sourceIndex: Int,
  val channelName: String,
  val probeTargetRecordName: String? = null,
  val recordName: String? = null,
  val protocol: MonitorProtocol? = null,
  val recordType: String = "",
  val value: String,
  val definitionKey: String? = null,
  val canPut: Boolean = false,
  val canProcess: Boolean = false,
  val fieldCells: List<EpicsPvlistWidgetFieldCellViewState> = emptyList(),
)

internal data class EpicsPvlistWidgetFieldCellViewState(
  val name: String,
  val value: String,
)

internal data class EpicsPvlistWidgetViewState(
  val rows: List<EpicsPvlistWidgetRowViewState>,
  val fieldColumns: List<String>,
  val message: String? = null,
)

internal class EpicsPvlistWidgetSession(
  private val project: Project,
  private val caContext: Context,
  private val pvaClient: PVAClient,
) : AutoCloseable {
  private val statesByKey = ConcurrentHashMap<String, ProbeValueState>()
  private val sessionsByKey = ConcurrentHashMap<String, RuntimeEditorSession>()

  fun updatePlan(plan: EpicsPvlistWidgetPlan) {
    val desiredByKey = plan.definitions.associateBy(EpicsPvlistWidgetDefinition::key)

    sessionsByKey.keys
      .filterNot(desiredByKey::containsKey)
      .forEach(::disposeKey)

    desiredByKey.forEach { (key, definition) ->
      if (statesByKey.containsKey(key)) {
        return@forEach
      }

      val state = ProbeValueState(
        protocol = definition.protocol,
        pvName = definition.pvName,
        key = key,
        fieldName = definition.pvName,
      )
      val session = RuntimeEditorSession(project, caContext, pvaClient, listOf(state))
      statesByKey[key] = state
      sessionsByKey[key] = session
      session.start()
    }
  }

  fun buildViewState(
    plan: EpicsPvlistWidgetPlan,
    monitoringActive: Boolean,
    fieldColumns: List<String>,
  ): EpicsPvlistWidgetViewState {
    val rows = plan.rows.map { row -> buildRowViewState(row, monitoringActive) }
    val message = when {
      rows.isEmpty() -> "Add channels to start monitoring."
      else -> null
    }
    return EpicsPvlistWidgetViewState(rows = rows, fieldColumns = fieldColumns, message = message)
  }

  fun getResolvedRecordTypes(): Map<String, String> {
    val resolvedTypes = linkedMapOf<String, String>()
    statesByKey.values.forEach { state ->
      if (!state.pvName.endsWith(".RTYP", ignoreCase = true)) {
        return@forEach
      }
      val recordName = EpicsPvlistWidgetSupport.getRecordName(state.pvName).trim()
      val recordType = state.displayValue.trim()
      if (recordName.isBlank() || recordType.isBlank() || recordType == CONNECTING_DISPLAY) {
        return@forEach
      }
      resolvedTypes[recordName] = recordType
    }
    return resolvedTypes
  }

  override fun close() {
    sessionsByKey.keys.toList().forEach(::disposeKey)
  }

  private fun buildRowViewState(
    row: EpicsPvlistWidgetRowPlan,
    monitoringActive: Boolean,
  ): EpicsPvlistWidgetRowViewState {
    if (row.definitionKey == null) {
      return EpicsPvlistWidgetRowViewState(
        sourceIndex = row.sourceIndex,
        channelName = row.channelName,
        value = row.unresolvedValue ?: if (monitoringActive) "(connecting...)" else "(stopped)",
        recordType = row.recordType.orEmpty(),
        fieldCells = row.fieldCells.map { fieldCell ->
          EpicsPvlistWidgetFieldCellViewState(
            name = fieldCell.name,
            value = fieldCell.value,
          )
        },
      )
    }

    val state = statesByKey[row.definitionKey]
    val recordTypeState = row.recordTypeDefinitionKey?.let(statesByKey::get)
    val value = getDisplayValue(state, showConnectingText = true, monitoringActive = monitoringActive)
    val recordType = row.recordType
      ?: getDisplayValue(recordTypeState, showConnectingText = false, monitoringActive = monitoringActive)
    return EpicsPvlistWidgetRowViewState(
      sourceIndex = row.sourceIndex,
      channelName = row.channelName,
      probeTargetRecordName = row.recordName?.takeIf(String::isNotBlank) ?: row.channelName.trim().takeIf(String::isNotBlank),
      recordName = row.recordName,
      protocol = row.protocol,
      recordType = recordType,
      value = value,
      definitionKey = row.definitionKey,
      canPut = monitoringActive && state?.canPut == true,
      canProcess = monitoringActive && !row.recordName.isNullOrBlank(),
      fieldCells = row.fieldCells.map { fieldCell ->
        val fieldState = fieldCell.definitionKey?.let(statesByKey::get)
        EpicsPvlistWidgetFieldCellViewState(
          name = fieldCell.name,
          value = fieldCell.value.ifBlank {
            getDisplayValue(fieldState, showConnectingText = false, monitoringActive = monitoringActive)
          },
        )
      },
    )
  }

  private fun getDisplayValue(
    state: ProbeValueState?,
    showConnectingText: Boolean,
    monitoringActive: Boolean,
  ): String {
    val connectingText = if (showConnectingText) "(Connecting)" else ""
    if (!monitoringActive) {
      return if (showConnectingText) "(stopped)" else ""
    }
    if (state == null) {
      return connectingText
    }
    return if (state.displayValue == CONNECTING_DISPLAY) {
      connectingText
    } else {
      state.displayValue
    }
  }

  private fun disposeKey(key: String) {
    sessionsByKey.remove(key)?.dispose()
    statesByKey.remove(key)?.dispose()
  }

  private companion object {
    private const val CONNECTING_DISPLAY = "(connecting...)"
  }
}
