package org.epics.workbench.runtime

import gov.aps.jca.Context
import org.epics.pva.client.PVAClient
import org.epics.workbench.pvlist.EpicsPvlistWidgetDefinition
import org.epics.workbench.pvlist.EpicsPvlistWidgetPlan
import org.epics.workbench.pvlist.EpicsPvlistWidgetRowPlan
import com.intellij.openapi.project.Project
import java.util.concurrent.ConcurrentHashMap

internal data class EpicsPvlistWidgetRowViewState(
  val channelName: String,
  val value: String,
  val definitionKey: String? = null,
  val canPut: Boolean = false,
)

internal data class EpicsPvlistWidgetViewState(
  val rows: List<EpicsPvlistWidgetRowViewState>,
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

  fun buildViewState(plan: EpicsPvlistWidgetPlan, monitoringActive: Boolean): EpicsPvlistWidgetViewState {
    val rows = plan.rows.map { row -> buildRowViewState(row, monitoringActive) }
    val message = when {
      rows.isEmpty() -> "Add channels to start monitoring."
      else -> null
    }
    return EpicsPvlistWidgetViewState(rows = rows, message = message)
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
        channelName = row.channelName,
        value = row.unresolvedValue ?: if (monitoringActive) "(connecting...)" else "(stopped)",
      )
    }

    val state = statesByKey[row.definitionKey]
    val value = when {
      !monitoringActive -> "(stopped)"
      state == null -> "(connecting...)"
      else -> state.displayValue
    }
    return EpicsPvlistWidgetRowViewState(
      channelName = row.channelName,
      value = value,
      definitionKey = row.definitionKey,
      canPut = monitoringActive && state?.canPut == true,
    )
  }

  private fun disposeKey(key: String) {
    sessionsByKey.remove(key)?.dispose()
    statesByKey.remove(key)?.dispose()
  }
}
