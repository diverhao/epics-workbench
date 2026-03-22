package org.epics.workbench.runtime

import com.intellij.openapi.components.service
import com.intellij.openapi.options.ConfigurationException
import com.intellij.openapi.options.SearchableConfigurable
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.ComboBox
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBTextArea
import com.intellij.util.ui.FormBuilder
import java.awt.BorderLayout
import javax.swing.JComponent
import javax.swing.JPanel
import javax.swing.JScrollPane

class EpicsRuntimeProjectConfigurable(
  private val project: Project,
) : SearchableConfigurable {
  private val configurationService = project.service<EpicsRuntimeProjectConfigurationService>()
  private var component: JPanel? = null
  private lateinit var protocolCombo: ComboBox<EpicsRuntimeProtocol>
  private lateinit var caAutoAddrListCombo: ComboBox<EpicsCaAutoAddrList>
  private lateinit var caAddrListArea: JBTextArea

  override fun getId(): String = "org.epics.workbench.runtime.configuration"

  override fun getDisplayName(): String = "EPICS Workbench"

  override fun createComponent(): JComponent {
    if (component != null) {
      return component as JPanel
    }

    protocolCombo = ComboBox(EpicsRuntimeProtocol.entries.toTypedArray())
    caAutoAddrListCombo = ComboBox(EpicsCaAutoAddrList.entries.toTypedArray())
    caAddrListArea = JBTextArea(5, 32).apply {
      lineWrap = true
      wrapStyleWord = true
      emptyText.text = "One address per line"
    }

    val note = buildString {
      append("Saved to ")
      append(EpicsRuntimeProjectConfigurationService.CONFIG_FILE_NAME)
      append(" in the current project root.")
    }

    component = FormBuilder.createFormBuilder()
      .addLabeledComponent("Protocol", protocolCombo)
      .addLabeledComponent("EPICS_CA_AUTO_ADDR_LIST", caAutoAddrListCombo)
      .addLabeledComponent(
        "EPICS_CA_ADDR_LIST",
        JScrollPane(caAddrListArea),
        true,
      )
      .addComponent(
        JPanel(BorderLayout()).apply {
          add(JBLabel(note), BorderLayout.WEST)
        },
      )
      .panel

    reset()
    return component as JPanel
  }

  override fun isModified(): Boolean {
    if (component == null) {
      return false
    }
    return readConfigurationFromUi() != configurationService.loadConfiguration()
  }

  override fun apply() {
    try {
      configurationService.saveConfiguration(readConfigurationFromUi())
      project.service<EpicsMonitorRuntimeService>().restartMonitoringIfActive()
    } catch (error: Exception) {
      throw ConfigurationException(
        error.message ?: "Failed to save EPICS Workbench configuration.",
        "EPICS Workbench",
      )
    }
  }

  override fun reset() {
    if (component == null) {
      return
    }
    val configuration = configurationService.loadConfiguration()
    protocolCombo.selectedItem = configuration.protocol
    caAutoAddrListCombo.selectedItem = configuration.caAutoAddrList
    caAddrListArea.text = configuration.caAddrList.joinToString("\n")
    caAddrListArea.caretPosition = 0
  }

  override fun disposeUIResources() {
    component = null
  }

  private fun readConfigurationFromUi(): EpicsRuntimeProjectConfiguration {
    return EpicsRuntimeProjectConfiguration(
      protocol = protocolCombo.item ?: EpicsRuntimeProtocol.CA,
      caAddrList = caAddrListArea.text
        .lineSequence()
        .map(String::trim)
        .filter(String::isNotEmpty)
        .toList(),
      caAutoAddrList = caAutoAddrListCombo.item ?: EpicsCaAutoAddrList.YES,
    )
  }
}
