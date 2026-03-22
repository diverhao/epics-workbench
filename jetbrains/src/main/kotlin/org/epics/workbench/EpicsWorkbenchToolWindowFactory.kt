package org.epics.workbench

import com.intellij.openapi.Disposable
import com.intellij.openapi.components.service
import com.intellij.openapi.options.ShowSettingsUtil
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.wm.ToolWindow
import com.intellij.openapi.wm.ToolWindowFactory
import com.intellij.ui.content.ContentFactory
import org.epics.workbench.export.openEpicsExcelImportPreview
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsMonitorRuntimeStateListener
import org.epics.workbench.runtime.EpicsRuntimeProjectConfigurable
import java.awt.BorderLayout
import java.awt.FlowLayout
import javax.swing.JButton
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.JScrollPane
import javax.swing.JTextArea

class EpicsWorkbenchToolWindowFactory : ToolWindowFactory, DumbAware {
  override fun createToolWindowContent(project: Project, toolWindow: ToolWindow) {
    val panel = EpicsWorkbenchToolWindowPanel(project)
    val content = ContentFactory.getInstance().createContent(panel, "", false)
    content.setDisposer(panel)
    toolWindow.contentManager.addContent(content)
  }
}

private class EpicsWorkbenchToolWindowPanel(
  private val project: Project,
) : JPanel(BorderLayout()), Disposable {
  private val runtimeService = project.service<EpicsMonitorRuntimeService>()
  private val statusLabel = JLabel()
  private val toggleButton = JButton()
  private val importPreviewButton = JButton("Excel Import Preview...")
  private val configureButton = JButton("Configuration...")

  init {
    runtimeService.initialize()

    val controlsPanel = JPanel(FlowLayout(FlowLayout.LEFT, 8, 8))
    controlsPanel.add(JLabel("Monitor Runtime:"))
    controlsPanel.add(statusLabel)
    controlsPanel.add(toggleButton)
    controlsPanel.add(importPreviewButton)
    controlsPanel.add(configureButton)

    toggleButton.addActionListener {
      runtimeService.toggleMonitoring()
    }
    importPreviewButton.addActionListener {
      openEpicsExcelImportPreview(project)
    }
    configureButton.addActionListener {
      ShowSettingsUtil.getInstance().showSettingsDialog(project, EpicsRuntimeProjectConfigurable::class.java)
    }

    val textArea = JTextArea(
      """
      Use Start Monitoring to create the EPICS CA/PVA context for this project.

      While running:
      - `.pvlist` files connect and show live channel values after each line.
      - Database TOCs show live values in the `Value` column.
      - `.probe` files show a live record page inline in the editor.
      - Stop Monitoring disposes the active sessions and destroys the EPICS context.
      - Use Configuration... to set the default protocol and CA address settings.
      - Use Excel Import Preview... to drag an EPICS Excel workbook into a preview tab and open generated database tabs.
      - Use Probe from database/startup editor context menus to open a file-less EPICS probe widget.
      """.trimIndent(),
    )
    textArea.isEditable = false
    textArea.lineWrap = true
    textArea.wrapStyleWord = true

    add(controlsPanel, BorderLayout.NORTH)
    add(JScrollPane(textArea), BorderLayout.CENTER)

    val connection = project.messageBus.connect(this)
    connection.subscribe(
      EpicsMonitorRuntimeStateListener.TOPIC,
      object : EpicsMonitorRuntimeStateListener {
        override fun monitoringStateChanged(active: Boolean) {
          updateState(active)
        }
      },
    )

    updateState(runtimeService.isMonitoringActive())
  }

  private fun updateState(active: Boolean) {
    statusLabel.text = if (active) "Running" else "Stopped"
    toggleButton.text = if (active) "Stop Monitoring" else "Start Monitoring"
  }

  override fun dispose() = Unit
}
