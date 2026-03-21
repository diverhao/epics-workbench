package org.epics.workbench

import com.intellij.openapi.Disposable
import com.intellij.openapi.components.service
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.wm.ToolWindow
import com.intellij.openapi.wm.ToolWindowFactory
import com.intellij.ui.content.ContentFactory
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsMonitorRuntimeStateListener
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

  init {
    runtimeService.initialize()

    val controlsPanel = JPanel(FlowLayout(FlowLayout.LEFT, 8, 8))
    controlsPanel.add(JLabel("Monitor Runtime:"))
    controlsPanel.add(statusLabel)
    controlsPanel.add(toggleButton)

    toggleButton.addActionListener {
      runtimeService.toggleMonitoring()
    }

    val textArea = JTextArea(
      """
      Use Start Monitoring to create the EPICS CA/PVA context for this project.

      While running:
      - `.monitor` files connect and show live channel values after each line.
      - Stop Monitoring disposes the active sessions and destroys the EPICS context.
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
