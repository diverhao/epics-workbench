package org.epics.workbench

import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.wm.ToolWindow
import com.intellij.openapi.wm.ToolWindowFactory
import com.intellij.ui.content.ContentFactory
import java.awt.BorderLayout
import javax.swing.JPanel
import javax.swing.JScrollPane
import javax.swing.JTextArea

class EpicsWorkbenchToolWindowFactory : ToolWindowFactory, DumbAware {
  override fun createToolWindowContent(project: Project, toolWindow: ToolWindow) {
    val textArea = JTextArea(
      """
      EPICS Workbench for JetBrains is scaffolded.

      Start building features under src/main/kotlin/org/epics/workbench.

      Suggested next steps:
      1. Add EPICS file types and syntax support.
      2. Add navigation, hover, and inspections.
      3. Add runtime CA/PVA integration.
      4. Keep IntelliJ IDEA as the primary development IDE and test the same plugin ZIP in PyCharm.
      """.trimIndent(),
    )
    textArea.isEditable = false
    textArea.lineWrap = true
    textArea.wrapStyleWord = true

    val panel = JPanel(BorderLayout())
    panel.add(JScrollPane(textArea), BorderLayout.CENTER)

    val content = ContentFactory.getInstance().createContent(panel, "", false)
    toolWindow.contentManager.addContent(content)
  }
}

