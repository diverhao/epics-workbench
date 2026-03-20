package org.epics.workbench

import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages

class ShowEpicsWorkbenchStatusAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    Messages.showInfoMessage(
      event.project,
      "EPICS Workbench plugin scaffold is loaded.",
      "EPICS Workbench",
    )
  }
}

