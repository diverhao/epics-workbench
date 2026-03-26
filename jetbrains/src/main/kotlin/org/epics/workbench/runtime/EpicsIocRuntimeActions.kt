package org.epics.workbench.runtime

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.components.service
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.widget.EpicsIocRuntimePageType
import org.epics.workbench.widget.openEpicsIocRuntimePage

class StartIocAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { !project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    if (!EpicsIocRuntimeService.isIocBootStartupFile(file)) {
      return
    }
    FileDocumentManager.getInstance().getDocument(file)?.let { document ->
      FileDocumentManager.getInstance().saveDocument(document)
    }
    project.service<EpicsIocRuntimeService>().startIoc(file).exceptionOrNull()?.let { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to start ${file.name}.", "Start IOC")
    }
  }
}

class StopIocAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    project.service<EpicsIocRuntimeService>().stopIoc(file)
  }
}

class ShowRunningTerminalAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    project.service<EpicsIocRuntimeService>().showRunningConsole(file)
  }
}

class OpenIocRuntimeCommandsAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.COMMANDS, file)
  }
}

class OpenIocRuntimeVariablesAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.VARIABLES, file)
  }
}

class OpenIocRuntimeEnvironmentAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val file = getTargetFile(event)
    val visible = project != null &&
      EpicsIocRuntimeService.isIocBootStartupFile(file) &&
      file?.let { project.service<EpicsIocRuntimeService>().isRunning(it) } == true
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = getTargetFile(event) ?: return
    openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.ENVIRONMENT, file)
  }
}

private fun getTargetFile(event: AnActionEvent): VirtualFile? {
  return event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
}
