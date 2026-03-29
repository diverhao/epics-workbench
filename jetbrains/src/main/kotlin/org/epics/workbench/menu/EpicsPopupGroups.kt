package org.epics.workbench.menu

import com.intellij.openapi.actionSystem.ActionManager
import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnAction
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.DefaultActionGroup
import com.intellij.openapi.actionSystem.Separator
import com.intellij.openapi.actionSystem.ex.ActionUtil
import com.intellij.openapi.project.DumbAware

abstract class AbstractEpicsPopupGroup(
  private vararg val childActionIds: String,
) : DefaultActionGroup(), DumbAware {
  init {
    templatePresentation.setHideGroupIfEmpty(true)
    templatePresentation.setDisableGroupIfEmpty(false)
    templatePresentation.putClientProperty(ActionUtil.HIDE_DISABLED_CHILDREN, true)
  }

  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun getChildren(event: AnActionEvent?): Array<AnAction> {
    val actionManager = ActionManager.getInstance()
    return childActionIds.mapNotNull { actionId ->
      when (actionId) {
        SEPARATOR -> Separator.create()
        else -> actionManager.getAction(actionId)
      }
    }.toTypedArray()
  }

  companion object {
    const val SEPARATOR: String = "__separator__"
  }
}

class EpicsBuildPopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.BuildWithMakefileAction",
  "org.epics.workbench.CleanAndBuildWithMakefileAction",
  SEPARATOR,
  "org.epics.workbench.ShowMakefileAction",
  "org.epics.workbench.ShowReleaseAction",
  SEPARATOR,
  "org.epics.workbench.BuildProjectAction",
  "org.epics.workbench.CleanAndBuildProjectAction",
  "org.epics.workbench.DistCleanProjectAction",
)

class EpicsWidgetPopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.OpenInProbeAction",
  "org.epics.workbench.OpenInPvlistAction",
  "org.epics.workbench.OpenInMonitorAction",
)

class EpicsImportExportPopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.ExportDatabaseToExcelAction",
  "org.epics.workbench.ImportDatabaseFromExcelAction",
)

class EpicsFormatPopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.FormatDatabaseFileAction",
)

class EpicsDatabaseFilePopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.AddDbToMakefileAction",
  SEPARATOR,
  "org.epics.workbench.UpdateDatabaseTocAction",
  "org.epics.workbench.ToggleMonitorChannelsAction",
  SEPARATOR,
  "org.epics.workbench.CopyAllRecordNamesAction",
  "org.epics.workbench.ExpandAllRecordsAction",
  "org.epics.workbench.CollapseAllRecordsAction",
)

class EpicsIocRuntimePopupGroup : AbstractEpicsPopupGroup(
  "org.epics.workbench.ShowRunningTerminalAction",
  "org.epics.workbench.OpenIocRuntimeCommandsAction",
  "org.epics.workbench.OpenIocRuntimeVariablesAction",
  "org.epics.workbench.OpenIocRuntimeEnvironmentAction",
)
