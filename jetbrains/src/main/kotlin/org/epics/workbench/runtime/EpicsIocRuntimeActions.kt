package org.epics.workbench.runtime

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.progress.ProgressManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.ui.popup.JBPopupFactory
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.ui.ColoredListCellRenderer
import com.intellij.ui.SimpleTextAttributes
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSourceKind
import org.epics.workbench.widget.openEpicsPvlistWidget
import org.epics.workbench.widget.openEpicsWidget
import org.epics.workbench.widget.EpicsIocRuntimePageType
import org.epics.workbench.widget.openEpicsIocRuntimePage
import java.awt.Component
import javax.swing.DefaultListCellRenderer
import javax.swing.JList

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
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Show Running Terminal") { selectedItem ->
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "Show Running Terminal")
        ?: return@chooseRunningIocForAction
      project.service<EpicsIocRuntimeService>().showRunningConsole(startupFile)
    }
  }
}

class OpenIocRuntimeCommandsAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "IOC Runtime Commands") { selectedItem ->
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "IOC Runtime Commands")
        ?: return@chooseRunningIocForAction
      openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.COMMANDS, startupFile)
    }
  }
}

class OpenIocRuntimeVariablesAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "IOC Runtime Variables") { selectedItem ->
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "IOC Runtime Variables")
        ?: return@chooseRunningIocForAction
      openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.VARIABLES, startupFile)
    }
  }
}

class OpenIocRuntimeEnvironmentAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "IOC Runtime Environment") { selectedItem ->
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "IOC Runtime Environment")
        ?: return@chooseRunningIocForAction
      openEpicsIocRuntimePage(project, EpicsIocRuntimePageType.ENVIRONMENT, startupFile)
    }
  }
}

class DumpAllRecordsAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Dump All Records") { selectedItem ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "Dump All Records")
        ?: return@chooseRunningIocForAction

      var outputText = ""
      var failure: Throwable? = null
      val completed = ProgressManager.getInstance().runProcessWithProgressSynchronously(
        {
          try {
            outputText = runtimeService.captureCommandOutput(startupFile, "dbDumpRecord")
          } catch (error: Throwable) {
            failure = error
          }
        },
        "Dump All Records",
        true,
        project,
      )
      if (!completed) {
        return@chooseRunningIocForAction
      }
      failure?.let { error ->
        Messages.showErrorDialog(
          project,
          error.message ?: "Failed to capture dbDumpRecord output.",
          "Dump All Records",
        )
        return@chooseRunningIocForAction
      }

      runtimeService.openTemporaryOutputFile(
        buildDumpAllRecordsFileName(selectedItem),
        ensureTrailingNewline(outputText),
      )
    }
  }
}

class DumpRecordAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val editor = event.getData(CommonDataKeys.EDITOR)
    val runtimeService = project.service<EpicsIocRuntimeService>()
    chooseRunningIocForAction(event, "Dump Record") { selectedIoc ->
      val recordItems = loadDumpRecordSelectionItems(project, "Dump Record", selectedIoc) ?: return@chooseRunningIocForAction
      showDumpRecordChooser(project, editor, "Dump Record", recordItems) { selectedItem ->
        runtimeService.openTemporaryOutputFile(
          buildDumpRecordFileName(selectedItem),
          ensureTrailingNewline(selectedItem.recordEntry.dumpText),
        )
      }
    }
  }
}

class ProbeRecordAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val editor = event.getData(CommonDataKeys.EDITOR)
    chooseRunningIocForAction(event, "Probe a Record") { selectedIoc ->
      val recordItems = loadDumpRecordSelectionItems(project, "Probe a Record", selectedIoc) ?: return@chooseRunningIocForAction
      showDumpRecordChooser(project, editor, "Probe a Record", recordItems) { selectedItem ->
        openEpicsWidget(project, selectedItem.recordEntry.recordName)
      }
    }
  }
}

class PvlistAllRecordsAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "PV List All Records") { selectedItem ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "PV List All Records")
        ?: return@chooseRunningIocForAction

      var recordNames = emptyList<String>()
      var failure: Throwable? = null
      val completed = ProgressManager.getInstance().runProcessWithProgressSynchronously(
        {
          try {
            val outputText = runtimeService.captureCommandOutput(startupFile, "dbDumpRecord")
            val seenRecordNames = linkedSetOf<String>()
            EpicsIocRuntimeService.parseDbDumpRecordOutput(outputText).forEach { recordEntry ->
              val recordName = recordEntry.recordName.trim()
              if (recordName.isNotBlank()) {
                seenRecordNames += recordName
              }
            }
            recordNames = seenRecordNames.toList()
          } catch (error: Throwable) {
            failure = error
          }
        },
        "PV List All Records",
        true,
        project,
      )
      if (!completed) {
        return@chooseRunningIocForAction
      }
      failure?.let { error ->
        Messages.showErrorDialog(
          project,
          error.message ?: "Failed to capture dbDumpRecord output.",
          "PV List All Records",
        )
        return@chooseRunningIocForAction
      }
      if (recordNames.isEmpty()) {
        Messages.showInfoMessage(
          project,
          "No records were found in dbDumpRecord output.",
          "PV List All Records",
        )
        return@chooseRunningIocForAction
      }

      openEpicsPvlistWidget(
        project,
        EpicsPvlistWidgetModel(
          sourceLabel = "${selectedItem.label} (All Records)",
          sourcePath = null,
          sourceKind = EpicsPvlistWidgetSourceKind.PVLIST,
          rawPvNames = recordNames.toMutableList(),
          macroNames = mutableListOf(),
          macroValues = linkedMapOf(),
        ),
      )
    }
  }
}

private fun getTargetFile(event: AnActionEvent): VirtualFile? {
  return event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
}

private data class RunningIocSelectionItem(
  val startup: EpicsRunningIocStartup,
  val label: String,
  val detail: String,
)

private data class DumpRecordSelectionItem(
  val runningIoc: RunningIocSelectionItem,
  val recordEntry: EpicsIocDumpRecordEntry,
)

private fun hasRunningIoc(project: Project?): Boolean {
  return project?.service<EpicsIocRuntimeService>()?.listRunningIocStartups()?.isNotEmpty() == true
}

private fun collectRunningIocSelectionItems(project: Project): List<RunningIocSelectionItem> {
  return project.service<EpicsIocRuntimeService>()
    .listRunningIocStartups()
    .map { startup ->
      RunningIocSelectionItem(
        startup = startup,
        label = buildRunningIocLabel(startup.startupPath),
        detail = startup.startupPath,
      )
    }
}

private fun findContextRunningIocSelectionItem(project: Project, file: VirtualFile?): RunningIocSelectionItem? {
  if (file == null || !EpicsIocRuntimeService.isIocBootStartupFile(file)) {
    return null
  }
  val runtimeService = project.service<EpicsIocRuntimeService>()
  if (!runtimeService.isRunning(file)) {
    return null
  }
  return collectRunningIocSelectionItems(project)
    .firstOrNull { it.startup.startupPath == file.path }
}

private fun chooseRunningIocForAction(
  event: AnActionEvent,
  title: String,
  onChosen: (RunningIocSelectionItem) -> Unit,
) {
  val project = event.project ?: return
  val directItem = findContextRunningIocSelectionItem(project, getTargetFile(event))
  if (directItem != null) {
    onChosen(directItem)
    return
  }

  val runningItems = collectRunningIocSelectionItems(project)
  if (runningItems.isEmpty()) {
    Messages.showInfoMessage(project, "No running IOCs are available.", title)
    return
  }

  showRunningIocChooser(
    project = project,
    editor = event.getData(CommonDataKeys.EDITOR),
    title = title,
    items = runningItems,
    onChosen = onChosen,
  )
}

private fun resolveRunningIocStartupFile(
  project: Project,
  item: RunningIocSelectionItem,
  title: String,
): VirtualFile? {
  val startupFile = LocalFileSystem.getInstance().findFileByPath(item.startup.startupPath)
  if (startupFile == null) {
    Messages.showErrorDialog(project, "Could not resolve ${item.startup.startupPath}.", title)
  }
  return startupFile
}

private fun loadDumpRecordSelectionItems(
  project: Project,
  title: String,
  runningItem: RunningIocSelectionItem,
): List<DumpRecordSelectionItem>? {
  val runtimeService = project.service<EpicsIocRuntimeService>()
  val startupFile = resolveRunningIocStartupFile(project, runningItem, title)
    ?: return null

  var recordItems = emptyList<DumpRecordSelectionItem>()
  var failure: Throwable? = null
  val completed = ProgressManager.getInstance().runProcessWithProgressSynchronously(
    {
      try {
        val outputText = runtimeService.captureCommandOutput(startupFile, "dbDumpRecord")
        recordItems = EpicsIocRuntimeService.parseDbDumpRecordOutput(outputText).map { recordEntry ->
          DumpRecordSelectionItem(
            runningIoc = runningItem,
            recordEntry = recordEntry,
          )
        }
      } catch (error: Throwable) {
        failure = error
      }
    },
    "Loading IOC Records",
    true,
    project,
  )
  if (!completed) {
    return null
  }
  if (recordItems.isEmpty()) {
    val message = failure?.message ?: "Could not read any records from dbDumpRecord output."
    Messages.showInfoMessage(project, message, title)
    return null
  }
  return recordItems
}

private fun showRunningIocChooser(
  project: com.intellij.openapi.project.Project,
  editor: Editor?,
  title: String,
  items: List<RunningIocSelectionItem>,
  onChosen: (RunningIocSelectionItem) -> Unit,
) {
  val popup = JBPopupFactory.getInstance()
    .createPopupChooserBuilder(items)
    .setTitle(title)
    .setNamerForFiltering { item -> "${item.label} ${item.detail} ${item.startup.consoleTitle}" }
    .setRenderer(RunningIocPopupRenderer())
    .setItemChosenCallback(onChosen)
    .setResizable(true)
    .setMovable(true)
    .createPopup()

  if (editor != null) {
    popup.showInBestPositionFor(editor)
  } else {
    popup.showCenteredInCurrentWindow(project)
  }
}

private fun showDumpRecordChooser(
  project: com.intellij.openapi.project.Project,
  editor: Editor?,
  title: String,
  items: List<DumpRecordSelectionItem>,
  onChosen: (DumpRecordSelectionItem) -> Unit,
) {
  val popup = JBPopupFactory.getInstance()
    .createPopupChooserBuilder(items)
    .setTitle(title)
    .setNamerForFiltering { item ->
      "${item.recordEntry.recordName} ${item.recordEntry.recordDesc} ${item.runningIoc.label} ${item.runningIoc.detail}"
    }
    .setRenderer(DumpRecordPopupRenderer())
    .setItemChosenCallback(onChosen)
    .setResizable(true)
    .setMovable(true)
    .createPopup()

  if (editor != null) {
    popup.showInBestPositionFor(editor)
  } else {
    popup.showCenteredInCurrentWindow(project)
  }
}

private fun buildRunningIocLabel(startupPath: String): String {
  val normalizedPath = startupPath.replace('\\', '/')
  val iocBootIndex = normalizedPath.lastIndexOf("/iocBoot/")
  return when {
    iocBootIndex >= 0 -> normalizedPath.substring(iocBootIndex + 1)
    normalizedPath.isBlank() -> "IOC"
    else -> normalizedPath.substringAfterLast('/')
  }
}

private fun buildDumpAllRecordsFileName(item: RunningIocSelectionItem): String {
  return "epics-dbdumprecord-${sanitizeDumpFileToken(item.startup.startupName)}.db"
}

private fun buildDumpRecordFileName(item: DumpRecordSelectionItem): String {
  return "epics-dbdumprecord-${sanitizeDumpFileToken(item.recordEntry.recordName)}-${sanitizeDumpFileToken(item.runningIoc.startup.startupName)}.db"
}

private fun sanitizeDumpFileToken(value: String): String {
  val normalized = value.trim().lowercase().replace(Regex("[^a-z0-9._-]+"), "-").trim('-')
  return normalized.ifEmpty { "ioc" }
}

private fun ensureTrailingNewline(text: String): String {
  return if (text.endsWith('\n')) text else "$text\n"
}

private fun truncateDumpDescription(text: String, maxLength: Int = 120): String {
  val trimmed = text.trim()
  return if (trimmed.length <= maxLength) trimmed else "${trimmed.take(maxLength - 3)}..."
}

private class RunningIocPopupRenderer : ColoredListCellRenderer<RunningIocSelectionItem>() {
  override fun customizeCellRenderer(
    list: JList<out RunningIocSelectionItem>,
    value: RunningIocSelectionItem?,
    index: Int,
    selected: Boolean,
    hasFocus: Boolean,
  ) {
    val item = value ?: return
    append(item.label, SimpleTextAttributes.REGULAR_ATTRIBUTES)
    append("   ${item.startup.consoleTitle}", SimpleTextAttributes.GRAY_ATTRIBUTES)
    append("   ${item.detail}", SimpleTextAttributes.GRAYED_ATTRIBUTES)
  }
}

private class DumpRecordPopupRenderer : DefaultListCellRenderer() {
  override fun getListCellRendererComponent(
    list: JList<*>,
    value: Any?,
    index: Int,
    isSelected: Boolean,
    cellHasFocus: Boolean,
  ): Component {
    val component = super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus)
    val label = component as? javax.swing.JLabel ?: return component
    val item = value as? DumpRecordSelectionItem ?: return component
    val descText = truncateDumpDescription(item.recordEntry.recordDesc.ifBlank { "(No DESC)" })
    label.text =
      """
      <html>
        <div><b>${escapeHtml(item.recordEntry.recordName)}</b></div>
        <div>${escapeHtml(descText)}</div>
        <div>${escapeHtml(item.runningIoc.label)}</div>
      </html>
      """.trimIndent()
    return label
  }
}

private fun escapeHtml(text: String): String {
  return text
    .replace("&", "&amp;")
    .replace("<", "&lt;")
    .replace(">", "&gt;")
}
