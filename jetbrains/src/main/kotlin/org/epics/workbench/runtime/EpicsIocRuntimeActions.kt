package org.epics.workbench.runtime

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.components.service
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.progress.ProgressManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.DialogWrapper
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.ui.popup.JBPopupFactory
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.ui.ColoredListCellRenderer
import com.intellij.ui.SimpleTextAttributes
import com.intellij.ui.components.JBCheckBox
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBList
import com.intellij.ui.components.JBScrollPane
import com.intellij.ui.components.JBTextField
import com.intellij.util.ui.JBUI
import org.epics.workbench.export.exportDatabaseTextToExcel
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSourceKind
import org.epics.workbench.widget.EpicsIocRuntimePageType
import org.epics.workbench.widget.openEpicsIocRuntimePage
import org.epics.workbench.widget.openEpicsPvlistWidget
import org.epics.workbench.widget.openEpicsWidget
import java.awt.BorderLayout
import java.awt.Component
import java.awt.Dimension
import java.awt.Font
import java.awt.event.ActionEvent
import java.awt.event.KeyEvent
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.nio.file.Files
import java.nio.file.Path
import javax.swing.Action
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.DefaultListCellRenderer
import javax.swing.DefaultListModel
import javax.swing.JComponent
import javax.swing.JList
import javax.swing.JPanel
import javax.swing.KeyStroke
import javax.swing.ListSelectionModel

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

class DumpAllRecordNamesAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Dump All Record Names") { selectedItem ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "Dump All Record Names")
        ?: return@chooseRunningIocForAction

      var outputText = ""
      var failure: Throwable? = null
      val completed = ProgressManager.getInstance().runProcessWithProgressSynchronously(
        {
          try {
            outputText = runtimeService.captureCommandOutput(startupFile, "dbl")
          } catch (error: Throwable) {
            failure = error
          }
        },
        "Dump All Record Names",
        true,
        project,
      )
      if (!completed) {
        return@chooseRunningIocForAction
      }
      failure?.let { error ->
        Messages.showErrorDialog(
          project,
          error.message ?: "Failed to capture dbl output.",
          "Dump All Record Names",
        )
        return@chooseRunningIocForAction
      }

      runtimeService.openTemporaryPvlistOutputFile(
        buildDumpAllRecordNamesFileName(selectedItem),
        ensureTrailingNewline(outputText),
      )
    }
  }
}

class DumpAllRecordsToExcelAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val visible = project != null &&
      project.service<EpicsIocRuntimeService>().listRunningIocStartups().isNotEmpty()
    event.presentation.isEnabledAndVisible = visible
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Dump All Records to Excel") { selectedItem ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedItem, "Dump All Records to Excel")
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
        "Dump All Records to Excel",
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
          "Dump All Records to Excel",
        )
        return@chooseRunningIocForAction
      }

      exportDatabaseTextToExcel(
        project = project,
        sourceText = ensureTrailingNewline(outputText),
        defaultName = buildDumpAllRecordsExcelFileName(selectedItem),
        initialDirectory = Path.of(startupFile.path).parent,
        sourceLabel = selectedItem.label,
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
    chooseRunningIocForAction(event, "Probe a Record") { selectedIoc ->
      val recordItems = loadDumpRecordSelectionItems(project, "Probe a Record", selectedIoc) ?: return@chooseRunningIocForAction
      ProbeRecordDialog(project, selectedIoc, recordItems).show()
    }
  }
}

class RunCommandAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Run Command") { selectedIoc ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedIoc, "Run Command")
        ?: return@chooseRunningIocForAction

      var commandNames = emptyList<String>()
      var helpByCommand = emptyMap<String, String>()
      var failure: Throwable? = null
      val completed = ProgressManager.getInstance().runProcessWithProgressSynchronously(
        {
          try {
            commandNames = runtimeService.getCommandNames(startupFile)
            helpByCommand = runtimeService.getCommandHelp(startupFile, commandNames)
          } catch (error: Throwable) {
            failure = error
          }
        },
        "Run Command",
        true,
        project,
      )
      if (!completed) {
        return@chooseRunningIocForAction
      }
      failure?.let { error ->
        Messages.showErrorDialog(
          project,
          error.message ?: "Failed to load IOC runtime commands.",
          "Run Command",
        )
        return@chooseRunningIocForAction
      }

      RunIocCommandDialog(
        project = project,
        startupFile = startupFile,
        runningIoc = selectedIoc,
        commandNames = commandNames,
        helpByCommand = helpByCommand,
      ).show()
    }
  }
}

class CommandHistoryAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = hasRunningIoc(event.project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    chooseRunningIocForAction(event, "Command History") { selectedIoc ->
      val runtimeService = project.service<EpicsIocRuntimeService>()
      val startupFile = resolveRunningIocStartupFile(project, selectedIoc, "Command History")
        ?: return@chooseRunningIocForAction
      CommandHistoryDialog(
        project = project,
        startupFile = startupFile,
        runningIoc = selectedIoc,
        historyEntries = runtimeService.readCommandHistory(),
      ).show()
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

private val RUN_IOC_RECORD_ARGUMENT_COMMANDS = setOf(
  "dbpf",
  "dba",
  "dbb",
  "dbap",
  "dbc",
  "dbcar",
  "dbd",
  "dbel",
  "dbgf",
  "dbgrep",
  "dbjlr",
  "dblsr",
  "dbp",
  "dbpr",
  "dbs",
  "dbtgf",
  "dbtpf",
  "dbtpn",
  "dbtr",
  "gft",
  "pft",
  "tpn",
)

private val RUN_IOC_ENVIRONMENT_ARGUMENT_COMMANDS = setOf(
  "epicsenvset",
  "epicsenvshow",
  "epicsenvunset",
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
  if (runningItems.size == 1) {
    onChosen(runningItems.single())
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
    .setNamerForFiltering { item -> "${item.recordEntry.recordName} ${item.recordEntry.recordDesc}" }
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

private data class RunIocCommandListItem(
  val label: String,
  val description: String,
  val detail: String,
  val insertText: String,
)

private data class RunIocCommandContext(
  val commandName: String,
  val suggestionType: String,
  val partialText: String,
)

private class RunIocCommandDialog(
  private val project: com.intellij.openapi.project.Project,
  private val startupFile: VirtualFile,
  private val runningIoc: RunningIocSelectionItem,
  private val commandNames: List<String>,
  private val helpByCommand: Map<String, String>,
) : DialogWrapper(project, true) {
  private val runtimeService = project.service<EpicsIocRuntimeService>()
  private val commandField = JBTextField().apply {
    putClientProperty("JTextField.placeholderText", "Type a command")
  }
  private val captureOutputCheck = JBCheckBox("Capture output")
  private val sendButton = JButton("Send")
  private val statusLabel = JBLabel("Type a command or click one below to fill the input.")
  private val itemModel = DefaultListModel<RunIocCommandListItem>()
  private val itemList = JBList(itemModel)
  private val closeAction = DialogWrapperExitAction("Close", CANCEL_EXIT_CODE)
  private var recordNames: List<String>? = null
  private var environmentNames: List<String>? = null
  private var loadingRecordNames = false
  private var loadingEnvironmentNames = false

  init {
    title = "Run Command"
    setResizable(true)
    commandField.focusTraversalKeysEnabled = false
    commandField.registerKeyboardAction(
      { applySelectedItem() },
      KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0),
      JComponent.WHEN_FOCUSED,
    )
    itemList.selectionMode = ListSelectionModel.SINGLE_SELECTION
    itemList.visibleRowCount = 16
    itemList.cellRenderer = RunIocCommandListRenderer()
    itemList.addMouseListener(
      object : MouseAdapter() {
        override fun mouseClicked(event: MouseEvent) {
          if (event.clickCount >= 1) {
            applySelectedItem()
          }
        }
      },
    )
    installDialogDocumentListener(commandField) {
      updateSendState()
      rebuildItems()
    }
    sendButton.addActionListener { sendCommand() }
    init()
    updateSendState()
    rebuildItems()
  }

  override fun createCenterPanel(): JComponent {
    val titleLabel = JBLabel("Type a command or click one below to fill the input.").apply {
      font = font.deriveFont(Font.BOLD, font.size2D + 2f)
    }
    val metaLabel = JBLabel(runningIoc.label).apply {
      foreground = SimpleTextAttributes.GRAYED_ATTRIBUTES.fgColor
    }
    val inputRow = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.X_AXIS)
      isOpaque = false
      alignmentX = Component.LEFT_ALIGNMENT
      commandField.columns = 36
      commandField.maximumSize = Dimension(Int.MAX_VALUE, commandField.preferredSize.height)
      add(commandField)
      add(Box.createHorizontalStrut(10))
      add(captureOutputCheck)
      add(Box.createHorizontalStrut(10))
      add(sendButton)
    }
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.empty(0, 0, 10, 0)
      add(titleLabel)
      add(Box.createVerticalStrut(4))
      add(metaLabel)
      add(Box.createVerticalStrut(10))
      add(inputRow)
      add(Box.createVerticalStrut(8))
      add(statusLabel)
    }
    return JPanel(BorderLayout(0, 10)).apply {
      preferredSize = Dimension(920, 520)
      add(topPanel, BorderLayout.NORTH)
      add(JBScrollPane(itemList), BorderLayout.CENTER)
    }
  }

  override fun createActions(): Array<Action> {
    return arrayOf(closeAction)
  }

  private fun updateSendState() {
    sendButton.isEnabled = commandField.text.trim().isNotEmpty()
  }

  private fun applySelectedItem() {
    val selectedItem = itemList.selectedValue ?: return
    commandField.text = selectedItem.insertText
    commandField.requestFocusInWindow()
    commandField.caretPosition = commandField.text.length
  }

  private fun rebuildItems() {
    val inputText = commandField.text
    val items = mutableListOf<RunIocCommandListItem>()
    val context = parseRunIocCommandContext(inputText)
    when (context?.suggestionType) {
      "record" -> {
        val cachedRecordNames = recordNames
        if (cachedRecordNames != null) {
          cachedRecordNames
            .asSequence()
            .filter { it.contains(context.partialText, ignoreCase = true) }
            .take(200)
            .forEach { recordName ->
              items += RunIocCommandListItem(
                label = recordName,
                description = "",
                detail = "",
                insertText = "${context.commandName} $recordName",
              )
            }
        } else {
          items += RunIocCommandListItem(
            label = "Loading record names...",
            description = "dbl",
            detail = runningIoc.label,
            insertText = inputText,
          )
          ensureRecordNamesLoaded()
        }
      }
      "environment" -> {
        val cachedEnvironmentNames = environmentNames
        if (cachedEnvironmentNames != null) {
          cachedEnvironmentNames
            .asSequence()
            .filter { it.contains(context.partialText, ignoreCase = true) }
            .take(200)
            .forEach { variableName ->
              items += RunIocCommandListItem(
                label = variableName,
                description = "Environment variable for ${context.commandName}",
                detail = runningIoc.label,
                insertText = "${context.commandName} $variableName",
              )
            }
        } else {
          items += RunIocCommandListItem(
            label = "Loading environment variables...",
            description = "epicsEnvShow",
            detail = runningIoc.label,
            insertText = inputText,
          )
          ensureEnvironmentNamesLoaded()
        }
      }
      "path" -> {
        val workingDirectory = runCatching {
          runtimeService.getWorkingDirectory(startupFile)
        }.getOrElse {
          Path.of(startupFile.parent?.path ?: project.basePath.orEmpty())
        }
        collectRunIocPathSuggestions(workingDirectory, context.partialText)
          .asSequence()
          .take(120)
          .forEach { pathText ->
            items += RunIocCommandListItem(
              label = pathText,
              description = "Directory",
              detail = workingDirectory.toString(),
              insertText = "cd ${formatRunIocPathSuggestion(pathText)}",
            )
          }
      }
    }

    val commandQuery = inputText.trimStart().substringBefore(' ', "").lowercase()
    commandNames
      .asSequence()
      .filter { commandName ->
        if (commandQuery.isBlank()) {
          true
        } else {
          "$commandName ${extractRunCommandSummary(helpByCommand[commandName].orEmpty())}"
            .lowercase()
            .contains(commandQuery)
        }
      }
      .take(240)
      .forEach { commandName ->
        items += RunIocCommandListItem(
          label = commandName,
          description = "",
          detail = extractRunCommandSummary(helpByCommand[commandName].orEmpty()),
          insertText = commandName,
        )
      }

    itemModel.removeAllElements()
    if (items.isEmpty()) {
      itemModel.addElement(
        RunIocCommandListItem(
          label = "No commands or suggestions match the current input.",
          description = "",
          detail = "",
          insertText = inputText,
        ),
      )
    } else {
      items.forEach(itemModel::addElement)
    }
    if (itemModel.size > 0) {
      itemList.selectedIndex = 0
    }
  }

  private fun ensureRecordNamesLoaded() {
    if (recordNames != null || loadingRecordNames) {
      return
    }
    loadingRecordNames = true
    ApplicationManager.getApplication().executeOnPooledThread {
      val loadedRecordNames = runCatching {
        parseDblRecordNames(
          runtimeService.captureCommandOutput(startupFile, "dbl", recordHistory = false),
        )
      }.getOrDefault(emptyList())
      ApplicationManager.getApplication().invokeLater {
        recordNames = loadedRecordNames
        loadingRecordNames = false
        rebuildItems()
      }
    }
  }

  private fun ensureEnvironmentNamesLoaded() {
    if (environmentNames != null || loadingEnvironmentNames) {
      return
    }
    loadingEnvironmentNames = true
    ApplicationManager.getApplication().executeOnPooledThread {
      val loadedEnvironmentNames = runCatching {
        runtimeService.listRuntimeEnvironment(startupFile)
          .map(EpicsIocRuntimeEnvironmentEntry::name)
          .distinct()
      }.getOrDefault(emptyList())
      ApplicationManager.getApplication().invokeLater {
        environmentNames = loadedEnvironmentNames
        loadingEnvironmentNames = false
        rebuildItems()
      }
    }
  }

  private fun sendCommand() {
    val rawCommandText = commandField.text.trim()
    if (rawCommandText.isEmpty()) {
      return
    }

    sendButton.isEnabled = false
    statusLabel.text = if (captureOutputCheck.isSelected) {
      "Capturing output for $rawCommandText"
    } else {
      "Sending $rawCommandText"
    }
    ApplicationManager.getApplication().executeOnPooledThread {
      val result = runCatching {
        val normalizedCommandText = normalizeRunIocCommandForExecution(
          rawCommandText,
          runtimeService.getCommandNames(startupFile),
        )
        if (captureOutputCheck.isSelected) {
          val output = runtimeService.captureCommandOutput(
            startupFile,
            normalizedCommandText,
            historyCommandText = rawCommandText,
          )
          runtimeService.openCapturedOutput(
            startupFile,
            "ioc-command",
            normalizedCommandText,
            output,
          )
        } else {
          runtimeService.sendCommandText(
            startupFile,
            normalizedCommandText,
            historyCommandText = rawCommandText,
          )
        }
      }
      ApplicationManager.getApplication().invokeLater {
        sendButton.isEnabled = true
        updateSendState()
        result.exceptionOrNull()?.let { error ->
          statusLabel.text = error.message ?: "Failed to run IOC command."
          Messages.showErrorDialog(project, statusLabel.text, "Run Command")
          return@invokeLater
        }
        recordNames = null
        environmentNames = null
        statusLabel.text = "Sent ${truncateDumpDescription(rawCommandText, 120)}"
        rebuildItems()
      }
    }
  }
}

private class CommandHistoryDialog(
  private val project: com.intellij.openapi.project.Project,
  private val startupFile: VirtualFile,
  private val runningIoc: RunningIocSelectionItem,
  historyEntries: List<String>,
) : DialogWrapper(project, true) {
  private val runtimeService = project.service<EpicsIocRuntimeService>()
  private val captureOutputCheck = JBCheckBox("Capture output")
  private val commandField = JBTextField().apply {
    putClientProperty("JTextField.placeholderText", "Filter command history, press Tab or click to fill, Enter to send.")
  }
  private val statusLabel = JBLabel("Newest commands first.")
  private val itemModel = DefaultListModel<RunIocCommandListItem>()
  private val itemList = JBList(itemModel)
  private val closeAction = DialogWrapperExitAction("Close", CANCEL_EXIT_CODE)
  private val allHistoryEntries = historyEntries

  init {
    title = "Command History"
    setResizable(true)
    commandField.focusTraversalKeysEnabled = false
    commandField.registerKeyboardAction(
      { applySelectedItem() },
      KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0),
      JComponent.WHEN_FOCUSED,
    )
    itemList.selectionMode = ListSelectionModel.SINGLE_SELECTION
    itemList.visibleRowCount = 16
    itemList.cellRenderer = RunIocCommandListRenderer()
    itemList.addMouseListener(
      object : MouseAdapter() {
        override fun mouseClicked(event: MouseEvent) {
          if (event.clickCount >= 1) {
            applySelectedItem()
          }
        }
      },
    )
    installDialogDocumentListener(commandField) {
      rebuildItems()
    }
    commandField.addActionListener { sendCommand() }
    init()
    rebuildItems()
  }

  override fun createCenterPanel(): JComponent {
    val titleLabel = JBLabel("Filter command history, then press Enter to send.").apply {
      font = font.deriveFont(Font.BOLD, font.size2D + 2f)
    }
    val metaLabel = JBLabel(runningIoc.label).apply {
      foreground = SimpleTextAttributes.GRAYED_ATTRIBUTES.fgColor
    }
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.empty(0, 0, 10, 0)
      add(titleLabel)
      add(Box.createVerticalStrut(4))
      add(metaLabel)
      add(Box.createVerticalStrut(10))
      add(captureOutputCheck)
      add(Box.createVerticalStrut(8))
      add(commandField)
      add(Box.createVerticalStrut(8))
      add(statusLabel)
    }
    return JPanel(BorderLayout(0, 10)).apply {
      preferredSize = Dimension(920, 520)
      add(topPanel, BorderLayout.NORTH)
      add(JBScrollPane(itemList), BorderLayout.CENTER)
    }
  }

  override fun createActions(): Array<Action> {
    return arrayOf(closeAction)
  }

  private fun applySelectedItem() {
    val selectedItem = itemList.selectedValue ?: return
    commandField.text = selectedItem.insertText
    commandField.requestFocusInWindow()
    commandField.caretPosition = commandField.text.length
  }

  private fun rebuildItems() {
    val filterText = commandField.text.trim().lowercase()
    val filteredEntries = allHistoryEntries.filter { entry ->
      filterText.isBlank() || entry.lowercase().contains(filterText)
    }
    itemModel.removeAllElements()
    if (filteredEntries.isEmpty()) {
      itemModel.addElement(
        RunIocCommandListItem(
          label = if (allHistoryEntries.isEmpty()) {
            "No IOC command history is available."
          } else {
            "No command history entries match the current filter."
          },
          description = "",
          detail = "",
          insertText = commandField.text,
        ),
      )
    } else {
      filteredEntries.forEach { entry ->
        itemModel.addElement(
          RunIocCommandListItem(
            label = entry,
            description = "",
            detail = "",
            insertText = entry,
          ),
        )
      }
      itemList.selectedIndex = 0
    }
  }

  private fun sendCommand() {
    val rawCommandText = commandField.text.trim()
    if (rawCommandText.isEmpty()) {
      close(CANCEL_EXIT_CODE)
      return
    }
    statusLabel.text =
      if (captureOutputCheck.isSelected) {
        "Capturing output for ${truncateDumpDescription(rawCommandText, 120)}"
      } else {
        "Sending ${truncateDumpDescription(rawCommandText, 120)}"
      }
    ApplicationManager.getApplication().executeOnPooledThread {
      val result = runCatching {
        val normalizedCommandText = normalizeRunIocCommandForExecution(
          rawCommandText,
          runtimeService.getCommandNames(startupFile),
        )
        if (captureOutputCheck.isSelected) {
          val output = runtimeService.captureCommandOutput(
            startupFile,
            normalizedCommandText,
            historyCommandText = rawCommandText,
          )
          runtimeService.openCapturedOutput(
            startupFile,
            "ioc-command",
            normalizedCommandText,
            output,
          )
        } else {
          runtimeService.sendCommandText(
            startupFile,
            normalizedCommandText,
            historyCommandText = rawCommandText,
          )
        }
      }
      ApplicationManager.getApplication().invokeLater {
        result.exceptionOrNull()?.let { error ->
          statusLabel.text = error.message ?: "Failed to run IOC command."
          Messages.showErrorDialog(project, statusLabel.text, "Command History")
          return@invokeLater
        }
        close(OK_EXIT_CODE)
      }
    }
  }
}

private class RunIocCommandListRenderer : DefaultListCellRenderer() {
  override fun getListCellRendererComponent(
    list: JList<*>,
    value: Any?,
    index: Int,
    isSelected: Boolean,
    cellHasFocus: Boolean,
  ): Component {
    val component = super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus)
    val label = component as? javax.swing.JLabel ?: return component
    val item = value as? RunIocCommandListItem ?: return component
    label.border = JBUI.Borders.empty(6, 8)
    label.text =
      """
      <html>
        <div><b>${escapeHtml(item.label)}</b></div>
        ${if (item.description.isNotBlank()) "<div>${escapeHtml(item.description)}</div>" else ""}
        ${if (item.detail.isNotBlank()) "<div>${escapeHtml(item.detail)}</div>" else ""}
      </html>
      """.trimIndent()
    return label
  }
}

private fun parseRunIocCommandContext(commandText: String): RunIocCommandContext? {
  val trimmedText = commandText.trimStart()
  val match = Regex("""^([A-Za-z_][A-Za-z0-9_]*)(\s+)([\s\S]*)$""").matchEntire(trimmedText) ?: return null
  val commandName = match.groupValues[1]
  val normalizedCommandName = commandName.lowercase()
  val rawArgumentText = match.groupValues[3]
  if (normalizedCommandName == "cd") {
    return RunIocCommandContext(commandName, "path", rawArgumentText)
  }
  val trimmedArgumentText = rawArgumentText.trim()
  if (trimmedArgumentText.isNotEmpty() && (rawArgumentText.lastOrNull()?.isWhitespace() == true || trimmedArgumentText.any(Char::isWhitespace))) {
    return null
  }
  return when {
    normalizedCommandName in RUN_IOC_RECORD_ARGUMENT_COMMANDS ->
      RunIocCommandContext(commandName, "record", trimmedArgumentText)
    normalizedCommandName in RUN_IOC_ENVIRONMENT_ARGUMENT_COMMANDS ->
      RunIocCommandContext(commandName, "environment", trimmedArgumentText)
    else -> null
  }
}

private class ProbeRecordDialog(
  private val project: com.intellij.openapi.project.Project,
  private val runningIoc: RunningIocSelectionItem,
  private val allRecordItems: List<DumpRecordSelectionItem>,
) : DialogWrapper(project, true) {
  private val recordField = JBTextField().apply {
    putClientProperty(
      "JTextField.placeholderText",
      "Filter records, press Tab to fill, Enter to open Probe.",
    )
  }
  private val statusLabel = JBLabel("Press Tab to fill the selected record, then Enter to open Probe.")
  private val itemModel = DefaultListModel<DumpRecordSelectionItem>()
  private val itemList = JBList(itemModel)
  private val closeAction = DialogWrapperExitAction("Close", CANCEL_EXIT_CODE)

  init {
    title = "Probe a Record"
    setResizable(true)
    recordField.focusTraversalKeysEnabled = false
    recordField.registerKeyboardAction(
      { applySelectedItem() },
      KeyStroke.getKeyStroke(KeyEvent.VK_TAB, 0),
      JComponent.WHEN_FOCUSED,
    )
    itemList.selectionMode = ListSelectionModel.SINGLE_SELECTION
    itemList.visibleRowCount = 16
    itemList.cellRenderer = DumpRecordPopupRenderer()
    itemList.addMouseListener(
      object : MouseAdapter() {
        override fun mouseClicked(event: MouseEvent) {
          if (event.clickCount >= 1) {
            applySelectedItem()
          }
        }
      },
    )
    installDialogDocumentListener(recordField) {
      rebuildItems()
    }
    recordField.addActionListener { probeRecord() }
    init()
    rebuildItems()
  }

  override fun createCenterPanel(): JComponent {
    val titleLabel = JBLabel("Filter records, press Tab to fill, then Enter to open Probe.").apply {
      font = font.deriveFont(Font.BOLD, font.size2D + 2f)
    }
    val metaLabel = JBLabel(runningIoc.label).apply {
      foreground = SimpleTextAttributes.GRAYED_ATTRIBUTES.fgColor
    }
    val topPanel = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      isOpaque = false
      border = JBUI.Borders.empty(0, 0, 10, 0)
      add(titleLabel)
      add(Box.createVerticalStrut(4))
      add(metaLabel)
      add(Box.createVerticalStrut(10))
      add(recordField)
      add(Box.createVerticalStrut(8))
      add(statusLabel)
    }
    return JPanel(BorderLayout(0, 10)).apply {
      preferredSize = Dimension(920, 520)
      add(topPanel, BorderLayout.NORTH)
      add(JBScrollPane(itemList), BorderLayout.CENTER)
    }
  }

  override fun createActions(): Array<Action> {
    return arrayOf(closeAction)
  }

  private fun applySelectedItem() {
    val selectedItem = itemList.selectedValue ?: return
    recordField.text = selectedItem.recordEntry.recordName
    recordField.requestFocusInWindow()
    recordField.caretPosition = recordField.text.length
  }

  private fun rebuildItems() {
    val filterText = recordField.text.trim().lowercase()
    val filteredEntries = allRecordItems.filter { item ->
      filterText.isBlank() ||
        item.recordEntry.recordName.lowercase().contains(filterText) ||
        item.recordEntry.recordDesc.lowercase().contains(filterText)
    }
    itemModel.removeAllElements()
    filteredEntries.forEach(itemModel::addElement)
    if (itemModel.size > 0) {
      itemList.selectedIndex = 0
      statusLabel.text = "Press Tab to fill the selected record, then Enter to open Probe."
    } else {
      statusLabel.text = "No records match the current filter."
    }
  }

  private fun probeRecord() {
    val typedRecordName = recordField.text.trim()
    if (typedRecordName.isEmpty()) {
      close(CANCEL_EXIT_CODE)
      return
    }

    val selectedItem =
      allRecordItems.firstOrNull {
        it.recordEntry.recordName.equals(typedRecordName, ignoreCase = true)
      } ?: itemList.selectedValue
    if (selectedItem == null) {
      statusLabel.text = "No record matches ${truncateDumpDescription(typedRecordName, 120)}."
      return
    }

    openEpicsWidget(project, selectedItem.recordEntry.recordName)
    close(OK_EXIT_CODE)
  }
}

private fun containsRunIocShellControlSyntax(text: String): Boolean {
  var inDoubleQuote = false
  var escaped = false
  var nestedParentheses = 0
  for (character in text) {
    if (escaped) {
      escaped = false
      continue
    }
    if (inDoubleQuote && character == '\\') {
      escaped = true
      continue
    }
    if (character == '"') {
      inDoubleQuote = !inDoubleQuote
      continue
    }
    if (inDoubleQuote) {
      continue
    }
    when (character) {
      '(' -> nestedParentheses += 1
      ')' -> if (nestedParentheses > 0) nestedParentheses -= 1
      '>', '<', '|', '&', ';' -> if (nestedParentheses == 0) return true
    }
  }
  return false
}

private fun extractWrappedRunIocArgumentText(text: String): String? {
  val trimmedText = text.trim()
  if (!trimmedText.startsWith("(")) {
    return null
  }

  var inDoubleQuote = false
  var escaped = false
  var depth = 0
  trimmedText.forEachIndexed { index, character ->
    if (escaped) {
      escaped = false
      return@forEachIndexed
    }
    if (inDoubleQuote && character == '\\') {
      escaped = true
      return@forEachIndexed
    }
    if (character == '"') {
      inDoubleQuote = !inDoubleQuote
      return@forEachIndexed
    }
    if (inDoubleQuote) {
      return@forEachIndexed
    }
    when (character) {
      '(' -> depth += 1
      ')' -> {
        depth -= 1
        if (depth == 0) {
          return if (index == trimmedText.lastIndex) {
            trimmedText.substring(1, index)
          } else {
            null
          }
        }
        if (depth < 0) {
          return null
        }
      }
    }
  }
  return null
}

private fun splitRunIocArgumentTokens(argumentText: String): List<String>? {
  val args = mutableListOf<String>()
  val currentToken = StringBuilder()
  var inDoubleQuote = false
  var escaped = false
  var nestedParentheses = 0

  fun flushCurrentToken() {
    val trimmedToken = currentToken.toString().trim()
    if (trimmedToken.isNotEmpty()) {
      args += trimmedToken
    }
    currentToken.setLength(0)
  }

  for (character in argumentText) {
    if (escaped) {
      currentToken.append(character)
      escaped = false
      continue
    }
    if (inDoubleQuote && character == '\\') {
      currentToken.append(character)
      escaped = true
      continue
    }
    if (character == '"') {
      currentToken.append(character)
      inDoubleQuote = !inDoubleQuote
      continue
    }
    if (!inDoubleQuote) {
      when (character) {
        '(' -> {
          nestedParentheses += 1
          currentToken.append(character)
          continue
        }
        ')' -> {
          if (nestedParentheses <= 0) {
            return null
          }
          nestedParentheses -= 1
          currentToken.append(character)
          continue
        }
        ',', ' ', '\t', '\n', '\r' -> if (nestedParentheses == 0) {
          flushCurrentToken()
          continue
        }
      }
    }
    currentToken.append(character)
  }

  if (escaped || inDoubleQuote || nestedParentheses != 0) {
    return null
  }
  flushCurrentToken()
  return args
}

private fun normalizeRunIocCommandForExecution(
  commandText: String,
  knownCommandNames: Collection<String>,
): String {
  val trimmedText = commandText.trim()
  if (trimmedText.isEmpty()) {
    return ""
  }

  val match = Regex("""^([A-Za-z_][A-Za-z0-9_]*)([\s\S]*)$""").matchEntire(trimmedText) ?: return trimmedText
  val commandName = match.groupValues[1]
  val remainderText = match.groupValues[2]
  if (remainderText.isBlank()) {
    return trimmedText
  }
  if (containsRunIocShellControlSyntax(remainderText)) {
    return trimmedText
  }
  if (knownCommandNames.none { it.equals(commandName, ignoreCase = true) }) {
    return trimmedText
  }

  val trimmedRemainder = remainderText.trim()
  val wrappedArgumentText =
    if (trimmedRemainder.startsWith("(")) extractWrappedRunIocArgumentText(trimmedRemainder) else null
  if (trimmedRemainder.startsWith("(") && wrappedArgumentText == null) {
    return trimmedText
  }
  val argumentTokens = splitRunIocArgumentTokens(wrappedArgumentText ?: trimmedRemainder) ?: return trimmedText
  if (argumentTokens.isEmpty()) {
    return if (wrappedArgumentText != null) "$commandName()" else commandName
  }

  return buildString {
    append(commandName)
    append('(')
    append(argumentTokens.joinToString(", "))
    append(')')
  }
}

private fun parseDblRecordNames(output: String): List<String> {
  return output.lineSequence()
    .map(String::trim)
    .filter(String::isNotEmpty)
    .filterNot { it.startsWith("epics>") }
    .distinct()
    .toList()
}

private fun extractRunCommandSummary(helpText: String): String {
  return truncateDumpDescription(
    helpText.lineSequence().map(String::trim).firstOrNull(String::isNotEmpty)
      ?: "No detailed help is available.",
    160,
  )
}

private fun collectRunIocPathSuggestions(baseDirectory: Path, partialText: String): List<String> {
  val rawPartialText = partialText
  val lookupPrefix = if (rawPartialText.startsWith("~")) {
    System.getProperty("user.home").orEmpty() + rawPartialText.removePrefix("~")
  } else {
    rawPartialText
  }
  val rawDirectoryPrefix =
    if (rawPartialText.contains('/') || rawPartialText.contains('\\') || rawPartialText.endsWith('/')) {
      rawPartialText.replace(Regex("""[^/\\]*$"""), "")
    } else {
      ""
    }
  val lookupDirectoryPrefix =
    if (lookupPrefix.contains('/') || lookupPrefix.contains('\\') || lookupPrefix.endsWith('/')) {
      lookupPrefix.replace(Regex("""[^/\\]*$"""), "")
    } else {
      ""
    }
  val fileNamePrefix =
    if (rawPartialText.endsWith('/') || rawPartialText.endsWith('\\')) {
      ""
    } else {
      rawPartialText.substringAfterLast('/').substringAfterLast('\\')
    }
  val lookupDirectory = runCatching {
    baseDirectory.resolve(lookupDirectoryPrefix.ifBlank { "." }).normalize()
  }.getOrElse {
    baseDirectory
  }
  val directoryEntries = runCatching {
    Files.list(lookupDirectory).use { stream ->
      stream
        .filter(Files::isDirectory)
        .map { path -> path.fileName.toString() }
        .sorted(String.CASE_INSENSITIVE_ORDER)
        .toList()
    }
  }.getOrDefault(emptyList())

  return directoryEntries.filter { entryName ->
    entryName.contains(fileNamePrefix, ignoreCase = true)
  }.map { entryName ->
    rawDirectoryPrefix + entryName
  }
}

private fun formatRunIocPathSuggestion(pathText: String): String {
  return if (pathText.any(Char::isWhitespace)) {
    "\"${pathText.replace("\"", "\\\"")}\""
  } else {
    pathText
  }
}

private fun installDialogDocumentListener(field: JBTextField, callback: () -> Unit) {
  field.document.addDocumentListener(object : javax.swing.event.DocumentListener {
    override fun insertUpdate(event: javax.swing.event.DocumentEvent?) = callback()
    override fun removeUpdate(event: javax.swing.event.DocumentEvent?) = callback()
    override fun changedUpdate(event: javax.swing.event.DocumentEvent?) = callback()
  })
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

private fun buildDumpAllRecordNamesFileName(item: RunningIocSelectionItem): String {
  return "epics-dbl-${sanitizeDumpFileToken(item.startup.startupName)}.pvlist"
}

private fun buildDumpAllRecordsExcelFileName(item: RunningIocSelectionItem): String {
  return "epics-dbdumprecord-${sanitizeDumpFileToken(item.startup.startupName)}.xlsx"
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
    val descText = truncateDumpDescription(item.recordEntry.recordDesc)
    label.text =
      """
      <html>
        <div><b>${escapeHtml(item.recordEntry.recordName)}</b></div>
        ${if (descText.isNotBlank()) "<div>${escapeHtml(descText)}</div>" else ""}
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
