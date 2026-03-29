package org.epics.workbench.runtime

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnAction
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.components.service
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.ui.DialogWrapper
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.openapi.vfs.toNioPathOrNull
import com.intellij.openapi.actionSystem.DefaultActionGroup
import com.intellij.ui.ColoredListCellRenderer
import com.intellij.ui.JBColor
import com.intellij.ui.SimpleTextAttributes
import com.intellij.ui.components.JBLabel
import com.intellij.ui.components.JBList
import com.intellij.ui.components.JBScrollPane
import org.epics.workbench.build.EpicsBuildModelService
import org.epics.workbench.build.collectEpicsProjectRoots
import org.epics.workbench.build.findContainingBuildProjectRoot
import org.epics.workbench.build.projectHasEpicsRoot
import java.awt.BorderLayout
import java.awt.Dimension
import java.awt.Font
import java.awt.event.ActionEvent
import java.awt.event.MouseAdapter
import java.awt.event.MouseEvent
import java.nio.file.Files
import java.nio.file.Path
import javax.swing.Action
import javax.swing.JComponent
import javax.swing.JList
import javax.swing.ListSelectionModel
import kotlin.io.path.extension
import kotlin.io.path.isDirectory
import kotlin.io.path.isRegularFile
import kotlin.io.path.pathString

class ManageProjectIocAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    event.presentation.isEnabledAndVisible = project != null && projectHasEpicsRoot(project)
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    showProjectIocChooser(project, file)
  }
}

class EpicsProjectIocStartPopupGroup : DefaultActionGroup(), DumbAware {
  init {
    templatePresentation.setHideGroupIfEmpty(true)
    templatePresentation.setDisableGroupIfEmpty(false)
  }

  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    if (project == null || !projectHasEpicsRoot(project)) {
      event.presentation.isEnabledAndVisible = false
      return
    }
    val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    event.presentation.isEnabledAndVisible = collectProjectStartupMenuItems(project, file).isNotEmpty()
  }

  override fun getChildren(event: AnActionEvent?): Array<AnAction> {
    val project = event?.project ?: return emptyArray()
    val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    return collectProjectStartupMenuItems(project, file)
      .map { startupItem -> ProjectStartupIocStartAction(startupItem) }
      .toTypedArray()
  }
}

class EpicsProjectIocStopPopupGroup : DefaultActionGroup(), DumbAware {
  init {
    templatePresentation.setHideGroupIfEmpty(true)
    templatePresentation.setDisableGroupIfEmpty(false)
  }

  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    if (project == null || !projectHasEpicsRoot(project)) {
      event.presentation.isEnabledAndVisible = false
      return
    }
    val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    event.presentation.isEnabledAndVisible = collectProjectStartupMenuItems(project, file).isNotEmpty()
  }

  override fun getChildren(event: AnActionEvent?): Array<AnAction> {
    val project = event?.project ?: return emptyArray()
    val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
      ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    return collectProjectStartupMenuItems(project, file)
      .map { startupItem -> ProjectStartupIocStopAction(startupItem) }
      .toTypedArray()
  }
}

internal fun showProjectIocChooser(
  project: com.intellij.openapi.project.Project,
  contextFile: VirtualFile? = null,
) {
  val startupFiles = collectProjectStartupFiles(project, contextFile)
  if (startupFiles.isEmpty()) {
    Messages.showWarningDialog(project, "No st.cmd-like files were found under this project's iocBoot folders.", TITLE)
    return
  }

  ProjectIocChooserDialog(project, startupFiles, contextFile).show()
}

private fun collectProjectStartupFiles(
  project: com.intellij.openapi.project.Project,
  contextFile: VirtualFile?,
): List<VirtualFile> {
  return collectProjectStartupMenuItems(project, contextFile).map { it.startupFile }
}

private data class ProjectStartupMenuItem(
  val startupFile: VirtualFile,
  val label: String,
)

private fun collectProjectStartupMenuItems(
  project: com.intellij.openapi.project.Project,
  contextFile: VirtualFile?,
): List<ProjectStartupMenuItem> {
  val roots = findContainingBuildProjectRoot(contextFile)?.let(::listOf)
    ?: collectEpicsProjectRoots(project)
  if (roots.isEmpty()) {
    return emptyList()
  }

  data class StartupCandidate(
    val rootName: String,
    val startupFile: VirtualFile,
    val iocBootRelativePath: String,
  )

  val buildModelService = project.service<EpicsBuildModelService>()
  val results = linkedMapOf<String, StartupCandidate>()
  roots.forEach { root ->
    val rootPath = root.toNioPathOrNull()?.normalize() ?: return@forEach
    val iocBootRootPath = rootPath.resolve("iocBoot").normalize()
    val startupPaths = buildModelService.collectStartupEntryPoints(rootPath).ifEmpty {
      collectStartupFilesUnder(iocBootRootPath)
    }
    startupPaths.forEach startupPathLoop@{ startupPath ->
      val normalizedPath = startupPath.normalize()
      val startupFile = LocalFileSystem.getInstance().refreshAndFindFileByPath(normalizedPath.pathString)
        ?: return@startupPathLoop
      val relativePath = runCatching {
        iocBootRootPath.relativize(normalizedPath).toString().replace('\\', '/')
      }.getOrDefault(startupFile.name)
      results.putIfAbsent(
        startupFile.path,
        StartupCandidate(
          rootName = root.name,
          startupFile = startupFile,
          iocBootRelativePath = relativePath,
        ),
      )
    }
  }

  val duplicateLabels = results.values
    .groupingBy { it.iocBootRelativePath.lowercase() }
    .eachCount()

  return results.values
    .sortedBy { it.startupFile.path.lowercase() }
    .map { candidate ->
      val label = if ((duplicateLabels[candidate.iocBootRelativePath.lowercase()] ?: 0) > 1) {
        "${candidate.rootName}/${candidate.iocBootRelativePath}"
      } else {
        candidate.iocBootRelativePath
      }
      ProjectStartupMenuItem(
        startupFile = candidate.startupFile,
        label = label,
      )
    }
}

private fun collectStartupFilesUnder(iocBootDirectory: Path): List<Path> {
  if (!iocBootDirectory.isDirectory()) {
    return emptyList()
  }
  return try {
    Files.walk(iocBootDirectory).use { stream ->
      stream
        .filter { candidate -> candidate.isRegularFile() && isStartupFilePath(candidate) }
        .sorted(compareBy<Path> { it.pathString.lowercase() })
        .toList()
    }
  } catch (_: Exception) {
    emptyList()
  }
}

private fun isStartupFilePath(path: Path): Boolean {
  val fileName = path.fileName?.toString().orEmpty()
  val extension = path.extension.lowercase()
  return extension == "cmd" || extension == "iocsh" || fileName == "st.cmd"
}

private class ProjectStartupIocStartAction(
  private val startupItem: ProjectStartupMenuItem,
) : DumbAwareAction(startupItem.label) {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val runtimeService = project?.service<EpicsIocRuntimeService>()
    val running = runtimeService?.isRunning(startupItem.startupFile) == true
    event.presentation.text = startupItem.label
    event.presentation.description = "Start the IOC for ${startupItem.startupFile.name}"
    event.presentation.isVisible =
      project != null && startupItem.startupFile.isValid && EpicsIocRuntimeService.isIocBootStartupFile(startupItem.startupFile)
    event.presentation.isEnabled = !running
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val runtimeService = project.service<EpicsIocRuntimeService>()
    val startupFile = startupItem.startupFile
    if (!startupFile.isValid || runtimeService.isRunning(startupFile)) {
      return
    }

    FileDocumentManager.getInstance().getDocument(startupFile)?.let(FileDocumentManager.getInstance()::saveDocument)
    FileEditorManager.getInstance(project).openFile(startupFile, true, true)
    runtimeService.startIoc(startupFile).exceptionOrNull()?.let { error ->
      Messages.showErrorDialog(project, error.message ?: "Failed to start ${startupFile.name}.", TITLE)
    }
  }
}

private class ProjectStartupIocStopAction(
  private val startupItem: ProjectStartupMenuItem,
) : DumbAwareAction(startupItem.label) {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val project = event.project
    val runtimeService = project?.service<EpicsIocRuntimeService>()
    val running = runtimeService?.isRunning(startupItem.startupFile) == true
    event.presentation.text = startupItem.label
    event.presentation.description = "Stop the IOC for ${startupItem.startupFile.name}"
    event.presentation.isVisible =
      project != null && startupItem.startupFile.isValid && EpicsIocRuntimeService.isIocBootStartupFile(startupItem.startupFile)
    event.presentation.isEnabled = running
  }

  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    val runtimeService = project.service<EpicsIocRuntimeService>()
    val startupFile = startupItem.startupFile
    if (!startupFile.isValid || !runtimeService.isRunning(startupFile)) {
      return
    }
    runtimeService.stopIoc(startupFile)
  }
}

private class ProjectIocChooserDialog(
  private val project: com.intellij.openapi.project.Project,
  startupFiles: List<VirtualFile>,
  initialSelection: VirtualFile?,
) : DialogWrapper(project, true) {
  private val runtimeService = project.service<EpicsIocRuntimeService>()
  private val startupList = JBList(startupFiles)
  private val openFileAction = object : DialogWrapperAction("Open File") {
    override fun doAction(event: ActionEvent?) {
      openSelectedStartupFile(closeAfterOpen = false)
    }
  }
  private val startOrBringAction = object : DialogWrapperAction("Start") {
    override fun doAction(event: ActionEvent?) {
      val startupFile = selectedStartupFile() ?: return
      if (runtimeService.isRunning(startupFile)) {
        runtimeService.showRunningConsole(startupFile)
      } else {
        saveStartupDocument(startupFile)
        runtimeService.startIoc(startupFile).exceptionOrNull()?.let { error ->
          Messages.showErrorDialog(project, error.message ?: "Failed to start ${startupFile.name}.", TITLE)
          return
        }
      }
      refreshActionState()
      startupList.repaint()
    }
  }
  private val stopAction = object : DialogWrapperAction("Stop") {
    override fun doAction(event: ActionEvent?) {
      val startupFile = selectedStartupFile() ?: return
      runtimeService.stopIoc(startupFile)
      refreshActionState()
      startupList.repaint()
    }
  }
  private val closeAction = DialogWrapperExitAction("Close", CANCEL_EXIT_CODE)

  init {
    title = TITLE
    setResizable(true)
    startupList.selectionMode = ListSelectionModel.SINGLE_SELECTION
    startupList.visibleRowCount = 14
    startupList.cellRenderer = StartupListRenderer(runtimeService)
    startupList.addListSelectionListener { refreshActionState() }
    startupList.addMouseListener(
      object : MouseAdapter() {
        override fun mouseClicked(event: MouseEvent) {
          if (event.clickCount >= 2 && startupList.selectedValue != null) {
            openSelectedStartupFile(closeAfterOpen = true)
          }
        }
      },
    )
    init()

    val initialIndex = startupFiles.indexOfFirst { candidate ->
      initialSelection?.path == candidate.path
    }.takeIf { it >= 0 } ?: 0
    startupList.selectedIndex = initialIndex
    refreshActionState()
  }

  override fun createCenterPanel(): JComponent {
    val titleLabel = JBLabel("Select an IOC startup file.").apply {
      font = font.deriveFont(Font.BOLD, font.size2D + 2f)
    }
    return javax.swing.JPanel(BorderLayout(0, 10)).apply {
      preferredSize = Dimension(860, 420)
      add(titleLabel, BorderLayout.NORTH)
      add(JBScrollPane(startupList), BorderLayout.CENTER)
    }
  }

  override fun createActions(): Array<Action> {
    return arrayOf(openFileAction, startOrBringAction, stopAction, closeAction)
  }

  private fun refreshActionState() {
    val startupFile = selectedStartupFile()
    val running = startupFile?.let(runtimeService::isRunning) == true
    openFileAction.isEnabled = startupFile != null
    startOrBringAction.isEnabled = startupFile != null
    stopAction.isEnabled = startupFile != null
    startOrBringAction.putValue(Action.NAME, if (running) "Bring to Front" else "Start")
  }

  private fun selectedStartupFile(): VirtualFile? = startupList.selectedValue

  private fun openSelectedStartupFile(closeAfterOpen: Boolean) {
    val startupFile = selectedStartupFile() ?: return
    FileEditorManager.getInstance(project).openFile(startupFile, true, true)
    if (closeAfterOpen) {
      close(CANCEL_EXIT_CODE)
    }
  }

  private fun saveStartupDocument(startupFile: VirtualFile) {
    FileDocumentManager.getInstance().getDocument(startupFile)?.let(FileDocumentManager.getInstance()::saveDocument)
  }
}

private class StartupListRenderer(
  private val runtimeService: EpicsIocRuntimeService,
) : ColoredListCellRenderer<VirtualFile>() {
  override fun customizeCellRenderer(
    list: JList<out VirtualFile>,
    value: VirtualFile?,
    index: Int,
    selected: Boolean,
    hasFocus: Boolean,
  ) {
    val startupFile = value ?: return
    font = list.font.deriveFont(list.font.size2D + 2f)
    append(startupFile.path, SimpleTextAttributes.REGULAR_ATTRIBUTES)
    if (runtimeService.isRunning(startupFile)) {
      append(
        "   Running",
        SimpleTextAttributes(SimpleTextAttributes.STYLE_PLAIN, JBColor(0x2E7D32, 0x73D07C)),
      )
    }
  }
}

private const val TITLE = "EPICS IOC Start/Stop"
