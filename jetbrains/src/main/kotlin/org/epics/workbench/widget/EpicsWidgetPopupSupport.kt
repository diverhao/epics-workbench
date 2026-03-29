package org.epics.workbench.widget

import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.projectHasEpicsRoot
import org.epics.workbench.build.runMakeBuildCommand
import org.epics.workbench.build.selectEpicsProjectRoot
import org.epics.workbench.export.canExportDatabaseFile
import org.epics.workbench.export.exportDatabaseFileToExcel
import org.epics.workbench.export.promptImportEpicsExcelWorkbook
import org.epics.workbench.pvlist.EpicsPvlistWidgetModel
import org.epics.workbench.pvlist.EpicsPvlistWidgetSourceKind
import java.awt.Component
import java.nio.file.Path
import java.awt.event.ContainerAdapter
import java.awt.event.ContainerEvent
import javax.swing.JComponent
import javax.swing.JMenu
import javax.swing.JMenuItem
import javax.swing.JPopupMenu
import javax.swing.event.PopupMenuEvent
import javax.swing.event.PopupMenuListener

internal fun installEpicsWidgetPopupMenu(
  project: Project,
  component: JComponent,
  channelsProvider: () -> List<String>,
  primaryChannelProvider: () -> String?,
  sourceLabelProvider: () -> String,
  exportFileProvider: () -> VirtualFile? = { null },
) {
  val popupMenu = JPopupMenu()
  popupMenu.addPopupMenuListener(
    object : PopupMenuListener {
      override fun popupMenuWillBecomeVisible(event: PopupMenuEvent?) {
        rebuildWidgetPopupMenu(
          popupMenu = popupMenu,
          project = project,
          channelsProvider = channelsProvider,
          primaryChannelProvider = primaryChannelProvider,
          sourceLabelProvider = sourceLabelProvider,
          exportFileProvider = exportFileProvider,
        )
      }

      override fun popupMenuWillBecomeInvisible(event: PopupMenuEvent?) = Unit

      override fun popupMenuCanceled(event: PopupMenuEvent?) = Unit
    },
  )
  applyPopupMenu(component, popupMenu)
}

private fun rebuildWidgetPopupMenu(
  popupMenu: JPopupMenu,
  project: Project,
  channelsProvider: () -> List<String>,
  primaryChannelProvider: () -> String?,
  sourceLabelProvider: () -> String,
  exportFileProvider: () -> VirtualFile?,
) {
  popupMenu.removeAll()

  popupMenu.add(buildBuildMenu(project))
  popupMenu.add(buildWidgetsMenu(project, channelsProvider, primaryChannelProvider, sourceLabelProvider))
  popupMenu.add(buildImportExportMenu(project, exportFileProvider))
}

private fun buildBuildMenu(project: Project): JMenu {
  val menu = JMenu("EPICS Build")
  val canBuild = projectHasEpicsRoot(project)
  menu.add(
    JMenuItem("Build Project").apply {
      isEnabled = canBuild
      addActionListener {
        val root = selectEpicsProjectRoot(project, null, "Build Project", "Select an EPICS project root to build.")
          ?: return@addActionListener
        runMakeBuildCommand(project, Path.of(root.path), "Build Project ${root.name}", listOf(emptyList()))
      }
    },
  )
  menu.add(
    JMenuItem("Clean Project").apply {
      isEnabled = canBuild
      addActionListener {
        val root = selectEpicsProjectRoot(project, null, "Clean Project", "Select an EPICS project root to clean.")
          ?: return@addActionListener
        runMakeBuildCommand(project, Path.of(root.path), "Clean Project ${root.name}", listOf(listOf("distclean")))
      }
    },
  )
  return menu
}

private fun buildWidgetsMenu(
  project: Project,
  channelsProvider: () -> List<String>,
  primaryChannelProvider: () -> String?,
  sourceLabelProvider: () -> String,
): JMenu {
  val channels = normalizeChannels(channelsProvider())
  val primaryChannel = primaryChannelProvider()?.trim()?.takeIf(String::isNotBlank)
  val menu = JMenu("EPICS Widgets")

  menu.add(
    JMenuItem("Probe").apply {
      isEnabled = primaryChannel != null
      addActionListener {
        openEpicsWidget(project, primaryChannel.orEmpty())
      }
    },
  )
  menu.add(
    JMenuItem("PV List").apply {
      isEnabled = channels.isNotEmpty()
      addActionListener {
        openEpicsPvlistWidget(
          project,
          EpicsPvlistWidgetModel(
            sourceLabel = sourceLabelProvider(),
            sourcePath = null,
            sourceKind = EpicsPvlistWidgetSourceKind.PVLIST,
            rawPvNames = channels.toMutableList(),
            macroNames = mutableListOf(),
            macroValues = linkedMapOf(),
          ),
        )
      }
    },
  )
  menu.add(
    JMenuItem("PV Monitor").apply {
      isEnabled = channels.isNotEmpty()
      addActionListener {
        openEpicsMonitorWidget(project, channels)
      }
    },
  )
  return menu
}

private fun buildImportExportMenu(
  project: Project,
  exportFileProvider: () -> VirtualFile?,
): JMenu {
  val menu = JMenu("EPICS Import/Export")
  menu.add(
    JMenuItem("Import Excel as Database").apply {
      addActionListener {
        promptImportEpicsExcelWorkbook(project)
      }
    },
  )

  val exportFile = exportFileProvider()?.takeIf(::canExportDatabaseFile)
  if (exportFile != null) {
    menu.add(
      JMenuItem("Export Database to Excel").apply {
        addActionListener {
          exportDatabaseFileToExcel(project, exportFile)
        }
      },
    )
  }
  return menu
}

private fun normalizeChannels(channels: List<String>): List<String> {
  val results = linkedSetOf<String>()
  channels.forEach { rawValue ->
    val trimmed = rawValue.trim()
    if (trimmed.isBlank()) {
      return@forEach
    }
    results += stripMonitorProtocol(trimmed)
  }
  return results.toList()
}

private fun stripMonitorProtocol(channelName: String): String {
  val separatorIndex = channelName.indexOf("://")
  return if (separatorIndex > 0) {
    channelName.substring(separatorIndex + 3)
  } else {
    channelName
  }
}

private fun applyPopupMenu(component: Component, popupMenu: JPopupMenu) {
  if (component is JComponent) {
    component.componentPopupMenu = popupMenu
    component.inheritsPopupMenu = true
    component.addContainerListener(
      object : ContainerAdapter() {
        override fun componentAdded(event: ContainerEvent) {
          event.child?.let { child -> applyPopupMenu(child, popupMenu) }
        }
      },
    )
    component.components.forEach { child -> applyPopupMenu(child, popupMenu) }
  }
}
