package org.epics.workbench.widget

import com.intellij.openapi.Disposable
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.components.service
import com.intellij.openapi.fileEditor.FileEditor
import com.intellij.openapi.fileEditor.FileEditorLocation
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorPolicy
import com.intellij.openapi.fileEditor.FileEditorProvider
import com.intellij.openapi.fileEditor.FileEditorState
import com.intellij.openapi.fileEditor.FileEditorStateLevel
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import org.epics.workbench.runtime.EpicsMonitorRuntimeService
import org.epics.workbench.runtime.EpicsProbeViewPanel
import java.awt.BorderLayout
import java.awt.Dimension
import java.beans.PropertyChangeListener
import java.util.UUID
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JComponent
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.JTextField

internal class EpicsWidgetVirtualFile(
  initialRecordName: String = "",
) : LightVirtualFile(TAB_TITLE) {
  val widgetId: String = UUID.randomUUID().toString()
  var initialRecordName: String = initialRecordName

  companion object {
    const val TAB_TITLE: String = "EPICS Probe"
  }
}

class OpenEpicsWidgetAction : DumbAwareAction() {
  override fun actionPerformed(event: AnActionEvent) {
    val project = event.project ?: return
    openEpicsWidget(project)
  }

  override fun update(event: AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null
  }
}

internal fun openEpicsWidget(project: Project, initialRecordName: String = "") {
  FileEditorManager.getInstance(project).openFile(EpicsWidgetVirtualFile(initialRecordName), true, true)
}

class EpicsWidgetFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: VirtualFile): Boolean {
    return file is EpicsWidgetVirtualFile
  }

  override fun createEditor(project: Project, file: VirtualFile): FileEditor {
    return EpicsWidgetFileEditor(project, file as EpicsWidgetVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-widget-editor"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsWidgetFileEditor(
  private val project: Project,
  private val file: EpicsWidgetVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val runtimeService = project.service<EpicsMonitorRuntimeService>()
  private val channelField = JTextField(file.initialRecordName, 28)
  private var currentRecordName: String = file.initialRecordName.trim()

  private val probePanel = EpicsProbeViewPanel(
    stateProvider = { runtimeService.getWidgetProbeViewState(file.widgetId, currentRecordName) },
    putHandler = { key -> runtimeService.requestPutWidgetValue(file.widgetId, currentRecordName, key) },
    isMonitoringActive = { runtimeService.isMonitoringActive() },
    startHandler = { runtimeService.startMonitoring() },
    stopHandler = { runtimeService.stopMonitoring() },
    processHandler = { runtimeService.requestProcessWidget(file.widgetId, currentRecordName) },
    showStartStopControls = false,
  )

  private val component = JPanel(BorderLayout())

  init {
    runtimeService.initialize()

    val controlRow = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.X_AXIS)
      add(JLabel("Channel name:"))
      add(Box.createHorizontalStrut(8))
      channelField.maximumSize = Dimension(360, channelField.preferredSize.height)
      add(channelField)
      add(Box.createHorizontalGlue())
    }

    val container = JPanel(BorderLayout(0, 12)).apply {
      add(controlRow, BorderLayout.NORTH)
      add(probePanel, BorderLayout.CENTER)
    }

    channelField.addActionListener { applyChannelName() }

    component.add(container, BorderLayout.CENTER)
    if (currentRecordName.isNotBlank()) {
      applyChannelName()
    }
    probePanel.refreshFromService()
  }

  override fun getComponent(): JComponent = component

  override fun getPreferredFocusedComponent(): JComponent? = channelField

  override fun getFile(): VirtualFile = file

  override fun getName(): String = EpicsWidgetVirtualFile.TAB_TITLE

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() {
    runtimeService.releaseWidgetProbeSession(file.widgetId)
    probePanel.dispose()
  }

  private fun applyChannelName() {
    currentRecordName = channelField.text.trim()
    file.initialRecordName = currentRecordName
    if (currentRecordName.isNotBlank()) {
      runtimeService.startMonitoring()
    } else {
      runtimeService.releaseWidgetProbeSession(file.widgetId)
    }
    probePanel.refreshFromService()
  }
}
