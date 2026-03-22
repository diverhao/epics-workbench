package org.epics.workbench.export

import com.intellij.openapi.Disposable
import com.intellij.openapi.fileEditor.FileEditor
import com.intellij.openapi.fileEditor.FileEditorLocation
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileEditor.FileEditorPolicy
import com.intellij.openapi.fileEditor.FileEditorProvider
import com.intellij.openapi.fileEditor.FileEditorState
import com.intellij.openapi.fileEditor.FileEditorStateLevel
import com.intellij.openapi.project.DumbAware
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.util.UserDataHolderBase
import com.intellij.testFramework.LightVirtualFile
import java.awt.BorderLayout
import java.awt.Color
import java.awt.Cursor
import java.awt.datatransfer.DataFlavor
import java.awt.datatransfer.Transferable
import java.beans.PropertyChangeListener
import java.nio.file.Path
import javax.swing.BorderFactory
import javax.swing.Box
import javax.swing.BoxLayout
import javax.swing.JButton
import javax.swing.JComponent
import javax.swing.JEditorPane
import javax.swing.JLabel
import javax.swing.JPanel
import javax.swing.TransferHandler

internal class EpicsExcelImportPreviewVirtualFile : LightVirtualFile(TAB_TITLE) {
  companion object {
    const val TAB_TITLE: String = "EPICS Excel Import Preview"
  }
}

class OpenExcelImportPreviewAction : com.intellij.openapi.project.DumbAwareAction() {
  override fun actionPerformed(event: com.intellij.openapi.actionSystem.AnActionEvent) {
    val project = event.project ?: return
    openEpicsExcelImportPreview(project)
  }

  override fun update(event: com.intellij.openapi.actionSystem.AnActionEvent) {
    event.presentation.isEnabledAndVisible = event.project != null
  }
}

internal fun openEpicsExcelImportPreview(project: Project) {
  FileEditorManager.getInstance(project).openFile(EpicsExcelImportPreviewVirtualFile(), true, true)
}

class EpicsExcelImportPreviewFileEditorProvider : FileEditorProvider, DumbAware {
  override fun accept(project: Project, file: com.intellij.openapi.vfs.VirtualFile): Boolean {
    return file is EpicsExcelImportPreviewVirtualFile
  }

  override fun createEditor(project: Project, file: com.intellij.openapi.vfs.VirtualFile): FileEditor {
    return EpicsExcelImportPreviewFileEditor(project, file as EpicsExcelImportPreviewVirtualFile)
  }

  override fun getEditorTypeId(): String = "epics-excel-import-preview"

  override fun getPolicy(): FileEditorPolicy = FileEditorPolicy.HIDE_DEFAULT_EDITOR
}

private class EpicsExcelImportPreviewFileEditor(
  private val project: Project,
  private val file: EpicsExcelImportPreviewVirtualFile,
) : UserDataHolderBase(), FileEditor, Disposable {
  private val statusLabel = JLabel("Drop an .xlsx workbook here.")
  private val component = JPanel(BorderLayout())

  init {
    val container = JPanel().apply {
      layout = BoxLayout(this, BoxLayout.Y_AXIS)
      border = BorderFactory.createEmptyBorder(24, 24, 24, 24)
    }

    val titleLabel = JLabel("EPICS Excel Import Preview").apply {
      alignmentX = JComponent.LEFT_ALIGNMENT
    }
    val instructions = JEditorPane("text/html", """
      <html>
        <body style="font-family: sans-serif; font-size: 12px;">
          <p>Drop an <b>.xlsx</b> workbook here. If one or more sheets start with
          <code>Record</code> and <code>Type</code> in the first row, each EPICS-style
          sheet will open as a temporary EPICS database tab.</p>
        </body>
      </html>
    """.trimIndent()).apply {
      isEditable = false
      isOpaque = false
      border = BorderFactory.createEmptyBorder()
      cursor = Cursor.getDefaultCursor()
      alignmentX = JComponent.LEFT_ALIGNMENT
    }
    val dropPanel = JPanel(BorderLayout()).apply {
      alignmentX = JComponent.LEFT_ALIGNMENT
      border = BorderFactory.createCompoundBorder(
        BorderFactory.createDashedBorder(Color.GRAY, 4f, 6f),
        BorderFactory.createEmptyBorder(24, 24, 24, 24),
      )
      add(JLabel("Drop Excel workbook here"), BorderLayout.CENTER)
      transferHandler = ExcelWorkbookDropHandler(project) { message ->
        statusLabel.text = message
      }
    }
    val buttonRow = JPanel(BorderLayout()).apply {
      alignmentX = JComponent.LEFT_ALIGNMENT
      add(
        JButton("Choose Workbook...").apply {
          addActionListener {
            val descriptor = com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
              .createSingleFileDescriptor("xlsx")
              .withTitle("Import Excel as EPICS Database")
            com.intellij.openapi.fileChooser.FileChooser.chooseFile(descriptor, project, null) { file ->
              importWorkbookPath(project, Path.of(file.path), statusLabel::setText)
            }
          }
        },
        BorderLayout.WEST,
      )
    }

    container.add(titleLabel)
    container.add(Box.createVerticalStrut(12))
    container.add(instructions)
    container.add(Box.createVerticalStrut(16))
    container.add(dropPanel)
    container.add(Box.createVerticalStrut(12))
    container.add(buttonRow)
    container.add(Box.createVerticalStrut(12))
    container.add(statusLabel)

    component.add(container, BorderLayout.CENTER)
  }

  override fun getComponent(): JComponent = component

  override fun getPreferredFocusedComponent(): JComponent? = component

  override fun getFile(): com.intellij.openapi.vfs.VirtualFile = file

  override fun getName(): String = EpicsExcelImportPreviewVirtualFile.TAB_TITLE

  override fun setState(state: FileEditorState) = Unit

  override fun isModified(): Boolean = false

  override fun isValid(): Boolean = true

  override fun addPropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun removePropertyChangeListener(listener: PropertyChangeListener) = Unit

  override fun getCurrentLocation(): FileEditorLocation? = null

  override fun getState(level: FileEditorStateLevel): FileEditorState = FileEditorState.INSTANCE

  override fun dispose() = Unit
}

private class ExcelWorkbookDropHandler(
  private val project: Project,
  private val updateStatus: (String) -> Unit,
) : TransferHandler() {
  override fun canImport(support: TransferSupport): Boolean {
    if (!support.isDataFlavorSupported(DataFlavor.javaFileListFlavor)) {
      return false
    }
    val files = extractDroppedFiles(support.transferable)
    return files.any { it.fileName.toString().endsWith(".xlsx", ignoreCase = true) }
  }

  override fun importData(support: TransferSupport): Boolean {
    if (!canImport(support)) {
      return false
    }
    val files = extractDroppedFiles(support.transferable)
      .filter { it.fileName.toString().endsWith(".xlsx", ignoreCase = true) }
    if (files.isEmpty()) {
      return false
    }

    files.forEach { path -> importWorkbookPath(project, path, updateStatus) }
    return true
  }

  @Suppress("UNCHECKED_CAST")
  private fun extractDroppedFiles(transferable: Transferable): List<Path> {
    val files = transferable.getTransferData(DataFlavor.javaFileListFlavor) as? List<*> ?: return emptyList()
    return files.mapNotNull { item ->
      when (item) {
        is java.io.File -> item.toPath()
        else -> null
      }
    }
  }
}

private fun importWorkbookPath(project: Project, path: Path, updateStatus: (String) -> Unit) {
  val importedSheets = runCatching {
    EpicsDatabaseExcelImporter.importWorkbook(path)
  }.getOrElse { error ->
    updateStatus("Failed to import ${path.fileName}: ${error.message ?: "unknown error"}")
    Messages.showErrorDialog(project, error.message ?: "Failed to import ${path.fileName}.", "Import as EPICS Database")
    return
  }

  if (importedSheets.isEmpty()) {
    updateStatus("No EPICS-style sheets were found in ${path.fileName}.")
    Messages.showWarningDialog(project, "No EPICS-style sheets were found in ${path.fileName}.", "Import as EPICS Database")
    return
  }

  EpicsDatabaseExcelImporter.openImportedSheets(project, importedSheets)
  val sheetLabel = "${importedSheets.size} sheet${if (importedSheets.size == 1) "" else "s"}"
  updateStatus("Imported $sheetLabel from ${path.fileName}.")
}
