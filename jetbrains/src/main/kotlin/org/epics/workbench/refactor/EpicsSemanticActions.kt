package org.epics.workbench.refactor

import com.intellij.openapi.actionSystem.ActionUpdateThread
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.actionSystem.CommonDataKeys
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.fileEditor.OpenFileDescriptor
import com.intellij.openapi.project.DumbAwareAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.InputValidatorEx
import com.intellij.openapi.ui.Messages
import com.intellij.openapi.ui.popup.JBPopupFactory
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.ui.ColoredListCellRenderer
import com.intellij.ui.SimpleTextAttributes
import java.nio.file.Path
import javax.swing.JList

class FindEpicsReferencesAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val symbol = resolveSemanticContext(event)?.symbol
    event.presentation.isEnabledAndVisible = symbol != null
  }

  override fun actionPerformed(event: AnActionEvent) {
    val context = resolveSemanticContext(event) ?: return
    val occurrences = EpicsSemanticSymbolSupport.collectOccurrences(
      project = context.project,
      symbol = context.symbol,
      includeDeclarations = true,
    )
    if (occurrences.isEmpty()) {
      Messages.showInfoMessage(
        context.project,
        "No EPICS references were found for \"${context.symbol.name}\".",
        "Find EPICS References",
      )
      return
    }

    val popup = JBPopupFactory.getInstance()
      .createPopupChooserBuilder(occurrences)
      .setTitle("EPICS References for ${context.symbol.name} (${occurrences.size})")
      .setNamerForFiltering { occurrence ->
        buildReferencePresentation(context.project, occurrence)
      }
      .setRenderer(EpicsReferenceOccurrenceRenderer(context.project))
      .setItemChosenCallback { occurrence ->
        OpenFileDescriptor(context.project, occurrence.file, occurrence.startOffset).navigate(true)
      }
      .setResizable(true)
      .setMovable(true)
      .createPopup()

    if (context.editor != null) {
      popup.showInBestPositionFor(context.editor)
    } else {
      popup.showCenteredInCurrentWindow(context.project)
    }
  }
}

class RenameEpicsSymbolAction : DumbAwareAction() {
  override fun getActionUpdateThread(): ActionUpdateThread = ActionUpdateThread.BGT

  override fun update(event: AnActionEvent) {
    val symbol = resolveSemanticContext(event)?.symbol
    event.presentation.isEnabledAndVisible = symbol != null
  }

  override fun actionPerformed(event: AnActionEvent) {
    val context = resolveSemanticContext(event) ?: return
    val validator = object : InputValidatorEx {
      override fun getErrorText(inputString: String?): String? {
        val candidate = inputString?.trim().orEmpty()
        return EpicsSemanticSymbolSupport.validateRename(context.symbol, candidate)
          ?: EpicsSemanticSymbolSupport.detectRenameConflict(context.project, context.symbol, candidate)
      }

      override fun checkInput(inputString: String?): Boolean {
        return getErrorText(inputString) == null
      }

      override fun canClose(inputString: String?): Boolean {
        return checkInput(inputString)
      }
    }

    val newName = Messages.showInputDialog(
      context.project,
      "Rename EPICS ${describeSymbol(context.symbol)}:",
      "Rename EPICS Symbol",
      null,
      context.symbol.name,
      validator,
    )?.trim() ?: return

    if (newName == context.symbol.name) {
      return
    }

    val updatedCount = EpicsSemanticSymbolSupport.applyRename(
      project = context.project,
      symbol = context.symbol,
      newName = newName,
    )
    if (updatedCount == 0) {
      Messages.showInfoMessage(
        context.project,
        "No EPICS occurrences were updated for \"${context.symbol.name}\".",
        "Rename EPICS Symbol",
      )
      return
    }

    Messages.showInfoMessage(
      context.project,
      "Renamed ${describeSymbol(context.symbol)} to \"$newName\" in $updatedCount location(s).",
      "Rename EPICS Symbol",
    )
  }
}

private data class EpicsSemanticActionContext(
  val project: Project,
  val editor: Editor?,
  val file: VirtualFile,
  val symbol: EpicsSemanticSymbol,
)

private fun resolveSemanticContext(event: AnActionEvent): EpicsSemanticActionContext? {
  val project = event.project ?: return null
  val editor = event.getData(CommonDataKeys.EDITOR)
  val file = event.getData(CommonDataKeys.VIRTUAL_FILE)
    ?: event.getData(CommonDataKeys.PSI_FILE)?.virtualFile
    ?: editor?.let { currentEditor ->
      FileDocumentManager.getInstance().getFile(currentEditor.document)
    }
    ?: return null
  val offset = editor?.caretModel?.offset ?: return null
  val symbol = EpicsSemanticSymbolSupport.findSymbol(project, file, offset) ?: return null
  return EpicsSemanticActionContext(project, editor, file, symbol)
}

private fun describeSymbol(symbol: EpicsSemanticSymbol): String {
  return when (symbol.kind) {
    EpicsSemanticSymbolKind.RECORD -> "record"
    EpicsSemanticSymbolKind.MACRO -> "macro"
    EpicsSemanticSymbolKind.RECORD_TYPE -> "record type"
    EpicsSemanticSymbolKind.FIELD -> "field"
    EpicsSemanticSymbolKind.DEVICE_SUPPORT -> "device support"
    EpicsSemanticSymbolKind.DRIVER -> "driver"
    EpicsSemanticSymbolKind.REGISTRAR -> "registrar"
    EpicsSemanticSymbolKind.FUNCTION -> "function"
    EpicsSemanticSymbolKind.VARIABLE -> "variable"
  }
}

private fun buildReferencePresentation(project: Project, occurrence: EpicsSemanticOccurrence): String {
  val location = buildRelativePath(project, occurrence.file)
  val kind = if (occurrence.declaration) "declaration" else "reference"
  return "$location:${occurrence.lineNumber} [$kind] ${occurrence.lineText}"
}

private fun buildRelativePath(project: Project, file: VirtualFile): String {
  val basePath = project.basePath ?: return file.path
  return runCatching {
    Path.of(basePath).relativize(Path.of(file.path)).toString()
  }.getOrDefault(file.path)
}

private class EpicsReferenceOccurrenceRenderer(
  private val project: Project,
) : ColoredListCellRenderer<EpicsSemanticOccurrence>() {
  override fun customizeCellRenderer(
    list: JList<out EpicsSemanticOccurrence>,
    value: EpicsSemanticOccurrence?,
    index: Int,
    selected: Boolean,
    hasFocus: Boolean,
  ) {
    if (value == null) {
      return
    }
    append(
      if (value.declaration) "Declaration  " else "Reference   ",
      SimpleTextAttributes.GRAYED_ATTRIBUTES,
    )
    append("${buildRelativePath(project, value.file)}:${value.lineNumber}  ")
    append(value.lineText.trim(), SimpleTextAttributes.REGULAR_ATTRIBUTES)
  }
}
