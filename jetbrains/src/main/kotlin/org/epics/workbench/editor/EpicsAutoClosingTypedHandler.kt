package org.epics.workbench.editor

import com.intellij.codeInsight.editorActions.TypedHandlerDelegate
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileTypes.FileType
import com.intellij.openapi.project.Project
import com.intellij.psi.PsiFile
import org.epics.workbench.completion.EpicsDatabaseCompletionContributor

class EpicsAutoClosingTypedHandler : TypedHandlerDelegate() {
  override fun beforeSelectionRemoved(
    charTyped: Char,
    project: Project,
    editor: Editor,
    file: PsiFile,
  ): Result {
    if (!isEpicsFile(file)) {
      return Result.CONTINUE
    }

    val closingChar = OPEN_TO_CLOSE[charTyped] ?: return Result.CONTINUE
    val selectionModel = editor.selectionModel
    if (!selectionModel.hasSelection()) {
      return Result.CONTINUE
    }

    val selectionStart = selectionModel.selectionStart
    val selectionEnd = selectionModel.selectionEnd
    val selectedText = selectionModel.selectedText ?: ""
    editor.document.replaceString(
      selectionStart,
      selectionEnd,
      buildString(selectedText.length + 2) {
        append(charTyped)
        append(selectedText)
        append(closingChar)
      },
    )
    selectionModel.removeSelection()
    editor.caretModel.moveToOffset(selectionEnd + 2)
    return Result.STOP
  }

  override fun beforeCharTyped(
    charTyped: Char,
    project: Project,
    editor: Editor,
    file: PsiFile,
    fileType: FileType,
  ): Result {
    if (!isEpicsFile(file)) {
      return Result.CONTINUE
    }

    if (charTyped in CLOSE_TO_OPEN && isClosingCharAtCaret(editor, charTyped)) {
      editor.caretModel.moveToOffset(editor.caretModel.offset + 1)
      return Result.STOP
    }

    val closingChar = OPEN_TO_CLOSE[charTyped] ?: return Result.CONTINUE
    if (!shouldAutoCloseBeforeInsert(editor, charTyped, closingChar)) {
      return Result.CONTINUE
    }

    val offset = editor.caretModel.offset
    editor.document.insertString(offset, "$charTyped$closingChar")
    editor.caretModel.moveToOffset(offset + 1)
    EpicsDatabaseCompletionContributor.maybeScheduleAutoPopupForTypedChar(
      file = file,
      editor = editor,
      project = project,
      charTyped = charTyped,
    )
    return Result.STOP
  }

  override fun charTyped(
    charTyped: Char,
    project: Project,
    editor: Editor,
    file: PsiFile,
  ): Result {
    if (isEpicsFile(file)) {
      EpicsDatabaseCompletionContributor.maybeScheduleAutoPopupForTypedChar(
        file = file,
        editor = editor,
        project = project,
        charTyped = charTyped,
      )
    }
    return Result.CONTINUE
  }

  private fun isEpicsFile(file: PsiFile): Boolean {
    val extension = file.virtualFile?.extension?.lowercase()
      ?: file.name.substringAfterLast('.', "").lowercase()
    return file.fileType.name in EPICS_FILE_TYPES ||
      file.language.id in EPICS_LANGUAGE_IDS ||
      extension in EPICS_EXTENSIONS
  }

  private fun isClosingCharAtCaret(editor: Editor, closingChar: Char): Boolean {
    return charAt(editor, editor.caretModel.offset) == closingChar
  }

  private fun shouldAutoCloseBeforeInsert(editor: Editor, openingChar: Char, closingChar: Char): Boolean {
    val offset = editor.caretModel.offset
    val nextChar = charAt(editor, offset)
    if (nextChar == closingChar) {
      return false
    }
    if (openingChar == '"' && charAt(editor, offset - 1) == '\\') {
      return false
    }
    if (nextChar == null) {
      return true
    }
    if (nextChar.isWhitespace()) {
      return true
    }
    if (nextChar in SAFE_FOLLOWING_CHARS || nextChar in CLOSE_TO_OPEN) {
      return true
    }
    if (openingChar == '"') {
      return !nextChar.isLetterOrDigit() && nextChar != '_'
    }
    return nextChar !in OPEN_TO_CLOSE.keys
  }

  private fun charAt(editor: Editor, offset: Int): Char? {
    val chars = editor.document.charsSequence
    if (offset < 0 || offset >= chars.length) {
      return null
    }
    return chars[offset]
  }

  companion object {
    private val OPEN_TO_CLOSE = linkedMapOf(
      '{' to '}',
      '(' to ')',
      '"' to '"',
    )
    private val CLOSE_TO_OPEN = OPEN_TO_CLOSE.entries.associate { (open, close) -> close to open }
    private val SAFE_FOLLOWING_CHARS = setOf(')', '}', ']', ',', ';', ':')
    private val EPICS_FILE_TYPES = setOf(
      "EPICS Database",
      "EPICS Substitutions",
      "EPICS Startup",
      "EPICS Database Definition",
      "EPICS Protocol",
      "EPICS Sequencer",
      "EPICS Monitor",
    )
    private val EPICS_LANGUAGE_IDS = EPICS_FILE_TYPES
    private val EPICS_EXTENSIONS = setOf(
      "db",
      "vdb",
      "template",
      "substitutions",
      "sub",
      "subs",
      "cmd",
      "iocsh",
      "dbd",
      "proto",
      "st",
      "monitor",
    )
  }
}
