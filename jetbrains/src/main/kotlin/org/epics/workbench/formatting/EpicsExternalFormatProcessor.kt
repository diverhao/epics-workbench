package org.epics.workbench.formatting

import com.intellij.application.options.CodeStyle
import com.intellij.openapi.editor.Document
import com.intellij.openapi.util.TextRange
import com.intellij.psi.PsiDocumentManager
import com.intellij.psi.PsiFile
import com.intellij.psi.codeStyle.ExternalFormatProcessor

class EpicsExternalFormatProcessor : ExternalFormatProcessor {
  override fun activeForFile(file: PsiFile): Boolean {
    return resolveFormatKind(file) != null
  }

  override fun format(
    file: PsiFile,
    range: TextRange,
    canChangeWhiteSpacesOnly: Boolean,
    quickFormat: Boolean,
    reformatContext: Boolean,
    indentAdjustment: Int,
  ): TextRange {
    val documentManager = PsiDocumentManager.getInstance(file.project)
    val document = documentManager.getDocument(file) ?: return range
    val formattedText = formatFileText(file, document.text) ?: return range

    if (formattedText != document.text) {
      document.replaceString(0, document.textLength, formattedText)
      documentManager.commitDocument(document)
    }

    return TextRange(0, document.textLength)
  }

  override fun indent(file: PsiFile, offset: Int): String {
    val document = PsiDocumentManager.getInstance(file.project).getDocument(file) ?: return ""
    val formattedText = formatFileText(file, document.text) ?: return currentLineIndent(document, offset)
    return lineIndentForOffset(formattedText, document, offset)
  }

  override fun getId(): String = "epics-workbench"

  private fun formatFileText(file: PsiFile, text: String): String? {
    return when (resolveFormatKind(file)) {
      FormatKind.DATABASE -> EpicsTextFormatter.formatDatabaseText(text, getIndentUnit(file))
      FormatKind.STARTUP -> EpicsTextFormatter.formatStartupText(text)
      FormatKind.SUBSTITUTIONS -> EpicsTextFormatter.formatSubstitutionText(text, getIndentUnit(file))
      FormatKind.MAKEFILE -> EpicsTextFormatter.formatMakefileText(text)
      FormatKind.PROTOCOL -> EpicsTextFormatter.formatProtocolText(text, getIndentUnit(file))
      FormatKind.MONITOR -> EpicsTextFormatter.formatMonitorText(text)
      FormatKind.SEQUENCER -> EpicsTextFormatter.formatSequencerText(text)
      null -> null
    }
  }

  private fun getIndentUnit(file: PsiFile): String {
    val indentOptions = CodeStyle.getIndentOptions(file)
    return if (indentOptions.USE_TAB_CHARACTER) {
      "\t"
    } else {
      " ".repeat(indentOptions.INDENT_SIZE.coerceAtLeast(DEFAULT_INDENT_SIZE))
    }
  }

  private fun lineIndentForOffset(formattedText: String, originalDocument: Document, offset: Int): String {
    val safeOffset = offset.coerceIn(0, originalDocument.textLength)
    val originalLine = originalDocument.getLineNumber(safeOffset)
    val formattedLines = formattedText.replace("\r\n", "\n").split('\n')
    val formattedLine = formattedLines.getOrElse(originalLine) { formattedLines.lastOrNull().orEmpty() }
    return formattedLine.takeWhile { it == ' ' || it == '\t' }
  }

  private fun currentLineIndent(document: Document, offset: Int): String {
    val safeOffset = offset.coerceIn(0, document.textLength)
    val lineNumber = document.getLineNumber(safeOffset)
    val lineStart = document.getLineStartOffset(lineNumber)
    val lineEnd = document.getLineEndOffset(lineNumber)
    val lineText = document.getText(TextRange(lineStart, lineEnd))
    return lineText.takeWhile { it == ' ' || it == '\t' }
  }

  private fun resolveFormatKind(file: PsiFile): FormatKind? {
    val extension = file.virtualFile?.extension?.lowercase()
      ?: file.name.substringAfterLast('.', "").lowercase()

    return when {
      extension in DATABASE_EXTENSIONS -> FormatKind.DATABASE
      extension in STARTUP_EXTENSIONS -> FormatKind.STARTUP
      extension in SUBSTITUTIONS_EXTENSIONS -> FormatKind.SUBSTITUTIONS
      extension == PROTOCOL_EXTENSION -> FormatKind.PROTOCOL
      extension == MONITOR_EXTENSION -> FormatKind.MONITOR
      extension == SEQUENCER_EXTENSION -> FormatKind.SEQUENCER
      isMakefile(file) -> FormatKind.MAKEFILE
      file.fileType.name == DATABASE_FILE_TYPE || file.language.id == DATABASE_LANGUAGE_ID -> FormatKind.DATABASE
      file.fileType.name == STARTUP_FILE_TYPE || file.language.id == STARTUP_LANGUAGE_ID -> FormatKind.STARTUP
      file.fileType.name == SUBSTITUTIONS_FILE_TYPE || file.language.id == SUBSTITUTIONS_LANGUAGE_ID -> FormatKind.SUBSTITUTIONS
      file.fileType.name == PROTOCOL_FILE_TYPE || file.language.id == PROTOCOL_LANGUAGE_ID -> FormatKind.PROTOCOL
      file.fileType.name == MONITOR_FILE_TYPE || file.language.id == MONITOR_LANGUAGE_ID -> FormatKind.MONITOR
      file.fileType.name == SEQUENCER_FILE_TYPE || file.language.id == SEQUENCER_LANGUAGE_ID -> FormatKind.SEQUENCER
      else -> null
    }
  }

  private fun isMakefile(file: PsiFile): Boolean = file.virtualFile?.name == "Makefile" || file.name == "Makefile"

  companion object {
    private const val DATABASE_LANGUAGE_ID = "EPICS Database"
    private const val STARTUP_LANGUAGE_ID = "EPICS Startup"
    private const val SUBSTITUTIONS_LANGUAGE_ID = "EPICS Substitutions"
    private const val PROTOCOL_LANGUAGE_ID = "EPICS Protocol"
    private const val SEQUENCER_LANGUAGE_ID = "EPICS Sequencer"
    private const val MONITOR_LANGUAGE_ID = "EPICS PV List"
    private const val DATABASE_FILE_TYPE = "EPICS Database"
    private const val STARTUP_FILE_TYPE = "EPICS Startup"
    private const val SUBSTITUTIONS_FILE_TYPE = "EPICS Substitutions"
    private const val PROTOCOL_FILE_TYPE = "EPICS Protocol"
    private const val SEQUENCER_FILE_TYPE = "EPICS Sequencer"
    private const val MONITOR_FILE_TYPE = "EPICS PV List"
    private const val DEFAULT_INDENT_SIZE = 4
    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
    private val STARTUP_EXTENSIONS = setOf("cmd", "iocsh")
    private val SUBSTITUTIONS_EXTENSIONS = setOf("substitutions", "sub", "subs")
    private const val PROTOCOL_EXTENSION = "proto"
    private const val SEQUENCER_EXTENSION = "st"
    private const val MONITOR_EXTENSION = "pvlist"
  }

  private enum class FormatKind {
    DATABASE,
    STARTUP,
    SUBSTITUTIONS,
    MAKEFILE,
    PROTOCOL,
    MONITOR,
    SEQUENCER,
  }
}
