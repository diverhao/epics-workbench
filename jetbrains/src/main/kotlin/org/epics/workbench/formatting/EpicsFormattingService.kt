package org.epics.workbench.formatting

import com.intellij.application.options.CodeStyle
import com.intellij.formatting.FormattingContext
import com.intellij.formatting.service.AbstractDocumentFormattingService
import com.intellij.formatting.service.FormattingService
import com.intellij.openapi.editor.Document
import com.intellij.openapi.util.TextRange
import com.intellij.psi.PsiFile

class EpicsFormattingService : AbstractDocumentFormattingService() {
  override fun getFeatures(): MutableSet<FormattingService.Feature> {
    return mutableSetOf(
      FormattingService.Feature.AD_HOC_FORMATTING,
      FormattingService.Feature.FORMAT_FRAGMENTS,
    )
  }

  override fun canFormat(file: PsiFile): Boolean {
    return resolveFormatKind(file) != null
  }

  override fun formatDocument(
    document: Document,
    formattingRanges: List<TextRange>,
    formattingContext: FormattingContext,
    canChangeWhiteSpaceOnly: Boolean,
    quickFormat: Boolean,
  ) {
    val file = formattingContext.containingFile
    val formattedText = when (resolveFormatKind(file)) {
      FormatKind.DATABASE -> EpicsTextFormatter.formatDatabaseText(
        document.text,
        getIndentUnit(file),
      )
      FormatKind.STARTUP -> EpicsTextFormatter.formatStartupText(document.text)
      FormatKind.SUBSTITUTIONS -> EpicsTextFormatter.formatSubstitutionText(
        document.text,
        getIndentUnit(file),
      )
      FormatKind.MAKEFILE -> EpicsTextFormatter.formatMakefileText(document.text)
      FormatKind.PROTOCOL -> EpicsTextFormatter.formatProtocolText(
        document.text,
        getIndentUnit(file),
      )
      FormatKind.MONITOR -> EpicsTextFormatter.formatMonitorText(document.text)
      FormatKind.SEQUENCER -> EpicsTextFormatter.formatSequencerText(document.text)
      else -> return
    }

    if (formattedText != document.text) {
      document.replaceString(0, document.textLength, formattedText)
    }
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

  private fun getIndentUnit(file: PsiFile): String {
    val indentOptions = CodeStyle.getIndentOptions(file)
    return if (indentOptions.USE_TAB_CHARACTER) {
      "\t"
    } else {
      " ".repeat(indentOptions.INDENT_SIZE.coerceAtLeast(DEFAULT_INDENT_SIZE))
    }
  }

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
