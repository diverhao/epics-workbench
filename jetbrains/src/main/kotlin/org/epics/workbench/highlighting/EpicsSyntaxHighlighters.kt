package org.epics.workbench.highlighting

import com.intellij.lexer.Lexer
import com.intellij.openapi.fileTypes.SyntaxHighlighter
import com.intellij.openapi.fileTypes.SyntaxHighlighterBase
import com.intellij.openapi.fileTypes.SyntaxHighlighterFactory
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.tree.IElementType

internal val DATABASE_KEYWORDS = setOf(
  "record",
  "grecord",
  "field",
  "info",
  "alias",
  "menu",
  "choice",
  "device",
  "driver",
  "registrar",
  "function",
  "variable",
  "include",
  "breaktable",
)

internal val SUBSTITUTIONS_KEYWORDS = setOf(
  "file",
  "global",
  "pattern",
)

internal val STARTUP_KEYWORDS = setOf(
  "dbLoadDatabase",
  "dbLoadRecords",
  "dbLoadTemplate",
  "epicsEnvSet",
  "iocInit",
  "cd",
  "var",
  "require",
  "seq",
  "exec",
  "help",
  "dbl",
  "dbpf",
  "dbpr",
  "dbgf",
  "dbtr",
  "dbnr",
  "errlogInit",
)

private class EpicsSimpleSyntaxHighlighter(
  private val profile: EpicsLexingProfile,
) : SyntaxHighlighterBase() {
  override fun getHighlightingLexer(): Lexer = EpicsSimpleLexer(profile)

  override fun getTokenHighlights(tokenType: IElementType): Array<com.intellij.openapi.editor.colors.TextAttributesKey> {
    return epicsHighlights(tokenType)
  }
}

class EpicsDatabaseSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(EpicsLexingProfile(DATABASE_KEYWORDS))
  }
}

class EpicsSubstitutionsSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(EpicsLexingProfile(SUBSTITUTIONS_KEYWORDS))
  }
}

class EpicsStartupSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(EpicsLexingProfile(STARTUP_KEYWORDS))
  }
}

class EpicsMonitorSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return object : SyntaxHighlighterBase() {
      override fun getHighlightingLexer(): Lexer = EpicsMonitorLexer()

      override fun getTokenHighlights(tokenType: IElementType): Array<com.intellij.openapi.editor.colors.TextAttributesKey> {
        return epicsHighlights(tokenType)
      }
    }
  }
}
