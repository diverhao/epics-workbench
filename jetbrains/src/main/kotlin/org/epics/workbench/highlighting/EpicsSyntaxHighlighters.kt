package org.epics.workbench.highlighting

import com.intellij.lexer.Lexer
import com.intellij.openapi.editor.colors.TextAttributesKey
import com.intellij.openapi.fileTypes.SyntaxHighlighter
import com.intellij.openapi.fileTypes.SyntaxHighlighterBase
import com.intellij.openapi.fileTypes.SyntaxHighlighterFactory
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.TokenType
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

internal val PROTOCOL_KEYWORDS = setOf(
  "in",
  "out",
  "connect",
  "disconnect",
  "wait",
  "event",
  "exec",
  "init",
  "true",
  "false",
  "ignore",
  "cr",
  "lf",
  "nl",
  "terminator",
  "outterminator",
  "interminator",
  "replytimeout",
  "readtimeout",
  "writetimeout",
  "pollperiod",
  "extrainput",
  "separator",
  "locktimeout",
  "waittimeout",
)

internal val DBD_KEYWORDS = setOf(
  "menu",
  "choice",
  "recordtype",
  "field",
  "breaktable",
  "include",
  "device",
  "driver",
  "registrar",
  "function",
  "variable",
  "prompt",
  "promptgroup",
  "special",
  "size",
  "initial",
  "interest",
  "base",
  "pp",
  "asl",
  "extra",
)

internal val MAKEFILE_KEYWORDS = setOf(
  "include",
  "ifdef",
  "ifndef",
  "ifeq",
  "ifneq",
  "else",
  "endif",
  "define",
  "endef",
  "export",
  "unexport",
  "override",
  "private",
  "vpath",
)

private fun syntaxHighlights(tokenType: IElementType?): Array<TextAttributesKey> = when (tokenType) {
  TokenType.BAD_CHARACTER -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.BAD_CHARACTER)
  EpicsTokenTypes.COMMENT -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.COMMENT)
  EpicsTokenTypes.STRING -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.STRING)
  EpicsTokenTypes.KEYWORD -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.KEYWORD)
  EpicsTokenTypes.NUMBER -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.NUMBER)
  EpicsTokenTypes.MACRO -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.MACRO)
  EpicsTokenTypes.RECORD_NAME -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.RECORD_NAME)
  EpicsTokenTypes.BRACE -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.BRACE)
  EpicsTokenTypes.PAREN -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.PAREN)
  EpicsTokenTypes.BRACKET -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.BRACKET)
  EpicsTokenTypes.COMMA -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.COMMA)
  EpicsTokenTypes.OPERATOR -> SyntaxHighlighterBase.pack(EpicsHighlightingKeys.OPERATOR)
  else -> TextAttributesKey.EMPTY_ARRAY
}

private class EpicsSimpleSyntaxHighlighter(
  private val profile: EpicsLexingProfile,
) : SyntaxHighlighterBase() {
  override fun getHighlightingLexer(): Lexer = EpicsSimpleLexer(profile)

  override fun getTokenHighlights(tokenType: IElementType): Array<com.intellij.openapi.editor.colors.TextAttributesKey> {
    return syntaxHighlights(tokenType)
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

class EpicsDatabaseDefinitionSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(EpicsLexingProfile(DBD_KEYWORDS))
  }
}

class EpicsProtocolSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(
      EpicsLexingProfile(
        keywords = PROTOCOL_KEYWORDS,
        allowSingleQuotedStrings = true,
        extraIdentifierChars = setOf('-'),
        caseInsensitiveKeywords = true,
      ),
    )
  }
}

class EpicsMonitorSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return object : SyntaxHighlighterBase() {
      override fun getHighlightingLexer(): Lexer = EpicsMonitorLexer()

      override fun getTokenHighlights(tokenType: IElementType): Array<com.intellij.openapi.editor.colors.TextAttributesKey> {
        return syntaxHighlights(tokenType)
      }
    }
  }
}

class EpicsProbeSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return object : SyntaxHighlighterBase() {
      override fun getHighlightingLexer(): Lexer = EpicsProbeLexer()

      override fun getTokenHighlights(tokenType: IElementType): Array<com.intellij.openapi.editor.colors.TextAttributesKey> {
        return syntaxHighlights(tokenType)
      }
    }
  }
}

class EpicsMakefileSyntaxHighlighterFactory : SyntaxHighlighterFactory() {
  override fun getSyntaxHighlighter(project: Project?, virtualFile: VirtualFile?): SyntaxHighlighter {
    return EpicsSimpleSyntaxHighlighter(
      EpicsLexingProfile(
        keywords = MAKEFILE_KEYWORDS,
        extraIdentifierChars = setOf('-', '.'),
      ),
    )
  }
}
