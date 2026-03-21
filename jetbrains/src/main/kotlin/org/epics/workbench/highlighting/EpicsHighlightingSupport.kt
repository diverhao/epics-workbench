package org.epics.workbench.highlighting

import com.intellij.lang.Language
import com.intellij.openapi.editor.DefaultLanguageHighlighterColors
import com.intellij.openapi.editor.HighlighterColors
import com.intellij.openapi.editor.colors.TextAttributesKey
import com.intellij.openapi.fileTypes.SyntaxHighlighterBase
import com.intellij.psi.TokenType
import com.intellij.psi.tree.IElementType

private object EpicsHighlightingLanguage : Language("EPICS Highlighting")

internal class EpicsTokenType(debugName: String) : IElementType(debugName, EpicsHighlightingLanguage)

internal object EpicsTokenTypes {
  val COMMENT = EpicsTokenType("EPICS_COMMENT")
  val STRING = EpicsTokenType("EPICS_STRING")
  val KEYWORD = EpicsTokenType("EPICS_KEYWORD")
  val NUMBER = EpicsTokenType("EPICS_NUMBER")
  val MACRO = EpicsTokenType("EPICS_MACRO")
  val RECORD_NAME = EpicsTokenType("EPICS_RECORD_NAME")
  val BRACE = EpicsTokenType("EPICS_BRACE")
  val PAREN = EpicsTokenType("EPICS_PAREN")
  val BRACKET = EpicsTokenType("EPICS_BRACKET")
  val COMMA = EpicsTokenType("EPICS_COMMA")
  val OPERATOR = EpicsTokenType("EPICS_OPERATOR")
  val IDENTIFIER = EpicsTokenType("EPICS_IDENTIFIER")
  val TEXT = EpicsTokenType("EPICS_TEXT")
}

internal object EpicsHighlightingKeys {
  val COMMENT = TextAttributesKey.createTextAttributesKey(
    "EPICS_COMMENT",
    DefaultLanguageHighlighterColors.LINE_COMMENT,
  )
  val STRING = TextAttributesKey.createTextAttributesKey(
    "EPICS_STRING",
    DefaultLanguageHighlighterColors.STRING,
  )
  val KEYWORD = TextAttributesKey.createTextAttributesKey(
    "EPICS_KEYWORD",
    DefaultLanguageHighlighterColors.KEYWORD,
  )
  val NUMBER = TextAttributesKey.createTextAttributesKey(
    "EPICS_NUMBER",
    DefaultLanguageHighlighterColors.NUMBER,
  )
  val MACRO = TextAttributesKey.createTextAttributesKey(
    "EPICS_MACRO",
    DefaultLanguageHighlighterColors.PARAMETER,
  )
  val RECORD_NAME = TextAttributesKey.createTextAttributesKey(
    "EPICS_RECORD_NAME",
    DefaultLanguageHighlighterColors.INSTANCE_FIELD,
  )
  val BRACE = TextAttributesKey.createTextAttributesKey(
    "EPICS_BRACE",
    DefaultLanguageHighlighterColors.BRACES,
  )
  val PAREN = TextAttributesKey.createTextAttributesKey(
    "EPICS_PAREN",
    DefaultLanguageHighlighterColors.PARENTHESES,
  )
  val BRACKET = TextAttributesKey.createTextAttributesKey(
    "EPICS_BRACKET",
    DefaultLanguageHighlighterColors.BRACKETS,
  )
  val COMMA = TextAttributesKey.createTextAttributesKey(
    "EPICS_COMMA",
    DefaultLanguageHighlighterColors.COMMA,
  )
  val OPERATOR = TextAttributesKey.createTextAttributesKey(
    "EPICS_OPERATOR",
    DefaultLanguageHighlighterColors.OPERATION_SIGN,
  )
  val BAD_CHARACTER = TextAttributesKey.createTextAttributesKey(
    "EPICS_BAD_CHARACTER",
    HighlighterColors.BAD_CHARACTER,
  )
}

internal fun epicsHighlights(tokenType: IElementType?): Array<TextAttributesKey> = when (tokenType) {
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
