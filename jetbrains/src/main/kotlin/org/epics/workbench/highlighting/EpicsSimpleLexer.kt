package org.epics.workbench.highlighting

import com.intellij.lexer.LexerBase
import com.intellij.psi.TokenType
import com.intellij.psi.tree.IElementType

internal data class EpicsLexingProfile(
  val keywords: Set<String>,
)

internal class EpicsSimpleLexer(
  private val profile: EpicsLexingProfile,
) : LexerBase() {
  private var buffer: CharSequence = ""
  private var startOffset = 0
  private var endOffset = 0
  private var tokenStart = 0
  private var tokenEnd = 0
  private var tokenType: IElementType? = null

  override fun start(
    buffer: CharSequence,
    startOffset: Int,
    endOffset: Int,
    initialState: Int,
  ) {
    this.buffer = buffer
    this.startOffset = startOffset
    this.endOffset = endOffset
    locateToken(startOffset)
  }

  override fun getState(): Int = 0

  override fun getTokenType(): IElementType? = tokenType

  override fun getTokenStart(): Int = tokenStart

  override fun getTokenEnd(): Int = tokenEnd

  override fun advance() {
    locateToken(tokenEnd)
  }

  override fun getBufferSequence(): CharSequence = buffer

  override fun getBufferEnd(): Int = endOffset

  private fun locateToken(offset: Int) {
    if (offset >= endOffset) {
      tokenStart = endOffset
      tokenEnd = endOffset
      tokenType = null
      return
    }

    tokenStart = offset
    val current = buffer[offset]

    if (current.isWhitespace()) {
      tokenEnd = scanWhile(offset) { it.isWhitespace() }
      tokenType = TokenType.WHITE_SPACE
      return
    }

    if (current == '#') {
      tokenEnd = scanToLineEnd(offset)
      tokenType = EpicsTokenTypes.COMMENT
      return
    }

    if (current == '"') {
      tokenEnd = scanString(offset)
      tokenType = EpicsTokenTypes.STRING
      return
    }

    if (isMacroStart(offset)) {
      tokenEnd = scanMacro(offset)
      tokenType = EpicsTokenTypes.MACRO
      return
    }

    if (isNumberStart(offset)) {
      tokenEnd = scanNumber(offset)
      tokenType = EpicsTokenTypes.NUMBER
      return
    }

    if (isIdentifierStart(current)) {
      tokenEnd = scanWhile(offset + 1) { isIdentifierPart(it) }
      val text = buffer.subSequence(offset, tokenEnd).toString()
      tokenType = if (profile.keywords.contains(text)) EpicsTokenTypes.KEYWORD else EpicsTokenTypes.IDENTIFIER
      return
    }

    tokenEnd = offset + 1
    tokenType = when (current) {
      '{', '}' -> EpicsTokenTypes.BRACE
      '(', ')' -> EpicsTokenTypes.PAREN
      '[', ']' -> EpicsTokenTypes.BRACKET
      ',' -> EpicsTokenTypes.COMMA
      '=', '<', '>', ':', ';', '@' -> EpicsTokenTypes.OPERATOR
      else -> EpicsTokenTypes.TEXT
    }
  }

  private fun scanToLineEnd(offset: Int): Int {
    var index = offset + 1
    while (index < endOffset && buffer[index] != '\n' && buffer[index] != '\r') {
      index += 1
    }
    return index
  }

  private fun scanString(offset: Int): Int {
    var index = offset + 1
    while (index < endOffset) {
      when (buffer[index]) {
        '\\' -> {
          index += 2
        }
        '"' -> {
          return index + 1
        }
        else -> {
          index += 1
        }
      }
    }
    return endOffset
  }

  private fun scanMacro(offset: Int): Int {
    var index = offset + 2
    val terminator = if (buffer[offset + 1] == '(') ')' else '}'
    while (index < endOffset && buffer[index] != terminator) {
      index += 1
    }
    return if (index < endOffset) index + 1 else endOffset
  }

  private fun scanNumber(offset: Int): Int {
    var index = offset
    if (buffer[index] == '+' || buffer[index] == '-') {
      index += 1
    }

    if (index + 1 < endOffset && buffer[index] == '0' && (buffer[index + 1] == 'x' || buffer[index + 1] == 'X')) {
      index += 2
      while (index < endOffset && buffer[index].isHexDigit()) {
        index += 1
      }
      return index
    }

    index = scanWhile(index) { it.isDigit() }
    if (index < endOffset && buffer[index] == '.') {
      index += 1
      index = scanWhile(index) { it.isDigit() }
    }
    if (index < endOffset && (buffer[index] == 'e' || buffer[index] == 'E')) {
      val exponentStart = index
      index += 1
      if (index < endOffset && (buffer[index] == '+' || buffer[index] == '-')) {
        index += 1
      }
      val exponentDigits = scanWhile(index) { it.isDigit() }
      index = if (exponentDigits == index) exponentStart else exponentDigits
    }
    return index
  }

  private fun scanWhile(offset: Int, predicate: (Char) -> Boolean): Int {
    var index = offset
    while (index < endOffset && predicate(buffer[index])) {
      index += 1
    }
    return index
  }

  private fun isMacroStart(offset: Int): Boolean {
    return offset + 1 < endOffset && buffer[offset] == '$' && (buffer[offset + 1] == '(' || buffer[offset + 1] == '{')
  }

  private fun isNumberStart(offset: Int): Boolean {
    val current = buffer[offset]
    if (current.isDigit()) {
      return true
    }
    if (current == '.' && offset + 1 < endOffset) {
      return buffer[offset + 1].isDigit()
    }
    if ((current == '+' || current == '-') && offset + 1 < endOffset) {
      val next = buffer[offset + 1]
      if (next.isDigit()) {
        return true
      }
      return next == '.' && offset + 2 < endOffset && buffer[offset + 2].isDigit()
    }
    return false
  }

  private fun isIdentifierStart(value: Char): Boolean = value == '_' || value.isLetter()

  private fun isIdentifierPart(value: Char): Boolean = value == '_' || value.isLetterOrDigit()

  private fun Char.isHexDigit(): Boolean = isDigit() || lowercaseChar() in 'a'..'f'
}
