package org.epics.workbench.highlighting

import com.intellij.lexer.LexerBase
import com.intellij.psi.TokenType
import com.intellij.psi.tree.IElementType

internal class EpicsMonitorLexer : LexerBase() {
  private var buffer: CharSequence = ""
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

    val lineStart = findLineStart(offset)
    val lineEnd = findLineEnd(offset)
    val firstNonWhitespace = findFirstNonWhitespace(lineStart, lineEnd)
    if (firstNonWhitespace == -1) {
      tokenEnd = lineEnd
      tokenType = TokenType.WHITE_SPACE
      return
    }

    if (offset == firstNonWhitespace && buffer[offset] == '#') {
      tokenEnd = lineEnd
      tokenType = EpicsTokenTypes.COMMENT
      return
    }

    if (isMacroStart(offset)) {
      tokenEnd = scanMacro(offset)
      tokenType = EpicsTokenTypes.MACRO
      return
    }

    val assignment = parseAssignment(lineStart, lineEnd, firstNonWhitespace)
    if (assignment != null) {
      when {
        offset in assignment.nameStart until assignment.nameEnd -> {
          tokenStart = assignment.nameStart
          tokenEnd = assignment.nameEnd
          tokenType = EpicsTokenTypes.MACRO
          return
        }

        offset == assignment.equalsOffset -> {
          tokenStart = assignment.equalsOffset
          tokenEnd = assignment.equalsOffset + 1
          tokenType = EpicsTokenTypes.OPERATOR
          return
        }

        offset > assignment.equalsOffset -> {
          tokenEnd = scanValueToken(offset, lineEnd)
          tokenType = EpicsTokenTypes.STRING
          return
        }
      }
    }

    tokenEnd = scanRecordToken(offset, lineEnd)
    tokenType = EpicsTokenTypes.RECORD_NAME
  }

  private fun scanWhile(offset: Int, predicate: (Char) -> Boolean): Int {
    var index = offset
    while (index < endOffset && predicate(buffer[index])) {
      index += 1
    }
    return index
  }

  private fun findLineStart(offset: Int): Int {
    var index = offset
    while (index > 0 && buffer[index - 1] != '\n' && buffer[index - 1] != '\r') {
      index -= 1
    }
    return index
  }

  private fun findLineEnd(offset: Int): Int {
    var index = offset
    while (index < endOffset && buffer[index] != '\n' && buffer[index] != '\r') {
      index += 1
    }
    return index
  }

  private fun findFirstNonWhitespace(lineStart: Int, lineEnd: Int): Int {
    var index = lineStart
    while (index < lineEnd && buffer[index].isWhitespace()) {
      index += 1
    }
    return if (index < lineEnd) index else -1
  }

  private fun parseAssignment(
    lineStart: Int,
    lineEnd: Int,
    firstNonWhitespace: Int,
  ): AssignmentContext? {
    val equalsOffset = findEqualsOffset(lineStart, lineEnd)
    if (equalsOffset <= firstNonWhitespace) {
      return null
    }

    var nameEnd = equalsOffset
    while (nameEnd > firstNonWhitespace && buffer[nameEnd - 1].isWhitespace()) {
      nameEnd -= 1
    }
    if (nameEnd <= firstNonWhitespace) {
      return null
    }

    val name = buffer.subSequence(firstNonWhitespace, nameEnd).toString()
    if (!name.matches(ASSIGNMENT_NAME_REGEX)) {
      return null
    }

    return AssignmentContext(
      nameStart = firstNonWhitespace,
      nameEnd = nameEnd,
      equalsOffset = equalsOffset,
    )
  }

  private fun findEqualsOffset(lineStart: Int, lineEnd: Int): Int {
    var index = lineStart
    while (index < lineEnd) {
      if (buffer[index] == '=') {
        return index
      }
      index += 1
    }
    return -1
  }

  private fun scanValueToken(offset: Int, lineEnd: Int): Int {
    var index = offset
    while (index < lineEnd) {
      if (buffer[index].isWhitespace() || isMacroStart(index)) {
        break
      }
      index += 1
    }
    return if (index == offset) offset + 1 else index
  }

  private fun scanRecordToken(offset: Int, lineEnd: Int): Int {
    var index = offset
    while (index < lineEnd) {
      if (buffer[index].isWhitespace() || isMacroStart(index)) {
        break
      }
      index += 1
    }
    return if (index == offset) offset + 1 else index
  }

  private fun isMacroStart(offset: Int): Boolean {
    return offset + 1 < endOffset && buffer[offset] == '$' && (buffer[offset + 1] == '(' || buffer[offset + 1] == '{')
  }

  private fun scanMacro(offset: Int): Int {
    var index = offset + 2
    val terminator = if (buffer[offset + 1] == '(') ')' else '}'
    while (index < endOffset && buffer[index] != terminator) {
      index += 1
    }
    return if (index < endOffset) index + 1 else endOffset
  }

  private data class AssignmentContext(
    val nameStart: Int,
    val nameEnd: Int,
    val equalsOffset: Int,
  )

  private companion object {
    val ASSIGNMENT_NAME_REGEX = Regex("""[A-Za-z_][A-Za-z0-9_]*""")
  }
}
