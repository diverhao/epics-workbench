package org.epics.workbench.protocol

import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.vfs.LocalFileSystem
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import java.nio.file.Files
import java.nio.file.Path
import java.util.concurrent.ConcurrentHashMap
import kotlin.io.path.pathString

internal data class StreamProtocolInvocation(
  val protocolPath: String,
  val commandName: String?,
  val commandStart: Int,
  val commandEnd: Int,
  val busValue: String?,
  val busStart: Int,
  val busEnd: Int,
)

internal data class StreamProtocolCommandDefinition(
  val name: String,
  val startOffset: Int,
  val endOffset: Int,
  val line: Int,
  val definitionText: String,
)

internal data class StreamProtocolCommandReference(
  val recordType: String,
  val recordName: String,
  val fieldName: String,
  val protocolPath: String,
  val commandName: String,
  val commandStart: Int,
  val commandEnd: Int,
)

internal object EpicsStreamProtocolSupport {
  private data class StreamProtocolCommandCacheEntry(
    val cacheTag: String,
    val commandDefinitions: List<StreamProtocolCommandDefinition>,
  )

  private data class StreamProtocolTextSnapshot(
    val cacheTag: String,
    val text: String,
  )

  private val streamProtocolCommandCache = ConcurrentHashMap<String, StreamProtocolCommandCacheEntry>()
  private val streamProtocolCommandDefinitionRegex =
    Regex("""([A-Za-z_][A-Za-z0-9_-]*)\s*\{""")

  internal fun extractInvocation(
    fieldDeclaration: EpicsRecordCompletionSupport.FieldDeclaration,
  ): StreamProtocolInvocation? {
    return extractInvocation(fieldDeclaration.value, fieldDeclaration.valueStart)
  }

  internal fun extractInvocation(
    fieldValue: String,
    valueStart: Int,
  ): StreamProtocolInvocation? {
    var index = 0
    while (index < fieldValue.length && fieldValue[index].isWhitespace()) {
      index += 1
    }
    if (index >= fieldValue.length || fieldValue[index] != '@') {
      return null
    }

    val protocolStart = index + 1
    var protocolEnd = protocolStart
    while (protocolEnd < fieldValue.length && !fieldValue[protocolEnd].isWhitespace() && fieldValue[protocolEnd] !in "\"'`") {
      protocolEnd += 1
    }

    val protocolPath = fieldValue.substring(protocolStart, protocolEnd)
    if (protocolPath.isBlank() || EpicsRecordCompletionSupport.containsEpicsMacroReference(protocolPath)) {
      return null
    }

    index = protocolEnd
    while (index < fieldValue.length && fieldValue[index].isWhitespace()) {
      index += 1
    }

    val commandStartIndex = index
    if (commandStartIndex >= fieldValue.length || !isStreamProtocolCommandStart(fieldValue[commandStartIndex])) {
      return StreamProtocolInvocation(
        protocolPath = protocolPath,
        commandName = null,
        commandStart = valueStart + commandStartIndex,
        commandEnd = valueStart + commandStartIndex,
        busValue = null,
        busStart = valueStart + commandStartIndex,
        busEnd = valueStart + commandStartIndex,
      )
    }

    index += 1
    while (index < fieldValue.length && isStreamProtocolCommandPart(fieldValue[index])) {
      index += 1
    }

    val commandName = fieldValue.substring(commandStartIndex, index)
    val commandEndIndex = parseOptionalStreamProtocolArguments(fieldValue, index)
    val bus = parseStreamProtocolBus(fieldValue, commandEndIndex)
    return StreamProtocolInvocation(
      protocolPath = protocolPath,
      commandName = commandName,
      commandStart = valueStart + commandStartIndex,
      commandEnd = valueStart + commandEndIndex,
      busValue = bus?.value,
      busStart = valueStart + (bus?.start ?: commandEndIndex),
      busEnd = valueStart + (bus?.end ?: commandEndIndex),
    )
  }

  internal fun getCommandNames(protocolPath: Path): Set<String> {
    return getCommandDefinitions(protocolPath).mapTo(linkedSetOf()) { it.name }
  }

  internal fun findCommandDefinition(
    protocolPath: Path,
    commandName: String,
  ): StreamProtocolCommandDefinition? {
    return getCommandDefinitions(protocolPath).firstOrNull { it.name == commandName }
  }

  internal fun findCommandReferenceAtOffset(
    text: String,
    offset: Int,
  ): StreamProtocolCommandReference? {
    val recordDeclaration = EpicsRecordCompletionSupport.extractRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.recordStart..declaration.recordEnd }
      ?: return null
    val fieldDeclarations = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, recordDeclaration)
    if (!isStreamDeviceRecord(fieldDeclarations)) {
      return null
    }

    val fieldDeclaration = fieldDeclarations
      .firstOrNull { declaration -> offset in declaration.valueStart..declaration.valueEnd }
      ?: return null
    if (!isLinkField(recordDeclaration.recordType, fieldDeclaration.fieldName)) {
      return null
    }

    val invocation = extractInvocation(fieldDeclaration) ?: return null
    val commandName = invocation.commandName ?: return null
    if (offset !in invocation.commandStart until invocation.commandEnd) {
      return null
    }

    return StreamProtocolCommandReference(
      recordType = recordDeclaration.recordType,
      recordName = recordDeclaration.name,
      fieldName = fieldDeclaration.fieldName,
      protocolPath = invocation.protocolPath,
      commandName = commandName,
      commandStart = invocation.commandStart,
      commandEnd = invocation.commandEnd,
    )
  }

  private fun getCommandDefinitions(protocolPath: Path): List<StreamProtocolCommandDefinition> {
    val normalizedPath = protocolPath.normalize().pathString
    val snapshot = readProtocolTextSnapshot(protocolPath) ?: return emptyList()

    streamProtocolCommandCache[normalizedPath]?.takeIf { it.cacheTag == snapshot.cacheTag }?.let { entry ->
      return entry.commandDefinitions
    }

    val commandDefinitions = extractCommandDefinitions(snapshot.text)
    streamProtocolCommandCache[normalizedPath] = StreamProtocolCommandCacheEntry(
      cacheTag = snapshot.cacheTag,
      commandDefinitions = commandDefinitions,
    )
    return commandDefinitions
  }

  private fun readProtocolTextSnapshot(protocolPath: Path): StreamProtocolTextSnapshot? {
    val virtualFile = LocalFileSystem.getInstance().findFileByNioFile(protocolPath)
    return if (virtualFile != null) {
      val document = FileDocumentManager.getInstance().getCachedDocument(virtualFile)
      if (document != null) {
        StreamProtocolTextSnapshot(
          cacheTag = "open:${document.modificationStamp}",
          text = document.text,
        )
      } else {
        val text = runCatching {
          virtualFile.inputStream.bufferedReader().use { it.readText() }
        }.getOrElse { return null }
        StreamProtocolTextSnapshot(
          cacheTag = "vfs:${virtualFile.modificationStamp}",
          text = text,
        )
      }
    } else {
      val size = runCatching { Files.size(protocolPath) }.getOrNull() ?: return null
      val modified = runCatching { Files.getLastModifiedTime(protocolPath).toMillis() }.getOrNull() ?: return null
      val text = runCatching { Files.readString(protocolPath) }.getOrElse { return null }
      StreamProtocolTextSnapshot(
        cacheTag = "fs:$size:$modified",
        text = text,
      )
    }
  }

  private fun extractCommandDefinitions(text: String): List<StreamProtocolCommandDefinition> {
    val definitions = mutableListOf<StreamProtocolCommandDefinition>()
    var index = 0
    var depth = 0
    var lineStart = true
    var inComment = false
    var inString = false
    var stringQuote = '\u0000'
    var escaped = false

    while (index < text.length) {
      val character = text[index]

      if (inComment) {
        if (character == '\n') {
          inComment = false
          lineStart = true
        }
        index += 1
        continue
      }

      if (inString) {
        when {
          escaped -> escaped = false
          character == '\\' -> escaped = true
          character == stringQuote -> {
            inString = false
            stringQuote = '\u0000'
          }
        }
        if (character == '\n') {
          lineStart = true
        } else if (!character.isWhitespace()) {
          lineStart = false
        }
        index += 1
        continue
      }

      if (character == '#') {
        inComment = true
        index += 1
        continue
      }

      if (character == '"' || character == '\'') {
        inString = true
        stringQuote = character
        lineStart = false
        index += 1
        continue
      }

      if (character == '\r') {
        index += 1
        continue
      }

      if (character == '\n') {
        lineStart = true
        index += 1
        continue
      }

      if (depth == 0 && lineStart) {
        if (character.isWhitespace()) {
          index += 1
          continue
        }

        val match = streamProtocolCommandDefinitionRegex.find(text, index)
        if (match != null && match.range.first == index) {
          val commandName = match.groups[1]?.value.orEmpty()
          val openBraceOffset = match.value.lastIndexOf('{')
            .takeIf { it >= 0 }
            ?.let { index + it }
          val closeBraceOffset = openBraceOffset?.let { findMatchingBrace(text, it) }
          if (commandName.isNotBlank() && openBraceOffset != null && closeBraceOffset != null) {
            definitions += StreamProtocolCommandDefinition(
              name = commandName,
              startOffset = index,
              endOffset = closeBraceOffset + 1,
              line = lineNumberAt(text, index),
              definitionText = text.substring(index, closeBraceOffset + 1).trim(),
            )
            index = closeBraceOffset + 1
            lineStart = false
            continue
          }
        }

        lineStart = false
      }

      when (character) {
        '{' -> depth += 1
        '}' -> depth = (depth - 1).coerceAtLeast(0)
      }
      if (!character.isWhitespace()) {
        lineStart = false
      }
      index += 1
    }

    return definitions
  }

  private fun findMatchingBrace(text: String, openBraceOffset: Int): Int? {
    var index = openBraceOffset
    var depth = 0
    var inComment = false
    var inString = false
    var stringQuote = '\u0000'
    var escaped = false

    while (index < text.length) {
      val character = text[index]

      if (inComment) {
        if (character == '\n') {
          inComment = false
        }
        index += 1
        continue
      }

      if (inString) {
        when {
          escaped -> escaped = false
          character == '\\' -> escaped = true
          character == stringQuote -> {
            inString = false
            stringQuote = '\u0000'
          }
        }
        index += 1
        continue
      }

      when (character) {
        '#' -> inComment = true
        '"', '\'' -> {
          inString = true
          stringQuote = character
        }
        '{' -> depth += 1
        '}' -> {
          depth -= 1
          if (depth == 0) {
            return index
          }
        }
      }
      index += 1
    }

    return null
  }

  private fun lineNumberAt(text: String, offset: Int): Int {
    return text.take(offset.coerceIn(0, text.length)).count { it == '\n' } + 1
  }

  private fun parseOptionalStreamProtocolArguments(text: String, commandEnd: Int): Int {
    var index = commandEnd
    while (index < text.length && text[index].isWhitespace()) {
      index += 1
    }
    if (index >= text.length || text[index] != '(') {
      return commandEnd
    }

    var depth = 0
    var inString = false
    var stringQuote = '\u0000'
    var escaped = false
    while (index < text.length) {
      val character = text[index]
      if (inString) {
        when {
          escaped -> escaped = false
          character == '\\' -> escaped = true
          character == stringQuote -> {
            inString = false
            stringQuote = '\u0000'
          }
        }
        index += 1
        continue
      }

      if (character == '"' || character == '\'') {
        inString = true
        stringQuote = character
        index += 1
        continue
      }

      if (character == '(') {
        depth += 1
      } else if (character == ')') {
        depth -= 1
        if (depth == 0) {
          return index + 1
        }
      }
      index += 1
    }

    return commandEnd
  }

  private fun parseStreamProtocolBus(
    text: String,
    startOffset: Int,
  ): ParsedStreamProtocolBus? {
    var index = startOffset
    while (index < text.length && text[index].isWhitespace()) {
      index += 1
    }
    if (index >= text.length) {
      return null
    }

    val busStart = index
    val busEnd = when {
      text[index] == '$' -> parseEpicsMacroToken(text, index)
      else -> {
        var tokenEnd = index
        while (tokenEnd < text.length && !text[tokenEnd].isWhitespace()) {
          tokenEnd += 1
        }
        tokenEnd
      }
    }
    if (busEnd <= busStart) {
      return null
    }

    return ParsedStreamProtocolBus(
      value = text.substring(busStart, busEnd),
      start = busStart,
      end = busEnd,
    )
  }

  private fun parseEpicsMacroToken(text: String, startOffset: Int): Int {
    if (startOffset + 1 >= text.length) {
      return startOffset + 1
    }

    return when (text[startOffset + 1]) {
      '(' -> parseDelimitedToken(text, startOffset + 1, '(', ')')
      '{' -> parseDelimitedToken(text, startOffset + 1, '{', '}')
      else -> {
        var index = startOffset + 1
        while (index < text.length && (text[index] == '_' || text[index].isLetterOrDigit())) {
          index += 1
        }
        index
      }
    }
  }

  private fun parseDelimitedToken(
    text: String,
    delimiterOffset: Int,
    openDelimiter: Char,
    closeDelimiter: Char,
  ): Int {
    var index = delimiterOffset
    var depth = 0
    var inString = false
    var stringQuote = '\u0000'
    var escaped = false

    while (index < text.length) {
      val character = text[index]
      if (inString) {
        when {
          escaped -> escaped = false
          character == '\\' -> escaped = true
          character == stringQuote -> {
            inString = false
            stringQuote = '\u0000'
          }
        }
        index += 1
        continue
      }

      if (character == '"' || character == '\'') {
        inString = true
        stringQuote = character
        index += 1
        continue
      }

      if (character == openDelimiter) {
        depth += 1
      } else if (character == closeDelimiter) {
        depth -= 1
        if (depth == 0) {
          return index + 1
        }
      }
      index += 1
    }

    return text.length
  }

  private fun isStreamProtocolCommandStart(character: Char): Boolean {
    return character == '_' || character.isLetter()
  }

  private fun isStreamProtocolCommandPart(character: Char): Boolean {
    return character == '_' || character == '-' || character.isLetterOrDigit()
  }

  private fun isStreamDeviceRecord(
    fieldDeclarations: List<EpicsRecordCompletionSupport.FieldDeclaration>,
  ): Boolean {
    val dtypField = fieldDeclarations.firstOrNull { it.fieldName == "DTYP" } ?: return false
    return dtypField.value.trim().equals("stream", ignoreCase = true)
  }

  private fun isLinkField(recordType: String, fieldName: String): Boolean {
    return EpicsRecordCompletionSupport.getFieldType(recordType, fieldName)?.contains("LINK") == true
  }

  private data class ParsedStreamProtocolBus(
    val value: String,
    val start: Int,
    val end: Int,
  )
}
