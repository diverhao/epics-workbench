package org.epics.workbench.inspections

import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsPathResolver
import java.util.ArrayDeque
import java.util.concurrent.ConcurrentHashMap
import kotlin.io.path.pathString

internal object EpicsDatabaseValueValidator {
  internal enum class ValidationSeverity {
    ERROR,
    WARNING,
  }

  internal data class ValidationIssue(
    val startOffset: Int,
    val endOffset: Int,
    val message: String,
    val code: String? = null,
    val severity: ValidationSeverity = ValidationSeverity.ERROR,
  )

  private data class StreamProtocolInvocation(
    val protocolPath: String,
    val commandName: String?,
    val commandStart: Int,
    val commandEnd: Int,
  )

  private data class StreamProtocolCommandCacheEntry(
    val cacheTag: String,
    val commandNames: Set<String>,
  )

  private val streamProtocolCommandCache = ConcurrentHashMap<String, StreamProtocolCommandCacheEntry>()
  private val streamProtocolInvocationRegex =
    Regex("""^\s*@([^\s"'`]+)(?:\s+([A-Za-z_][A-Za-z0-9_-]*))?""")
  private val streamProtocolCommandDefinitionRegex =
    Regex("""^([A-Za-z_][A-Za-z0-9_-]*)\s*\{""")

  fun collectIssues(text: String): List<ValidationIssue> = collectIssues(null, null, text)

  fun collectIssues(project: Project?, hostFile: VirtualFile?, text: String): List<ValidationIssue> {
    val issues = mutableListOf<ValidationIssue>()
    issues += collectUnmatchedDelimiterIssues(text)

    val declarations = EpicsRecordCompletionSupport.extractRecordDeclarations(text)
    issues += collectDuplicateRecordIssues(declarations)

    for (recordDeclaration in declarations) {
      val fieldDeclarations = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, recordDeclaration)
      issues += collectDuplicateFieldIssues(recordDeclaration, fieldDeclarations)

      val allowedFields = EpicsRecordCompletionSupport.getDeclaredFieldNamesForRecordType(recordDeclaration.recordType)
      if (!allowedFields.isNullOrEmpty()) {
        for (fieldDeclaration in fieldDeclarations) {
          if (fieldDeclaration.fieldName !in allowedFields) {
            issues += ValidationIssue(
              startOffset = fieldDeclaration.fieldNameStart,
              endOffset = fieldDeclaration.fieldNameEnd,
              message = "Field \"${fieldDeclaration.fieldName}\" is not valid for record type \"${recordDeclaration.recordType}\".",
              code = "epics.database.invalidFieldName",
            )
          }
        }
      }

      for (fieldDeclaration in fieldDeclarations) {
        val dbfType = EpicsRecordCompletionSupport.getFieldType(
          recordDeclaration.recordType,
          fieldDeclaration.fieldName,
        ) ?: continue

        if (EpicsRecordCompletionSupport.isNumericFieldType(dbfType)) {
          if (EpicsRecordCompletionSupport.isSkippableNumericFieldValue(fieldDeclaration.value)) {
            continue
          }
          if (EpicsRecordCompletionSupport.isValidNumericFieldValue(fieldDeclaration.value, dbfType)) {
            continue
          }

          issues += ValidationIssue(
            startOffset = fieldDeclaration.valueStart,
            endOffset = fieldDeclaration.valueEnd,
            message = "Field \"${fieldDeclaration.fieldName}\" expects a $dbfType numeric value.",
          )
          continue
        }

        if (dbfType != "DBF_MENU") {
          continue
        }
        if (EpicsRecordCompletionSupport.containsEpicsMacroReference(fieldDeclaration.value)) {
          continue
        }

        val allowedChoices = EpicsRecordCompletionSupport.getMenuFieldChoices(
          recordDeclaration.recordType,
          fieldDeclaration.fieldName,
        )
        if (allowedChoices.isEmpty() || allowedChoices.contains(fieldDeclaration.value)) {
          continue
        }

        issues += ValidationIssue(
          startOffset = fieldDeclaration.valueStart,
          endOffset = fieldDeclaration.valueEnd,
          message = "Field \"${fieldDeclaration.fieldName}\" must be one of the menu choices for \"${recordDeclaration.recordType}\".",
          code = "epics.database.invalidMenuFieldValue",
        )
      }
    }

    if (project != null && hostFile != null) {
      issues += collectInvalidStreamProtocolCommandIssues(project, hostFile, declarations, text)
    }

    return issues
  }

  private fun collectInvalidStreamProtocolCommandIssues(
    project: Project,
    hostFile: VirtualFile,
    declarations: List<EpicsRecordCompletionSupport.RecordDeclaration>,
    text: String,
  ): List<ValidationIssue> {
    val issues = mutableListOf<ValidationIssue>()

    for (recordDeclaration in declarations) {
      val fieldDeclarations = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, recordDeclaration)
      if (!isStreamDeviceRecord(fieldDeclarations)) {
        continue
      }

      for (fieldDeclaration in fieldDeclarations) {
        if (!isLinkField(recordDeclaration.recordType, fieldDeclaration.fieldName)) {
          continue
        }

        val invocation = extractStreamProtocolInvocation(fieldDeclaration) ?: continue
        val commandName = invocation.commandName ?: continue
        val protocolFiles = EpicsPathResolver.resolveStreamProtocolPaths(project, hostFile, invocation.protocolPath)
        if (protocolFiles.isEmpty()) {
          continue
        }

        if (protocolFiles.any { protocolFile -> getStreamProtocolCommandNames(protocolFile).contains(commandName) }) {
          continue
        }

        val protocolLabel = invocation.protocolPath
          .substringAfterLast('/')
          .substringAfterLast('\\')
          .ifBlank { invocation.protocolPath }
        issues += ValidationIssue(
          startOffset = invocation.commandStart,
          endOffset = invocation.commandEnd,
          message = if (protocolFiles.size == 1) {
            "StreamDevice command \"$commandName\" was not found in protocol file \"$protocolLabel\"."
          } else {
            "StreamDevice command \"$commandName\" was not found in any resolved \"$protocolLabel\" protocol file."
          },
          code = "epics.database.unknownStreamProtocolCommand",
        )
      }
    }

    return issues
  }

  private fun extractStreamProtocolInvocation(
    fieldDeclaration: EpicsRecordCompletionSupport.FieldDeclaration,
  ): StreamProtocolInvocation? {
    val match = streamProtocolInvocationRegex.find(fieldDeclaration.value) ?: return null
    val protocolPath = match.groups[1]?.value.orEmpty()
    if (protocolPath.isBlank() || EpicsRecordCompletionSupport.containsEpicsMacroReference(protocolPath)) {
      return null
    }

    val commandGroup = match.groups[2]
    val commandName = commandGroup?.value
    val commandStart = commandGroup?.range?.first?.let { fieldDeclaration.valueStart + it } ?: fieldDeclaration.valueStart
    val commandEnd = commandGroup?.range?.last?.plus(1)?.let { fieldDeclaration.valueStart + it } ?: fieldDeclaration.valueStart
    return StreamProtocolInvocation(
      protocolPath = protocolPath,
      commandName = commandName,
      commandStart = commandStart,
      commandEnd = commandEnd,
    )
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

  private fun getStreamProtocolCommandNames(protocolPath: java.nio.file.Path): Set<String> {
    val normalizedPath = protocolPath.normalize().pathString
    val virtualFile = LocalFileSystem.getInstance().findFileByNioFile(protocolPath)

    val cacheTag: String
    val text: String
    if (virtualFile != null) {
      val document = FileDocumentManager.getInstance().getCachedDocument(virtualFile)
      if (document != null) {
        cacheTag = "open:${document.modificationStamp}"
        text = document.text
      } else {
        cacheTag = "vfs:${virtualFile.modificationStamp}"
        text = runCatching { virtualFile.inputStream.bufferedReader().use { it.readText() } }.getOrElse { return emptySet() }
      }
    } else {
      val size = runCatching { java.nio.file.Files.size(protocolPath) }.getOrNull() ?: return emptySet()
      val modified = runCatching { java.nio.file.Files.getLastModifiedTime(protocolPath).toMillis() }.getOrNull()
        ?: return emptySet()
      cacheTag = "fs:$size:$modified"
      text = runCatching { java.nio.file.Files.readString(protocolPath) }.getOrElse { return emptySet() }
    }

    streamProtocolCommandCache[normalizedPath]?.takeIf { it.cacheTag == cacheTag }?.let { entry ->
      return entry.commandNames
    }

    val commandNames = extractStreamProtocolCommandNames(text)
    streamProtocolCommandCache[normalizedPath] = StreamProtocolCommandCacheEntry(
      cacheTag = cacheTag,
      commandNames = commandNames,
    )
    return commandNames
  }

  private fun extractStreamProtocolCommandNames(text: String): Set<String> {
    val commandNames = linkedSetOf<String>()
    var depth = 0
    var index = 0
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

        val match = streamProtocolCommandDefinitionRegex.find(text.substring(index))
        if (match != null) {
          commandNames += match.groups[1]?.value.orEmpty()
          depth += 1
          lineStart = false
          index += match.value.length
          continue
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

    return commandNames
  }

  private fun collectDuplicateRecordIssues(
    declarations: List<EpicsRecordCompletionSupport.RecordDeclaration>,
  ): List<ValidationIssue> {
    val declarationsByName = linkedMapOf<String, MutableList<EpicsRecordCompletionSupport.RecordDeclaration>>()
    declarations.forEach { declaration ->
      declarationsByName.getOrPut(declaration.name) { mutableListOf() } += declaration
    }

    val issues = mutableListOf<ValidationIssue>()
    for ((recordName, duplicates) in declarationsByName) {
      if (recordName.isBlank() || duplicates.size < 2) {
        continue
      }
      duplicates.forEach { declaration ->
        issues += ValidationIssue(
          startOffset = declaration.nameStart,
          endOffset = declaration.nameEnd,
          message = "Duplicate record name \"$recordName\" in this file.",
          code = "epics.database.duplicateRecordName",
        )
      }
    }
    return issues
  }

  private fun collectDuplicateFieldIssues(
    recordDeclaration: EpicsRecordCompletionSupport.RecordDeclaration,
    fieldDeclarations: List<EpicsRecordCompletionSupport.FieldDeclaration>,
  ): List<ValidationIssue> {
    val fieldsByName = linkedMapOf<String, MutableList<EpicsRecordCompletionSupport.FieldDeclaration>>()
    fieldDeclarations.forEach { declaration ->
      fieldsByName.getOrPut(declaration.fieldName) { mutableListOf() } += declaration
    }

    val issues = mutableListOf<ValidationIssue>()
    for ((fieldName, duplicates) in fieldsByName) {
      if (duplicates.size < 2) {
        continue
      }
      duplicates.forEach { declaration ->
        issues += ValidationIssue(
          startOffset = declaration.fieldNameStart,
          endOffset = declaration.fieldNameEnd,
          message = "Duplicate field \"$fieldName\" in record \"${recordDeclaration.name}\".",
        )
      }
    }
    return issues
  }

  private fun collectUnmatchedDelimiterIssues(text: String): List<ValidationIssue> {
    data class Delimiter(val character: Char, val index: Int)

    val issues = mutableListOf<ValidationIssue>()
    val delimiterStack = ArrayDeque<Delimiter>()
    var inString = false
    var escaped = false
    var inComment = false

    for ((index, character) in text.withIndex()) {
      if (inComment) {
        if (character == '\n') {
          inComment = false
        }
        continue
      }

      if (inString) {
        when {
          escaped -> escaped = false
          character == '\\' -> escaped = true
          character == '"' -> inString = false
        }
        continue
      }

      when (character) {
        '#' -> {
          inComment = true
        }

        '"' -> {
          inString = true
        }

        '(', '{' -> {
          delimiterStack.addLast(Delimiter(character, index))
        }

        ')', '}' -> {
          val expectedOpening = if (character == ')') '(' else '{'
          val lastOpening = if (delimiterStack.isEmpty()) null else delimiterStack.removeLast()
          if (lastOpening == null || lastOpening.character != expectedOpening) {
            if (lastOpening != null) {
              delimiterStack.addLast(lastOpening)
            }
            issues += ValidationIssue(
              startOffset = index,
              endOffset = index + 1,
              message = "Unmatched \"$character\".",
            )
          }
        }
      }
    }

    while (delimiterStack.isNotEmpty()) {
      val unmatchedOpening = delimiterStack.removeLast()
      val expectedClosing = if (unmatchedOpening.character == '(') ')' else '}'
      issues += ValidationIssue(
        startOffset = unmatchedOpening.index,
        endOffset = unmatchedOpening.index + 1,
        message = "Unmatched \"${unmatchedOpening.character}\"; missing \"$expectedClosing\".",
      )
    }

    return issues
  }
}
