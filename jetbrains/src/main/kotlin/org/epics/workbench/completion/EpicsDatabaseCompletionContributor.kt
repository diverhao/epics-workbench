package org.epics.workbench.completion

import com.intellij.codeInsight.AutoPopupController
import com.intellij.codeInsight.completion.CodeCompletionHandlerBase
import com.intellij.codeInsight.completion.CompletionContributor
import com.intellij.codeInsight.completion.CompletionParameters
import com.intellij.codeInsight.completion.CompletionResultSet
import com.intellij.codeInsight.completion.CompletionType
import com.intellij.codeInsight.completion.InsertHandler
import com.intellij.codeInsight.lookup.LookupElement
import com.intellij.codeInsight.lookup.LookupElementBuilder
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.editor.Document
import com.intellij.openapi.util.TextRange
import com.intellij.psi.PsiDocumentManager
import com.intellij.psi.PsiElement
import org.epics.workbench.navigation.EpicsPathCompletionCandidate
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import java.nio.file.Files
import java.nio.file.Path

class EpicsDatabaseCompletionContributor : CompletionContributor() {
  override fun invokeAutoPopup(position: PsiElement, typeChar: Char): Boolean {
    val fileName = position.containingFile.name
    return when {
      isDatabaseFile(fileName) -> typeChar == '(' || typeChar == '"'
      isStartupFile(fileName) -> typeChar == '"' || typeChar.isLetterOrDigit() || typeChar == '_'
      else -> false
    }
  }

  override fun fillCompletionVariants(parameters: CompletionParameters, result: CompletionResultSet) {
    val file = parameters.originalFile
    if (!isDatabaseFile(file.name) && !isStartupFile(file.name)) {
      return
    }

    val editor = parameters.editor
    val document = editor.document
    val offset = editor.caretModel.offset
    val linePrefix = getLinePrefix(document, offset)

    if (isStartupFile(file.name)) {
      getStartupCommandContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        for (commandName in STARTUP_COMMANDS) {
          if (!matchesCompletionQuery(commandName, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(commandName)
              .withTypeText("EPICS startup command", true)
              .withInsertHandler(buildStartupCommandInsertHandler(commandName, context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getStartupLoadPathContext(offset, linePrefix)?.let { context ->
        val hostFile = file.virtualFile
        val resultSet = result.withPrefixMatcher(context.partial)
        val candidates = EpicsPathResolver.collectStartupPathCompletionCandidates(
          project = file.project,
          hostFile = hostFile,
          text = document.text,
          untilOffset = offset,
          rawPartial = context.partial,
          kind = context.pathKind,
        )
        for (candidate in candidates) {
          if (!matchesStartupPathQuery(candidate, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(candidate.insertPath)
              .withTypeText(candidate.detail, true)
              .withInsertHandler(buildStartupPathInsertHandler(candidate, context.startOffset, context.pathKind)),
          )
        }
        result.stopHere()
        return
      }

      return
    }

    getRecordTypeContext(offset, linePrefix)?.let { context ->
      val resultSet = result.withPrefixMatcher(context.partial)
      val hostFile = file.virtualFile ?: file.viewProvider.virtualFile
      for (recordType in EpicsRecordCompletionSupport.getRecordTypes(file.project, hostFile)) {
        if (!matchesCompletionQuery(recordType, context.partial)) {
          continue
        }
        resultSet.addElement(
          LookupElementBuilder.create(recordType)
            .withTypeText("EPICS record type", true)
            .withInsertHandler(buildRecordTypeInsertHandler(recordType, context.startOffset)),
        )
      }
      result.stopHere()
      return
    }

    getRecordNameContext(offset, linePrefix)?.let { context ->
      val names = linkedSetOf<String>()
      for (declaration in EpicsRecordCompletionSupport.extractRecordDeclarations(document.text)) {
        if (declaration.name.isNotBlank() && declaration.name != context.partial) {
          names += declaration.name
        }
      }
      val resultSet = result.withPrefixMatcher(context.partial)
      for (name in names.sortedWith(String.CASE_INSENSITIVE_ORDER)) {
        if (!matchesCompletionQuery(name, context.partial)) {
          continue
        }
        resultSet.addElement(
          LookupElementBuilder.create(name)
            .withTypeText("Existing record name in this file", true)
            .withInsertHandler(buildSimpleReplacementInsertHandler(name, context.startOffset)),
        )
      }
      result.stopHere()
      return
    }

    getFieldNameContext(offset, linePrefix)?.let { context ->
      val recordType = EpicsRecordCompletionSupport.findEnclosingRecordType(document.text, offset)
      val fieldNames = EpicsRecordCompletionSupport.getAvailableFieldNamesForRecordInstance(
        documentText = document.text,
        offset = offset,
        recordType = recordType,
      )
      val resultSet = result.withPrefixMatcher(context.partial)

      for (fieldName in fieldNames) {
        if (!matchesCompletionQuery(fieldName, context.partial)) {
          continue
        }
        resultSet.addElement(
          LookupElementBuilder.create(fieldName)
            .withTypeText(buildFieldTypeText(recordType, fieldName), true)
            .withInsertHandler(buildFieldNameInsertHandler(recordType, fieldName, context.startOffset)),
        )
      }
      result.stopHere()
      return
    }

    getFieldValueContext(offset, linePrefix)?.let { context ->
      val recordType = EpicsRecordCompletionSupport.findEnclosingRecordType(document.text, offset)
      if (recordType.isNullOrBlank()) {
        return
      }
      if (EpicsRecordCompletionSupport.getFieldType(recordType, context.fieldName) != "DBF_MENU") {
        return
      }

      val choices = EpicsRecordCompletionSupport.getMenuFieldChoices(recordType, context.fieldName)
      if (choices.isEmpty()) {
        return
      }

      val resultSet = result.withPrefixMatcher(context.partial)
      for (choice in choices) {
        if (!matchesCompletionQuery(choice, context.partial)) {
          continue
        }
        resultSet.addElement(
          LookupElementBuilder.create(choice)
            .withTypeText("${context.fieldName} menu choice", true)
            .withInsertHandler(buildSimpleReplacementInsertHandler(choice, context.startOffset)),
        )
      }
      result.stopHere()
    }
  }

  companion object {
    internal fun maybeScheduleAutoPopupForTypedChar(
      file: com.intellij.psi.PsiFile,
      editor: com.intellij.openapi.editor.Editor,
      project: com.intellij.openapi.project.Project,
      charTyped: Char,
    ) {
      if (!isDatabaseFile(file.name) && !isStartupFile(file.name)) {
        return
      }

      val offset = editor.caretModel.offset
      val linePrefix = getLinePrefix(editor.document, offset)
      val shouldPopup = when (charTyped) {
        '(' -> getRecordTypeContext(offset, linePrefix) != null ||
          getFieldNameContext(offset, linePrefix) != null
        '"' -> getRecordNameContext(offset, linePrefix) != null ||
          shouldPopupForMenuFieldValue(editor.document, offset, linePrefix) ||
          getStartupLoadPathContext(offset, linePrefix) != null
        in 'A'..'Z', in 'a'..'z', in '0'..'9', '_' -> getStartupCommandContext(offset, linePrefix) != null
        else -> false
      }

      if (shouldPopup) {
        requestCompletionPopup(project, editor)
      }
    }

    private fun requestCompletionPopup(
      project: com.intellij.openapi.project.Project,
      editor: com.intellij.openapi.editor.Editor,
    ) {
      ApplicationManager.getApplication().invokeLater {
        if (editor.isDisposed) {
          return@invokeLater
        }
        PsiDocumentManager.getInstance(project).commitDocument(editor.document)
        AutoPopupController.getInstance(project).scheduleAutoPopup(editor)
        CodeCompletionHandlerBase(CompletionType.BASIC).invokeCompletion(project, editor)
      }
    }

    private fun buildRecordTypeInsertHandler(
      recordType: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      val document = context.document

      document.replaceString(replacementStartOffset, context.tailOffset, recordType)
      val tailStart = replacementStartOffset + recordType.length
      val tailEnd = getRecordTailInsertionEnd(document, tailStart)
      val baseIndent = getLineIndent(document, tailStart)
      val indentUnit = EpicsRecordCompletionSupport.getIndentUnit(context.file)
      val tailText = buildRecordTemplateTail(recordType, indentUnit, baseIndent)
      document.replaceString(tailStart, tailEnd, tailText)
      context.commitDocument()

      val nameOffset = tailStart + 3
      context.editor.caretModel.moveToOffset(nameOffset)
      context.editor.selectionModel.removeSelection()

      requestCompletionPopup(context.project, context.editor)
    }

    private fun buildSimpleReplacementInsertHandler(
      value: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      context.document.replaceString(replacementStartOffset, context.tailOffset, value)
      context.commitDocument()
      context.editor.caretModel.moveToOffset(replacementStartOffset + value.length)
      context.editor.selectionModel.removeSelection()
    }

    private fun buildStartupPathInsertHandler(
      candidate: EpicsPathCompletionCandidate,
      replacementStartOffset: Int,
      pathKind: EpicsPathKind,
    ): InsertHandler<LookupElement> {
      if (pathKind != EpicsPathKind.DATABASE || candidate.isDirectory || candidate.absolutePath == null) {
        return buildSimpleReplacementInsertHandler(candidate.insertPath, replacementStartOffset)
      }

      return InsertHandler { context, _ ->
        context.setAddCompletionChar(false)
        val document = context.document

        document.replaceString(replacementStartOffset, context.tailOffset, candidate.insertPath)
        val tailStart = replacementStartOffset + candidate.insertPath.length
        val macroNames = extractMacroNamesFromDatabaseFile(candidate.absolutePath)
        val tailEnd = getDbLoadRecordsTailInsertionEnd(document, tailStart)
        val tailText = buildDbLoadRecordsTail(macroNames)
        document.replaceString(tailStart, tailEnd, tailText)
        context.commitDocument()

        if (macroNames.isEmpty()) {
          context.editor.caretModel.moveToOffset(tailStart + tailText.length)
          context.editor.selectionModel.removeSelection()
          return@InsertHandler
        }

        val firstValueOffset = tailStart + 5 + macroNames.first().length
        context.editor.caretModel.moveToOffset(firstValueOffset)
        context.editor.selectionModel.removeSelection()
      }
    }

    private fun buildStartupCommandInsertHandler(
      commandName: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      val document = context.document

      document.replaceString(replacementStartOffset, context.tailOffset, commandName)
      val tailStart = replacementStartOffset + commandName.length
      val tailEnd = getIdentifierTailInsertionEnd(document, tailStart)
      document.replaceString(tailStart, tailEnd, "(\"\")")
      context.commitDocument()

      val pathOffset = tailStart + 2
      context.editor.caretModel.moveToOffset(pathOffset)
      context.editor.selectionModel.removeSelection()
      requestCompletionPopup(context.project, context.editor)
    }

    private fun buildFieldNameInsertHandler(
      recordType: String?,
      fieldName: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      val document = context.document

      document.replaceString(replacementStartOffset, context.tailOffset, fieldName)
      val tailStart = replacementStartOffset + fieldName.length
      val tailEnd = getFieldTailInsertionEnd(document, tailStart)
      val defaultValue = EpicsRecordCompletionSupport.getDefaultFieldValue(recordType.orEmpty(), fieldName)
      val tailText = buildFieldTail(defaultValue)
      document.replaceString(tailStart, tailEnd, tailText)
      context.commitDocument()

      val valueStart = tailStart + 3
      val valueEnd = valueStart + defaultValue.length
      context.editor.caretModel.moveToOffset(valueEnd)
      context.editor.selectionModel.setSelection(valueStart, valueEnd)
    }

    private fun buildRecordTemplateTail(
      recordType: String,
      indentUnit: String,
      baseIndent: String,
    ): String {
      val fields = EpicsRecordCompletionSupport.getTemplateFields(recordType)
      val builder = StringBuilder()
      builder.append(", \"\") {")
      for (fieldName in fields) {
        builder.append('\n')
          .append(baseIndent)
          .append(indentUnit)
          .append("field(")
          .append(fieldName)
          .append(", \"")
          .append(escapeEpicsString(EpicsRecordCompletionSupport.getDefaultFieldValue(recordType, fieldName)))
          .append("\")")
      }
      builder.append('\n').append(baseIndent).append('}')
      return builder.toString()
    }

    private fun buildFieldTail(defaultValue: String): String {
      return ", \"${escapeEpicsString(defaultValue)}\")"
    }

    private fun buildDbLoadRecordsTail(macroNames: List<String>): String {
      if (macroNames.isEmpty()) {
        return "\")"
      }
      val assignments = macroNames.joinToString(",") { "${it}=" }
      return "\", \"$assignments\")"
    }

    private fun getRecordTailInsertionEnd(document: Document, startOffset: Int): Int {
      val lineNumber = document.getLineNumber(startOffset.coerceAtMost(document.textLength))
      val lineEnd = document.getLineEndOffset(lineNumber)
      var offset = startOffset
      val chars = document.charsSequence

      while (offset < lineEnd && chars[offset].isLetterOrDigitOrUnderscore()) {
        offset += 1
      }
      if (offset < lineEnd && chars[offset] == ')') {
        offset += 1
      }
      return offset
    }

    private fun getIdentifierTailInsertionEnd(document: Document, startOffset: Int): Int {
      val lineNumber = document.getLineNumber(startOffset.coerceAtMost(document.textLength))
      val lineEnd = document.getLineEndOffset(lineNumber)
      var offset = startOffset
      val chars = document.charsSequence

      while (offset < lineEnd && chars[offset].isLetterOrDigitOrUnderscore()) {
        offset += 1
      }
      return offset
    }

    private fun getFieldTailInsertionEnd(document: Document, startOffset: Int): Int {
      val lineNumber = document.getLineNumber(startOffset.coerceAtMost(document.textLength))
      val lineEnd = document.getLineEndOffset(lineNumber)
      var offset = startOffset
      val chars = document.charsSequence

      while (offset < lineEnd && chars[offset].isLetterOrDigitOrUnderscore()) {
        offset += 1
      }
      if (offset < lineEnd && chars[offset] == '"') {
        offset += 1
      }
      if (offset < lineEnd && chars[offset] == ')') {
        offset += 1
      }
      return offset
    }

    private fun getDbLoadRecordsTailInsertionEnd(document: Document, startOffset: Int): Int {
      val lineNumber = document.getLineNumber(startOffset.coerceAtMost(document.textLength))
      val lineEnd = document.getLineEndOffset(lineNumber)
      val chars = document.charsSequence
      var offset = startOffset

      if (offset < lineEnd && chars[offset] == '"') {
        offset += 1
      }
      while (offset < lineEnd && chars[offset].isWhitespace()) {
        offset += 1
      }
      if (offset < lineEnd && chars[offset] == ',') {
        offset += 1
        while (offset < lineEnd && chars[offset].isWhitespace()) {
          offset += 1
        }
        if (offset < lineEnd && chars[offset] == '"') {
          offset += 1
          var escaped = false
          while (offset < lineEnd) {
            val character = chars[offset]
            offset += 1
            if (escaped) {
              escaped = false
              continue
            }
            if (character == '\\') {
              escaped = true
              continue
            }
            if (character == '"') {
              break
            }
          }
        }
        while (offset < lineEnd && chars[offset].isWhitespace()) {
          offset += 1
        }
      }
      if (offset < lineEnd && chars[offset] == ')') {
        offset += 1
      }
      return offset
    }

    private fun getLinePrefix(document: Document, offset: Int): String {
      val lineNumber = document.getLineNumber(offset.coerceAtMost(document.textLength))
      val lineStart = document.getLineStartOffset(lineNumber)
      return document.getText(TextRange(lineStart, offset.coerceAtMost(document.textLength)))
    }

    private fun getLineIndent(document: Document, offset: Int): String {
      val lineNumber = document.getLineNumber(offset.coerceAtMost(document.textLength))
      val lineStart = document.getLineStartOffset(lineNumber)
      val lineEnd = document.getLineEndOffset(lineNumber)
      val lineText = document.getText(TextRange(lineStart, lineEnd))
      return lineText.takeWhile { it == ' ' || it == '\t' }
    }

    private fun getRecordTypeContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = RECORD_TYPE_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getRecordNameContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = RECORD_NAME_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getFieldNameContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = FIELD_NAME_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getFieldValueContext(offset: Int, linePrefix: String): FieldValueCompletionContext? {
      val match = FIELD_VALUE_CONTEXT_REGEX.find(linePrefix) ?: return null
      val fieldName = match.groups[1]?.value ?: match.groups[2]?.value ?: return null
      val partial = match.groups[3]?.value.orEmpty()
      return FieldValueCompletionContext(
        fieldName = fieldName,
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getStartupCommandContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = STARTUP_COMMAND_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getStartupLoadPathContext(offset: Int, linePrefix: String): StartupPathCompletionContext? {
      DB_LOAD_RECORDS_PATH_CONTEXT_REGEX.find(linePrefix)?.let { match ->
        val partial = match.groups[1]?.value.orEmpty()
        return StartupPathCompletionContext(
          partial = partial,
          startOffset = offset - partial.length,
          pathKind = EpicsPathKind.DATABASE,
        )
      }

      DB_LOAD_TEMPLATE_PATH_CONTEXT_REGEX.find(linePrefix)?.let { match ->
        val partial = match.groups[1]?.value.orEmpty()
        return StartupPathCompletionContext(
          partial = partial,
          startOffset = offset - partial.length,
          pathKind = EpicsPathKind.SUBSTITUTIONS,
        )
      }

      return null
    }

    private fun shouldPopupForMenuFieldValue(
      document: Document,
      offset: Int,
      linePrefix: String,
    ): Boolean {
      val context = getFieldValueContext(offset, linePrefix) ?: return false
      val recordType = EpicsRecordCompletionSupport.findEnclosingRecordType(document.text, offset) ?: return false
      if (EpicsRecordCompletionSupport.getFieldType(recordType, context.fieldName) != "DBF_MENU") {
        return false
      }
      return EpicsRecordCompletionSupport.getMenuFieldChoices(recordType, context.fieldName).isNotEmpty()
    }

    private fun buildFieldTypeText(recordType: String?, fieldName: String): String {
      val dbfType = recordType?.let { EpicsRecordCompletionSupport.getFieldType(it, fieldName) }
      return if (recordType != null && dbfType != null) {
        "$fieldName for $recordType ($dbfType)"
      } else if (recordType != null) {
        "Field for $recordType"
      } else {
        "EPICS field"
      }
    }

    private fun matchesStartupPathQuery(
      candidate: EpicsPathCompletionCandidate,
      partial: String,
    ): Boolean {
      if (partial.isBlank()) {
        return true
      }
      val normalizedPartial = partial.replace('\\', '/')
      return candidate.insertPath.startsWith(normalizedPartial, ignoreCase = true) ||
        candidate.insertPath.substringAfterLast('/').startsWith(
          normalizedPartial.substringAfterLast('/'),
          ignoreCase = true,
        )
    }

    private fun matchesCompletionQuery(label: String, partial: String): Boolean {
      if (partial.isBlank()) {
        return true
      }
      return label.startsWith(partial, ignoreCase = true)
    }

    private fun isDatabaseFile(fileName: String): Boolean {
      val extension = fileName.substringAfterLast('.', "").lowercase()
      return extension in setOf("db", "vdb", "template")
    }

    private fun isStartupFile(fileName: String): Boolean {
      val extension = fileName.substringAfterLast('.', "").lowercase()
      return extension in setOf("cmd", "iocsh") || fileName == "st.cmd"
    }

    private fun extractMacroNamesFromDatabaseFile(path: Path): List<String> {
      val text = runCatching { Files.readString(path) }.getOrNull() ?: return emptyList()
      return extractMacroNames(maskDatabaseComments(text))
    }

    private fun maskDatabaseComments(text: String): String {
      val sanitized = StringBuilder(text.length)
      var inString = false
      var escaped = false
      var index = 0

      while (index < text.length) {
        val character = text[index]
        if (inString) {
          sanitized.append(character)
          when {
            escaped -> escaped = false
            character == '\\' -> escaped = true
            character == '"' -> inString = false
          }
          index += 1
          continue
        }

        if (character == '"') {
          inString = true
          sanitized.append(character)
          index += 1
          continue
        }

        if (character == '#') {
          while (index < text.length && text[index] != '\n') {
            sanitized.append(' ')
            index += 1
          }
          if (index < text.length && text[index] == '\n') {
            sanitized.append('\n')
            index += 1
          }
          continue
        }

        sanitized.append(character)
        index += 1
      }

      return sanitized.toString()
    }

    private fun extractMacroNames(text: String): List<String> {
      val names = linkedSetOf<String>()
      EPICS_MACRO_REFERENCE_REGEX.findAll(text).forEach { match ->
        val name = match.groups[1]?.value ?: match.groups[2]?.value
        if (!name.isNullOrBlank()) {
          names += name
        }
      }
      return names.toList()
    }

    private fun Char.isLetterOrDigitOrUnderscore(): Boolean {
      return isLetterOrDigit() || this == '_'
    }

    private fun escapeEpicsString(value: String): String {
      return value.replace("\\", "\\\\").replace("\"", "\\\"")
    }

    private data class CompletionContext(
      val partial: String,
      val startOffset: Int,
    )

    private data class FieldValueCompletionContext(
      val fieldName: String,
      val partial: String,
      val startOffset: Int,
    )

    private data class StartupPathCompletionContext(
      val partial: String,
      val startOffset: Int,
      val pathKind: EpicsPathKind,
    )

    private val RECORD_TYPE_CONTEXT_REGEX = Regex("""record\(\s*([A-Za-z0-9_]*)$""")
    private val RECORD_NAME_CONTEXT_REGEX = Regex("""record\(\s*[A-Za-z0-9_]+\s*,\s*"([^"\n]*)$""")
    private val FIELD_NAME_CONTEXT_REGEX = Regex("""field\(\s*(?:"?([A-Za-z0-9_]*))$""")
    private val FIELD_VALUE_CONTEXT_REGEX =
      Regex("""field\(\s*(?:"([A-Za-z0-9_]+)"|([A-Za-z0-9_]+))\s*,\s*"([^"\n]*)$""")
    private val STARTUP_COMMAND_CONTEXT_REGEX = Regex("""^\s*([A-Za-z0-9_]*)$""")
    private val DB_LOAD_RECORDS_PATH_CONTEXT_REGEX = Regex("""dbLoadRecords\(\s*"([^"\n]*)$""")
    private val DB_LOAD_TEMPLATE_PATH_CONTEXT_REGEX = Regex("""dbLoadTemplate\(\s*"([^"\n]*)$""")
    private val EPICS_MACRO_REFERENCE_REGEX =
      Regex("""\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}""")
    private val STARTUP_COMMANDS = listOf("dbLoadRecords", "dbLoadTemplate")
  }
}
