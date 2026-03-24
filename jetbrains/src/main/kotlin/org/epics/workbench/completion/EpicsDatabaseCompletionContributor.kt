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
import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.util.TextRange
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiDocumentManager
import com.intellij.psi.PsiElement
import org.epics.workbench.navigation.EpicsPathCompletionCandidate
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import org.epics.workbench.navigation.EpicsRecordResolver
import java.nio.file.Files
import java.nio.file.Path
import java.util.concurrent.ConcurrentHashMap
import kotlin.io.path.extension

class EpicsDatabaseCompletionContributor : CompletionContributor() {
  override fun invokeAutoPopup(position: PsiElement, typeChar: Char): Boolean {
    val fileName = position.containingFile.name
    return when {
      isDatabaseFile(fileName) -> typeChar == '(' || typeChar == '"'
      isStartupFile(fileName) -> typeChar == '"' || typeChar.isLetterOrDigit() || typeChar == '_'
      isDbdFile(fileName) -> typeChar == '"' || typeChar == '(' || typeChar == ',' || typeChar.isLetterOrDigit() || typeChar == '_'
      else -> false
    }
  }

  override fun fillCompletionVariants(parameters: CompletionParameters, result: CompletionResultSet) {
    val file = parameters.originalFile
    if (!isDatabaseFile(file.name) && !isStartupFile(file.name) && !isDbdFile(file.name)) {
      return
    }

    val editor = parameters.editor
    val document = editor.document
    val offset = editor.caretModel.offset
    val linePrefix = getLinePrefix(document, offset)

    if (isDbdFile(file.name)) {
      fillDbdCompletionVariants(file.project, file.virtualFile ?: file.viewProvider.virtualFile, document, offset, linePrefix, result)
      return
    }

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

      getStartupLoadedRecordNameContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        val hostFile = file.virtualFile
        val names = EpicsRecordResolver.collectStartupLoadedRecordNames(
          project = file.project,
          hostFile = hostFile,
          text = document.text,
          untilOffset = offset,
        )
        for (name in names) {
          if (!matchesCompletionQuery(name, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(name, name)
              .withTypeText("Record loaded by this startup script", true)
              .withInsertHandler(buildSimpleReplacementInsertHandler(name, context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getStartupLoadPathContext(offset, linePrefix)?.let { context ->
        val hostFile = file.virtualFile
        val resultSet = result.withPrefixMatcher(context.partial)
        val candidates = if (context.pathKind == EpicsPathKind.DATABASE) {
          EpicsPathResolver.collectStartupDbLoadRecordsCompletionCandidates(
            project = file.project,
            hostFile = hostFile,
            text = document.text,
            untilOffset = offset,
            rawPartial = context.partial,
          )
        } else {
          EpicsPathResolver.collectStartupPathCompletionCandidates(
            project = file.project,
            hostFile = hostFile,
            text = document.text,
            untilOffset = offset,
            rawPartial = context.partial,
            kind = context.pathKind,
          )
        }
        for (candidate in candidates) {
          if (!matchesStartupPathQuery(candidate, context.partial)) {
            continue
          }
          val lookupLabel = if (context.pathKind == EpicsPathKind.DATABASE) {
            candidate.insertPath.substringAfterLast('/')
          } else {
            candidate.insertPath
          }
          resultSet.addElement(
            LookupElementBuilder.create(lookupLabel)
              .withLookupString(candidate.insertPath)
              .withTypeText(candidate.detail, true)
              .withInsertHandler(buildStartupPathInsertHandler(candidate, context.startOffset, context.pathKind)),
          )
        }
        result.stopHere()
        return
      }

      getStartupLoadMacroTailContext(offset, linePrefix)?.let { context ->
        val macroNames = extractMacroNamesFromStartupDatabaseFile(
          file.project,
          file.virtualFile,
          document.text,
          offset,
          context.path,
        )
        if (macroNames.isNotEmpty()) {
          result.addElement(
            LookupElementBuilder.create(buildDbLoadRecordsAssignmentLabel(macroNames))
              .withTypeText("Macros used by ${context.path.substringAfterLast('/')}", true)
              .withInsertHandler(buildStartupLoadMacroTailInsertHandler(macroNames, context.startOffset)),
          )
          result.stopHere()
        }
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
      if (!isDatabaseFile(file.name) && !isStartupFile(file.name) && !isDbdFile(file.name)) {
        return
      }

      val offset = editor.caretModel.offset
      val linePrefix = getLinePrefix(editor.document, offset)
      val shouldPopup = when (charTyped) {
        '(' -> getRecordTypeContext(offset, linePrefix) != null ||
          getFieldNameContext(offset, linePrefix) != null ||
          isDbdFile(file.name) && (
            getDbdDeviceRecordTypeContext(offset, linePrefix) != null ||
              getDbdKeywordContext(offset, linePrefix) != null ||
              getDbdSimpleNameContext(offset, linePrefix, "driver") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "registrar") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "function") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "variable") != null
            )
        '"' -> getRecordNameContext(offset, linePrefix) != null ||
          shouldPopupForMenuFieldValue(editor.document, offset, linePrefix) ||
          getStartupLoadPathContext(offset, linePrefix) != null ||
          getStartupLoadedRecordNameContext(offset, linePrefix) != null ||
          getStartupLoadMacroTailContext(offset, linePrefix) != null ||
          isDbdFile(file.name) && getDbdDeviceChoiceContext(offset, linePrefix) != null
        in 'A'..'Z', in 'a'..'z', in '0'..'9', '_' -> getStartupCommandContext(offset, linePrefix) != null
          || isDbdFile(file.name) && (
            getDbdKeywordContext(offset, linePrefix) != null ||
              getDbdDeviceRecordTypeContext(offset, linePrefix) != null ||
              getDbdDeviceLinkTypeContext(offset, linePrefix) != null ||
              getDbdDeviceSupportNameContext(offset, linePrefix) != null ||
              getDbdSimpleNameContext(offset, linePrefix, "driver") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "registrar") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "function") != null ||
              getDbdSimpleNameContext(offset, linePrefix, "variable") != null
            )
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

    private fun buildStartupLoadMacroTailInsertHandler(
      macroNames: List<String>,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      val document = context.document
      val tailEnd = getDbLoadRecordsTailInsertionEnd(document, replacementStartOffset)
      val tailText = buildDbLoadRecordsTail(macroNames)
      document.replaceString(replacementStartOffset, tailEnd, tailText)
      context.commitDocument()

      val firstValueOffset = replacementStartOffset + 5 + macroNames.first().length
      context.editor.caretModel.moveToOffset(firstValueOffset)
      context.editor.selectionModel.removeSelection()
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

    private fun buildDbLoadRecordsAssignmentLabel(macroNames: List<String>): String {
      return macroNames.joinToString(",") { "${it}=" }
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

    private fun getStartupLoadedRecordNameContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = STARTUP_LOADED_RECORD_NAME_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getStartupLoadMacroTailContext(offset: Int, linePrefix: String): StartupMacroTailCompletionContext? {
      val match = DB_LOAD_RECORDS_MACRO_TAIL_CONTEXT_REGEX.find(linePrefix) ?: return null
      val path = match.groups[1]?.value.orEmpty()
      return StartupMacroTailCompletionContext(
        path = path,
        startOffset = offset,
      )
    }

    private fun fillDbdCompletionVariants(
      project: Project,
      hostFile: VirtualFile,
      document: Document,
      offset: Int,
      linePrefix: String,
      result: CompletionResultSet,
    ) {
      getDbdDeviceRecordTypeContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        for (recordType in EpicsRecordCompletionSupport.getRecordTypes(project, hostFile)) {
          if (!matchesCompletionQuery(recordType, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(recordType)
              .withTypeText("EPICS record type", true)
              .withInsertHandler(buildSimpleReplacementAndPopupInsertHandler("$recordType, ", context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getDbdDeviceLinkTypeContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        for (linkType in DBD_DEVICE_LINK_TYPES) {
          if (!matchesCompletionQuery(linkType, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(linkType)
              .withTypeText("Device link type", true)
              .withInsertHandler(buildSimpleReplacementAndPopupInsertHandler("$linkType, ", context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getDbdDeviceSupportNameContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        val definitionsByName = getDbdSourceIndex(project).deviceSupportDefinitionsByName
        for ((supportName, definitions) in definitionsByName.entries.sortedBy { it.key.lowercase() }) {
          if (!matchesCompletionQuery(supportName, context.partial)) {
            continue
          }
          val definition = definitions.firstOrNull()
          resultSet.addElement(
            LookupElementBuilder.create(supportName)
              .withTypeText(
                definition?.let { "epicsExportAddress(${it.exportType}, $supportName) @ ${it.relativePath}:${it.line}" }
                  ?: "Exported device support structure",
                true,
              )
              .withInsertHandler(buildSimpleReplacementAndPopupInsertHandler("$supportName, \"", context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getDbdDeviceChoiceContext(offset, linePrefix)?.let { context ->
        val suggestions = linkedSetOf<String>()
        inferDeviceChoiceName(context.supportName)?.let(suggestions::add)
        suggestions += "device_name"
        val resultSet = result.withPrefixMatcher(context.partial)
        for (choice in suggestions) {
          if (!matchesCompletionQuery(choice, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(choice)
              .withTypeText("DTYP choice name", true)
              .withInsertHandler(buildSimpleReplacementInsertHandler("$choice\"", context.startOffset)),
          )
        }
        result.stopHere()
        return
      }

      getDbdSimpleNameContext(offset, linePrefix, "driver")?.let { context ->
        addDbdNamedDefinitionCompletions(
          result = result,
          context = context,
          definitionsByName = getDbdSourceIndex(project).driverDefinitionsByName,
          typeText = { definition -> "epicsExportAddress(${definition.exportType}, ${definition.name}) @ ${definition.relativePath}:${definition.line}" },
          insertValue = { definition -> definition.name },
          hostFile = hostFile,
        )
        return
      }

      getDbdSimpleNameContext(offset, linePrefix, "registrar")?.let { context ->
        addDbdNamedDefinitionCompletions(
          result = result,
          context = context,
          definitionsByName = getDbdSourceIndex(project).registrarDefinitionsByName,
          typeText = { definition -> "epicsExportRegistrar(${definition.name}) @ ${definition.relativePath}:${definition.line}" },
          insertValue = { definition -> definition.name },
          hostFile = hostFile,
        )
        return
      }

      getDbdSimpleNameContext(offset, linePrefix, "function")?.let { context ->
        addDbdNamedDefinitionCompletions(
          result = result,
          context = context,
          definitionsByName = getDbdSourceIndex(project).functionDefinitionsByName,
          typeText = { definition -> "epicsRegisterFunction(${definition.name}) @ ${definition.relativePath}:${definition.line}" },
          insertValue = { definition -> definition.name },
          hostFile = hostFile,
        )
        return
      }

      getDbdSimpleNameContext(offset, linePrefix, "variable")?.let { context ->
        addDbdNamedDefinitionCompletions(
          result = result,
          context = context,
          definitionsByName = getDbdSourceIndex(project).variableDefinitionsByName,
          typeText = { definition -> "epicsExportAddress(${definition.exportType}, ${definition.name}) @ ${definition.relativePath}:${definition.line}" },
          insertValue = { definition -> "${definition.name}, ${definition.exportType}" },
          hostFile = hostFile,
        )
        return
      }

      getDbdKeywordContext(offset, linePrefix)?.let { context ->
        val resultSet = result.withPrefixMatcher(context.partial)
        for (keyword in DBD_COMPLETION_KEYWORDS) {
          if (!matchesCompletionQuery(keyword, context.partial)) {
            continue
          }
          resultSet.addElement(
            LookupElementBuilder.create(keyword)
              .withTypeText("EPICS database definition keyword", true)
              .withInsertHandler(buildDbdKeywordInsertHandler(keyword, context.startOffset)),
          )
        }
        result.stopHere()
      }
    }

    private fun addDbdNamedDefinitionCompletions(
      result: CompletionResultSet,
      context: CompletionContext,
      definitionsByName: Map<String, List<DbdExportDefinition>>,
      typeText: (DbdExportDefinition) -> String,
      insertValue: (DbdExportDefinition) -> String,
      hostFile: VirtualFile,
    ) {
      val resultSet = result.withPrefixMatcher(context.partial)
      val hostDirectory = hostFile.parent?.toNioPath()?.normalize()
      for ((name, definitions) in definitionsByName.entries.sortedBy { it.key.lowercase() }) {
        if (!matchesCompletionQuery(name, context.partial)) {
          continue
        }
        val preferredDefinition = selectPreferredDbdDefinition(definitions, hostDirectory) ?: continue
        resultSet.addElement(
          LookupElementBuilder.create(name)
            .withTypeText(typeText(preferredDefinition), true)
            .withInsertHandler(buildSimpleReplacementInsertHandler(insertValue(preferredDefinition), context.startOffset)),
        )
      }
      result.stopHere()
    }

    private fun buildDbdKeywordInsertHandler(
      keyword: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      val document = context.document
      document.replaceString(replacementStartOffset, context.tailOffset, keyword)
      val tailStart = replacementStartOffset + keyword.length
      val tailEnd = getIdentifierTailInsertionEnd(document, tailStart)
      document.replaceString(tailStart, tailEnd, "(")
      context.commitDocument()
      context.editor.caretModel.moveToOffset(tailStart + 1)
      context.editor.selectionModel.removeSelection()
      requestCompletionPopup(context.project, context.editor)
    }

    private fun buildSimpleReplacementAndPopupInsertHandler(
      value: String,
      replacementStartOffset: Int,
    ): InsertHandler<LookupElement> = InsertHandler { context, _ ->
      context.setAddCompletionChar(false)
      context.document.replaceString(replacementStartOffset, context.tailOffset, value)
      context.commitDocument()
      context.editor.caretModel.moveToOffset(replacementStartOffset + value.length)
      context.editor.selectionModel.removeSelection()
      requestCompletionPopup(context.project, context.editor)
    }

    private fun getDbdKeywordContext(offset: Int, linePrefix: String): CompletionContext? {
      if ('#' in linePrefix) {
        return null
      }
      val match = DBD_KEYWORD_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getDbdDeviceRecordTypeContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = DBD_DEVICE_RECORD_TYPE_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getDbdDeviceLinkTypeContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = DBD_DEVICE_LINK_TYPE_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getDbdDeviceSupportNameContext(offset: Int, linePrefix: String): CompletionContext? {
      val match = DBD_DEVICE_SUPPORT_NAME_CONTEXT_REGEX.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getDbdDeviceChoiceContext(offset: Int, linePrefix: String): DbdDeviceChoiceCompletionContext? {
      val match = DBD_DEVICE_CHOICE_CONTEXT_REGEX.find(linePrefix) ?: return null
      val supportName = match.groups[1]?.value.orEmpty()
      val partial = match.groups[2]?.value.orEmpty()
      return DbdDeviceChoiceCompletionContext(
        supportName = supportName,
        partial = partial,
        startOffset = offset - partial.length,
      )
    }

    private fun getDbdSimpleNameContext(
      offset: Int,
      linePrefix: String,
      keyword: String,
    ): CompletionContext? {
      val regex = Regex("""^\s*${Regex.escape(keyword)}\(\s*([A-Za-z0-9_]*)$""")
      val match = regex.find(linePrefix) ?: return null
      val partial = match.groups[1]?.value.orEmpty()
      return CompletionContext(
        partial = partial,
        startOffset = offset - partial.length,
      )
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

    private fun getDbdSourceIndex(project: Project): DbdSourceIndex {
      val roots = ProjectRootManager.getInstance(project).contentRoots
        .filter { it.isDirectory }
        .map { it.toNioPath().normalize() }
        .sortedBy { it.toString().lowercase() }
      val cacheKey = roots.joinToString("|") { it.toString() }
      return dbdSourceIndexCache.computeIfAbsent(cacheKey) {
        scanDbdSourceIndex(roots)
      }
    }

    private fun scanDbdSourceIndex(roots: List<Path>): DbdSourceIndex {
      val deviceSupportDefinitions = mutableListOf<DbdExportDefinition>()
      val driverDefinitions = mutableListOf<DbdExportDefinition>()
      val registrarDefinitions = mutableListOf<DbdExportDefinition>()
      val functionDefinitions = mutableListOf<DbdExportDefinition>()
      val variableDefinitions = mutableListOf<DbdExportDefinition>()

      for (root in roots) {
        try {
          Files.walk(root).use { stream ->
            stream
              .filter { Files.isRegularFile(it) }
              .filter { it.extension.lowercase() in DBD_SOURCE_EXTENSIONS }
              .forEach { path ->
                val text = runCatching { Files.readString(path) }.getOrNull() ?: return@forEach
                deviceSupportDefinitions += extractDeviceSupportDefinitions(path, root, text)
                driverDefinitions += extractDriverDefinitions(path, root, text)
                registrarDefinitions += extractRegistrarDefinitions(path, root, text)
                functionDefinitions += extractFunctionDefinitions(path, root, text)
                variableDefinitions += extractVariableDefinitions(path, root, text)
              }
          }
        } catch (_: Exception) {
          continue
        }
      }

      return DbdSourceIndex(
        deviceSupportDefinitionsByName = groupDefinitionsByName(deviceSupportDefinitions),
        driverDefinitionsByName = groupDefinitionsByName(driverDefinitions),
        registrarDefinitionsByName = groupDefinitionsByName(registrarDefinitions),
        functionDefinitionsByName = groupDefinitionsByName(functionDefinitions),
        variableDefinitionsByName = groupDefinitionsByName(variableDefinitions),
      )
    }

    private fun groupDefinitionsByName(
      definitions: List<DbdExportDefinition>,
    ): Map<String, List<DbdExportDefinition>> {
      return definitions
        .groupBy { it.name }
        .mapValues { (_, entries) -> entries.sortedBy { it.relativePath.lowercase() } }
    }

    private fun extractDeviceSupportDefinitions(
      path: Path,
      root: Path,
      text: String,
    ): List<DbdExportDefinition> {
      val definitions = mutableListOf<DbdExportDefinition>()
      EPICS_EXPORT_ADDRESS_REGEX.findAll(text).forEach { match ->
        val exportType = match.groups[1]?.value.orEmpty()
        if (!DEVICE_SUPPORT_EXPORT_TYPE_REGEX.matches(exportType)) {
          return@forEach
        }
        val name = match.groups[2]?.value.orEmpty()
        if (name.isBlank()) {
          return@forEach
        }
        definitions += DbdExportDefinition(
          name = name,
          exportType = exportType,
          absolutePath = path.normalize(),
          relativePath = buildDefinitionRelativePath(root, path),
          line = lineNumberAt(text, match.range.first),
        )
      }
      return definitions
    }

    private fun extractDriverDefinitions(
      path: Path,
      root: Path,
      text: String,
    ): List<DbdExportDefinition> {
      val definitions = mutableListOf<DbdExportDefinition>()
      EPICS_EXPORT_ADDRESS_REGEX.findAll(text).forEach { match ->
        val exportType = match.groups[1]?.value.orEmpty()
        if (!exportType.equals("drvet", ignoreCase = true)) {
          return@forEach
        }
        val name = match.groups[2]?.value.orEmpty()
        if (name.isBlank()) {
          return@forEach
        }
        definitions += DbdExportDefinition(
          name = name,
          exportType = exportType,
          absolutePath = path.normalize(),
          relativePath = buildDefinitionRelativePath(root, path),
          line = lineNumberAt(text, match.range.first),
        )
      }
      return definitions
    }

    private fun extractRegistrarDefinitions(
      path: Path,
      root: Path,
      text: String,
    ): List<DbdExportDefinition> {
      val definitions = mutableListOf<DbdExportDefinition>()
      EPICS_EXPORT_REGISTRAR_REGEX.findAll(text).forEach { match ->
        val name = match.groups[1]?.value.orEmpty()
        if (name.isBlank()) {
          return@forEach
        }
        definitions += DbdExportDefinition(
          name = name,
          exportType = null,
          absolutePath = path.normalize(),
          relativePath = buildDefinitionRelativePath(root, path),
          line = lineNumberAt(text, match.range.first),
        )
      }
      return definitions
    }

    private fun extractFunctionDefinitions(
      path: Path,
      root: Path,
      text: String,
    ): List<DbdExportDefinition> {
      val definitions = mutableListOf<DbdExportDefinition>()
      EPICS_REGISTER_FUNCTION_REGEX.findAll(text).forEach { match ->
        val name = match.groups[1]?.value.orEmpty()
        if (name.isBlank()) {
          return@forEach
        }
        definitions += DbdExportDefinition(
          name = name,
          exportType = null,
          absolutePath = path.normalize(),
          relativePath = buildDefinitionRelativePath(root, path),
          line = lineNumberAt(text, match.range.first),
        )
      }
      return definitions
    }

    private fun extractVariableDefinitions(
      path: Path,
      root: Path,
      text: String,
    ): List<DbdExportDefinition> {
      val definitions = mutableListOf<DbdExportDefinition>()
      EPICS_EXPORT_ADDRESS_REGEX.findAll(text).forEach { match ->
        val exportType = match.groups[1]?.value.orEmpty()
        if (DEVICE_SUPPORT_EXPORT_TYPE_REGEX.matches(exportType) || exportType.equals("drvet", ignoreCase = true)) {
          return@forEach
        }
        val name = match.groups[2]?.value.orEmpty()
        if (name.isBlank()) {
          return@forEach
        }
        definitions += DbdExportDefinition(
          name = name,
          exportType = exportType,
          absolutePath = path.normalize(),
          relativePath = buildDefinitionRelativePath(root, path),
          line = lineNumberAt(text, match.range.first),
        )
      }
      return definitions
    }

    private fun selectPreferredDbdDefinition(
      definitions: List<DbdExportDefinition>,
      hostDirectory: Path?,
    ): DbdExportDefinition? {
      if (definitions.isEmpty()) {
        return null
      }
      if (hostDirectory == null) {
        return definitions.first()
      }
      return definitions.firstOrNull { definition ->
        definition.absolutePath.startsWith(hostDirectory)
      } ?: definitions.first()
    }

    private fun buildDefinitionRelativePath(root: Path, path: Path): String {
      return runCatching { root.relativize(path).toString() }
        .getOrElse { path.toString() }
        .replace('\\', '/')
    }

    private fun lineNumberAt(text: String, offset: Int): Int {
      return text.take(offset.coerceIn(0, text.length)).count { it == '\n' } + 1
    }

    private fun inferDeviceChoiceName(supportName: String): String? {
      val text = supportName.trim()
      if (text.isBlank()) {
        return null
      }
      val stripped = text
        .removePrefix("dev")
        .removePrefix("Dev")
        .removePrefix("DSET_")
        .trimStart('_')
      return stripped.ifBlank { text }
    }

    private fun isDatabaseFile(fileName: String): Boolean {
      val extension = fileName.substringAfterLast('.', "").lowercase()
      return extension in setOf("db", "vdb", "template")
    }

    private fun isStartupFile(fileName: String): Boolean {
      val extension = fileName.substringAfterLast('.', "").lowercase()
      return extension in setOf("cmd", "iocsh") || fileName == "st.cmd"
    }

    private fun isDbdFile(fileName: String): Boolean {
      return fileName.substringAfterLast('.', "").lowercase() == "dbd"
    }

    private fun extractMacroNamesFromDatabaseFile(path: Path): List<String> {
      val text = runCatching { Files.readString(path) }.getOrNull() ?: return emptyList()
      return extractMacroNames(maskDatabaseComments(text))
    }

    private fun extractMacroNamesFromStartupDatabaseFile(
      project: com.intellij.openapi.project.Project,
      hostFile: com.intellij.openapi.vfs.VirtualFile,
      text: String,
      untilOffset: Int,
      rawPath: String,
    ): List<String> {
      val ownerRoot = EpicsPathResolver.findOwningEpicsRoot(project, hostFile)
      val resolved = EpicsPathResolver.resolveStartupDatabasePath(
        hostFile = hostFile,
        ownerRoot = ownerRoot,
        text = text,
        untilOffset = untilOffset,
        rawPath = rawPath,
      ) ?: return emptyList()
      return extractMacroNamesFromDatabaseFile(resolved)
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

    private data class StartupMacroTailCompletionContext(
      val path: String,
      val startOffset: Int,
    )

    private data class DbdDeviceChoiceCompletionContext(
      val supportName: String,
      val partial: String,
      val startOffset: Int,
    )

    private data class DbdExportDefinition(
      val name: String,
      val exportType: String?,
      val absolutePath: Path,
      val relativePath: String,
      val line: Int,
    )

    private data class DbdSourceIndex(
      val deviceSupportDefinitionsByName: Map<String, List<DbdExportDefinition>>,
      val driverDefinitionsByName: Map<String, List<DbdExportDefinition>>,
      val registrarDefinitionsByName: Map<String, List<DbdExportDefinition>>,
      val functionDefinitionsByName: Map<String, List<DbdExportDefinition>>,
      val variableDefinitionsByName: Map<String, List<DbdExportDefinition>>,
    )

    private val RECORD_TYPE_CONTEXT_REGEX = Regex("""record\(\s*([A-Za-z0-9_]*)$""")
    private val RECORD_NAME_CONTEXT_REGEX = Regex("""record\(\s*[A-Za-z0-9_]+\s*,\s*"([^"\n]*)$""")
    private val FIELD_NAME_CONTEXT_REGEX = Regex("""field\(\s*(?:"?([A-Za-z0-9_]*))$""")
    private val FIELD_VALUE_CONTEXT_REGEX =
      Regex("""field\(\s*(?:"([A-Za-z0-9_]+)"|([A-Za-z0-9_]+))\s*,\s*"([^"\n]*)$""")
    private val STARTUP_COMMAND_CONTEXT_REGEX = Regex("""^\s*([A-Za-z0-9_]*)$""")
    private val STARTUP_LOADED_RECORD_NAME_CONTEXT_REGEX = Regex("""dbpf\(\s*"([^"\n]*)$""")
    private val DB_LOAD_RECORDS_PATH_CONTEXT_REGEX = Regex("""dbLoadRecords\(\s*"([^"\n]*)$""")
    private val DB_LOAD_RECORDS_MACRO_TAIL_CONTEXT_REGEX = Regex("""dbLoadRecords\(\s*"([^"\n]+)"\s*$""")
    private val DB_LOAD_TEMPLATE_PATH_CONTEXT_REGEX = Regex("""dbLoadTemplate\(\s*"([^"\n]*)$""")
    private val DBD_KEYWORD_CONTEXT_REGEX = Regex("""^\s*([A-Za-z0-9_]*)$""")
    private val DBD_DEVICE_RECORD_TYPE_CONTEXT_REGEX = Regex("""^\s*device\(\s*([A-Za-z0-9_]*)$""")
    private val DBD_DEVICE_LINK_TYPE_CONTEXT_REGEX = Regex("""^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]*)$""")
    private val DBD_DEVICE_SUPPORT_NAME_CONTEXT_REGEX =
      Regex("""^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]*)$""")
    private val DBD_DEVICE_CHOICE_CONTEXT_REGEX =
      Regex("""^\s*device\(\s*[A-Za-z0-9_]+\s*,\s*[A-Za-z0-9_]+\s*,\s*([A-Za-z0-9_]+)\s*,\s*"([^"\n]*)$""")
    private val EPICS_MACRO_REFERENCE_REGEX =
      Regex("""\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}""")
    private val EPICS_EXPORT_ADDRESS_REGEX =
      Regex("""epicsExportAddress\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
    private val EPICS_EXPORT_REGISTRAR_REGEX =
      Regex("""epicsExportRegistrar\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
    private val EPICS_REGISTER_FUNCTION_REGEX =
      Regex("""epicsRegisterFunction\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
    private val DEVICE_SUPPORT_EXPORT_TYPE_REGEX = Regex("""(?:^|_)dset$""", RegexOption.IGNORE_CASE)
    private val STARTUP_COMMANDS = listOf("dbLoadRecords", "dbLoadTemplate")
    private val DBD_COMPLETION_KEYWORDS = listOf(
      "device",
      "driver",
      "registrar",
      "function",
      "variable",
      "recordtype",
      "menu",
      "field",
      "choice",
      "breaktable",
    )
    private val DBD_DEVICE_LINK_TYPES = listOf("INST_IO")
    private val DBD_SOURCE_EXTENSIONS = setOf("c", "cc", "cpp", "cxx", "h", "hh", "hpp")
    private val dbdSourceIndexCache = ConcurrentHashMap<String, DbdSourceIndex>()
  }
}
