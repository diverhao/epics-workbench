package org.epics.workbench.refactor

import com.intellij.openapi.command.WriteCommandAction
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.openapi.vfs.toNioPathOrNull
import com.intellij.psi.search.FilenameIndex
import com.intellij.psi.search.GlobalSearchScope
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.toc.EpicsDatabaseToc
import java.nio.charset.Charset
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.extension
import kotlin.io.path.isRegularFile

internal enum class EpicsSemanticSymbolKind {
  RECORD,
  MACRO,
  RECORD_TYPE,
  FIELD,
  DEVICE_SUPPORT,
  DRIVER,
  REGISTRAR,
  FUNCTION,
  VARIABLE,
}

internal data class EpicsSemanticSymbol(
  val kind: EpicsSemanticSymbolKind,
  val name: String,
  val file: VirtualFile,
  val startOffset: Int,
  val endOffset: Int,
  val searchNames: Set<String> = linkedSetOf(name),
  val recordType: String? = null,
)

internal data class EpicsSemanticOccurrence(
  val file: VirtualFile,
  val startOffset: Int,
  val endOffset: Int,
  val declaration: Boolean,
  val lineNumber: Int,
  val lineText: String,
)

internal object EpicsSemanticSymbolSupport {
  fun findSymbol(
    project: Project,
    file: VirtualFile,
    offset: Int,
  ): EpicsSemanticSymbol? {
    val text = readText(file) ?: return null
    return when {
      isDatabaseFile(file) -> {
        findDatabaseRecordSymbol(file, text, offset)
          ?: findDatabaseRecordTypeSymbol(file, text, offset)
          ?: findDatabaseFieldSymbol(file, text, offset)
          ?: findMacroSymbol(file, text, offset)
      }

      isStartupFile(file) -> {
        findStartupRecordSymbol(project, file, text, offset)
          ?: findStartupMacroSymbol(file, text, offset)
          ?: findMacroReferenceSymbol(file, text, offset)
      }

      isSubstitutionsFile(file) -> {
        findSubstitutionsMacroSymbol(file, text, offset)
          ?: findMacroReferenceSymbol(file, text, offset)
      }

      isPvlistFile(file) -> {
        findPvlistRecordSymbol(file, text, offset)
          ?: findPvlistMacroSymbol(file, text, offset)
          ?: findMacroReferenceSymbol(file, text, offset)
      }

      isProbeFile(file) -> findProbeRecordSymbol(file, text, offset)
      isDbdFile(file) -> {
        findDbdRecordTypeSymbol(file, text, offset)
          ?: findDbdFieldSymbol(file, text, offset)
          ?: findDbdNamedSymbol(file, text, offset)
      }

      isSourceFile(file) -> findSourceNamedSymbol(file, text, offset)
      else -> null
    }
  }

  fun collectOccurrences(
    project: Project,
    symbol: EpicsSemanticSymbol,
    includeDeclarations: Boolean = true,
  ): List<EpicsSemanticOccurrence> {
    val occurrences = mutableListOf<EpicsSemanticOccurrence>()
    val seen = linkedSetOf<String>()
    for (file in collectSemanticFiles(project, symbol.file)) {
      val text = readText(file) ?: continue
      when (symbol.kind) {
        EpicsSemanticSymbolKind.RECORD -> collectRecordOccurrences(file, text, symbol.searchNames, includeDeclarations, occurrences, seen)
        EpicsSemanticSymbolKind.MACRO -> collectMacroOccurrences(file, text, symbol.name, includeDeclarations, occurrences, seen)
        EpicsSemanticSymbolKind.RECORD_TYPE -> collectRecordTypeOccurrences(file, text, symbol.name, includeDeclarations, occurrences, seen)
        EpicsSemanticSymbolKind.FIELD -> collectFieldOccurrences(file, text, symbol.recordType, symbol.name, includeDeclarations, occurrences, seen)
        EpicsSemanticSymbolKind.DEVICE_SUPPORT,
        EpicsSemanticSymbolKind.DRIVER,
        EpicsSemanticSymbolKind.REGISTRAR,
        EpicsSemanticSymbolKind.FUNCTION,
        EpicsSemanticSymbolKind.VARIABLE,
        -> collectNamedSymbolOccurrences(file, text, symbol.kind, symbol.name, includeDeclarations, occurrences, seen)
      }
    }
    return occurrences.sortedWith(
      compareBy<EpicsSemanticOccurrence>({ it.file.path.lowercase() }, { it.startOffset }),
    )
  }

  fun validateRename(symbol: EpicsSemanticSymbol, newName: String): String? {
    val trimmed = newName.trim()
    if (trimmed.isEmpty()) {
      return "Name cannot be empty."
    }

    return when (symbol.kind) {
      EpicsSemanticSymbolKind.MACRO ->
        if (MACRO_NAME_REGEX.matches(trimmed)) null else "EPICS macro names must match [A-Za-z_][A-Za-z0-9_]*."
      EpicsSemanticSymbolKind.RECORD_TYPE ->
        if (SYMBOL_NAME_REGEX.matches(trimmed)) null else "EPICS record types must match [A-Za-z_][A-Za-z0-9_]*."
      EpicsSemanticSymbolKind.FIELD ->
        if (FIELD_NAME_REGEX.matches(trimmed)) null else "EPICS field names must match [A-Z][A-Z0-9_]*."
      EpicsSemanticSymbolKind.DEVICE_SUPPORT,
      EpicsSemanticSymbolKind.DRIVER,
      EpicsSemanticSymbolKind.REGISTRAR,
      EpicsSemanticSymbolKind.FUNCTION,
      EpicsSemanticSymbolKind.VARIABLE,
      -> if (SYMBOL_NAME_REGEX.matches(trimmed)) null else "EPICS symbols must match [A-Za-z_][A-Za-z0-9_]*."
      EpicsSemanticSymbolKind.RECORD -> if (trimmed.contains('"') || trimmed.contains('\n')) "EPICS record names cannot contain quotes or newlines." else null
    }
  }

  fun detectRenameConflict(
    project: Project,
    symbol: EpicsSemanticSymbol,
    newName: String,
  ): String? {
    val trimmed = newName.trim()
    if (trimmed == symbol.name) {
      return null
    }

    val files = collectSemanticFiles(project, symbol.file)
    for (file in files) {
      val text = readText(file) ?: continue
      when (symbol.kind) {
        EpicsSemanticSymbolKind.RECORD -> {
          val declarations = extractDatabaseRecordDeclarations(text)
          if (declarations.any { declaration -> declaration.name == trimmed && !sameRange(file, declaration.nameStart, declaration.nameEnd, symbol) }) {
            return "Record \"$trimmed\" already exists in ${file.name}."
          }
        }

        EpicsSemanticSymbolKind.MACRO -> {
          if (collectMacroDefinitionRanges(file, text, trimmed).any { occurrence -> !sameOccurrence(occurrence, symbol) }) {
            return "Macro \"$trimmed\" is already defined in ${file.name}."
          }
        }

        EpicsSemanticSymbolKind.RECORD_TYPE -> {
          if (collectRecordTypeDeclarationRanges(file, text, trimmed).any { occurrence -> !sameOccurrence(occurrence, symbol) }) {
            return "Record type \"$trimmed\" already exists in ${file.name}."
          }
        }

        EpicsSemanticSymbolKind.FIELD -> {
          if (symbol.recordType != null && collectFieldDeclarationRanges(file, text, symbol.recordType, trimmed).any { occurrence -> !sameOccurrence(occurrence, symbol) }) {
            return "Field \"$trimmed\" is already declared for record type \"${symbol.recordType}\" in ${file.name}."
          }
        }

        else -> {
          if (collectNamedSymbolDeclarationRanges(file, text, symbol.kind, trimmed).any { occurrence -> !sameOccurrence(occurrence, symbol) }) {
            return "Symbol \"$trimmed\" already exists in ${file.name}."
          }
        }
      }
    }
    return null
  }

  fun applyRename(
    project: Project,
    symbol: EpicsSemanticSymbol,
    newName: String,
  ): Int {
    val occurrences = collectOccurrences(project, symbol, includeDeclarations = true)
    if (occurrences.isEmpty()) {
      return 0
    }

    val editsByFile = occurrences.groupBy { it.file }
    WriteCommandAction.runWriteCommandAction(project, "Rename EPICS Symbol", null, Runnable {
      val documentManager = FileDocumentManager.getInstance()
      editsByFile.forEach { (file, fileOccurrences) ->
        val document = documentManager.getDocument(file) ?: return@forEach
        fileOccurrences
          .sortedByDescending { it.startOffset }
          .forEach { occurrence ->
            document.replaceString(occurrence.startOffset, occurrence.endOffset, newName)
          }
        documentManager.saveDocument(document)
      }
    })
    return occurrences.size
  }

  private fun collectRecordOccurrences(
    file: VirtualFile,
    text: String,
    searchNames: Set<String>,
    includeDeclarations: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    if (isDatabaseFile(file)) {
      if (includeDeclarations) {
        for (declaration in extractDatabaseRecordDeclarations(text)) {
          val declarationNames = getDatabaseRecordSearchNames(text, declaration.name)
          if (declarationNames.any(searchNames::contains)) {
            addOccurrence(file, text, declaration.nameStart, declaration.nameEnd, true, occurrences, seen)
          }
        }
      }
      val tocMacros = EpicsDatabaseToc.extractMacroAssignmentValues(text)
      for (declaration in extractDatabaseRecordDeclarations(text)) {
        for (fieldDeclaration in extractDatabaseFieldDeclarations(text, declaration)) {
          for (candidate in extractLinkedRecordCandidates(fieldDeclaration.value, fieldDeclaration.valueStart, tocMacros)) {
            if (searchNames.contains(candidate.name)) {
              addOccurrence(file, text, candidate.start, candidate.end, false, occurrences, seen)
            }
          }
        }
      }
      return
    }

    if (isStartupFile(file)) {
      val regex = Regex("""dbpf\(\s*"((?:[^"\\]|\\.)*)"""")
      val lines = splitLinesWithOffsets(text)
      for ((lineText, lineStart) in lines) {
        val sanitizedLine = maskHashComments(lineText)
        for (match in regex.findAll(sanitizedLine)) {
          val value = match.groups[1]?.value.orEmpty()
          val valueStart = lineStart + match.range.first + match.value.length - value.length - 1
          for (candidate in extractLinkedRecordCandidates(value, valueStart, emptyMap())) {
            if (searchNames.contains(candidate.name)) {
              addOccurrence(file, text, candidate.start, candidate.end, false, occurrences, seen)
            }
          }
        }
      }
      return
    }

    if (isPvlistFile(file)) {
      for (reference in extractSimpleRecordLineReferences(text)) {
        if (searchNames.contains(reference.name)) {
          addOccurrence(file, text, reference.start, reference.end, false, occurrences, seen)
        }
      }
      return
    }

    if (isProbeFile(file)) {
      for (reference in extractProbeRecordReferences(text)) {
        if (searchNames.contains(reference.name)) {
          addOccurrence(file, text, reference.start, reference.end, false, occurrences, seen)
        }
      }
    }
  }

  private fun collectMacroOccurrences(
    file: VirtualFile,
    text: String,
    macroName: String,
    includeDeclarations: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    if (includeDeclarations) {
      collectMacroDefinitionRanges(file, text, macroName)
        .forEach { occurrence ->
          addOccurrence(file, text, occurrence.startOffset, occurrence.endOffset, true, occurrences, seen)
        }
    }
    for (reference in extractMacroReferenceRanges(text, macroName)) {
      addOccurrence(file, text, reference.first, reference.second, false, occurrences, seen)
    }
  }

  private fun collectRecordTypeOccurrences(
    file: VirtualFile,
    text: String,
    recordTypeName: String,
    includeDeclarations: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    if (isDatabaseFile(file)) {
      for (declaration in extractDatabaseRecordDeclarations(text)) {
        if (declaration.recordType == recordTypeName) {
          addOccurrence(file, text, declaration.recordTypeStart, declaration.recordTypeEnd, false, occurrences, seen)
        }
      }
      return
    }

    if (!isDbdFile(file)) {
      return
    }

    if (includeDeclarations) {
      for (declaration in extractDbdRecordTypeDeclarations(text)) {
        if (declaration.name == recordTypeName) {
          addOccurrence(file, text, declaration.nameStart, declaration.nameEnd, true, occurrences, seen)
        }
      }
    }

    for (entry in extractDbdDeviceDeclarations(text)) {
      if (entry.recordType == recordTypeName) {
        addOccurrence(file, text, entry.recordTypeStart, entry.recordTypeEnd, false, occurrences, seen)
      }
    }
  }

  private fun collectFieldOccurrences(
    file: VirtualFile,
    text: String,
    recordType: String?,
    fieldName: String,
    includeDeclarations: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    if (recordType.isNullOrBlank()) {
      return
    }

    if (isDatabaseFile(file)) {
      for (recordDeclaration in extractDatabaseRecordDeclarations(text)) {
        if (recordDeclaration.recordType != recordType) {
          continue
        }
        for (fieldDeclaration in extractDatabaseFieldDeclarations(text, recordDeclaration)) {
          if (fieldDeclaration.fieldName == fieldName) {
            addOccurrence(file, text, fieldDeclaration.fieldNameStart, fieldDeclaration.fieldNameEnd, false, occurrences, seen)
          }
        }
      }
      return
    }

    if (!isDbdFile(file) || !includeDeclarations) {
      return
    }

    for (recordTypeDeclaration in extractDbdRecordTypeDeclarations(text)) {
      if (recordTypeDeclaration.name != recordType) {
        continue
      }
      for (fieldDeclaration in extractDbdFieldDeclarations(text, recordTypeDeclaration)) {
        if (fieldDeclaration.fieldName == fieldName) {
          addOccurrence(file, text, fieldDeclaration.fieldNameStart, fieldDeclaration.fieldNameEnd, true, occurrences, seen)
        }
      }
    }
  }

  private fun collectNamedSymbolOccurrences(
    file: VirtualFile,
    text: String,
    kind: EpicsSemanticSymbolKind,
    name: String,
    includeDeclarations: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    if (isDbdFile(file)) {
      when (kind) {
        EpicsSemanticSymbolKind.DEVICE_SUPPORT ->
          extractDbdDeviceDeclarations(text).filter { it.supportName == name }.forEach { entry ->
            addOccurrence(file, text, entry.supportNameStart, entry.supportNameEnd, true, occurrences, seen)
          }

        EpicsSemanticSymbolKind.DRIVER,
        EpicsSemanticSymbolKind.REGISTRAR,
        EpicsSemanticSymbolKind.FUNCTION,
        EpicsSemanticSymbolKind.VARIABLE,
        -> extractDbdNamedDeclarations(text, kind).filter { it.name == name }.forEach { entry ->
          addOccurrence(file, text, entry.nameStart, entry.nameEnd, true, occurrences, seen)
        }

        else -> Unit
      }
    }

    if (!isSourceFile(file) || !includeDeclarations) {
      return
    }

    extractSourceNamedDeclarations(text)
      .filter { entry -> entry.kind == kind && entry.name == name }
      .forEach { entry ->
        addOccurrence(file, text, entry.nameStart, entry.nameEnd, true, occurrences, seen)
      }
  }

  private fun collectSemanticFiles(project: Project, currentFile: VirtualFile): List<VirtualFile> {
    val files = linkedMapOf<String, VirtualFile>()
    files[currentFile.path] = currentFile
    for (contentRoot in ProjectRootManager.getInstance(project).contentRoots) {
      collectSemanticFilesRecursively(contentRoot, files)
    }
    return files.values.toList()
  }

  private fun collectSemanticFilesRecursively(directory: VirtualFile, files: MutableMap<String, VirtualFile>) {
    if (!directory.isDirectory) {
      if (isSemanticFile(directory)) {
        files[directory.path] = directory
      }
      return
    }
    if (directory.name in IGNORED_DIRECTORY_NAMES) {
      return
    }
    directory.children.forEach { child -> collectSemanticFilesRecursively(child, files) }
  }

  private fun addOccurrence(
    file: VirtualFile,
    text: String,
    startOffset: Int,
    endOffset: Int,
    declaration: Boolean,
    occurrences: MutableList<EpicsSemanticOccurrence>,
    seen: MutableSet<String>,
  ) {
    val key = "${file.path}:$startOffset:$endOffset:$declaration"
    if (!seen.add(key)) {
      return
    }
    val (lineNumber, lineText) = getLineContext(text, startOffset)
    occurrences += EpicsSemanticOccurrence(file, startOffset, endOffset, declaration, lineNumber, lineText)
  }

  private fun sameRange(file: VirtualFile, start: Int, end: Int, symbol: EpicsSemanticSymbol): Boolean {
    return file.path == symbol.file.path && start == symbol.startOffset && end == symbol.endOffset
  }

  private fun sameOccurrence(occurrence: EpicsSemanticOccurrence, symbol: EpicsSemanticSymbol): Boolean {
    return sameRange(occurrence.file, occurrence.startOffset, occurrence.endOffset, symbol)
  }

  private fun findDatabaseRecordSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    val tocMacros = EpicsDatabaseToc.extractMacroAssignmentValues(text)
    for (declaration in extractDatabaseRecordDeclarations(text)) {
      if (offset in declaration.nameStart..declaration.nameEnd) {
        return EpicsSemanticSymbol(
          kind = EpicsSemanticSymbolKind.RECORD,
          name = declaration.name,
          file = file,
          startOffset = declaration.nameStart,
          endOffset = declaration.nameEnd,
          searchNames = getDatabaseRecordSearchNames(text, declaration.name),
        )
      }
      for (fieldDeclaration in extractDatabaseFieldDeclarations(text, declaration)) {
        for (candidate in extractLinkedRecordCandidates(fieldDeclaration.value, fieldDeclaration.valueStart, tocMacros)) {
          if (offset in candidate.start..candidate.end) {
            return EpicsSemanticSymbol(
              kind = EpicsSemanticSymbolKind.RECORD,
              name = candidate.name,
              file = file,
              startOffset = candidate.start,
              endOffset = candidate.end,
              searchNames = candidate.searchNames,
            )
          }
        }
      }
    }
    return null
  }

  private fun findStartupRecordSymbol(project: Project, file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    val definition = EpicsRecordResolver.resolveRecordDefinition(project, file, offset) ?: return null
    return EpicsSemanticSymbol(
      kind = EpicsSemanticSymbolKind.RECORD,
      name = definition.recordName,
      file = file,
      startOffset = offset.coerceAtLeast(0),
      endOffset = offset.coerceAtLeast(0) + 1,
      searchNames = linkedSetOf(definition.recordName),
    )
  }

  private fun findPvlistRecordSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    return extractSimpleRecordLineReferences(text)
      .firstOrNull { candidate -> offset in candidate.start..candidate.end }
      ?.let { candidate ->
        EpicsSemanticSymbol(EpicsSemanticSymbolKind.RECORD, candidate.name, file, candidate.start, candidate.end)
      }
  }

  private fun findProbeRecordSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    return extractProbeRecordReferences(text)
      .firstOrNull { candidate -> offset in candidate.start..candidate.end }
      ?.let { candidate ->
        EpicsSemanticSymbol(EpicsSemanticSymbolKind.RECORD, candidate.name, file, candidate.start, candidate.end)
      }
  }

  private fun findMacroSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    return findMacroReferenceSymbol(file, text, offset)
  }

  private fun findStartupMacroSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    for (statement in extractStartupStatements(text)) {
      if (statement.kind == "envSet" && statement.nameStart != null && statement.nameEnd != null && offset in statement.nameStart..statement.nameEnd) {
        return EpicsSemanticSymbol(EpicsSemanticSymbolKind.MACRO, statement.name.orEmpty(), file, statement.nameStart, statement.nameEnd)
      }
      if (statement.kind == "load" && statement.command == "dbLoadRecords" && statement.macros != null && statement.macroValueStart != null) {
        for ((name, range) in extractNamedAssignmentsWithRanges(statement.macros, statement.macroValueStart).entries) {
          if (offset in range.first..range.second) {
            return EpicsSemanticSymbol(EpicsSemanticSymbolKind.MACRO, name, file, range.first, range.second)
          }
        }
      }
    }
    return null
  }

  private fun findSubstitutionsMacroSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    collectMacroDefinitionRanges(file, text, null).firstOrNull { occurrence -> offset in occurrence.startOffset..occurrence.endOffset }?.let { occurrence ->
      val name = text.substring(occurrence.startOffset, occurrence.endOffset)
      return EpicsSemanticSymbol(EpicsSemanticSymbolKind.MACRO, name, file, occurrence.startOffset, occurrence.endOffset)
    }
    return null
  }

  private fun findPvlistMacroSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    collectMacroDefinitionRanges(file, text, null).firstOrNull { occurrence -> offset in occurrence.startOffset..occurrence.endOffset }?.let { occurrence ->
      val name = text.substring(occurrence.startOffset, occurrence.endOffset)
      return EpicsSemanticSymbol(EpicsSemanticSymbolKind.MACRO, name, file, occurrence.startOffset, occurrence.endOffset)
    }
    return null
  }

  private fun findMacroReferenceSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    val macroRegex = EPICS_MACRO_REFERENCE_REGEX
    for (match in macroRegex.findAll(maskHashComments(text))) {
      val name = match.groups[1]?.value ?: match.groups[2]?.value ?: continue
      val nameStart = match.range.first + 2
      val nameEnd = nameStart + name.length
      if (offset in nameStart..nameEnd) {
        return EpicsSemanticSymbol(EpicsSemanticSymbolKind.MACRO, name, file, nameStart, nameEnd)
      }
    }
    return null
  }

  private fun findDatabaseRecordTypeSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    return extractDatabaseRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.recordTypeStart..declaration.recordTypeEnd }
      ?.let { declaration ->
        EpicsSemanticSymbol(EpicsSemanticSymbolKind.RECORD_TYPE, declaration.recordType, file, declaration.recordTypeStart, declaration.recordTypeEnd)
      }
  }

  private fun findDatabaseFieldSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    for (declaration in extractDatabaseRecordDeclarations(text)) {
      if (offset !in declaration.recordStart..declaration.recordEnd) {
        continue
      }
      for (fieldDeclaration in extractDatabaseFieldDeclarations(text, declaration)) {
        if (offset in fieldDeclaration.fieldNameStart..fieldDeclaration.fieldNameEnd) {
          return EpicsSemanticSymbol(
            EpicsSemanticSymbolKind.FIELD,
            fieldDeclaration.fieldName,
            file,
            fieldDeclaration.fieldNameStart,
            fieldDeclaration.fieldNameEnd,
            recordType = declaration.recordType,
          )
        }
      }
    }
    return null
  }

  private fun findDbdRecordTypeSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    extractDbdRecordTypeDeclarations(text).firstOrNull { declaration -> offset in declaration.nameStart..declaration.nameEnd }?.let { declaration ->
      return EpicsSemanticSymbol(EpicsSemanticSymbolKind.RECORD_TYPE, declaration.name, file, declaration.nameStart, declaration.nameEnd)
    }
    extractDbdDeviceDeclarations(text).firstOrNull { declaration -> offset in declaration.recordTypeStart..declaration.recordTypeEnd }?.let { declaration ->
      return EpicsSemanticSymbol(EpicsSemanticSymbolKind.RECORD_TYPE, declaration.recordType, file, declaration.recordTypeStart, declaration.recordTypeEnd)
    }
    return null
  }

  private fun findDbdFieldSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    for (recordTypeDeclaration in extractDbdRecordTypeDeclarations(text)) {
      if (offset !in recordTypeDeclaration.blockStart..recordTypeDeclaration.blockEnd) {
        continue
      }
      for (fieldDeclaration in extractDbdFieldDeclarations(text, recordTypeDeclaration)) {
        if (offset in fieldDeclaration.fieldNameStart..fieldDeclaration.fieldNameEnd) {
          return EpicsSemanticSymbol(
            EpicsSemanticSymbolKind.FIELD,
            fieldDeclaration.fieldName,
            file,
            fieldDeclaration.fieldNameStart,
            fieldDeclaration.fieldNameEnd,
            recordType = recordTypeDeclaration.name,
          )
        }
      }
    }
    return null
  }

  private fun findDbdNamedSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    extractDbdDeviceDeclarations(text).firstOrNull { entry -> offset in entry.supportNameStart..entry.supportNameEnd }?.let { entry ->
      return EpicsSemanticSymbol(EpicsSemanticSymbolKind.DEVICE_SUPPORT, entry.supportName, file, entry.supportNameStart, entry.supportNameEnd)
    }
    for (kind in listOf(EpicsSemanticSymbolKind.DRIVER, EpicsSemanticSymbolKind.REGISTRAR, EpicsSemanticSymbolKind.FUNCTION, EpicsSemanticSymbolKind.VARIABLE)) {
      extractDbdNamedDeclarations(text, kind).firstOrNull { entry -> offset in entry.nameStart..entry.nameEnd }?.let { entry ->
        return EpicsSemanticSymbol(kind, entry.name, file, entry.nameStart, entry.nameEnd)
      }
    }
    return null
  }

  private fun findSourceNamedSymbol(file: VirtualFile, text: String, offset: Int): EpicsSemanticSymbol? {
    return extractSourceNamedDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.nameStart..declaration.nameEnd }
      ?.let { declaration ->
        EpicsSemanticSymbol(declaration.kind, declaration.name, file, declaration.nameStart, declaration.nameEnd)
      }
  }

  private fun getDatabaseRecordSearchNames(text: String, recordName: String): Set<String> {
    val names = linkedSetOf<String>()
    if (recordName.isNotBlank()) {
      names += recordName
      val expanded = expandEpicsMacros(recordName, EpicsDatabaseToc.extractMacroAssignmentValues(text))
      if (expanded.isNotBlank()) {
        names += expanded
      }
    }
    return names
  }

  private fun collectMacroDefinitionRanges(file: VirtualFile, text: String, macroName: String?): List<EpicsSemanticOccurrence> {
    val occurrences = mutableListOf<EpicsSemanticOccurrence>()
    val seen = linkedSetOf<String>()
    if (isStartupFile(file)) {
      for (statement in extractStartupStatements(text)) {
        if (statement.kind == "envSet" && statement.nameStart != null && statement.nameEnd != null && (macroName == null || statement.name == macroName)) {
          addOccurrence(file, text, statement.nameStart, statement.nameEnd, true, occurrences, seen)
        }
        if (statement.kind == "load" && statement.command == "dbLoadRecords" && statement.macros != null && statement.macroValueStart != null) {
          for ((name, range) in extractNamedAssignmentsWithRanges(statement.macros, statement.macroValueStart).entries) {
            if (macroName == null || macroName == name) {
              addOccurrence(file, text, range.first, range.second, true, occurrences, seen)
            }
          }
        }
      }
    }
    if (isPvlistFile(file)) {
      for (line in splitLinesWithOffsets(text)) {
        val match = PVLIST_MACRO_DEFINITION_REGEX.find(line.first) ?: continue
        val name = match.groups[2]?.value.orEmpty()
        if (macroName == null || name == macroName) {
          val start = line.second + (match.groups[2]?.range?.first ?: continue)
          addOccurrence(file, text, start, start + name.length, true, occurrences, seen)
        }
      }
    }
    if (isSubstitutionsFile(file)) {
      for (match in SUBSTITUTIONS_MACRO_DEFINITION_REGEX.findAll(text)) {
        val name = match.groups[1]?.value.orEmpty()
        if (macroName == null || name == macroName) {
          val start = match.range.first
          addOccurrence(file, text, start, start + name.length, true, occurrences, seen)
        }
      }
    }
    return occurrences
  }

  private fun collectRecordTypeDeclarationRanges(file: VirtualFile, text: String, recordTypeName: String): List<EpicsSemanticOccurrence> {
    val occurrences = mutableListOf<EpicsSemanticOccurrence>()
    val seen = linkedSetOf<String>()
    if (isDbdFile(file)) {
      extractDbdRecordTypeDeclarations(text).filter { it.name == recordTypeName }.forEach { declaration ->
        addOccurrence(file, text, declaration.nameStart, declaration.nameEnd, true, occurrences, seen)
      }
    }
    return occurrences
  }

  private fun collectFieldDeclarationRanges(file: VirtualFile, text: String, recordType: String, fieldName: String): List<EpicsSemanticOccurrence> {
    val occurrences = mutableListOf<EpicsSemanticOccurrence>()
    val seen = linkedSetOf<String>()
    if (isDbdFile(file)) {
      for (recordTypeDeclaration in extractDbdRecordTypeDeclarations(text)) {
        if (recordTypeDeclaration.name != recordType) {
          continue
        }
        for (fieldDeclaration in extractDbdFieldDeclarations(text, recordTypeDeclaration)) {
          if (fieldDeclaration.fieldName == fieldName) {
            addOccurrence(file, text, fieldDeclaration.fieldNameStart, fieldDeclaration.fieldNameEnd, true, occurrences, seen)
          }
        }
      }
    }
    return occurrences
  }

  private fun collectNamedSymbolDeclarationRanges(file: VirtualFile, text: String, kind: EpicsSemanticSymbolKind, name: String): List<EpicsSemanticOccurrence> {
    val occurrences = mutableListOf<EpicsSemanticOccurrence>()
    val seen = linkedSetOf<String>()
    collectNamedSymbolOccurrences(file, text, kind, name, includeDeclarations = true, occurrences, seen)
    return occurrences.filter { it.declaration }
  }

  private fun extractMacroReferenceRanges(text: String, macroName: String): List<Pair<Int, Int>> {
    val ranges = mutableListOf<Pair<Int, Int>>()
    for (match in EPICS_MACRO_REFERENCE_REGEX.findAll(maskHashComments(text))) {
      val name = match.groups[1]?.value ?: match.groups[2]?.value ?: continue
      if (name != macroName) {
        continue
      }
      val start = match.range.first + 2
      ranges += start to (start + name.length)
    }
    return ranges
  }

  private fun readText(file: VirtualFile): String? {
    return try {
      String(file.contentsToByteArray(), Charset.forName(file.charset.name()))
    } catch (_: Exception) {
      null
    }
  }

  private fun getLineContext(text: String, offset: Int): Pair<Int, String> {
    val safeOffset = offset.coerceIn(0, text.length)
    val lineNumber = text.take(safeOffset).count { it == '\n' } + 1
    val lineStart = text.lastIndexOf('\n', (safeOffset - 1).coerceAtLeast(0)).let { if (it < 0) 0 else it + 1 }
    val lineEnd = text.indexOf('\n', safeOffset).let { if (it < 0) text.length else it }
    return lineNumber to text.substring(lineStart, lineEnd).trimEnd()
  }

  private fun splitLinesWithOffsets(text: String): List<Pair<String, Int>> {
    val lines = mutableListOf<Pair<String, Int>>()
    var offset = 0
    for (line in text.split('\n')) {
      lines += line to offset
      offset += line.length + 1
    }
    return lines
  }

  private fun isSemanticFile(file: VirtualFile): Boolean {
    return isDatabaseFile(file) ||
      isStartupFile(file) ||
      isSubstitutionsFile(file) ||
      isPvlistFile(file) ||
      isProbeFile(file) ||
      isDbdFile(file) ||
      isSourceFile(file)
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean = file.extension?.lowercase() in DATABASE_EXTENSIONS
  private fun isStartupFile(file: VirtualFile): Boolean = file.extension?.lowercase() in STARTUP_EXTENSIONS || file.name == "st.cmd"
  private fun isSubstitutionsFile(file: VirtualFile): Boolean = file.extension?.lowercase() in SUBSTITUTIONS_EXTENSIONS
  private fun isPvlistFile(file: VirtualFile): Boolean = file.extension?.lowercase() == "pvlist"
  private fun isProbeFile(file: VirtualFile): Boolean = file.extension?.lowercase() == "probe"
  private fun isDbdFile(file: VirtualFile): Boolean = file.extension?.lowercase() == "dbd"
  private fun isSourceFile(file: VirtualFile): Boolean = file.extension?.lowercase() in SOURCE_EXTENSIONS

  private data class DatabaseRecordDeclaration(
    val recordType: String,
    val recordTypeStart: Int,
    val recordTypeEnd: Int,
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
    val recordStart: Int,
    val recordEnd: Int,
  )

  private data class DatabaseFieldDeclaration(
    val fieldName: String,
    val fieldNameStart: Int,
    val fieldNameEnd: Int,
    val value: String,
    val valueStart: Int,
  )

  private data class LinkedRecordCandidate(
    val name: String,
    val start: Int,
    val end: Int,
    val searchNames: Set<String>,
  )

  private data class StartupStatement(
    val kind: String,
    val command: String? = null,
    val name: String? = null,
    val nameStart: Int? = null,
    val nameEnd: Int? = null,
    val macros: String? = null,
    val macroValueStart: Int? = null,
  )

  private data class DbdRecordTypeDeclaration(
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
    val blockStart: Int,
    val blockEnd: Int,
  )

  private data class DbdFieldDeclaration(
    val fieldName: String,
    val fieldNameStart: Int,
    val fieldNameEnd: Int,
  )

  private data class DbdDeviceDeclaration(
    val recordType: String,
    val recordTypeStart: Int,
    val recordTypeEnd: Int,
    val supportName: String,
    val supportNameStart: Int,
    val supportNameEnd: Int,
  )

  private data class NamedDeclaration(
    val kind: EpicsSemanticSymbolKind,
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
  )

  private data class SimpleRecordReference(
    val name: String,
    val start: Int,
    val end: Int,
  )

  private fun extractDatabaseRecordDeclarations(text: String): List<DatabaseRecordDeclaration> {
    val declarations = mutableListOf<DatabaseRecordDeclaration>()
    val sanitized = maskHashComments(text)
    for (match in DATABASE_RECORD_REGEX.findAll(sanitized)) {
      val recordTypeGroup = match.groups[1] ?: continue
      val nameGroup = match.groups[2] ?: continue
      val recordStart = match.range.first
      val recordEnd = findBlockEnd(sanitized, recordStart)
      declarations += DatabaseRecordDeclaration(
        recordType = recordTypeGroup.value,
        recordTypeStart = recordTypeGroup.range.first,
        recordTypeEnd = recordTypeGroup.range.last + 1,
        name = nameGroup.value,
        nameStart = nameGroup.range.first,
        nameEnd = nameGroup.range.last + 1,
        recordStart = recordStart,
        recordEnd = recordEnd,
      )
    }
    return declarations
  }

  private fun extractDatabaseFieldDeclarations(
    text: String,
    recordDeclaration: DatabaseRecordDeclaration,
  ): List<DatabaseFieldDeclaration> {
    val declarations = mutableListOf<DatabaseFieldDeclaration>()
    val recordText = text.substring(recordDeclaration.recordStart, recordDeclaration.recordEnd)
    val sanitized = maskHashComments(recordText)
    for (match in DATABASE_FIELD_REGEX.findAll(sanitized)) {
      val fieldGroup = match.groups[1] ?: match.groups[2] ?: continue
      val valueGroup = match.groups[3] ?: continue
      declarations += DatabaseFieldDeclaration(
        fieldName = fieldGroup.value.uppercase(),
        fieldNameStart = recordDeclaration.recordStart + fieldGroup.range.first,
        fieldNameEnd = recordDeclaration.recordStart + fieldGroup.range.last + 1,
        value = valueGroup.value,
        valueStart = recordDeclaration.recordStart + valueGroup.range.first,
      )
    }
    return declarations
  }

  private fun extractLinkedRecordCandidates(
    fieldValue: String,
    baseOffset: Int,
    tocMacros: Map<String, String>,
  ): List<LinkedRecordCandidate> {
    val leading = fieldValue.indexOfFirst { !it.isWhitespace() }.let { if (it < 0) return emptyList() else it }
    var token = fieldValue.substring(leading).takeWhile { !it.isWhitespace() && it != ',' }
    if (token.startsWith("ca://") || token.startsWith("pva://")) {
      token = token.substringAfter("://")
    }
    if (token.startsWith("@")) {
      return emptyList()
    }
    val recordPortion = token.substringBeforeLast('.').ifBlank { token }
    val start = baseOffset + leading
    val end = start + recordPortion.length
    val names = linkedSetOf<String>()
    if (recordPortion.isNotBlank()) {
      names += recordPortion
      val expanded = expandEpicsMacros(recordPortion, tocMacros)
      if (expanded.isNotBlank()) {
        names += expanded
      }
    }
    return names.map { name ->
      LinkedRecordCandidate(name, start, end, names)
    }
  }

  private fun extractStartupStatements(text: String): List<StartupStatement> {
    val statements = mutableListOf<StartupStatement>()
    for ((lineText, lineStart) in splitLinesWithOffsets(text)) {
      val sanitized = maskHashComments(lineText)
      STARTUP_ENV_SET_REGEX.find(sanitized)?.let { match ->
        val nameGroup = match.groups[1] ?: return@let
        statements += StartupStatement(
          kind = "envSet",
          name = nameGroup.value,
          nameStart = lineStart + nameGroup.range.first,
          nameEnd = lineStart + nameGroup.range.last + 1,
        )
      }
      DB_LOAD_RECORDS_REGEX.find(sanitized)?.let { match ->
        val macrosGroup = match.groups[2]
        statements += StartupStatement(
          kind = "load",
          command = "dbLoadRecords",
          macros = macrosGroup?.value,
          macroValueStart = macrosGroup?.let { lineStart + it.range.first },
        )
      }
    }
    return statements
  }

  private fun extractNamedAssignmentsWithRanges(text: String, absoluteStartOffset: Int): Map<String, Pair<Int, Int>> {
    val ranges = linkedMapOf<String, Pair<Int, Int>>()
    var segmentStart = 0
    var escaped = false
    fun flush(segmentEnd: Int) {
      val segment = text.substring(segmentStart, segmentEnd)
      val match = NAMED_ASSIGNMENT_REGEX.find(segment) ?: return
      val name = match.groups[1]?.value?.trim().orEmpty()
      if (name.isBlank()) {
        return
      }
      val start = absoluteStartOffset + segmentStart + (match.groups[1]?.range?.first ?: return)
      ranges[name] = start to (start + name.length)
    }
    for (index in text.indices) {
      when {
        escaped -> escaped = false
        text[index] == '\\' -> escaped = true
        text[index] == ',' -> {
          flush(index)
          segmentStart = index + 1
        }
      }
    }
    flush(text.length)
    return ranges
  }

  private fun extractDbdRecordTypeDeclarations(text: String): List<DbdRecordTypeDeclaration> {
    val declarations = mutableListOf<DbdRecordTypeDeclaration>()
    for (match in DBD_RECORDTYPE_REGEX.findAll(text)) {
      val nameGroup = match.groups[1] ?: continue
      val start = match.range.first
      val end = findBlockEnd(text, start)
      declarations += DbdRecordTypeDeclaration(
        name = nameGroup.value,
        nameStart = nameGroup.range.first,
        nameEnd = nameGroup.range.last + 1,
        blockStart = start,
        blockEnd = end,
      )
    }
    return declarations
  }

  private fun extractDbdFieldDeclarations(text: String, recordTypeDeclaration: DbdRecordTypeDeclaration): List<DbdFieldDeclaration> {
    val declarations = mutableListOf<DbdFieldDeclaration>()
    val blockText = text.substring(recordTypeDeclaration.blockStart, recordTypeDeclaration.blockEnd)
    for (match in DBD_FIELD_REGEX.findAll(blockText)) {
      val fieldGroup = match.groups[1] ?: continue
      declarations += DbdFieldDeclaration(
        fieldName = fieldGroup.value.uppercase(),
        fieldNameStart = recordTypeDeclaration.blockStart + fieldGroup.range.first,
        fieldNameEnd = recordTypeDeclaration.blockStart + fieldGroup.range.last + 1,
      )
    }
    return declarations
  }

  private fun extractDbdDeviceDeclarations(text: String): List<DbdDeviceDeclaration> {
    val declarations = mutableListOf<DbdDeviceDeclaration>()
    for (match in DBD_DEVICE_REGEX.findAll(text)) {
      val recordTypeGroup = match.groups[1] ?: continue
      val supportNameGroup = match.groups[2] ?: continue
      declarations += DbdDeviceDeclaration(
        recordType = recordTypeGroup.value,
        recordTypeStart = recordTypeGroup.range.first,
        recordTypeEnd = recordTypeGroup.range.last + 1,
        supportName = supportNameGroup.value,
        supportNameStart = supportNameGroup.range.first,
        supportNameEnd = supportNameGroup.range.last + 1,
      )
    }
    return declarations
  }

  private fun extractDbdNamedDeclarations(text: String, kind: EpicsSemanticSymbolKind): List<NamedDeclaration> {
    val keyword = when (kind) {
      EpicsSemanticSymbolKind.DRIVER -> "driver"
      EpicsSemanticSymbolKind.REGISTRAR -> "registrar"
      EpicsSemanticSymbolKind.FUNCTION -> "function"
      EpicsSemanticSymbolKind.VARIABLE -> "variable"
      else -> return emptyList()
    }
    val regex = Regex("""\b$keyword\(\s*([A-Za-z_][A-Za-z0-9_]*)""")
    return regex.findAll(text).mapNotNull { match ->
      val group = match.groups[1] ?: return@mapNotNull null
      NamedDeclaration(kind, group.value, group.range.first, group.range.last + 1)
    }.toList()
  }

  private fun extractSourceNamedDeclarations(text: String): List<NamedDeclaration> {
    val declarations = mutableListOf<NamedDeclaration>()
    for (match in EPICS_EXPORT_ADDRESS_REGEX.findAll(text)) {
      val className = match.groups[1]?.value?.trim().orEmpty()
      val nameGroup = match.groups[2] ?: continue
      val kind = when {
        className.contains("dset", ignoreCase = true) -> EpicsSemanticSymbolKind.DEVICE_SUPPORT
        className.contains("drvet", ignoreCase = true) -> EpicsSemanticSymbolKind.DRIVER
        else -> EpicsSemanticSymbolKind.VARIABLE
      }
      declarations += NamedDeclaration(kind, nameGroup.value, nameGroup.range.first, nameGroup.range.last + 1)
    }
    for (match in EPICS_EXPORT_REGISTRAR_REGEX.findAll(text)) {
      val group = match.groups[1] ?: continue
      declarations += NamedDeclaration(EpicsSemanticSymbolKind.REGISTRAR, group.value, group.range.first, group.range.last + 1)
    }
    for (match in EPICS_REGISTER_FUNCTION_REGEX.findAll(text)) {
      val group = match.groups[1] ?: continue
      declarations += NamedDeclaration(EpicsSemanticSymbolKind.FUNCTION, group.value, group.range.first, group.range.last + 1)
    }
    return declarations
  }

  private fun extractSimpleRecordLineReferences(text: String): List<SimpleRecordReference> {
    return splitLinesWithOffsets(text).mapNotNull { (lineText, lineStart) ->
      val trimmed = lineText.trim()
      if (trimmed.isEmpty() || trimmed.startsWith("#") || PVLIST_MACRO_DEFINITION_REGEX.matches(lineText)) {
        return@mapNotNull null
      }
      val leading = lineText.indexOfFirst { !it.isWhitespace() }.takeIf { it >= 0 } ?: return@mapNotNull null
      val token = lineText.substring(leading).takeWhile { !it.isWhitespace() }
      if (token.isBlank()) {
        return@mapNotNull null
      }
      SimpleRecordReference(token, lineStart + leading, lineStart + leading + token.length)
    }
  }

  private fun extractProbeRecordReferences(text: String): List<SimpleRecordReference> {
    return splitLinesWithOffsets(text).mapNotNull { (lineText, lineStart) ->
      val match = PROBE_RECORD_LINE_REGEX.find(lineText) ?: return@mapNotNull null
      val name = match.groups[1]?.value.orEmpty()
      val start = lineStart + (match.groups[1]?.range?.first ?: return@mapNotNull null)
      SimpleRecordReference(name, start, start + name.length)
    }
  }

  private fun findBlockEnd(text: String, start: Int): Int {
    val openingBrace = text.indexOf('{', start)
    if (openingBrace < 0) {
      return text.length
    }
    var depth = 0
    var inString = false
    var escaped = false
    for (index in openingBrace until text.length) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }
      if (character == '\\') {
        escaped = true
        continue
      }
      if (character == '"') {
        inString = !inString
        continue
      }
      if (inString) {
        continue
      }
      when (character) {
        '{' -> depth += 1
        '}' -> {
          depth -= 1
          if (depth == 0) {
            return index + 1
          }
        }
      }
    }
    return text.length
  }

  private fun maskHashComments(text: String): String {
    val sanitized = StringBuilder(text.length)
    var inString = false
    var escaped = false
    var inComment = false
    for (character in text) {
      if (inComment) {
        if (character == '\n') {
          inComment = false
          sanitized.append(character)
        } else {
          sanitized.append(' ')
        }
        continue
      }
      if (escaped) {
        sanitized.append(character)
        escaped = false
        continue
      }
      if (character == '\\') {
        sanitized.append(character)
        escaped = true
        continue
      }
      if (character == '"') {
        inString = !inString
        sanitized.append(character)
        continue
      }
      if (!inString && character == '#') {
        inComment = true
        sanitized.append(' ')
        continue
      }
      sanitized.append(character)
    }
    return sanitized.toString()
  }

  private fun expandEpicsMacros(value: String, macros: Map<String, String>): String {
    var expanded = value
    repeat(10) {
      val next = EPICS_VALUE_MACRO_REGEX.replace(expanded) { match ->
        val name = match.groups[1]?.value ?: match.groups[3]?.value ?: match.groups[5]?.value ?: return@replace match.value
        val defaultValue = match.groups[2]?.value ?: match.groups[4]?.value
        macros[name] ?: defaultValue ?: match.value
      }
      if (next == expanded) {
        return next
      }
      expanded = next
    }
    return expanded
  }

  private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  private val STARTUP_EXTENSIONS = setOf("cmd", "iocsh")
  private val SUBSTITUTIONS_EXTENSIONS = setOf("sub", "subs", "substitutions")
  private val SOURCE_EXTENSIONS = setOf("c", "cc", "cpp", "cxx", "h", "hh", "hpp", "hxx")
  private val IGNORED_DIRECTORY_NAMES = setOf(".git", ".hg", ".svn", "node_modules", "out", "build", ".idea")

  private val DATABASE_RECORD_REGEX = Regex("""\brecord\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"""")
  private val DATABASE_FIELD_REGEX = Regex("""field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"""")
  private val STARTUP_ENV_SET_REGEX = Regex("""\bepicsEnvSet\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)""")
  private val DB_LOAD_RECORDS_REGEX = Regex("""\bdbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"((?:[^"\\]|\\.)*)")?""")
  private val NAMED_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=""")
  private val DBD_RECORDTYPE_REGEX = Regex("""\brecordtype\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
  private val DBD_FIELD_REGEX = Regex("""\bfield\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,""")
  private val DBD_DEVICE_REGEX = Regex("""\bdevice\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*[^,]+,\s*([A-Za-z_][A-Za-z0-9_]*)\s*,""")
  private val EPICS_EXPORT_ADDRESS_REGEX = Regex("""epicsExportAddress\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
  private val EPICS_EXPORT_REGISTRAR_REGEX = Regex("""epicsExportRegistrar\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
  private val EPICS_REGISTER_FUNCTION_REGEX = Regex("""epicsRegisterFunction\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)""")
  private val PROBE_RECORD_LINE_REGEX = Regex("""^\s*([A-Za-z0-9_:$(){}\-.]+)\b""")
  private val PVLIST_MACRO_DEFINITION_REGEX = Regex("""^(\s*)([A-Za-z_][A-Za-z0-9_]*)\s*=""")
  private val SUBSTITUTIONS_MACRO_DEFINITION_REGEX = Regex("""(?m)(?<=\{|,|\s)([A-Za-z_][A-Za-z0-9_]*)(?=\s*=)""")
  private val EPICS_MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}""")
  private val EPICS_VALUE_MACRO_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
  private val MACRO_NAME_REGEX = Regex("""[A-Za-z_][A-Za-z0-9_]*""")
  private val SYMBOL_NAME_REGEX = Regex("""[A-Za-z_][A-Za-z0-9_]*""")
  private val FIELD_NAME_REGEX = Regex("""[A-Z][A-Z0-9_]*""")
}
