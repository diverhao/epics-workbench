package org.epics.workbench.navigation

import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.epicsBuildModelService
import java.io.File
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.exists
import kotlin.io.path.extension
import kotlin.io.path.isDirectory
import kotlin.io.path.name
import kotlin.io.path.pathString

internal data class EpicsResolvedRecordDefinition(
  val targetFile: VirtualFile,
  val recordName: String,
  val recordType: String,
  val recordStartOffset: Int,
  val recordEndOffset: Int,
  val nameStartOffset: Int,
  val nameEndOffset: Int,
  val line: Int,
)

private data class RecordDeclaration(
  val recordType: String,
  val name: String,
  val nameStart: Int,
  val nameEnd: Int,
  val recordStart: Int,
  val recordEnd: Int,
)

private data class FieldDeclaration(
  val fieldName: String,
  val value: String,
  val valueStart: Int,
  val valueEnd: Int,
)

private data class StartupState(
  var currentDirectory: Path,
  val envVariables: MutableMap<String, String>,
  val searchRoots: List<Path>,
)

private data class StartupLoadStatement(
  val command: String,
  val path: String,
  val macros: String? = null,
)

private data class SubstitutionLoad(
  val templatePath: String,
  val rows: List<Map<String, String>>,
)

private data class SubstitutionBlock(
  val kind: String,
  val templatePath: String?,
  val body: String,
)

private data class BraceSegment(
  val text: String,
)

internal object EpicsRecordResolver {
  internal fun resolveRecordDefinitionInFile(
    hostFile: VirtualFile,
    recordName: String,
    recordType: String? = null,
  ): EpicsResolvedRecordDefinition? {
    return findRecordDefinitionInFile(hostFile, recordName, recordType)
  }

  internal fun resolveRecordDefinitions(
    project: Project,
    hostFile: VirtualFile,
    offset: Int,
  ): List<EpicsResolvedRecordDefinition> {
    val text = readText(hostFile) ?: return emptyList()

    return when {
      isDatabaseFile(hostFile) -> resolveDatabaseRecordDefinitions(project, hostFile, text, offset)
      isStartupFile(hostFile) -> resolveStartupRecordDefinitions(project, hostFile, text, offset)
      else -> emptyList()
    }
  }

  internal fun resolveRecordDefinitionsForName(
    project: Project,
    hostFile: VirtualFile,
    recordName: String,
  ): List<EpicsResolvedRecordDefinition> {
    val candidateNames = extractLinkedRecordCandidates(recordName)
    if (candidateNames.isEmpty()) {
      return emptyList()
    }

    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val envPathsVariables = loadEnvPathsVariables(hostFile.parent?.toNioPath(), releaseVariables)
    val searchDirectories = linkedSetOf<Path>()

    hostFile.parent?.toNioPath()?.let(searchDirectories::add)
    for (root in buildSearchRoots(project, ownerRoot, releaseVariables, envPathsVariables)) {
      searchDirectories.add(root.resolve("db"))
      searchDirectories.add(root.resolve("Db"))
    }

    return resolveLinkedRecordsFromSearchPaths(
      currentFile = hostFile,
      candidateNames = candidateNames,
      searchDirectories = searchDirectories.toList(),
    )
  }

  internal fun collectStartupLoadedRecordNames(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    untilOffset: Int,
  ): List<String> {
    if (!isStartupFile(hostFile)) {
      return emptyList()
    }
    return collectStartupLoadedDefinitions(project, hostFile, text, untilOffset)
      .keys
      .sortedWith(String.CASE_INSENSITIVE_ORDER)
  }

  fun resolveRecordDefinition(
    project: Project,
    hostFile: VirtualFile,
    offset: Int,
  ): EpicsResolvedRecordDefinition? {
    return resolveRecordDefinitions(project, hostFile, offset).firstOrNull()
  }

  private fun resolveDatabaseRecordDefinitions(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    offset: Int,
  ): List<EpicsResolvedRecordDefinition> {
    val linkedValue = getDatabaseLinkedRecordValueAtOffset(text, offset) ?: return emptyList()
    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val envPathsVariables = loadEnvPathsVariables(hostFile.parent?.toNioPath(), releaseVariables)
    val searchDirectories = linkedSetOf<Path>()

    hostFile.parent?.toNioPath()?.let(searchDirectories::add)
    for (root in buildSearchRoots(project, ownerRoot, releaseVariables, envPathsVariables)) {
      searchDirectories.add(root.resolve("db"))
      searchDirectories.add(root.resolve("Db"))
    }

    return resolveLinkedRecordsFromSearchPaths(
      currentFile = hostFile,
      candidateNames = extractLinkedRecordCandidates(linkedValue),
      searchDirectories = searchDirectories.toList(),
    )
  }

  private fun resolveStartupRecordDefinitions(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    offset: Int,
  ): List<EpicsResolvedRecordDefinition> {
    val dbpfRecordName = getStartupDbpfRecordAtOffset(text, offset) ?: return emptyList()
    val definitionsByName = collectStartupLoadedDefinitions(project, hostFile, text, offset)
    val results = mutableListOf<EpicsResolvedRecordDefinition>()
    val seenKeys = mutableSetOf<String>()

    for (candidate in extractLinkedRecordCandidates(dbpfRecordName)) {
      val definitions = definitionsByName[candidate].orEmpty()
      for (definition in definitions) {
        val key = buildRecordDefinitionKey(definition)
        if (seenKeys.add(key)) {
          results += definition
        }
      }
    }

    return results
  }

  private fun resolveLinkedRecordsFromSearchPaths(
    currentFile: VirtualFile,
    candidateNames: List<String>,
    searchDirectories: List<Path>,
  ): List<EpicsResolvedRecordDefinition> {
    if (candidateNames.isEmpty()) {
      return emptyList()
    }

    val results = mutableListOf<EpicsResolvedRecordDefinition>()
    val seenKeys = mutableSetOf<String>()

    for (candidate in candidateNames) {
      findRecordDefinitionInFile(currentFile, candidate)?.let { definition ->
        val key = buildRecordDefinitionKey(definition)
        if (seenKeys.add(key)) {
          results += definition
        }
      }
    }

    val seenPaths = mutableSetOf(currentFile.path)
    for (directory in searchDirectories) {
      for (candidateFile in collectDatabaseFiles(directory, recursive = directory != currentFile.parent?.toNioPath())) {
        val normalizedPath = candidateFile.pathString
        if (!seenPaths.add(normalizedPath)) {
          continue
        }

        val virtualFile = LocalFileSystem.getInstance().findFileByIoFile(candidateFile.toFile()) ?: continue
        for (candidate in candidateNames) {
          findRecordDefinitionInFile(virtualFile, candidate)?.let { definition ->
            val key = buildRecordDefinitionKey(definition)
            if (seenKeys.add(key)) {
              results += definition
            }
          }
        }
      }
    }

    return results
  }

  private fun collectStartupLoadedDefinitions(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    untilOffset: Int,
  ): Map<String, List<EpicsResolvedRecordDefinition>> {
    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val envPathsVariables = loadEnvPathsVariables(hostFile.parent?.toNioPath(), releaseVariables)
    val envVariables = linkedMapOf<String, String>().apply {
      putAll(releaseVariables)
      putAll(envPathsVariables)
      putIfAbsent("TOP", ownerRoot.pathString)
    }
    val state = StartupState(
      currentDirectory = hostFile.parent?.toNioPath() ?: ownerRoot,
      envVariables = envVariables,
      searchRoots = buildSearchRoots(project, ownerRoot, releaseVariables, envPathsVariables),
    )
    val definitionsByName = linkedMapOf<String, MutableList<EpicsResolvedRecordDefinition>>()

    var runningOffset = 0
    for (line in text.split('\n')) {
      val lineStart = runningOffset
      if (lineStart >= untilOffset) {
        break
      }

      val sanitizedLine = maskHashComments(line)

      parseStartupLoadStatement(sanitizedLine)?.let { statement ->
        when (statement.command) {
          "dbLoadRecords" -> {
            val resolvedFile = resolveStartupPath(state, statement.path, EpicsPathKind.DATABASE)
            if (resolvedFile != null) {
              val loadMacros = parseMacroAssignments(statement.macros, state.envVariables)
              addLoadedRecordDefinitions(
                definitionsByName,
                resolvedFile,
                listOf(loadMacros, state.envVariables, System.getenv()),
              )
            }
          }

          "dbLoadTemplate" -> {
            val substitutionsFile = resolveStartupPath(state, statement.path, EpicsPathKind.SUBSTITUTIONS)
            if (substitutionsFile != null) {
              addLoadedSubstitutionDefinitions(
                definitionsByName,
                substitutionsFile,
                state,
              )
            }
          }
        }
      }

      applyStartupLine(sanitizedLine, state)
      runningOffset += line.length + 1
    }

    return definitionsByName
  }

  private fun addLoadedSubstitutionDefinitions(
    definitionsByName: MutableMap<String, MutableList<EpicsResolvedRecordDefinition>>,
    substitutionsFile: VirtualFile,
    state: StartupState,
  ) {
    val substitutionsText = readText(substitutionsFile) ?: return
    for (load in parseSubstitutionLoads(substitutionsText)) {
      val templateFile = resolveTemplatePath(state, substitutionsFile, load.templatePath) ?: continue
      val rows = if (load.rows.isEmpty()) listOf(emptyMap()) else load.rows
      for (rowMacros in rows) {
        addLoadedRecordDefinitions(
          definitionsByName,
          templateFile,
          listOf(rowMacros, state.envVariables, System.getenv()),
        )
      }
    }
  }

  private fun addLoadedRecordDefinitions(
    definitionsByName: MutableMap<String, MutableList<EpicsResolvedRecordDefinition>>,
    databaseFile: VirtualFile,
    macroSources: List<Map<String, String>>,
  ) {
    val text = readText(databaseFile) ?: return
    for (declaration in extractRecordDeclarations(text)) {
      val expandedName = expandEpicsValue(declaration.name, macroSources)
      if (expandedName.isBlank()) {
        continue
      }

      definitionsByName.getOrPut(expandedName) { mutableListOf() }.add(
        EpicsResolvedRecordDefinition(
          targetFile = databaseFile,
          recordName = expandedName,
          recordType = declaration.recordType,
          recordStartOffset = declaration.recordStart,
          recordEndOffset = declaration.recordEnd,
          nameStartOffset = declaration.nameStart,
          nameEndOffset = declaration.nameEnd,
          line = getLineNumberAtOffset(text, declaration.recordStart),
        ),
      )
    }
  }

  private fun resolveTemplatePath(
    state: StartupState,
    substitutionsFile: VirtualFile,
    templatePath: String,
  ): VirtualFile? {
    resolveStartupPath(state, templatePath, EpicsPathKind.DATABASE)?.let { return it }

    val expanded = expandEpicsValue(templatePath, listOf(state.envVariables, System.getenv()))
    val fallback = substitutionsFile.parent?.toNioPath()?.resolve(expanded)?.normalize() ?: return null
    if (!fallback.exists() || fallback.isDirectory()) {
      return null
    }
    return LocalFileSystem.getInstance().findFileByIoFile(fallback.toFile())
  }

  private fun resolveStartupPath(
    state: StartupState,
    rawPath: String,
    kind: EpicsPathKind,
  ): VirtualFile? {
    val expandedPath = expandEpicsValue(rawPath, listOf(state.envVariables, System.getenv()))
    if (expandedPath.isBlank()) {
      return null
    }

    val directCandidate = resolveAbsoluteOrRelative(state.currentDirectory, expandedPath)
    if (directCandidate.exists() && !directCandidate.isDirectory()) {
      return LocalFileSystem.getInstance().findFileByIoFile(directCandidate.toFile())
    }

    val basename = runCatching { Path.of(expandedPath).fileName?.toString().orEmpty() }
      .getOrElse { expandedPath.substringAfterLast('/').substringAfterLast('\\') }

    for (root in state.searchRoots) {
      for (candidate in candidatePathsForKind(root, expandedPath, basename, kind)) {
        if (candidate.exists() && !candidate.isDirectory()) {
          return LocalFileSystem.getInstance().findFileByIoFile(candidate.toFile())
        }
      }
    }

    for (root in state.searchRoots) {
      searchPreferredDirectories(root, basename, kind)?.let { return it }
    }

    return null
  }

  private fun applyStartupLine(line: String, state: StartupState) {
    val epicsEnvMatch = STARTUP_ENV_SET_REGEX.find(line)
    if (epicsEnvMatch != null) {
      val name = epicsEnvMatch.groups[1]?.value?.trim().orEmpty()
      val rawValue = epicsEnvMatch.groups[2]?.value.orEmpty()
      if (name.isNotEmpty()) {
        state.envVariables[name] = expandEpicsValue(rawValue, listOf(state.envVariables, System.getenv()))
      }
    }

    val cdMatch = STARTUP_CD_REGEX.find(line)
    if (cdMatch != null) {
      val rawDirectory = cdMatch.groups[1]?.value ?: cdMatch.groups[2]?.value ?: ""
      val expandedDirectory = expandEpicsValue(rawDirectory, listOf(state.envVariables, System.getenv()))
      val resolvedDirectory = resolveAbsoluteOrRelative(state.currentDirectory, expandedDirectory)
      if (resolvedDirectory.exists() && resolvedDirectory.isDirectory()) {
        state.currentDirectory = resolvedDirectory.normalize()
      }
    }
  }

  private fun parseStartupLoadStatement(line: String): StartupLoadStatement? {
    val dbLoadRecordsMatch = DB_LOAD_RECORDS_REGEX.find(line)
    if (dbLoadRecordsMatch != null) {
      return StartupLoadStatement(
        command = "dbLoadRecords",
        path = dbLoadRecordsMatch.groups[1]?.value.orEmpty(),
        macros = dbLoadRecordsMatch.groups[2]?.value,
      )
    }

    val dbLoadTemplateMatch = DB_LOAD_TEMPLATE_REGEX.find(line)
    if (dbLoadTemplateMatch != null) {
      return StartupLoadStatement(
        command = "dbLoadTemplate",
        path = dbLoadTemplateMatch.groups[1]?.value.orEmpty(),
      )
    }

    return null
  }

  private fun parseSubstitutionLoads(text: String): List<SubstitutionLoad> {
    val loads = mutableListOf<SubstitutionLoad>()
    var globalMacros = emptyMap<String, String>()

    for (block in extractSubstitutionBlocks(text)) {
      when (block.kind) {
        "global" -> {
          globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body))
        }

        "file" -> {
          val templatePath = block.templatePath ?: continue
          val rows = parseSubstitutionRows(block.body).map { row ->
            mergeMacroMaps(globalMacros, row)
          }
          loads += SubstitutionLoad(templatePath, rows)
        }
      }
    }

    return loads
  }

  private fun extractSubstitutionBlocks(text: String): List<SubstitutionBlock> {
    val blocks = mutableListOf<SubstitutionBlock>()
    val blockPattern = Regex("""(?:^|\n)\s*(global|file)(?:\s+("(?:[^"\\]|\\.)*"|[^\s{]+))?\s*\{""")
    var searchIndex = 0

    while (searchIndex < text.length) {
      val match = blockPattern.find(text, searchIndex) ?: break
      val braceIndex = text.indexOf('{', match.range.first)
      if (braceIndex < 0) {
        break
      }

      val blockEnd = findBalancedBlockEnd(text, braceIndex)
      val bodyStart = braceIndex + 1
      val bodyEnd = (blockEnd - 1).coerceAtLeast(bodyStart)
      val rawTemplatePath = match.groups[2]?.value
      val templatePath = rawTemplatePath?.trim()?.removeSurrounding("\"")
      blocks += SubstitutionBlock(
        kind = match.groups[1]?.value.orEmpty(),
        templatePath = templatePath,
        body = text.substring(bodyStart, bodyEnd),
      )
      searchIndex = blockEnd
    }

    return blocks
  }

  private fun parseSubstitutionRows(body: String): List<Map<String, String>> {
    val segments = extractTopLevelBraceSegments(body)
    if (segments.isEmpty()) {
      return emptyList()
    }

    if (body.trimStart().startsWith("pattern")) {
      val columns = tokenizeSubstitutionValues(segments.first().text)
      return segments.drop(1).map { segment ->
        createPatternMacroAssignments(columns, tokenizeSubstitutionValues(segment.text))
      }.filter { it.isNotEmpty() }
    }

    return segments.map { segment ->
      extractNamedAssignments(segment.text)
    }.filter { it.isNotEmpty() }
  }

  private fun extractTopLevelBraceSegments(text: String): List<BraceSegment> {
    val segments = mutableListOf<BraceSegment>()
    var searchIndex = 0
    while (searchIndex < text.length) {
      val braceIndex = text.indexOf('{', searchIndex)
      if (braceIndex < 0) {
        break
      }
      val blockEnd = findBalancedBlockEnd(text, braceIndex)
      val contentStart = braceIndex + 1
      val contentEnd = (blockEnd - 1).coerceAtLeast(contentStart)
      segments += BraceSegment(text.substring(contentStart, contentEnd))
      searchIndex = blockEnd
    }
    return segments
  }

  private fun createPatternMacroAssignments(
    columns: List<String>,
    values: List<String>,
  ): Map<String, String> {
    val assignments = linkedMapOf<String, String>()
    columns.forEachIndexed { index, name ->
      if (name.isBlank()) {
        return@forEachIndexed
      }
      assignments[name] = values.getOrElse(index) { "" }
    }
    return assignments
  }

  private fun tokenizeSubstitutionValues(text: String): List<String> {
    val tokens = mutableListOf<String>()
    val tokenRegex = Regex("""\"((?:[^"\\]|\\.)*)\"|([^,\s{}]+)""")
    tokenRegex.findAll(text).forEach { match ->
      tokens += (match.groups[1]?.value ?: match.groups[2]?.value ?: "").trim()
    }
    return tokens
  }

  private fun mergeMacroMaps(
    left: Map<String, String>,
    right: Map<String, String>,
  ): Map<String, String> {
    return linkedMapOf<String, String>().apply {
      putAll(left)
      putAll(right)
    }
  }

  private fun parseMacroAssignments(
    rawAssignments: String?,
    envVariables: Map<String, String>,
  ): Map<String, String> {
    if (rawAssignments.isNullOrBlank()) {
      return emptyMap()
    }

    val assignments = linkedMapOf<String, String>()
    for ((name, rawValue) in extractNamedAssignments(rawAssignments)) {
      assignments[name] = expandEpicsValue(rawValue, listOf(envVariables, System.getenv()))
    }
    return assignments
  }

  private fun extractNamedAssignments(rawText: String): Map<String, String> {
    if (rawText.isBlank()) {
      return emptyMap()
    }

    val assignments = linkedMapOf<String, String>()
    for (token in splitTopLevelCommaSeparated(rawText)) {
      val separatorIndex = token.indexOf('=')
      if (separatorIndex <= 0) {
        continue
      }

      val name = token.substring(0, separatorIndex).trim()
      if (name.isBlank()) {
        continue
      }

      val value = token.substring(separatorIndex + 1).trim().removeSurrounding("\"")
      assignments[name] = value
    }
    return assignments
  }

  private fun splitTopLevelCommaSeparated(text: String): List<String> {
    val parts = mutableListOf<String>()
    val current = StringBuilder()
    var inString = false
    var escaped = false

    for (character in text) {
      when {
        escaped -> {
          current.append(character)
          escaped = false
        }

        character == '\\' -> {
          current.append(character)
          escaped = true
        }

        character == '"' -> {
          current.append(character)
          inString = !inString
        }

        character == ',' && !inString -> {
          parts += current.toString()
          current.setLength(0)
        }

        else -> current.append(character)
      }
    }

    if (current.isNotEmpty()) {
      parts += current.toString()
    }

    return parts
  }

  private fun getDatabaseLinkedRecordValueAtOffset(text: String, offset: Int): String? {
    for (recordDeclaration in extractRecordDeclarations(text)) {
      if (offset < recordDeclaration.recordStart || offset > recordDeclaration.recordEnd) {
        continue
      }

      for (fieldDeclaration in extractFieldDeclarationsInRecord(text, recordDeclaration)) {
        if (offset < fieldDeclaration.valueStart || offset > fieldDeclaration.valueEnd) {
          continue
        }
        if (!isLinkField(fieldDeclaration.fieldName)) {
          continue
        }
        return fieldDeclaration.value
      }
    }

    return null
  }

  private fun getStartupDbpfRecordAtOffset(text: String, offset: Int): String? {
    val sanitizedText = maskHashComments(text)
    val regex = Regex("""dbpf\(\s*"((?:[^"\\]|\\.)*)"""")
    for (match in regex.findAll(sanitizedText)) {
      val value = match.groups[1]?.value.orEmpty()
      val valueStart = match.range.first + match.value.length - value.length - 1
      val valueEnd = valueStart + value.length
      if (offset in valueStart..valueEnd) {
        return value
      }
    }
    return null
  }

  private fun findRecordDefinitionInFile(
    virtualFile: VirtualFile,
    recordName: String,
    recordType: String? = null,
  ): EpicsResolvedRecordDefinition? {
    val text = readText(virtualFile) ?: return null
    for (declaration in extractRecordDeclarations(text)) {
      if (declaration.name != recordName) {
        continue
      }
      if (recordType != null && declaration.recordType != recordType) {
        continue
      }

      return EpicsResolvedRecordDefinition(
        targetFile = virtualFile,
        recordName = declaration.name,
        recordType = declaration.recordType,
        recordStartOffset = declaration.recordStart,
        recordEndOffset = declaration.recordEnd,
        nameStartOffset = declaration.nameStart,
        nameEndOffset = declaration.nameEnd,
        line = getLineNumberAtOffset(text, declaration.recordStart),
      )
    }
    return null
  }

  private fun extractRecordDeclarations(text: String): List<RecordDeclaration> {
    val declarations = mutableListOf<RecordDeclaration>()
    val sanitizedText = maskHashComments(text)
    val regex = Regex("""\brecord\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"""")

    for (match in regex.findAll(sanitizedText)) {
      val recordType = match.groups[1]?.value.orEmpty()
      val recordName = match.groups[2]?.value.orEmpty()
      val prefixLength = Regex("""record\(\s*[A-Za-z0-9_]+\s*,\s*"""").find(match.value)?.value?.length ?: continue
      val recordStart = match.range.first
      val recordEnd = findRecordBlockEnd(sanitizedText, recordStart)
      declarations += RecordDeclaration(
        recordType = recordType,
        name = recordName,
        nameStart = match.range.first + prefixLength,
        nameEnd = match.range.first + prefixLength + recordName.length,
        recordStart = recordStart,
        recordEnd = recordEnd,
      )
    }

    return declarations
  }

  private fun extractFieldDeclarationsInRecord(
    text: String,
    recordDeclaration: RecordDeclaration,
  ): List<FieldDeclaration> {
    val declarations = mutableListOf<FieldDeclaration>()
    val sanitizedText = maskHashComments(text)
    val recordText = sanitizedText.substring(recordDeclaration.recordStart, recordDeclaration.recordEnd)
    val regex = Regex("""field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"""")

    for (match in regex.findAll(recordText)) {
      val fieldName = match.groups[1]?.value ?: match.groups[2]?.value ?: continue
      val valuePrefixLength = Regex("""field\(\s*(?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+)\s*,\s*"""").find(match.value)?.value?.length ?: continue
      declarations += FieldDeclaration(
        fieldName = fieldName,
        value = match.groups[3]?.value.orEmpty(),
        valueStart = recordDeclaration.recordStart + match.range.first + valuePrefixLength,
        valueEnd = recordDeclaration.recordStart + match.range.first + valuePrefixLength + (match.groups[3]?.value?.length ?: 0),
      )
    }

    return declarations
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

  private fun findRecordBlockEnd(text: String, recordStart: Int): Int {
    val openingBraceIndex = text.indexOf('{', recordStart)
    if (openingBraceIndex < 0) {
      return recordStart
    }

    return findBalancedBlockEnd(text, openingBraceIndex)
  }

  private fun findBalancedBlockEnd(text: String, openingBraceIndex: Int): Int {
    var depth = 0
    var inString = false
    var escaped = false
    var inComment = false

    for (index in openingBraceIndex until text.length) {
      val character = text[index]

      if (inComment) {
        if (character == '\n') {
          inComment = false
        }
        continue
      }

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

      if (character == '#') {
        inComment = true
        continue
      }

      if (character == '{') {
        depth += 1
        continue
      }

      if (character == '}') {
        depth -= 1
        if (depth == 0) {
          return index + 1
        }
      }
    }

    return text.length
  }

  private fun getLineNumberAtOffset(text: String, offset: Int): Int {
    return text.take(offset.coerceIn(0, text.length)).count { it == '\n' } + 1
  }

  private fun isLinkField(fieldName: String): Boolean {
    if (fieldName.isBlank()) {
      return false
    }
    val normalizedFieldName = fieldName.uppercase()

    return when {
      normalizedFieldName in FALLBACK_LINK_FIELDS -> true
      Regex("""^INP[A-U]$""").matches(normalizedFieldName) -> true
      Regex("""^OUT[A-U]$""").matches(normalizedFieldName) -> true
      Regex("""^DOL[0-9A-F]$""").matches(normalizedFieldName) -> true
      Regex("""^LNK[0-9A-F]$""").matches(normalizedFieldName) -> true
      else -> false
    }
  }

  private fun extractLinkedRecordCandidates(fieldValue: String): List<String> {
    val trimmedValue = fieldValue.trim()
    if (trimmedValue.isBlank() || trimmedValue.startsWith("@")) {
      return emptyList()
    }

    val firstToken = trimmedValue.split(Regex("""\s+""")).firstOrNull().orEmpty()
    if (firstToken.isBlank()) {
      return emptyList()
    }

    val candidates = linkedSetOf<String>()
    val normalized = normalizeLinkedRecordCandidate(firstToken)
    if (normalized.isNotBlank()) {
      candidates += normalized
    }

    val lastDotIndex = firstToken.lastIndexOf('.')
    if (lastDotIndex > 0) {
      val suffix = firstToken.substring(lastDotIndex + 1)
      if (Regex("""^[A-Z0-9_]+$""").matches(suffix)) {
        candidates += normalizeLinkedRecordCandidate(firstToken.substring(0, lastDotIndex))
      }
    }

    return candidates.filter { it.isNotBlank() }
  }

  private fun normalizeLinkedRecordCandidate(candidate: String): String {
    return candidate.trim().replace(Regex("""[),;]+$"""), "")
  }

  private fun collectDatabaseFiles(
    directory: Path,
    recursive: Boolean,
  ): Sequence<Path> {
    if (!directory.exists() || !directory.isDirectory()) {
      return emptySequence()
    }

    return if (!recursive) {
      directory.toFile().listFiles()
        .orEmpty()
        .asSequence()
        .filter { it.isFile }
        .map { it.toPath() }
        .filter(::isDatabasePath)
        .sortedBy { it.pathString }
    } else {
      Files.walk(directory).use { stream ->
        stream
          .filter { path -> Files.isRegularFile(path) && isDatabasePath(path) }
          .iterator()
          .asSequence()
          .toList()
          .sortedBy { it.pathString }
          .asSequence()
      }
    }
  }

  private fun buildRecordDefinitionKey(definition: EpicsResolvedRecordDefinition): String {
    return "${definition.targetFile.path}:${definition.recordName}:${definition.recordStartOffset}:${definition.recordEndOffset}"
  }

  private fun isDatabasePath(path: Path): Boolean {
    return path.extension.lowercase() in DATABASE_EXTENSIONS
  }

  private fun readText(file: VirtualFile): String? {
    return try {
      String(file.contentsToByteArray(), file.charset)
    } catch (_: Exception) {
      null
    }
  }

  private fun findOwningEpicsRoot(project: Project, hostFile: VirtualFile): Path {
    var current = hostFile.parent?.toNioPath()
    while (current != null) {
      if (current.resolve("configure").resolve("RELEASE").exists()) {
        return current
      }
      current = current.parent
    }

    ProjectRootManager.getInstance(project).contentRoots.firstOrNull()?.let { contentRoot ->
      val path = contentRoot.toNioPath()
      if (path.resolve("configure").resolve("RELEASE").exists()) {
        return path
      }
    }

    return Path.of(project.basePath ?: hostFile.parent?.path ?: ".").normalize()
  }

  private fun buildSearchRoots(
    project: Project?,
    ownerRoot: Path,
    releaseVariables: Map<String, String>,
    envVariables: Map<String, String>,
  ): List<Path> {
    project?.let { currentProject ->
      return currentProject.epicsBuildModelService().collectSearchRoots(ownerRoot, releaseVariables, envVariables)
    }

    val roots = linkedSetOf<Path>()
    roots.add(ownerRoot.normalize())
    (releaseVariables.values + envVariables.values).forEach { rawValue ->
      val candidate = runCatching { Path.of(rawValue) }.getOrNull() ?: return@forEach
      if (candidate.exists() && candidate.isDirectory()) {
        roots.add(candidate.normalize())
      }
    }
    return roots.toList()
  }

  private fun loadReleaseVariables(ownerRoot: Path): Map<String, String> {
    val configureDirectory = ownerRoot.resolve("configure")
    val releaseFiles = listOf(
      configureDirectory.resolve("RELEASE"),
      configureDirectory.resolve("RELEASE.local"),
    )
    val rawValues = linkedMapOf("TOP" to ownerRoot.pathString)
    for (releaseFile in releaseFiles) {
      if (!releaseFile.exists()) {
        continue
      }

      releaseFile.toFile().readLines().forEach { line ->
        val match = RELEASE_ASSIGNMENT_REGEX.find(line) ?: return@forEach
        rawValues[match.groups[1]?.value.orEmpty()] = match.groups[2]?.value?.trim().orEmpty()
      }
    }

    val resolved = linkedMapOf<String, String>()
    val resolving = mutableSetOf<String>()

    fun resolve(name: String): String? {
      resolved[name]?.let { return it }
      if (name in resolving) {
        return rawValues[name]
      }
      val rawValue = rawValues[name] ?: return null
      resolving += name
      val expanded = expandEpicsValue(
        rawValue,
        listOf(rawValues, System.getenv()),
      ) { nestedName -> resolve(nestedName) }
      val normalized = if (expanded.isNotBlank() && !Path.of(expanded).isAbsolute) {
        ownerRoot.resolve(expanded).normalize().pathString
      } else {
        expanded
      }
      resolving -= name
      resolved[name] = normalized
      return normalized
    }

    rawValues.keys.forEach { resolve(it) }
    return resolved
  }

  private fun loadEnvPathsVariables(
    startupDirectory: Path?,
    baseVariables: Map<String, String>,
  ): Map<String, String> {
    val directory = startupDirectory ?: return emptyMap()
    val candidateFiles = directory.toFile().listFiles { file ->
      file.isFile && (file.name == "envPaths" || file.name.startsWith("envPaths."))
    }?.sortedBy { if (it.name == "envPaths") 0 else 1 }.orEmpty()
    if (candidateFiles.isEmpty()) {
      return emptyMap()
    }

    val values = linkedMapOf<String, String>()
    for (candidate in candidateFiles) {
      candidate.readLines().forEach { line ->
        val match = STARTUP_ENV_SET_REGEX.find(line) ?: return@forEach
        val name = match.groups[1]?.value?.trim().orEmpty()
        val rawValue = match.groups[2]?.value.orEmpty()
        if (name.isEmpty()) {
          return@forEach
        }
        values[name] = expandEpicsValue(rawValue, listOf(values, baseVariables, System.getenv()))
      }
    }
    return values
  }

  private fun expandEpicsValue(
    rawValue: String,
    macroSources: List<Map<String, String>>,
    fallbackResolver: ((String) -> String?)? = null,
  ): String {
    var expanded = rawValue
    repeat(10) {
      val next = EPICS_VARIABLE_REGEX.replace(expanded) { match ->
        val name = match.groups[1]?.value
          ?: match.groups[3]?.value
          ?: match.groups[5]?.value
          ?: ""
        val defaultValue = match.groups[2]?.value ?: match.groups[4]?.value
        macroSources.firstNotNullOfOrNull { it[name] } ?: fallbackResolver?.invoke(name) ?: defaultValue ?: match.value
      }
      if (next == expanded) {
        return next.trim().trim('"')
      }
      expanded = next
    }
    return expanded.trim().trim('"')
  }

  private fun resolveAbsoluteOrRelative(baseDirectory: Path, rawPath: String): Path {
    val candidate = runCatching { Path.of(rawPath) }.getOrNull()
    return if (candidate != null && candidate.isAbsolute) {
      candidate.normalize()
    } else {
      baseDirectory.resolve(rawPath).normalize()
    }
  }

  private fun candidatePathsForKind(
    searchRoot: Path,
    expandedPath: String,
    basename: String,
    kind: EpicsPathKind,
  ): List<Path> {
    if (basename.isBlank()) {
      return emptyList()
    }

    return when (kind) {
      EpicsPathKind.DATABASE, EpicsPathKind.SUBSTITUTIONS -> {
        val normalized = expandedPath.replace('\\', '/')
        buildList {
          if (normalized.startsWith("db/") || normalized.startsWith("Db/")) {
            add(searchRoot.resolve(normalized))
          }
          add(searchRoot.resolve("db").resolve(basename))
          add(searchRoot.resolve("Db").resolve(basename))
          add(searchRoot.resolve(normalized))
        }
      }

      EpicsPathKind.PROTOCOL -> emptyList()

      EpicsPathKind.DBD -> buildList {
        val normalized = expandedPath.replace('\\', '/')
        if (normalized.startsWith("dbd/")) {
          add(searchRoot.resolve(normalized))
        }
        add(searchRoot.resolve("dbd").resolve(basename))
        add(searchRoot.resolve(normalized))
      }

      EpicsPathKind.LIBRARY -> emptyList()
    }
  }

  private fun searchPreferredDirectories(
    root: Path,
    basename: String,
    kind: EpicsPathKind,
  ): VirtualFile? {
    val directories = when (kind) {
      EpicsPathKind.DATABASE, EpicsPathKind.SUBSTITUTIONS -> listOf(root.resolve("db"), root.resolve("Db"))
      EpicsPathKind.PROTOCOL -> emptyList()
      EpicsPathKind.DBD -> listOf(root.resolve("dbd"))
      EpicsPathKind.LIBRARY -> emptyList()
    }

    for (directory in directories) {
      if (!directory.exists() || !directory.isDirectory()) {
        continue
      }
      val found = findFileRecursively(directory.toFile(), setOf(basename))
      if (found != null) {
        return LocalFileSystem.getInstance().findFileByIoFile(found)
      }
    }
    return null
  }

  private fun findFileRecursively(
    directory: File,
    candidateNames: Set<String>,
  ): File? {
    val queue = ArrayDeque<File>()
    queue += directory
    while (queue.isNotEmpty()) {
      val next = queue.removeFirst()
      val children = next.listFiles().orEmpty().sortedBy { if (it.isDirectory) 0 else 1 }
      for (child in children) {
        if (child.isDirectory) {
          queue += child
          continue
        }
        if (candidateNames.contains(child.name)) {
          return child
        }
      }
    }
    return null
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in DATABASE_EXTENSIONS
  }

  private fun isStartupFile(file: VirtualFile): Boolean {
    val extension = file.extension?.lowercase()
    return extension == "cmd" || extension == "iocsh" || file.name == "st.cmd"
  }

  private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
  private val FALLBACK_LINK_FIELDS = setOf("INP", "OUT", "FLNK", "SELL", "DOL", "SDIS", "SIOL", "TSEL")
  private val STARTUP_ENV_SET_REGEX = Regex("""\bepicsEnvSet\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)""")
  private val STARTUP_CD_REGEX = Regex("""^\s*cd\s+(?:"([^"]+)"|([^\s#]+))""")
  private val DB_LOAD_RECORDS_REGEX = Regex("""\bdbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"([^"\n]*)")?\s*\)""")
  private val DB_LOAD_TEMPLATE_REGEX = Regex("""\bdbLoadTemplate\(\s*"([^"\n]+)"\s*\)""")
  private val RELEASE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_.-]*)\s*=\s*(.*?)\s*(?:#.*)?$""")
  private val EPICS_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
}
