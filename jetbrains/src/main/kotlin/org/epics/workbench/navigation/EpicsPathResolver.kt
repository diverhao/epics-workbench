package org.epics.workbench.navigation

import com.intellij.openapi.project.Project
import com.intellij.openapi.roots.ProjectRootManager
import com.intellij.openapi.vfs.LocalFileSystem
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import java.io.File
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.extension
import kotlin.io.path.exists
import kotlin.io.path.isDirectory
import kotlin.io.path.name
import kotlin.io.path.pathString

internal enum class EpicsPathKind {
  DATABASE,
  SUBSTITUTIONS,
  PROTOCOL,
  DBD,
  LIBRARY,
}

internal data class EpicsResolvedReference(
  val targetFile: VirtualFile,
  val rawPath: String,
  val kind: EpicsPathKind,
)

internal data class EpicsPathCompletionCandidate(
  val insertPath: String,
  val detail: String,
  val isDirectory: Boolean,
  val absolutePath: Path? = null,
)

private data class EpicsReferenceContext(
  val kind: EpicsPathKind,
  val rawPath: String,
)

private data class StartupExecutionState(
  var currentDirectory: Path,
  val variables: MutableMap<String, String>,
  val searchRoots: List<Path>,
)

object EpicsPathResolver {
  internal fun collectStartupDbLoadRecordsCompletionCandidates(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    untilOffset: Int,
    rawPartial: String,
  ): List<EpicsPathCompletionCandidate> {
    if (!isStartupFile(hostFile)) {
      return emptyList()
    }

    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val state = createStartupExecutionState(hostFile, ownerRoot)
    applyStartupStateUntilOffset(text, untilOffset, state)
    return collectDbLoadRecordsCompletionCandidates(state, rawPartial)
  }

  internal fun resolveStartupDatabasePath(
    hostFile: VirtualFile,
    ownerRoot: Path,
    text: String,
    untilOffset: Int,
    rawPath: String,
  ): Path? {
    if (!isStartupFile(hostFile)) {
      return null
    }

    val state = createStartupExecutionState(hostFile, ownerRoot)
    applyStartupStateUntilOffset(text, untilOffset, state)
    val expandedPath = expandEpicsValue(rawPath, state.variables).trim()
    if (expandedPath.isBlank()) {
      return null
    }

    val directCandidate = resolveAbsoluteOrRelative(state.currentDirectory, expandedPath)
    if (directCandidate.exists() && !directCandidate.isDirectory()) {
      return directCandidate.normalize()
    }

    val basename = runCatching { Path.of(expandedPath).fileName?.toString().orEmpty() }
      .getOrElse { expandedPath.substringAfterLast('/').substringAfterLast('\\') }
    for (root in state.searchRoots) {
      candidatePathsForKind(root, expandedPath, basename, EpicsPathKind.DATABASE).forEach { candidate ->
        if (candidate.exists() && !candidate.isDirectory()) {
          return candidate.normalize()
        }
      }
    }

    for (root in state.searchRoots) {
      searchPreferredDirectories(
        root,
        basename,
        EpicsReferenceContext(EpicsPathKind.DATABASE, rawPath),
      )?.let { return it.targetFile.toNioPath() }
    }

    return null
  }

  internal fun resolveSubstitutionsReferences(
    project: Project,
    hostFile: VirtualFile,
    offset: Int,
  ): List<EpicsResolvedReference> {
    if (!isSubstitutionsFile(hostFile)) {
      return emptyList()
    }

    val text = runCatching { hostFile.inputStream.bufferedReader().use { it.readText() } }.getOrNull()
      ?: return emptyList()
    val context = getSubstitutionsReferenceAtOffset(text, offset) ?: return emptyList()
    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    return collectSubstitutionsReferences(hostFile, ownerRoot, context)
  }

  internal fun collectStartupPathCompletionCandidates(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    untilOffset: Int,
    rawPartial: String,
    kind: EpicsPathKind,
  ): List<EpicsPathCompletionCandidate> {
    if (!isStartupFile(hostFile)) {
      return emptyList()
    }

    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val state = createStartupExecutionState(hostFile, ownerRoot)
    applyStartupStateUntilOffset(text, untilOffset, state)
    return collectPathCompletionCandidates(state, rawPartial, kind)
  }

  internal fun resolveReference(project: Project, hostFile: VirtualFile, offset: Int): EpicsResolvedReference? {
    val text = hostFile.inputStream.bufferedReader().use { it.readText() }
    val context = when {
      isDatabaseFile(hostFile) -> getDatabaseReferenceAtOffset(text, offset)
      isMakefile(hostFile) -> getMakefileReferenceAtOffset(text, offset)
      isStartupFile(hostFile) -> getStartupReferenceAtOffset(project, hostFile, text, offset)
      isSubstitutionsFile(hostFile) -> getSubstitutionsReferenceAtOffset(text, offset)
      else -> null
    } ?: return null

    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    return when {
      isDatabaseFile(hostFile) -> resolveDatabaseReference(text, offset, ownerRoot, context)
      isMakefile(hostFile) -> resolveMakefileReference(project, hostFile, ownerRoot, context)
      isStartupFile(hostFile) -> resolveStartupReference(project, hostFile, text, offset, ownerRoot, context)
      isSubstitutionsFile(hostFile) -> resolveSubstitutionsReference(hostFile, ownerRoot, context)
      else -> null
    }
  }

  fun resolveReferencedFile(project: Project, hostFile: VirtualFile, offset: Int): VirtualFile? {
    return resolveReference(project, hostFile, offset)?.targetFile
  }

  private fun resolveStartupReference(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    offset: Int,
    ownerRoot: Path,
    initialContext: EpicsReferenceContext,
  ): EpicsResolvedReference? {
    val state = createStartupExecutionState(hostFile, ownerRoot)
    val lineMatches = collectStartupLineMatches(text)
    for (match in lineMatches) {
      if (match.startOffset > offset) {
        break
      }
      if (offset in match.rangeStart until match.rangeEnd) {
        return resolvePath(
          project,
          currentDirectory = state.currentDirectory,
          searchRoots = state.searchRoots,
          context = match.context,
          expansionVariables = state.variables,
        )
      }
      applyStartupLine(text.substring(match.lineStart, match.lineEnd), state)
    }
    return resolvePath(
      project,
      currentDirectory = state.currentDirectory,
      searchRoots = state.searchRoots,
      context = initialContext,
      expansionVariables = state.variables,
    )
  }

  private fun resolveDatabaseReference(
    text: String,
    offset: Int,
    ownerRoot: Path,
    context: EpicsReferenceContext,
  ): EpicsResolvedReference? {
    return when (context.kind) {
      EpicsPathKind.PROTOCOL -> resolveStreamProtocolReference(text, offset, ownerRoot, context)
      else -> null
    }
  }

  private fun resolveMakefileReference(
    project: Project,
    hostFile: VirtualFile,
    ownerRoot: Path,
    context: EpicsReferenceContext,
  ): EpicsResolvedReference? {
    val hostDirectory = hostFile.parent?.toNioPath() ?: ownerRoot
    val searchRoots = buildSearchRoots(ownerRoot, emptyMap(), emptyMap())

    if (context.kind == EpicsPathKind.DATABASE) {
      resolveFileCandidate(hostDirectory.resolve(context.rawPath), context)?.let { return it }
      toSubstitutionsSiblingPath(context.rawPath)?.let { substitutionsRelativePath ->
        val substitutionsCandidate = hostDirectory.resolve(substitutionsRelativePath)
        resolveFileCandidate(
          substitutionsCandidate,
          context.copy(kind = EpicsPathKind.SUBSTITUTIONS),
        )?.let { return it }
      }
    }

    return resolvePath(
      project,
      currentDirectory = hostDirectory,
      searchRoots = searchRoots,
      context = context,
      expansionVariables = emptyMap(),
    )
  }

  private fun resolveSubstitutionsReference(
    hostFile: VirtualFile,
    ownerRoot: Path,
    context: EpicsReferenceContext,
  ): EpicsResolvedReference? {
    return collectSubstitutionsReferences(hostFile, ownerRoot, context).firstOrNull()
  }

  private fun resolveStreamProtocolReference(
    text: String,
    offset: Int,
    ownerRoot: Path,
    context: EpicsReferenceContext,
  ): EpicsResolvedReference? {
    val protocolPath = getStreamProtocolReferenceAtOffset(text, offset) ?: return null
    val searchDirectories = collectStreamProtocolSearchDirectories(ownerRoot)
    for (searchDirectory in searchDirectories) {
      val candidate = resolveAbsoluteOrRelative(searchDirectory, protocolPath)
      resolveFileCandidate(candidate, context)?.let { return it }
    }
    return null
  }

  private fun resolvePath(
    project: Project,
    currentDirectory: Path,
    searchRoots: List<Path>,
    context: EpicsReferenceContext,
    expansionVariables: Map<String, String>,
  ): EpicsResolvedReference? {
    val expandedPath = expandEpicsValue(context.rawPath, expansionVariables).trim()

    val directCandidate = resolveAbsoluteOrRelative(currentDirectory, expandedPath)
    resolveFileCandidate(directCandidate, context)?.let { return it }

    val basename = runCatching { Path.of(expandedPath).fileName?.toString().orEmpty() }
      .getOrElse { expandedPath.substringAfterLast('/').substringAfterLast('\\') }
    for (root in searchRoots) {
      candidatePathsForKind(root, expandedPath, basename, context.kind).forEach { candidate ->
        resolveFileCandidate(candidate, context)?.let { return it }
      }
    }

    for (root in searchRoots) {
      searchPreferredDirectories(root, basename, context)?.let { return it }
    }

    for (contentRoot in ProjectRootManager.getInstance(project).contentRoots) {
      if (!contentRoot.isDirectory) {
        continue
      }
      searchPreferredDirectories(contentRoot.toNioPath(), basename, context)?.let { return it }
    }

    return null
  }

  private fun collectPathCompletionCandidates(
    state: StartupExecutionState,
    rawPartial: String,
    kind: EpicsPathKind,
  ): List<EpicsPathCompletionCandidate> {
    val expandedPartial = expandEpicsValue(rawPartial, state.variables)
    if (containsMakeVariableReference(expandedPartial)) {
      return emptyList()
    }

    val normalizedPartial = expandedPartial.replace('\\', '/')
    val partialWithinDb = when {
      normalizedPartial == "db" -> ""
      normalizedPartial.startsWith("db/") -> normalizedPartial.removePrefix("db/")
      else -> normalizedPartial
    }
    val hasSeparator = partialWithinDb.contains('/')
    val hasTrailingSeparator = partialWithinDb.endsWith('/')
    val relativeDirectory = when {
      !hasSeparator -> "."
      hasTrailingSeparator -> partialWithinDb.removeSuffix("/")
      else -> partialWithinDb.substringBeforeLast('/')
    }.ifBlank { "." }
    val namePrefix = when {
      hasTrailingSeparator -> ""
      hasSeparator -> partialWithinDb.substringAfterLast('/')
      else -> partialWithinDb
    }

    val candidates = linkedMapOf<String, EpicsPathCompletionCandidate>()
    collectDbDirectoryCompletionCandidates(
      baseDirectory = state.currentDirectory,
      dbDirectory = resolveAbsoluteOrRelative(
        state.currentDirectory,
        if (relativeDirectory == "." || relativeDirectory == "db") "db" else "db/$relativeDirectory",
      ),
      namePrefix = namePrefix,
      kind = kind,
      detailPrefix = "db",
      candidates = candidates,
    )

    for (root in state.searchRoots) {
      if (root.normalize() == state.currentDirectory.normalize()) {
        continue
      }
      val dbDirectory = root.resolve("db")
      if (!dbDirectory.exists() || !dbDirectory.isDirectory()) {
        continue
      }
      collectDbDirectoryCompletionCandidates(
        baseDirectory = state.currentDirectory,
        dbDirectory = if (relativeDirectory == "." || relativeDirectory == "db") {
          dbDirectory
        } else {
          dbDirectory.resolve(relativeDirectory)
        },
        namePrefix = namePrefix,
        kind = kind,
        detailPrefix = root.fileName?.toString()?.let { "$it/db" } ?: "db",
        candidates = candidates,
      )
    }

    return candidates.values.sortedBy { it.insertPath.lowercase() }
  }

  private fun collectDbLoadRecordsCompletionCandidates(
    state: StartupExecutionState,
    rawPartial: String,
  ): List<EpicsPathCompletionCandidate> {
    val expandedPartial = expandEpicsValue(rawPartial, state.variables)
    if (containsMakeVariableReference(expandedPartial)) {
      return emptyList()
    }

    val normalizedPartial = expandedPartial.replace('\\', '/')
    val query = normalizedPartial
      .removePrefix("db/")
      .removePrefix("Db/")
      .substringAfterLast('/')

    val candidates = linkedMapOf<String, EpicsPathCompletionCandidate>()
    collectDbLoadRecordsDirectoryFiles(
      baseDirectory = state.currentDirectory,
      directory = state.currentDirectory,
      namePrefix = query,
      detail = "current dir",
      candidates = candidates,
    )
    collectDbLoadRecordsDirectoryFiles(
      baseDirectory = state.currentDirectory,
      directory = state.currentDirectory.resolve("db"),
      namePrefix = query,
      detail = "db",
      candidates = candidates,
    )
    return candidates.values.sortedBy { it.insertPath.lowercase() }
  }

  private fun collectDbLoadRecordsDirectoryFiles(
    baseDirectory: Path,
    directory: Path,
    namePrefix: String,
    detail: String,
    candidates: MutableMap<String, EpicsPathCompletionCandidate>,
  ) {
    if (!directory.exists() || !directory.isDirectory()) {
      return
    }

    val allowedExtensions = allowedExtensionsForKind(EpicsPathKind.DATABASE)
    directory.toFile().listFiles().orEmpty().sortedBy { it.name.lowercase() }.forEach { child ->
      if (child.isDirectory) {
        return@forEach
      }
      if (child.extension.lowercase() !in allowedExtensions) {
        return@forEach
      }
      if (namePrefix.isNotBlank() && !child.name.startsWith(namePrefix, ignoreCase = true)) {
        return@forEach
      }

      val childPath = child.toPath()
      val insertPath = relativizePath(baseDirectory, childPath)
      candidates.putIfAbsent(
        insertPath,
        EpicsPathCompletionCandidate(
          insertPath = insertPath,
          detail = detail,
          isDirectory = false,
          absolutePath = childPath,
        ),
      )
    }
  }

  private fun collectDbDirectoryCompletionCandidates(
    baseDirectory: Path,
    dbDirectory: Path,
    namePrefix: String,
    kind: EpicsPathKind,
    detailPrefix: String,
    candidates: MutableMap<String, EpicsPathCompletionCandidate>,
  ) {
    if (!dbDirectory.exists() || !dbDirectory.isDirectory()) {
      return
    }
    val allowedExtensions = allowedExtensionsForKind(kind)
    dbDirectory.toFile().listFiles().orEmpty().sortedBy { if (it.isDirectory) 0 else 1 }.forEach { child ->
      if (namePrefix.isNotBlank() && !child.name.startsWith(namePrefix, ignoreCase = true)) {
        return@forEach
      }

      val childPath = child.toPath()
      val insertPath = relativizePath(baseDirectory, childPath)
      if (child.isDirectory) {
        candidates.putIfAbsent(
          "$insertPath/",
          EpicsPathCompletionCandidate(
            insertPath = "$insertPath/",
            detail = "$detailPrefix/${child.name}",
            isDirectory = true,
            absolutePath = childPath,
          ),
        )
        return@forEach
      }

      if (child.extension.lowercase() !in allowedExtensions) {
        return@forEach
      }
      candidates.putIfAbsent(
        insertPath,
        EpicsPathCompletionCandidate(
          insertPath = insertPath,
          detail = "$detailPrefix/${child.name}",
          isDirectory = false,
          absolutePath = childPath,
        ),
      )
    }
  }

  private fun applyStartupStateUntilOffset(
    text: String,
    untilOffset: Int,
    state: StartupExecutionState,
  ) {
    var runningOffset = 0
    for (line in text.split('\n')) {
      val lineStart = runningOffset
      val lineEnd = lineStart + line.length
      if (lineStart >= untilOffset) {
        break
      }
      if (lineEnd < untilOffset) {
        applyStartupLine(line, state)
      }
      runningOffset = lineEnd + 1
    }
  }

  private fun allowedExtensionsForKind(kind: EpicsPathKind): Set<String> = when (kind) {
    EpicsPathKind.DATABASE -> setOf("db", "vdb", "template")
    EpicsPathKind.SUBSTITUTIONS -> setOf("substitutions", "sub", "subs")
    EpicsPathKind.PROTOCOL -> setOf("proto")
    EpicsPathKind.DBD -> setOf("dbd")
    EpicsPathKind.LIBRARY -> emptySet()
  }

  private fun relativizePath(baseDirectory: Path, target: Path): String {
    return runCatching {
      baseDirectory.normalize().relativize(target.normalize()).pathString
    }.getOrElse {
      target.normalize().pathString
    }.replace(File.separatorChar, '/')
  }

  private fun createStartupExecutionState(hostFile: VirtualFile, ownerRoot: Path): StartupExecutionState {
    val releaseVariables = loadReleaseVariables(ownerRoot).toMutableMap()
    val envPathsVariables = loadEnvPathsVariables(hostFile.parent?.toNioPath(), releaseVariables)
    val variables = linkedMapOf<String, String>()
    variables.putAll(releaseVariables)
    variables.putAll(envPathsVariables)
    variables.putIfAbsent("TOP", ownerRoot.pathString)
    val currentDirectory = hostFile.parent?.toNioPath() ?: ownerRoot
    return StartupExecutionState(
      currentDirectory = currentDirectory,
      variables = variables,
      searchRoots = buildSearchRoots(ownerRoot, releaseVariables, envPathsVariables),
    )
  }

  private fun applyStartupLine(line: String, state: StartupExecutionState) {
    val sanitizedLine = maskHashCommentLine(line)

    val epicsEnvMatch = STARTUP_ENV_SET_REGEX.find(sanitizedLine)
    if (epicsEnvMatch != null) {
      val name = epicsEnvMatch.groups[1]?.value?.trim().orEmpty()
      val rawValue = epicsEnvMatch.groups[2]?.value.orEmpty()
      if (name.isNotEmpty()) {
        state.variables[name] = expandEpicsValue(rawValue, state.variables)
      }
    }

    val cdMatch = STARTUP_CD_REGEX.find(sanitizedLine)
    if (cdMatch != null) {
      val rawDirectory = cdMatch.groups[1]?.value ?: cdMatch.groups[2]?.value ?: ""
      val expandedDirectory = expandEpicsValue(rawDirectory, state.variables)
      val resolvedDirectory = resolveAbsoluteOrRelative(state.currentDirectory, expandedDirectory)
      if (resolvedDirectory.exists() && resolvedDirectory.isDirectory()) {
        state.currentDirectory = resolvedDirectory.normalize()
      }
    }
  }

  private data class StartupLineMatch(
    val startOffset: Int,
    val lineStart: Int,
    val lineEnd: Int,
    val rangeStart: Int,
    val rangeEnd: Int,
    val context: EpicsReferenceContext,
  )

  private fun collectStartupLineMatches(text: String): List<StartupLineMatch> {
    val matches = mutableListOf<StartupLineMatch>()
    var offset = 0
    for (line in text.split('\n')) {
      val lineStart = offset
      val lineEnd = lineStart + line.length
      findStartupReferenceOnLine(line, lineStart)?.let { matches += it.copy(lineEnd = lineEnd) }
      offset = lineEnd + 1
    }
    return matches
  }

  private fun findStartupReferenceOnLine(line: String, lineStart: Int): StartupLineMatch? {
    val sanitizedLine = maskHashCommentLine(line)
    for ((regex, kind) in STARTUP_LOAD_PATTERNS) {
      val match = regex.find(sanitizedLine) ?: continue
      val group = match.groups[1] ?: continue
      return StartupLineMatch(
        startOffset = lineStart,
        lineStart = lineStart,
        lineEnd = lineStart + line.length,
        rangeStart = lineStart + group.range.first,
        rangeEnd = lineStart + group.range.last + 1,
        context = EpicsReferenceContext(kind, group.value),
      )
    }
    return null
  }

  private fun getStartupReferenceAtOffset(
    project: Project,
    hostFile: VirtualFile,
    text: String,
    offset: Int,
  ): EpicsReferenceContext? {
    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val state = createStartupExecutionState(hostFile, ownerRoot)
    var runningOffset = 0
    for (line in text.split('\n')) {
      findStartupReferenceOnLine(line, runningOffset)?.let { match ->
        if (offset in match.rangeStart until match.rangeEnd) {
          return match.context
        }
      }
      applyStartupLine(line, state)
      runningOffset += line.length + 1
    }
    return null
  }

  private fun getDatabaseReferenceAtOffset(
    text: String,
    offset: Int,
  ): EpicsReferenceContext? {
    val protocolPath = getStreamProtocolReferenceAtOffset(text, offset) ?: return null
    return EpicsReferenceContext(EpicsPathKind.PROTOCOL, protocolPath)
  }

  private fun getMakefileReferenceAtOffset(text: String, offset: Int): EpicsReferenceContext? {
    var runningOffset = 0
    for (line in text.split('\n')) {
      val assignment = MAKEFILE_ASSIGNMENT_REGEX.find(line)
      if (assignment != null) {
        val variableName = assignment.groups[1]?.value.orEmpty()
        val kind = getMakefileReferenceKind(variableName)
        if (kind == null) {
          runningOffset += line.length + 1
          continue
        }
        val valueStart = assignment.range.last + 1
        val commentIndex = line.indexOf('#', valueStart).let { if (it >= 0) it else line.length }
        TOKEN_REGEX.findAll(line.substring(valueStart, commentIndex)).forEach { tokenMatch ->
          val token = tokenMatch.value
          if (containsMakeVariableReference(token) || token == "-nil-") {
            return@forEach
          }
          val tokenStart = runningOffset + valueStart + tokenMatch.range.first
          val tokenEnd = runningOffset + valueStart + tokenMatch.range.last + 1
          if (offset in tokenStart until tokenEnd) {
            return EpicsReferenceContext(kind, token)
          }
        }
      }
      runningOffset += line.length + 1
    }
    return null
  }

  private fun getSubstitutionsReferenceAtOffset(text: String, offset: Int): EpicsReferenceContext? {
    var runningOffset = 0
    for (line in text.split('\n')) {
      val match = SUBSTITUTIONS_FILE_REGEX.find(line)
      if (match != null) {
        val group = match.groups[1] ?: continue
        val start = runningOffset + group.range.first
        val end = runningOffset + group.range.last + 1
        if (offset in start until end) {
          val raw = group.value.trim().trim('"')
          return EpicsReferenceContext(
            if (raw.lowercase().endsWith(".substitutions")) EpicsPathKind.SUBSTITUTIONS else EpicsPathKind.DATABASE,
            raw,
          )
        }
      }
      runningOffset += line.length + 1
    }
    return null
  }

  private fun getMakefileReferenceKind(variableName: String): EpicsPathKind? = when {
    DBD_VARIABLE_REGEX.matches(variableName) -> EpicsPathKind.DBD
    LIB_VARIABLE_REGEX.matches(variableName) -> EpicsPathKind.LIBRARY
    DB_VARIABLE_REGEX.matches(variableName) -> EpicsPathKind.DATABASE
    else -> null
  }

  private fun candidatePathsForKind(
    searchRoot: Path,
    expandedPath: String,
    basename: String,
    kind: EpicsPathKind,
  ): List<Path> {
    if (basename.isEmpty()) {
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
      EpicsPathKind.PROTOCOL -> buildList {
        val normalized = expandedPath.replace('\\', '/')
        add(searchRoot.resolve(normalized))
      }
      EpicsPathKind.DBD -> buildList {
        val normalized = expandedPath.replace('\\', '/')
        if (normalized.startsWith("dbd/")) {
          add(searchRoot.resolve(normalized))
        }
        add(searchRoot.resolve("dbd").resolve(basename))
        add(searchRoot.resolve(normalized))
      }
      EpicsPathKind.LIBRARY -> buildList {
        addAll(libraryCandidatePaths(searchRoot, basename))
      }
    }
  }

  private fun libraryCandidatePaths(searchRoot: Path, libraryName: String): List<Path> {
    val libRoot = searchRoot.resolve("lib")
    if (!libRoot.exists() || !libRoot.isDirectory()) {
      return emptyList()
    }
    val fileNameCandidates = buildList {
      if (libraryName.contains('/')) {
        add(libraryName)
      } else if (libraryName.endsWith(".a") || libraryName.endsWith(".so") || libraryName.endsWith(".dylib") || libraryName.endsWith(".lib")) {
        add(libraryName)
      } else {
        add("lib$libraryName.a")
        add("lib$libraryName.dylib")
        add("lib$libraryName.so")
        add("lib$libraryName.dll")
        add("$libraryName.lib")
      }
    }
    val architectureHint = detectHostArchitecture()
    val architectureDirectories = libRoot.toFile().listFiles { file -> file.isDirectory }?.map { it.toPath() }.orEmpty()
    val preferred = architectureDirectories.sortedBy { directory ->
      if (architectureHint != null && directory.name == architectureHint) 0 else 1
    }
    return buildList {
      for (directory in preferred) {
        for (candidate in fileNameCandidates) {
          add(directory.resolve(candidate))
        }
      }
    }
  }

  private fun searchPreferredDirectories(root: Path, basename: String, context: EpicsReferenceContext): EpicsResolvedReference? {
    val directories = when (context.kind) {
      EpicsPathKind.DATABASE, EpicsPathKind.SUBSTITUTIONS -> listOf(root.resolve("db"), root.resolve("Db"))
      EpicsPathKind.PROTOCOL -> emptyList()
      EpicsPathKind.DBD -> listOf(root.resolve("dbd"))
      EpicsPathKind.LIBRARY -> root.resolve("lib").toFile().listFiles { file -> file.isDirectory }?.map { it.toPath() }.orEmpty()
    }

    val names = if (context.kind == EpicsPathKind.LIBRARY && !basename.contains('.')) {
      setOf(
        "lib$basename.a",
        "lib$basename.dylib",
        "lib$basename.so",
        "lib$basename.dll",
        "$basename.lib",
      )
    } else {
      setOf(basename)
    }

    for (directory in directories) {
      if (!directory.exists() || !directory.isDirectory()) {
        continue
      }
      val found = findFileRecursively(directory.toFile(), names)
      if (found != null) {
        return LocalFileSystem.getInstance().findFileByIoFile(found)?.let { file ->
          EpicsResolvedReference(file, context.rawPath, context.kind)
        }
      }
    }
    return null
  }

  private fun findFileRecursively(directory: File, candidateNames: Set<String>): File? {
    val queue = ArrayDeque<File>()
    queue.add(directory)
    while (queue.isNotEmpty()) {
      val next = queue.removeFirst()
      val children = next.listFiles().orEmpty().sortedBy { if (it.isDirectory) 0 else 1 }
      for (child in children) {
        if (child.isDirectory) {
          queue.add(child)
          continue
        }
        if (candidateNames.contains(child.name)) {
          return child
        }
      }
    }
    return null
  }

  private fun buildSearchRoots(
    ownerRoot: Path,
    releaseVariables: Map<String, String>,
    envVariables: Map<String, String>,
  ): List<Path> {
    val roots = linkedSetOf<Path>()
    roots.add(ownerRoot.normalize())
    (releaseVariables.values + envVariables.values).forEach { value ->
      val candidate = runCatching { Path.of(value) }.getOrNull() ?: return@forEach
      if (candidate.exists() && candidate.isDirectory()) {
        roots.add(candidate.normalize())
      }
    }
    return roots.toList()
  }

  private fun toSubstitutionsSiblingPath(rawPath: String): String? {
    val normalizedPath = rawPath.replace('\\', '/')
    val lowerCasePath = normalizedPath.lowercase()
    val matchingExtension = listOf(".db", ".vdb", ".template")
      .firstOrNull { extension -> lowerCasePath.endsWith(extension) }
      ?: return null
    return normalizedPath.dropLast(matchingExtension.length) + ".substitutions"
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
      val raw = rawValues[name] ?: return null
      resolving += name
      val expanded = expandEpicsValue(raw, rawValues) { nested -> resolve(nested) }
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

  private fun loadEnvPathsVariables(startupDirectory: Path?, baseVariables: Map<String, String>): Map<String, String> {
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
        values[name] = expandEpicsValue(rawValue, values + baseVariables)
      }
    }
    return values
  }

  private fun expandEpicsValue(
    rawValue: String,
    variables: Map<String, String>,
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
        variables[name] ?: fallbackResolver?.invoke(name) ?: defaultValue ?: match.value
      }
      if (next == expanded) {
        return next.trim().trim('"')
      }
      expanded = next
    }
    return expanded.trim().trim('"')
  }

  private fun resolveAbsoluteOrRelative(baseDirectory: Path, rawPath: String): Path {
    val candidate = Path.of(rawPath)
    return if (candidate.isAbsolute) candidate.normalize() else baseDirectory.resolve(rawPath).normalize()
  }

  private fun resolveFileCandidate(candidate: Path, context: EpicsReferenceContext): EpicsResolvedReference? {
    if (!candidate.exists() || candidate.isDirectory()) {
      return null
    }
    val file = LocalFileSystem.getInstance().findFileByIoFile(candidate.toFile()) ?: return null
    return EpicsResolvedReference(file, context.rawPath, context.kind)
  }

  internal fun findOwningEpicsRoot(project: Project, hostFile: VirtualFile): Path {
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

  private fun collectStreamProtocolSearchDirectories(ownerRoot: Path): List<Path> {
    val iocBootDirectory = ownerRoot.resolve("iocBoot")
    if (!iocBootDirectory.exists() || !iocBootDirectory.isDirectory()) {
      return emptyList()
    }

    val searchDirectories = linkedSetOf<Path>()
    val startupFiles = collectStartupFiles(iocBootDirectory)
    for (startupPath in startupFiles) {
      val startupFile = LocalFileSystem.getInstance().findFileByIoFile(startupPath.toFile()) ?: continue
      val text = runCatching { startupFile.inputStream.bufferedReader().use { it.readText() } }.getOrNull() ?: continue
      if (!text.contains(STREAM_PROTOCOL_PATH_VARIABLE)) {
        continue
      }

      val state = createStartupExecutionState(startupFile, ownerRoot)
      for (line in text.split('\n')) {
        val envMatch = STARTUP_ENV_SET_REGEX.find(line)
        if (envMatch != null) {
          val name = envMatch.groups[1]?.value?.trim().orEmpty()
          val rawValue = envMatch.groups[2]?.value.orEmpty()
          if (name == STREAM_PROTOCOL_PATH_VARIABLE) {
            splitStreamProtocolSearchDirectories(
              expandEpicsValue(rawValue, state.variables),
              state.currentDirectory,
            ).forEach(searchDirectories::add)
          }
        }
        applyStartupLine(line, state)
      }
    }

    return searchDirectories.filter { it.exists() && it.isDirectory() }
  }

  private fun collectStartupFiles(rootDirectory: Path): List<Path> {
    return try {
      Files.walk(rootDirectory).use { stream ->
        stream
          .filter { path -> Files.isRegularFile(path) && isStartupFilePath(path) }
          .sorted(compareBy<Path> { it.pathString.lowercase() })
          .toList()
      }
    } catch (_: Exception) {
      emptyList()
    }
  }

  private fun isStartupFilePath(path: Path): Boolean {
    val fileName = path.fileName?.toString().orEmpty()
    val extension = path.extension.lowercase()
    return extension == "cmd" || extension == "iocsh" || fileName == "st.cmd"
  }

  private fun splitStreamProtocolSearchDirectories(value: String, baseDirectory: Path): List<Path> {
    return value
      .split(File.pathSeparatorChar)
      .map(String::trim)
      .filter(String::isNotBlank)
      .map { entry ->
        val candidate = runCatching { Path.of(entry) }.getOrNull()
        if (candidate != null && candidate.isAbsolute) {
          candidate.normalize()
        } else {
          baseDirectory.resolve(entry).normalize()
        }
      }
  }

  private fun maskHashCommentLine(line: String): String {
    val sanitized = StringBuilder(line.length)
    var inString = false
    var escaped = false

    for (character in line) {
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
        sanitized.append(' ')
        repeat(line.length - sanitized.length) {
          sanitized.append(' ')
        }
        break
      }

      sanitized.append(character)
    }

    return sanitized.toString()
  }

  private fun getStreamProtocolReferenceAtOffset(text: String, offset: Int): String? {
    val recordDeclaration = getEnclosingRecordDeclaration(text, offset) ?: return null
    val fieldDeclarations = extractFieldDeclarationsInRecord(text, recordDeclaration)
    if (!isStreamDeviceRecord(fieldDeclarations)) {
      return null
    }

    for (fieldDeclaration in fieldDeclarations) {
      if (!isLinkField(recordDeclaration.recordType, fieldDeclaration.fieldName)) {
        continue
      }

      val match = STREAM_PROTOCOL_REFERENCE_REGEX.find(fieldDeclaration.value) ?: continue
      val protocolPath = match.groups[1]?.value.orEmpty()
      if (protocolPath.isBlank() || EPICS_VARIABLE_REGEX.containsMatchIn(protocolPath)) {
        continue
      }

      val protocolStart = fieldDeclaration.valueStart + (match.groups[1]?.range?.first ?: continue)
      val protocolEnd = protocolStart + protocolPath.length
      if (offset in protocolStart until protocolEnd) {
        return protocolPath
      }
    }

    return null
  }

  private data class DatabaseRecordDeclaration(
    val recordType: String,
    val recordStart: Int,
    val recordEnd: Int,
  )

  private data class DatabaseFieldDeclaration(
    val fieldName: String,
    val value: String,
    val valueStart: Int,
  )

  private fun getEnclosingRecordDeclaration(text: String, offset: Int): DatabaseRecordDeclaration? {
    val sanitizedText = maskDatabaseComments(text)
    val regex = Regex("""\brecord\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"""")
    for (match in regex.findAll(sanitizedText)) {
      val recordType = match.groups[1]?.value.orEmpty()
      val recordStart = match.range.first
      val recordEnd = findRecordBlockEnd(sanitizedText, recordStart)
      if (offset in recordStart..recordEnd) {
        return DatabaseRecordDeclaration(recordType, recordStart, recordEnd)
      }
    }
    return null
  }

  private fun extractFieldDeclarationsInRecord(
    text: String,
    recordDeclaration: DatabaseRecordDeclaration,
  ): List<DatabaseFieldDeclaration> {
    val declarations = mutableListOf<DatabaseFieldDeclaration>()
    val sanitizedText = maskDatabaseComments(text)
    val recordText = sanitizedText.substring(recordDeclaration.recordStart, recordDeclaration.recordEnd)
    val regex = Regex("""field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"""")
    val valuePrefixRegex = Regex("""field\(\s*(?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+)\s*,\s*"""")

    for (match in regex.findAll(recordText)) {
      val rawFieldName = match.groups[1]?.value ?: match.groups[2]?.value ?: continue
      val valuePrefixLength = valuePrefixRegex.find(match.value)?.value?.length ?: continue
      declarations += DatabaseFieldDeclaration(
        fieldName = rawFieldName.uppercase(),
        value = match.groups[3]?.value.orEmpty(),
        valueStart = recordDeclaration.recordStart + match.range.first + valuePrefixLength,
      )
    }

    return declarations
  }

  private fun isStreamDeviceRecord(fieldDeclarations: List<DatabaseFieldDeclaration>): Boolean {
    val dtypField = fieldDeclarations.firstOrNull { it.fieldName == "DTYP" } ?: return false
    return dtypField.value.trim().equals("stream", ignoreCase = true)
  }

  private fun isLinkField(recordType: String, fieldName: String): Boolean {
    val dbfType = EpicsRecordCompletionSupport.getFieldType(recordType, fieldName)
    return dbfType?.contains("LINK") == true
  }

  private fun maskDatabaseComments(text: String): String {
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
    val openingBrace = text.indexOf('{', recordStart)
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

  private fun detectHostArchitecture(): String? {
    val osName = System.getProperty("os.name")?.lowercase().orEmpty()
    val osArch = System.getProperty("os.arch")?.lowercase().orEmpty()
    return when {
      osName.contains("mac") && osArch.contains("aarch64") -> "darwin-aarch64"
      osName.contains("mac") && (osArch.contains("x86_64") || osArch.contains("amd64")) -> "darwin-x86_64"
      osName.contains("linux") && (osArch.contains("x86_64") || osArch.contains("amd64")) -> "linux-x86_64"
      osName.contains("linux") && osArch.contains("aarch64") -> "linux-aarch64"
      osName.contains("windows") && (osArch.contains("x86_64") || osArch.contains("amd64")) -> "windows-x64"
      else -> null
    }
  }

  internal fun isStartupFile(file: VirtualFile): Boolean {
    val extension = file.extension?.lowercase()
    return extension == "cmd" || extension == "iocsh" || file.name == "st.cmd"
  }

  private fun isDatabaseFile(file: VirtualFile): Boolean {
    val extension = file.extension?.lowercase()
    return extension in setOf("db", "vdb", "template")
  }

  private fun isSubstitutionsFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("substitutions", "sub", "subs")
  }

  private fun isMakefile(file: VirtualFile): Boolean = file.name == "Makefile"

  private val STARTUP_LOAD_PATTERNS = listOf(
    Regex("""\bdbLoadDatabase(?:\(\s*|\s+)\"([^\\"\n]+)\"""") to EpicsPathKind.DBD,
    Regex("""\bdbLoadRecords\(\s*\"([^\\"\n]+)\"""") to EpicsPathKind.DATABASE,
    Regex("""\bdbLoadTemplate\(\s*\"([^\\"\n]+)\"""") to EpicsPathKind.SUBSTITUTIONS,
  )
  private val STARTUP_ENV_SET_REGEX = Regex("""\bepicsEnvSet\(\s*\"([^\"]+)\"\s*,\s*\"([^\"]*)\"\s*\)""")
  private val STARTUP_CD_REGEX = Regex("""^\s*cd\s+(?:\"([^\"]+)\"|([^\s#]+))""")
  private val STREAM_PROTOCOL_REFERENCE_REGEX = Regex("""^\s*@([^\s"'`]+)""")
  private const val STREAM_PROTOCOL_PATH_VARIABLE = "STREAM_PROTOCOL_PATH"
  private val MAKEFILE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z0-9_.-]+)\s*(?:\+?=|:=|\?=)\s*""")
  private val TOKEN_REGEX = Regex("""[^\s]+""")
  private val DBD_VARIABLE_REGEX = Regex("""^(?:[A-Za-z0-9_.-]+_)?DBD$""")
  private val LIB_VARIABLE_REGEX = Regex("""^(?:[A-Za-z0-9_.-]+_)?LIBS$""")
  private val DB_VARIABLE_REGEX = Regex("""^(?:[A-Za-z0-9_.-]+_)?DB$""")
  private val SUBSTITUTIONS_FILE_REGEX = Regex("""\bfile\s+("?[^"\s{]+"?)\s*\{""")
  private val RELEASE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_.-]*)\s*=\s*(.*?)\s*(?:#.*)?$""")
  private val EPICS_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")

  private fun containsMakeVariableReference(value: String): Boolean {
    return value.contains("\$(") || value.contains("\${")
  }

  private fun collectSubstitutionsReferences(
    hostFile: VirtualFile,
    ownerRoot: Path,
    context: EpicsReferenceContext,
  ): List<EpicsResolvedReference> {
    val hostDirectory = hostFile.parent?.toNioPath() ?: ownerRoot
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val expandedPath = expandEpicsValue(context.rawPath, releaseVariables).trim()
    if (expandedPath.isBlank()) {
      return emptyList()
    }

    val basename = runCatching { Path.of(expandedPath).fileName?.toString().orEmpty() }
      .getOrElse { expandedPath.substringAfterLast('/').substringAfterLast('\\') }
    val references = mutableListOf<EpicsResolvedReference>()
    val seenPaths = linkedSetOf<Path>()

    fun addCandidate(candidate: Path) {
      val normalized = candidate.normalize()
      if (!seenPaths.add(normalized)) {
        return
      }
      resolveFileCandidate(normalized, context)?.let(references::add)
    }

    addCandidate(resolveAbsoluteOrRelative(hostDirectory, expandedPath))

    candidatePathsForKind(ownerRoot, expandedPath, basename, context.kind).forEach(::addCandidate)

    buildSearchRoots(ownerRoot, releaseVariables, emptyMap())
      .asSequence()
      .filter { it.normalize() != ownerRoot.normalize() }
      .forEach { searchRoot ->
        candidatePathsForKind(searchRoot, expandedPath, basename, context.kind).forEach(::addCandidate)
      }

    return references
  }
}
