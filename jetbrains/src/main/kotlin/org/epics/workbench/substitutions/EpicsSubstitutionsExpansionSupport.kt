package org.epics.workbench.substitutions

import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.exists
import kotlin.io.path.isDirectory
import kotlin.io.path.name
import kotlin.io.path.pathString

internal data class EpicsExpandedSubstitutionsResult(
  val expandedText: String?,
  val issues: List<String>,
)

internal object EpicsSubstitutionsExpansionSupport {
  fun isSubstitutionsFile(file: VirtualFile?): Boolean {
    return file?.extension?.lowercase() in SUBSTITUTIONS_EXTENSIONS
  }

  fun expandToDatabaseText(
    project: Project,
    file: VirtualFile,
  ): EpicsExpandedSubstitutionsResult {
    val sourcePath = runCatching { Path.of(file.path) }.getOrNull()
      ?: return EpicsExpandedSubstitutionsResult(
        expandedText = null,
        issues = listOf("Failed to resolve ${file.name} on disk."),
      )
    val sourceText = runCatching { Files.readString(sourcePath) }.getOrElse { error ->
      return EpicsExpandedSubstitutionsResult(
        expandedText = null,
        issues = listOf(error.message ?: "Failed to read ${file.name}."),
      )
    }

    val ownerRoot = findOwningEpicsRoot(project, file)
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val searchRoots = buildSearchRoots(ownerRoot, releaseVariables)
    var globalMacros = emptyMap<String, String>()
    val expandedSections = mutableListOf<String>()
    val issues = mutableListOf<String>()

    for (block in extractSubstitutionBlocks(sourceText)) {
      when (block.kind) {
        "global" -> {
          globalMacros = mergeMacroMaps(globalMacros, extractNamedAssignments(block.body))
        }

        "file" -> {
          val templatePath = block.templatePath
          if (templatePath.isNullOrBlank()) {
            continue
          }
          val templateFile = resolveTemplatePath(file, templatePath, releaseVariables, searchRoots)
          if (templateFile == null) {
            issues += """Cannot resolve substitutions database/template file "$templatePath"."""
            continue
          }
          val templateText = runCatching { Files.readString(templateFile) }.getOrElse { error ->
            issues += error.message ?: "Failed to read ${templateFile.name}."
            null
          } ?: continue
          val templateTextWithoutToc = removeDatabaseTocBlock(templateText)
          val rows = parseSubstitutionRows(block.body)
          val effectiveRows = if (rows.isEmpty()) {
            listOf(globalMacros)
          } else {
            rows.map { row -> mergeMacroMaps(globalMacros, row) }
          }

          effectiveRows.forEach { rowMacros ->
            val expanded = expandEpicsValue(
              templateTextWithoutToc,
              listOf(rowMacros, releaseVariables, System.getenv()),
            ).trimEnd()
            if (expanded.isNotBlank()) {
              expandedSections += expanded
            }
          }
        }
      }
    }

    if (issues.isNotEmpty()) {
      return EpicsExpandedSubstitutionsResult(expandedText = null, issues = issues)
    }

    val expandedText = expandedSections.joinToString("\n\n").trim()
    if (expandedText.isBlank()) {
      return EpicsExpandedSubstitutionsResult(
        expandedText = null,
        issues = listOf("No database/template expansions were found in ${file.name}."),
      )
    }

    return EpicsExpandedSubstitutionsResult(expandedText = expandedText, issues = emptyList())
  }

  private fun findOwningEpicsRoot(project: Project, hostFile: VirtualFile): Path {
    var current = runCatching { hostFile.parent?.toNioPath() }.getOrNull()
    while (current != null) {
      if (current.resolve("configure").resolve("RELEASE").exists()) {
        return current
      }
      current = current.parent
    }
    return Path.of(project.basePath ?: hostFile.parent?.path ?: ".").normalize()
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
      val expanded = expandEpicsValue(raw, listOf(rawValues)) { nested -> resolve(nested) }
      val normalized = if (expanded.isNotBlank() && !runCatching { Path.of(expanded).isAbsolute }.getOrDefault(false)) {
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

  private fun buildSearchRoots(
    ownerRoot: Path,
    releaseVariables: Map<String, String>,
  ): List<Path> {
    val roots = linkedSetOf<Path>()
    roots.add(ownerRoot.normalize())
    releaseVariables.values.forEach { value ->
      val candidate = runCatching { Path.of(value) }.getOrNull() ?: return@forEach
      if (candidate.exists() && candidate.isDirectory()) {
        roots.add(candidate.normalize())
      }
    }
    return roots.toList()
  }

  private fun resolveTemplatePath(
    substitutionsFile: VirtualFile,
    templatePath: String,
    releaseVariables: Map<String, String>,
    searchRoots: List<Path>,
  ): Path? {
    val expandedPath = expandEpicsValue(templatePath, listOf(releaseVariables, System.getenv()))
    if (expandedPath.isBlank()) {
      return null
    }

    val baseDirectory = runCatching { substitutionsFile.parent?.toNioPath() }.getOrNull()
      ?: searchRoots.firstOrNull()
      ?: return null
    val directCandidate = resolveAbsoluteOrRelative(baseDirectory, expandedPath)
    if (directCandidate.exists() && !directCandidate.isDirectory()) {
      return directCandidate
    }

    val basename = runCatching { Path.of(expandedPath).fileName?.toString().orEmpty() }
      .getOrElse { expandedPath.substringAfterLast('/').substringAfterLast('\\') }
    for (root in searchRoots) {
      for (candidate in candidatePaths(root, expandedPath, basename)) {
        if (candidate.exists() && !candidate.isDirectory()) {
          return candidate
        }
      }
    }
    for (root in searchRoots) {
      searchPreferredDirectories(root, basename)?.let { return it }
    }
    return null
  }

  private fun candidatePaths(
    searchRoot: Path,
    expandedPath: String,
    basename: String,
  ): List<Path> {
    if (basename.isBlank()) {
      return emptyList()
    }

    val normalized = expandedPath.replace('\\', '/')
    return buildList {
      if (normalized.startsWith("db/") || normalized.startsWith("Db/")) {
        add(searchRoot.resolve(normalized))
      }
      add(searchRoot.resolve("db").resolve(basename))
      add(searchRoot.resolve("Db").resolve(basename))
      add(searchRoot.resolve(normalized))
    }
  }

  private fun searchPreferredDirectories(root: Path, basename: String): Path? {
    val directories = listOf(root.resolve("db"), root.resolve("Db"))
    for (directory in directories) {
      if (!directory.exists() || !directory.isDirectory()) {
        continue
      }
      val found = findFileRecursively(directory, basename)
      if (found != null) {
        return found
      }
    }
    return null
  }

  private fun findFileRecursively(directory: Path, basename: String): Path? {
    val queue = ArrayDeque<Path>()
    queue.addLast(directory)
    while (queue.isNotEmpty()) {
      val next = queue.removeFirst()
      val children = runCatching { Files.list(next).use { it.toList() } }.getOrDefault(emptyList())
        .sortedBy { if (Files.isDirectory(it)) 0 else 1 }
      for (child in children) {
        if (Files.isDirectory(child)) {
          queue.addLast(child)
          continue
        }
        if (child.name == basename) {
          return child
        }
      }
    }
    return null
  }

  private fun resolveAbsoluteOrRelative(baseDirectory: Path, rawPath: String): Path {
    val candidate = runCatching { Path.of(rawPath) }.getOrNull()
    return if (candidate != null && candidate.isAbsolute) {
      candidate.normalize()
    } else {
      baseDirectory.resolve(rawPath).normalize()
    }
  }

  private fun removeDatabaseTocBlock(text: String): String {
    val start = text.indexOf(TOC_BEGIN_MARKER)
    if (start < 0) {
      return text
    }
    val endMarkerStart = text.indexOf(TOC_END_MARKER, start)
    if (endMarkerStart < 0) {
      return text
    }
    var end = endMarkerStart + TOC_END_MARKER.length
    if (text.startsWith("\r\n", end)) {
      end += 2
    } else if (end < text.length && text[end] == '\n') {
      end += 1
    }
    while (text.startsWith("\r\n", end)) {
      end += 2
    }
    while (end < text.length && text[end] == '\n') {
      end += 1
    }
    return buildString(text.length - (end - start)) {
      append(text.substring(0, start))
      append(text.substring(end))
    }
  }

  private fun extractSubstitutionBlocks(text: String): List<SubstitutionBlock> {
    val blocks = mutableListOf<SubstitutionBlock>()
    var searchIndex = 0
    while (searchIndex < text.length) {
      val match = SUBSTITUTION_BLOCK_REGEX.find(text, searchIndex) ?: break
      val braceIndex = text.indexOf('{', match.range.first)
      if (braceIndex < 0) {
        break
      }
      val blockEnd = findBalancedBlockEnd(text, braceIndex)
      val bodyStart = braceIndex + 1
      val bodyEnd = (blockEnd - 1).coerceAtLeast(bodyStart)
      val rawTemplatePath = match.groups[2]?.value
      blocks += SubstitutionBlock(
        kind = match.groups[1]?.value.orEmpty(),
        templatePath = rawTemplatePath?.trim()?.removeSurrounding("\""),
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
      val columns = tokenizeSubstitutionValues(segments.first())
      return segments.drop(1).map { segment ->
        createPatternMacroAssignments(columns, tokenizeSubstitutionValues(segment))
      }.filter { it.isNotEmpty() }
    }

    return segments.map(::extractNamedAssignments).filter { it.isNotEmpty() }
  }

  private fun extractTopLevelBraceSegments(text: String): List<String> {
    val segments = mutableListOf<String>()
    var searchIndex = 0
    while (searchIndex < text.length) {
      val braceIndex = text.indexOf('{', searchIndex)
      if (braceIndex < 0) {
        break
      }
      val blockEnd = findBalancedBlockEnd(text, braceIndex)
      val contentStart = braceIndex + 1
      val contentEnd = (blockEnd - 1).coerceAtLeast(contentStart)
      segments += text.substring(contentStart, contentEnd)
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
    TOKEN_REGEX.findAll(text).forEach { match ->
      tokens += (match.groups[1]?.value ?: match.groups[2]?.value ?: "").trim()
    }
    return tokens
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
      assignments[name] = token.substring(separatorIndex + 1).trim().removeSurrounding("\"")
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
    parts += current.toString()
    return parts
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

  private data class SubstitutionBlock(
    val kind: String,
    val templatePath: String?,
    val body: String,
  )

  private val SUBSTITUTIONS_EXTENSIONS = setOf("substitutions", "sub", "subs")
  private const val TOC_BEGIN_MARKER = "# EPICS TOC BEGIN"
  private const val TOC_END_MARKER = "# EPICS TOC END"
  private val SUBSTITUTION_BLOCK_REGEX = Regex("""(?:^|\n)\s*(global|file)(?:\s+("(?:[^"\\]|\\.)*"|[^\s{]+))?\s*\{""")
  private val TOKEN_REGEX = Regex("""\"((?:[^"\\]|\\.)*)\"|([^,\s{}]+)""")
  private val RELEASE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_.-]*)\s*=\s*(.*?)\s*(?:#.*)?$""")
  private val EPICS_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
}
