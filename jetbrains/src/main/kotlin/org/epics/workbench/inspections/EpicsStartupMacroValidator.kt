package org.epics.workbench.inspections

import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import java.nio.charset.Charset

internal object EpicsStartupMacroValidator {
  private data class StartupLoadStatement(
    val command: String,
    val pathStart: Int,
    val pathEnd: Int,
    val macros: String,
    val macrosStart: Int?,
  )

  private data class MacroAssignment(
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
  )

  fun collectIssues(
    project: Project,
    hostFile: VirtualFile,
    text: String,
  ): List<EpicsDatabaseValueValidator.ValidationIssue> {
    val issues = mutableListOf<EpicsDatabaseValueValidator.ValidationIssue>()

    for (statement in extractStartupLoadStatements(text)) {
      if (statement.command != "dbLoadRecords") {
        continue
      }

      val resolved = EpicsPathResolver.resolveReference(project, hostFile, statement.pathStart) ?: continue
      if (resolved.kind != EpicsPathKind.DATABASE) {
        continue
      }

      val targetText = readText(resolved.targetFile) ?: continue
      val sanitizedText = maskDatabaseComments(targetText)
      val requiredMacroNames = extractRequiredMacroNames(sanitizedText)
      val definedMacroNames = extractMacroNames(sanitizedText).toSet()
      val providedAssignments = extractNamedAssignments(statement.macros, statement.macrosStart)
      val providedMacroNames = providedAssignments.mapTo(linkedSetOf()) { it.name }

      val missingMacroNames = requiredMacroNames.filterNot(providedMacroNames::contains)
      if (missingMacroNames.isNotEmpty()) {
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = statement.pathStart,
          endOffset = statement.pathEnd,
          message = "dbLoadRecords is missing macro assignments for \"${resolved.targetFile.name}\": ${missingMacroNames.joinToString(", ")}.",
        )
      }

      for (assignment in providedAssignments) {
        if (assignment.name in definedMacroNames) {
          continue
        }
        issues += EpicsDatabaseValueValidator.ValidationIssue(
          startOffset = assignment.nameStart,
          endOffset = assignment.nameEnd,
          message = "Macro \"${assignment.name}\" is not defined in \"${resolved.targetFile.name}\".",
        )
      }
    }

    return issues
  }

  private fun extractStartupLoadStatements(text: String): List<StartupLoadStatement> {
    val statements = mutableListOf<StartupLoadStatement>()
    var lineOffset = 0

    for (line in text.split('\n')) {
      val match = DB_LOAD_RECORDS_REGEX.find(line)
      if (match != null) {
        val pathGroup = match.groups[1]
        if (pathGroup != null) {
          val macrosGroup = match.groups[2]
          statements += StartupLoadStatement(
            command = "dbLoadRecords",
            pathStart = lineOffset + pathGroup.range.first,
            pathEnd = lineOffset + pathGroup.range.last + 1,
            macros = macrosGroup?.value.orEmpty(),
            macrosStart = macrosGroup?.let { lineOffset + it.range.first },
          )
        }
      }

      lineOffset += line.length + 1
    }

    return statements
  }

  private fun extractNamedAssignments(
    text: String,
    absoluteStartOffset: Int?,
  ): List<MacroAssignment> {
    if (text.isBlank() || absoluteStartOffset == null) {
      return emptyList()
    }

    val assignments = mutableListOf<MacroAssignment>()
    var segmentStart = 0
    var escaped = false

    fun flushSegment(segmentEnd: Int) {
      val segment = text.substring(segmentStart, segmentEnd)
      val match = NAMED_ASSIGNMENT_REGEX.find(segment) ?: return
      val nameGroup = match.groups[1] ?: return
      val name = nameGroup.value
      if (name.isBlank()) {
        return
      }
      assignments += MacroAssignment(
        name = name,
        nameStart = absoluteStartOffset + segmentStart + nameGroup.range.first,
        nameEnd = absoluteStartOffset + segmentStart + nameGroup.range.last + 1,
      )
    }

    for (index in text.indices) {
      val character = text[index]
      when {
        escaped -> escaped = false
        character == '\\' -> escaped = true
        character == ',' -> {
          flushSegment(index)
          segmentStart = index + 1
        }
      }
    }

    flushSegment(text.length)
    return assignments
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
    DATABASE_MACRO_REGEX.findAll(text).forEach { match ->
      val name = match.groups[1]?.value ?: match.groups[2]?.value
      if (!name.isNullOrBlank()) {
        names += name
      }
    }
    return names.toList().sortedWith(String.CASE_INSENSITIVE_ORDER)
  }

  private fun extractRequiredMacroNames(text: String): List<String> {
    val names = linkedSetOf<String>()
    REQUIRED_DATABASE_MACRO_REGEX.findAll(text).forEach { match ->
      val parenthesizedName = match.groups[1]?.value
      val parenthesizedDefault = match.groups[2]?.value
      val bracedName = match.groups[3]?.value
      val bracedDefault = match.groups[4]?.value

      if (parenthesizedName != null && parenthesizedDefault == null) {
        names += parenthesizedName
      }
      if (bracedName != null && bracedDefault == null) {
        names += bracedName
      }
    }
    return names.toList().sortedWith(String.CASE_INSENSITIVE_ORDER)
  }

  private fun readText(file: VirtualFile): String? {
    return try {
      String(file.contentsToByteArray(), Charset.forName(file.charset.name()))
    } catch (_: Exception) {
      null
    }
  }

  private val DB_LOAD_RECORDS_REGEX =
    Regex("""^\s*dbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"((?:[^"\\]|\\.)*)")?""")
  private val NAMED_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=""")
  private val DATABASE_MACRO_REGEX =
    Regex("""\$\(([^)=,\s]+)(?:=[^)]*)?\)|\$\{([^}=,\s]+)(?:=[^}]*)?\}""")
  private val REQUIRED_DATABASE_MACRO_REGEX =
    Regex("""\$\(([^)=,\s]+)(?:=([^)]*))?\)|\$\{([^}=,\s]+)(?:=([^}]*))?\}""")
}
