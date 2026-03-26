package org.epics.workbench.inspections

import com.intellij.codeInspection.LocalQuickFix
import com.intellij.codeInspection.ProblemDescriptor
import com.intellij.openapi.command.WriteCommandAction
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiFile
import org.epics.workbench.build.epicsBuildModelService
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import java.nio.charset.Charset

internal object EpicsInspectionQuickFixSupport {
  fun buildQuickFixes(
    project: Project,
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): Array<LocalQuickFix> {
    return when (issue.code) {
      "epics.startup.missingDbLoadRecordsMacros" ->
        listOfNotNull(createStartupLoadMacroQuickFix(project, psiFile, issue)).toTypedArray()
      "epics.database.duplicateRecordName" ->
        listOfNotNull(createDuplicateRecordNameQuickFix(psiFile, issue)).toTypedArray()
      "epics.database.invalidFieldName" ->
        listOfNotNull(createInvalidFieldNameQuickFix(psiFile, issue)).toTypedArray()
      "epics.database.invalidMenuFieldValue" ->
        listOfNotNull(createInvalidMenuFieldValueQuickFix(psiFile, issue)).toTypedArray()
      "epics.startup.unknownIocRegistrationFunction" ->
        createUnknownIocRegistrationQuickFixes(project, psiFile, issue).toTypedArray()
      else -> emptyArray()
    }
  }

  private fun createStartupLoadMacroQuickFix(
    project: Project,
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): LocalQuickFix? {
    val file = psiFile.virtualFile ?: return null
    val statement = findStartupLoadStatement(psiFile.text, issue.startOffset, issue.endOffset) ?: return null
    val resolved = EpicsPathResolver.resolveReference(project, file, statement.pathStart) ?: return null
    if (resolved.kind != EpicsPathKind.DATABASE) {
      return null
    }

    val targetText = readText(resolved.targetFile) ?: return null
    val requiredMacroNames = extractRequiredMacroNames(maskDatabaseComments(targetText))
    val providedMacroNames = extractAssignedMacroNames(statement.macros)
    val missingMacroNames = requiredMacroNames.filterNot(providedMacroNames::contains)
    if (missingMacroNames.isEmpty()) {
      return null
    }

    val edits = if (statement.macroValueStart != null && statement.macroValueEnd != null) {
      val existingMacros = statement.macros
      val separator = if (existingMacros.trim().isNotEmpty()) "," else ""
      listOf(
        TextEdit(
          startOffset = statement.macroValueStart,
          endOffset = statement.macroValueEnd,
          replacement = existingMacros + separator + buildDbLoadRecordsAssignmentLabel(missingMacroNames),
        ),
      )
    } else {
      listOf(
        TextEdit(
          startOffset = statement.pathEnd + 1,
          endOffset = statement.pathEnd + 1,
          replacement = ", \"" + buildDbLoadRecordsAssignmentLabel(missingMacroNames) + "\"",
        ),
      )
    }

    return DocumentEditQuickFix(
      name = "Add missing dbLoadRecords macros: ${missingMacroNames.joinToString(", ")}",
      edits = edits,
    )
  }

  private fun createDuplicateRecordNameQuickFix(
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): LocalQuickFix? {
    val currentName = psiFile.text.safeSubstring(issue.startOffset, issue.endOffset).trim()
    if (currentName.isEmpty()) {
      return null
    }
    val replacement = suggestUniqueRecordName(psiFile.text, currentName)
    return DocumentEditQuickFix(
      name = "Rename duplicate record to \"$replacement\"",
      edits = listOf(TextEdit(issue.startOffset, issue.endOffset, replacement)),
    )
  }

  private fun createInvalidFieldNameQuickFix(
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): LocalQuickFix? {
    val text = psiFile.text
    val invalidFieldName = text.safeSubstring(issue.startOffset, issue.endOffset).trim()
    if (invalidFieldName.isEmpty()) {
      return null
    }
    val recordType = EpicsRecordCompletionSupport.findEnclosingRecordType(text, issue.startOffset) ?: return null
    val replacement = findBestMatchingLabel(
      EpicsRecordCompletionSupport.getFieldNamesForRecordType(recordType).filter { it != invalidFieldName },
      invalidFieldName,
    ) ?: return null
    return DocumentEditQuickFix(
      name = "Replace invalid field with \"$replacement\"",
      edits = listOf(TextEdit(issue.startOffset, issue.endOffset, replacement)),
    )
  }

  private fun createInvalidMenuFieldValueQuickFix(
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): LocalQuickFix? {
    val text = psiFile.text
    val invalidValue = text.safeSubstring(issue.startOffset, issue.endOffset)
    val context = EpicsRecordCompletionSupport.findMenuFieldValueContext(text, issue.startOffset) ?: return null
    val replacement = findBestMatchingLabel(
      context.choices.filter { it != invalidValue },
      invalidValue,
    ) ?: EpicsRecordCompletionSupport.getDefaultFieldValue(context.recordType, context.fieldName)
    if (replacement.isBlank()) {
      return null
    }
    return DocumentEditQuickFix(
      name = "Replace with menu value \"$replacement\"",
      edits = listOf(TextEdit(issue.startOffset, issue.endOffset, replacement)),
    )
  }

  private fun createUnknownIocRegistrationQuickFixes(
    project: Project,
    psiFile: PsiFile,
    issue: EpicsDatabaseValueValidator.ValidationIssue,
  ): List<LocalQuickFix> {
    val file = psiFile.virtualFile ?: return emptyList()
    val currentName = psiFile.text.safeSubstring(issue.startOffset, issue.endOffset).trim()
    if (currentName.isEmpty()) {
      return emptyList()
    }
    return collectKnownIocRegistrationFunctions(project, file)
      .filter { it != currentName }
      .map { replacement ->
        DocumentEditQuickFix(
          name = "Replace with \"$replacement\"",
          edits = listOf(TextEdit(issue.startOffset, issue.endOffset, replacement)),
        )
      }
  }

  private fun collectKnownIocRegistrationFunctions(project: Project, file: VirtualFile): List<String> {
    val ownerRoot = EpicsPathResolver.findOwningEpicsRoot(project, file)
    return project.epicsBuildModelService()
      .loadBuildModel(ownerRoot)
      ?.iocs
      ?.map { "${it.name}_registerRecordDeviceDriver" }
      ?.distinct()
      ?.sortedWith(String.CASE_INSENSITIVE_ORDER)
      .orEmpty()
  }

  private fun findStartupLoadStatement(
    text: String,
    pathStart: Int,
    pathEnd: Int,
  ): StartupLoadStatement? {
    var lineOffset = 0
    for (line in text.split('\n')) {
      val match = DB_LOAD_RECORDS_REGEX.find(line)
      if (match != null) {
        val pathValue = match.groups[1]?.value.orEmpty()
        val localPathStart = line.indexOf(pathValue)
        val absolutePathStart = lineOffset + localPathStart
        val absolutePathEnd = absolutePathStart + pathValue.length
        if (absolutePathStart == pathStart && absolutePathEnd == pathEnd) {
          val macrosGroup = match.groups[2]
          return StartupLoadStatement(
            pathStart = absolutePathStart,
            pathEnd = absolutePathEnd,
            macros = macrosGroup?.value.orEmpty(),
            macroValueStart = macrosGroup?.let { lineOffset + it.range.first },
            macroValueEnd = macrosGroup?.let { lineOffset + it.range.last + 1 },
          )
        }
      }
      lineOffset += line.length + 1
    }
    return null
  }

  private fun buildDbLoadRecordsAssignmentLabel(macroNames: List<String>): String {
    return macroNames.joinToString(",") { macroName -> "$macroName=" }
  }

  private fun extractAssignedMacroNames(text: String): Set<String> {
    if (text.isBlank()) {
      return emptySet()
    }

    val names = linkedSetOf<String>()
    var segmentStart = 0
    var escaped = false

    fun flush(segmentEnd: Int) {
      val segment = text.substring(segmentStart, segmentEnd)
      val match = NAMED_ASSIGNMENT_REGEX.find(segment) ?: return
      val name = match.groups[1]?.value.orEmpty()
      if (name.isNotBlank()) {
        names += name
      }
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
    return names
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

  private fun suggestUniqueRecordName(text: String, recordName: String): String {
    val existingNames = EpicsRecordCompletionSupport.extractRecordDeclarations(text)
      .mapTo(linkedSetOf()) { declaration -> declaration.name }
    var suffix = 1
    var candidate = "${recordName}_$suffix"
    while (candidate in existingNames) {
      suffix += 1
      candidate = "${recordName}_$suffix"
    }
    return candidate
  }

  private fun findBestMatchingLabel(labels: List<String>, target: String): String? {
    val normalizedTarget = target.trim()
    if (normalizedTarget.isEmpty()) {
      return null
    }
    val uniqueLabels = labels.filter { it.isNotBlank() }.distinct()
    if (uniqueLabels.isEmpty()) {
      return null
    }
    return uniqueLabels
      .map { label ->
        label to computeLevenshteinDistance(label.uppercase(), normalizedTarget.uppercase())
      }
      .sortedWith(compareBy<Pair<String, Int>>({ it.second }, { it.first }))
      .firstOrNull()
      ?.first
  }

  private fun computeLevenshteinDistance(left: String, right: String): Int {
    val matrix = Array(left.length + 1) { IntArray(right.length + 1) }
    for (row in 0..left.length) {
      matrix[row][0] = row
    }
    for (column in 0..right.length) {
      matrix[0][column] = column
    }
    for (row in 1..left.length) {
      for (column in 1..right.length) {
        val substitutionCost = if (left[row - 1] == right[column - 1]) 0 else 1
        matrix[row][column] = minOf(
          matrix[row - 1][column] + 1,
          matrix[row][column - 1] + 1,
          matrix[row - 1][column - 1] + substitutionCost,
        )
      }
    }
    return matrix[left.length][right.length]
  }

  private fun readText(file: VirtualFile): String? {
    return try {
      String(file.contentsToByteArray(), Charset.forName(file.charset.name()))
    } catch (_: Exception) {
      null
    }
  }

  private fun String.safeSubstring(startOffset: Int, endOffset: Int): String {
    val safeStart = startOffset.coerceIn(0, length)
    val safeEnd = endOffset.coerceAtLeast(safeStart).coerceIn(0, length)
    return substring(safeStart, safeEnd)
  }

  private data class TextEdit(
    val startOffset: Int,
    val endOffset: Int,
    val replacement: String,
  )

  private data class StartupLoadStatement(
    val pathStart: Int,
    val pathEnd: Int,
    val macros: String,
    val macroValueStart: Int?,
    val macroValueEnd: Int?,
  )

  private class DocumentEditQuickFix(
    private val name: String,
    private val edits: List<TextEdit>,
  ) : LocalQuickFix {
    override fun getName(): String = name

    override fun getFamilyName(): String = name

    override fun applyFix(project: Project, descriptor: ProblemDescriptor) {
      val psiFile = descriptor.psiElement.containingFile ?: return
      val document = psiFile.viewProvider.document ?: return
      WriteCommandAction.runWriteCommandAction(project, name, null, Runnable {
        edits.sortedByDescending { it.startOffset }.forEach { edit ->
          val safeStart = edit.startOffset.coerceIn(0, document.textLength)
          val safeEnd = edit.endOffset.coerceAtLeast(safeStart).coerceIn(0, document.textLength)
          document.replaceString(safeStart, safeEnd, edit.replacement)
        }
      }, psiFile)
    }
  }

  private val DB_LOAD_RECORDS_REGEX =
    Regex("""^\s*dbLoadRecords\(\s*"([^"\n]+)"(?:\s*,\s*"((?:[^"\\]|\\.)*)")?""")
  private val NAMED_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=""")
  private val REQUIRED_DATABASE_MACRO_REGEX =
    Regex("""\$\(([^)=,\s]+)(?:=([^)]*))?\)|\$\{([^}=,\s]+)(?:=([^}]*))?\}""")
}
