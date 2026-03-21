package org.epics.workbench.inspections

import org.epics.workbench.completion.EpicsRecordCompletionSupport
import java.util.ArrayDeque

internal object EpicsDatabaseValueValidator {
  internal data class ValidationIssue(
    val startOffset: Int,
    val endOffset: Int,
    val message: String,
  )

  fun collectIssues(text: String): List<ValidationIssue> {
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
        )
      }
    }

    return issues
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
