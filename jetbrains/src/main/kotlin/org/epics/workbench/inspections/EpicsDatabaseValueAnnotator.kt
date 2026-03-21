package org.epics.workbench.inspections

import com.intellij.lang.annotation.AnnotationHolder
import com.intellij.lang.annotation.Annotator
import com.intellij.lang.annotation.HighlightSeverity
import com.intellij.openapi.util.Key
import com.intellij.openapi.util.TextRange
import com.intellij.psi.PsiElement
import com.intellij.psi.PsiFile
import org.epics.workbench.completion.EpicsRecordCompletionSupport

class EpicsDatabaseValueAnnotator : Annotator {
  override fun annotate(element: PsiElement, holder: AnnotationHolder) {
    val file = element.containingFile ?: return
    if (!isDatabaseFile(file.name)) {
      return
    }
    val session = holder.currentAnnotationSession
    if (session.getUserData(PROCESSED_KEY) == true) {
      return
    }
    session.putUserData(PROCESSED_KEY, true)

    val text = file.text
    for (recordDeclaration in EpicsRecordCompletionSupport.extractRecordDeclarations(text)) {
      for (fieldDeclaration in EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, recordDeclaration)) {
        val dbfType = EpicsRecordCompletionSupport.getFieldType(recordDeclaration.recordType, fieldDeclaration.fieldName)
          ?: continue

        if (EpicsRecordCompletionSupport.isNumericFieldType(dbfType)) {
          if (EpicsRecordCompletionSupport.isSkippableNumericFieldValue(fieldDeclaration.value)) {
            continue
          }
          if (EpicsRecordCompletionSupport.isValidNumericFieldValue(fieldDeclaration.value, dbfType)) {
            continue
          }

          holder.newAnnotation(
            HighlightSeverity.ERROR,
            "Field \"${fieldDeclaration.fieldName}\" expects a $dbfType numeric value.",
          )
            .range(valueRange(file, fieldDeclaration.valueStart, fieldDeclaration.valueEnd))
            .create()
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

        holder.newAnnotation(
          HighlightSeverity.ERROR,
          "Field \"${fieldDeclaration.fieldName}\" must be one of the menu choices for \"${recordDeclaration.recordType}\".",
        )
          .range(valueRange(file, fieldDeclaration.valueStart, fieldDeclaration.valueEnd))
          .create()
      }
    }
  }

  private fun valueRange(file: PsiFile, start: Int, end: Int): TextRange {
    if (end > start) {
      return TextRange(start, end)
    }
    val safeStart = start.coerceAtMost(file.textLength)
    val safeEnd = (safeStart + 1).coerceAtMost(file.textLength)
    return TextRange(safeStart.coerceAtMost(safeEnd), safeEnd)
  }

  private fun isDatabaseFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("db", "vdb", "template")
  }

  companion object {
    private val PROCESSED_KEY = Key.create<Boolean>("org.epics.workbench.inspections.databaseValueAnnotator.processed")
  }
}
