package org.epics.workbench.inspections

import com.intellij.codeInspection.LocalInspectionTool
import com.intellij.codeInspection.LocalInspectionToolSession
import com.intellij.codeInspection.ProblemsHolder
import com.intellij.codeInspection.ProblemHighlightType
import com.intellij.psi.PsiElementVisitor
import com.intellij.psi.PsiFile
class EpicsDatabaseValueInspection : LocalInspectionTool() {
  override fun buildVisitor(
    holder: ProblemsHolder,
    isOnTheFly: Boolean,
    session: LocalInspectionToolSession,
  ): PsiElementVisitor {
    return object : PsiElementVisitor() {
      override fun visitFile(file: PsiFile) {
        if (!isDatabaseFile(file.name)) {
          return
        }

        for (issue in EpicsDatabaseValueValidator.collectIssues(file.text)) {
          holder.registerProblem(
            file,
            issue.message,
            ProblemHighlightType.GENERIC_ERROR,
            file.textRange(issue.startOffset, issue.endOffset),
          )
        }
      }
    }
  }

  private fun isDatabaseFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("db", "vdb", "template")
  }
}

private fun PsiFile.textRange(start: Int, end: Int) =
  com.intellij.openapi.util.TextRange(
    start.coerceIn(0, textLength),
    end.coerceAtLeast(start).coerceIn(0, textLength),
  )
