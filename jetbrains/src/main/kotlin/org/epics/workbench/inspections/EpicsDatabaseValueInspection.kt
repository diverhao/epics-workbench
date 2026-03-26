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
        val issues = when {
          isDatabaseFile(file.name) -> EpicsDatabaseValueValidator.collectIssues(file.text)
          isStartupFile(file.name) -> {
            val virtualFile = file.virtualFile ?: return
            EpicsStartupMacroValidator.collectIssues(file.project, virtualFile, file.text)
          }

          else -> emptyList()
        }
        if (issues.isEmpty()) {
          return
        }

        for (issue in issues) {
          val quickFixes = EpicsInspectionQuickFixSupport.buildQuickFixes(file.project, file, issue)
          holder.registerProblem(
            file,
            issue.message,
            ProblemHighlightType.GENERIC_ERROR,
            file.textRange(issue.startOffset, issue.endOffset),
            *quickFixes,
          )
        }
      }
    }
  }

  private fun isDatabaseFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("db", "vdb", "template")
  }

  private fun isStartupFile(fileName: String): Boolean {
    val extension = fileName.substringAfterLast('.', "").lowercase()
    return extension in setOf("cmd", "iocsh") || fileName == "st.cmd"
  }
}

private fun PsiFile.textRange(start: Int, end: Int) =
  com.intellij.openapi.util.TextRange(
    start.coerceIn(0, textLength),
    end.coerceAtLeast(start).coerceIn(0, textLength),
  )
