package org.epics.workbench.navigation

import com.intellij.codeInsight.navigation.actions.GotoDeclarationHandler
import com.intellij.openapi.actionSystem.DataContext
import com.intellij.openapi.editor.Editor
import com.intellij.openapi.fileEditor.OpenFileDescriptor
import com.intellij.openapi.project.Project
import com.intellij.openapi.util.TextRange
import com.intellij.psi.PsiFile
import com.intellij.psi.PsiElement
import com.intellij.psi.PsiManager
import com.intellij.psi.impl.FakePsiElement

class EpicsGotoDeclarationHandler : GotoDeclarationHandler {
  override fun getGotoDeclarationTargets(
    sourceElement: PsiElement?,
    offset: Int,
    editor: Editor,
  ): Array<PsiElement> {
    val project = editor.project ?: sourceElement?.project ?: return PsiElement.EMPTY_ARRAY
    val hostFile = sourceElement?.containingFile?.virtualFile ?: return PsiElement.EMPTY_ARRAY

    EpicsRecordResolver.resolveRecordDefinition(project, hostFile, offset)?.let { definition ->
      val psiFile = PsiManager.getInstance(project).findFile(definition.targetFile) ?: return PsiElement.EMPTY_ARRAY
      return arrayOf(
        EpicsNavigatableTargetElement(
          psiFile = psiFile,
          offset = definition.recordStartOffset,
        ),
      )
    }

    val target = EpicsPathResolver.resolveReferencedFile(project, hostFile, offset) ?: return PsiElement.EMPTY_ARRAY
    val psiTarget = PsiManager.getInstance(project).findFile(target) ?: return PsiElement.EMPTY_ARRAY
    return arrayOf(psiTarget)
  }

  override fun getActionText(context: DataContext): String = "Go to EPICS Definition"
}

private class EpicsNavigatableTargetElement(
  private val psiFile: PsiFile,
  private val offset: Int,
) : FakePsiElement() {
  override fun getProject(): Project = psiFile.project

  override fun getContainingFile(): PsiFile = psiFile

  override fun getParent(): PsiElement = psiFile

  override fun getTextOffset(): Int = offset

  override fun getTextRange(): TextRange = TextRange(offset, offset)

  override fun getNavigationElement(): PsiElement = this

  override fun navigate(requestFocus: Boolean) {
    OpenFileDescriptor(project, psiFile.virtualFile, offset).navigate(requestFocus)
  }

  override fun canNavigate(): Boolean = true

  override fun canNavigateToSource(): Boolean = true
}
