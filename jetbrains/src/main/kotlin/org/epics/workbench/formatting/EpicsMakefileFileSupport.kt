package org.epics.workbench.formatting

import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiFile

internal fun isMakefileStyleFile(file: PsiFile): Boolean {
  val virtualFile = file.virtualFile
  return if (virtualFile != null) {
    isMakefileStyleFile(virtualFile)
  } else {
    file.name == "Makefile"
  }
}

internal fun isMakefileStyleFile(file: VirtualFile?): Boolean {
  val target = file ?: return false
  if (!target.isValid || target.isDirectory) {
    return false
  }
  if (target.name == "Makefile") {
    return true
  }
  val parent = target.parent ?: return false
  if (parent.name != "configure") {
    return false
  }
  return isEpicsReleaseFileName(target.name)
}

private fun isEpicsReleaseFileName(fileName: String): Boolean {
  return fileName == "RELEASE" || fileName.startsWith("RELEASE.")
}
