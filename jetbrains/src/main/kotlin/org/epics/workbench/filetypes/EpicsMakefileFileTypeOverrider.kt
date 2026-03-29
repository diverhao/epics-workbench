package org.epics.workbench.filetypes

import com.intellij.openapi.fileTypes.FileType
import com.intellij.openapi.fileTypes.FileTypeManager
import com.intellij.openapi.fileTypes.UnknownFileType
import com.intellij.openapi.fileTypes.impl.FileTypeOverrider
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.formatting.isMakefileStyleFile

class EpicsMakefileFileTypeOverrider : FileTypeOverrider {
  override fun getOverriddenFileType(file: VirtualFile): FileType? {
    if (!isMakefileStyleFile(file) || file.name == "Makefile") {
      return null
    }

    val makefileType = FileTypeManager.getInstance().getFileTypeByFileName("Makefile")
    return if (makefileType == UnknownFileType.INSTANCE) null else makefileType
  }
}
