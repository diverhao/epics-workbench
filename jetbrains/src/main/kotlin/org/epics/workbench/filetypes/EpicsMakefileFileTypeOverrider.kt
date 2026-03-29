package org.epics.workbench.filetypes

import com.intellij.openapi.fileTypes.impl.FileTypeOverrider
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.formatting.isMakefileStyleFile

class EpicsMakefileFileTypeOverrider : FileTypeOverrider {
  override fun getOverriddenFileType(file: VirtualFile) =
    if (isMakefileStyleFile(file) && file.name != "Makefile") {
      EpicsMakefileFileType.INSTANCE
    } else {
      null
    }
}
