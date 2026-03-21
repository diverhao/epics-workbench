package org.epics.workbench.filetypes

import com.intellij.icons.AllIcons
import com.intellij.openapi.fileTypes.LanguageFileType
import com.intellij.openapi.util.NlsContexts.Label
import javax.swing.Icon

abstract class EpicsFileType(
  language: com.intellij.lang.Language,
  private val displayName: String,
  private val descriptionText: String,
  private val defaultExtensionText: String,
) : LanguageFileType(language) {
  override fun getName(): String = displayName

  override fun getDescription(): @Label String = descriptionText

  override fun getDefaultExtension(): String = defaultExtensionText

  override fun getIcon(): Icon = AllIcons.FileTypes.Text
}

class EpicsDatabaseFileType private constructor() : EpicsFileType(
  EpicsDatabaseLanguage.INSTANCE,
  "EPICS Database",
  "EPICS database and template files",
  "db",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsDatabaseFileType()
  }
}

class EpicsSubstitutionsFileType private constructor() : EpicsFileType(
  EpicsSubstitutionsLanguage.INSTANCE,
  "EPICS Substitutions",
  "EPICS substitutions files",
  "substitutions",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsSubstitutionsFileType()
  }
}

class EpicsStartupFileType private constructor() : EpicsFileType(
  EpicsStartupLanguage.INSTANCE,
  "EPICS Startup",
  "EPICS IOC shell startup files",
  "cmd",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsStartupFileType()
  }
}

class EpicsDatabaseDefinitionFileType private constructor() : EpicsFileType(
  EpicsDatabaseDefinitionLanguage.INSTANCE,
  "EPICS Database Definition",
  "EPICS database definition files",
  "dbd",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsDatabaseDefinitionFileType()
  }
}

class EpicsProtocolFileType private constructor() : EpicsFileType(
  EpicsProtocolLanguage.INSTANCE,
  "EPICS Protocol",
  "EPICS protocol files",
  "proto",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsProtocolFileType()
  }
}

class EpicsSequencerFileType private constructor() : EpicsFileType(
  EpicsSequencerLanguage.INSTANCE,
  "EPICS Sequencer",
  "EPICS sequencer files",
  "st",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsSequencerFileType()
  }
}

class EpicsMonitorFileType private constructor() : EpicsFileType(
  EpicsMonitorLanguage.INSTANCE,
  "EPICS Monitor",
  "EPICS monitor files",
  "monitor",
) {
  companion object {
    @JvmField
    val INSTANCE = EpicsMonitorFileType()
  }
}
