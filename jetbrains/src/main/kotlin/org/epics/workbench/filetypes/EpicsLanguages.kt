package org.epics.workbench.filetypes

import com.intellij.lang.Language

class EpicsDatabaseLanguage private constructor() : Language("EPICS Database") {
  companion object {
    @JvmField
    val INSTANCE = EpicsDatabaseLanguage()
  }
}

class EpicsSubstitutionsLanguage private constructor() : Language("EPICS Substitutions") {
  companion object {
    @JvmField
    val INSTANCE = EpicsSubstitutionsLanguage()
  }
}

class EpicsStartupLanguage private constructor() : Language("EPICS Startup") {
  companion object {
    @JvmField
    val INSTANCE = EpicsStartupLanguage()
  }
}

class EpicsDatabaseDefinitionLanguage private constructor() : Language("EPICS Database Definition") {
  companion object {
    @JvmField
    val INSTANCE = EpicsDatabaseDefinitionLanguage()
  }
}

class EpicsProtocolLanguage private constructor() : Language("EPICS Protocol") {
  companion object {
    @JvmField
    val INSTANCE = EpicsProtocolLanguage()
  }
}

class EpicsSequencerLanguage private constructor() : Language("EPICS Sequencer") {
  companion object {
    @JvmField
    val INSTANCE = EpicsSequencerLanguage()
  }
}

class EpicsMonitorLanguage private constructor() : Language("EPICS PV List") {
  companion object {
    @JvmField
    val INSTANCE = EpicsMonitorLanguage()
  }
}

class EpicsProbeLanguage private constructor() : Language("EPICS Probe") {
  companion object {
    @JvmField
    val INSTANCE = EpicsProbeLanguage()
  }
}
