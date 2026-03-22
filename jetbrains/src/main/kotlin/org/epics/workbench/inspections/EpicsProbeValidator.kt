package org.epics.workbench.inspections

import org.epics.workbench.probe.EpicsProbeSupport

internal object EpicsProbeValidator {
  fun collectIssues(text: String): List<EpicsDatabaseValueValidator.ValidationIssue> {
    return EpicsProbeSupport.analyzeText(text).issues
  }
}
