package org.epics.workbench.editor

import com.intellij.codeInsight.editorActions.SimpleTokenSetQuoteHandler
import org.epics.workbench.highlighting.EpicsTokenTypes

class EpicsQuoteHandler : SimpleTokenSetQuoteHandler(EpicsTokenTypes.STRING)
