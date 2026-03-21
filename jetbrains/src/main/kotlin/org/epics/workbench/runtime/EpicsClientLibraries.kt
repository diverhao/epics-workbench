package org.epics.workbench.runtime

import com.cosylab.epics.caj.CAJContext
import gov.aps.jca.Context
import org.epics.pva.client.PVAClient

/**
 * Thin bridge around the Java EPICS client libraries that are bundled into the plugin build.
 *
 * Runtime lifecycle and configuration are handled later. For now, this file keeps the APIs
 * available to the plugin code and verifies they compile in this project.
 */
object EpicsClientLibraries {
  fun createCaContext(): Context = CAJContext().apply {
    setDoNotShareChannels(true)
  }

  fun createPvaClient(): PVAClient = PVAClient()
}
