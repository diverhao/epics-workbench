package org.epics.workbench.runtime

import com.cosylab.epics.caj.CAJContext
import gov.aps.jca.Context
import gov.aps.jca.configuration.DefaultConfiguration
import org.epics.pva.client.PVAClient

/**
 * Thin bridge around the Java EPICS client libraries that are bundled into the plugin build.
 *
 * Runtime lifecycle and configuration are handled later. For now, this file keeps the APIs
 * available to the plugin code and verifies they compile in this project.
 */
object EpicsClientLibraries {
  fun createCaContext(configuration: EpicsRuntimeProjectConfiguration): Context = CAJContext().apply {
    configure(
      DefaultConfiguration("CAJContext").apply {
        setAttribute("addr_list", configuration.caAddrList.joinToString(" "))
        setAttribute(
          "auto_addr_list",
          if (configuration.caAutoAddrList == EpicsCaAutoAddrList.YES) "true" else "false",
        )
      },
    )
    setDoNotShareChannels(true)
  }

  fun createPvaClient(@Suppress("UNUSED_PARAMETER") configuration: EpicsRuntimeProjectConfiguration): PVAClient = PVAClient()
}
