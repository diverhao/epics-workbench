package org.epics.workbench.runtime

import com.intellij.openapi.components.Service
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.LocalFileSystem
import kotlinx.serialization.SerialName
import kotlinx.serialization.Serializable
import kotlinx.serialization.encodeToString
import kotlinx.serialization.json.Json
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.exists
import kotlin.io.path.readText
import kotlin.io.path.writeText

data class EpicsRuntimeProjectConfiguration(
  val protocol: EpicsRuntimeProtocol = EpicsRuntimeProtocol.CA,
  val caAddrList: List<String> = emptyList(),
  val caAutoAddrList: EpicsCaAutoAddrList = EpicsCaAutoAddrList.YES,
)

enum class EpicsRuntimeProtocol(
  val displayName: String,
  val serializedValue: String,
) {
  CA("Channel Access", "ca"),
  PVA("PV Access", "pva"),
  ;

  override fun toString(): String = displayName

  companion object {
    fun fromSerializedValue(value: String?): EpicsRuntimeProtocol {
      return entries.firstOrNull { it.serializedValue.equals(value, ignoreCase = true) } ?: CA
    }
  }
}

enum class EpicsCaAutoAddrList(
  val displayName: String,
  val serializedValue: String,
) {
  YES("Yes", "Yes"),
  NO("No", "No"),
  ;

  override fun toString(): String = displayName

  companion object {
    fun fromSerializedValue(value: String?): EpicsCaAutoAddrList {
      return entries.firstOrNull { it.serializedValue.equals(value, ignoreCase = true) } ?: YES
    }
  }
}

@Service(Service.Level.PROJECT)
class EpicsRuntimeProjectConfigurationService(
  private val project: Project,
) {
  fun loadConfiguration(): EpicsRuntimeProjectConfiguration {
    val configPath = getConfigurationPath() ?: return EpicsRuntimeProjectConfiguration()
    if (!configPath.exists()) {
      return EpicsRuntimeProjectConfiguration()
    }

    return runCatching {
      val raw = configPath.readText()
      normalizeConfiguration(json.decodeFromString<EpicsRuntimeProjectConfigurationFile>(raw))
    }.getOrElse {
      EpicsRuntimeProjectConfiguration()
    }
  }

  fun saveConfiguration(configuration: EpicsRuntimeProjectConfiguration) {
    val configPath = getConfigurationPath()
      ?: throw IllegalStateException("The current project does not have a filesystem location.")
    Files.createDirectories(configPath.parent)
    configPath.writeText(
      json.encodeToString(
        EpicsRuntimeProjectConfigurationFile(
          protocol = configuration.protocol.takeIf { it == EpicsRuntimeProtocol.PVA }?.serializedValue,
          caAddrList = configuration.caAddrList,
          caAutoAddrList = configuration.caAutoAddrList.serializedValue,
        ),
      ),
    )
    LocalFileSystem.getInstance().refreshNioFiles(listOf(configPath))
  }

  fun getConfigurationPath(): Path? {
    val basePath = project.basePath ?: return null
    return Path.of(basePath, CONFIG_FILE_NAME)
  }

  companion object {
    const val CONFIG_FILE_NAME: String = ".epics-workbench-config.json"

    private val json = Json {
      prettyPrint = true
      encodeDefaults = true
      ignoreUnknownKeys = true
      explicitNulls = false
    }

    private fun normalizeConfiguration(
      file: EpicsRuntimeProjectConfigurationFile,
    ): EpicsRuntimeProjectConfiguration {
      return EpicsRuntimeProjectConfiguration(
        protocol = EpicsRuntimeProtocol.fromSerializedValue(file.protocol),
        caAddrList = file.caAddrList.orEmpty().map(String::trim).filter(String::isNotEmpty),
        caAutoAddrList = EpicsCaAutoAddrList.fromSerializedValue(file.caAutoAddrList),
      )
    }
  }
}

@Serializable
private data class EpicsRuntimeProjectConfigurationFile(
  val protocol: String? = null,
  @SerialName("EPICS_CA_ADDR_LIST")
  val caAddrList: List<String>? = null,
  @SerialName("EPICS_CA_AUTO_ADDR_LIST")
  val caAutoAddrList: String? = null,
)
