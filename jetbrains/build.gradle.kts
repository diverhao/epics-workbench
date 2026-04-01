import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

fun propertyOrEnv(propertyName: String, envName: String = propertyName): String? =
  providers.gradleProperty(propertyName).orNull ?: System.getenv(envName)

fun nonBlankPropertyOrEnv(propertyName: String, envName: String = propertyName): String? =
  propertyOrEnv(propertyName, envName)?.trim()?.takeIf { it.isNotEmpty() }

fun firstExistingFile(vararg candidates: String?): File? =
  candidates
    .asSequence()
    .filterNotNull()
    .map(String::trim)
    .filter(String::isNotEmpty)
    .map(::file)
    .firstOrNull(File::isFile)

val epicsJcaVersion = "2.4.7"
val phoebusPvaVersion = "5.0.2"
val pluginId = "org.epics.workbench"
val pluginName = "EPICS Workbench"
val pluginVersion = nonBlankPropertyOrEnv("pluginVersion", "PLUGIN_VERSION") ?: "0.1.0-SNAPSHOT"
val pluginVendorName = nonBlankPropertyOrEnv("pluginVendorName", "PLUGIN_VENDOR_NAME") ?: pluginName
val pluginVendorEmail = nonBlankPropertyOrEnv("pluginVendorEmail", "PLUGIN_VENDOR_EMAIL")
val pluginVendorUrl = nonBlankPropertyOrEnv("pluginVendorUrl", "PLUGIN_VENDOR_URL")
val pluginChangeNotes =
  nonBlankPropertyOrEnv("pluginChangeNotes", "PLUGIN_CHANGE_NOTES")
    ?: "Initial JetBrains plugin scaffold for EPICS Workbench."
val publishToken = nonBlankPropertyOrEnv("jetbrainsPublishToken", "PUBLISH_TOKEN")
val publishChannel = nonBlankPropertyOrEnv("jetbrainsPublishChannel", "JETBRAINS_PUBLISH_CHANNEL") ?: "default"
val publishHidden = nonBlankPropertyOrEnv("jetbrainsPublishHidden", "JETBRAINS_PUBLISH_HIDDEN")?.toBoolean() ?: false
val signingPrivateKeyFile = firstExistingFile(
  nonBlankPropertyOrEnv("jetbrainsSigningPrivateKeyFile", "PRIVATE_KEY_FILE"),
  "../dev-doc/private.pem",
  "../dev-doc/private_encrypted.pem",
)
val signingCertificateChainFile = firstExistingFile(
  nonBlankPropertyOrEnv("jetbrainsSigningCertificateChainFile", "CERTIFICATE_CHAIN_FILE"),
  "../dev-doc/chain.crt",
)
val signingPrivateKeyPassword = nonBlankPropertyOrEnv("jetbrainsSigningPrivateKeyPassword", "PRIVATE_KEY_PASSWORD")

plugins {
  id("java")
  id("org.jetbrains.kotlin.jvm") version "1.9.25"
  id("org.jetbrains.intellij.platform") version "2.3.0"
}

group = "org.epics.workbench"
version = pluginVersion

repositories {
  mavenCentral()
  maven(url = "https://s01.oss.sonatype.org/content/repositories/releases")
  intellijPlatform {
    defaultRepositories()
  }
}

dependencies {
  implementation("org.jetbrains.kotlinx:kotlinx-serialization-json:1.7.3")
  implementation("org.epics:jca:$epicsJcaVersion")
  implementation("org.phoebus:core-pva:$phoebusPvaVersion")

  intellijPlatform {
    create("IC", "2024.3.6")
    testFramework(org.jetbrains.intellij.platform.gradle.TestFrameworkType.Platform)
  }
}

intellijPlatform {
  pluginConfiguration {
    id = pluginId
    name = pluginName
    version = project.version.toString()
    ideaVersion {
      sinceBuild = "243"
    }
    vendor {
      name = pluginVendorName
      if (pluginVendorEmail != null) {
        email = pluginVendorEmail
      }
      if (pluginVendorUrl != null) {
        url = pluginVendorUrl
      }
    }
    changeNotes = pluginChangeNotes
  }

  signing {
    if (signingPrivateKeyFile != null) {
      privateKeyFile = signingPrivateKeyFile
    }
    if (signingCertificateChainFile != null) {
      certificateChainFile = signingCertificateChainFile
    }
    if (signingPrivateKeyPassword != null) {
      password = signingPrivateKeyPassword
    }
  }

  publishing {
    if (publishToken != null) {
      token = publishToken
    }
    channels = listOf(publishChannel)
    hidden = publishHidden
  }
}

tasks {
  processResources {
    from("LICENSE") {
      into("META-INF")
    }
    from("../vscode/data") {
      into("data")
    }
    from("../vscode/scripts/epics-build-model.js") {
      into("build-model")
    }
  }

  withType<JavaCompile> {
    sourceCompatibility = "21"
    targetCompatibility = "21"
  }

  withType<KotlinCompile> {
    kotlinOptions.jvmTarget = "21"
  }
}
