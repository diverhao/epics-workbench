import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

val epicsJcaVersion = "2.4.7"
val phoebusPvaVersion = "5.0.2"

plugins {
  id("java")
  id("org.jetbrains.kotlin.jvm") version "1.9.25"
  id("org.jetbrains.intellij.platform") version "2.3.0"
}

group = "org.epics.workbench"
version = "0.1.0-SNAPSHOT"

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
    ideaVersion {
      sinceBuild = "243"
    }
    changeNotes = """
      Initial JetBrains plugin scaffold for EPICS Workbench.
    """.trimIndent()
  }
}

tasks {
  processResources {
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
