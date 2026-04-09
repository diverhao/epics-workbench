package org.epics.workbench.inspections

import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver

internal object EpicsMakefileDatabaseValidator {
  fun collectIssues(
    project: Project,
    file: VirtualFile,
    text: String,
  ): List<EpicsDatabaseValueValidator.ValidationIssue> {
    if (file.name != "Makefile") {
      return emptyList()
    }

    return EpicsPathResolver.extractMakefileReferences(text)
      .filter { it.kind == EpicsPathKind.DATABASE || it.kind == EpicsPathKind.SUBSTITUTIONS }
      .mapNotNull { reference ->
        if (EpicsPathResolver.resolveReference(project, file, reference.startOffset) != null) {
          null
        } else {
          EpicsDatabaseValueValidator.ValidationIssue(
            startOffset = reference.startOffset,
            endOffset = reference.endOffset,
            message = buildDiagnosticMessage(project, file, reference.name),
            code = "epics.makefile.unknownDatabaseFile",
          )
        }
      }
  }

  private fun buildDiagnosticMessage(
    project: Project,
    file: VirtualFile,
    token: String,
  ): String {
    if (containsMakeVariableReference(token)) {
      val resolution = EpicsPathResolver.resolveMakefileTokenPath(project, file, token)
      if (resolution.absolutePath != null) {
        return "Unknown database/template file \"$token\". Expanded path \"${resolution.absolutePath}\" does not exist."
      }
      if (resolution.unresolvedVariables.isNotEmpty()) {
        return "Unknown database/template file \"$token\". Could not resolve make variables from the Makefile or configure/RELEASE: ${resolution.unresolvedVariables.joinToString(", ")}."
      }
      return "Unknown database/template file \"$token\". It was not found at the expanded path from local Makefile variables or configure/RELEASE."
    }

    return "Unknown database/template file \"$token\". It was not found beside the Makefile or in the module roots from configure/RELEASE."
  }

  private fun containsMakeVariableReference(value: String): Boolean {
    return MAKE_VARIABLE_REGEX.containsMatchIn(value)
  }

  private val MAKE_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
}
