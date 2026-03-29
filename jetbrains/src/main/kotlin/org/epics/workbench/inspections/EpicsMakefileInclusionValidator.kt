package org.epics.workbench.inspections

import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.findLocalMakefile
import org.epics.workbench.build.getMakefileDbInstallToken
import org.epics.workbench.build.parseMakeAssignments
import org.epics.workbench.build.readCurrentText
import java.nio.file.Path
import kotlin.io.path.name
import kotlin.io.path.pathString

internal object EpicsMakefileInclusionValidator {
  fun collectIssues(file: VirtualFile, text: String): List<EpicsDatabaseValueValidator.ValidationIssue> {
    val makefile = findLocalMakefile(file) ?: return emptyList()
    val installToken = getMakefileDbInstallToken(file) ?: return emptyList()
    val makefileText = readCurrentText(makefile) ?: return emptyList()
    val installedTokens = parseMakeAssignments(makefileText)["DB"].orEmpty()
    if (installedTokens.any { it == installToken }) {
      return emptyList()
    }

    if (isDatabaseTemplateFile(file) && isReferencedBySiblingSubstitutionsFile(file)) {
      return emptyList()
    }

    val highlightEnd = if (text.isEmpty()) 0 else 1
    return listOf(
      EpicsDatabaseValueValidator.ValidationIssue(
        startOffset = 0,
        endOffset = highlightEnd,
        message = "This file is not included in Makefile.",
        code = "epics.makefile.notIncluded",
        severity = EpicsDatabaseValueValidator.ValidationSeverity.WARNING,
      ),
    )
  }

  private fun isReferencedBySiblingSubstitutionsFile(file: VirtualFile): Boolean {
    val targetPath = runCatching { file.toNioPath().normalize() }.getOrNull() ?: return false
    val directory = file.parent ?: return false
    for (sibling in directory.children) {
      if (!isSubstitutionsFile(sibling) || sibling.path == file.path) {
        continue
      }
      val siblingText = readCurrentText(sibling) ?: continue
      for (templatePath in extractSubstitutionTemplatePaths(siblingText)) {
        if (doesTemplatePathReferenceTarget(sibling, templatePath, targetPath)) {
          return true
        }
      }
    }
    return false
  }

  private fun doesTemplatePathReferenceTarget(
    substitutionsFile: VirtualFile,
    templatePath: String,
    targetPath: Path,
  ): Boolean {
    val rawTemplatePath = templatePath.trim().removeSurrounding("\"")
    if (rawTemplatePath.isEmpty()) {
      return false
    }

    val directCandidate = runCatching {
      val candidate = Path.of(rawTemplatePath)
      if (candidate.isAbsolute) {
        candidate.normalize()
      } else {
        substitutionsFile.parent.toNioPath().resolve(rawTemplatePath).normalize()
      }
    }.getOrNull()
    if (directCandidate?.pathString == targetPath.pathString) {
      return true
    }

    val basename = runCatching { Path.of(rawTemplatePath).name }.getOrDefault(rawTemplatePath.substringAfterLast('/').substringAfterLast('\\'))
    return basename == targetPath.name
  }

  private fun extractSubstitutionTemplatePaths(text: String): List<String> {
    return FILE_BLOCK_REGEX.findAll(text)
      .mapNotNull { match -> match.groups[1]?.value?.trim()?.removeSurrounding("\"") }
      .filter(String::isNotBlank)
      .toList()
  }

  private fun isDatabaseTemplateFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("db", "template")
  }

  private fun isSubstitutionsFile(file: VirtualFile): Boolean {
    return file.extension?.lowercase() in setOf("substitutions", "sub", "subs")
  }

  private val FILE_BLOCK_REGEX = Regex("""(?:^|\n)\s*file(?:\s+("(?:[^"\\]|\\.)*"|[^\s{]+))?\s*\{""")
}
