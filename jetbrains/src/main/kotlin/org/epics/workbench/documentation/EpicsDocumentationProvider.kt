package org.epics.workbench.documentation

import com.intellij.lang.documentation.AbstractDocumentationProvider
import com.intellij.openapi.fileEditor.FileDocumentManager
import com.intellij.openapi.editor.colors.EditorColorsManager
import com.intellij.openapi.editor.colors.TextAttributesKey
import com.intellij.openapi.project.Project
import com.intellij.openapi.util.text.StringUtil
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.psi.PsiElement
import com.intellij.psi.PsiFile
import com.intellij.psi.PsiManager
import com.intellij.psi.impl.FakePsiElement
import com.intellij.psi.tree.IElementType
import org.epics.workbench.build.EpicsBuildModelService
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.formatting.isMakefileStyleFile
import org.epics.workbench.highlighting.EpicsHighlightingKeys
import org.epics.workbench.highlighting.EpicsLexingProfile
import org.epics.workbench.highlighting.EpicsSimpleLexer
import org.epics.workbench.highlighting.EpicsTokenTypes
import org.epics.workbench.highlighting.PROTOCOL_KEYWORDS
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.navigation.EpicsResolvedReference
import org.epics.workbench.navigation.EpicsResolvedRecordDefinition
import org.epics.workbench.protocol.EpicsStreamProtocolSupport
import org.epics.workbench.protocol.StreamProtocolCommandReference
import org.epics.workbench.toc.EpicsDatabaseToc
import java.awt.Color
import java.io.File
import java.net.URLEncoder
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.exists
import kotlin.io.path.isDirectory
import kotlin.io.path.isRegularFile
import kotlin.io.path.name
import kotlin.io.path.pathString

class EpicsDocumentationProvider : AbstractDocumentationProvider() {
  override fun getCustomDocumentationElement(
    editor: com.intellij.openapi.editor.Editor,
    file: PsiFile,
    contextElement: PsiElement?,
    targetOffset: Int,
  ): PsiElement? {
    return createDocumentationElement(file, targetOffset)
  }

  override fun generateDoc(element: PsiElement?, originalElement: PsiElement?): String? {
    val referencesElement = element as? EpicsReferencedFilesElement
    if (referencesElement != null) {
      val hostFileName = referencesElement.hostFile.virtualFile?.name ?: referencesElement.hostFile.name
      return buildDocumentation(referencesElement.references, hostFileName)
    }
    val referenceElement = element as? EpicsReferencedFileElement ?: return null
    val hostFileName = referenceElement.hostFile.virtualFile?.name ?: referenceElement.hostFile.name
    return buildDocumentation(referenceElement.reference, hostFileName)
  }

  override fun generateHoverDoc(element: PsiElement, originalElement: PsiElement?): String? {
    return generateDoc(element, originalElement)
  }

  companion object {
    internal fun createDocumentationElement(file: PsiFile, offset: Int): EpicsReferencedFileElement? {
      val virtualFile = file.virtualFile ?: return null
      val startupTraceReferences = EpicsPathResolver.resolveStartupDbLoadRecordsTraceReferences(file.project, virtualFile, offset)
      if (startupTraceReferences.size > 1) {
        return EpicsReferencedFilesElement(file.manager, file, startupTraceReferences)
      }
      if (startupTraceReferences.size == 1) {
        return EpicsReferencedFileElement(file.manager, file, startupTraceReferences.first())
      }
      if (isSubstitutionsFile(virtualFile)) {
        val resolvedReferences = EpicsPathResolver.resolveSubstitutionsReferences(file.project, virtualFile, offset)
        if (resolvedReferences.size > 1) {
          return EpicsReferencedFilesElement(file.manager, file, resolvedReferences)
        }
      }
      val resolved = EpicsPathResolver.resolveReference(file.project, virtualFile, offset) ?: return null
      return EpicsReferencedFileElement(file.manager, file, resolved)
    }

    internal fun createDocumentationPreview(project: Project, hostFile: VirtualFile, offset: Int): EpicsDocumentationPreview? {
      if (hostFile.extension?.lowercase() in DATABASE_EXTENSIONS) {
        val text = readText(hostFile)
        if (text != null) {
          findDatabaseFieldValueContext(text, offset)
            ?.takeIf { it.fieldName == "DTYP" }
            ?.let { context ->
              val matches = resolveDtypDeviceHoverMatches(project, hostFile, context)
              if (matches.isNotEmpty()) {
                val referenceKey = buildDtypReferenceKey(hostFile, context, matches)
                return EpicsDocumentationPreview(
                  referenceKey,
                  buildDtypDeviceDocumentation(context, matches),
                )
              }
            }

          EpicsStreamProtocolSupport.findCommandReferenceAtOffset(text, offset)?.let { context ->
            val matches = resolveStreamProtocolCommandHoverMatches(project, hostFile, context)
            if (matches.isNotEmpty()) {
              val referenceKey = buildStreamProtocolCommandReferenceKey(hostFile, context, matches)
              return EpicsDocumentationPreview(
                referenceKey,
                buildStreamProtocolCommandDocumentation(context, matches),
              )
            }
          }

          EpicsRecordCompletionSupport.findMenuFieldValueContext(text, offset)?.let { context ->
            val referenceKey = "${hostFile.path}:menu:${context.recordName}:${context.fieldName}:${context.valueStart}:${context.valueEnd}"
            return EpicsDocumentationPreview(referenceKey, buildMenuFieldDocumentation(hostFile, context))
          }

          EpicsDatabaseToc.findRecordReferenceAtTypeOffset(text, offset)?.let { tocReference ->
            EpicsRecordResolver.resolveRecordDefinitionInFile(
              hostFile,
              tocReference.recordName,
              tocReference.recordType,
            )?.let { definition ->
              val referenceKey = "${hostFile.path}:toc-type:${definition.recordName}:${definition.line}"
              return EpicsDocumentationPreview(referenceKey, buildRecordDocumentation(definition, hostFile))
            }
          }
        }
      }

      if (hostFile.extension?.lowercase() == PVLIST_EXTENSION) {
        val text = readText(hostFile)
        val recordName = text?.let { findPvlistHoverRecordName(it, offset) }
        if (!recordName.isNullOrBlank()) {
          val definitions = EpicsRecordResolver.resolveRecordDefinitionsForName(project, hostFile, recordName)
          if (definitions.isNotEmpty()) {
            val referenceKey = buildRecordReferenceKey(hostFile, definitions)
            return EpicsDocumentationPreview(referenceKey, buildRecordDocumentation(definitions, hostFile))
          }
        }
      }

      findVariableDocumentationPreview(project, hostFile, offset)?.let { return it }

      if (isSubstitutionsFile(hostFile)) {
        val resolvedReferences = EpicsPathResolver.resolveSubstitutionsReferences(project, hostFile, offset)
        if (resolvedReferences.isNotEmpty()) {
          val referenceKey = buildReferenceKey(hostFile, resolvedReferences)
          return EpicsDocumentationPreview(referenceKey, buildDocumentation(resolvedReferences, hostFile.name))
        }
      }

      val startupTraceReferences = EpicsPathResolver.resolveStartupDbLoadRecordsTraceReferences(project, hostFile, offset)
      if (startupTraceReferences.isNotEmpty()) {
        val referenceKey = buildReferenceKey(hostFile, startupTraceReferences)
        return EpicsDocumentationPreview(referenceKey, buildDocumentation(startupTraceReferences, hostFile.name))
      }

      EpicsPathResolver.resolveReference(project, hostFile, offset)?.let { resolved ->
        val referenceKey = "${hostFile.path}:${resolved.rawPath}:${resolved.targetFile.path}"
        return EpicsDocumentationPreview(referenceKey, buildDocumentation(resolved, hostFile.name))
      }

      EpicsRecordResolver.resolveRecordDefinitions(project, hostFile, offset)
        .takeIf { it.isNotEmpty() }
        ?.let { definitions ->
          val referenceKey = buildRecordReferenceKey(hostFile, definitions)
          return EpicsDocumentationPreview(referenceKey, buildRecordDocumentation(definitions, hostFile))
        }

      return null
    }

    private fun findVariableDocumentationPreview(
      project: Project,
      hostFile: VirtualFile,
      offset: Int,
    ): EpicsDocumentationPreview? {
      if (!isMakefileStyleFile(hostFile) && !isStartupStateFile(hostFile)) {
        return null
      }

      val text = readText(hostFile) ?: return null
      val reference = findVariableReferenceAtOffset(text, offset) ?: return null
      val resolution = when {
        isMakefileStyleFile(hostFile) ->
          resolveMakefileVariableDocumentation(project, hostFile, text, reference.variableName)

        isStartupStateFile(hostFile) ->
          resolveStartupVariableDocumentation(project, hostFile, text, reference.startOffset, reference.variableName)

        else -> null
      } ?: return null

      val referenceKey = buildString {
        append(hostFile.path)
        append(":variable:")
        append(reference.variableName)
        append(':')
        append(resolution.resolvedValue)
        resolution.sourceInfo?.sourcePath?.let {
          append(':')
          append(it)
        }
        resolution.sourceInfo?.line?.let {
          append(':')
          append(it)
        }
      }
      return EpicsDocumentationPreview(referenceKey, buildVariableDocumentation(reference.variableName, resolution))
    }

    private fun resolveMakefileVariableDocumentation(
      project: Project,
      hostFile: VirtualFile,
      hostText: String,
      variableName: String,
    ): VariableHoverResolution? {
      val ownerRoot = EpicsPathResolver.findOwningEpicsRoot(project, hostFile)
      val releaseState = loadReleaseVariableState(ownerRoot)
      val hostDefinitions = linkedMapOf<String, VariableDefinition>()
      applyMakefileVariableDefinitions(
        text = hostText,
        sourcePath = hostFile.path,
        sourceKind = if (isEpicsReleaseFile(hostFile)) "RELEASE" else "Makefile",
        baseDirectory = hostFile.parent?.toNioPath() ?: ownerRoot,
        target = hostDefinitions,
      )

      val cache = linkedMapOf<String, VariableHoverResolution?>()
      val resolving = linkedSetOf<String>()

      fun resolve(name: String): VariableHoverResolution? {
        if (cache.containsKey(name)) {
          return cache[name]
        }
        if (!resolving.add(name)) {
          return null
        }

        val definition = hostDefinitions[name] ?: releaseState.definitions[name]
        val rawValue = when {
          definition != null -> definition.rawValue
          releaseState.resolvedValues.containsKey(name) -> releaseState.resolvedValues[name].orEmpty()
          else -> System.getenv(name)
        }

        if (rawValue == null) {
          resolving.remove(name)
          cache[name] = null
          return null
        }

        val expandedValue = expandDocumentationVariableValue(rawValue) { nestedName ->
          resolve(nestedName)?.resolvedValue
            ?: releaseState.resolvedValues[nestedName]
            ?: System.getenv(nestedName)
        }
        val absolutePath = computeAbsoluteVariablePath(
          expandedValue,
          definition?.baseDirectory ?: (hostFile.parent?.toNioPath() ?: ownerRoot),
        )
        val result = VariableHoverResolution(
          resolvedValue = absolutePath?.pathString ?: expandedValue,
          absolutePath = absolutePath,
          sourceInfo = definition?.toSourceInfo(),
        )
        resolving.remove(name)
        cache[name] = result
        return result
      }

      return resolve(variableName)
    }

    private fun resolveStartupVariableDocumentation(
      project: Project,
      hostFile: VirtualFile,
      hostText: String,
      untilOffset: Int,
      variableName: String,
    ): VariableHoverResolution? {
      val ownerRoot = EpicsPathResolver.findOwningEpicsRoot(project, hostFile)
      val releaseState = loadReleaseVariableState(ownerRoot)
      val envPathsState = loadEnvPathsVariableState(hostFile.parent?.toNioPath(), releaseState.resolvedValues)
      val state = StartupVariableHoverState(
        currentDirectory = hostFile.parent?.toNioPath() ?: ownerRoot,
        variables = linkedMapOf<String, String>().apply {
          putAll(releaseState.resolvedValues)
          putAll(envPathsState.resolvedValues)
        },
        sources = linkedMapOf<String, VariableHoverSourceInfo>().apply {
          releaseState.definitions.forEach { (name, definition) ->
            definition.toSourceInfo()?.let { put(name, it) }
          }
          envPathsState.definitions.forEach { (name, definition) ->
            definition.toSourceInfo()?.let { put(name, it) }
          }
        },
      )
      if (isStartupFile(hostFile)) {
        applyStartupVariableStateUntilOffset(hostText, untilOffset, state, hostFile)
      }

      val resolvedValue = state.variables[variableName] ?: System.getenv(variableName) ?: return null
      val sourceInfo = state.sources[variableName]
      val absolutePath = computeAbsoluteVariablePath(
        resolvedValue,
        sourceInfo?.baseDirectory ?: state.currentDirectory,
      )
      return VariableHoverResolution(
        resolvedValue = absolutePath?.pathString ?: resolvedValue,
        absolutePath = absolutePath,
        sourceInfo = sourceInfo,
      )
    }

    private fun buildVariableDocumentation(
      variableName: String,
      resolution: VariableHoverResolution,
    ): String {
      val terminalDirectory = resolution.absolutePath?.takeIf { it.exists() && it.isDirectory() }
      return buildString {
        append("<html><body>")
        append("<h3>").append(escape(variableName)).append("</h3>")
        append(paragraph("Resolved value", resolution.resolvedValue))

        if (resolution.absolutePath != null && resolution.absolutePath.pathString != resolution.resolvedValue) {
          append(paragraph("Absolute path", resolution.absolutePath.pathString))
        }

        resolution.sourceInfo?.sourcePath?.let { sourcePath ->
          val line = resolution.sourceInfo.line
          val offset = line?.let { lineNumberToOffset(sourcePath, it) }
          val label = if (line != null) "$sourcePath:$line" else sourcePath
          append(linkedPathParagraph("Defined in", label, sourcePath, offset))
        }
        resolution.sourceInfo?.sourceKind?.let { append(paragraph("Source", it)) }
        resolution.sourceInfo?.rawValue?.let { append(paragraph("Assigned value", it)) }
        terminalDirectory?.let {
          append("<p><a href=\"")
          append(escapeAttribute(buildOpenDirectoryInTerminalHref(it)))
          append("\">")
          append(escape("Open in terminal"))
          append("</a></p>")
        }
        append("</body></html>")
      }
    }

    private fun buildDtypReferenceKey(
      hostFile: VirtualFile,
      context: DatabaseFieldValueContext,
      matches: List<DbdDeviceHoverMatch>,
    ): String {
      return matches.joinToString(
        separator = "|",
        prefix = "${hostFile.path}:dtyp:${context.recordName}:${context.recordType}:${context.value}:",
      ) { match ->
        "${match.absolutePath.pathString}:${match.startOffset}:${match.supportName}"
      }
    }

    private fun buildDtypDeviceDocumentation(
      context: DatabaseFieldValueContext,
      matches: List<DbdDeviceHoverMatch>,
    ): String {
      return buildString {
        append("<html><body>")
        append("<h3>").append(escape("EPICS DTYP device support")).append("</h3>")
        append(paragraph("Record", context.recordName))
        append(paragraph("Type", context.recordType))
        append(paragraph("Field", context.fieldName))
        append(paragraph("Current value", context.value))

        matches.take(5).forEachIndexed { index, match ->
          append("<hr/>")
          append("<h4>").append(escape("${index + 1}. ${match.supportName}")).append("</h4>")
          append(pathParagraph("DBD", match.absolutePath.pathString, match.startOffset))
          append(paragraph("Line", match.line.toString()))
          append(paragraph("Link type", match.linkType))
          append(paragraph("Search root", match.searchLabel))
          appendPreview("device(...)", match.declarationText)
        }

        if (matches.size > 5) {
          append(paragraph("Omitted", "${matches.size - 5} more matching device declarations"))
        }
        append("</body></html>")
      }
    }

    private fun buildMenuFieldDocumentation(
      hostFile: VirtualFile,
      context: EpicsRecordCompletionSupport.MenuFieldValueContext,
    ): String {
      return buildString {
        append("<html><body>")
        append("<h3>").append(escape("EPICS menu field choices")).append("</h3>")
        append(paragraph("Record", context.recordName))
        append(paragraph("Type", context.recordType))
        append(paragraph("Field", context.fieldName))
        append(paragraph("Current value", context.value))
        append("<p><b>").append(escape("Choices")).append(":</b></p>")
        append("<pre>")
        context.choices.forEachIndexed { index, choice ->
          val renderedChoice = if (choice.isEmpty()) "\"\"" else choice
          val prefix = if (choice == context.value) "* " else "  "
          val href = buildMenuChoiceHref(hostFile, context.valueStart, context.valueEnd, choice)
          append(escape(prefix))
          append("<a href=\"").append(escapeAttribute(href)).append("\">")
          append(escape("[$index] $renderedChoice"))
          append("</a>")
          append("\n")
        }
        append("</pre>")
        append("</body></html>")
      }
    }

    private fun buildMenuChoiceHref(
      hostFile: VirtualFile,
      valueStart: Int,
      valueEnd: Int,
      choice: String,
    ): String {
      val encodedFile = urlEncode(File(hostFile.path).toURI().toASCIIString())
      val encodedChoice = urlEncode(choice)
      return "epics-menu://replace?file=$encodedFile&start=$valueStart&end=$valueEnd&value=$encodedChoice"
    }

    private fun buildStreamProtocolCommandReferenceKey(
      hostFile: VirtualFile,
      context: StreamProtocolCommandReference,
      matches: List<StreamProtocolCommandHoverMatch>,
    ): String {
      return matches.joinToString(
        separator = "|",
        prefix = "${hostFile.path}:proto-command:${context.recordName}:${context.fieldName}:${context.protocolPath}:${context.commandName}:",
      ) { match ->
        "${match.absolutePath.pathString}:${match.startOffset}"
      }
    }

    private fun buildStreamProtocolCommandDocumentation(
      context: StreamProtocolCommandReference,
      matches: List<StreamProtocolCommandHoverMatch>,
    ): String {
      if (matches.size == 1) {
        val match = matches.first()
        return buildString {
          append("<html><body>")
          append("<h3>").append(escape("EPICS StreamDevice protocol command")).append("</h3>")
          append(paragraph("Record", context.recordName))
          append(paragraph("Type", context.recordType))
          append(paragraph("Field", context.fieldName))
          append(paragraph("Protocol", context.protocolPath))
          append(paragraph("Command", context.commandName))
          append(pathParagraph("Path", match.absolutePath.pathString, match.startOffset))
          append(paragraph("Line", match.line.toString()))
          appendProtocolHighlightedPreview("Command preview", match.definitionText)
          append("</body></html>")
        }
      }

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape("EPICS StreamDevice protocol command")).append("</h3>")
        append(paragraph("Record", context.recordName))
        append(paragraph("Type", context.recordType))
        append(paragraph("Field", context.fieldName))
        append(paragraph("Protocol", context.protocolPath))
        append(paragraph("Command", context.commandName))
        append(paragraph("Matches", matches.size.toString()))

        matches.forEachIndexed { index, match ->
          append("<hr/>")
          append("<h4>").append(escape("${index + 1}. ${match.absolutePath.name}")).append("</h4>")
          append(pathParagraph("Path", match.absolutePath.pathString, match.startOffset))
          append(paragraph("Line", match.line.toString()))
          appendProtocolHighlightedPreview("Command preview", match.definitionText)
        }

        append("</body></html>")
      }
    }

    private fun buildDocumentation(reference: EpicsResolvedReference, hostFileName: String): String {
      val text = readText(reference.targetFile)
      val title = when (reference.kind) {
        EpicsPathKind.DATABASE -> {
          if (reference.targetFile.extension?.lowercase() in DATABASE_EXTENSIONS) {
            "EPICS database/template file"
          } else {
            "EPICS database file"
          }
        }
        EpicsPathKind.SUBSTITUTIONS -> "EPICS substitutions file"
        EpicsPathKind.PROTOCOL -> "EPICS StreamDevice protocol file"
        EpicsPathKind.DBD -> "EPICS database definition file"
        EpicsPathKind.LIBRARY -> "EPICS library file"
      }
      val readOnlySuffix = startupReadOnlySuffix(hostFileName, reference)

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape(title + readOnlySuffix)).append("</h3>")
        append(pathParagraph("Path", reference.targetFile.path))

        if (hostFileName == "Makefile") {
          val installedName = reference.rawPath.substringAfterLast('/').substringAfterLast('\\')
          if (installedName.isNotBlank() && installedName != reference.targetFile.name) {
            append(paragraph("Installed name", installedName))
          }
        }

        when (reference.kind) {
          EpicsPathKind.DATABASE -> appendDatabaseSummary(reference.targetFile, text)
          EpicsPathKind.SUBSTITUTIONS -> appendSubstitutionSummary(reference.targetFile, text)
          EpicsPathKind.PROTOCOL -> if (text != null) appendPreview("Content preview", previewLines(text, 200))
          EpicsPathKind.DBD -> if (text != null) appendPreview("Content preview", previewLines(text, 120))
          EpicsPathKind.LIBRARY -> {
            val parentPath = reference.targetFile.parent?.path
            if (!parentPath.isNullOrBlank()) {
              append(paragraph("Library directory", parentPath))
            }
          }
        }

        append("</body></html>")
      }
    }

    private fun buildDocumentation(references: List<EpicsResolvedReference>, hostFileName: String): String {
      if (references.size == 1) {
        return buildDocumentation(references.first(), hostFileName)
      }

      val firstKind = references.firstOrNull()?.kind
      val title = when (firstKind) {
        EpicsPathKind.DATABASE -> "EPICS database/template file candidates"
        EpicsPathKind.SUBSTITUTIONS -> "EPICS substitutions file candidates"
        EpicsPathKind.PROTOCOL -> "EPICS StreamDevice protocol file candidates"
        EpicsPathKind.DBD -> "EPICS database definition file candidates"
        EpicsPathKind.LIBRARY -> "EPICS library file candidates"
        null -> "EPICS file candidates"
      }

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape(title)).append("</h3>")
        append(paragraph("Matches", references.size.toString()))

        references.forEachIndexed { index, reference ->
          val text = readText(reference.targetFile)
          append("<hr/>")
          append("<h4>")
            .append(escape("${index + 1}. ${reference.targetFile.name}${startupReadOnlySuffix(hostFileName, reference)}"))
            .append("</h4>")
          append(pathParagraph("Path", reference.targetFile.path))

          when (reference.kind) {
            EpicsPathKind.DATABASE -> appendDatabaseSummary(reference.targetFile, text)
            EpicsPathKind.SUBSTITUTIONS -> appendSubstitutionSummary(reference.targetFile, text)
            EpicsPathKind.PROTOCOL -> if (text != null) appendPreview("Content preview", previewLines(text, 200))
            EpicsPathKind.DBD -> if (text != null) appendPreview("Content preview", previewLines(text, 120))
            EpicsPathKind.LIBRARY -> {
              val parentPath = reference.targetFile.parent?.path
              if (!parentPath.isNullOrBlank()) {
                append(paragraph("Library directory", parentPath))
              }
            }
          }
        }

        append("</body></html>")
      }
    }

    private fun startupReadOnlySuffix(hostFileName: String, reference: EpicsResolvedReference): String {
      if (!isStartupHostFileName(hostFileName)) {
        return ""
      }
      if (reference.kind != EpicsPathKind.DATABASE && reference.kind != EpicsPathKind.SUBSTITUTIONS) {
        return ""
      }
      return if (reference.targetFile.isWritable) "" else " (read only)"
    }

    private fun isStartupHostFileName(hostFileName: String): Boolean {
      val normalizedName = hostFileName.lowercase()
      return normalizedName.endsWith(".cmd") || normalizedName.endsWith(".iocsh") || normalizedName == "st.cmd"
    }

    private fun buildReferenceKey(
      hostFile: VirtualFile,
      references: List<EpicsResolvedReference>,
    ): String {
      return references.joinToString(
        separator = "|",
        prefix = "${hostFile.path}:reference:",
      ) { reference ->
        "${reference.kind}:${reference.rawPath}:${reference.targetFile.path}"
      }
    }

    private fun buildRecordReferenceKey(
      hostFile: VirtualFile,
      definitions: List<EpicsResolvedRecordDefinition>,
    ): String {
      return definitions.joinToString(
        separator = "|",
        prefix = "${hostFile.path}:record:",
      ) { definition ->
        "${definition.targetFile.path}:${definition.recordName}:${definition.line}"
      }
    }

    private fun buildRecordDocumentation(
      definitions: List<EpicsResolvedRecordDefinition>,
      hostFile: VirtualFile,
    ): String {
      if (definitions.size == 1) {
        return buildRecordDocumentation(definitions.first(), hostFile)
      }

      val referenceLabel = when {
        isStartupFile(hostFile) -> "dbpf"
        else -> "record link"
      }

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape("EPICS record definitions")).append("</h3>")
        append(paragraph("Matches", definitions.size.toString()))
        append(paragraph("Referenced by", referenceLabel))

        definitions.forEachIndexed { index, definition ->
          append("<hr/>")
          append("<h4>")
            .append(escape("${index + 1}. ${definition.recordName}"))
            .append("</h4>")
          append(paragraph("Type", definition.recordType))
          append(pathParagraph("Path", definition.targetFile.path, definition.recordStartOffset))
          append(paragraph("Line", definition.line.toString()))
          appendHighlightedPreview("Record preview", buildRecordPreview(readText(definition.targetFile), definition))
        }

        append("</body></html>")
      }
    }

    private fun buildRecordDocumentation(
      definition: EpicsResolvedRecordDefinition,
      hostFile: VirtualFile,
    ): String {
      val text = readText(definition.targetFile)
      val preview = buildRecordPreview(text, definition)
      val referenceLabel = when {
        isStartupFile(hostFile) -> "dbpf"
        else -> "record link"
      }

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape("EPICS record definition")).append("</h3>")
        append(paragraph("Record", definition.recordName))
        append(paragraph("Type", definition.recordType))
        append(paragraph("Referenced by", referenceLabel))
        append(pathParagraph("Path", definition.targetFile.path, definition.recordStartOffset))
        append(paragraph("Line", definition.line.toString()))
        appendHighlightedPreview("Record preview", preview)
        append("</body></html>")
      }
    }

    private fun StringBuilder.appendDatabaseSummary(targetFile: VirtualFile, text: String?) {
      if (text == null) {
        return
      }
      val filteredText = text.lineSequence()
        .filterNot { it.trimStart().startsWith("#") }
        .joinToString("\n")
      val recordDeclarations = RECORD_DECLARATION_REGEX.findAll(filteredText)
        .map { it.groupValues[1] }
        .toList()
      val macroNames = extractMacroNames(filteredText)

      append(paragraph("Records", recordDeclarations.size.toString()))
      append(paragraph("Macros", if (macroNames.isEmpty()) "none" else macroNames.joinToString(", ")))
      if (recordDeclarations.isNotEmpty()) {
        val previewNames = recordDeclarations.take(100)
        appendCompactCodeLines(previewNames)
        if (recordDeclarations.size > previewNames.size) {
          append(paragraph("Omitted", "${recordDeclarations.size - previewNames.size} more record names"))
        }
      } else {
        append(paragraph("File", targetFile.name))
      }
    }

    private fun StringBuilder.appendSubstitutionSummary(targetFile: VirtualFile, text: String?) {
      if (text == null) {
        return
      }
      val blocks = parseSubstitutionBlocks(text)
      val expansionCount = blocks.sumOf { countSubstitutionExpansions(it.body) }
      append(paragraph("Blocks", blocks.size.toString()))
      append(paragraph("Expansions", expansionCount.toString()))
      appendPreview("Content preview", previewLines(text, 200))
      val omitted = text.lineSequence().count() - text.lineSequence().take(200).count()
      if (omitted > 0) {
        append(paragraph("Omitted", "$omitted more lines"))
      }
      if (blocks.isEmpty()) {
        append(paragraph("File", targetFile.name))
      }
    }

    private fun paragraph(label: String, value: String): String {
      return "<p><b>${escape(label)}:</b> <code>${escape(value)}</code></p>"
    }

    private fun pathParagraph(label: String, value: String, offset: Int? = null): String {
      val baseHref = File(value).toURI().toASCIIString()
      val href = if (offset != null && offset >= 0) {
        "$baseHref#offset=$offset"
      } else {
        baseHref
      }
      return "<p><b>${escape(label)}:</b> <a href=\"$href\"><code>${escape(value)}</code></a></p>"
    }

    private fun linkedPathParagraph(
      label: String,
      displayValue: String,
      targetPath: String,
      offset: Int? = null,
    ): String {
      val baseHref = File(targetPath).toURI().toASCIIString()
      val href = if (offset != null && offset >= 0) {
        "$baseHref#offset=$offset"
      } else {
        baseHref
      }
      return "<p><b>${escape(label)}:</b> <a href=\"$href\"><code>${escape(displayValue)}</code></a></p>"
    }

    private fun StringBuilder.appendPreview(label: String, content: String) {
      if (content.isBlank()) {
        return
      }
      append("<p><b>").append(escape(label)).append(":</b></p>")
      append("<pre style=\"margin: 4px 0 0; line-height: 1.15;\">")
      append(escape(content))
      append("</pre>")
    }

    private fun StringBuilder.appendHighlightedPreview(label: String, content: String) {
      if (content.isBlank()) {
        return
      }
      append("<p><b>").append(escape(label)).append(":</b></p>")
      append("<pre style=\"margin: 4px 0 0; line-height: 1.15;\">")
      append(renderDatabasePreviewHtml(content))
      append("</pre>")
    }

    private fun StringBuilder.appendProtocolHighlightedPreview(label: String, content: String) {
      if (content.isBlank()) {
        return
      }
      append("<p><b>").append(escape(label)).append(":</b></p>")
      append("<pre style=\"margin: 4px 0 0; line-height: 1.15;\">")
      append(renderProtocolPreviewHtml(content))
      append("</pre>")
    }

    private fun StringBuilder.appendCompactCodeLines(lines: List<String>) {
      if (lines.isEmpty()) {
        return
      }
      append("<div style=\"margin: 4px 0 0; line-height: 1.15;\">")
      lines.forEachIndexed { index, line ->
        if (index > 0) {
          append("<br/>")
        }
        append("<code>").append(escape(line)).append("</code>")
      }
      append("</div>")
    }

    private fun previewLines(text: String, lineLimit: Int): String {
      return text.lineSequence().take(lineLimit).joinToString("\n")
    }

    private fun buildRecordPreview(
      text: String?,
      definition: EpicsResolvedRecordDefinition,
    ): String {
      if (text.isNullOrEmpty()) {
        return """record(${definition.recordType}, "${definition.recordName}")"""
      }

      val rawPreview = text
        .substring(
          definition.recordStartOffset.coerceAtLeast(0).coerceAtMost(text.length),
          definition.recordEndOffset.coerceAtLeast(0).coerceAtMost(text.length),
        )
        .trim()

      if (rawPreview.isBlank()) {
        return """record(${definition.recordType}, "${definition.recordName}")"""
      }

      val normalizedPreview = RECORD_DECLARATION_PREFIX_REGEX.find(rawPreview)?.let { match ->
        val replacement = match.groups[1]?.value.orEmpty() +
          escapeDoubleQuotedString(definition.recordName) +
          match.groups[3]?.value.orEmpty()
        rawPreview.replaceRange(match.range, replacement)
      } ?: rawPreview

      val previewLines = normalizedPreview.lineSequence().toList()
      val truncatedLines = if (previewLines.size > RECORD_PREVIEW_MAX_LINES) {
        previewLines.take(RECORD_PREVIEW_MAX_LINES) + "..."
      } else {
        previewLines
      }
      val truncatedPreview = truncatedLines.joinToString("\n")
      return if (truncatedPreview.length > RECORD_PREVIEW_MAX_CHARACTERS) {
        truncatedPreview.take(RECORD_PREVIEW_MAX_CHARACTERS - 3) + "..."
      } else {
        truncatedPreview
      }
    }

    private fun renderDatabasePreviewHtml(content: String): String {
      return renderHighlightedPreviewHtml(
        content,
        EpicsLexingProfile(DOCUMENTATION_DATABASE_KEYWORDS),
      )
    }

    private fun renderProtocolPreviewHtml(content: String): String {
      return renderHighlightedPreviewHtml(
        content,
        EpicsLexingProfile(
          keywords = PROTOCOL_KEYWORDS,
          allowSingleQuotedStrings = true,
          extraIdentifierChars = setOf('-'),
          caseInsensitiveKeywords = true,
        ),
      )
    }

    private fun renderHighlightedPreviewHtml(
      content: String,
      profile: EpicsLexingProfile,
    ): String {
      val lexer = EpicsSimpleLexer(profile)
      lexer.start(content, 0, content.length, 0)
      val html = StringBuilder()
      while (true) {
        val tokenType = lexer.tokenType ?: break
        val tokenText = content.substring(lexer.tokenStart, lexer.tokenEnd)
        val escapedText = escape(tokenText)
        val color = resolveTokenColor(tokenType)
        val bold = tokenType == EpicsTokenTypes.KEYWORD

        if (color == null) {
          html.append(escapedText)
        } else {
          if (bold) {
            html.append("<b>")
          }
          html.append("<font color=\"").append(color).append("\">")
          html.append(escapedText)
          html.append("</font>")
          if (bold) {
            html.append("</b>")
          }
        }
        lexer.advance()
      }
      return html.toString()
    }

    private fun resolveTokenColor(tokenType: IElementType): String? {
      val key = when (tokenType) {
        EpicsTokenTypes.COMMENT -> EpicsHighlightingKeys.COMMENT
        EpicsTokenTypes.STRING -> EpicsHighlightingKeys.STRING
        EpicsTokenTypes.KEYWORD -> EpicsHighlightingKeys.KEYWORD
        EpicsTokenTypes.NUMBER -> EpicsHighlightingKeys.NUMBER
        EpicsTokenTypes.MACRO -> EpicsHighlightingKeys.MACRO
        EpicsTokenTypes.BRACE -> EpicsHighlightingKeys.BRACE
        EpicsTokenTypes.PAREN -> EpicsHighlightingKeys.PAREN
        EpicsTokenTypes.BRACKET -> EpicsHighlightingKeys.BRACKET
        EpicsTokenTypes.COMMA -> EpicsHighlightingKeys.COMMA
        EpicsTokenTypes.OPERATOR -> EpicsHighlightingKeys.OPERATOR
        else -> null
      } ?: return null

      val scheme = EditorColorsManager.getInstance().globalScheme
      val color = scheme.getAttributes(key)?.foregroundColor
        ?: key.defaultAttributes.foregroundColor
        ?: DEFAULT_TOKEN_COLORS[key]
      return color?.let(::toHtmlColor)
    }

    private fun extractMacroNames(text: String): List<String> {
      val names = linkedSetOf<String>()
      EPICS_VARIABLE_REGEX.findAll(text).forEach { match ->
        val name = match.groups[1]?.value
          ?: match.groups[3]?.value
          ?: match.groups[5]?.value
          ?: return@forEach
        if (name.isNotBlank()) {
          names += name
        }
      }
      return names.toList().sorted()
    }

    private fun parseSubstitutionBlocks(text: String): List<SubstitutionBlock> {
      val blocks = mutableListOf<SubstitutionBlock>()
      var searchIndex = 0
      while (searchIndex < text.length) {
        val match = SUBSTITUTION_BLOCK_START_REGEX.find(text, searchIndex) ?: break
        val headerEnd = match.range.last + 1
        var depth = 1
        var index = headerEnd
        while (index < text.length && depth > 0) {
          when (text[index]) {
            '{' -> depth += 1
            '}' -> depth -= 1
          }
          index += 1
        }
        if (depth == 0) {
          blocks += SubstitutionBlock(match.groups[1]?.value.orEmpty(), text.substring(headerEnd, index - 1))
          searchIndex = index
        } else {
          break
        }
      }
      return blocks
    }

    private fun countSubstitutionExpansions(body: String): Int {
      var count = 0
      body.lineSequence().forEach { line ->
        val trimmed = line.trim()
        if (trimmed.isEmpty() || trimmed.startsWith("#") || trimmed.startsWith("pattern")) {
          return@forEach
        }
        if (trimmed.startsWith("{")) {
          count += 1
        }
      }
      return count
    }

    private fun findVariableReferenceAtOffset(text: String, offset: Int): VariableReference? {
      return EPICS_VARIABLE_REGEX.findAll(text).firstNotNullOfOrNull { match ->
        val start = match.range.first
        val end = match.range.last + 1
        if (offset !in start until end) {
          return@firstNotNullOfOrNull null
        }
        val variableName = match.groups[1]?.value
          ?: match.groups[3]?.value
          ?: match.groups[5]?.value
          ?: return@firstNotNullOfOrNull null
        VariableReference(variableName, start, end)
      }
    }

    private fun loadReleaseVariableState(ownerRoot: Path): VariableDefinitionState {
      val configureDirectory = ownerRoot.resolve("configure")
      val definitions = linkedMapOf<String, VariableDefinition>()
      definitions["TOP"] = VariableDefinition(ownerRoot.pathString, null, null, null, ownerRoot)
      collectReleaseFiles(configureDirectory).forEach { releaseFile ->
        val text = runCatching { Files.readString(releaseFile) }.getOrNull() ?: return@forEach
        applyMakefileVariableDefinitions(
          text = text,
          sourcePath = releaseFile.pathString,
          sourceKind = "RELEASE",
          baseDirectory = releaseFile.parent ?: configureDirectory,
          target = definitions,
        )
      }
      return resolveVariableDefinitionState(definitions, ownerRoot)
    }

    private fun loadEnvPathsVariableState(
      startupDirectory: Path?,
      baseVariables: Map<String, String>,
    ): VariableDefinitionState {
      val definitions = linkedMapOf<String, VariableDefinition>()
      val resolvedValues = linkedMapOf<String, String>()
      val files = startupDirectory
        ?.toFile()
        ?.listFiles { file -> file.isFile && (file.name == "envPaths" || file.name.startsWith("envPaths.")) }
        ?.sortedBy { if (it.name == "envPaths") 0 else 1 }
        .orEmpty()
      for (envPathsFile in files) {
        val text = runCatching { Files.readString(envPathsFile.toPath()) }.getOrNull() ?: continue
        text.split(Regex("""\r?\n""")).forEachIndexed { index, rawLine ->
          val match = STARTUP_ENV_SET_REGEX.find(maskHashCommentLine(rawLine)) ?: return@forEachIndexed
          val name = match.groups[1]?.value?.trim().orEmpty()
          val rawValue = match.groups[2]?.value.orEmpty()
          if (name.isEmpty()) {
            return@forEachIndexed
          }
          definitions[name] = VariableDefinition(
            rawValue = rawValue,
            sourcePath = envPathsFile.path,
            line = index + 1,
            sourceKind = "envPaths",
            baseDirectory = envPathsFile.parentFile.toPath(),
          )
          resolvedValues[name] = expandDocumentationVariableValue(rawValue) { variableName ->
            resolvedValues[variableName] ?: baseVariables[variableName] ?: System.getenv(variableName)
          }
        }
      }
      return VariableDefinitionState(resolvedValues, definitions)
    }

    private fun resolveVariableDefinitionState(
      definitions: Map<String, VariableDefinition>,
      ownerRoot: Path,
    ): VariableDefinitionState {
      val cache = linkedMapOf<String, String>()
      val resolving = linkedSetOf<String>()

      fun resolve(name: String): String? {
        if (cache.containsKey(name)) {
          return cache[name]
        }
        if (!resolving.add(name)) {
          return null
        }

        val definition = definitions[name]
        val rawValue = definition?.rawValue ?: System.getenv(name)
        if (rawValue == null) {
          resolving.remove(name)
          return null
        }

        val expanded = expandDocumentationVariableValue(rawValue) { nestedName ->
          resolve(nestedName) ?: System.getenv(nestedName)
        }
        resolving.remove(name)
        cache[name] = expanded
        return expanded
      }

      definitions.keys.forEach { resolve(it) }
      cache.putIfAbsent("TOP", ownerRoot.pathString)
      return VariableDefinitionState(cache, LinkedHashMap(definitions))
    }

    private fun collectReleaseFiles(configureDirectory: Path): List<Path> {
      if (!configureDirectory.exists() || !configureDirectory.isDirectory()) {
        return emptyList()
      }

      val releaseFiles = mutableListOf<Path>()
      Files.list(configureDirectory).use { stream ->
        stream
          .filter { file ->
            Files.isRegularFile(file) && (
              file.fileName.toString() == "RELEASE" ||
                file.fileName.toString().startsWith("RELEASE.")
            )
          }
          .sorted { left, right ->
            releaseFileSortKey(left.fileName.toString()).compareTo(releaseFileSortKey(right.fileName.toString()))
          }
          .forEach { releaseFiles.add(it) }
      }
      return releaseFiles
    }

    private fun releaseFileSortKey(fileName: String): String {
      return when (fileName) {
        "RELEASE" -> "0:$fileName"
        "RELEASE.local" -> "1:$fileName"
        else -> "2:$fileName"
      }
    }

    private fun applyMakefileVariableDefinitions(
      text: String,
      sourcePath: String,
      sourceKind: String,
      baseDirectory: Path,
      target: LinkedHashMap<String, VariableDefinition>,
    ) {
      text.split(Regex("""\r?\n""")).forEachIndexed { index, rawLine ->
        val match = MAKEFILE_ASSIGNMENT_REGEX.find(rawLine) ?: return@forEachIndexed
        val variableName = match.groups[1]?.value.orEmpty()
        val operator = match.groups[2]?.value.orEmpty()
        val rawValue = match.groups[3]?.value?.trim().orEmpty()
        if (variableName.isBlank()) {
          return@forEachIndexed
        }

        val definition = when {
          operator == "?=" && target.containsKey(variableName) -> null
          operator == "+=" && target.containsKey(variableName) ->
            target[variableName]?.copy(
              rawValue = "${target[variableName]?.rawValue.orEmpty()} $rawValue".trim(),
              sourcePath = sourcePath,
              line = index + 1,
              sourceKind = sourceKind,
              baseDirectory = baseDirectory,
            )

          else -> VariableDefinition(rawValue, sourcePath, index + 1, sourceKind, baseDirectory)
        }
        if (definition != null) {
          target[variableName] = definition
        }
      }
    }

    private fun applyStartupVariableStateUntilOffset(
      text: String,
      untilOffset: Int,
      state: StartupVariableHoverState,
      hostFile: VirtualFile,
    ) {
      var runningOffset = 0
      for ((lineIndex, line) in text.split('\n').withIndex()) {
        val lineStart = runningOffset
        if (lineStart >= untilOffset) {
          break
        }
        applyStartupVariableLine(line, lineIndex + 1, state, hostFile)
        runningOffset = lineStart + line.length + 1
      }
    }

    private fun applyStartupVariableLine(
      line: String,
      lineNumber: Int,
      state: StartupVariableHoverState,
      hostFile: VirtualFile,
    ) {
      val sanitizedLine = maskHashCommentLine(line)
      STARTUP_ENV_SET_REGEX.find(sanitizedLine)?.let { match ->
        val name = match.groups[1]?.value?.trim().orEmpty()
        val rawValue = match.groups[2]?.value.orEmpty()
        if (name.isNotEmpty()) {
          state.variables[name] = expandDocumentationVariableValue(rawValue) { variableName ->
            state.variables[variableName] ?: System.getenv(variableName)
          }
          state.sources[name] = VariableHoverSourceInfo(
            sourcePath = hostFile.path,
            line = lineNumber,
            sourceKind = if (isEnvPathsFile(hostFile)) "envPaths" else "startup",
            rawValue = rawValue,
            baseDirectory = state.currentDirectory,
          )
        }
      }

      STARTUP_CD_REGEX.find(sanitizedLine)?.let { match ->
        val rawDirectory = match.groups[1]?.value ?: match.groups[2]?.value ?: ""
        val expandedDirectory = expandDocumentationVariableValue(rawDirectory) { variableName ->
          state.variables[variableName] ?: System.getenv(variableName)
        }
        val resolvedDirectory = resolveAbsoluteOrRelative(state.currentDirectory, expandedDirectory)
        if (resolvedDirectory.exists() && resolvedDirectory.isDirectory()) {
          state.currentDirectory = resolvedDirectory.normalize()
        }
      }
    }

    private fun expandDocumentationVariableValue(
      rawValue: String,
      resolver: (String) -> String?,
    ): String {
      var expanded = rawValue
      repeat(5) {
        val next = EPICS_VARIABLE_REGEX.replace(expanded) { match ->
          val variableName = match.groups[1]?.value
            ?: match.groups[3]?.value
            ?: match.groups[5]?.value
            ?: return@replace match.value
          val defaultValue = match.groups[2]?.value ?: match.groups[4]?.value
          resolver(variableName) ?: defaultValue ?: match.value
        }
        if (next == expanded) {
          return expanded
        }
        expanded = next
      }
      return expanded
    }

    private fun computeAbsoluteVariablePath(value: String, baseDirectory: Path?): Path? {
      if (value.isBlank() || baseDirectory == null || !looksLikePathValue(value) || EPICS_VARIABLE_REGEX.containsMatchIn(value)) {
        return null
      }

      val candidate = runCatching {
        when {
          value.startsWith("~/") -> Path.of(System.getProperty("user.home")).resolve(value.removePrefix("~/"))
          Path.of(value).isAbsolute -> Path.of(value)
          else -> baseDirectory.resolve(value)
        }
      }.getOrNull() ?: return null
      return candidate.normalize()
    }

    private fun looksLikePathValue(value: String): Boolean {
      return value.startsWith(".") ||
        value.startsWith("/") ||
        value.startsWith("~/") ||
        value.contains("/") ||
        value.contains("\\")
    }

    private fun resolveAbsoluteOrRelative(currentDirectory: Path, value: String): Path {
      val candidate = runCatching { Path.of(value) }.getOrNull()
      return if (candidate != null && candidate.isAbsolute) {
        candidate
      } else {
        currentDirectory.resolve(value)
      }
    }

    private fun lineNumberToOffset(sourcePath: String, lineNumber: Int): Int? {
      if (lineNumber <= 1) {
        return 0
      }
      val text = runCatching { Files.readString(Path.of(sourcePath)) }.getOrNull() ?: return null
      var line = 1
      for ((index, character) in text.withIndex()) {
        if (line >= lineNumber) {
          return index
        }
        if (character == '\n') {
          line += 1
        }
      }
      return if (line == lineNumber) text.length else null
    }

    private fun buildOpenDirectoryInTerminalHref(directory: Path): String {
      return "epics-terminal://open?path=${urlEncode(directory.pathString)}"
    }

    private fun maskHashCommentLine(line: String): String {
      val sanitized = StringBuilder(line.length)
      var inString = false
      var escaped = false

      for (character in line) {
        if (escaped) {
          sanitized.append(character)
          escaped = false
          continue
        }

        if (character == '\\') {
          sanitized.append(character)
          escaped = true
          continue
        }

        if (character == '"') {
          inString = !inString
          sanitized.append(character)
          continue
        }

        if (!inString && character == '#') {
          sanitized.append(' ')
          repeat(line.length - sanitized.length) {
            sanitized.append(' ')
          }
          break
        }

        sanitized.append(character)
      }

      return sanitized.toString()
    }

    private fun isStartupStateFile(file: VirtualFile): Boolean {
      return isStartupFile(file) || isEnvPathsFile(file)
    }

    private fun isEnvPathsFile(file: VirtualFile): Boolean {
      return file.name == "envPaths" || file.name.startsWith("envPaths.")
    }

    private fun isEpicsReleaseFile(file: VirtualFile): Boolean {
      return file.parent?.name == "configure" &&
        (file.name == "RELEASE" || file.name.startsWith("RELEASE."))
    }

    private fun findDatabaseFieldValueContext(text: String, offset: Int): DatabaseFieldValueContext? {
      val recordDeclaration = EpicsRecordCompletionSupport.extractRecordDeclarations(text)
        .firstOrNull { declaration -> offset in declaration.recordStart..declaration.recordEnd }
        ?: return null
      val fieldDeclaration = EpicsRecordCompletionSupport.extractFieldDeclarationsInRecord(text, recordDeclaration)
        .firstOrNull { declaration -> offset in declaration.valueStart..declaration.valueEnd }
        ?: return null

      return DatabaseFieldValueContext(
        recordType = recordDeclaration.recordType,
        recordName = recordDeclaration.name,
        fieldName = fieldDeclaration.fieldName,
        value = fieldDeclaration.value,
        valueStart = fieldDeclaration.valueStart,
        valueEnd = fieldDeclaration.valueEnd,
      )
    }

    private fun resolveStreamProtocolCommandHoverMatches(
      project: Project,
      hostFile: VirtualFile,
      context: StreamProtocolCommandReference,
    ): List<StreamProtocolCommandHoverMatch> {
      return EpicsPathResolver.resolveStreamProtocolPaths(project, hostFile, context.protocolPath)
        .mapNotNull { protocolPath ->
          EpicsStreamProtocolSupport.findCommandDefinition(protocolPath, context.commandName)
            ?.let { definition ->
              StreamProtocolCommandHoverMatch(
                absolutePath = protocolPath.normalize(),
                line = definition.line,
                startOffset = definition.startOffset,
                definitionText = definition.definitionText,
              )
            }
        }
    }

    private fun resolveDtypDeviceHoverMatches(
      project: Project,
      hostFile: VirtualFile,
      context: DatabaseFieldValueContext,
    ): List<DbdDeviceHoverMatch> {
      val ownerRoot = EpicsPathResolver.findOwningEpicsRoot(project, hostFile)
      val choiceName = context.value.trim()
      if (choiceName.isEmpty()) {
        return emptyList()
      }

      val matches = mutableListOf<DbdDeviceHoverMatch>()
      val seenFiles = linkedSetOf<String>()
      for (searchDirectory in collectDtypDeviceSearchDirectories(project, hostFile, ownerRoot)) {
        val directory = searchDirectory.directory
        if (!directory.exists() || !directory.isDirectory()) {
          continue
        }

        val paths = runCatching {
          Files.list(directory).use { stream ->
            stream
              .filter { candidate -> Files.isRegularFile(candidate) && candidate.name.lowercase().endsWith(".dbd") }
              .sorted(compareBy<Path> { it.name.lowercase() }.thenBy { it.pathString.lowercase() })
              .toList()
          }
        }.getOrDefault(emptyList())

        for (dbdPath in paths) {
          val normalizedPath = dbdPath.normalize()
          if (!seenFiles.add(normalizedPath.pathString)) {
            continue
          }

          val text = runCatching { Files.readString(normalizedPath) }.getOrNull() ?: continue
          for (declaration in extractDbdDeviceDeclarations(text)) {
            if (
              declaration.recordType != context.recordType ||
                declaration.choiceName != choiceName
            ) {
              continue
            }

            matches += DbdDeviceHoverMatch(
              absolutePath = normalizedPath,
              line = lineNumberAt(text, declaration.startOffset),
              startOffset = declaration.startOffset,
              recordType = declaration.recordType,
              linkType = declaration.linkType,
              supportName = declaration.supportName,
              choiceName = declaration.choiceName,
              declarationText = declaration.declarationText,
              searchLabel = searchDirectory.label,
            )
          }
        }
      }

      return matches
    }

    private fun collectDtypDeviceSearchDirectories(
      project: Project,
      hostFile: VirtualFile,
      ownerRoot: Path,
    ): List<DbdSearchDirectory> {
      val directories = mutableListOf<DbdSearchDirectory>()
      val seen = linkedSetOf<String>()

      fun addDirectory(path: Path?, label: String) {
        val normalizedPath = path?.normalize() ?: return
        if (!seen.add(normalizedPath.pathString)) {
          return
        }
        directories += DbdSearchDirectory(normalizedPath, label)
      }

      addDirectory(hostFile.parent?.toNioPath(), "current folder")
      addDirectory(ownerRoot.resolve("dbd"), "project dbd")

      val releaseVariables = project.getService(EpicsBuildModelService::class.java)
        ?.loadBuildModel(ownerRoot)
        ?.releaseVariables
        ?: loadReleaseVariables(ownerRoot)
      resolveReleaseModuleRoots(ownerRoot, releaseVariables).forEach { root ->
        addDirectory(root.rootPath.resolve("dbd"), "${root.variableName} dbd")
      }

      return directories
    }

    private fun loadReleaseVariables(ownerRoot: Path): Map<String, String> {
      val configureDirectory = ownerRoot.resolve("configure")
      val releaseFiles = listOf(
        configureDirectory.resolve("RELEASE"),
        configureDirectory.resolve("RELEASE.local"),
      )
      val rawValues = linkedMapOf<String, String>()
      for (releaseFile in releaseFiles) {
        if (!releaseFile.exists() || !releaseFile.isRegularFile()) {
          continue
        }

        runCatching { Files.readAllLines(releaseFile) }.getOrDefault(emptyList()).forEach { line ->
          val match = RELEASE_ASSIGNMENT_REGEX.find(line) ?: return@forEach
          rawValues[match.groups[1]?.value.orEmpty()] = match.groups[2]?.value?.trim().orEmpty()
        }
      }
      return rawValues
    }

    private fun resolveReleaseModuleRoots(
      ownerRoot: Path,
      releaseVariables: Map<String, String>,
    ): List<ResolvedReleaseRoot> {
      val resolved = linkedMapOf<String, Path?>()
      val resolving = linkedSetOf<String>()

      fun resolve(variableName: String): Path? {
        if (resolved.containsKey(variableName)) {
          return resolved[variableName]
        }
        if (!resolving.add(variableName)) {
          return null
        }

        val rawValue = releaseVariables[variableName]
        if (rawValue.isNullOrBlank()) {
          resolving.remove(variableName)
          resolved[variableName] = null
          return null
        }

        val expandedValue = EPICS_VARIABLE_REGEX.replace(rawValue) { match ->
          val nestedName = match.groups[1]?.value
            ?: match.groups[3]?.value
            ?: match.groups[5]?.value
            ?: return@replace ""
          val defaultValue = match.groups[2]?.value ?: match.groups[4]?.value
          resolve(nestedName)?.pathString
            ?: releaseVariables[nestedName]
            ?: System.getenv(nestedName)
            ?: defaultValue
            ?: ""
        }

        val resolvedPath = when {
          expandedValue.isBlank() ||
            expandedValue.contains("\$(") ||
            expandedValue.contains("\${") -> null
          else -> runCatching {
            val candidate = Path.of(expandedValue)
            if (candidate.isAbsolute) candidate.normalize() else ownerRoot.resolve(candidate).normalize()
          }.getOrNull()
        }?.takeIf { candidate -> candidate.exists() && candidate.isDirectory() }

        resolving.remove(variableName)
        resolved[variableName] = resolvedPath
        return resolvedPath
      }

      return releaseVariables.keys.mapNotNull { variableName ->
        resolve(variableName)?.let { rootPath ->
          ResolvedReleaseRoot(variableName, rootPath)
        }
      }
    }

    private fun extractDbdDeviceDeclarations(text: String): List<DbdDeviceDeclaration> {
      return DBD_DEVICE_DECLARATION_REGEX.findAll(text).mapNotNull { match ->
        val recordType = match.groups[1]?.value.orEmpty()
        val linkType = match.groups[2]?.value.orEmpty()
        val supportName = match.groups[3]?.value.orEmpty()
        val choiceName = match.groups[4]?.value.orEmpty()
        if (
          recordType.isBlank() ||
            linkType.isBlank() ||
            supportName.isBlank() ||
            choiceName.isBlank()
        ) {
          return@mapNotNull null
        }

        DbdDeviceDeclaration(
          recordType = recordType,
          linkType = linkType,
          supportName = supportName,
          choiceName = choiceName,
          declarationText = match.value,
          startOffset = match.range.first,
        )
      }.toList()
    }

    private fun lineNumberAt(text: String, offset: Int): Int {
      return text.take(offset.coerceIn(0, text.length)).count { it == '\n' } + 1
    }

    private fun readText(file: VirtualFile): String? {
      FileDocumentManager.getInstance().getCachedDocument(file)?.let { return it.text }
      return try {
        String(file.contentsToByteArray(), file.charset)
      } catch (_: Exception) {
        null
      }
    }

    private fun escape(value: String): String = StringUtil.escapeXmlEntities(value)

    private fun escapeAttribute(value: String): String = StringUtil.escapeXmlEntities(value)

    private fun urlEncode(value: String): String {
      return URLEncoder.encode(value, StandardCharsets.UTF_8)
    }

    private fun escapeDoubleQuotedString(value: String): String {
      return value.replace("\\", "\\\\").replace("\"", "\\\"")
    }

    private fun toHtmlColor(color: Color): String {
      return "#%02x%02x%02x".format(color.red, color.green, color.blue)
    }

    private fun isStartupFile(file: VirtualFile): Boolean {
      val extension = file.extension?.lowercase()
      return extension == "cmd" || extension == "iocsh" || file.name == "st.cmd"
    }

    private fun isSubstitutionsFile(file: VirtualFile): Boolean {
      return file.extension?.lowercase() in setOf("substitutions", "sub", "subs")
    }

    private fun findPvlistHoverRecordName(text: String, offset: Int): String? {
      val lineStart = text.lastIndexOf('\n', (offset - 1).coerceAtLeast(0)).let { it + 1 }
      val lineEnd = text.indexOf('\n', offset).let { if (it >= 0) it else text.length }
      if (lineStart >= lineEnd) {
        return null
      }

      val lineText = text.substring(lineStart, lineEnd)
      val trimmed = lineText.trim()
      if (
        trimmed.isBlank() ||
        trimmed.startsWith("#") ||
        PVLIST_MACRO_ASSIGNMENT_REGEX.matches(trimmed) ||
        trimmed.contains(' ') ||
        trimmed.contains('\t') ||
        trimmed.contains("=")
      ) {
        return null
      }

      val trimmedStart = lineText.indexOf(trimmed).takeIf { it >= 0 } ?: return null
      val tokenStart = lineStart + trimmedStart
      val tokenEnd = tokenStart + trimmed.length
      if (offset !in tokenStart..tokenEnd) {
        return null
      }

      val expanded = expandPvlistValue(trimmed, extractPvlistMacroDefinitions(text), linkedSetOf())
        ?.trim()
        ?.takeIf { it.isNotBlank() }
        ?: return null
      return stripPvlistProtocolPrefix(expanded)
    }

    private fun extractPvlistMacroDefinitions(text: String): Map<String, String> {
      val macroDefinitions = linkedMapOf<String, String>()
      text.split(Regex("""\r?\n""")).forEach { rawLine ->
        val trimmed = rawLine.trim()
        if (trimmed.isEmpty() || trimmed.startsWith("#")) {
          return@forEach
        }

        val match = PVLIST_MACRO_ASSIGNMENT_REGEX.matchEntire(trimmed) ?: return@forEach
        val macroName = match.groups[1]?.value.orEmpty()
        if (macroName.isNotBlank() && macroName !in macroDefinitions) {
          macroDefinitions[macroName] = match.groups[2]?.value.orEmpty()
        }
      }
      return macroDefinitions
    }

    private fun expandPvlistValue(
      text: String,
      macroDefinitions: Map<String, String>,
      stack: LinkedHashSet<String>,
    ): String? {
      var unresolved = false
      val expanded = PVLIST_MACRO_REFERENCE_REGEX.replace(text) { match ->
        val macroName = match.groups[1]?.value ?: match.groups[3]?.value.orEmpty()
        val defaultValue = match.groups[2]?.value
        val resolved = resolvePvlistMacro(macroName, defaultValue, macroDefinitions, stack)
        if (resolved == null) {
          unresolved = true
          ""
        } else {
          resolved
        }
      }
      return if (unresolved) null else expanded
    }

    private fun resolvePvlistMacro(
      macroName: String,
      defaultValue: String?,
      macroDefinitions: Map<String, String>,
      stack: LinkedHashSet<String>,
    ): String? {
      if (macroName in stack) {
        return null
      }

      val value = macroDefinitions[macroName] ?: return defaultValue
      val nextStack = LinkedHashSet(stack)
      nextStack += macroName
      return expandPvlistValue(value, macroDefinitions, nextStack)
    }

    private fun stripPvlistProtocolPrefix(value: String): String {
      return when {
        value.startsWith("pva://", ignoreCase = true) -> value.drop(6)
        value.startsWith("ca://", ignoreCase = true) -> value.drop(5)
        else -> value
      }
    }

    private data class SubstitutionBlock(
      val templatePath: String,
      val body: String,
    )

    private data class DatabaseFieldValueContext(
      val recordType: String,
      val recordName: String,
      val fieldName: String,
      val value: String,
      val valueStart: Int,
      val valueEnd: Int,
    )

    private data class VariableReference(
      val variableName: String,
      val startOffset: Int,
      val endOffset: Int,
    )

    private data class VariableDefinition(
      val rawValue: String,
      val sourcePath: String?,
      val line: Int?,
      val sourceKind: String?,
      val baseDirectory: Path?,
    ) {
      fun toSourceInfo(): VariableHoverSourceInfo? {
        if (sourcePath == null && sourceKind == null && rawValue.isBlank()) {
          return null
        }
        return VariableHoverSourceInfo(
          sourcePath = sourcePath,
          line = line,
          sourceKind = sourceKind,
          rawValue = rawValue,
          baseDirectory = baseDirectory,
        )
      }
    }

    private data class VariableDefinitionState(
      val resolvedValues: LinkedHashMap<String, String>,
      val definitions: LinkedHashMap<String, VariableDefinition>,
    )

    private data class VariableHoverSourceInfo(
      val sourcePath: String?,
      val line: Int?,
      val sourceKind: String?,
      val rawValue: String?,
      val baseDirectory: Path?,
    )

    private data class VariableHoverResolution(
      val resolvedValue: String,
      val absolutePath: Path?,
      val sourceInfo: VariableHoverSourceInfo?,
    )

    private data class StartupVariableHoverState(
      var currentDirectory: Path,
      val variables: LinkedHashMap<String, String>,
      val sources: LinkedHashMap<String, VariableHoverSourceInfo>,
    )

    private data class DbdSearchDirectory(
      val directory: Path,
      val label: String,
    )

    private data class ResolvedReleaseRoot(
      val variableName: String,
      val rootPath: Path,
    )

    private data class DbdDeviceDeclaration(
      val recordType: String,
      val linkType: String,
      val supportName: String,
      val choiceName: String,
      val declarationText: String,
      val startOffset: Int,
    )

    private data class DbdDeviceHoverMatch(
      val absolutePath: Path,
      val line: Int,
      val startOffset: Int,
      val recordType: String,
      val linkType: String,
      val supportName: String,
      val choiceName: String,
      val declarationText: String,
      val searchLabel: String,
    )

    private data class StreamProtocolCommandHoverMatch(
      val absolutePath: Path,
      val line: Int,
      val startOffset: Int,
      val definitionText: String,
    )

    private val DATABASE_EXTENSIONS = setOf("db", "vdb", "template")
    private const val PVLIST_EXTENSION = "pvlist"
    private const val RECORD_PREVIEW_MAX_LINES = 100
    private const val RECORD_PREVIEW_MAX_CHARACTERS = 12000
    private val DOCUMENTATION_DATABASE_KEYWORDS = setOf(
      "record",
      "grecord",
      "field",
      "info",
      "alias",
      "menu",
      "choice",
      "device",
      "driver",
      "registrar",
      "function",
      "variable",
      "include",
      "breaktable",
    )
    private val RECORD_DECLARATION_REGEX = Regex("""\b(?:g?record)\s*\(\s*[A-Za-z_][A-Za-z0-9_]*\s*,\s*"((?:[^"\\]|\\.)*)"""")
    private val RECORD_DECLARATION_PREFIX_REGEX = Regex("""(record\(\s*[A-Za-z0-9_]+\s*,\s*")((?:[^"\\]|\\.)*)(")""")
    private val SUBSTITUTION_BLOCK_START_REGEX = Regex("""(?m)^\s*file\s+("?[^"\s{]+"?)\s*\{""")
    private val RELEASE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$""")
    private val DBD_DEVICE_DECLARATION_REGEX = Regex(
      """device\(\s*([A-Za-z0-9_]+)\s*,\s*([A-Za-z0-9_]+)\s*,\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)""",
    )
    private val MAKEFILE_ASSIGNMENT_REGEX = Regex(
      """^\s*([A-Za-z_][A-Za-z0-9_.-]*)\s*(\+?=|:=|\?=)\s*(.*?)\s*(?:#.*)?$""",
    )
    private val STARTUP_ENV_SET_REGEX = Regex(
      """\bepicsEnvSet(?:\s*\(\s*|\s+)\"?([A-Za-z_][A-Za-z0-9_]*)\"?\s*,\s*\"((?:[^"\\]|\\.)*)\"\s*\)?""",
    )
    private val STARTUP_CD_REGEX = Regex("""^\s*cd\s+(?:\"([^\"]+)\"|([^\s#]+))""")
    private val EPICS_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
    private val PVLIST_MACRO_ASSIGNMENT_REGEX = Regex("""^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$""")
    private val PVLIST_MACRO_REFERENCE_REGEX = Regex("""\$\(([^)=\s]+)(?:=([^)]*))?\)|\$\{([^}\s]+)\}""")
    private val DEFAULT_TOKEN_COLORS = mapOf<TextAttributesKey, Color>(
      EpicsHighlightingKeys.COMMENT to Color(0x6A, 0x99, 0x55),
      EpicsHighlightingKeys.STRING to Color(0xCE, 0x91, 0x78),
      EpicsHighlightingKeys.KEYWORD to Color(0x56, 0x9C, 0xD6),
      EpicsHighlightingKeys.NUMBER to Color(0xB5, 0xCE, 0xA8),
      EpicsHighlightingKeys.MACRO to Color(0xC5, 0x86, 0xC0),
      EpicsHighlightingKeys.BRACE to Color(0xD4, 0xD4, 0xD4),
      EpicsHighlightingKeys.PAREN to Color(0xD4, 0xD4, 0xD4),
      EpicsHighlightingKeys.BRACKET to Color(0xD4, 0xD4, 0xD4),
      EpicsHighlightingKeys.COMMA to Color(0xD4, 0xD4, 0xD4),
      EpicsHighlightingKeys.OPERATOR to Color(0xD4, 0xD4, 0xD4),
    )
  }
}

internal data class EpicsDocumentationPreview(
  val referenceKey: String,
  val html: String,
)

internal open class EpicsReferencedFileElement(
  private val manager: PsiManager,
  val hostFile: PsiFile,
  val reference: EpicsResolvedReference,
) : FakePsiElement() {
  val referenceKey: String = "${hostFile.virtualFile?.path}:${reference.rawPath}:${reference.targetFile.path}"

  override fun getParent(): PsiElement = hostFile

  override fun getContainingFile(): PsiFile = hostFile

  override fun getManager(): PsiManager = manager

  override fun getName(): String = reference.targetFile.name

  override fun getPresentableText(): String = reference.targetFile.name
}

internal class EpicsReferencedFilesElement(
  manager: PsiManager,
  hostFile: PsiFile,
  val references: List<EpicsResolvedReference>,
) : EpicsReferencedFileElement(manager, hostFile, references.first()) {
}
