package org.epics.workbench.documentation

import com.intellij.lang.documentation.AbstractDocumentationProvider
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
import org.epics.workbench.completion.EpicsRecordCompletionSupport
import org.epics.workbench.highlighting.EpicsHighlightingKeys
import org.epics.workbench.highlighting.EpicsLexingProfile
import org.epics.workbench.highlighting.EpicsSimpleLexer
import org.epics.workbench.highlighting.EpicsTokenTypes
import org.epics.workbench.navigation.EpicsPathKind
import org.epics.workbench.navigation.EpicsPathResolver
import org.epics.workbench.navigation.EpicsRecordResolver
import org.epics.workbench.navigation.EpicsResolvedReference
import org.epics.workbench.navigation.EpicsResolvedRecordDefinition
import org.epics.workbench.toc.EpicsDatabaseToc
import java.awt.Color
import java.io.File
import java.net.URLEncoder
import java.nio.charset.StandardCharsets

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

      if (isSubstitutionsFile(hostFile)) {
        val resolvedReferences = EpicsPathResolver.resolveSubstitutionsReferences(project, hostFile, offset)
        if (resolvedReferences.isNotEmpty()) {
          val referenceKey = buildReferenceKey(hostFile, resolvedReferences)
          return EpicsDocumentationPreview(referenceKey, buildDocumentation(resolvedReferences, hostFile.name))
        }
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

      return buildString {
        append("<html><body>")
        append("<h3>").append(escape(title)).append("</h3>")
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
            .append(escape("${index + 1}. ${reference.targetFile.name}"))
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
        appendPreview("Record name preview", previewNames.joinToString("\n"))
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

    private fun StringBuilder.appendPreview(label: String, content: String) {
      if (content.isBlank()) {
        return
      }
      append("<p><b>").append(escape(label)).append(":</b></p>")
      append("<pre>")
      append(escape(content))
      append("</pre>")
    }

    private fun StringBuilder.appendHighlightedPreview(label: String, content: String) {
      if (content.isBlank()) {
        return
      }
      append("<p><b>").append(escape(label)).append(":</b></p>")
      append("<pre>")
      append(renderDatabasePreviewHtml(content))
      append("</pre>")
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
      val lexer = EpicsSimpleLexer(EpicsLexingProfile(DOCUMENTATION_DATABASE_KEYWORDS))
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

    private fun readText(file: VirtualFile): String? {
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
