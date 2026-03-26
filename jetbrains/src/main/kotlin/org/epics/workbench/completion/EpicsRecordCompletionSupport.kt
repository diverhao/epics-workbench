package org.epics.workbench.completion

import com.intellij.application.options.CodeStyle
import com.intellij.openapi.project.Project
import com.intellij.openapi.vfs.VirtualFile
import org.epics.workbench.build.EpicsBuildModelPathEntry
import org.epics.workbench.build.epicsBuildModelService
import kotlinx.serialization.json.Json
import kotlinx.serialization.json.JsonArray
import kotlinx.serialization.json.JsonElement
import kotlinx.serialization.json.JsonObject
import kotlinx.serialization.json.JsonPrimitive
import kotlinx.serialization.json.jsonObject
import kotlinx.serialization.json.jsonPrimitive
import java.nio.file.Files
import java.nio.file.Path
import java.util.concurrent.ConcurrentHashMap
import kotlin.io.path.extension
import kotlin.io.path.exists
import kotlin.io.path.isDirectory
import kotlin.io.path.name
import kotlin.io.path.pathString

internal object EpicsRecordCompletionSupport {
  private val json = Json { ignoreUnknownKeys = true }

  private data class StaticData(
    val fieldOrderByRecordType: Map<String, List<String>>,
    val templateFieldsByRecordType: Map<String, List<String>>,
    val fieldTypesByRecordType: Map<String, Map<String, String>>,
    val fieldMenuChoicesByRecordType: Map<String, Map<String, List<String>>>,
    val fieldInitialValuesByRecordType: Map<String, Map<String, String>>,
  )

  internal data class RecordDeclaration(
    val recordType: String,
    val name: String,
    val nameStart: Int,
    val nameEnd: Int,
    val recordStart: Int,
    val recordEnd: Int,
  )

  internal data class FieldDeclaration(
    val fieldName: String,
    val fieldNameStart: Int,
    val fieldNameEnd: Int,
    val value: String,
    val valueStart: Int,
    val valueEnd: Int,
  )

  internal data class MenuFieldValueContext(
    val recordType: String,
    val recordName: String,
    val fieldName: String,
    val value: String,
    val valueStart: Int,
    val valueEnd: Int,
    val choices: List<String>,
  )

  private val workspaceRecordTypeCache = ConcurrentHashMap<String, List<String>>()

  private val staticData: StaticData by lazy(LazyThreadSafetyMode.PUBLICATION) {
    StaticData(
      fieldOrderByRecordType = parseStringListMap("data/embedded-record-fields.json"),
      templateFieldsByRecordType = parseStringListMap("data/record-template-fields.json"),
      fieldTypesByRecordType = parseNestedStringMap("data/embedded-record-field-types.json"),
      fieldMenuChoicesByRecordType = parseNestedStringListMap("data/embedded-record-field-menus.json"),
      fieldInitialValuesByRecordType = parseNestedStringMap("data/embedded-record-field-initials.json"),
    )
  }

  fun getRecordTypes(project: Project, hostFile: VirtualFile): List<String> {
    val labels = linkedSetOf<String>()
    labels += staticData.templateFieldsByRecordType.keys
    labels += loadWorkspaceRecordTypes(project, hostFile)
    return labels.sortedWith(String.CASE_INSENSITIVE_ORDER)
  }

  fun getTemplateFields(recordType: String): List<String> {
    return staticData.templateFieldsByRecordType[recordType]
      ?: COMMON_RECORD_FIELDS
  }

  fun getFieldNamesForRecordType(recordType: String?): List<String> {
    if (!recordType.isNullOrBlank()) {
      val labels = linkedSetOf<String>()
      staticData.fieldOrderByRecordType[recordType]?.let(labels::addAll)
      staticData.templateFieldsByRecordType[recordType]?.let(labels::addAll)
      staticData.fieldTypesByRecordType[recordType]?.keys?.let(labels::addAll)
      if (labels.isNotEmpty()) {
        return labels.toList()
      }
    }

    val fallbackLabels = linkedSetOf<String>()
    fallbackLabels += COMMON_RECORD_FIELDS
    fallbackLabels += staticData.fieldTypesByRecordType.values.flatMap { it.keys }
    return fallbackLabels.toList()
  }

  fun getDeclaredFieldNamesForRecordType(recordType: String): Set<String>? {
    if (recordType.isBlank()) {
      return null
    }

    val labels = linkedSetOf<String>()
    staticData.templateFieldsByRecordType[recordType]?.let(labels::addAll)
    staticData.fieldTypesByRecordType[recordType]?.keys?.let(labels::addAll)
    if (labels.isEmpty()) {
      return null
    }
    return labels.mapTo(linkedSetOf()) { it.uppercase() }
  }

  fun getAvailableFieldNamesForRecordInstance(
    documentText: String,
    offset: Int,
    recordType: String?,
  ): List<String> {
    val fieldNames = getFieldNamesForRecordType(recordType)
    if (recordType.isNullOrBlank()) {
      return fieldNames
    }

    val recordDeclaration = extractRecordDeclarations(documentText)
      .firstOrNull { declaration -> offset in declaration.recordStart..declaration.recordEnd }
      ?: return fieldNames

    val existingFields = extractFieldDeclarationsInRecord(documentText, recordDeclaration)
      .map { declaration -> declaration.fieldName }
      .toSet()

    return fieldNames.filterNot(existingFields::contains)
  }

  fun getDefaultFieldValue(recordType: String, fieldName: String): String {
    val dbfType = staticData.fieldTypesByRecordType[recordType]?.get(fieldName)
    val explicit = staticData.fieldInitialValuesByRecordType[recordType]?.get(fieldName)
    if (explicit != null) {
      return resolveExplicitFieldInitialValue(recordType, fieldName, dbfType, explicit)
    }

    return when (dbfType) {
      in NUMERIC_DBF_TYPES -> "0"
      "DBF_MENU" -> getMenuFieldChoices(recordType, fieldName).firstOrNull().orEmpty()
      else -> ""
    }
  }

  fun getFieldType(recordType: String, fieldName: String): String? {
    return staticData.fieldTypesByRecordType[recordType]?.get(fieldName)
  }

  fun isNumericFieldType(dbfType: String?): Boolean {
    return dbfType in NUMERIC_DBF_TYPES
  }

  fun isIntegerFieldType(dbfType: String?): Boolean {
    return dbfType in INTEGER_DBF_TYPES
  }

  fun containsEpicsMacroReference(value: String): Boolean {
    return EPICS_MACRO_REFERENCE_REGEX.containsMatchIn(value)
  }

  fun isSkippableNumericFieldValue(value: String): Boolean {
    val trimmedValue = value.trim()
    return trimmedValue.isBlank() || containsEpicsMacroReference(trimmedValue)
  }

  fun isValidNumericFieldValue(value: String, dbfType: String): Boolean {
    val numericValue = value.trim().toDoubleOrNull() ?: return false
    if (!numericValue.isFinite()) {
      return false
    }
    if (isIntegerFieldType(dbfType) && numericValue % 1.0 != 0.0) {
      return false
    }
    return true
  }

  fun getMenuFieldChoices(recordType: String, fieldName: String): List<String> {
    if (recordType.isBlank() || fieldName.isBlank()) {
      return emptyList()
    }
    return staticData.fieldMenuChoicesByRecordType[recordType]?.get(fieldName).orEmpty()
  }

  fun findEnclosingRecordType(text: String, offset: Int): String? {
    return extractRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.recordStart..declaration.recordEnd }
      ?.recordType
  }

  fun findMenuFieldValueContext(text: String, offset: Int): MenuFieldValueContext? {
    val recordDeclaration = extractRecordDeclarations(text)
      .firstOrNull { declaration -> offset in declaration.recordStart..declaration.recordEnd }
      ?: return null

    val fieldDeclaration = extractFieldDeclarationsInRecord(text, recordDeclaration)
      .firstOrNull { declaration -> offset in declaration.valueStart until declaration.valueEnd }
      ?: return null

    if (getFieldType(recordDeclaration.recordType, fieldDeclaration.fieldName) != "DBF_MENU") {
      return null
    }

    val choices = getMenuFieldChoices(recordDeclaration.recordType, fieldDeclaration.fieldName)
    if (choices.isEmpty()) {
      return null
    }

    return MenuFieldValueContext(
      recordType = recordDeclaration.recordType,
      recordName = recordDeclaration.name,
      fieldName = fieldDeclaration.fieldName,
      value = fieldDeclaration.value,
      valueStart = fieldDeclaration.valueStart,
      valueEnd = fieldDeclaration.valueEnd,
      choices = choices,
    )
  }

  fun extractRecordDeclarations(text: String): List<RecordDeclaration> {
    val declarations = mutableListOf<RecordDeclaration>()
    val sanitizedText = maskHashComments(text)
    val regex = Regex("""\brecord\(\s*([A-Za-z0-9_]+)\s*,\s*"((?:[^"\\]|\\.)*)"""")

    for (match in regex.findAll(sanitizedText)) {
      val recordType = match.groups[1]?.value.orEmpty()
      val recordName = match.groups[2]?.value.orEmpty()
      val prefixLength = Regex("""record\(\s*[A-Za-z0-9_]+\s*,\s*"""").find(match.value)?.value?.length ?: continue
      val recordStart = match.range.first
      val recordEnd = findRecordBlockEnd(sanitizedText, recordStart)
      declarations += RecordDeclaration(
        recordType = recordType,
        name = recordName,
        nameStart = match.range.first + prefixLength,
        nameEnd = match.range.first + prefixLength + recordName.length,
        recordStart = recordStart,
        recordEnd = recordEnd,
      )
    }

    return declarations
  }

  fun extractFieldDeclarationsInRecord(
    text: String,
    recordDeclaration: RecordDeclaration,
  ): List<FieldDeclaration> {
    val declarations = mutableListOf<FieldDeclaration>()
    val sanitizedText = maskHashComments(text)
    val recordText = sanitizedText.substring(recordDeclaration.recordStart, recordDeclaration.recordEnd)
    val regex = Regex("""field\(\s*(?:"((?:[^"\\]|\\.)*)"|([A-Za-z0-9_]+))\s*,\s*"((?:[^"\\]|\\.)*)"""")
    val valuePrefixRegex = Regex("""field\(\s*(?:"(?:[^"\\]|\\.)*"|[A-Za-z0-9_]+)\s*,\s*"""")

    for (match in regex.findAll(recordText)) {
      val rawFieldName = match.groups[1]?.value ?: match.groups[2]?.value ?: continue
      val fieldName = rawFieldName.uppercase()
      val fieldNameStartInMatch = match.groups[1]?.range?.first ?: match.groups[2]?.range?.first ?: continue
      val fieldNameEndInMatch = fieldNameStartInMatch + rawFieldName.length
      val valuePrefixLength = valuePrefixRegex.find(match.value)?.value?.length ?: continue
      val value = match.groups[3]?.value.orEmpty()
      declarations += FieldDeclaration(
        fieldName = fieldName,
        fieldNameStart = recordDeclaration.recordStart + fieldNameStartInMatch,
        fieldNameEnd = recordDeclaration.recordStart + fieldNameEndInMatch,
        value = value,
        valueStart = recordDeclaration.recordStart + match.range.first + valuePrefixLength,
        valueEnd = recordDeclaration.recordStart + match.range.first + valuePrefixLength + value.length,
      )
    }

    return declarations
  }

  fun getIndentUnit(file: com.intellij.psi.PsiFile): String {
    val indentOptions = CodeStyle.getIndentOptions(file)
    return if (indentOptions.USE_TAB_CHARACTER) "\t" else " ".repeat(indentOptions.INDENT_SIZE.coerceAtLeast(4))
  }

  private fun parseStringListMap(resourcePath: String): Map<String, List<String>> {
    val root = loadJsonObject(resourcePath)
    return root.entries.associate { (key, value) ->
      key to value.asStringList()
    }
  }

  private fun parseNestedStringMap(resourcePath: String): Map<String, Map<String, String>> {
    val root = loadJsonObject(resourcePath)
    return root.entries.associate { (key, value) ->
      key to value.jsonObject.entries.associate { (nestedKey, nestedValue) ->
        nestedKey to nestedValue.jsonPrimitive.content
      }
    }
  }

  private fun parseNestedStringListMap(resourcePath: String): Map<String, Map<String, List<String>>> {
    val root = loadJsonObject(resourcePath)
    return root.entries.associate { (key, value) ->
      key to value.jsonObject.entries.associate { (nestedKey, nestedValue) ->
        nestedKey to nestedValue.asStringList()
      }
    }
  }

  private fun loadJsonObject(resourcePath: String): JsonObject {
    val stream = javaClass.classLoader.getResourceAsStream(resourcePath)
      ?: return JsonObject(emptyMap())
    val text = stream.bufferedReader().use { it.readText() }
    return json.parseToJsonElement(text).jsonObject
  }

  private fun JsonElement.asStringList(): List<String> {
    return (this as? JsonArray)?.mapNotNull { element ->
      (element as? JsonPrimitive)?.content
    }.orEmpty()
  }

  private fun loadWorkspaceRecordTypes(project: Project, hostFile: VirtualFile): List<String> {
    val ownerRoot = findOwningEpicsRoot(project, hostFile)
    val releaseVariables = loadReleaseVariables(ownerRoot)
    val buildModel = project.epicsBuildModelService().loadBuildModel(ownerRoot)
    val searchRoots = buildSearchRoots(project, ownerRoot, releaseVariables)
    val cacheKey = searchRoots.joinToString("|") { it.pathString }

    return workspaceRecordTypeCache.computeIfAbsent(cacheKey) {
      val names = linkedSetOf<String>()
      val buildModelDbdEntries = buildModel?.availableDbds.orEmpty()
      if (buildModelDbdEntries.isNotEmpty()) {
        collectRecordTypesFromBuildModelEntries(buildModelDbdEntries, names)
      }
      for (root in searchRoots) {
        val dbdDirectory = root.resolve("dbd")
        if (!dbdDirectory.exists() || !dbdDirectory.isDirectory()) {
          continue
        }
        try {
          Files.walk(dbdDirectory).use { stream ->
            stream
              .filter { path -> Files.isRegularFile(path) && path.fileName.toString().lowercase().endsWith(".dbd") }
              .forEach { path ->
                val text = runCatching { Files.readString(path) }.getOrNull() ?: return@forEach
                RECORD_TYPE_REGEX.findAll(text).forEach { match ->
                  val recordType = match.groups[1]?.value?.trim().orEmpty()
                  if (recordType.isNotEmpty()) {
                    names += recordType
                  }
                }
              }
          }
        } catch (_: Exception) {
          continue
        }
      }
      names.sortedWith(String.CASE_INSENSITIVE_ORDER)
    }
  }

  private fun findOwningEpicsRoot(project: Project, hostFile: VirtualFile): Path {
    var current = hostFile.parent?.toNioPath()
    while (current != null) {
      if (current.resolve("configure").resolve("RELEASE").exists()) {
        return current
      }
      current = current.parent
    }
    return Path.of(project.basePath ?: hostFile.parent?.path ?: ".").normalize()
  }

  private fun loadReleaseVariables(ownerRoot: Path): Map<String, String> {
    val configureDirectory = ownerRoot.resolve("configure")
    val releaseFiles = listOf(
      configureDirectory.resolve("RELEASE"),
      configureDirectory.resolve("RELEASE.local"),
    )
    val rawValues = linkedMapOf("TOP" to ownerRoot.pathString)
    for (releaseFile in releaseFiles) {
      if (!releaseFile.exists()) {
        continue
      }
      releaseFile.toFile().readLines().forEach { line ->
        val match = RELEASE_ASSIGNMENT_REGEX.find(line) ?: return@forEach
        rawValues[match.groups[1]?.value.orEmpty()] = match.groups[2]?.value?.trim().orEmpty()
      }
    }

    val resolved = linkedMapOf<String, String>()
    val resolving = mutableSetOf<String>()

    fun resolve(name: String): String? {
      resolved[name]?.let { return it }
      if (name in resolving) {
        return rawValues[name]
      }
      val raw = rawValues[name] ?: return null
      resolving += name
      val expanded = expandEpicsValue(raw, rawValues) { nested -> resolve(nested) }
      val normalized = if (expanded.isNotBlank() && !Path.of(expanded).isAbsolute) {
        ownerRoot.resolve(expanded).normalize().pathString
      } else {
        expanded
      }
      resolving -= name
      resolved[name] = normalized
      return normalized
    }

    rawValues.keys.forEach { resolve(it) }
    return resolved
  }

  private fun buildSearchRoots(
    project: Project?,
    ownerRoot: Path,
    releaseVariables: Map<String, String>,
  ): List<Path> {
    project?.let { currentProject ->
      return currentProject.epicsBuildModelService().collectSearchRoots(ownerRoot, releaseVariables, emptyMap())
    }

    val roots = linkedSetOf<Path>()
    roots.add(ownerRoot.normalize())
    releaseVariables.values.forEach { value ->
      val candidate = runCatching { Path.of(value) }.getOrNull() ?: return@forEach
      if (candidate.exists() && candidate.isDirectory()) {
        roots.add(candidate.normalize())
      }
    }
    return roots.toList()
  }

  private fun collectRecordTypesFromBuildModelEntries(
    entries: List<EpicsBuildModelPathEntry>,
    names: MutableSet<String>,
  ) {
    for (entry in entries) {
      val absolutePath = entry.absolutePath?.let { path ->
        runCatching { Path.of(path) }.getOrNull()
      } ?: continue
      if (!absolutePath.exists() || absolutePath.extension.lowercase() != "dbd") {
        continue
      }
      val text = runCatching { Files.readString(absolutePath) }.getOrNull() ?: continue
      RECORD_TYPE_REGEX.findAll(text).forEach { match ->
        val recordType = match.groups[1]?.value?.trim().orEmpty()
        if (recordType.isNotEmpty()) {
          names += recordType
        }
      }
    }
  }

  private fun resolveExplicitFieldInitialValue(
    recordType: String,
    fieldName: String,
    dbfType: String?,
    explicitInitialValue: String,
  ): String {
    if (dbfType != "DBF_MENU") {
      return explicitInitialValue
    }

    val choices = getMenuFieldChoices(recordType, fieldName)
    if (choices.isEmpty()) {
      return explicitInitialValue
    }
    if (choices.contains(explicitInitialValue)) {
      return explicitInitialValue
    }

    val trimmedValue = explicitInitialValue.trim()
    if (trimmedValue.all(Char::isDigit)) {
      val index = trimmedValue.toIntOrNull()
      if (index != null && index in choices.indices) {
        return choices[index]
      }
    }

    return explicitInitialValue
  }

  private fun expandEpicsValue(
    rawValue: String,
    variables: Map<String, String>,
    fallbackResolver: ((String) -> String?)? = null,
  ): String {
    var expanded = rawValue
    repeat(10) {
      val next = EPICS_VARIABLE_REGEX.replace(expanded) { match ->
        val name = match.groups[1]?.value
          ?: match.groups[3]?.value
          ?: match.groups[5]?.value
          ?: ""
        val defaultValue = match.groups[2]?.value ?: match.groups[4]?.value
        variables[name] ?: fallbackResolver?.invoke(name) ?: defaultValue ?: match.value
      }
      if (next == expanded) {
        return next.trim().trim('"')
      }
      expanded = next
    }
    return expanded.trim().trim('"')
  }

  private fun maskHashComments(text: String): String {
    val sanitized = StringBuilder(text.length)
    var inString = false
    var escaped = false
    var inComment = false

    for (character in text) {
      if (inComment) {
        if (character == '\n') {
          inComment = false
          sanitized.append(character)
        } else {
          sanitized.append(' ')
        }
        continue
      }

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
        inComment = true
        sanitized.append(' ')
        continue
      }

      sanitized.append(character)
    }

    return sanitized.toString()
  }

  private fun findRecordBlockEnd(text: String, recordStart: Int): Int {
    val openingBrace = text.indexOf('{', recordStart)
    if (openingBrace < 0) {
      return text.length
    }

    var depth = 0
    var inString = false
    var escaped = false

    for (index in openingBrace until text.length) {
      val character = text[index]
      if (escaped) {
        escaped = false
        continue
      }

      if (character == '\\') {
        escaped = true
        continue
      }

      if (character == '"') {
        inString = !inString
        continue
      }

      if (inString) {
        continue
      }

      when (character) {
        '{' -> depth += 1
        '}' -> {
          depth -= 1
          if (depth == 0) {
            return index + 1
          }
        }
      }
    }

    return text.length
  }

  private val COMMON_RECORD_FIELDS = listOf(
    "DESC",
    "SCAN",
    "PINI",
    "DTYP",
    "DISA",
    "SDIS",
    "FLNK",
    "VAL",
    "PREC",
    "EGU",
    "INP",
    "OUT",
    "TSEL",
    "DOL",
    "OMSL",
    "SIML",
    "SIMM",
    "SIMS",
    "SIOL",
    "HOPR",
    "LOPR",
  )

  private val NUMERIC_DBF_TYPES = setOf(
    "DBF_SHORT",
    "DBF_ENUM",
    "DBF_UCHAR",
    "DBF_UINT64",
    "DBF_ULONG",
    "DBF_USHORT",
    "DBF_DOUBLE",
    "DBF_INT64",
    "DBF_LONG",
  )

  private val INTEGER_DBF_TYPES = setOf(
    "DBF_SHORT",
    "DBF_ENUM",
    "DBF_UCHAR",
    "DBF_UINT64",
    "DBF_ULONG",
    "DBF_USHORT",
    "DBF_INT64",
    "DBF_LONG",
  )

  private val RELEASE_ASSIGNMENT_REGEX = Regex("""^\s*([A-Za-z_][A-Za-z0-9_.-]*)\s*=\s*(.*?)\s*(?:#.*)?$""")
  private val EPICS_VARIABLE_REGEX = Regex("""\$\(([^)=]+)(?:=([^)]*))?\)|\$\{([^}=]+)(?:=([^}]*))?\}|\$([A-Za-z_][A-Za-z0-9_.-]*)""")
  private val EPICS_MACRO_REFERENCE_REGEX = Regex("""\$\(|\$\{""")
  private val RECORD_TYPE_REGEX = Regex("""\brecordtype\(\s*([A-Za-z0-9_]+)\s*\)""")
}
