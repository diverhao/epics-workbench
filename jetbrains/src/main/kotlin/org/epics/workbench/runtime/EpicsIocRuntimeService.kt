package org.epics.workbench.runtime

import com.intellij.execution.executors.DefaultRunExecutor
import com.intellij.execution.process.ColoredProcessHandler
import com.intellij.execution.filters.TextConsoleBuilderFactory
import com.intellij.execution.process.ProcessEvent
import com.intellij.execution.process.ProcessListener
import com.intellij.execution.ui.ConsoleView
import com.intellij.execution.ui.RunContentDescriptor
import com.intellij.execution.ui.RunContentManager
import com.intellij.notification.NotificationAction
import com.intellij.notification.NotificationGroupManager
import com.intellij.notification.NotificationType
import com.intellij.openapi.Disposable
import com.intellij.openapi.application.ApplicationManager
import com.intellij.openapi.components.Service
import com.intellij.openapi.components.service
import com.intellij.openapi.fileEditor.FileEditorManager
import com.intellij.openapi.fileTypes.PlainTextFileType
import com.intellij.openapi.project.Project
import com.intellij.openapi.util.Disposer
import com.intellij.openapi.vfs.VirtualFile
import com.intellij.testFramework.LightVirtualFile
import com.intellij.util.messages.Topic
import com.intellij.util.execution.ParametersListUtil
import java.io.File
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.attribute.FileTime
import java.util.concurrent.atomic.AtomicReference

interface EpicsIocRuntimeStateListener {
  fun startupStateChanged(startupPath: String, running: Boolean)

  companion object {
    @JvmField
    val TOPIC: Topic<EpicsIocRuntimeStateListener> = Topic.create(
      "EPICS IOC runtime state",
      EpicsIocRuntimeStateListener::class.java,
    )
  }
}

data class EpicsIocRuntimeVariable(
  val type: String,
  val name: String,
  val value: String,
)

data class EpicsIocRuntimeEnvironmentEntry(
  val name: String,
  val value: String,
)

data class EpicsRunningIocStartup(
  val startupPath: String,
  val startupName: String,
  val consoleTitle: String,
)

data class EpicsIocDumpRecordEntry(
  val recordType: String,
  val recordName: String,
  val recordDesc: String,
  val dumpText: String,
)

internal data class EpicsStartupValidation(
  val hasShebangCommand: Boolean,
  val missingExecutableName: String? = null,
)

@Service(Service.Level.PROJECT)
class EpicsIocRuntimeService(
  private val project: Project,
) : Disposable {
  private val sessionsByStartupPath = linkedMapOf<String, IocProcessSession>()
  private val sessionLock = Any()
  private val configurationService = project.service<EpicsRuntimeProjectConfigurationService>()

  fun isRunning(startupFile: VirtualFile): Boolean {
    val startupPath = normalizeStartupPath(startupFile) ?: return false
    synchronized(sessionLock) {
      return sessionsByStartupPath[startupPath]?.isRunning == true
    }
  }

  fun getConsoleTitle(startupFile: VirtualFile): String? {
    val startupPath = normalizeStartupPath(startupFile) ?: return null
    synchronized(sessionLock) {
      return sessionsByStartupPath[startupPath]?.consoleTitle
    }
  }

  fun listRunningIocStartups(): List<EpicsRunningIocStartup> {
    return synchronized(sessionLock) {
      sessionsByStartupPath.values
        .filter { session -> session.isRunning }
        .sortedBy { session -> session.startupPath.lowercase() }
        .map { session ->
          EpicsRunningIocStartup(
            startupPath = session.startupPath,
            startupName = session.startupName,
            consoleTitle = session.consoleTitle,
          )
        }
    }
  }

  fun startIoc(startupFile: VirtualFile): Result<Unit> {
    val startupPath = normalizeStartupPath(startupFile)
      ?: return Result.failure(IllegalArgumentException("Unsupported IOC startup file."))
    synchronized(sessionLock) {
      val existing = sessionsByStartupPath[startupPath]
      if (existing != null && existing.isRunning) {
        showConsole(existing)
        return Result.success(Unit)
      }
    }

    val workingDirectory = Path.of(startupFile.path).parent
      ?: return Result.failure(IllegalStateException("Startup file does not have a parent directory."))
    val commandLine = buildStartupCommandLine(startupFile).getOrElse { error ->
      return Result.failure(error)
    }

    val process = runCatching {
      ProcessBuilder(commandLine)
        .directory(workingDirectory.toFile())
        .redirectErrorStream(true)
        .start()
    }.getOrElse { error ->
      return Result.failure(error)
    }

    val commandPresentation = commandLine.joinToString(" ")
    val handler = ColoredProcessHandler(process, commandPresentation, StandardCharsets.UTF_8)
    val console = TextConsoleBuilderFactory.getInstance().createBuilder(project).console
    console.attachToProcess(handler)

    val consoleTitle = buildConsoleTitle(startupFile)
    val descriptor = RunContentDescriptor(console, handler, console.component, consoleTitle)

    val session = IocProcessSession(
      startupPath = startupPath,
      startupName = startupFile.name,
      workingDirectory = workingDirectory,
      consoleTitle = consoleTitle,
      process = process,
      processHandler = handler,
      console = console,
      descriptor = descriptor,
    )

    handler.addProcessListener(
      object : ProcessListener {
        override fun processTerminated(event: ProcessEvent) {
          onProcessTerminated(session.startupPath)
        }
      },
    )

    synchronized(sessionLock) {
      sessionsByStartupPath[startupPath] = session
    }

    handler.startNotify()
    showConsole(session)
    notifyStartupStateChanged(startupPath, true)
    notifyConsoleEvent(
      startupFile.name,
      consoleTitle,
      "Started ${startupFile.name} in \"$consoleTitle\".",
      session,
    )
    return Result.success(Unit)
  }

  fun stopIoc(startupFile: VirtualFile) {
    val startupPath = normalizeStartupPath(startupFile) ?: return
    val session = synchronized(sessionLock) {
      sessionsByStartupPath.remove(startupPath)
    } ?: return
    session.processHandler.destroyProcess()
    notifyStartupStateChanged(startupPath, false)
    session.commandNames.set(null)
    session.commandHelp.clear()
    notifyConsoleEvent(
      startupFile.name,
      session.consoleTitle,
      "Stopped ${startupFile.name} in \"${session.consoleTitle}\".",
      session,
    )
  }

  fun showRunningConsole(startupFile: VirtualFile) {
    val startupPath = normalizeStartupPath(startupFile) ?: return
    val session = synchronized(sessionLock) { sessionsByStartupPath[startupPath] } ?: return
    showConsole(session)
  }

  fun getCommandNames(startupFile: VirtualFile): List<String> {
    val session = getRequiredSession(startupFile)
    session.commandNames.get()?.let { return it }
    val output = captureRawCommand(session, "help")
    val commands = parseCommandNames(output)
    session.commandNames.compareAndSet(null, commands)
    return session.commandNames.get().orEmpty()
  }

  fun getCommandHelp(startupFile: VirtualFile, commandNames: List<String>): Map<String, String> {
    if (commandNames.isEmpty()) {
      return emptyMap()
    }
    val session = getRequiredSession(startupFile)
    val result = linkedMapOf<String, String>()
    synchronized(session.commandLock) {
      commandNames.forEach { commandName ->
        session.commandHelp[commandName]?.let { helpText ->
          result[commandName] = helpText
        }
      }
      val missing = commandNames.filterNot { result.containsKey(it) }
      missing.chunked(HELP_BATCH_SIZE).forEach { batch ->
        val output = captureRawCommandLocked(session, "help ${batch.joinToString(" ")}")
        parseCommandHelp(batch, output).forEach { (commandName, helpText) ->
          session.commandHelp[commandName] = helpText
          result[commandName] = helpText
        }
      }
    }
    return commandNames.associateWith { result[it].orEmpty() }
  }

  fun listRuntimeVariables(startupFile: VirtualFile): List<EpicsIocRuntimeVariable> {
    val output = captureCommandOutput(startupFile, "var")
    return parseRuntimeVariables(output)
  }

  fun setRuntimeVariable(startupFile: VirtualFile, variableName: String, value: String) {
    sendCommandText(startupFile, "var $variableName $value")
  }

  fun listRuntimeEnvironment(startupFile: VirtualFile): List<EpicsIocRuntimeEnvironmentEntry> {
    val output = captureCommandOutput(startupFile, "epicsEnvShow")
    return parseRuntimeEnvironment(output)
  }

  fun setRuntimeEnvironment(startupFile: VirtualFile, name: String, value: String) {
    val escaped = value.replace("\"", "\\\"")
    sendCommandText(startupFile, "epicsEnvSet $name \"$escaped\"")
  }

  fun sendCommandText(startupFile: VirtualFile, commandText: String) {
    val session = getRequiredSession(startupFile)
    synchronized(session.commandLock) {
      sendLine(session, commandText)
    }
  }

  fun captureCommandOutput(startupFile: VirtualFile, commandText: String): String {
    val session = getRequiredSession(startupFile)
    return captureRawCommand(session, commandText)
  }

  fun openCapturedOutput(startupFile: VirtualFile, titlePrefix: String, commandText: String, output: String) {
    val content = buildString {
      append("# ")
      append(commandText)
      appendLine()
      appendLine()
      append(output.trimEnd())
      appendLine()
    }
    val suffix = titlePrefix.lowercase().replace(Regex("[^a-z0-9]+"), "-").trim('-')
    val file = LightVirtualFile(
      "epics-${suffix.ifEmpty { "ioc" }}-${startupFile.name}.txt",
      PlainTextFileType.INSTANCE,
      content,
    )
    ApplicationManager.getApplication().invokeLater {
      FileEditorManager.getInstance(project).openFile(file, true, true)
    }
  }

  fun openTemporaryOutputFile(fileName: String, content: String) {
    val file = LightVirtualFile(fileName, content)
    ApplicationManager.getApplication().invokeLater {
      FileEditorManager.getInstance(project).openFile(file, true, true)
    }
  }

  override fun dispose() {
    val sessions = synchronized(sessionLock) {
      sessionsByStartupPath.values.toList().also { sessionsByStartupPath.clear() }
    }
    sessions.forEach { session ->
      runCatching { session.processHandler.destroyProcess() }
      Disposer.dispose(session.console)
    }
  }

  private fun getRequiredSession(startupFile: VirtualFile): IocProcessSession {
    val startupPath = normalizeStartupPath(startupFile)
      ?: throw IllegalStateException("Unsupported IOC startup file.")
    return synchronized(sessionLock) {
      sessionsByStartupPath[startupPath]
    }?.takeIf { it.isRunning }
      ?: throw IllegalStateException("${startupFile.name} is not running in EPICS Workbench.")
  }

  private fun captureRawCommand(session: IocProcessSession, commandText: String): String {
    synchronized(session.commandLock) {
      return captureRawCommandLocked(session, commandText)
    }
  }

  private fun captureRawCommandLocked(session: IocProcessSession, commandText: String): String {
    val tempFile = Files.createTempFile(TEMP_FILE_PREFIX, ".txt")
    val output = runCatching {
      Files.deleteIfExists(tempFile)
      sendLine(session, "$commandText > ${tempFile.toAbsolutePath()}")
      waitForCapturedOutput(tempFile)
      sanitizeCapturedOutput(Files.readString(tempFile))
    }.also {
      runCatching { Files.deleteIfExists(tempFile) }
    }
    return output.getOrThrow()
  }

  private fun waitForCapturedOutput(tempFile: Path) {
    val deadline = System.currentTimeMillis() + CAPTURE_TIMEOUT_MS
    var lastSize = -1L
    var lastModified: FileTime? = null
    var stableReads = 0

    while (System.currentTimeMillis() < deadline) {
      if (Files.exists(tempFile)) {
        val currentSize = runCatching { Files.size(tempFile) }.getOrDefault(-1L)
        val currentModified = runCatching { Files.getLastModifiedTime(tempFile) }.getOrNull()
        if (currentSize >= 0L && currentModified != null) {
          if (currentSize == lastSize && currentModified == lastModified) {
            stableReads += 1
            if (stableReads >= 4) {
              return
            }
          } else {
            stableReads = 0
            lastSize = currentSize
            lastModified = currentModified
          }
        }
      }
      Thread.sleep(CAPTURE_POLL_INTERVAL_MS)
    }
    throw IllegalStateException("Timed out waiting for IOC output.")
  }

  private fun sendLine(session: IocProcessSession, line: String) {
    val writer = session.outputWriter
    writer.write(line)
    writer.newLine()
    writer.flush()
  }

  private fun onProcessTerminated(startupPath: String) {
    val session = synchronized(sessionLock) {
      sessionsByStartupPath.remove(startupPath)
    } ?: return
    notifyStartupStateChanged(startupPath, false)
    session.commandNames.set(null)
    session.commandHelp.clear()
  }

  private fun notifyStartupStateChanged(startupPath: String, running: Boolean) {
    project.messageBus.syncPublisher(EpicsIocRuntimeStateListener.TOPIC)
      .startupStateChanged(startupPath, running)
  }

  private fun showConsole(session: IocProcessSession) {
    RunContentManager.getInstance(project).showRunContent(
      DefaultRunExecutor.getRunExecutorInstance(),
      session.descriptor,
    )
  }

  private fun notifyConsoleEvent(
    startupName: String,
    consoleTitle: String,
    message: String,
    session: IocProcessSession,
  ) {
    NotificationGroupManager.getInstance()
      .getNotificationGroup(NOTIFICATION_GROUP_ID)
      .createNotification(
        startupName,
        "$message Console: $consoleTitle",
        NotificationType.INFORMATION,
      )
      .addAction(
        NotificationAction.createSimpleExpiring("Show Console") {
          showConsole(session)
        },
      )
      .notify(project)
  }

  private fun normalizeStartupPath(startupFile: VirtualFile): String? {
    if (!isIocBootStartupFile(startupFile)) {
      return null
    }
    return Path.of(startupFile.path).normalize().toString()
  }

  private fun buildConsoleTitle(startupFile: VirtualFile): String {
    val parentName = startupFile.parent?.name?.takeIf(String::isNotBlank) ?: startupFile.nameWithoutExtension
    return "IOC: $parentName"
  }

  private fun buildStartupCommandLine(startupFile: VirtualFile): Result<List<String>> {
    val startupPath = Path.of(startupFile.path).normalize()
    val shebangResolution = resolveStartupShebangCommandLine(startupPath)
    if (shebangResolution.missingExecutableName != null) {
      return Result.failure(
        IllegalStateException("Error: executable ${shebangResolution.missingExecutableName} missing"),
      )
    }
    shebangResolution.commandLine?.let { return Result.success(it) }

    val configuration = configurationService.loadConfiguration()
    val startupCommand = "./${startupFile.name}"
    val shellPath = configuration.iocStartupShell.trim()
    if (shellPath.isEmpty()) {
      return Result.success(listOf(startupCommand))
    }
    val shellArgs = ParametersListUtil.parse(configuration.iocStartupShellArgs.trim()).toMutableList()
    shellArgs += startupCommand
    return Result.success(listOf(shellPath) + shellArgs)
  }

  private data class IocProcessSession(
    val startupPath: String,
    val startupName: String,
    val workingDirectory: Path,
    val consoleTitle: String,
    val process: Process,
    val processHandler: ColoredProcessHandler,
    val console: ConsoleView,
    val descriptor: RunContentDescriptor,
  ) {
    val outputWriter = process.outputWriter(StandardCharsets.UTF_8)
    val commandLock = Any()
    val commandNames = AtomicReference<List<String>?>(null)
    val commandHelp = linkedMapOf<String, String>()

    val isRunning: Boolean
      get() = process.isAlive
  }

  companion object {
    private const val NOTIFICATION_GROUP_ID = "EPICS Workbench Notifications"
    private const val TEMP_FILE_PREFIX = "epics-workbench-ioc-output-"
    private const val CAPTURE_TIMEOUT_MS = 8_000L
    private const val CAPTURE_POLL_INTERVAL_MS = 50L
    private const val HELP_BATCH_SIZE = 24
    private val ANSI_ESCAPE_REGEX = Regex("""\u001B\[[0-9;?]*[ -/]*[@-~]""")

    fun isStartupFile(file: VirtualFile?): Boolean {
      val target = file ?: return false
      val extension = target.extension?.lowercase()
      return extension == "cmd" || extension == "iocsh" || target.name == "st.cmd"
    }

    internal fun validateStartupFile(file: VirtualFile?, text: String? = null): EpicsStartupValidation {
      val target = file ?: return EpicsStartupValidation(hasShebangCommand = false)
      val startupPath = runCatching { Path.of(target.path).normalize() }.getOrNull()
        ?: return EpicsStartupValidation(hasShebangCommand = false)
      val resolution = resolveStartupShebangCommandLine(startupPath, text)
      return EpicsStartupValidation(
        hasShebangCommand = resolution.hasShebangCommand,
        missingExecutableName = resolution.missingExecutableName,
      )
    }

    fun isIocBootStartupFile(file: VirtualFile?): Boolean {
      val target = file ?: return false
      if (!isStartupFile(target)) {
        return false
      }
      return Path.of(target.path).normalize().any { pathPart -> pathPart.toString() == "iocBoot" }
    }

    private fun displayExecutableName(executableText: String): String {
      val trimmed = executableText.trim().trimEnd('/', '\\')
      return trimmed.substringAfterLast('/').substringAfterLast('\\').ifEmpty { trimmed }
    }

    private fun resolveStartupShebangCommandLine(startupPath: Path, text: String? = null): StartupShebangResolution {
      val firstLine = text
        ?.substringBefore('\n')
        ?.removeSuffix("\r")
        ?: runCatching {
          Files.newBufferedReader(startupPath).use { reader -> reader.readLine().orEmpty() }
        }.getOrDefault("")
      if (!firstLine.startsWith("#!")) {
        return StartupShebangResolution()
      }

      val shebangCommand = firstLine.removePrefix("#!").trim()
      if (shebangCommand.isEmpty()) {
        return StartupShebangResolution()
      }

      val shebangParts = ParametersListUtil.parse(shebangCommand)
        .map(String::trim)
        .filter(String::isNotEmpty)
      if (shebangParts.isEmpty()) {
        return StartupShebangResolution()
      }

      val executablePath = resolveExistingExecutablePath(
        executableText = shebangParts.first(),
        baseDirectory = startupPath.parent,
      )
      if (executablePath == null) {
        return StartupShebangResolution(
          hasShebangCommand = true,
          missingExecutableName = displayExecutableName(shebangParts.first()),
        )
      }

      return StartupShebangResolution(
        hasShebangCommand = true,
        commandLine = buildList {
          add(executablePath.toString())
          addAll(shebangParts.drop(1))
          add(startupPath.toString())
        },
      )
    }

    private fun resolveExistingExecutablePath(executableText: String, baseDirectory: Path?): Path? {
      val normalizedExecutable = executableText.trim()
      if (normalizedExecutable.isEmpty()) {
        return null
      }

      if (normalizedExecutable.contains('/') || normalizedExecutable.contains('\\')) {
        val candidatePath = if (Path.of(normalizedExecutable).isAbsolute) {
          Path.of(normalizedExecutable)
        } else {
          (baseDirectory ?: Path.of("")).resolve(normalizedExecutable)
        }.normalize()
        return candidatePath.takeIf(Files::exists)
      }

      val pathEntries = System.getenv("PATH")
        .orEmpty()
        .split(File.pathSeparatorChar)
        .map(String::trim)
        .filter(String::isNotEmpty)
      for (pathEntry in pathEntries) {
        val candidatePath = Path.of(pathEntry, normalizedExecutable).normalize()
        if (Files.exists(candidatePath)) {
          return candidatePath
        }
      }

      return null
    }

    fun parseCommandNames(output: String): List<String> {
      val commands = linkedSetOf<String>()
      for (rawLine in output.lineSequence()) {
        val trimmed = rawLine.trim()
        if (trimmed.isEmpty()) {
          continue
        }
        if (trimmed.startsWith("Type 'help")) {
          break
        }
        val line = trimmed.removePrefix("#").trim()
        line.split(Regex("\\s+"))
          .filter { token -> COMMAND_NAME_REGEX.matches(token) }
          .forEach(commands::add)
      }
      return commands.toList()
    }

    fun parseCommandHelp(commandNames: List<String>, output: String): Map<String, String> {
      val requested = commandNames.toSet()
      val result = linkedMapOf<String, String>()
      var currentCommand: String? = null
      var currentBlock = mutableListOf<String>()

      fun flush() {
        val commandName = currentCommand ?: return
        val text = currentBlock.joinToString("\n").trim()
        if (text.isNotEmpty()) {
          result[commandName] = text
        }
      }

      output.lineSequence().forEach { rawLine ->
        val line = rawLine.trimEnd()
        if (line.isBlank()) {
          if (currentCommand != null) {
            currentBlock.add("")
          }
          return@forEach
        }
        val headerCandidate = line.trim().substringBefore(' ')
        if (headerCandidate in requested && !line.startsWith("Example:")) {
          flush()
          currentCommand = headerCandidate
          currentBlock = mutableListOf(line.trim())
        } else if (currentCommand != null) {
          currentBlock.add(line.trim())
        }
      }
      flush()
      return result
    }

    fun parseRuntimeVariables(output: String): List<EpicsIocRuntimeVariable> {
      return output.lineSequence()
        .map(String::trim)
        .filter(String::isNotEmpty)
        .mapNotNull { line ->
          VARIABLE_LINE_REGEX.matchEntire(line)?.let { match ->
            EpicsIocRuntimeVariable(
              type = match.groupValues[1].trim(),
              name = match.groupValues[2].trim(),
              value = match.groupValues[3].trim(),
            )
          }
        }
        .toList()
    }

    fun parseRuntimeEnvironment(output: String): List<EpicsIocRuntimeEnvironmentEntry> {
      return output.lineSequence()
        .map(String::trim)
        .filter(String::isNotEmpty)
        .mapNotNull { line ->
          val equalsIndex = line.indexOf('=')
          if (equalsIndex <= 0) {
            null
          } else {
            EpicsIocRuntimeEnvironmentEntry(
              name = line.substring(0, equalsIndex).trim(),
              value = line.substring(equalsIndex + 1),
            )
          }
        }
        .toList()
    }

    fun parseDbDumpRecordOutput(output: String): List<EpicsIocDumpRecordEntry> {
      val records = mutableListOf<EpicsIocDumpRecordEntry>()
      var activeRecordType = ""
      var activeRecordName = ""
      var activeRecordDesc = ""
      var activeLines = mutableListOf<String>()

      fun flushRecord() {
        if (activeRecordName.isBlank() || activeLines.isEmpty()) {
          activeRecordType = ""
          activeRecordName = ""
          activeRecordDesc = ""
          activeLines = mutableListOf()
          return
        }
        records += EpicsIocDumpRecordEntry(
          recordType = activeRecordType,
          recordName = activeRecordName,
          recordDesc = activeRecordDesc,
          dumpText = activeLines.joinToString("\n").trimEnd(),
        )
        activeRecordType = ""
        activeRecordName = ""
        activeRecordDesc = ""
        activeLines = mutableListOf()
      }

      output.replace("\r", "").lineSequence().forEach { line ->
        val recordMatch = DB_DUMP_RECORD_HEADER_REGEX.matchEntire(line)
        if (recordMatch != null) {
          flushRecord()
          activeRecordType = recordMatch.groupValues[1].trim()
          activeRecordName = recordMatch.groupValues[2].trim()
          activeLines += line
          return@forEach
        }

        if (activeRecordName.isBlank()) {
          return@forEach
        }

        activeLines += line
        if (activeRecordDesc.isBlank()) {
          val descMatch = DB_DUMP_RECORD_DESC_REGEX.matchEntire(line)
          if (descMatch != null) {
            activeRecordDesc = decodeDbDumpQuotedString(descMatch.groupValues[1])
          }
        }
        if (line.trim() == "}") {
          flushRecord()
        }
      }
      flushRecord()
      return records
    }

    private fun sanitizeCapturedOutput(text: String): String {
      return ANSI_ESCAPE_REGEX.replace(text, "")
    }

    private fun decodeDbDumpQuotedString(value: String): String {
      val result = StringBuilder()
      var index = 0
      while (index < value.length) {
        val current = value[index]
        if (current != '\\' || index + 1 >= value.length) {
          result.append(current)
          index += 1
          continue
        }

        when (val escaped = value[index + 1]) {
          '\\' -> result.append('\\')
          '"' -> result.append('"')
          'n' -> result.append('\n')
          'r' -> result.append('\r')
          't' -> result.append('\t')
          else -> result.append(escaped)
        }
        index += 2
      }
      return result.toString()
    }

    private val COMMAND_NAME_REGEX = Regex("""^[A-Za-z_][A-Za-z0-9_]*$""")
    private val VARIABLE_LINE_REGEX = Regex("""^(.*?)\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$""")
    private val DB_DUMP_RECORD_HEADER_REGEX =
      Regex("""^\s*record\(\s*([^,]+?)\s*,\s*"([^"\n]+)"\s*\)\s*\{\s*$""")
    private val DB_DUMP_RECORD_DESC_REGEX =
      Regex("""^\s*field\(\s*DESC\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)\s*$""")
  }

  private data class StartupShebangResolution(
    val hasShebangCommand: Boolean = false,
    val commandLine: List<String>? = null,
    val missingExecutableName: String? = null,
  )
}
