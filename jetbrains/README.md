# JetBrains Plugin Development

This folder contains the JetBrains version of `epics-workbench`.

Use one plugin project here for both `IntelliJ IDEA` and `PyCharm`.
Keep the implementation on common IntelliJ Platform APIs so the same plugin ZIP can be tested in both IDEs.

Planning files for development:

- `ROADMAP.md`: longer-term milestones and architecture direction
- `TASKS.md`: current backlog and active implementation queue

## Prerequisites

1. Install `IntelliJ IDEA`.
2. Use the bundled JBR 21 from IntelliJ IDEA as the Gradle JVM.
   On this machine that path is:

   ```bash
   /Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home
   ```

3. Ensure internet access is available the first time Gradle resolves IntelliJ Platform dependencies.

## Project Layout

```text
jetbrains/
  build.gradle.kts
  settings.gradle.kts
  gradle.properties
  src/main/kotlin/org/epics/workbench/
  src/main/resources/META-INF/plugin.xml
```

## Step-by-Step Development Loop

1. Open `~/projects/epics-workbench/jetbrains` in `IntelliJ IDEA`.
2. When IntelliJ asks for the Gradle JVM, select:

   ```text
   /Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home
   ```

3. Let IntelliJ import the Gradle project.
4. Start the sandbox IDE:

   ```bash
   cd ~/projects/epics-workbench/jetbrains
   export JAVA_HOME="/Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home"
   ./gradlew runIde
   ```

5. In the sandbox IDE, verify the scaffold:
   - Open `Tools -> Show EPICS Workbench Status`
   - Open the `EPICS Workbench` tool window

6. Edit plugin code under:

   ```text
   src/main/kotlin/org/epics/workbench/
   src/main/resources/META-INF/plugin.xml
   ```

7. Re-run the sandbox after each change:

   ```bash
   ./gradlew runIde
   ```

## Build and Test

### Build the installable plugin ZIP

```bash
cd ~/projects/epics-workbench/jetbrains
export JAVA_HOME="/Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home"
./gradlew buildPlugin
```

The ZIP will be created under:

```text
build/distributions/
```

### Test in IntelliJ IDEA

Use `./gradlew runIde` for the fastest development loop.

### Test in PyCharm

1. Build the plugin ZIP with `./gradlew buildPlugin`.
2. Open `PyCharm`.
3. Go to `Settings -> Plugins -> gear icon -> Install Plugin from Disk...`
4. Select the ZIP from `build/distributions/`.
5. Restart PyCharm and verify the plugin loads.

## Marketplace Release

The Gradle build is prepared to sign and publish the plugin:

- Signing files are auto-discovered from `../dev-doc/private.pem` and `../dev-doc/chain.crt`.
- You can override those paths with `PRIVATE_KEY_FILE` and `CERTIFICATE_CHAIN_FILE`.
- If the private key is encrypted, provide `PRIVATE_KEY_PASSWORD`.
- Publishing uses `PUBLISH_TOKEN`.
- Release metadata can come from environment variables or Gradle properties.

### Supported release variables

- `PLUGIN_VERSION`: plugin version to publish, for example `0.1.0`
- `PLUGIN_CHANGE_NOTES`: HTML or plain text change notes for this release
- `PLUGIN_VENDOR_NAME`: defaults to `EPICS Workbench`
- `PLUGIN_VENDOR_EMAIL`: vendor contact email shown in Marketplace
- `PLUGIN_VENDOR_URL`: vendor website shown in Marketplace
- `JETBRAINS_PUBLISH_CHANNEL`: defaults to `default`
- `JETBRAINS_PUBLISH_HIDDEN`: optional, `true` or `false`
- `PUBLISH_TOKEN`: JetBrains Marketplace permanent token

### Release commands

```bash
cd ~/projects/epics-workbench/jetbrains
export JAVA_HOME="/Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home"

export PLUGIN_VERSION="0.1.0"
export PLUGIN_CHANGE_NOTES="<p>First Marketplace release.</p>"
export PLUGIN_VENDOR_EMAIL="maintainer@example.org"
export PLUGIN_VENDOR_URL="https://example.org/epics-workbench"
export PUBLISH_TOKEN="perm:..."

./gradlew buildPlugin signPlugin publishPlugin
```

You can also place the non-secret values in `~/.gradle/gradle.properties` instead of exporting them every time:

```properties
pluginVersion=0.1.0
pluginChangeNotes=<p>First Marketplace release.</p>
pluginVendorName=EPICS Workbench
pluginVendorEmail=maintainer@example.org
pluginVendorUrl=https://example.org/epics-workbench
jetbrainsPublishChannel=default
jetbrainsPublishHidden=false
```

### Manual steps still required

1. Create or confirm the JetBrains Marketplace vendor profile.
2. Decide the public vendor email and vendor URL.
3. Create the Marketplace permanent token and export it as `PUBLISH_TOKEN`.
4. Pick the release version and final change notes for the release.
5. Run `./gradlew signPlugin publishPlugin` once those values are set.

## Recommended Next Milestones

1. Register EPICS file types: `.db`, `.template`, `.substitutions`, `.cmd`, `.dbd`, `.proto`, `.st`, `.monitor`
2. Add syntax highlighting and lexer/parser support
3. Add navigation and hover documentation
4. Add IOC/CA/PVA runtime panels
5. Add inspections, formatting, and editor actions

## Working With Codex

The fastest collaboration loop is task-based:

1. Keep feature work in `TASKS.md`
2. Ask Codex to implement one task or one milestone
3. Let Codex make the code changes and run Gradle checks
4. Verify behavior in the sandbox IDE started by `./gradlew runIde`

Preferred request style:

- `Implement J-01`
- `Implement the next 2 JetBrains tasks`
- `Finish Milestone 1 in jetbrains/TASKS.md`

Avoid line-by-line editing instructions unless the change is very small.
