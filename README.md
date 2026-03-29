# EPICS Workbench

EPICS Workbench is a developer toolkit for EPICS applications inside modern IDEs.
This repository currently contains:

- a VS Code extension with the broadest feature set
- a JetBrains plugin for IntelliJ Platform IDEs
- a sample EPICS application used for development and manual testing

The project is focused on everyday EPICS work: reading and editing database files, navigating links, validating Makefiles and startup scripts, starting IOCs, inspecting live PVs, and moving data between EPICS text files and Excel.

This work is sponsored by Oak Ridge National Laboratory.

## Repository Layout

- `vscode/`: VS Code extension source, docs, tasks, packaging files, and vendored TDM runtime
- `jetbrains/`: JetBrains plugin source, docs, and task tracker
- `appTest/`: sample EPICS application for trying features in both IDEs
- `README.md`: repo-level overview

## Current Status

- VS Code is the most complete implementation today.
- JetBrains now has solid parity for core editing, formatting, build, Makefile, IOC, widget, and context-menu workflows.
- TDM display hosting is currently a VS Code feature.

## What The Project Can Do

Across the two IDE integrations, EPICS Workbench can help with:

- EPICS-aware handling of `.db`, `.vdb`, `.template`, `.sub`, `.subs`, `.substitutions`, `st.cmd`, `.cmd`, `.iocsh`, `.dbd`, `.proto`, `.pvlist`, `.probe`, and EPICS `Makefile` files
- EPICS project detection from the real application root markers: top-level `Makefile`, `configure/RELEASE`, and `configure/RULES_TOP`
- navigation for records, links, macros, field names, record types, DBD symbols, and related EPICS source symbols
- find-references and rename for EPICS symbols instead of plain-text search/replace
- diagnostics for duplicate records, duplicate fields, invalid fields, invalid menu values, invalid numeric values, missing Makefile inclusion, startup path problems, and unknown IOC registration functions
- quick fixes for common database and startup-file errors
- formatters for database files, substitutions files, startup files, and Makefiles
- `Add to Makefile` for `.db`, `.template`, and substitutions files when they are not yet listed in `DB += ...`
- local `Build` / `Clean` from a folder `Makefile`
- project-wide `Build Project` / `Clean Project` from any file inside an EPICS application
- Excel import/export for database content
- runtime widgets: `Probe`, `PV List`, and `PV Monitor`
- IOC start/stop helpers for `st.cmd`-like files under `iocBoot/`
- IOC runtime helpers for commands, variables, and environment inspection

### VS Code Highlights

The VS Code extension additionally includes:

- grouped `EPICS ...` editor and explorer context menus
- macro-aware database link lookup using Table of Contents macro assignments
- widget context menus inside `Probe`, `PV List`, and `PV Monitor`
- TDM display hosting for `.tdl`, `.edl`, `.bob`, `.stp`, and `.plt`
- vendored TDM runtime support so the extension can launch its own packaged TDM copy

### JetBrains Highlights

The JetBrains plugin currently includes:

- the same core EPICS project detection rules as VS Code
- EPICS popup groups for `EPICS Database File`, `EPICS Build`, `EPICS Widgets`, `EPICS Import/Export`, `EPICS Format`, and IOC runtime actions
- dynamic `EPICS IOC Start` and `EPICS IOC Stop` submenus that list discovered `iocBoot/**` startup files
- startup and Makefile formatter parity

## Quick Start

### VS Code

To develop or try the VS Code extension from source:

```bash
cd vscode
code .
```

Then:

1. Press `F5` to launch the Extension Development Host.
2. In the new window, open `appTest/`.
3. Start with:
   - `appTest/test01App/Db/test01.db`
   - `appTest/test01App/Db/test02.substitutions`
   - `appTest/test01App/src/Makefile`
   - `appTest/iocBoot/ioctest01/st.cmd`
   - `appTest/dbd/test01.dbd`

### JetBrains

To run the JetBrains plugin sandbox:

```bash
cd jetbrains
export JAVA_HOME="/Applications/IntelliJ IDEA.app/Contents/jbr/Contents/Home"
./gradlew runIde
```

Then open `appTest/` in the sandbox IDE and try the same sample files.

## Quick Manual

### 1. Database files

Open:

```text
appTest/test01App/Db/test01.db
```

Useful things to try:

- `Go to Definition` or `Find All References` on a linked record name
- `Rename Symbol` on a record or macro
- `EPICS Database File -> Add to Makefile` on a file that is not yet in `DB += ...`
- `EPICS Format -> Format File`
- `EPICS Widgets -> Probe`, `PV List`, or `PV Monitor`

What to expect:

- record-aware navigation instead of plain text lookup
- EPICS-specific validation and quick fixes
- Makefile updates inserted between `include $(TOP)/configure/CONFIG` and `include $(TOP)/configure/RULES`

### 2. Substitutions files

Open:

```text
appTest/test01App/Db/test02.substitutions
```

Useful things to try:

- `EPICS Import/Export -> Export Database to Excel`
- `EPICS Widgets -> PV List`
- `EPICS Database File -> Add to Makefile`
- `EPICS Format -> Format File`

What to expect:

- substitutions expansion for runtime views and export
- Makefile handling that maps substitutions input to the generated `.db` install token

### 3. Startup files and IOC control

Open:

```text
appTest/iocBoot/ioctest01/st.cmd
```

Useful things to try:

- completion for `dbLoadRecords("...")`
- completion for missing macro assignments in `dbLoadRecords(...)`
- completion for `dbpf("...")`
- `EPICS IOC Start/Stop` in VS Code
- `EPICS IOC Start` or `EPICS IOC Stop` in JetBrains
- IOC runtime commands, variables, and environment views

What to expect:

- file suggestions for loaded DB files
- record suggestions based on DBs already loaded earlier in the startup file
- IOC launch helpers that understand `iocBoot/**` startup scripts

### 4. Makefiles and builds

Open:

```text
appTest/test01App/src/Makefile
```

Useful things to try:

- `EPICS Build -> Build`
- `EPICS Build -> Clean`
- `EPICS Build -> Build Project`
- `EPICS Build -> Clean Project`

What to expect:

- folder-local build actions when the current file's folder has a `Makefile`
- project-wide build actions anywhere inside an EPICS application

### 5. Runtime widgets

Open any database or startup file and use:

- `EPICS Widgets -> Probe`
- `EPICS Widgets -> PV List`
- `EPICS Widgets -> PV Monitor`

These tools stay inside IDE tabs instead of requiring a separate external runtime tool.

### 6. TDM displays in VS Code

VS Code can open supported TDM display files directly in editor tabs.

Supported display types:

- `.tdl`
- `.edl`
- `.bob`
- `.stp`
- `.plt`

Use:

- `EPICS: Open in TDM`
- `Open in TDM`

The extension hosts the display in a custom editor tab and uses the vendored TDM runtime by default.

## Recommended Reading Order

If you are new to the repository:

1. Read [vscode/README.md](vscode/README.md) for the most complete user-facing feature walkthrough.
2. Open [appTest](appTest) and try the example files above.
3. Check [vscode/TASKS.md](vscode/TASKS.md) and [jetbrains/TASKS.md](jetbrains/TASKS.md) for the current implementation history and backlog.
4. Read [jetbrains/README.md](jetbrains/README.md) if you are working on the IntelliJ Platform side.

## Notes

- The VS Code extension packages a vendored TDM runtime, so the extension footprint is larger than a normal text-only extension.
- The sample project in [appTest](appTest) is the fastest way to verify most features manually.

## License

This repository is licensed under the MIT License. See [LICENSE](LICENSE).

Vendored and third-party components may carry their own licenses and notices. In particular, the vendored TDM snapshot under `vscode/vendor/tdm` already declares `MIT`.
