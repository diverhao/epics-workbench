# EPICS Workbench for VS Code

This document describes what the VS Code extension can do today and how to use it in practice.

The examples below use the sample project in `appTest/`.

## What The Extension Understands

The extension understands these EPICS-oriented file types:

- `.db`, `.vdb`, `.template`
- `.sub`, `.subs`, `.substitutions`
- `st.cmd`, `.cmd`, `.iocsh`
- `.dbd`
- `.proto`
- `.pvlist`
- `.probe`
- EPICS C/C++ source files used by DBD exports
- EPICS `Makefile`, `configure/RELEASE`, and `configure/RELEASE.local`

It builds an internal EPICS project model automatically. That model is used for:

- navigation
- hover and peeker popups
- validation and diagnostics
- completion
- rename and find-references
- widgets such as Probe, PV List, and PV Monitor

The build-model layer also interrogates EPICS Makefiles with `make -pn` when possible. If that fails, the extension falls back to file-based project discovery automatically.

## How To Start

Open an EPICS application folder in VS Code, for example:

```bash
code /Users/1h7/projects/epics-workbench/appTest
```

For extension development, open the extension source and run the Extension Development Host:

```bash
code /Users/1h7/projects/epics-workbench/vscode
```

Then press `F5`, and in the Extension Development Host open:

```bash
/Users/1h7/projects/epics-workbench/appTest
```

## Database Files: `.db`, `.vdb`, `.template`

### What You Can Do

- hover linked records and jump to their definitions
- `Find All References` on records, record links, macros, record types, and field names
- `Rename Symbol` for the same symbol kinds
- validate duplicate records, duplicate fields, invalid fields, invalid numeric values, invalid menu values, and long `DESC`
- quick-fix duplicate record names
- quick-fix invalid field names
- quick-fix invalid menu values
- `Probe`
- `PV List`
- `PV Monitor`
- `Export Excel`
- `Import Excel as EPICS DB`
- `Format File`

### Practical Example: Record References

Open:

`appTest/test01App/Db/test01.db`

Put the caret on:

```db
field(INP, "abcd2")
```

Then run:

- `Shift+F12` for `Find All References`
- `F12` for `Go to Definition`

Expected result:

- the definition of `record(ao, "abcd2")`
- other references to `abcd2`

### Practical Example: Macro-Aware Record References

Open:

`appTest/test01App/Db/test01.db`

This file defines:

```text
#  - S = pv
```

and contains:

```db
field(INP, "$(S)2222")
```

The extension uses the file-local TOC macro assignments for DB link lookup. That means the lookup target is treated as `pv2222`, even though the editor still shows `$(S)2222`.

### Practical Example: Invalid Field Quick Fix

In a database file, change a valid field name to an invalid one:

```db
field(OUT, "33")
```

to:

```db
field(OTT, "33")
```

Then place the caret on the error and press:

- macOS: `Cmd+.`
- Windows/Linux: `Ctrl+.`

Expected result:

- a quick fix appears with the closest valid field name

### Practical Example: Invalid Menu Value Quick Fix

Change:

```db
field(ZSV, "MAJOR")
```

to:

```db
field(ZSV, "BAD")
```

Then trigger quick fix. The extension offers a valid menu replacement.

## Startup Files: `st.cmd`, `.cmd`, `.iocsh`

### What You Can Do

- hover `epicsEnvSet(...)` macros and see resolved values
- hover `dbLoadRecords(...)` and `dbLoadTemplate(...)` file paths
- hover `dbpf(...)` record names loaded earlier in the file
- completion for startup commands
- completion for `dbLoadRecords(...)` first argument
- completion for empty macro assignments in `dbLoadRecords(...)`
- completion for `dbpf(...)` record names loaded earlier in the startup file
- diagnostics for unresolved startup paths
- diagnostics for missing `dbLoadRecords(...)` macro assignments
- diagnostics for unknown IOC registration functions
- quick fix for missing `dbLoadRecords(...)` macros
- quick fix for unknown IOC registration function
- `Probe`
- `PV List`
- `PV Monitor`
- `Format File`
- `Import Excel`

### Practical Example: `dbLoadRecords(...)` Completion

Open:

`appTest/iocBoot/ioctest01/st.cmd`

Type:

```iocsh
dbLoadRecords("
```

Expected result:

- only `.db`, `.vdb`, and `.template` files from the current IOC shell directory and its `db/` folder are suggested
- folders are not shown
- the dropdown label shows only the file name

If the chosen DB file contains macros, selecting the file inserts the empty macro tail automatically. Example:

```iocsh
dbLoadRecords("db/asynInt32TimeSeries.db", "P=,R=,PORT=,ADDR=,TIMEOUT=,DRVINFO=,NELM=,SCAN=")
```

### Practical Example: `dbpf(...)` Completion

In the same startup file, type:

```iocsh
dbpf("
```

Expected result:

- suggested channel names come only from DB and substitutions files loaded earlier in the same startup file through `dbLoadRecords(...)` and `dbLoadTemplate(...)`

### Practical Example: IOC Registration Quick Fix

In `st.cmd`, change:

```iocsh
test01_registerRecordDeviceDriver(pdbbase)
```

to:

```iocsh
wrong_registerRecordDeviceDriver(pdbbase)
```

Then trigger quick fix on the error.

Expected result:

- a replacement suggestion using the known IOC registration function for the current application

## Substitutions Files: `.sub`, `.subs`, `.substitutions`

### What You Can Do

- hover the `file "...db"` or `file "...template"` path and peek matching template files
- expand substitutions into `PV List`
- expand substitutions into `Export Excel`
- `Probe` opens blank by design
- `PV Monitor` opens blank by design
- `Format File`
- `Import Excel`

### Practical Example

Right-click a substitutions file and choose:

- `PV List`

Expected result:

- the PV List widget opens with the expanded record set from the substitutions file

Right-click the same file and choose:

- `Export Excel`

Expected result:

- the expanded database text is exported to Excel
- the bottom-right notification contains an action to open the generated workbook in the operating system default app

## PV List Widget

### What It Does

The `PV List` widget is the runtime-oriented list view for channels. It replaces the old idea of showing live channel values inline inside `.pvlist` files.

### What You Can Do

- start from a database file, substitutions file, startup file, `.pvlist`, or blank widget
- monitor all listed channels
- edit macro values inline
- double-click a value to put a new value
- `Configure Channels` to edit the full raw channel list
- channels preserve their order
- macro values from database TOC and `.pvlist` files are honored when the widget opens

### Practical Example

Right-click:

`appTest/test01App/Db/test02.db`

Choose:

- `PV List`

Expected result:

- the widget opens with the channels from the DB
- if the source file defines macros, their values are preloaded into the widget

Use `Configure Channels` to open the editor page. That page shows the full current raw channel list, including macros. Edit the list, click `OK`, and the runtime list updates to match the content and order of the text box.

## Probe Widget

### What It Does

The `Probe` widget is the single-channel live inspector. It shows:

- value
- record type
- last update
- access
- fields for the record type

### What You Can Do

- open it blank from some contexts
- open it pre-targeted from DB/startup contexts when a record can be resolved
- double-click the main value or a field value to put a new value
- for record widgets with `Process`, force `record.PROC = 1`

### Practical Example

Right-click a record name in a DB file and choose:

- `Probe`

Expected result:

- a file-less probe widget opens
- if a resolvable record was under the cursor, it is preloaded

## PV Monitor Widget

### What It Does

The `PV Monitor` widget is event-driven runtime monitoring. It does not sample periodically. It appends a new line whenever CA/PVA delivers a new value.

### What You Can Do

- open the widget blank or from context
- add channels
- monitor values continuously
- keep the viewport pinned to the bottom when you are already at the bottom
- keep the viewport fixed when you scroll up to inspect older data
- export monitor history as plain text data

### Practical Example

Right-click a DB file or startup file and choose:

- `PV Monitor`

Then add one or more channels.

If you are at the bottom of the view, the widget auto-scrolls as new values arrive. If you scroll upward, the widget keeps appending new data without moving your visible region.

## `.pvlist` Files

### What They Are For

A `.pvlist` file is a plain EPICS list file, not a runtime file.

Typical content:

```text
SYS = aaa
SUBSYS =

$(SYS)1
val2
$(SUBSYS)3
```

### What You Can Do

- use record hover/peeker on channel-name lines
- `Probe` from the current line
- `PV Monitor` from the current line
- `PV List` from the whole file
- `Format File`

### Practical Example

Open a `.pvlist` file and place the caret on a record line such as:

```text
val2
```

Then:

- hover to peek the record definition
- use `Probe` to open a single-channel probe

## `.dbd` And EPICS Source Files

### What You Can Do

- autocompletion for DBD keywords and exported symbol names
- `Find All References` across `.dbd` and EPICS C/C++ source files
- `Rename Symbol` across those files
- `Probe`, `PV List`, and `PV Monitor` from `.dbd` context open blank
- `Format File`

### Practical Example: DBD / Source Cross-Reference

Open:

- `appTest/test01App/src/devAdd10.dbd`
- `appTest/test01App/src/devAdd10.c`

Try these symbols:

- `aaabbbccc`
- `AABBCC`
- `myParameter`
- `myFunction`

Use:

- `Shift+F12` for `Find All References`
- `F2` for `Rename Symbol`

Expected result:

- `.dbd` and C occurrences participate in the same reference/rename operation

## Build Model And Project Understanding

### What It Does

The extension automatically discovers EPICS applications from:

- `configure/RELEASE`
- `configure/RELEASE.local`
- application `Makefile`s
- generated runtime artifacts

It models:

- IOC names
- runtime DB, DBD, and substitutions artifacts
- RELEASE-expanded dependency roots
- available DBDs and libraries
- startup entry-point files

### Direct Debugging Use

The build model is internal to the IDE, but you can inspect the raw JSON with:

```bash
node vscode/scripts/epics-build-model.js --root /Users/1h7/projects/epics-workbench/appTest
```

This is useful when you want to confirm how the extension is interpreting an EPICS application.

## Safe Edit / Refactor Features

### What Is Implemented Today

- `Find All References` for:
  - records
  - macros
  - record types
  - field names
  - DBD/source symbols
- `Rename Symbol` for the same categories
- rename conflict analysis for records, macros, record types, field names, and DBD/source symbols
- quick fixes for common EPICS diagnostics
- one VS Code `WorkspaceEdit` per rename, so built-in rename preview can review the whole EPICS edit set

### Practical Example: Conflict Check

Open:

`appTest/test01App/Db/test01.db`

Try renaming:

- `abcd2` to `abcd3`

Expected result:

- rename is rejected because `abcd3` already exists in the workspace

## TDM Display Integration

The extension can open TDM display files directly inside VS Code tabs through a custom editor.

Supported file types:

- `.tdl`
- `.edl`
- `.bob`
- `.stp`
- `.plt`

How it works:

- EPICS Workbench starts TDM in `web` mode
- it creates a display window agent through TDM's `/command` API
- it hosts the resulting `DisplayWindow.html?...` page inside a VS Code webview tab
- when that TDM display opens another display, the extension intercepts the popup request and opens another VS Code tab instead

By default, the extension now prefers the vendored TDM runtime copied into:

- `vscode/vendor/tdm`

That vendored runtime is started through a Node launcher owned by the extension instead of a separately managed TDM binary. The old executable-based launch path is still available as an explicit fallback.

### Required Settings

Set these in VS Code if auto-detection is not enough:

- `epicsWorkbench.tdm.runtimeMode`
- `epicsWorkbench.tdm.rootPath`
- `epicsWorkbench.tdm.executablePath`
- `epicsWorkbench.tdm.profilesPath`
- `epicsWorkbench.tdm.profile`

If you want the extension-owned runtime, leave `epicsWorkbench.tdm.runtimeMode` at `vendoredSource`.

If you need to fall back to a separately built TDM checkout, switch `epicsWorkbench.tdm.runtimeMode` to `externalBinary` and point `epicsWorkbench.tdm.executablePath` at a runnable launcher if auto-detection is not enough.

The extension defaults to:

- `vendoredSource` launch mode when `vscode/vendor/tdm` exists
- `externalBinary` only when explicitly selected or when the vendored runtime is missing
- `TDM_ROOT`
- a sibling `../tdm` checkout next to the current workspace for `externalBinary`
- `~/.tdm/profiles.json`
- the first non-`For All Profiles` profile in that file

### Updating The Vendored TDM Copy

To refresh the copied TDM tree from a local checkout without editing the original source in place, run:

```sh
npm run sync:tdm-vendor
```

or:

```sh
python3 ./scripts/sync-vendored-tdm.py
```

The sync script copies a runtime-focused vendored snapshot:

- top-level TDM metadata needed to identify and launch the vendored runtime
- the built web assets and runtime layout under `vscode/vendor/tdm/dist`
- the runtime dependency closure needed by the vendored Node-hosted web mode

It intentionally does not copy `vscode/vendor/tdm/src`, because the current VS Code vendored runtime launches from `dist` and reads runtime assets from `dist/common/resources`.

### How To Use It

Open a supported TDM display file in VS Code, or run:

- `EPICS: Open in TDM`

You can also right-click a supported display file in the explorer and choose:

- `Open in TDM`

## Known Scope Limits

- DB TOC macro-aware record lookup is implemented for database record references. The references panel still shows the literal source text from the file.
- Record-type and field-schema rename is intended for schema declared in workspace `.dbd` files. Built-in EPICS base schema is not blindly renamed by text replacement.
- TDM integration still uses a local proxy and localhost-oriented display hosting. The vendored runtime removes the separate TDM executable dependency, but the Remote/SSH cleanup is not finished yet.
- The integration hosts TDM display windows, not the full TDM main-window/profile-management UI.
- Some runtime widgets intentionally open blank from certain file types:
  - `Probe` from substitutions
  - `PV Monitor` from substitutions
  - `PV List` from `.dbd`

## Recommended Manual Test Set

Use these files for a fast regression pass:

- `appTest/test01App/Db/test01.db`
- `appTest/test01App/Db/test02.db`
- `appTest/iocBoot/ioctest01/st.cmd`
- `appTest/test01App/src/devAdd10.dbd`
- `appTest/test01App/src/devAdd10.c`

For each one, test:

- hover
- `Go to Definition`
- `Find All References`
- `Rename Symbol`
- quick fix where applicable
- context-menu widgets where applicable
