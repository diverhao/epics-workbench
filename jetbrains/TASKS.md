# JetBrains Tasks

Status legend:

- `todo`
- `in_progress`
- `done`
- `blocked`

## Current Priority

### J-01 File Types

Status: `done`

Implement EPICS file type registration for:

- `.db`
- `.vdb`
- `.template`
- `.substitutions`
- `.sub`
- `.subs`
- `.cmd`
- `.iocsh`
- `.dbd`
- `.proto`
- `.st`
- `.monitor`

Acceptance:

- each file extension opens with a non-plain-text EPICS file type
- display names are sensible in the UI
- `./gradlew buildPlugin` passes

### J-02 Sample Files and Manual Smoke Test

Status: `blocked`

Add a small set of sample EPICS files under `jetbrains/src/testData/` for manual and future automated testing.

Acceptance:

- include at least one file for database, substitutions, startup, and dbd
- samples are small and readable
- README or TASKS references how they are used

### J-03 Basic Syntax Highlighting Strategy

Status: `done`

Choose and implement the initial highlighting approach for JetBrains:

- direct lexer-based path if it is still lightweight enough
- syntax colors should be similar to that in vscode epics-workbench extension
- syntax highlight should applied to the corresponding file types in vscode extension

Acceptance:

- database, substitutions, st.cmd-like files show visible syntax differentiation

### J-10 Resolve file path

Status: `done`

A test environment is in `~/projects/epics-workbench/appTest/`, it is a full EPICS application.

The file path should be resolved similar to the vscode version:
 - the current host file's location should be used first to resolve relative path
 - the relative path is w.r.t. the current folder or the current location in `st.cmd`. Because in `st.cmd`, we may have `cd $DIR_1`, where `DIR_1` is a macro defined in `configure/RELEASE` or `envPaths`.
 - then we search the folder locations defined in `configure/RELEASE` or `envPaths`. 

For particular file types:
 - For database (.db and .template) file, we should look for the `db` folder in search paths
 - for lib file, we should look for the `lib/architecture (e.g. darwin/linux_x86-64)` folder
 - for dbd files, should be in `dbd` folder

Acceptance:
 - in `st.cmd`-like file, the `db` file paths can be successfully resolved
 - in `Makefile`, the `.dbd` files and library files can be successfully resolved
 - in `.substitutions`, the `.db` and `.template` should be successfully resolved

### J-11 File information peek via floating window

Status: `done`

When mouse hovers over a `.db`, `.substitutions` file name in `st.cmd`-like, Makefile, and `.substitutions` file, show the vital info about this file and peek the file contents. Vscode version extension already provides a good example.

Acceptance:
 - mouse hover triggers a floating window
 - this floating window should be similar to the epics-workbench extension in vscode (see `~/projects/epics-workbench/vscode`)


## J-12 Resolve record definition location 

Status: `done`

in epics database, a record is referenced in a link-type field: the field could be a `INLINK`, `OUTLINK` or `FWDLINK`. These links are defined in epics base. There is an epics base at `/Users/1h7/Tools/epics-base`.

in `st.cmd`-like file, the record name is referenced in ioc shell (iocsh) commands like `dbpf("record:name", "value")`. 

The record referenced in a database file (`.db` or `.template`) or st.cmd-like file should be resolved to its definition location. The search path is
 - current db file's (`.db`, `.template`) folder, then try the `db` folders defined in `configure/RELEASE` and `envPaths`
 - for `st.cmd`-like file, use the current location defined in previous lines, e.g. `cd $ASYN`, then we should search for the `$ASYN/db` fodler. If in previous lines there is no path defined, then use this `st.cmd`'s current folder

this functionality is already in vscode version, you can refer to it.

Acceptance:
 - record definition location (file + line) can be resolved for a record name referenced in `st.cmd`-like, `.db` and `.template` file

### J-13 Show floating window when mosue hovers record name

Status: `done`

Show floating window when mouse hovers on a record name in `st.cmd`-like and database file. The ref to record name are in the locations specified in J-12. The vscode version of this extension already implemented this functionality. The floating window should be similar to the file content peeking window in this jetbrains extension.

The record content should be syntax highlighted in the floating box.

### J-14 Generate table of contents in database file

Status: `done`

Add an option in context menu (right click menu), clicking which will generate a table-of-contents of this database file. The TOC format should be similar to that in vscode version, i.e. looks like

```
# EPICS TOC BEGIN
# Macros:
#  - ABC
#  - IOCNAME
# Table of Contents
# | Record                | Value              | Type |
# | --------------------- | ------------------ | ---- |
# | $(IOCNAME)$(ABC)abcd1 |                    | ai   |
# | abcd2                 |                    | ao   |
# | abcd5                 |                    | aSub |
# | abcd3                 |                    | ao   |
# | abcd4                 |                    | bo   |
# EPICS TOC END
```

The `Value` column is reserved for runtime value which will be implemented later.

User can add field names, such as `INP`, `FLNK` after the `# Table of Contents`, then regenerate the TOC with these field values shown in the table.

There are many details, you should refer to the vscode version.


### J-15 EPICS file formatter

Status: `done`

Add file formatter for epics related files. The formatted result should be same as the vscode extension.

### J-16 Autoclosing 

Status: `done`

When user types, auto close the paired symbols, like those in this file

```
~/projects/epics-workbench/vscode/lang-config.json
```

### J-17 Autocompletion for record in database file

Status: `done`

When user types `record(`, it should show a list of available record types, `ai`, `ao`, ..., these types are defined in `xxxRecord.dbd` in epics base. After user selectes the record type, it autocompletes the whole record with commonly used fields implemented. Then bring the cursor to the location of record name, show a dropdown menu which lists all existing record names defined in this file to let the user select.

This functionality is already implemented in vscode version. Pls refer to it for details. 


### J-18 Dropdown menu for `menu` type record field

Status: `done`

A record's field could be a `menu` type, which is defined in `xxxRecord.dbd` and `menuXXX.dbd` in epics base. When user types `field(SCAN, "`, I want to show the choices for this field, e.g. `.1 second`, `.2 second` ... for `SCAN` field. Vscode Version Already Implemented This.

### J-18.5 Autocompletion for field in record 

Status: `done`

In database file, when I type `field(` inside a record, I would like to show a dropdown menu with all fields in this specific record type. This functionality is already in VsCode version.

### J-18.6 Autocompletion for field value in record

Status: `done`

Catch up request from J-18.5, after the field name is selected from dropdown menu, the field value should be filled with default value. For string type field, it is `""`, for number type value, it is `0`. Notice that some fields have special default value, as defined in dbd file, e.g. in `aaiRecord.dbd`

```
	field(NELM,DBF_ULONG) {
		prompt("Number of Elements")
		promptgroup("30 - Action")
		special(SPC_NOMOD)
		interest(1)
		initial("1")
		prop(YES)		# get_graphic_double, get_control_double
	}
```

the initial default value of NELM field is 1. This feature is already in vscode verion.

### J-19 Value type check for field types in record

Status: `done`

Each field in epics record has a data type, like `string`, `number` (INT32, DOUBLE ...), or menu choice type. Add the syntax semantic check for the field's value in database file, if the type is wrong, e.g. `field(SCAN, "1.5 second")` should be labeled as an error because there is no `1.5 second` option in `SCAN` field. For number type data, only check if it is a number, no need to verify the exact type (int, float). Vscode version already implemented this.


### J-20 Autocompletion for dbLoadRecords command in st.cmd-like file

Status: `done`

Add autocompletion for `dbLoadRecords` command in st.cmd-like file. After use types `dbL`, show the choice for `dbLoadRecords`, after the user selects this choice, complete the command, then show a dropdown menu showing the available `.db` and `.template` files from `db/` folder in current project and the folders/projects defined in `configure/RELEASE`. One tricky thing is `st.cmd` is a shell-like environment, the `cd $XXX` or `cd /abc/xxx` means currently the shell is switched to that folder. So, the relative path in the db file selection dropdown menu should be w.r.t. the current folder. This functionality is already implemented in VsCode version. You can refer to it.


Catch up: the file paths from dropdown menu should include only these folders
 - the `db/` folder in current dir (not the st.cmd's dir, but the ioc shell's current dir). If the IOCSH is not `cd`-ed to other folders, then use this project's `db/` folder
 - the `db/` folder in other folders defined in `configure/RELEASE` in this project

No other folders should be included

### J-21 Auto create macros input entries for dbLoadRecords

Status: `done`

After the user select db file in `dbLoadRecords`, automatically add the macros definitions in this db file, like

`dbLoadRecord("../../db/abc.db", "SYS=,SUBSYS=")`

where `SYS` and `SUBSYS` are macros defined in file `abc.db`. 

### J-22 Semantic check for db file

Status: `done`

in a database file (.db or .template), if the record name is duplicated, show an error: red underline below the record name.

in a record in the database file, each record has specific collection of fields, as defined in the record's dbd file in epics base, e.g. `aiRecord.dbd` for `ai` type record. If a record in database file contains fields that do not belong to this record, show it as an error: show a red underline under the field name. 

if the field is duplicated in one record, show an error. 

if the field's `(...)` is not closed, show an error

if the record body's `{...}` is not closed, show an error

if the record's header `(...)` is not closed, show an error

The above features are already in VsCode version.

### J-23 File navigation

Status: `done`

When the mouse hovers over files in st.cmd-like file or in Makefile, a floating window shows the basic info and refined content of this file. The file path is also in the floating box. Implement link to the path, when user clicks the path link, open the file.


### J-24 Semantic check for macro value assignments in dbLoadRecords command

Status: `done`

In dbLoadRecords commands, if the macro is not defined, or there are excessive macros definition, show an error.

### J-25 Add support for StreamDevice proto file

Status: `done`

If a record's DTYP is `"stream"`, its INP/OUT fields (maybe more link-type fields) may use StreamDevice protocol file, in format of 

```
field(OUT, "@checksums.proto")
```

the `checksums.proto` is a file. We should search it in `STREAM_PROTOCOL_PATH` which is defined in `st.cmd`-like file. Add a floating window/box when mouse hovers on the proto file name, this window/box includes similar info as db file's floating window.

### J-26 record resolution and navigation for epics command dbpf in st.cmd

Status: `done`

epics ioc shell (iocsh) has a command `dbpf`, it is invoked in form

```
dbpf("record_name_1", "record_value_1")
```

Show the record's basic info when mosue hovers over the record name. Record should be resolved according to the database files defined in this st.cmd file

Covered by the existing J-12 and J-13 implementation.

### J-27 Show dropdown menu for menu type field's value when mouse hovers

Status: `done`

When mouse hovers on the menu type field's value, e.g. when mouse is over the `"1 second"` in `field(SCAN, "1 second")`, show a dropdown menu showing all choices


Catch up: the choices in the dropdown menu, when click a choice, replace the field value with the choice value.

### J-28 Integreate JCA/CAJ and PVAClient Java libs 

Status: `done`

Download and install the JCA/CAJ and PVAClient java libs. Integrate them into this extension, so that we can later use their APIs.


### J-29 Fix db table of contents field value

Status: `done`

In table of contents in database file's header, if field is not explicitly defined in record, it is wrongly shown as `NA`. e.g.


```
# | Record                    | Value              | Type      | DTYP   | INP                                       | OUT                                       |
# | ------------------------- | ------------------ | --------- | ------ | ----------------------------------------- | ----------------------------------------- |
# | $(S):SET_CURR_SWEEP_LLIM4 |                    | ao        | NA     | NA                                        | NA                                        |
```

this record is ao type, it has OUT field (even the OUT field is not explicitly defined). In the OUT field cell, we should put a `""`.

The cell should be shown as its default value if this field exists in this record type. For string type field, the default value is empty string; for number type field, the default value is 0. Some fields have predefined default values which are defined in the field's dbd file, like

```
	field(NELM,DBF_ULONG) {
		prompt("Number of Elements")
		promptgroup("30 - Action")
		special(SPC_NOMOD)
		interest(1)
		initial("1")
		prop(YES)		# get_graphic_double, get_control_double
	}
```

where the inital value for NELM is `1` as it is DBF_ULONG type.


### J-30 Peek record in db table of contents

Status: `done`

When mouse hovers over the record type in db file's table of contents, show a floating window showing the record's info, same to other record name's floating peeking window.

### J-31 Update .gitignore

Status: `done`

Update .gitignore for JetBrains extension, excludes the unnecessary files


### J-32 Implement syntax highlighter for `.monitor` type file

Status: `done`

A file with suffix `.monitor` is a file type used in this extension. Such a file may look like this:

```
# test02.monitor

SYS = val
S = OK

$(SYS)1
val4
val5
val7
valabc
val2002
$(SYS)2003
```

A comment starts with `#`, a line that has `=` is a macro value assignment line. All other lines are either empty or channel (record) name lines. Each line can only contain only one record name. If there is more than one, mark it as an error. Implement the syntax highlighter for this type of file. the vscode version has already implemented this.


### J-33 Semantic check for .monitor file

Status: `done`

All macros in the channel names in `.monitor` files must be assigned. If a macros is not assigned, show an error for all macros. A empty assignment `SYS = ` is acceptable, it means the macro should be replaced by an empty string.

### J-34 runtime epics channel value in .monitor file

Status: `done`

Use the installed CA and PVA libs, create context, connect channel, get the basic inforamtion (e.g. choices for enum type record such as bo, mbbi), then monitor the channel. Then display the realtime channel value after the channel name in `.monitor` file. Align the values vertically. this functionality has been realized in vscode version.


### J-35 Runtime epics channel value in db file table of contents

Status: `done`

Similar to the epics channel monitoring in `.monitor` file, create an overlay and show the channel's runtime value (at a rate of 1Hz) in the `Value` column. this functionality is realized in vscode version. Keep the vertical alignment of the table.


### J-36 Put runtime record value

Status: `done`

When double click the runtime value, show a input box. User can type value in this box and submit it. 
