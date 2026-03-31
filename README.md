# EPICS Workbench for VS Code and JetBrains

EPICS Workbench adds useful EPICS tools to your editor. It helps you read files, move around a project, work with runtime data, and use table-based views without leaving the IDE.

This README applies to both VS Code and JetBrains unless noted otherwise.

- [Quick Start](#quick-start)
- [Static Tools](#static-tools)
  - [Syntax, Errors, and Formatting](#syntax-errors-and-formatting)
  - [File Preview and Navigation](#file-preview-and-navigation)
  - [Auto-Complete](#auto-complete)
  - [Record Preview and Navigation](#record-preview-and-navigation)
  - [Database Table of Contents](#database-table-of-contents)
  - [Environment Variable Resolution](#environment-variable-resolution)
- [Widgets](#widgets)
  - [Probe](#probe)
  - [PV List](#pv-list)
  - [PV Monitor](#pv-monitor)
  - [Spreadsheet](#spreadsheet)
- [IOC Runtime Tools](#ioc-runtime-tools)
  - [Configuration](#configuration)
  - [IOC Control](#ioc-control)
  - [Run IOC Shell Commands](#run-ioc-shell-commands)
  - [Live Channel Data](#live-channel-data)
- [TDM Integration](#tdm-integration)

## Quick Start

Open your EPICS application folder in the editor, for example:

```bash
code /home/abc/appTest/
```

## Static Tools

Static tools work from the files in your workspace. You do not need a running IOC to use them.

- Syntax highlighting, error checks, and formatting
- Quick file and record navigation
- Auto-complete for common EPICS files
- Database table of contents
- Environment variable lookup

### Syntax, Errors, and Formatting

EPICS Workbench highlights common EPICS files and helps you catch mistakes while you edit.

<img src="doc-figures/vscode-syntax-highlight.png" width="30%" />

It can also show problems as you type:

<img src="doc-figures/vscode-db-correction.gif" width="55%" />

Formatting is available for supported EPICS files:

<img src="doc-figures/vscode-db-format.gif" width="55%" />

### File Preview and Navigation

You can preview files in a side view and jump straight to the source.

<img src="doc-figures/vscode-file-nav.gif" width="55%" />

### Auto-Complete

Auto-complete helps you type less and choose valid names and values more easily.

<img src="doc-figures/vscode-auto-complete.gif" width="55%" />

It also works in files such as `st.cmd`.

<img src="doc-figures/vscode-cmd-auto-completion.gif" width="55%" />

### Record Preview and Navigation

You can preview records and jump to related definitions quickly.

<img src="doc-figures/vscode-record-nav.gif" width="55%" />

### Database Table of Contents

The database table of contents gives you a quick overview of records, macros, fields, and runtime values.

<img src="doc-figures/vscode-db-toc.gif" width="55%" />

### Environment Variable Resolution

Variables from `envPaths` and `RELEASE` can be resolved in the editor.

<img src="doc-figures/vscode-macro-resolution.gif" width="55%" />

## Widgets

EPICS Workbench includes four main widgets.

### Probe

`Probe` shows a record's live value and its fields.

<img src="doc-figures/vscode-widget-probe.gif" width="55%" />

### PV List

`PV List` shows values for multiple channels. You can build the list yourself or from a running IOC.

<img src="doc-figures/vscode-widget-pvlist.gif" width="55%" />

### PV Monitor

`PV Monitor` lets you watch changing values in one place.

<img src="doc-figures/vscode-widget-pvmonitor.gif" width="55%" />

### Spreadsheet

Note: this widget is available in the VS Code extension.

The `Spreadsheet` widget lets you view and edit EPICS data in a table. You can open Excel files, work with EPICS databases in spreadsheet form, save changes, and move between spreadsheet and database formats.

<img src="doc-figures/vscode-spreadsheet-widget.gif" width="55%" />

You can also open a database file directly in spreadsheet view:

<img src="doc-figures/vscode-spreadsheet-widget-2.gif" width="55%" />

Excel import and export are available as well:

<img src="doc-figures/vscode-export-to-excel.gif" width="55%" />

## IOC Runtime Tools

Runtime tools work with a live IOC.

### Configuration

Open `EPICS Runtime Configuration` to set the runtime options.

<img src="./doc-figures/vscode-config.png" width="50%" />

The settings page keeps the main options in one place:

<img src="./doc-figures/vscode-config-page.png" width="50%" />

### IOC Control

You can start, stop, and manage IOCs from the context menu.

<img src="./doc-figures/vscode-start-stop-ioc.jpg" width="50%" />

### Run IOC Shell Commands

You can run IOC shell commands in a widget and send the output to a buffer.

<img src="./doc-figures/vscode-widget-iocsh-commands.gif" width="50%" />

### Live Channel Data

Runtime widgets make it easier to browse and choose record names without typing everything by hand.

<img src="./doc-figures/vscode-widget-ioc-runtime-pvlist-probe.gif" width="50%" />

## TDM Integration

Note: this is currently available only in the VS Code extension and is still under development.

[TDM](https://github.com/diverhao/tdm) integration lets you open operator displays inside VS Code so they stay close to your source files and IOC tools.

To use it, add `Web Server` to the first TDM profile and choose `No Authentication`.

Open a `.tdl` file in VS Code to view and edit the display.

<img src="./doc-figures/vscode-tdm-01.gif" width="50%" />
