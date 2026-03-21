# JetBrains Roadmap

This roadmap is for the JetBrains version of `epics-workbench`.

Scope assumptions:

- one plugin project for both `IntelliJ IDEA` and `PyCharm`
- common IntelliJ Platform APIs only
- no Python-specific APIs unless explicitly needed later
- runtime support should remain aligned with the VS Code extension, but the UI does not need to be a 1:1 clone

## Milestone 1: Project Skeleton and File Types

Goal:
- make EPICS files first-class citizens in the IDE

Scope:
- register EPICS file types and icons
- recognize `.db`, `.template`, `.substitutions`, `.sub`, `.subs`, `.cmd`, `.iocsh`, `.dbd`, `.proto`, `.st`, `.monitor`
- add basic language display names
- keep the current tool window and scaffold healthy

Exit criteria:
- files open with EPICS-specific file types
- plugin builds with `./gradlew buildPlugin`
- sandbox IDE launches with `./gradlew runIde`

## Milestone 2: Syntax and Structure

Goal:
- provide useful editing for EPICS text formats

Scope:
- syntax highlighting for database, substitutions, startup, dbd, proto, sequencer, and monitor files
- comment/string/keyword/token recognition
- basic brace matching and editor behavior where applicable

Exit criteria:
- the common EPICS files are readable and tokenized correctly
- no obvious regressions in editor performance

## Milestone 3: Navigation and Documentation

Goal:
- make EPICS projects navigable

Scope:
- go to declaration for records, startup-loaded files, Makefile DB references, and linked records
- hover previews for records, substitutions templates, and Makefile DB entries
- project-aware file resolution similar to the VS Code extension

Exit criteria:
- major navigation paths work in a representative test project

## Milestone 4: Diagnostics and Formatting

Goal:
- provide editing assistance beyond highlighting

Scope:
- inspections for common EPICS mistakes
- quick fixes where the value is high and the implementation is safe
- formatting for DB and substitutions files

Exit criteria:
- at least the most common static validation paths are covered

## Milestone 5: IOC and Runtime UI

Goal:
- make JetBrains useful for runtime EPICS work, not just static editing

Scope:
- runtime tool window
- project runtime configuration
- monitor list and basic runtime status
- prepare the architecture for CA/PVA integration

Exit criteria:
- runtime panel exists and can host future monitor features cleanly

## Milestone 6: CA/PVA Integration

Goal:
- support live EPICS interaction

Scope:
- choose the runtime integration path:
  - JVM-native CA/PVA libraries
  - or a sidecar service
- monitor channel values
- support put operations for scalar writable values
- handle enum/choice presentation

Exit criteria:
- a user can monitor and put a representative set of PVs from the IDE

## Milestone 7: Editor Runtime Integration

Goal:
- bring runtime information back into the editor

Scope:
- inline runtime value presentation
- click or action-driven put flow
- record-centric runtime actions from hovers, gutters, or popups

Exit criteria:
- runtime data is easy to inspect from the editor without overwhelming the UI

## Architectural Notes

- Prefer a shared core for EPICS parsing/resolution logic where feasible.
- Keep JetBrains-specific code focused on PSI, editor integration, actions, settings, and tool windows.
- Do not try to copy VS Code decoration behavior exactly if JetBrains has a better native interaction model.
