#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import os
from collections import deque
from datetime import datetime, timezone
from pathlib import Path
import shutil
import sys
from typing import Set


RUNTIME_PACKAGE_ROOTS = [
    "epics-tca",
    "express",
    "express-session",
    "node-fetch",
    "passport",
    "passport-ldapauth",
    "pidusage",
    "process",
    "selfsigned",
    "uuid",
    "ws",
    "xml2js",
]

VENDORED_ROOT_FILES = [
    "package.json",
    "package-lock.json",
    "README.md",
    "tsconfig.json",
    "webpack.config.js",
]

VENDORED_ROOT_DIRECTORIES = [
    "scripts",
    "dist/mainProcess",
    "dist/common",
    "dist/webpack",
]

PRUNED_TARGET_PATHS = [
    "src",
]


def parse_args() -> argparse.Namespace:
    script_path = Path(__file__).resolve()
    workspace_root = script_path.parent.parent
    default_source_root = (
        Path(os.environ["TDM_SOURCE_ROOT"]).expanduser().resolve()
        if os.environ.get("TDM_SOURCE_ROOT")
        else (workspace_root / "../../tdm").resolve()
    )

    parser = argparse.ArgumentParser(
        description="Copy a runtime-focused TDM snapshot into vscode/vendor/tdm.",
    )
    parser.add_argument(
        "source_root",
        nargs="?",
        default=str(default_source_root),
        help="Path to the source TDM checkout. Defaults to $TDM_SOURCE_ROOT or ../tdm.",
    )
    parser.add_argument(
        "--target-root",
        default=str(workspace_root / "vendor" / "tdm"),
        help="Path to the vendored TDM target folder. Defaults to vscode/vendor/tdm.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_root = Path(args.source_root).expanduser().resolve()
    target_root = Path(args.target_root).expanduser().resolve()

    assert_exists(source_root, "TDM source root")

    for relative_path in VENDORED_ROOT_FILES:
        copy_file(source_root, target_root, relative_path)
    for relative_path in VENDORED_ROOT_DIRECTORIES:
        copy_directory(source_root, target_root, relative_path)

    prune_target_paths(target_root)
    prune_webpack_hot_updates(target_root)
    sync_runtime_dependencies(source_root, target_root)
    write_manifest(source_root, target_root)

    print(f"Vendored TDM synced from {source_root} to {target_root}")
    return 0


def copy_file(source_root: Path, target_root: Path, relative_path: str) -> None:
    from_path = source_root / relative_path
    to_path = target_root / relative_path
    assert_exists(from_path, relative_path)
    to_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(from_path, to_path)


def copy_directory(source_root: Path, target_root: Path, relative_path: str) -> None:
    from_path = source_root / relative_path
    to_path = target_root / relative_path
    assert_exists(from_path, relative_path)
    shutil.rmtree(to_path, ignore_errors=True)
    to_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        from_path,
        to_path,
        ignore=ignore_names,
        symlinks=True,
        ignore_dangling_symlinks=True,
    )


def prune_target_paths(target_root: Path) -> None:
    for relative_path in PRUNED_TARGET_PATHS:
        shutil.rmtree(target_root / relative_path, ignore_errors=True)


def prune_webpack_hot_updates(target_root: Path) -> None:
    webpack_root = target_root / "dist" / "webpack"
    if not webpack_root.exists():
        return

    for entry in webpack_root.iterdir():
        if "hot-update" in entry.name or entry.name.endswith(".LICENSE.txt"):
            if entry.is_dir():
                shutil.rmtree(entry, ignore_errors=True)
            else:
                entry.unlink(missing_ok=True)


def sync_runtime_dependencies(source_root: Path, target_root: Path) -> None:
    lock_path = source_root / "package-lock.json"
    lock = json.loads(lock_path.read_text(encoding="utf8"))
    packages = lock.get("packages", {})
    visited: Set[str] = set()
    queue = deque(RUNTIME_PACKAGE_ROOTS)

    while queue:
        package_name = queue.popleft()
        package_key = f"node_modules/{package_name}"
        if package_key in visited or package_key not in packages:
            continue

        visited.add(package_key)
        dependencies = packages[package_key].get("dependencies", {})
        queue.extend(dependencies.keys())

    target_node_modules = target_root / "node_modules"
    shutil.rmtree(target_node_modules, ignore_errors=True)
    target_node_modules.mkdir(parents=True, exist_ok=True)

    for package_key in sorted(visited):
        from_path = source_root / package_key
        relative_node_module_path = package_key[len("node_modules/") :]
        to_path = target_node_modules / relative_node_module_path
        assert_exists(from_path, package_key)
        to_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.rmtree(to_path, ignore_errors=True)
        shutil.copytree(
            from_path,
            to_path,
            ignore=ignore_names,
            symlinks=True,
            ignore_dangling_symlinks=True,
        )


def write_manifest(source_root: Path, target_root: Path) -> None:
    manifest = {
        "sourceRoot": str(source_root),
        "syncedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "rootFiles": VENDORED_ROOT_FILES,
        "rootDirectories": VENDORED_ROOT_DIRECTORIES,
        "prunedTargetPaths": PRUNED_TARGET_PATHS,
        "runtimePackageRoots": RUNTIME_PACKAGE_ROOTS,
    }
    manifest_path = target_root / "vendor-manifest.json"
    manifest_path.write_text(f"{json.dumps(manifest, indent=2)}\n", encoding="utf8")


def ignore_names(_: str, names) -> Set[str]:
    return {name for name in names if name == ".DS_Store"}


def assert_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Cannot find {label} at {path}")


if __name__ == "__main__":
    sys.exit(main())
