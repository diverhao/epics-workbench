#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const RUNTIME_PACKAGE_ROOTS = [
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
];

const VENDORED_ROOT_FILES = [
  "package.json",
  "package-lock.json",
  "README.md",
  "tsconfig.json",
  "webpack.config.js",
];

const VENDORED_ROOT_DIRECTORIES = [
  "scripts",
  "dist/mainProcess",
  "dist/common",
  "dist/webpack",
];

const PRUNED_TARGET_PATHS = [
  "src",
];

const workspaceRoot = path.resolve(__dirname, "..");
const sourceRoot = path.resolve(
  process.argv[2] || process.env.TDM_SOURCE_ROOT || path.join(workspaceRoot, "../../tdm"),
);
const targetRoot = path.join(workspaceRoot, "vendor", "tdm");

main();

function main() {
  assertExists(sourceRoot, "TDM source root");
  for (const relativePath of VENDORED_ROOT_FILES) {
    copyFile(relativePath);
  }
  for (const relativePath of VENDORED_ROOT_DIRECTORIES) {
    copyDirectory(relativePath);
  }
  pruneTargetPaths();
  pruneWebpackHotUpdates();
  syncRuntimeDependencies();
  writeManifest();
  console.log(`Vendored TDM synced from ${sourceRoot} to ${targetRoot}`);
}

function copyFile(relativePath) {
  const from = path.join(sourceRoot, relativePath);
  const to = path.join(targetRoot, relativePath);
  assertExists(from, relativePath);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
}

function copyDirectory(relativePath) {
  const from = path.join(sourceRoot, relativePath);
  const to = path.join(targetRoot, relativePath);
  assertExists(from, relativePath);
  fs.rmSync(to, { recursive: true, force: true });
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.cpSync(from, to, {
    recursive: true,
    force: true,
    filter: (item) => path.basename(item) !== ".DS_Store",
  });
}

function pruneWebpackHotUpdates() {
  const webpackRoot = path.join(targetRoot, "dist", "webpack");
  if (!fs.existsSync(webpackRoot)) {
    return;
  }
  for (const entry of fs.readdirSync(webpackRoot)) {
    if (entry.includes("hot-update") || entry.endsWith(".LICENSE.txt")) {
      fs.rmSync(path.join(webpackRoot, entry), { force: true });
    }
  }
}

function pruneTargetPaths() {
  for (const relativePath of PRUNED_TARGET_PATHS) {
    fs.rmSync(path.join(targetRoot, relativePath), { recursive: true, force: true });
  }
}

function syncRuntimeDependencies() {
  const lockPath = path.join(sourceRoot, "package-lock.json");
  const lock = JSON.parse(fs.readFileSync(lockPath, "utf8"));
  const packages = lock.packages || {};
  const visited = new Set();

  const visit = (packageName) => {
    const key = `node_modules/${packageName}`;
    if (visited.has(key) || !packages[key]) {
      return;
    }
    visited.add(key);
    const dependencies = packages[key].dependencies || {};
    for (const dependencyName of Object.keys(dependencies)) {
      visit(dependencyName);
    }
  };

  for (const packageName of RUNTIME_PACKAGE_ROOTS) {
    visit(packageName);
  }

  const targetNodeModules = path.join(targetRoot, "node_modules");
  fs.rmSync(targetNodeModules, { recursive: true, force: true });
  fs.mkdirSync(targetNodeModules, { recursive: true });

  for (const key of [...visited].sort()) {
    const relativeNodeModulePath = key.replace(/^node_modules\//, "");
    const from = path.join(sourceRoot, key);
    const to = path.join(targetNodeModules, relativeNodeModulePath);
    assertExists(from, key);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.cpSync(from, to, { recursive: true, force: true });
  }
}

function writeManifest() {
  const manifestPath = path.join(targetRoot, "vendor-manifest.json");
  const manifest = {
    sourceRoot,
    syncedAt: new Date().toISOString(),
    rootFiles: VENDORED_ROOT_FILES,
    rootDirectories: VENDORED_ROOT_DIRECTORIES,
    prunedTargetPaths: PRUNED_TARGET_PATHS,
    runtimePackageRoots: RUNTIME_PACKAGE_ROOTS,
  };
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
}

function assertExists(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Cannot find ${label} at ${filePath}`);
  }
}
