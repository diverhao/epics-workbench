#!/usr/bin/env node

const path = require("path");
const Module = require("module");
const { createElectronShim } = require("./electronShim");
const { oracledbStub, ssh2Stub } = require("./optionalStubs");

const argv = process.argv.slice(2);
const vendorRoot = resolveVendorRoot(argv);
const forwardedArgs = stripVendorArgs(argv);
const startScriptPath = path.join(vendorRoot, "dist", "mainProcess", "startMainProcess.js");

installModuleShims(vendorRoot);
process.chdir(vendorRoot);
process.env.EPICS_WORKBENCH_TDM_VENDOR_ROOT = vendorRoot;
process.argv = [process.argv[0], ".", ...forwardedArgs];

require(startScriptPath);

function resolveVendorRoot(args) {
  const explicitIndex = args.indexOf("--vendor-root");
  if (explicitIndex >= 0 && args[explicitIndex + 1]) {
    return path.resolve(args[explicitIndex + 1]);
  }
  return path.resolve(__dirname, "../../vendor/tdm");
}

function stripVendorArgs(args) {
  const nextArgs = [];
  for (let index = 0; index < args.length; index += 1) {
    if (args[index] === "--vendor-root") {
      index += 1;
      continue;
    }
    nextArgs.push(args[index]);
  }
  return nextArgs;
}

function installModuleShims(vendorRoot) {
  const electronShim = createElectronShim(vendorRoot);
  const originalLoad = Module._load;

  Module._load = function patchedLoad(request, parent, isMain) {
    if (request === "electron") {
      return electronShim;
    }
    if (request === "oracledb") {
      return oracledbStub;
    }
    if (request === "ssh2") {
      return ssh2Stub;
    }
    return originalLoad.call(this, request, parent, isMain);
  };
}
