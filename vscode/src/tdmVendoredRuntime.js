const fs = require("fs");
const path = require("path");

const VENDORED_RUNTIME_MODE = "vendoredSource";
const EXTERNAL_RUNTIME_MODE = "externalBinary";

function getVendoredTdmRoot(extensionPath) {
  const vendorRoot = path.join(extensionPath, "vendor", "tdm");
  return fs.existsSync(path.join(vendorRoot, "dist", "mainProcess", "startMainProcess.js"))
    ? vendorRoot
    : undefined;
}

function getVendoredTdmLaunchInfo(extensionPath) {
  const vendorRoot = getVendoredTdmRoot(extensionPath);
  if (!vendorRoot) {
    return undefined;
  }

  return {
    rootPath: vendorRoot,
    command: process.execPath,
    args: [path.join(extensionPath, "src", "tdmVendoredRuntime", "startVendoredTdm.js"), "--vendor-root", vendorRoot],
    cwd: vendorRoot,
  };
}

module.exports = {
  EXTERNAL_RUNTIME_MODE,
  VENDORED_RUNTIME_MODE,
  getVendoredTdmLaunchInfo,
  getVendoredTdmRoot,
};
