const fs = require("fs");
const path = require("path");
const EventEmitter = require("events");

function createNoOpFunction(returnValue) {
  return () => returnValue;
}

class ShimWebContents extends EventEmitter {
  constructor() {
    super();
    this.id = Date.now();
  }

  send() {}
  loadURL() {
    return Promise.resolve();
  }
  loadFile() {
    return Promise.resolve();
  }
  setWindowOpenHandler() {
    return { action: "deny" };
  }
  setZoomFactor() {}
  print() {}
  printToPDF() {
    return Promise.resolve(Buffer.alloc(0));
  }
  capturePage() {
    return Promise.resolve({
      toPNG() {
        return Buffer.alloc(0);
      },
    });
  }
  getURL() {
    return "";
  }
  isDestroyed() {
    return false;
  }
}

class ShimBrowserWindow extends EventEmitter {
  static windows = new Set();

  static getAllWindows() {
    return [...this.windows];
  }

  static getFocusedWindow() {
    return [...this.windows].find((window) => window._focused) || undefined;
  }

  static fromWebContents() {
    return undefined;
  }

  constructor(options = {}) {
    super();
    this.options = options;
    this._bounds = {
      x: 0,
      y: 0,
      width: options.width || 800,
      height: options.height || 600,
    };
    this._focused = false;
    this._destroyed = false;
    this.webContents = new ShimWebContents();
    ShimBrowserWindow.windows.add(this);
  }

  loadURL() {
    return Promise.resolve();
  }

  loadFile() {
    return Promise.resolve();
  }

  show() {
    this._focused = true;
  }

  hide() {
    this._focused = false;
  }

  focus() {
    this._focused = true;
  }

  blur() {
    this._focused = false;
  }

  close() {
    this.emit("close", { preventDefault() {} });
    this._destroyed = true;
    ShimBrowserWindow.windows.delete(this);
    this.emit("closed");
  }

  destroy() {
    this.close();
  }

  isDestroyed() {
    return this._destroyed;
  }

  getBounds() {
    return { ...this._bounds };
  }

  setBounds(bounds) {
    this._bounds = { ...this._bounds, ...bounds };
  }

  setSize(width, height) {
    this._bounds.width = width;
    this._bounds.height = height;
  }

  getSize() {
    return [this._bounds.width, this._bounds.height];
  }

  setPosition(x, y) {
    this._bounds.x = x;
    this._bounds.y = y;
  }

  getPosition() {
    return [this._bounds.x, this._bounds.y];
  }

  setTitle() {}
  setMenuBarVisibility() {}
  setAutoHideMenuBar() {}
  setAlwaysOnTop() {}
  setResizable() {}
  center() {}
  maximize() {}
  unmaximize() {}
  isMaximized() {
    return false;
  }
  showInactive() {}
  moveTop() {}
  minimize() {}
  restore() {}
}

class ShimBrowserView {
  constructor() {
    this.webContents = new ShimWebContents();
  }
}

class ShimMenu {
  static buildFromTemplate() {
    return new ShimMenu();
  }

  static setApplicationMenu() {}

  append() {}
  popup() {}
}

class ShimMenuItem {
  constructor(options = {}) {
    Object.assign(this, options);
  }
}

function createClipboardShim() {
  let text = "";
  return {
    readText: () => text,
    writeText: (nextText) => {
      text = nextText;
    },
  };
}

function createDialogShim() {
  return {
    showOpenDialogSync: createNoOpFunction(undefined),
    showSaveDialogSync: createNoOpFunction(undefined),
    showMessageBoxSync: createNoOpFunction(0),
    showMessageBox: async () => ({ response: 0 }),
    showOpenDialog: async () => ({ canceled: true, filePaths: [] }),
    showSaveDialog: async () => ({ canceled: true, filePath: undefined }),
  };
}

function createAppShim(vendorRoot) {
  const emitter = new EventEmitter();
  const packageJsonPath = path.join(vendorRoot, "package.json");
  const packageJson = fs.existsSync(packageJsonPath)
    ? JSON.parse(fs.readFileSync(packageJsonPath, "utf8"))
    : {};

  const baseApp = {
    commandLine: {
      appendSwitch() {},
    },
    requestSingleInstanceLock() {
      return true;
    },
    releaseSingleInstanceLock() {},
    whenReady() {
      return Promise.resolve();
    },
    isReady() {
      return true;
    },
    quit() {
      process.exit(0);
    },
    exit(code = 0) {
      process.exit(code);
    },
    relaunch() {},
    focus() {},
    getVersion() {
      return packageJson.version || "0.0.0";
    },
    getName() {
      return packageJson.productName || packageJson.name || "TDM";
    },
    getAppPath() {
      return vendorRoot;
    },
    getPath(name) {
      if (name === "home") {
        return process.env.HOME || process.cwd();
      }
      return process.cwd();
    },
    setPath() {},
    dock: {
      setMenu() {},
      hide() {},
      show() {},
    },
    on: emitter.on.bind(emitter),
    once: emitter.once.bind(emitter),
    emit: emitter.emit.bind(emitter),
  };

  return new Proxy(baseApp, {
    get(target, property) {
      if (property in target) {
        return target[property];
      }
      return createNoOpFunction(undefined);
    },
  });
}

function createElectronShim(vendorRoot) {
  return {
    app: createAppShim(vendorRoot),
    BrowserWindow: ShimBrowserWindow,
    BrowserView: ShimBrowserView,
    Menu: ShimMenu,
    MenuItem: ShimMenuItem,
    dialog: createDialogShim(),
    clipboard: createClipboardShim(),
    screen: {
      getCursorScreenPoint: () => ({ x: 0, y: 0 }),
      getPrimaryDisplay: () => ({ workAreaSize: { width: 1920, height: 1080 } }),
      getAllDisplays: () => [],
    },
    desktopCapturer: {
      getSources: async () => [],
    },
    nativeImage: {
      createFromPath() {
        return {
          isEmpty: () => true,
        };
      },
    },
    shell: {
      openExternal: async () => {},
    },
    ipcMain: new EventEmitter(),
    ipcRenderer: new EventEmitter(),
    contextBridge: {
      exposeInMainWorld() {},
    },
    webUtils: {
      getPathForFile(file) {
        return file?.path || "";
      },
    },
    systemPreferences: {
      getUserDefault: createNoOpFunction(""),
      setUserDefault() {},
    },
  };
}

module.exports = {
  createElectronShim,
};
