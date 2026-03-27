const EventEmitter = require("events");

const oracledbStub = {
  BIND_OUT: 0,
  CURSOR: 0,
  async getConnection() {
    throw new Error("Oracle archive support is unavailable in the vendored VS Code TDM runtime.");
  },
};

class SshClientStub extends EventEmitter {
  connect() {
    queueMicrotask(() => {
      this.emit("error", new Error("SSH support is unavailable in the vendored VS Code TDM runtime."));
    });
  }

  end() {}
  destroy() {}
}

class ClientChannelStub extends EventEmitter {}

module.exports = {
  oracledbStub,
  ssh2Stub: {
    Client: SshClientStub,
    ClientChannel: ClientChannelStub,
  },
};
