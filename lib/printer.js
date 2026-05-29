const net = require("net");
const { printer } = require("../config");

function sendToPrinter(zplData, options = {}) {
  const ip = options.ip || printer.ip;
  const port = Number(options.port || printer.port);
  const timeoutMs = Number(options.timeoutMs || printer.timeoutMs || 10000);

  if (!zplData || typeof zplData !== "string") {
    return Promise.reject(new Error("No ZPL data was provided."));
  }

  return new Promise((resolve, reject) => {
    const client = new net.Socket();
    let settled = false;

    function finish(error, result) {
      if (settled) return;
      settled = true;
      client.destroy();
      if (error) reject(error);
      else resolve(result);
    }

    client.setTimeout(timeoutMs);

    client.connect(port, ip, () => {
      client.write(zplData, () => {
        client.end();
      });
    });

    client.on("timeout", () => {
      finish(new Error(`Connection to ${ip}:${port} timed out after ${timeoutMs} ms.`));
    });

    client.on("error", (error) => {
      finish(error);
    });

    client.on("close", () => {
      finish(null, { ip, port, bytesSent: Buffer.byteLength(zplData, "utf8") });
    });
  });
}

async function testConnection(options = {}) {
  const ip = options.ip || printer.ip;
  const port = Number(options.port || printer.port);
  const timeoutMs = Number(options.timeoutMs || printer.timeoutMs || 10000);

  return new Promise((resolve, reject) => {
    const client = new net.Socket();
    let settled = false;

    function finish(error, result) {
      if (settled) return;
      settled = true;
      client.destroy();
      if (error) reject(error);
      else resolve(result);
    }

    client.setTimeout(timeoutMs);
    client.connect(port, ip, () => {
      finish(null, { ok: true, ip, port });
    });
    client.on("timeout", () => finish(new Error(`Connection to ${ip}:${port} timed out after ${timeoutMs} ms.`)));
    client.on("error", finish);
  });
}

function buildFeedCommand(labelCount = 1) {
  const count = Math.max(1, Math.min(10, Number(labelCount) || 1));
  return `^XA^PQ${count}^XZ`;
}

function buildCalibrateCommand() {
  return "~JC";
}

module.exports = {
  sendToPrinter,
  testConnection,
  buildFeedCommand,
  buildCalibrateCommand,
};
