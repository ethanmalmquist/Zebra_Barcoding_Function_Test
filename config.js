module.exports = {
  printer: {
    ip: process.env.PRINTER_IP || "192.168.1.126",
    port: Number(process.env.PRINTER_PORT || 9100),
    timeoutMs: Number(process.env.PRINTER_TIMEOUT_MS || 10000),
  },
};
