// stop-server.js
const { exec } = require("child_process");

const PORT = process.argv[2] || "3000";

function run(command) {
  return new Promise((resolve) => {
    exec(command, (error, stdout, stderr) => {
      resolve({
        error,
        stdout,
        stderr,
      });
    });
  });
}

async function stopServer() {
  console.log(`Looking for server using port ${PORT}...`);

  const netstatResult = await run(`netstat -ano | findstr :${PORT}`);
  const output = netstatResult.stdout.trim();

  if (!output) {
    console.log(`No server found using port ${PORT}.`);
    return;
  }

  const lines = output.split(/\r?\n/);
  const pids = new Set();

  for (const line of lines) {
    if (!line.includes("LISTENING")) {
      continue;
    }

    const parts = line.trim().split(/\s+/);
    const pid = parts[parts.length - 1];

    if (pid && /^\d+$/.test(pid)) {
      pids.add(pid);
    }
  }

  if (pids.size === 0) {
    console.log(`No listening server found on port ${PORT}.`);
    return;
  }

  for (const pid of pids) {
    console.log(`Stopping server PID ${pid} on port ${PORT}...`);

    const killResult = await run(`taskkill /PID ${pid} /F`);

    if (killResult.error) {
      console.log(`Failed to stop PID ${pid}.`);
      console.log(killResult.stderr || killResult.error.message);
    } else {
      console.log(killResult.stdout.trim());
    }
  }

  console.log(`Done. Port ${PORT} is free.`);
}

stopServer().catch((error) => {
  console.error("Failed to stop server:");
  console.error(error.message);
});