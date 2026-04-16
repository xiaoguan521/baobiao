const fs = require("fs");
const os = require("os");
const path = require("path");
const test = require("node:test");
const assert = require("node:assert/strict");
const { createApp } = require("../src/server");

async function startTestServer(app) {
  return new Promise((resolve) => {
    const server = app.listen(0, "127.0.0.1", () => resolve(server));
  });
}

async function stopTestServer(server) {
  await new Promise((resolve, reject) => {
    server.close((error) => (error ? reject(error) : resolve()));
  });
}

test("health endpoint stays open without auth", async (t) => {
  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({}),
    generateReport: async () => ({ ok: true }),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot: os.tmpdir()
  });
  const server = await startTestServer(app);
  t.after(async () => stopTestServer(server));

  const { port } = server.address();
  const response = await fetch(`http://127.0.0.1:${port}/api/health`);
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), { ok: true });
});

test("protected endpoints require token when REPORT_API_TOKEN is configured", async (t) => {
  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({ ok: true }),
    generateReport: async () => ({ ok: true }),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot: os.tmpdir()
  });
  const server = await startTestServer(app);
  t.after(async () => stopTestServer(server));

  const { port } = server.address();
  const response = await fetch(`http://127.0.0.1:${port}/api/reports/generate`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ month: "2025-12" })
  });

  assert.equal(response.status, 401);
  assert.deepEqual(await response.json(), { error: "unauthorized" });
});

test("protected endpoints accept bearer token", async (t) => {
  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({ ok: true }),
    generateReport: async (payload) => ({ ok: true, month: payload.month }),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot: os.tmpdir()
  });
  const server = await startTestServer(app);
  t.after(async () => stopTestServer(server));

  const { port } = server.address();
  const response = await fetch(`http://127.0.0.1:${port}/api/reports/generate`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: "Bearer secret-token"
    },
    body: JSON.stringify({ month: "2025-12" })
  });

  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), { ok: true, month: "2025-12" });
});

test("download endpoint rejects files outside configured output root", async (t) => {
  const downloadRoot = fs.mkdtempSync(path.join(os.tmpdir(), "report-download-"));
  const insideFile = path.join(downloadRoot, "inside.txt");
  const outsideFile = path.join(os.tmpdir(), `outside-${Date.now()}.txt`);
  fs.writeFileSync(insideFile, "inside");
  fs.writeFileSync(outsideFile, "outside");

  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({}),
    generateReport: async () => ({}),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot
  });
  const server = await startTestServer(app);
  t.after(async () => {
    await stopTestServer(server);
    fs.rmSync(downloadRoot, { recursive: true, force: true });
    fs.rmSync(outsideFile, { force: true });
  });

  const { port } = server.address();
  const outsideResponse = await fetch(
    `http://127.0.0.1:${port}/api/reports/download?path=${encodeURIComponent(outsideFile)}`,
    { headers: { authorization: "Bearer secret-token" } }
  );
  assert.equal(outsideResponse.status, 400);
  assert.deepEqual(await outsideResponse.json(), { error: "path is outside download root" });

  const insideResponse = await fetch(
    `http://127.0.0.1:${port}/api/reports/download?path=${encodeURIComponent(insideFile)}`,
    { headers: { authorization: "Bearer secret-token" } }
  );
  assert.equal(insideResponse.status, 200);
  assert.equal(await insideResponse.text(), "inside");
});
