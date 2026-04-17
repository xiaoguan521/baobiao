const fs = require("fs");
const os = require("os");
const path = require("path");
const test = require("node:test");
const assert = require("node:assert/strict");
const { createApp, encodeFileId } = require("../src/server");

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
  assert.deepEqual(await response.json(), { ok: true, authRequired: true });
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
  const downloadRoot = fs.mkdtempSync(path.join(os.tmpdir(), "report-generate-"));
  const generatedFile = path.join(downloadRoot, "report.xlsx");
  fs.writeFileSync(generatedFile, "generated");
  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({ ok: true }),
    generateReport: async (payload) => ({
      month: payload.month,
      sheetName: payload.sheetName || null,
      sheetOnly: Boolean(payload.sheetOnly),
      outputPath: generatedFile,
      stats: { scannedRows: 10, classifiedRows: 9, unmatchedOutletRows: 1 }
    }),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot
  });
  const server = await startTestServer(app);
  t.after(async () => {
    await stopTestServer(server);
    fs.rmSync(downloadRoot, { recursive: true, force: true });
  });

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
  const payload = await response.json();
  assert.equal(payload.month, "2025-12");
  assert.equal(payload.file.name, "report.xlsx");
  assert.ok(payload.file.downloadUrl.endsWith(`/api/reports/download/${payload.file.id}`));
  assert.equal("outputPath" in payload, false);
});

test("download endpoint rejects files outside configured output root", async (t) => {
  const downloadRoot = fs.mkdtempSync(path.join(os.tmpdir(), "report-download-"));
  const insideFile = path.join(downloadRoot, "inside.txt");
  fs.writeFileSync(insideFile, "inside");

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
  });

  const { port } = server.address();
  const insideFileId = encodeFileId("inside.txt");
  const outsideResponse = await fetch(
    `http://127.0.0.1:${port}/api/reports/download/${encodeURIComponent(encodeFileId("../outside.txt"))}`,
    { headers: { authorization: "Bearer secret-token" } }
  );
  assert.equal(outsideResponse.status, 400);
  assert.deepEqual(await outsideResponse.json(), { error: "file is outside download root" });

  const insideResponse = await fetch(
    `http://127.0.0.1:${port}/api/reports/download/${encodeURIComponent(insideFileId)}`,
    { headers: { authorization: "Bearer secret-token" } }
  );
  assert.equal(insideResponse.status, 200);
  assert.equal(await insideResponse.text(), "inside");
});

test("root page is served as static html", async (t) => {
  const app = createApp({
    analyzeUnmatched: async () => ({}),
    generateReport: async () => ({}),
    reportSheets: ["总表"],
    configPath: "/tmp/report-rules.json",
    downloadRoot: os.tmpdir()
  });
  const server = await startTestServer(app);
  t.after(async () => stopTestServer(server));

  const { port } = server.address();
  const response = await fetch(`http://127.0.0.1:${port}/`);
  const html = await response.text();

  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") || "", /text\/html/);
  assert.match(html, /Oracle 报表服务台/);
});
