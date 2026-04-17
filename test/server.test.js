const fs = require("fs");
const os = require("os");
const path = require("path");
const test = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");
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

test("preview endpoint returns worksheet data for generated workbook", async (t) => {
  const downloadRoot = fs.mkdtempSync(path.join(os.tmpdir(), "report-preview-"));
  const generatedFile = path.join(downloadRoot, "report.xlsx");
  const workbook = new ExcelJS.Workbook();
  const regionSheet = workbook.addWorksheet("市本级");
  regionSheet.getCell("B2").value = "网点";
  regionSheet.getCell("C2").value = "归集";
  regionSheet.getCell("B3").value = "工行锦山支行";
  regionSheet.getCell("C3").value = 12;
  workbook.addWorksheet("排名").getCell("A1").value = "排名";
  await workbook.xlsx.writeFile(generatedFile);

  const app = createApp({
    apiToken: "secret-token",
    analyzeUnmatched: async () => ({}),
    generateReport: async () => ({}),
    reportSheets: ["市本级", "排名"],
    configPath: "/tmp/report-rules.json",
    downloadRoot
  });
  const server = await startTestServer(app);
  t.after(async () => {
    await stopTestServer(server);
    fs.rmSync(downloadRoot, { recursive: true, force: true });
  });

  const { port } = server.address();
  const response = await fetch(
    `http://127.0.0.1:${port}/api/reports/preview/${encodeURIComponent(encodeFileId("report.xlsx"))}?sheet=${encodeURIComponent("市本级")}`,
    { headers: { authorization: "Bearer secret-token" } }
  );

  assert.equal(response.status, 200);
  const payload = await response.json();
  assert.deepEqual(payload.sheets, ["市本级", "排名"]);
  assert.equal(payload.sheetName, "市本级");
  assert.deepEqual(payload.preview.columns.map((item) => item.label), ["B", "C"]);
  assert.equal(payload.preview.rows[0].rowNumber, 2);
  assert.deepEqual(
    payload.preview.rows[1].cells.map((item) => item.value),
    ["工行锦山支行", "12"]
  );
});

test("preview endpoint keeps merged cells for frontend rendering", async (t) => {
  const downloadRoot = fs.mkdtempSync(path.join(os.tmpdir(), "report-preview-merge-"));
  const generatedFile = path.join(downloadRoot, "merge.xlsx");
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("总表");
  worksheet.getCell("B2").value = "丹东公积金报表";
  worksheet.mergeCells("B2:D2");
  worksheet.getCell("B3").value = "地区";
  worksheet.getCell("C3").value = "归集";
  await workbook.xlsx.writeFile(generatedFile);

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
  const response = await fetch(
    `http://127.0.0.1:${port}/api/reports/preview/${encodeURIComponent(encodeFileId("merge.xlsx"))}`,
    { headers: { authorization: "Bearer secret-token" } }
  );

  assert.equal(response.status, 200);
  const payload = await response.json();
  assert.equal(payload.preview.rows[0].cells[0].value, "丹东公积金报表");
  assert.equal(payload.preview.rows[0].cells[0].colSpan, 3);
  assert.equal(payload.preview.rows[0].cells[0].rowSpan, 1);
  assert.equal(payload.preview.rows[1].cells[0].value, "地区");
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
