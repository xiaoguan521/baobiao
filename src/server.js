const express = require("express");
const path = require("path");
const { analyzeUnmatched, generateReport, REPORT_SHEETS, CONFIG_PATH } = require("./report-engine");

const app = express();
app.use(express.json({ limit: "1mb" }));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/reports/sheets", (_req, res) => {
  res.json({ sheets: REPORT_SHEETS, configPath: CONFIG_PATH });
});

app.get("/api/reports/debug/unmatched", async (req, res) => {
  try {
    const month = req.query.month;
    if (!month) {
      res.status(400).json({ error: "month is required" });
      return;
    }
    const result = await analyzeUnmatched({
      month,
      limit: Number(req.query.limit || 30),
      dbConfig: {
        user: req.query.user || process.env.ORACLE_USER || "damoxing",
        password: req.query.password || process.env.ORACLE_PASSWORD || "Damoxing123!",
        connectString: req.query.dsn || process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
      }
    });
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post("/api/reports/generate", async (req, res) => {
  try {
    const body = req.body || {};
    const result = await generateReport({
      month: body.month,
      sheetName: body.sheetName,
      sheetOnly: Boolean(body.sheetOnly),
      templatePath: body.templatePath || process.env.TEMPLATE_PATH || path.join(process.cwd(), "模板.xlsx"),
      outputPath: body.outputPath,
      outputDir: body.outputDir || process.env.OUTPUT_DIR,
      user: body.user || process.env.ORACLE_USER || "damoxing",
      password: body.password || process.env.ORACLE_PASSWORD || "Damoxing123!",
      dsn: body.dsn || process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
    });
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.get("/api/reports/download", async (req, res) => {
  const filePath = req.query.path;
  if (!filePath) {
    res.status(400).json({ error: "path is required" });
    return;
  }
  res.download(filePath);
});

const port = Number(process.env.PORT || 3000);
app.listen(port, () => {
  console.log(`report service listening on ${port}`);
});
