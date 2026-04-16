const express = require("express");
const fs = require("fs");
const path = require("path");
const {
  analyzeUnmatched,
  generateReport,
  REPORT_SHEETS,
  CONFIG_PATH
} = require("./report-engine");

function normalizePathForCompare(filePath) {
  return path.resolve(filePath);
}

function isPathInsideRoot(filePath, rootPath) {
  const resolvedFile = normalizePathForCompare(filePath);
  const resolvedRoot = normalizePathForCompare(rootPath);
  return resolvedFile === resolvedRoot || resolvedFile.startsWith(`${resolvedRoot}${path.sep}`);
}

function createAuthMiddleware(apiToken) {
  return (req, res, next) => {
    if (!apiToken) {
      next();
      return;
    }

    const authHeader = req.get("authorization") || "";
    const bearerToken = authHeader.startsWith("Bearer ") ? authHeader.slice("Bearer ".length).trim() : "";
    const headerToken = req.get("x-api-token") || "";

    if (bearerToken === apiToken || headerToken === apiToken) {
      next();
      return;
    }

    res.status(401).json({ error: "unauthorized" });
  };
}

function resolveDownloadRoot(downloadRoot) {
  return normalizePathForCompare(downloadRoot || process.env.DOWNLOAD_ROOT || process.env.OUTPUT_DIR || path.join(process.cwd(), "generated"));
}

function createApp(options = {}) {
  const services = {
    analyzeUnmatched: options.analyzeUnmatched || analyzeUnmatched,
    generateReport: options.generateReport || generateReport,
    reportSheets: options.reportSheets || REPORT_SHEETS,
    configPath: options.configPath || CONFIG_PATH
  };
  const apiToken = options.apiToken ?? process.env.REPORT_API_TOKEN ?? "";
  const downloadRoot = resolveDownloadRoot(options.downloadRoot);
  const requireAuth = createAuthMiddleware(apiToken);

  const app = express();
  app.use(express.json({ limit: "1mb" }));

  app.get("/api/health", (_req, res) => {
    res.json({ ok: true });
  });

  app.get("/api/reports/sheets", (_req, res) => {
    res.json({ sheets: services.reportSheets, configPath: services.configPath });
  });

  app.get("/api/reports/debug/unmatched", requireAuth, async (req, res) => {
    try {
      const month = req.query.month;
      if (!month) {
        res.status(400).json({ error: "month is required" });
        return;
      }
      const result = await services.analyzeUnmatched({
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

  app.post("/api/reports/generate", requireAuth, async (req, res) => {
    try {
      const body = req.body || {};
      const result = await services.generateReport({
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

  app.get("/api/reports/download", requireAuth, async (req, res) => {
    const requestedPath = req.query.path;
    if (!requestedPath) {
      res.status(400).json({ error: "path is required" });
      return;
    }

    const resolvedPath = normalizePathForCompare(String(requestedPath));
    if (!isPathInsideRoot(resolvedPath, downloadRoot)) {
      res.status(400).json({ error: "path is outside download root" });
      return;
    }

    if (!fs.existsSync(resolvedPath)) {
      res.status(404).json({ error: "file not found" });
      return;
    }

    res.download(resolvedPath);
  });

  return app;
}

function startServer(options = {}) {
  const app = createApp(options);
  const port = Number(options.port || process.env.PORT || 3000);
  return app.listen(port, () => {
    console.log(`report service listening on ${port}`);
  });
}

if (require.main === module) {
  startServer();
}

module.exports = {
  createApp,
  createAuthMiddleware,
  isPathInsideRoot,
  resolveDownloadRoot,
  startServer
};
