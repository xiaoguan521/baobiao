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

function normalizePublicBaseUrl(publicBaseUrl) {
  return publicBaseUrl ? String(publicBaseUrl).replace(/\/+$/, "") : "";
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

function resolvePublicBaseUrl(publicBaseUrl) {
  return normalizePublicBaseUrl(publicBaseUrl || process.env.REPORT_PUBLIC_BASE_URL || "");
}

function encodeFileId(relativePath) {
  return Buffer.from(relativePath, "utf8").toString("base64url");
}

function decodeFileId(fileId) {
  return Buffer.from(String(fileId || ""), "base64url").toString("utf8");
}

function createRequestBaseUrl(req, publicBaseUrl) {
  if (publicBaseUrl) return publicBaseUrl;
  return `${req.protocol}://${req.get("host")}`;
}

function createFileDescriptor(filePath, downloadRoot, req, publicBaseUrl) {
  const resolvedPath = normalizePathForCompare(filePath);
  if (!isPathInsideRoot(resolvedPath, downloadRoot)) {
    throw new Error("generated file is outside download root");
  }

  const relativePath = path.relative(downloadRoot, resolvedPath);
  const stats = fs.statSync(resolvedPath);
  const fileId = encodeFileId(relativePath);
  const baseUrl = createRequestBaseUrl(req, publicBaseUrl);

  return {
    id: fileId,
    name: path.basename(resolvedPath),
    sizeBytes: stats.size,
    downloadUrl: `${baseUrl}/api/reports/download/${encodeURIComponent(fileId)}`
  };
}

function parseGenerateRequest(body, reportSheets) {
  const payload = body || {};
  const month = String(payload.month || "").trim();
  const sheetName = payload.sheetName == null ? "" : String(payload.sheetName).trim();
  const sheetOnly = Boolean(payload.sheetOnly);

  if (!month) {
    throw new Error("month is required");
  }

  if (sheetOnly && !sheetName) {
    throw new Error("sheetOnly requires sheetName");
  }

  if (sheetName && !reportSheets.includes(sheetName)) {
    throw new Error(`sheetName must be one of: ${reportSheets.join(", ")}`);
  }

  return {
    month,
    sheetName: sheetName || undefined,
    sheetOnly
  };
}

function parseDebugRequest(query) {
  const month = String(query.month || "").trim();
  const rawLimit = query.limit == null ? 30 : Number(query.limit);

  if (!month) {
    throw new Error("month is required");
  }

  if (!Number.isInteger(rawLimit) || rawLimit <= 0 || rawLimit > 200) {
    throw new Error("limit must be an integer between 1 and 200");
  }

  return {
    month,
    limit: rawLimit
  };
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
  const publicBaseUrl = resolvePublicBaseUrl(options.publicBaseUrl);
  const publicDir = options.publicDir || path.join(process.cwd(), "public");
  const requireAuth = createAuthMiddleware(apiToken);

  const app = express();
  app.set("trust proxy", true);
  app.use(express.json({ limit: "1mb" }));

  app.get("/api/health", (_req, res) => {
    res.json({ ok: true, authRequired: Boolean(apiToken) });
  });

  app.get("/api/reports/sheets", (_req, res) => {
    res.json({
      sheets: services.reportSheets,
      configPath: services.configPath,
      authRequired: Boolean(apiToken)
    });
  });

  app.get("/api/reports/debug/unmatched", requireAuth, async (req, res) => {
    try {
      const { month, limit } = parseDebugRequest(req.query);
      const result = await services.analyzeUnmatched({
        month,
        limit,
        dbConfig: {
          user: process.env.ORACLE_USER || "damoxing",
          password: process.env.ORACLE_PASSWORD || "Damoxing123!",
          connectString: process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
        }
      });
      res.json(result);
    } catch (error) {
      res.status(400).json({ error: error.message });
    }
  });

  app.post("/api/reports/generate", requireAuth, async (req, res) => {
    try {
      const payload = parseGenerateRequest(req.body, services.reportSheets);
      const result = await services.generateReport({
        month: payload.month,
        sheetName: payload.sheetName,
        sheetOnly: payload.sheetOnly,
        templatePath: process.env.TEMPLATE_PATH || path.join(process.cwd(), "模板.xlsx"),
        outputDir: process.env.OUTPUT_DIR || downloadRoot,
        user: process.env.ORACLE_USER || "damoxing",
        password: process.env.ORACLE_PASSWORD || "Damoxing123!",
        dsn: process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
      });
      res.json({
        month: result.month,
        sheetName: result.sheetName,
        sheetOnly: result.sheetOnly,
        stats: result.stats,
        generatedAt: new Date().toISOString(),
        file: createFileDescriptor(result.outputPath, downloadRoot, req, publicBaseUrl)
      });
    } catch (error) {
      res.status(400).json({ error: error.message });
    }
  });

  app.get("/api/reports/download/:fileId", requireAuth, async (req, res) => {
    if (!req.params.fileId) {
      res.status(400).json({ error: "fileId is required" });
      return;
    }

    let relativePath;
    try {
      relativePath = decodeFileId(req.params.fileId);
    } catch (_error) {
      res.status(400).json({ error: "invalid fileId" });
      return;
    }

    const resolvedPath = normalizePathForCompare(path.join(downloadRoot, relativePath));
    if (!isPathInsideRoot(resolvedPath, downloadRoot)) {
      res.status(400).json({ error: "file is outside download root" });
      return;
    }

    if (!fs.existsSync(resolvedPath)) {
      res.status(404).json({ error: "file not found" });
      return;
    }

    res.download(resolvedPath);
  });

  app.use(express.static(publicDir, { index: "index.html" }));

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
  createFileDescriptor,
  createApp,
  createAuthMiddleware,
  decodeFileId,
  encodeFileId,
  isPathInsideRoot,
  normalizePublicBaseUrl,
  parseDebugRequest,
  parseGenerateRequest,
  resolveDownloadRoot,
  resolvePublicBaseUrl,
  startServer
};
