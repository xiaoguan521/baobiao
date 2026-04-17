const express = require("express");
const ExcelJS = require("exceljs");
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

function normalizeSheetName(sheetName) {
  return String(sheetName == null ? "" : sheetName).trim();
}

function resolveSheetName(sheetName, availableSheetNames) {
  const exactName = String(sheetName == null ? "" : sheetName);
  if (!exactName) return undefined;
  const exactMatch = availableSheetNames.find((item) => item === exactName);
  if (exactMatch) return exactMatch;
  const normalizedName = normalizeSheetName(exactName);
  return availableSheetNames.find((item) => normalizeSheetName(item) === normalizedName);
}

function columnNumberToLabel(columnNumber) {
  let current = Number(columnNumber || 0);
  let label = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label || "A";
}

function columnLabelToNumber(columnLabel) {
  return String(columnLabel || "")
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
}

function parseCellReference(cellReference) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(cellReference || "").trim());
  if (!match) return null;
  return {
    column: columnLabelToNumber(match[1]),
    row: Number(match[2])
  };
}

function parseMergeRange(rangeReference) {
  const [startRef, endRef] = String(rangeReference || "").split(":");
  const start = parseCellReference(startRef);
  const end = parseCellReference(endRef || startRef);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column)
  };
}

function getCellDisplayValue(cell) {
  if (!cell || cell.value == null) return "";
  const value = cell.value;

  if (value instanceof Date) {
    return cell.text || value.toISOString();
  }

  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text || "").join("");
    }
    if (value.formula || value.sharedFormula) {
      if (value.result != null) return String(value.result);
      return value.formula ? `=${value.formula}` : "";
    }
    if (value.hyperlink) return String(value.text || value.hyperlink);
    if (value.text != null) return String(value.text);
    if (value.error) return String(value.error);
  }

  if (cell.text) return cell.text;
  return String(value);
}

function getWorksheetMergeRegions(worksheet) {
  const mergeRanges = worksheet.model?.merges || [];
  return mergeRanges
    .map((range) => parseMergeRange(range))
    .filter(Boolean);
}

function findWorksheetBounds(worksheet) {
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxColumn = 0;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const displayValue = getCellDisplayValue(cell);
      if (!String(displayValue).trim()) return;
      minRow = Math.min(minRow, rowNumber);
      maxRow = Math.max(maxRow, rowNumber);
      minColumn = Math.min(minColumn, columnNumber);
      maxColumn = Math.max(maxColumn, columnNumber);
    });
  });

  if (!Number.isFinite(minRow) || !Number.isFinite(minColumn)) {
    return null;
  }

  getWorksheetMergeRegions(worksheet).forEach((region) => {
    const masterValue = getCellDisplayValue(worksheet.getCell(region.startRow, region.startColumn));
    if (!String(masterValue).trim()) return;
    minRow = Math.min(minRow, region.startRow);
    maxRow = Math.max(maxRow, region.endRow);
    minColumn = Math.min(minColumn, region.startColumn);
    maxColumn = Math.max(maxColumn, region.endColumn);
  });

  return { minRow, maxRow, minColumn, maxColumn };
}

function extractWorksheetPreview(worksheet, options = {}) {
  const bounds = findWorksheetBounds(worksheet);
  if (!bounds) {
    return {
      columns: [],
      rows: [],
      range: null,
      truncatedRows: false,
      truncatedColumns: false
    };
  }

  const maxRows = Number(options.maxRows || 40);
  const maxColumns = Number(options.maxColumns || 12);
  const endRow = Math.min(bounds.maxRow, bounds.minRow + maxRows - 1);
  const endColumn = Math.min(bounds.maxColumn, bounds.minColumn + maxColumns - 1);
  const columns = [];
  const rows = [];
  const mergeRegions = getWorksheetMergeRegions(worksheet)
    .filter((region) => (
      region.endRow >= bounds.minRow
      && region.startRow <= endRow
      && region.endColumn >= bounds.minColumn
      && region.startColumn <= endColumn
    ));
  const mergedCellsByKey = new Map();

  for (let columnNumber = bounds.minColumn; columnNumber <= endColumn; columnNumber += 1) {
    columns.push({
      index: columnNumber,
      label: columnNumberToLabel(columnNumber)
    });
  }

  mergeRegions.forEach((region) => {
    for (let rowNumber = region.startRow; rowNumber <= Math.min(region.endRow, endRow); rowNumber += 1) {
      for (let columnNumber = region.startColumn; columnNumber <= Math.min(region.endColumn, endColumn); columnNumber += 1) {
        mergedCellsByKey.set(`${rowNumber}:${columnNumber}`, region);
      }
    }
  });

  for (let rowNumber = bounds.minRow; rowNumber <= endRow; rowNumber += 1) {
    const worksheetRow = worksheet.getRow(rowNumber);
    const cells = [];

    for (let columnNumber = bounds.minColumn; columnNumber <= endColumn; columnNumber += 1) {
      const mergeRegion = mergedCellsByKey.get(`${rowNumber}:${columnNumber}`);
      if (mergeRegion) {
        const isMasterCell = mergeRegion.startRow === rowNumber && mergeRegion.startColumn === columnNumber;
        if (!isMasterCell) continue;
        cells.push({
          column: columnNumber,
          value: getCellDisplayValue(worksheetRow.getCell(columnNumber)),
          rowSpan: Math.min(mergeRegion.endRow, endRow) - mergeRegion.startRow + 1,
          colSpan: Math.min(mergeRegion.endColumn, endColumn) - mergeRegion.startColumn + 1,
          merged: true
        });
        continue;
      }

      cells.push({
        column: columnNumber,
        value: getCellDisplayValue(worksheetRow.getCell(columnNumber)),
        rowSpan: 1,
        colSpan: 1,
        merged: false
      });
    }

    rows.push({
      rowNumber,
      cells
    });
  }

  return {
    columns,
    rows,
    range: {
      startRow: bounds.minRow,
      endRow,
      startColumn: bounds.minColumn,
      endColumn
    },
    totalRows: bounds.maxRow - bounds.minRow + 1,
    totalColumns: bounds.maxColumn - bounds.minColumn + 1,
    mergeRegions: mergeRegions.map((region) => ({
      startRow: region.startRow,
      endRow: Math.min(region.endRow, endRow),
      startColumn: region.startColumn,
      endColumn: Math.min(region.endColumn, endColumn)
    })),
    truncatedRows: endRow < bounds.maxRow,
    truncatedColumns: endColumn < bounds.maxColumn
  };
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
  const requestedSheetName = payload.sheetName == null ? "" : String(payload.sheetName);
  const sheetOnly = Boolean(payload.sheetOnly);
  const resolvedSheetName = resolveSheetName(requestedSheetName, reportSheets);

  if (!month) {
    throw new Error("month is required");
  }

  if (sheetOnly && !normalizeSheetName(requestedSheetName)) {
    throw new Error("sheetOnly requires sheetName");
  }

  if (normalizeSheetName(requestedSheetName) && !resolvedSheetName) {
    throw new Error(`sheetName must be one of: ${reportSheets.map((item) => normalizeSheetName(item)).join(", ")}`);
  }

  return {
    month,
    sheetName: resolvedSheetName,
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

function resolveFileIdToPath(fileId, downloadRoot) {
  let relativePath;
  try {
    relativePath = decodeFileId(fileId);
  } catch (_error) {
    throw new Error("invalid fileId");
  }

  const resolvedPath = normalizePathForCompare(path.join(downloadRoot, relativePath));
  if (!isPathInsideRoot(resolvedPath, downloadRoot)) {
    throw new Error("file is outside download root");
  }

  return resolvedPath;
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

  app.get("/api/reports/preview/:fileId", requireAuth, async (req, res) => {
    try {
      const resolvedPath = resolveFileIdToPath(req.params.fileId, downloadRoot);
      if (!fs.existsSync(resolvedPath)) {
        res.status(404).json({ error: "file not found" });
        return;
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(resolvedPath);
      const availableSheetNames = workbook.worksheets.map((item) => item.name);
      const requestedSheetName = req.query.sheet == null ? "" : String(req.query.sheet);
      const resolvedSheetName = requestedSheetName ? resolveSheetName(requestedSheetName, availableSheetNames) : undefined;
      const worksheet = resolvedSheetName ? workbook.getWorksheet(resolvedSheetName) : workbook.worksheets[0];

      if (!worksheet) {
        res.status(400).json({ error: requestedSheetName ? `sheet not found: ${normalizeSheetName(requestedSheetName)}` : "workbook has no worksheets" });
        return;
      }

      res.json({
        file: {
          id: req.params.fileId,
          name: path.basename(resolvedPath)
        },
        sheets: workbook.worksheets.map((item) => item.name),
        sheetName: worksheet.name,
        preview: extractWorksheetPreview(worksheet)
      });
    } catch (error) {
      const statusCode = error.message === "invalid fileId" || error.message === "file is outside download root" ? 400 : 500;
      res.status(statusCode).json({ error: error.message });
    }
  });

  app.get("/api/reports/download/:fileId", requireAuth, async (req, res) => {
    if (!req.params.fileId) {
      res.status(400).json({ error: "fileId is required" });
      return;
    }

    try {
      const resolvedPath = resolveFileIdToPath(req.params.fileId, downloadRoot);
      if (!fs.existsSync(resolvedPath)) {
        res.status(404).json({ error: "file not found" });
        return;
      }

      res.download(resolvedPath);
    } catch (error) {
      res.status(400).json({ error: error.message });
      return;
    }
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
  resolveFileIdToPath,
  resolvePublicBaseUrl,
  startServer
};
