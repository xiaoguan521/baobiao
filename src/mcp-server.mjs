#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as z from "zod/v4";
import sqlExplainerModule from "./report-sql-explainer.js";

const { explainFieldSql } = sqlExplainerModule;

const serviceBaseUrl = normalizeBaseUrl(process.env.REPORT_API_BASE_URL || "http://127.0.0.1:3000");
const apiToken = String(process.env.REPORT_API_TOKEN || "").trim();

function normalizeBaseUrl(url) {
  return String(url || "").replace(/\/+$/, "");
}

async function callApi(pathname, options = {}) {
  const url = `${serviceBaseUrl}${pathname}`;
  const headers = {
    accept: "application/json",
    ...(options.headers || {})
  };

  if (apiToken) {
    headers.authorization = `Bearer ${apiToken}`;
  }

  const response = await fetch(url, {
    ...options,
    headers
  });

  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json") ? await response.json() : { raw: await response.text() };

  if (!response.ok) {
    throw new Error(payload.error || payload.raw || `Request failed with status ${response.status}`);
  }

  return payload;
}

function asTextResult(payload) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(payload, null, 2)
      }
    ]
  };
}

const server = new McpServer({
  name: "oracle-report-service-mcp",
  version: "1.0.0"
});

server.registerTool(
  "health_check",
  {
    description: "Check whether the centralized Oracle report service is reachable.",
    outputSchema: {
      ok: z.boolean(),
      authRequired: z.boolean()
    }
  },
  async () => {
    const payload = await callApi("/api/health");
    return {
      ...asTextResult(payload),
      structuredContent: payload
    };
  }
);

server.registerTool(
  "list_report_sheets",
  {
    description: "List the report sheets supported by the centralized report service.",
    outputSchema: {
      sheets: z.array(z.string()),
      configPath: z.string(),
      authRequired: z.boolean()
    }
  },
  async () => {
    const payload = await callApi("/api/reports/sheets");
    return {
      ...asTextResult(payload),
      structuredContent: payload
    };
  }
);

server.registerTool(
  "generate_report",
  {
    description: "Generate an Excel report for a month. Use sheetOnly only when exporting a single sheet.",
    inputSchema: {
      month: z.string().describe("Month in YYYY-MM format"),
      sheetName: z.string().optional().describe("Optional sheet name, for example 排名"),
      sheetOnly: z.boolean().optional().default(false).describe("Whether to export only the selected sheet")
    }
  },
  async ({ month, sheetName, sheetOnly }) => {
    const payload = await callApi("/api/reports/generate", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        month,
        sheetName,
        sheetOnly
      })
    });
    return asTextResult(payload);
  }
);

server.registerTool(
  "debug_unmatched_outlets",
  {
    description: "Inspect unmatched outlets for a month to help maintain outletRules.",
    inputSchema: {
      month: z.string().describe("Month in YYYY-MM format"),
      limit: z.number().int().min(1).max(200).optional().default(20).describe("Maximum rows per top list")
    }
  },
  async ({ month, limit }) => {
    const params = new URLSearchParams({
      month,
      limit: String(limit)
    });
    const payload = await callApi(`/api/reports/debug/unmatched?${params.toString()}`);
    return asTextResult(payload);
  }
);

server.registerTool(
  "explain_sheet_field_sql",
  {
    description: "Explain which SQL and post-processing logic are used for a field in a report sheet.",
    inputSchema: {
      sheetName: z.string().describe("Sheet name, for example 市本级, 排名, 占比"),
      fieldName: z.string().describe("Field or column name, for example 归集, 比上月增减, 合计"),
      rowName: z.string().optional().describe("Optional row label such as a branch name, bank name, item name, or reviewer row item")
    }
  },
  async ({ sheetName, fieldName, rowName }) => {
    const payload = explainFieldSql({ sheetName, fieldName, rowName });
    return {
      ...asTextResult(payload),
      structuredContent: payload
    };
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`oracle-report-service MCP server connected via stdio -> ${serviceBaseUrl}`);
}

main().catch((error) => {
  console.error("MCP server error:", error);
  process.exit(1);
});
