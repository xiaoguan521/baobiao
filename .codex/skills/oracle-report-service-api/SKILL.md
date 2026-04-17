---
name: oracle-report-service-api
description: Use this skill when working with the centralized Oracle Excel report service in this repository, especially to generate monthly reports, export a single sheet, inspect unmatched outlets, or guide users to the built-in web page and MCP wrapper instead of using the local CLI or direct database access.
---

# Oracle Report Service API

This project should be operated through the centralized HTTP service, not by asking end users to install Node.js, run the CLI, or connect to Oracle directly.

## When to use

Use this skill when the user asks to:

- generate a monthly Excel report
- export only one sheet such as `排名`
- inspect unmatched outlets for rule maintenance
- connect AI tools to the deployed report service
- explain whether to use the page, skill, or MCP entrypoint

## Preferred entrypoints

Choose the lightest entrypoint that fits the user:

- Business users: use the built-in page at `/`
- Codex users in this repo: call the HTTP API directly
- External AI clients: use the MCP server in `src/mcp-server.mjs`

Do not recommend local CLI for end users unless they are explicitly doing development or troubleshooting.

## API workflow

1. Check service health with `GET /api/health`.
2. Fetch supported sheets with `GET /api/reports/sheets` before proposing `sheetName`.
3. Generate reports with `POST /api/reports/generate`.
4. Use `GET /api/reports/debug/unmatched` only for rule-maintenance or troubleshooting requests.
5. Download files from the returned `file.downloadUrl` or `file.id`.

## Auth

If `authRequired` is true, send either:

- `Authorization: Bearer <token>`
- or `X-API-Token: <token>`

Prefer `Authorization: Bearer`.

## Request patterns

Generate a full workbook:

```json
{
  "month": "2025-12"
}
```

Generate a single sheet:

```json
{
  "month": "2025-12",
  "sheetName": "排名",
  "sheetOnly": true
}
```

Debug unmatched outlets:

```text
GET /api/reports/debug/unmatched?month=2025-12&limit=20
```

## Response handling

For `POST /api/reports/generate`, expect:

- `month`
- `sheetName`
- `sheetOnly`
- `stats`
- `generatedAt`
- `file.id`
- `file.name`
- `file.sizeBytes`
- `file.downloadUrl`

Do not expect or ask for absolute server file paths.

## Files to consult when needed

- API and deployment notes: `references/api.md`
- Repository docs: `/Users/xiaochen/Desktop/123/docs/配置使用手册.md`
- MCP usage: `/Users/xiaochen/Desktop/123/docs/MCP接入说明.md`

## Guardrails

- Keep database credentials on the server side only.
- Do not tell users to install local Node.js just to generate reports.
- If the user is choosing between page, skill, and MCP:
  - recommend page for non-technical users
  - recommend skill for Codex-internal use
  - recommend MCP for cross-client AI integration
