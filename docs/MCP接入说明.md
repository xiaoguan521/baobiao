# MCP 接入说明

## 1. 目的

本项目提供一个本地 `stdio` MCP server，用来把集中部署的报表 HTTP 服务暴露给 AI 工具。

请注意：

- 这是本地进程型 MCP，不是部署在服务器上的 remote MCP 服务
- 它需要运行在 AI 客户端所在机器上，例如 Codex、Claude Desktop 或其他支持 MCP 的客户端宿主机
- 它本身不直接连 Oracle，只转调已经部署好的报表 HTTP 服务

MCP server 自身不连数据库，只调用已经部署好的：

- `GET /api/health`
- `GET /api/reports/sheets`
- `POST /api/reports/generate`
- `GET /api/reports/debug/unmatched`

## 2. 启动方式

```bash
cd /Users/xiaochen/Desktop/123
REPORT_API_BASE_URL='https://report.example.com' \
REPORT_API_TOKEN='replace-with-a-strong-token' \
npm run mcp
```

适用前提：

1. 报表 HTTP 服务已经部署并可访问
2. 运行 MCP 的那台机器能访问 `REPORT_API_BASE_URL`
3. 如果服务端开启了鉴权，需要同时提供 `REPORT_API_TOKEN`

环境变量说明：

| 变量名 | 说明 |
|---|---|
| `REPORT_API_BASE_URL` | 已部署报表服务的基地址 |
| `REPORT_API_TOKEN` | 如果服务端开启鉴权则需要提供 |

## 3. MCP 工具列表

当前 MCP server 暴露 5 个工具：

- `health_check`
- `list_report_sheets`
- `generate_report`
- `debug_unmatched_outlets`
- `explain_sheet_field_sql`

## 4. 推荐调用方式

### 4.1 生成整本

参数：

```json
{
  "month": "2025-12"
}
```

### 4.2 生成单个 sheet

参数：

```json
{
  "month": "2025-12",
  "sheetName": "排名",
  "sheetOnly": true
}
```

### 4.3 排查未匹配网点

参数：

```json
{
  "month": "2025-12",
  "limit": 20
}
```

### 4.4 查询某个 sheet 字段的取值 SQL

参数示例：

```json
{
  "sheetName": "市本级",
  "fieldName": "比上月增减",
  "rowName": "工行锦山支行"
}
```

说明：

- `sheetName` 填 sheet 名称
- `fieldName` 填列名或字段名，例如 `归集`、`合计`、`比上月增减`
- `rowName` 可选，但强烈建议提供
- 对地区 sheet，`rowName` 一般是网点名称
- 对“各渠道复核业务来源统计”和“复核业务量统计”，`rowName` 一般是 B 列事项名称

返回内容会包含：

- 实际明细来源 SQL
- 用于校验的 SQL 模板
- 这个字段是否还有代码层二次计算
- 依赖哪些规则配置

## 5. 设计边界

- 报表文件由服务端生成和保存
- MCP 响应里会包含下载地址，但下载仍然通过服务端完成
- 客户端不需要 Node CLI 权限，也不需要直连 Oracle
- 普通业务用户更适合使用页面入口，不必接触 MCP
- 如果后续想让外部客户端直接通过网络接入 MCP，需要再实现 remote MCP 服务形态
