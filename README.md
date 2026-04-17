# Oracle Excel Report Service

按指定月份从 Oracle 查询基础数据，写入 Excel 模板中的基础值单元格，并保留模板原有公式。现在同时提供集中部署 API、内置调用页面、项目级 skill 和 MCP 封装。

## 快速开始

安装依赖：

```bash
cd /Users/xiaochen/Desktop/123
npm install
```

生成整本：

```bash
node src/cli.js --month 2025-12
```

启动 HTTP 服务：

```bash
npm start
```

启动 MCP server（供 AI 工具接入）：

```bash
REPORT_API_BASE_URL=http://127.0.0.1:3000 npm run mcp
```

## 文档导航

- 配置使用手册：`docs/配置使用手册.md`
- 部署手册：`docs/部署手册.md`
- MCP 接入说明：`docs/MCP接入说明.md`
- 规则配置文件：`config/report-rules.json`
- 环境变量示例：`.env.example`

## 访问入口

- 页面入口：启动服务后访问 `http://127.0.0.1:3000/`
- API：`/api/reports/generate`、`/api/reports/debug/unmatched`、`/api/reports/download/:fileId`
- 项目级 skill：仓库内 `.codex/skills/oracle-report-service-api/`
- MCP：`npm run mcp`
- GHCR 镜像：`ghcr.io/xiaoguan521/baobiao:baobiao-YYYYMMDDHHmm`，时间按东八区生成

## 常用命令

单个 sheet 导出：

```bash
node src/cli.js --month 2025-12 --sheet 排名 --sheet-only
```

查看未匹配网点：

```bash
node src/cli.js --month 2025-12 --debug-unmatched --limit 20
```

使用自定义规则文件启动：

```bash
REPORT_RULES_PATH=/your/path/report-rules.json npm start
```

启用 API 鉴权：

```bash
REPORT_API_TOKEN=replace-with-a-strong-token npm start
```

把项目级 skill 安装到本机 Codex：

```bash
npm run skill:install
```

## Docker 镜像发布

仓库已包含 GitHub Actions 多架构构建配置：

`/Users/xiaochen/Desktop/123/.github/workflows/docker-multiarch.yml`

默认行为：

- `push` 到 `main` 或 `master` 时构建并推送 `linux/amd64` 和 `linux/arm64`
- 打版本标签 `v*` 时也会推送镜像
- `pull_request` 只做构建校验，不推送镜像
- 镜像默认发布到 `ghcr.io/<owner>/<repo>`

如果仓库启用了 GitHub Packages，使用内置 `GITHUB_TOKEN` 就可以推送到 GHCR。

另外还包含一个常规 CI 工作流，会执行 `npm test` 和 Docker build 校验。
