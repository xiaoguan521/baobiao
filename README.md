# Oracle Excel Report Service

按指定月份从 Oracle 查询基础数据，写入 Excel 模板中的基础值单元格，并保留模板原有公式。

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

## 文档导航

- 配置使用手册：`docs/配置使用手册.md`
- 部署手册：`docs/部署手册.md`
- 规则配置文件：`config/report-rules.json`
- 环境变量示例：`.env.example`

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
