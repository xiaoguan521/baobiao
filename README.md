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
