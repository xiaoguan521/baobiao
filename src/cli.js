const path = require("path");
const { analyzeUnmatched, generateReport, REPORT_SHEETS } = require("./report-engine");

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const item = argv[i];
    if (!item.startsWith("--")) continue;
    const key = item.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    i += 1;
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.month) {
    console.log("Usage: node src/cli.js --month YYYY-MM [--sheet SHEET_NAME] [--sheet-only] [--template PATH] [--output PATH] [--debug-unmatched]");
    console.log(`Supported sheets: ${REPORT_SHEETS.join(", ")}`);
    process.exit(args.help ? 0 : 1);
  }

  if (args["debug-unmatched"]) {
    const debugResult = await analyzeUnmatched({
      month: args.month,
      limit: Number(args.limit || 30),
      dbConfig: {
        user: args.user || process.env.ORACLE_USER || "damoxing",
        password: args.password || process.env.ORACLE_PASSWORD || "Damoxing123!",
        connectString: args.dsn || process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
      }
    });
    console.log(JSON.stringify(debugResult, null, 2));
    return;
  }

  const result = await generateReport({
    month: args.month,
    sheetName: args.sheet,
    sheetOnly: Boolean(args["sheet-only"]),
    templatePath: args.template || process.env.TEMPLATE_PATH || path.join(process.cwd(), "模板.xlsx"),
    outputPath: args.output,
    outputDir: args["output-dir"] || process.env.OUTPUT_DIR,
    user: args.user || process.env.ORACLE_USER || "damoxing",
    password: args.password || process.env.ORACLE_PASSWORD || "Damoxing123!",
    dsn: args.dsn || process.env.ORACLE_DSN || "127.0.0.1:51521/FREEPDB1"
  });

  console.log(JSON.stringify(result, null, 2));
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
