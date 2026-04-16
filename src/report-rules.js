const fs = require("fs");
const path = require("path");

const DEFAULT_CONFIG_PATH = process.env.REPORT_RULES_PATH || path.join(process.cwd(), "config", "report-rules.json");

function loadRules(configPath = DEFAULT_CONFIG_PATH) {
  const raw = fs.readFileSync(configPath, "utf8");
  const parsed = JSON.parse(raw);
  return {
    BANK_NAME_BY_OUTLET: parsed.bankNameByOutlet,
    BANK_ORDER: parsed.bankOrder,
    BUSINESS_RULES: parsed.businessRules,
    CENTER_NAME_BY_REGION: parsed.centerNameByRegion,
    CHANNEL_BUCKET_ORDER: parsed.channelBucketOrder,
    CHANNEL_BUCKETS: Object.fromEntries(
      Object.entries(parsed.channelBuckets).map(([key, values]) => [key, new Set(values)])
    ),
    OUTLET_RULES: parsed.outletRules,
    REGION_NAME_BY_CODE: parsed.regionNameByCode,
    REPORT_SHEETS: parsed.reportSheets,
    REVIEWER_NAMES: parsed.reviewerNames,
    SUMMARY_GROUP_ORDER: parsed.summaryGroupOrder,
    WORK_TIME: parsed.workTime,
    CONFIG_PATH: configPath
  };
}

module.exports = loadRules();
module.exports.loadRules = loadRules;
module.exports.DEFAULT_CONFIG_PATH = DEFAULT_CONFIG_PATH;
