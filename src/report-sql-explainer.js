const {
  BANK_NAME_BY_OUTLET,
  BANK_ORDER,
  BUSINESS_RULES,
  CENTER_NAME_BY_REGION,
  CHANNEL_BUCKET_ORDER,
  CHANNEL_BUCKETS,
  OUTLET_RULES,
  REGION_NAME_BY_CODE,
  REPORT_SHEETS,
  REVIEWER_NAMES,
  SUMMARY_GROUP_ORDER
} = require("./report-rules");

const REGION_SHEETS = new Set(["市本级", "东港", "凤城", "宽甸"]);
const CHANNEL_SHEET = "各渠道复核业务来源统计";
const REVIEWER_SHEET = "复核业务量统计";
const TOTAL_SHEET = "总表";
const SHARE_SHEET = "占比";
const RANKING_SHEET = "排名";

const BASE_DETAIL_SQL = `
select
  dx_05_dxcjsj,
  nvl(dx_05_dxmc, ' ') as dxmc,
  nvl(dx_05_dxms, ' ') as dxms,
  trim(nvl(dx_05_dxcjqd, ' ')) as channel,
  nvl(dx_05_spr, ' ') as reviewer,
  nvl(dx_05_glb, ' ') as region_code,
  nvl(dx_05_dxcjzmc, ' ') as creator
from PT_DXSL_2106032001_05
where dx_05_dxcjsj >= :start_time
  and dx_05_dxcjsj < :end_time
`.trim();

const BASE_TEXT_EXPR = "nvl(dx_05_dxmc, ' ') || ' ' || nvl(dx_05_dxms, ' ')";
const DXMS_EXPR = "nvl(dx_05_dxms, ' ')";
const CREATOR_EXPR = "nvl(dx_05_dxcjzmc, ' ')";
const CHANNEL_EXPR = "trim(nvl(dx_05_dxcjqd, ' '))";
const REVIEWER_EXPR = "nvl(dx_05_spr, ' ')";

const REGION_CODE_BY_NAME = Object.fromEntries(
  Object.entries(REGION_NAME_BY_CODE).map(([code, name]) => [name, code])
);

const OUTLET_REGION_BY_NAME = Object.fromEntries(
  Object.entries(OUTLET_RULES).flatMap(([region, rules]) => rules.map((rule) => [rule.name, region]))
);

function escapeSqlLiteral(text) {
  return String(text).replace(/'/g, "''");
}

function normalizeSheetName(sheetName) {
  return String(sheetName || "").trim();
}

function normalizeFieldName(fieldName) {
  return String(fieldName || "").trim();
}

function normalizeRowName(rowName) {
  const normalized = String(rowName || "").trim();
  return normalized || null;
}

function buildContainsCondition(patterns, expr = BASE_TEXT_EXPR) {
  const uniquePatterns = [...new Set((patterns || []).filter(Boolean))];
  if (!uniquePatterns.length) return null;
  return uniquePatterns.map((pattern) => `instr(${expr}, '${escapeSqlLiteral(pattern)}') > 0`).join(" or ");
}

function buildRegexCondition(patterns, expr = DXMS_EXPR) {
  const uniquePatterns = [...new Set((patterns || []).filter(Boolean))];
  if (!uniquePatterns.length) return null;
  return uniquePatterns.map((pattern) => `regexp_like(${expr}, '${escapeSqlLiteral(pattern)}')`).join(" or ");
}

function buildInCondition(values, expr) {
  const uniqueValues = [...new Set((values || []).filter(Boolean))];
  if (!uniqueValues.length) return null;
  return `${expr} in (${uniqueValues.map((value) => `'${escapeSqlLiteral(value)}'`).join(", ")})`;
}

function buildRuleCondition(rule) {
  const regexCondition = buildRegexCondition(rule.patterns || []);
  const creatorCondition = buildInCondition(rule.creators || [], CREATOR_EXPR);
  return [regexCondition, creatorCondition].filter(Boolean).join(" or ") || null;
}

function buildSummaryGroupCondition(summaryGroup) {
  const patterns = BUSINESS_RULES
    .filter((rule) => rule.summaryGroup === summaryGroup)
    .flatMap((rule) => rule.patterns || []);
  return buildContainsCondition(patterns);
}

function buildItemCondition(itemName) {
  const patterns = BUSINESS_RULES
    .filter((rule) => rule.item === itemName)
    .flatMap((rule) => rule.patterns || []);
  return buildContainsCondition(patterns);
}

function buildClassifiedCondition() {
  return SUMMARY_GROUP_ORDER.map((group) => `(${buildSummaryGroupCondition(group)})`).join(" or ");
}

function buildRegionCondition(regionName) {
  const code = REGION_CODE_BY_NAME[regionName];
  if (!code) return null;
  return `nvl(dx_05_glb, ' ') = '${escapeSqlLiteral(code)}'`;
}

function buildOutletCondition(regionName, rowName) {
  const regionRules = OUTLET_RULES[regionName] || [];
  const targetName = rowName || "<网点名称>";
  const centerName = CENTER_NAME_BY_REGION[regionName];

  if (!rowName) {
    return {
      condition: "-- 这里需要替换成具体网点规则",
      notes: ["未提供 rowName，SQL 中保留了网点条件占位符。"]
    };
  }

  if (rowName === centerName) {
    const allConditions = regionRules.map(buildRuleCondition).filter(Boolean);
    const centerConditions = regionRules.filter((rule) => rule.name === centerName).map(buildRuleCondition).filter(Boolean);
    const unmatchedCondition = allConditions.length ? `not (${allConditions.join(" or ")})` : "1 = 1";
    const explicitCenterCondition = centerConditions.length ? `(${centerConditions.join(" or ")})` : null;
    const condition = explicitCenterCondition ? `(${explicitCenterCondition} or ${unmatchedCondition})` : unmatchedCondition;
    return {
      condition,
      notes: [
        `中心行 ${targetName} 的值 = 显式命中中心规则的记录 + 未命中任何网点规则的兜底记录。`,
        "这和报表代码里的 `outlet || centerName` 回退逻辑一致。"
      ]
    };
  }

  const targetRules = regionRules.filter((rule) => rule.name === rowName).map(buildRuleCondition).filter(Boolean);
  if (!targetRules.length) {
    return {
      condition: "-- 未找到该网点的 outletRules，请先检查 config/report-rules.json",
      notes: [`未在地区 ${regionName} 的 outletRules 中找到网点 ${targetName}。`]
    };
  }

  return {
    condition: `(${targetRules.join(" or ")})`,
    notes: [`网点 ${targetName} 通过 outletRules 中的 patterns/creators 匹配。`]
  };
}

function buildMonthWindow(scope) {
  if (scope === "current_month") {
    return {
      label: "当月",
      where: "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
      binds: [":month_start", ":next_month"]
    };
  }
  if (scope === "previous_month") {
    return {
      label: "上月",
      where: "dx_05_dxcjsj >= :prev_month_start and dx_05_dxcjsj < :month_start",
      binds: [":prev_month_start", ":month_start"]
    };
  }
  return {
    label: "年初至本月",
    where: "dx_05_dxcjsj >= :year_start and dx_05_dxcjsj < :next_month",
    binds: [":year_start", ":next_month"]
  };
}

function buildCountSql(whereClauses) {
  const clauses = whereClauses.filter(Boolean);
  return `
select count(*) as field_value
from PT_DXSL_2106032001_05
where ${clauses.join("\n  and ")}
`.trim();
}

function buildRegionFieldExplanation(sheetName, fieldName, rowName) {
  const regionCondition = buildRegionCondition(sheetName);
  const outlet = buildOutletCondition(sheetName, rowName);
  const classifiedCondition = buildClassifiedCondition();

  if (SUMMARY_GROUP_ORDER.includes(fieldName)) {
    const businessCondition = buildSummaryGroupCondition(fieldName);
    return {
      supported: true,
      sheetName,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        regionCondition,
        `(${businessCondition})`,
        outlet.condition
      ]),
      explanation: `${sheetName} sheet 的 ${fieldName} 列，按地区 + 网点 + summaryGroup=${fieldName} 统计当月条数。`,
      requiredBinds: [":month_start", ":next_month"],
      notes: outlet.notes
    };
  }

  if (fieldName === "合计") {
    return {
      supported: true,
      sheetName,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        regionCondition,
        `(${classifiedCondition})`,
        outlet.condition
      ]),
      explanation: `${sheetName} sheet 的合计列 = 归集 + 提取 + 贷款 + 贷后，当月范围内按地区和网点汇总。`,
      requiredBinds: [":month_start", ":next_month"],
      notes: outlet.notes
    };
  }

  if (fieldName === "年初至本月累计数量") {
    return {
      supported: true,
      sheetName,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :year_start and dx_05_dxcjsj < :next_month",
        regionCondition,
        `(${classifiedCondition})`,
        outlet.condition
      ]),
      explanation: `${sheetName} sheet 的年初至本月累计数量，按年初到本月末的真实明细累计。`,
      requiredBinds: [":year_start", ":next_month"],
      notes: outlet.notes
    };
  }

  if (fieldName === "比上月增减") {
    const currentSql = buildCountSql([
      "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
      regionCondition,
      `(${classifiedCondition})`,
      outlet.condition
    ]);
    const previousSql = buildCountSql([
      "dx_05_dxcjsj >= :prev_month_start and dx_05_dxcjsj < :month_start",
      regionCondition,
      `(${classifiedCondition})`,
      outlet.condition
    ]);
    return {
      supported: true,
      sheetName,
      fieldName,
      rowName,
      sqlKind: "derived",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: `with current_month as (\n  ${indentSql(currentSql)}\n), previous_month as (\n  ${indentSql(previousSql)}\n)\nselect case\n  when previous_month.field_value = 0 then null\n  else (current_month.field_value - previous_month.field_value) / previous_month.field_value\nend as ratio_value\nfrom current_month, previous_month`,
      explanation: `${sheetName} sheet 的比上月增减不是单条原始 SQL 字段，而是“本月合计 - 上月合计”再除以上月合计的代码层计算。`,
      requiredBinds: [":prev_month_start", ":month_start", ":next_month"],
      notes: outlet.notes
    };
  }

  return unsupportedResponse(sheetName, fieldName, rowName, [
    "地区 sheet 当前支持字段：归集、提取、贷款、贷后、合计、比上月增减、年初至本月累计数量。"
  ]);
}

function buildTotalSheetExplanation(fieldName, rowName) {
  const monthMatch = fieldName.match(/^(\d{1,2})月$/);
  const bankName = rowName || "<银行名称或中心办件>";

  if (monthMatch) {
    return {
      supported: true,
      sheetName: TOTAL_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${buildClassifiedCondition()})`,
        "-- 银行行需要按 bankNameByOutlet 反推 outletRules，再按月汇总；中心办件则汇总未映射银行的记录"
      ]),
      explanation: `总表的 ${fieldName} 列来自真实 SQL 明细，但银行/中心的月度汇总是在代码里按 bankNameByOutlet 和 outletRules 再次聚合。`,
      requiredBinds: [":month_start", ":next_month"],
      notes: [`当前 rowName=${bankName}，如需精确校验，建议同时提供银行名称或“中心办件”。`]
    };
  }

  if (["合计", "年度合计"].includes(fieldName)) {
    return {
      supported: true,
      sheetName: TOTAL_SHEET,
      fieldName,
      rowName,
      sqlKind: "derived",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :year_start and dx_05_dxcjsj < :next_year_start",
        `(${buildClassifiedCondition()})`,
        "-- 银行行需要按 bankNameByOutlet 反推 outletRules，再汇总全年"
      ]),
      explanation: "总表合计列是全年各月累计，不是数据库里的单独字段。",
      requiredBinds: [":year_start", ":next_year_start"],
      notes: ["如果要精确核银行行，请补充 rowName。"]
    };
  }

  return unsupportedResponse(TOTAL_SHEET, fieldName, rowName, [
    "总表建议按“银行/中心办件 + 月份列”来问，例如 rowName=工商银行, fieldName=12月。"
  ]);
}

function buildShareSheetExplanation(fieldName, rowName) {
  if (![...SUMMARY_GROUP_ORDER, "合计"].includes(fieldName)) {
    return unsupportedResponse(SHARE_SHEET, fieldName, rowName, [
      "占比 sheet 当前支持字段：归集、提取、贷款、贷后、合计。"
    ]);
  }

  const bankName = rowName || "<银行名称或中心办件>";
  const businessCondition = fieldName === "合计" ? buildClassifiedCondition() : buildSummaryGroupCondition(fieldName);

  return {
    supported: true,
    sheetName: SHARE_SHEET,
    fieldName,
    rowName,
    sqlKind: "aggregate",
    sourceSql: BASE_DETAIL_SQL,
    verificationSql: buildCountSql([
      "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
      `(${businessCondition})`,
      "-- 需要按 bankNameByOutlet 把 outlet 汇总到银行；中心办件汇总 bankName 为空的记录"
    ]),
    explanation: `占比 sheet 的 ${fieldName} 列来自当月真实明细，但“银行/中心办件”层级是在代码里根据 bankNameByOutlet 再汇总。`,
    requiredBinds: [":month_start", ":next_month"],
    notes: [`当前 rowName=${bankName}。如需精确校验，请明确是哪个银行或“中心办件”。`]
  };
}

function buildRankingSheetExplanation(fieldName, rowName) {
  const outletName = rowName || "<网点名称>";

  if (fieldName === "本月业务量") {
    return {
      supported: true,
      sheetName: RANKING_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${buildClassifiedCondition()})`,
        "-- 需要先根据 outletRules 判断该网点属于哪个地区，再按 outlet 汇总"
      ]),
      explanation: "排名 sheet 的本月业务量来自当月真实明细汇总。",
      requiredBinds: [":month_start", ":next_month"],
      notes: [`当前 rowName=${outletName}。网点定位依赖 outletRules。`]
    };
  }

  if (fieldName === "年初至本月累计数量") {
    return {
      supported: true,
      sheetName: RANKING_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :year_start and dx_05_dxcjsj < :next_month",
        `(${buildClassifiedCondition()})`,
        "-- 需要先根据 outletRules 判断该网点属于哪个地区，再按 outlet 汇总"
      ]),
      explanation: "排名 sheet 的年初至本月累计数量来自年累计真实明细汇总。",
      requiredBinds: [":year_start", ":next_month"],
      notes: [`当前 rowName=${outletName}。网点定位依赖 outletRules。`]
    };
  }

  if (fieldName === "本月排名" || fieldName === "年累计排名") {
    return {
      supported: true,
      sheetName: RANKING_SHEET,
      fieldName,
      rowName,
      sqlKind: "derived",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: "-- 排名字段不是数据库中的直接列，而是先算出各网点数量后，再在代码中按数量排序得到名次",
      explanation: `${fieldName} 不是单条 SQL 直接取值，而是代码在汇总各网点数量后进行排序生成。`,
      requiredBinds: fieldName === "本月排名" ? [":month_start", ":next_month"] : [":year_start", ":next_month"],
      notes: [`当前 rowName=${outletName}。可先查询对应的本月业务量或年累计数量，再对比排序。`]
    };
  }

  return unsupportedResponse(RANKING_SHEET, fieldName, rowName, [
    "排名 sheet 当前支持字段：本月业务量、本月排名、年初至本月累计数量、年累计排名。"
  ]);
}

function buildChannelSheetExplanation(fieldName, rowName) {
  const itemName = rowName || "<事项名称>";
  const itemCondition = rowName ? buildItemCondition(rowName) : "-- 请提供 rowName=事项名称";

  if (!rowName) {
    return unsupportedResponse(CHANNEL_SHEET, fieldName, rowName, [
      "各渠道复核业务来源统计需要提供 rowName，对应 B 列事项名称。"
    ]);
  }

  if (fieldName === "合计") {
    return {
      supported: true,
      sheetName: CHANNEL_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${itemCondition})`
      ]),
      explanation: "各渠道复核业务来源统计的合计列 = 当前事项在当月的全部渠道数量。",
      requiredBinds: [":month_start", ":next_month"],
      notes: ["事项命中依赖 businessRules.item。"]
    };
  }

  if (CHANNEL_BUCKET_ORDER.includes(fieldName) || fieldName === "其他") {
    const bucketValues = CHANNEL_BUCKETS[fieldName] ? [...CHANNEL_BUCKETS[fieldName]] : [];
    const channelCondition = bucketValues.length
      ? buildInCondition(bucketValues, CHANNEL_EXPR)
      : `not (${Object.values(CHANNEL_BUCKETS).flatMap((values) => [...values]).map((value) => `${CHANNEL_EXPR} = '${escapeSqlLiteral(value)}'`).join(" or ")})`;
    return {
      supported: true,
      sheetName: CHANNEL_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${itemCondition})`,
        `(${channelCondition})`
      ]),
      explanation: `${fieldName} 列按事项 + 渠道桶统计当月数量。`,
      requiredBinds: [":month_start", ":next_month"],
      notes: ["渠道分桶依赖 channelBuckets/channelBucketOrder。"]
    };
  }

  if (["工作时间", "中午时间", "非工作时间"].includes(fieldName)) {
    return {
      supported: true,
      sheetName: CHANNEL_SHEET,
      fieldName,
      rowName,
      sqlKind: "derived",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${itemCondition})`,
        `-- ${fieldName} 由配置里的 workTime 时间段判断，不是 SQL 原始列`
      ]),
      explanation: `${fieldName} 列基于真实时间字段 dx_05_dxcjsj，但工作时间/中午时间/非工作时间是在代码里按 workTime 切分得到。`,
      requiredBinds: [":month_start", ":next_month"],
      notes: ["时间段定义来自 config/report-rules.json 中的 workTime。"]
    };
  }

  return unsupportedResponse(CHANNEL_SHEET, fieldName, rowName, [
    "各渠道复核业务来源统计当前支持字段：各渠道列、合计、工作时间、中午时间、非工作时间。"
  ]);
}

function buildReviewerSheetExplanation(fieldName, rowName) {
  if (!rowName) {
    return unsupportedResponse(REVIEWER_SHEET, fieldName, rowName, [
      "复核业务量统计需要提供 rowName，对应 B 列事项名称。"
    ]);
  }

  const itemCondition = buildItemCondition(rowName);
  if (!itemCondition) {
    return unsupportedResponse(REVIEWER_SHEET, fieldName, rowName, [
      `未在 businessRules 中找到事项 ${rowName}。`
    ]);
  }

  if (fieldName === "合计") {
    return {
      supported: true,
      sheetName: REVIEWER_SHEET,
      fieldName,
      rowName,
      sqlKind: "aggregate",
      sourceSql: BASE_DETAIL_SQL,
      verificationSql: buildCountSql([
        "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
        `(${itemCondition})`,
        `(${buildReviewerCondition(REVIEWER_NAMES)})`
      ]),
      explanation: "复核业务量统计的合计列 = 当前事项在当月被 reviewerNames 名单内复核人的总次数。",
      requiredBinds: [":month_start", ":next_month"],
      notes: ["复核人名单来自 reviewerNames。"]
    };
  }

  if (!REVIEWER_NAMES.includes(fieldName)) {
    return unsupportedResponse(REVIEWER_SHEET, fieldName, rowName, [
      `字段 ${fieldName} 不在 reviewerNames 名单内。`
    ]);
  }

  return {
    supported: true,
    sheetName: REVIEWER_SHEET,
    fieldName,
    rowName,
    sqlKind: "aggregate",
    sourceSql: BASE_DETAIL_SQL,
    verificationSql: buildCountSql([
      "dx_05_dxcjsj >= :month_start and dx_05_dxcjsj < :next_month",
      `(${itemCondition})`,
      `(${buildReviewerCondition([fieldName])})`
    ]),
    explanation: `复核业务量统计的 ${fieldName} 列 = 当前事项在当月被该复核人处理的数量。`,
    requiredBinds: [":month_start", ":next_month"],
    notes: ["复核人拆分逻辑和代码里的 splitReviewers 一致，按中文逗号、英文逗号、顿号拆分。"]
  };
}

function buildReviewerCondition(reviewers) {
  const list = reviewers.filter(Boolean);
  return list.map((name) => `instr(replace(replace(replace(${REVIEWER_EXPR}, '，', ','), '、', ','), ' ', ''), '${escapeSqlLiteral(name)}') > 0`).join(" or ");
}

function unsupportedResponse(sheetName, fieldName, rowName, notes = []) {
  return {
    supported: false,
    sheetName,
    fieldName,
    rowName,
    sqlKind: "unsupported",
    sourceSql: BASE_DETAIL_SQL,
    verificationSql: null,
    explanation: "当前 MCP 还没有为这个 sheet/字段组合生成专门的 SQL 模板。",
    requiredBinds: [],
    notes
  };
}

function indentSql(sql) {
  return String(sql)
    .split("\n")
    .map((line) => `  ${line}`)
    .join("\n");
}

function explainFieldSql({ sheetName, fieldName, rowName = null }) {
  const normalizedSheet = normalizeSheetName(sheetName);
  const normalizedField = normalizeFieldName(fieldName);
  const normalizedRow = normalizeRowName(rowName);

  if (!normalizedSheet || !normalizedField) {
    return unsupportedResponse(normalizedSheet || "<missing sheetName>", normalizedField || "<missing fieldName>", normalizedRow, [
      "sheetName 和 fieldName 都是必填。"
    ]);
  }

  if (REGION_SHEETS.has(normalizedSheet)) {
    return buildRegionFieldExplanation(normalizedSheet, normalizedField, normalizedRow);
  }
  if (normalizedSheet === CHANNEL_SHEET) {
    return buildChannelSheetExplanation(normalizedField, normalizedRow);
  }
  if (normalizedSheet === REVIEWER_SHEET) {
    return buildReviewerSheetExplanation(normalizedField, normalizedRow);
  }
  if (normalizedSheet === TOTAL_SHEET) {
    return buildTotalSheetExplanation(normalizedField, normalizedRow);
  }
  if (normalizedSheet === SHARE_SHEET) {
    return buildShareSheetExplanation(normalizedField, normalizedRow);
  }
  if (normalizedSheet === RANKING_SHEET) {
    return buildRankingSheetExplanation(normalizedField, normalizedRow);
  }

  return unsupportedResponse(normalizedSheet, normalizedField, normalizedRow, [
    `当前支持的 sheet 包括：${REPORT_SHEETS.map((item) => item.trim()).join("、")}`
  ]);
}

module.exports = {
  BASE_DETAIL_SQL,
  CHANNEL_SHEET,
  REGION_SHEETS,
  REVIEWER_SHEET,
  RANKING_SHEET,
  SHARE_SHEET,
  TOTAL_SHEET,
  explainFieldSql
};
