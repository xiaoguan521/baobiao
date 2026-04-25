const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const oracledb = require("oracledb");

const {
  BANK_NAME_BY_OUTLET,
  BANK_ORDER,
  BUSINESS_RULES,
  CENTER_NAME_BY_REGION,
  CHANNEL_BUCKET_ORDER,
  CHANNEL_BUCKETS,
  CONFIG_PATH,
  OUTLET_RULES,
  REGION_NAME_BY_CODE,
  REPORT_SHEETS,
  REVIEWER_NAMES,
  SUMMARY_GROUP_ORDER,
  WORK_TIME
} = require("./report-rules");

const EXCLUDED_ZBMBH_FOR_NON_CHANNEL_REPORTS = new Set([
  "2106032001",
  "210603200102",
  "210603200103",
  "210603200104",
  "210603200105"
]);

function compileContext() {
  return {
    businessRules: BUSINESS_RULES.map((item) => ({ ...item })),
    outletRules: Object.fromEntries(
      Object.entries(OUTLET_RULES).map(([region, rules]) => [
        region,
        rules.map((item) => ({
          ...item,
          regexes: item.patterns.map((pattern) => new RegExp(pattern))
        }))
      ])
    )
  };
}

function parseMonthBounds(monthText) {
  if (!/^\d{4}-\d{2}$/.test(monthText)) {
    throw new Error("month must be in YYYY-MM format");
  }
  const [year, month] = monthText.split("-").map(Number);
  const monthStart = new Date(year, month - 1, 1);
  const nextMonth = new Date(year, month, 1);
  const yearStart = new Date(year, 0, 1);
  const prevMonth = new Date(year, month - 2, 1);
  const queryStart = prevMonth < yearStart ? prevMonth : yearStart;
  return { year, month, monthKey: monthText, monthStart, nextMonth, yearStart, queryStart, prevMonthKey: formatMonthKey(prevMonth) };
}

function formatMonthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function createEmptyData() {
  return {
    monthlyBankTotals: new Map(),
    monthlyCenterTotals: new Map(),
    monthlyAllTotals: new Map(),
    regionMonthOutletGroup: new Map(),
    regionYtdOutletGroup: new Map(),
    regionPrevOutletGroup: new Map(),
    channelItemCounts: new Map(),
    channelItemTimeCounts: new Map(),
    reviewerItemCounts: new Map()
  };
}

function createStats() {
  return {
    scannedRows: 0,
    classifiedRows: 0,
    unclassifiedRows: 0,
    matchedOutletRows: 0,
    unmatchedOutletRows: 0
  };
}

function textFromCellValue(value) {
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "object" && Array.isArray(value.richText)) {
    return value.richText.map((item) => item.text || "").join("");
  }
  return String(value);
}

function numericFromCellValue(value) {
  if (value == null) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "string") {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  if (typeof value === "object") {
    if (Object.prototype.hasOwnProperty.call(value, "result")) {
      return numericFromCellValue(value.result);
    }
    if (Object.prototype.hasOwnProperty.call(value, "text")) {
      return numericFromCellValue(value.text);
    }
  }
  return 0;
}

function createTopCollector(limit) {
  return {
    limit,
    map: new Map()
  };
}

function collectTopItem(collector, key, sample) {
  const current = collector.map.get(key);
  if (current) {
    current.count += 1;
    return;
  }
  collector.map.set(key, { count: 1, sample });
}

function topCollectorToArray(collector) {
  return [...collector.map.entries()]
    .map(([key, value]) => ({ key, count: value.count, sample: value.sample }))
    .sort((a, b) => b.count - a.count || String(a.key).localeCompare(String(b.key)))
    .slice(0, collector.limit);
}

function getNestedCount(map, keys) {
  let current = map;
  for (const key of keys.slice(0, -1)) {
    if (!current.has(key)) current.set(key, new Map());
    current = current.get(key);
  }
  const lastKey = keys[keys.length - 1];
  current.set(lastKey, (current.get(lastKey) || 0) + 1);
}

function getMapValue(map, ...keys) {
  let current = map;
  for (const key of keys) {
    if (!(current instanceof Map) || !current.has(key)) return undefined;
    current = current.get(key);
  }
  return current;
}

function sumMapValues(mapLike) {
  if (!(mapLike instanceof Map)) return 0;
  let total = 0;
  for (const value of mapLike.values()) total += value || 0;
  return total;
}

function classifyBusiness(dxmc, dxms, businessRules) {
  const combined = `${dxmc} ${dxms}`;
  return businessRules.find((rule) => rule.patterns.some((pattern) => combined.includes(pattern))) || null;
}

function classifyChannel(channel) {
  for (const [bucket, values] of Object.entries(CHANNEL_BUCKETS)) {
    if (values.has(channel)) return bucket;
  }
  return "其他";
}

function classifyTimeBucket(date) {
  const hm = [date.getHours(), date.getMinutes()];
  if ((compareHm(hm, WORK_TIME.start) >= 0 && compareHm(hm, WORK_TIME.lunchStart) < 0) || (compareHm(hm, WORK_TIME.lunchEnd) >= 0 && compareHm(hm, WORK_TIME.end) < 0)) {
    return "工作时间";
  }
  if (compareHm(hm, WORK_TIME.lunchStart) >= 0 && compareHm(hm, WORK_TIME.lunchEnd) < 0) {
    return "中午时间";
  }
  return "非工作时间";
}

function compareHm(left, right) {
  return left[0] === right[0] ? left[1] - right[1] : left[0] - right[0];
}

function splitReviewers(text) {
  return String(text || "")
    .split(/[，,、]/)
    .map((item) => item.trim())
    .filter((item) => REVIEWER_NAMES.includes(item));
}

function matchOutlet(region, dxms, creator, outletRules) {
  const regionRules = outletRules[region] || [];
  for (const rule of regionRules) {
    if (rule.regexes.some((regex) => regex.test(dxms))) return rule.name;
    if ((rule.creators || []).includes(creator)) return rule.name;
  }
  return null;
}

function bankForOutlet(outlet) {
  return BANK_NAME_BY_OUTLET[outlet] || null;
}

function totalForGroups(groupMap) {
  return SUMMARY_GROUP_ORDER.reduce((total, group) => total + (groupMap?.get(group) || 0), 0);
}

function safeRatio(numerator, denominator) {
  return denominator ? numerator / denominator : null;
}

function isFormulaValue(value) {
  if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "formula")) return true;
  if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "sharedFormula")) return true;
  if (typeof value === "string" && value.startsWith("=")) return true;
  return false;
}

function setCellBaseValue(worksheet, address, value) {
  const cell = worksheet.getCell(address);
  if (isFormulaValue(cell.value)) return;
  cell.value = value == null ? null : value;
}

function setCellBaseValueByRowCol(worksheet, row, col, value) {
  const cell = worksheet.getCell(row, col);
  if (isFormulaValue(cell.value)) return;
  cell.value = value == null ? null : value;
}

function setCellComputedValueByRowCol(worksheet, row, col, value) {
  const cell = worksheet.getCell(row, col);
  if (cell.value && typeof cell.value === "object") {
    if (Object.prototype.hasOwnProperty.call(cell.value, "formula") || Object.prototype.hasOwnProperty.call(cell.value, "sharedFormula")) {
      cell.value = { ...cell.value, result: value == null ? null : value };
      return;
    }
  }
  cell.value = value == null ? null : value;
}

function setCellPercentValueByRowCol(worksheet, row, col, value) {
  setCellBaseValueByRowCol(worksheet, row, col, value);
  worksheet.getCell(row, col).numFmt = "0.00%";
}

function shouldExcludeFromNonChannelReports(zbmbh) {
  return EXCLUDED_ZBMBH_FOR_NON_CHANNEL_REPORTS.has(String(zbmbh || "").trim());
}

function accumulateClassifiedRecord(data, bounds, {
  dt,
  monthKey,
  region,
  rule,
  outlet,
  bankName,
  channel,
  spr,
  excludeFromNonChannelReports
}) {
  if (monthKey === bounds.monthKey) {
    const channelBucket = classifyChannel(channel);
    getNestedCount(data.channelItemCounts, [rule.item, channelBucket]);
    getNestedCount(data.channelItemTimeCounts, [rule.item, classifyTimeBucket(dt)]);
  }

  if (excludeFromNonChannelReports) return;

  if (dt.getFullYear() === bounds.year) {
    data.monthlyAllTotals.set(monthKey, (data.monthlyAllTotals.get(monthKey) || 0) + 1);
    if (bankName) {
      if (!data.monthlyBankTotals.has(monthKey)) data.monthlyBankTotals.set(monthKey, new Map());
      const bucket = data.monthlyBankTotals.get(monthKey);
      bucket.set(bankName, (bucket.get(bankName) || 0) + 1);
    } else {
      data.monthlyCenterTotals.set(monthKey, (data.monthlyCenterTotals.get(monthKey) || 0) + 1);
    }
  }

  if (rule.summaryGroup && monthKey === bounds.monthKey) {
    const bucketOutlet = outlet || CENTER_NAME_BY_REGION[region];
    getNestedCount(data.regionMonthOutletGroup, [region, bucketOutlet, rule.summaryGroup]);
  }

  if (rule.summaryGroup && dt.getFullYear() === bounds.year) {
    const bucketOutlet = outlet || CENTER_NAME_BY_REGION[region];
    getNestedCount(data.regionYtdOutletGroup, [region, bucketOutlet, rule.summaryGroup]);
  }

  if (rule.summaryGroup && monthKey === bounds.prevMonthKey) {
    const bucketOutlet = outlet || CENTER_NAME_BY_REGION[region];
    getNestedCount(data.regionPrevOutletGroup, [region, bucketOutlet, rule.summaryGroup]);
  }

  if (monthKey === bounds.monthKey) {
    for (const reviewer of splitReviewers(spr)) {
      getNestedCount(data.reviewerItemCounts, [rule.item, reviewer]);
    }
  }
}

async function collectReportData({ month, dbConfig }) {
  const bounds = parseMonthBounds(month);
  const context = compileContext();
  const data = createEmptyData();
  const stats = createStats();
  const conn = await oracledb.getConnection(dbConfig);

  try {
    const sql = `
      select
        dx_05_dxcjsj,
        nvl(dx_05_dxmc, ' '),
        nvl(dx_05_dxms, ' '),
        trim(nvl(dx_05_dxcjqd, ' ')),
        nvl(dx_05_spr, ' '),
        nvl(dx_05_glb, ' '),
        nvl(dx_05_dxcjzmc, ' '),
        trim(nvl(dx_05_zbmbh, ' '))
      from PT_DXSL_2106032001_05
      where dx_05_dxcjsj >= :start_time
        and dx_05_dxcjsj < :end_time
    `;

    const result = await conn.execute(
      sql,
      { start_time: bounds.queryStart, end_time: bounds.nextMonth },
      { resultSet: true, outFormat: oracledb.OUT_FORMAT_ARRAY, fetchArraySize: 2000 }
    );

    const rs = result.resultSet;
    while (true) {
      const rows = await rs.getRows(2000);
      if (!rows.length) break;
      for (const row of rows) {
        stats.scannedRows += 1;
        const [dt, rawDxmc, rawDxms, rawChannel, rawSpr, rawGlb, rawCreator, rawZbmbh] = row;
        const dxmc = String(rawDxmc || "").trim();
        const dxms = String(rawDxms || "").trim();
        const channel = String(rawChannel || "").trim();
        const spr = String(rawSpr || "").trim();
        const glb = String(rawGlb || "").trim();
        const creator = String(rawCreator || "").trim();
        const zbmbh = String(rawZbmbh || "").trim();
        const region = REGION_NAME_BY_CODE[glb];
        if (!region || !(dt instanceof Date)) continue;

        const rule = classifyBusiness(dxmc, dxms, context.businessRules);
        if (!rule) {
          stats.unclassifiedRows += 1;
          continue;
        }
        stats.classifiedRows += 1;

        const outlet = matchOutlet(region, dxms, creator, context.outletRules);
        let bankName = null;
        if (outlet) {
          stats.matchedOutletRows += 1;
          bankName = bankForOutlet(outlet);
        } else {
          stats.unmatchedOutletRows += 1;
        }

        const monthKey = formatMonthKey(dt);
        accumulateClassifiedRecord(data, bounds, {
          dt,
          monthKey,
          region,
          rule,
          outlet,
          bankName,
          channel,
          spr,
          excludeFromNonChannelReports: shouldExcludeFromNonChannelReports(zbmbh)
        });
      }
    }
    await rs.close();
    return { bounds, data, stats };
  } finally {
    await conn.close();
  }
}

async function analyzeUnmatched({ month, dbConfig, limit = 30 }) {
  const bounds = parseMonthBounds(month);
  const context = compileContext();
  const conn = await oracledb.getConnection(dbConfig);
  const summary = {
    month,
    configPath: CONFIG_PATH,
    limit,
    scannedRows: 0,
    classifiedRows: 0,
    unmatchedRows: 0,
    unmatchedCreators: createTopCollector(limit),
    unmatchedDxms: createTopCollector(limit),
    unmatchedByRegion: createTopCollector(limit),
    unmatchedByItem: createTopCollector(limit)
  };

  try {
    const sql = `
      select
        dx_05_dxcjsj,
        nvl(dx_05_dxmc, ' '),
        nvl(dx_05_dxms, ' '),
        trim(nvl(dx_05_dxcjqd, ' ')),
        nvl(dx_05_glb, ' '),
        nvl(dx_05_dxcjzmc, ' ')
      from PT_DXSL_2106032001_05
      where dx_05_dxcjsj >= :start_time
        and dx_05_dxcjsj < :end_time
    `;
    const result = await conn.execute(
      sql,
      { start_time: bounds.queryStart, end_time: bounds.nextMonth },
      { resultSet: true, outFormat: oracledb.OUT_FORMAT_ARRAY, fetchArraySize: 2000 }
    );
    const rs = result.resultSet;
    while (true) {
      const rows = await rs.getRows(2000);
      if (!rows.length) break;
      for (const row of rows) {
        summary.scannedRows += 1;
        const [dt, rawDxmc, rawDxms, rawChannel, rawGlb, rawCreator] = row;
        if (!(dt instanceof Date)) continue;
        const dxmc = String(rawDxmc || "").trim();
        const dxms = String(rawDxms || "").trim();
        const channel = String(rawChannel || "").trim();
        const glb = String(rawGlb || "").trim();
        const creator = String(rawCreator || "").trim();
        const region = REGION_NAME_BY_CODE[glb];
        if (!region) continue;
        const rule = classifyBusiness(dxmc, dxms, context.businessRules);
        if (!rule) continue;
        summary.classifiedRows += 1;
        const outlet = matchOutlet(region, dxms, creator, context.outletRules);
        if (outlet) continue;
        summary.unmatchedRows += 1;
        collectTopItem(summary.unmatchedCreators, `${region} / ${creator || "(空)"}`, { region, creator, channel });
        collectTopItem(summary.unmatchedDxms, `${region} / ${dxms || "(空)"}`, { region, dxmc, channel });
        collectTopItem(summary.unmatchedByRegion, region, { region });
        collectTopItem(summary.unmatchedByItem, rule.item, { item: rule.item, region, channel });
      }
    }
    await rs.close();
    return {
      month: summary.month,
      configPath: summary.configPath,
      limit: summary.limit,
      scannedRows: summary.scannedRows,
      classifiedRows: summary.classifiedRows,
      unmatchedRows: summary.unmatchedRows,
      unmatchedCreators: topCollectorToArray(summary.unmatchedCreators),
      unmatchedDxms: topCollectorToArray(summary.unmatchedDxms),
      unmatchedByRegion: topCollectorToArray(summary.unmatchedByRegion),
      unmatchedByItem: topCollectorToArray(summary.unmatchedByItem)
    };
  } finally {
    await conn.close();
  }
}

function buildRowMap(worksheet, column) {
  const result = new Map();
  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const value = worksheet.getCell(`${column}${row}`).value;
    if (typeof value === "string" && value.trim()) result.set(value.trim(), row);
  }
  return result;
}

function buildRegionRowMap(worksheet) {
  const result = new Map();
  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const regionType = textFromCellValue(worksheet.getCell(`A${row}`).value).trim();
    const outletName = textFromCellValue(worksheet.getCell(`B${row}`).value).trim();
    if (!outletName) continue;
    if (regionType === "中心" || regionType === "银行网点" || regionType === "合计" || outletName === "小计") {
      result.set(outletName, row);
    }
  }
  return result;
}

function updateHeaderText(text, replacement) {
  return String(text).replace(/\d{4}年\d{1,2}月(?:-\d{1,2}月)?/g, replacement);
}

function formatReportDate(date = new Date()) {
  return `填报日期：${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

function updateReportDateCells(worksheet, date = new Date()) {
  const nextText = formatReportDate(date);
  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    for (let col = 1; col <= worksheet.columnCount; col += 1) {
      const cell = worksheet.getCell(row, col);
      const currentText = textFromCellValue(cell.value);
      if (!currentText.includes("填报时间：") && !currentText.includes("填报日期：")) continue;
      cell.value = currentText.replace(/填报(?:时间|日期)：\d{4}年\d{1,2}月\d{1,2}日/g, nextText);
    }
  }
}

function fillTotalSheet(worksheet, data, bounds, generatedAt = new Date()) {
  worksheet.getCell("A3").value = updateHeaderText(worksheet.getCell("A3").value, `${bounds.year}年1月-12月`);
  const rowMap = {
    "工商银行": 6,
    "农业银行": 9,
    "建设银行": 12,
    "丹东银行": 15,
    "邮储银行": 18,
    "大连银行": 21,
    "交通银行": 24,
    "中国银行": 27,
    "银行办件": 30,
    "中心办件": 33,
    "办件总数合计": 36
  };
  const monthKeys = Array.from({ length: 12 }, (_, index) => `${bounds.year}-${String(index + 1).padStart(2, "0")}`);
  const runningBank = Object.fromEntries(BANK_ORDER.map((bank) => [bank, 0]));
  let runningCenter = 0;

  for (let monthNumber = 1; monthNumber <= 12; monthNumber += 1) {
    const monthKey = monthKeys[monthNumber - 1];
    const col = 2 + monthNumber;
    const shouldFillMonth = monthNumber <= bounds.month;
    const bankBucket = data.monthlyBankTotals.get(monthKey) || new Map();
    const bankTotal = BANK_ORDER.reduce((sum, bank) => sum + (bankBucket.get(bank) || 0), 0);
    const totalValue = data.monthlyAllTotals.get(monthKey) || 0;
    const centerValue = data.monthlyCenterTotals.get(monthKey) || 0;
    const prevMonthKey = monthNumber > 1 ? monthKeys[monthNumber - 2] : null;
    const prevBankBucket = prevMonthKey ? data.monthlyBankTotals.get(prevMonthKey) || new Map() : new Map();
    const prevBankTotal = BANK_ORDER.reduce((sum, bank) => sum + (prevBankBucket.get(bank) || 0), 0);
    const prevCenterValue = prevMonthKey ? data.monthlyCenterTotals.get(prevMonthKey) || 0 : 0;

    for (const bank of BANK_ORDER) {
      const value = shouldFillMonth ? bankBucket.get(bank) || 0 : null;
      if (value != null) runningBank[bank] += value;
      setCellComputedValueByRowCol(worksheet, rowMap[bank], col, value);
      setCellComputedValueByRowCol(worksheet, rowMap[bank] + 1, col, shouldFillMonth ? safeRatio(value || 0, bankTotal) : null);
      setCellComputedValueByRowCol(worksheet, rowMap[bank] + 2, col, shouldFillMonth ? safeRatio(value || 0, totalValue) : null);
    }
    if (shouldFillMonth) runningCenter += centerValue;
    setCellComputedValueByRowCol(worksheet, rowMap["银行办件"], col, shouldFillMonth ? bankTotal : null);
    setCellComputedValueByRowCol(worksheet, rowMap["银行办件"] + 1, col, shouldFillMonth ? safeRatio(bankTotal, totalValue) : null);
    setCellComputedValueByRowCol(worksheet, rowMap["银行办件"] + 2, col, shouldFillMonth ? safeRatio(bankTotal - prevBankTotal, prevBankTotal) : null);
    setCellComputedValueByRowCol(worksheet, rowMap["中心办件"], col, shouldFillMonth ? centerValue : null);
    setCellComputedValueByRowCol(worksheet, rowMap["中心办件"] + 1, col, shouldFillMonth ? safeRatio(centerValue, totalValue) : null);
    setCellComputedValueByRowCol(worksheet, rowMap["中心办件"] + 2, col, shouldFillMonth ? safeRatio(centerValue - prevCenterValue, prevCenterValue) : null);
    setCellComputedValueByRowCol(worksheet, rowMap["办件总数合计"], col, shouldFillMonth ? totalValue : null);
  }

  const cumulativeBankTotal = BANK_ORDER.reduce((sum, bank) => sum + runningBank[bank], 0);
  const cumulativeTotal = cumulativeBankTotal + runningCenter;
  for (const bank of BANK_ORDER) {
    const total = runningBank[bank];
    setCellComputedValueByRowCol(worksheet, rowMap[bank], 15, total);
    setCellComputedValueByRowCol(worksheet, rowMap[bank] + 1, 15, safeRatio(total, cumulativeBankTotal));
    setCellComputedValueByRowCol(worksheet, rowMap[bank] + 2, 15, safeRatio(total, cumulativeTotal));
  }
  setCellComputedValueByRowCol(worksheet, rowMap["银行办件"], 15, cumulativeBankTotal);
  setCellComputedValueByRowCol(worksheet, rowMap["银行办件"] + 1, 15, safeRatio(cumulativeBankTotal, cumulativeTotal));
  setCellComputedValueByRowCol(worksheet, rowMap["银行办件"] + 2, 15, null);
  setCellComputedValueByRowCol(worksheet, rowMap["中心办件"], 15, runningCenter);
  setCellComputedValueByRowCol(worksheet, rowMap["中心办件"] + 1, 15, safeRatio(runningCenter, cumulativeTotal));
  setCellComputedValueByRowCol(worksheet, rowMap["中心办件"] + 2, 15, null);
  setCellComputedValueByRowCol(worksheet, rowMap["办件总数合计"], 15, cumulativeTotal);
  updateReportDateCells(worksheet, generatedAt);
}

function fillRegionSheet(worksheet, region, data, bounds, generatedAt = new Date()) {
  worksheet.getCell("A3").value = updateHeaderText(worksheet.getCell("A3").value, `${bounds.year}年${bounds.month}月`);
  const rowMap = buildRegionRowMap(worksheet);
  const current = data.regionMonthOutletGroup.get(region) || new Map();
  const ytd = data.regionYtdOutletGroup.get(region) || new Map();
  const previous = data.regionPrevOutletGroup.get(region) || new Map();
  const totalRegionCurrent = Array.from(current.values()).reduce((sum, groupMap) => sum + totalForGroups(groupMap), 0);
  const totalRegionYtd = Array.from(ytd.values()).reduce((sum, groupMap) => sum + totalForGroups(groupMap), 0);

  for (const [outletName, row] of rowMap.entries()) {
    if (outletName === "小计" || outletName === "合计") continue;
    const currentGroup = current.get(outletName) || new Map();
    const ytdGroup = ytd.get(outletName) || new Map();
    const prevGroup = previous.get(outletName) || new Map();
    SUMMARY_GROUP_ORDER.forEach((group, idx) => {
      setCellComputedValueByRowCol(worksheet, row, 3 + idx, currentGroup.get(group) || 0);
    });
    const currentTotal = totalForGroups(currentGroup);
    const previousTotal = totalForGroups(prevGroup);
    const ytdTotal = totalForGroups(ytdGroup);
    setCellComputedValueByRowCol(worksheet, row, 7, currentTotal);
    setCellComputedValueByRowCol(worksheet, row, 8, safeRatio(currentTotal, totalRegionCurrent));
    setCellComputedValueByRowCol(worksheet, row, 9, safeRatio(currentTotal - previousTotal, previousTotal));
    setCellComputedValueByRowCol(worksheet, row, 10, ytdTotal);
    setCellComputedValueByRowCol(worksheet, row, 11, safeRatio(ytdTotal, totalRegionYtd));
  }

  const centerName = CENTER_NAME_BY_REGION[region];
  const centerRow = rowMap.get(centerName);
  const subtotalRow = rowMap.get("小计");
  const totalRow = Array.from({ length: worksheet.rowCount }, (_, idx) => idx + 1).find((row) => textFromCellValue(worksheet.getCell(`A${row}`).value).trim() === "合计");

  if (subtotalRow) {
    const bankRows = [...rowMap.entries()]
      .filter(([name]) => name !== centerName && name !== "小计" && name !== "合计")
      .map(([, row]) => row);
    const subtotalCurrent = bankRows.reduce((sum, row) => sum + numericFromCellValue(worksheet.getCell(row, 7).value), 0);
    const subtotalPrevious = [...rowMap.entries()]
      .filter(([name]) => name !== centerName && name !== "小计" && name !== "合计")
      .reduce((sum, [name]) => sum + totalForGroups(previous.get(name) || new Map()), 0);
    const subtotalYtd = bankRows.reduce((sum, row) => sum + numericFromCellValue(worksheet.getCell(row, 10).value), 0);
    for (let col = 3; col <= 6; col += 1) {
      setCellComputedValueByRowCol(
        worksheet,
        subtotalRow,
        col,
        bankRows.reduce((sum, row) => sum + numericFromCellValue(worksheet.getCell(row, col).value), 0)
      );
    }
    setCellComputedValueByRowCol(worksheet, subtotalRow, 7, subtotalCurrent);
    setCellComputedValueByRowCol(worksheet, subtotalRow, 8, safeRatio(subtotalCurrent, totalRegionCurrent));
    setCellComputedValueByRowCol(worksheet, subtotalRow, 9, safeRatio(subtotalCurrent - subtotalPrevious, subtotalPrevious));
    setCellComputedValueByRowCol(worksheet, subtotalRow, 10, subtotalYtd);
    setCellComputedValueByRowCol(worksheet, subtotalRow, 11, safeRatio(subtotalYtd, totalRegionYtd));
  }

  if (totalRow && centerRow && subtotalRow) {
    for (let col = 3; col <= 7; col += 1) {
      setCellComputedValueByRowCol(
        worksheet,
        totalRow,
        col,
        numericFromCellValue(worksheet.getCell(centerRow, col).value) + numericFromCellValue(worksheet.getCell(subtotalRow, col).value)
      );
    }
    setCellComputedValueByRowCol(worksheet, totalRow, 8, null);
    setCellComputedValueByRowCol(worksheet, totalRow, 9, null);
    setCellComputedValueByRowCol(worksheet, totalRow, 10, numericFromCellValue(worksheet.getCell(centerRow, 10).value) + numericFromCellValue(worksheet.getCell(subtotalRow, 10).value));
    setCellComputedValueByRowCol(worksheet, totalRow, 11, null);
  }

  updateReportDateCells(worksheet, generatedAt);
}

function fillRankingSheet(worksheet, data, bounds, generatedAt = new Date()) {
  worksheet.getCell("A2").value = updateHeaderText(worksheet.getCell("A2").value, `${bounds.year}年${bounds.month}月`);
  setCellBaseValueByRowCol(worksheet, 3, 2, `${bounds.month}月办理业务数量`);
  setCellBaseValueByRowCol(worksheet, 3, 3, `${bounds.month}月办理业务数量`);
  const monthTotals = [];
  const ytdTotals = [];

  for (const [region, outlets] of data.regionMonthOutletGroup.entries()) {
    for (const [outlet, groupMap] of outlets.entries()) {
      if (outlet === CENTER_NAME_BY_REGION[region]) continue;
      monthTotals.push([outlet, totalForGroups(groupMap)]);
    }
  }
  for (const [region, outlets] of data.regionYtdOutletGroup.entries()) {
    for (const [outlet, groupMap] of outlets.entries()) {
      if (outlet === CENTER_NAME_BY_REGION[region]) continue;
      ytdTotals.push([outlet, totalForGroups(groupMap)]);
    }
  }

  const currentMap = new Map(monthTotals);
  const ytdMap = new Map(ytdTotals);
  const names = [...new Set([...currentMap.keys(), ...ytdMap.keys()])];
  const currentRank = new Map([...names].sort((a, b) => (currentMap.get(b) || 0) - (currentMap.get(a) || 0) || a.localeCompare(b)).map((name, index) => [name, index + 1]));
  const ytdRank = new Map([...names].sort((a, b) => (ytdMap.get(b) || 0) - (ytdMap.get(a) || 0) || a.localeCompare(b)).map((name, index) => [name, index + 1]));
  const ranked = [...names].sort(
    (a, b) =>
      (currentMap.get(b) || 0) - (currentMap.get(a) || 0) ||
      (ytdMap.get(b) || 0) - (ytdMap.get(a) || 0) ||
      a.localeCompare(b)
  );

  for (let row = 4; row <= 23; row += 1) {
    const name = ranked[row - 4];
    setCellBaseValueByRowCol(worksheet, row, 1, name || null);
    setCellBaseValueByRowCol(worksheet, row, 2, name ? currentMap.get(name) || 0 : null);
    setCellBaseValueByRowCol(worksheet, row, 4, name ? currentRank.get(name) || null : null);
    setCellBaseValueByRowCol(worksheet, row, 5, name ? ytdMap.get(name) || 0 : null);
    setCellBaseValueByRowCol(worksheet, row, 7, name ? ytdRank.get(name) || null : null);
  }
  updateReportDateCells(worksheet, generatedAt);
}

function fillShareSheet(worksheet, data, bounds, generatedAt = new Date()) {
  worksheet.getCell("A2").value = updateHeaderText(worksheet.getCell("A2").value, `${bounds.year}年${bounds.month}月`);
  const bankCounts = new Map(BANK_ORDER.map((bank) => [bank, new Map()]));
  const centerCounts = new Map();

  for (const [region, outlets] of data.regionMonthOutletGroup.entries()) {
    for (const [outlet, groupMap] of outlets.entries()) {
      if (outlet === CENTER_NAME_BY_REGION[region]) {
        for (const group of SUMMARY_GROUP_ORDER) {
          centerCounts.set(group, (centerCounts.get(group) || 0) + (groupMap.get(group) || 0));
        }
        continue;
      }
      const bank = bankForOutlet(outlet);
      if (!bank) continue;
      for (const group of SUMMARY_GROUP_ORDER) {
        const current = bankCounts.get(bank);
        current.set(group, (current.get(group) || 0) + (groupMap.get(group) || 0));
      }
    }
  }

  const rowMap = {
    "工商银行": 4,
    "农业银行": 7,
    "建设银行": 10,
    "丹东银行": 13,
    "邮储银行": 16,
    "大连银行": 19,
    "交通银行": 22,
    "中国银行": 25,
    "中心办件": 30
  };
  for (const bank of BANK_ORDER) {
    const groupMap = bankCounts.get(bank);
    SUMMARY_GROUP_ORDER.forEach((group, idx) => {
      setCellBaseValueByRowCol(worksheet, rowMap[bank], 3 + idx, groupMap.get(group) || 0);
    });
    setCellBaseValueByRowCol(worksheet, rowMap[bank], 7, totalForGroups(groupMap));
  }
  SUMMARY_GROUP_ORDER.forEach((group, idx) => {
    setCellBaseValueByRowCol(worksheet, rowMap["中心办件"], 3 + idx, centerCounts.get(group) || 0);
  });
  setCellBaseValueByRowCol(worksheet, rowMap["中心办件"], 7, totalForGroups(centerCounts));
  updateReportDateCells(worksheet, generatedAt);
}

function fillChannelSheet(worksheet, data, bounds) {
  worksheet.getCell("K2").value = `查询日期：${bounds.year}年${bounds.month}月`;
  for (let row = 5; row <= worksheet.rowCount; row += 1) {
    const item = worksheet.getCell(`B${row}`).value;
    if (typeof item !== "string" || !item.trim()) continue;
    const itemName = item.trim();
    const channelMap = data.channelItemCounts.get(itemName) || new Map();
    const timeMap = data.channelItemTimeCounts.get(itemName) || new Map();
    CHANNEL_BUCKET_ORDER.forEach((bucket, index) => {
      setCellBaseValueByRowCol(worksheet, row, 3 + index, channelMap.get(bucket) || 0);
    });
    setCellBaseValueByRowCol(worksheet, row, 10, sumMapValues(channelMap));
    setCellBaseValueByRowCol(worksheet, row, 11, timeMap.get("工作时间") || 0);
    setCellBaseValueByRowCol(worksheet, row, 12, timeMap.get("中午时间") || 0);
    setCellBaseValueByRowCol(worksheet, row, 13, timeMap.get("非工作时间") || 0);
  }
  const totalRow = 44;
  for (let col = 3; col <= 13; col += 1) {
    const total = Array.from({ length: totalRow - 5 }, (_, index) => 5 + index)
      .reduce((sum, row) => sum + (Number(worksheet.getCell(row, col).value) || 0), 0);
    setCellBaseValueByRowCol(worksheet, totalRow, col, total);
  }
}

function fillReviewerSheet(worksheet, data) {
  const reviewerColumns = new Map(REVIEWER_NAMES.map((name, index) => [name, 3 + index]));
  for (let row = 3; row <= 49; row += 1) {
    const item = worksheet.getCell(`B${row}`).value;
    if (typeof item !== "string" || !item.trim()) continue;
    const itemName = item.trim();
    const reviewerMap = data.reviewerItemCounts.get(itemName) || new Map();
    let total = 0;
    for (const [reviewer, col] of reviewerColumns.entries()) {
      const value = reviewerMap.get(reviewer) || 0;
      total += value;
      setCellBaseValueByRowCol(worksheet, row, col, value);
    }
    setCellBaseValueByRowCol(worksheet, row, 9, total);
  }

  const totalRows = [
    [3, 19, 20, 21],
    [22, 33, 34, 35],
    [36, 36, 37, 38],
    [39, 43, 44, 45],
    [46, 49, 50, 51]
  ];
  for (const [start, end, totalRow, ratioRow] of totalRows) {
    for (let col = 3; col <= 9; col += 1) {
      let total = 0;
      for (let row = start; row <= end; row += 1) total += worksheet.getCell(row, col).value || 0;
      setCellBaseValueByRowCol(worksheet, totalRow, col, total);
    }
    const groupTotal = Number(worksheet.getCell(totalRow, 9).value) || 0;
    for (let col = 3; col <= 9; col += 1) {
      setCellPercentValueByRowCol(worksheet, ratioRow, col, safeRatio(Number(worksheet.getCell(totalRow, col).value) || 0, groupTotal));
    }
  }
  for (let col = 3; col <= 9; col += 1) {
    const total = [20, 34, 37, 44, 50].reduce((sum, row) => sum + (Number(worksheet.getCell(row, col).value) || 0), 0);
    setCellBaseValueByRowCol(worksheet, 52, col, total);
  }
  const grandTotal = Number(worksheet.getCell(52, 9).value) || 0;
  for (let col = 3; col <= 9; col += 1) {
    setCellPercentValueByRowCol(worksheet, 53, col, safeRatio(Number(worksheet.getCell(52, col).value) || 0, grandTotal));
  }
}

function maybeFillSheet(workbook, sheetName, fillFn) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new Error(`sheet not found: ${sheetName}`);
  fillFn(worksheet);
}

function cloneValue(value) {
  if (value == null) return value;
  if (value instanceof Date) return new Date(value.getTime());
  if (Array.isArray(value)) return value.map(cloneValue);
  if (typeof value === "object") return JSON.parse(JSON.stringify(value));
  return value;
}

function wrapFormulaWithIfError(formula) {
  const text = String(formula || "").trim();
  if (!text) return text;
  if (/^IFERROR\s*\(/i.test(text)) return text;
  return `IFERROR(${text},"")`;
}

function sanitizeWorksheetFormulaErrors(worksheet) {
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value;
      if (!value || typeof value !== "object") return;
      if (Object.prototype.hasOwnProperty.call(value, "formula")) {
        cell.value = {
          ...value,
          formula: wrapFormulaWithIfError(value.formula)
        };
      }
    });
  });
}

function cloneWorksheetToWorkbook(sourceSheet, targetWorkbook) {
  const targetSheet = targetWorkbook.addWorksheet(sourceSheet.name, {
    views: sourceSheet.views,
    properties: sourceSheet.properties,
    pageSetup: sourceSheet.pageSetup,
    headerFooter: sourceSheet.headerFooter,
    state: sourceSheet.state
  });

  sourceSheet.columns.forEach((column, index) => {
    const targetColumn = targetSheet.getColumn(index + 1);
    targetColumn.width = column.width;
    targetColumn.hidden = column.hidden;
    targetColumn.outlineLevel = column.outlineLevel;
    targetColumn.style = JSON.parse(JSON.stringify(column.style || {}));
  });

  sourceSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const targetRow = targetSheet.getRow(rowNumber);
    targetRow.height = row.height;
    targetRow.hidden = row.hidden;
    targetRow.outlineLevel = row.outlineLevel;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber);
      targetCell.value = cloneValue(cell.value);
      targetCell.style = JSON.parse(JSON.stringify(cell.style || {}));
      if (cell.numFmt) targetCell.numFmt = cell.numFmt;
      if (cell.note) targetCell.note = cloneValue(cell.note);
    });
    targetRow.commit();
  });

  const merges = sourceSheet.model?.merges || [];
  merges.forEach((range) => targetSheet.mergeCells(range));
  return targetSheet;
}

async function writeReportFile({ templatePath, outputPath, report, sheetName, sheetOnly }) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  workbook.calcProperties.fullCalcOnLoad = true;
  const generatedAt = new Date();
  const requestedSheets = sheetName ? new Set([sheetName]) : new Set(REPORT_SHEETS);

  if (requestedSheets.has("总表")) maybeFillSheet(workbook, "总表", (ws) => fillTotalSheet(ws, report.data, report.bounds, generatedAt));
  if (requestedSheets.has("市本级")) maybeFillSheet(workbook, "市本级", (ws) => fillRegionSheet(ws, "市本级", report.data, report.bounds, generatedAt));
  if (requestedSheets.has("东港")) maybeFillSheet(workbook, "东港", (ws) => fillRegionSheet(ws, "东港", report.data, report.bounds, generatedAt));
  if (requestedSheets.has("凤城")) maybeFillSheet(workbook, "凤城", (ws) => fillRegionSheet(ws, "凤城", report.data, report.bounds, generatedAt));
  if (requestedSheets.has("宽甸")) maybeFillSheet(workbook, "宽甸", (ws) => fillRegionSheet(ws, "宽甸", report.data, report.bounds, generatedAt));
  if (requestedSheets.has("排名")) maybeFillSheet(workbook, "排名", (ws) => fillRankingSheet(ws, report.data, report.bounds, generatedAt));
  if (requestedSheets.has("占比")) maybeFillSheet(workbook, "占比", (ws) => fillShareSheet(ws, report.data, report.bounds, generatedAt));
  if (requestedSheets.has("各渠道复核业务来源统计")) maybeFillSheet(workbook, "各渠道复核业务来源统计", (ws) => fillChannelSheet(ws, report.data, report.bounds));
  if (requestedSheets.has(" 复核业务量统计")) maybeFillSheet(workbook, " 复核业务量统计", (ws) => fillReviewerSheet(ws, report.data));
  workbook.eachSheet((worksheet) => sanitizeWorksheetFormulaErrors(worksheet));

  if (sheetOnly) {
    if (!sheetName) throw new Error("sheetOnly requires sheetName");
    const selected = workbook.getWorksheet(sheetName);
    if (!selected) throw new Error(`sheet not found: ${sheetName}`);
    const singleWorkbook = new ExcelJS.Workbook();
    singleWorkbook.calcProperties.fullCalcOnLoad = true;
    cloneWorksheetToWorkbook(selected, singleWorkbook);
    await ensureParentDir(outputPath);
    await singleWorkbook.xlsx.writeFile(outputPath);
    return;
  }

  await ensureParentDir(outputPath);
  await workbook.xlsx.writeFile(outputPath);
}

async function ensureParentDir(filePath) {
  await fs.promises.mkdir(path.dirname(filePath), { recursive: true });
}

function buildDefaultOutputPath({ templatePath, outputDir, month, sheetName, sheetOnly }) {
  const parsed = path.parse(templatePath);
  const safeSheet = sheetName ? `_${sheetName.replace(/[\\/:*?"<>|]/g, "_").trim()}` : "";
  const suffix = sheetOnly ? "_single" : "";
  const dir = outputDir || path.join(parsed.dir, "generated");
  return path.join(dir, `${parsed.name}_${month}${safeSheet}${suffix}${parsed.ext}`);
}

async function generateReport(options) {
  const templatePath = options.templatePath || path.join(process.cwd(), "模板.xlsx");
  const outputPath = options.outputPath || buildDefaultOutputPath(options);
  const dbConfig = {
    user: options.user,
    password: options.password,
    connectString: options.dsn
  };
  const report = await collectReportData({ month: options.month, dbConfig });
  await writeReportFile({
    templatePath,
    outputPath,
    report,
    sheetName: options.sheetName,
    sheetOnly: options.sheetOnly
  });
  return {
    outputPath,
    month: options.month,
    sheetName: options.sheetName || null,
    sheetOnly: Boolean(options.sheetOnly),
    stats: report.stats
  };
}

module.exports = {
  REPORT_SHEETS,
  CONFIG_PATH,
  analyzeUnmatched,
  buildDefaultOutputPath,
  collectReportData,
  generateReport
};

module.exports.__test__ = {
  fillChannelSheet,
  fillRegionSheet,
  fillRankingSheet,
  fillReviewerSheet,
  fillTotalSheet,
  sanitizeWorksheetFormulaErrors,
  accumulateClassifiedRecord,
  compileContext,
  formatReportDate,
  matchOutlet,
  textFromCellValue,
  numericFromCellValue,
  buildRegionRowMap,
  shouldExcludeFromNonChannelReports,
  wrapFormulaWithIfError
};
