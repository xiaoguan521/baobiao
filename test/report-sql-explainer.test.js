const test = require("node:test");
const assert = require("node:assert/strict");
const { explainFieldSql } = require("../src/report-sql-explainer");

test("region sheet monthly group field returns aggregate sql", () => {
  const result = explainFieldSql({
    sheetName: "市本级",
    fieldName: "归集",
    rowName: "工行锦山支行"
  });

  assert.equal(result.supported, true);
  assert.equal(result.sqlKind, "aggregate");
  assert.match(result.verificationSql, /count\(\*\) as field_value/);
  assert.match(result.verificationSql, /dx_05_dxcjsj >= :month_start/);
  assert.match(result.verificationSql, /regexp_like\(nvl\(dx_05_dxms, ' '\), '工商银行\.\*锦山支行'\)/);
});

test("region sheet ratio field returns derived sql", () => {
  const result = explainFieldSql({
    sheetName: "东港",
    fieldName: "比上月增减",
    rowName: "东港分部"
  });

  assert.equal(result.supported, true);
  assert.equal(result.sqlKind, "derived");
  assert.match(result.explanation, /不是单条原始 SQL 字段/);
  assert.match(result.verificationSql, /with current_month as/);
  assert.match(result.verificationSql, /previous_month/);
});

test("reviewer sheet supports reviewer totals", () => {
  const result = explainFieldSql({
    sheetName: "复核业务量统计",
    fieldName: "李俊昌",
    rowName: "房租提取"
  });

  assert.equal(result.supported, true);
  assert.equal(result.sqlKind, "aggregate");
  assert.match(result.verificationSql, /李俊昌/);
  assert.match(result.explanation, /被该复核人处理的数量/);
});
