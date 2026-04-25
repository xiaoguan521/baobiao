const test = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");
const { __test__ } = require("../src/report-engine");

test("unit-level patterns fall back to center outlet", () => {
  const context = __test__.compileContext();

  assert.equal(
    __test__.matchOutlet("东港", "东港市农业综合行政执法队(02089497)缴存人基数调整。（来自网厅）", "", context.outletRules),
    "东港分部"
  );

  assert.equal(
    __test__.matchOutlet("市本级", "中国建设银行股份有限公司丹东分行(01012286)缴存人基数调整。（来自网厅）", "", context.outletRules),
    "公积金大厅"
  );
});

test("bank keyword rules still win before center fallback", () => {
  const context = __test__.compileContext();

  assert.equal(
    __test__.matchOutlet("东港", "工商银行东港支行办理贷款业务", "", context.outletRules),
    "工行东港支行"
  );
});

test("region row map ignores merged title rows", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("市本级");
  worksheet.mergeCells("A2:L2");
  worksheet.mergeCells("A3:L3");
  worksheet.getCell("A2").value = "市本级住房公积金柜台业务办理情况统计表";
  worksheet.getCell("A3").value = "统计月度：2025年12月";
  worksheet.getCell("A6").value = "中心";
  worksheet.getCell("B6").value = "公积金大厅";
  worksheet.getCell("A7").value = "银行网点";
  worksheet.getCell("B7").value = "工行锦山支行";
  worksheet.getCell("A8").value = "银行网点";
  worksheet.getCell("B8").value = "小计";

  const rowMap = __test__.buildRegionRowMap(worksheet);

  assert.equal(rowMap.has("市本级住房公积金柜台业务办理情况统计表"), false);
  assert.equal(rowMap.has("统计月度：2025年12月"), false);
  assert.equal(rowMap.get("公积金大厅"), 6);
  assert.equal(rowMap.get("工行锦山支行"), 7);
  assert.equal(rowMap.get("小计"), 8);
});

test("fill total sheet leaves future months blank and refreshes report date", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("总表");
  worksheet.getCell("A3").value = "统计月度：2025年1月-12月";
  worksheet.getCell("L40").value = "填报时间：2026年1月5日";

  const data = {
    monthlyBankTotals: new Map([
      ["2026-01", new Map([["工商银行", 11]])],
      ["2026-02", new Map([["工商银行", 12]])],
      ["2026-03", new Map([["工商银行", 13]])],
      ["2026-04", new Map([["工商银行", 14]])],
      ["2026-05", new Map([["工商银行", 99]])]
    ]),
    monthlyAllTotals: new Map([
      ["2026-01", 12],
      ["2026-02", 14],
      ["2026-03", 16],
      ["2026-04", 18],
      ["2026-05", 187]
    ]),
    monthlyCenterTotals: new Map([
      ["2026-01", 1],
      ["2026-02", 2],
      ["2026-03", 3],
      ["2026-04", 4],
      ["2026-05", 88]
    ])
  };

  __test__.fillTotalSheet(worksheet, data, { year: 2026, month: 4 }, new Date("2026-04-24T00:00:00Z"));

  assert.equal(worksheet.getCell("A3").value, "统计月度：2026年1月-12月");
  assert.equal(worksheet.getCell("C6").value, 11);
  assert.equal(worksheet.getCell("F6").value, 14);
  assert.equal(worksheet.getCell("G6").value, null);
  assert.equal(worksheet.getCell("O6").value, 50);
  assert.equal(worksheet.getCell("F30").value, 14);
  assert.equal(worksheet.getCell("G30").value, null);
  assert.equal(worksheet.getCell("O30").value, 50);
  assert.equal(worksheet.getCell("F33").value, 4);
  assert.equal(worksheet.getCell("G33").value, null);
  assert.equal(worksheet.getCell("O33").value, 10);
  assert.equal(worksheet.getCell("O36").value, 60);
  assert.equal(worksheet.getCell("L40").value, "填报日期：2026年4月24日");
});

test("fill reviewer sheet populates totals and ratios", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("复核");
  worksheet.getCell("B3").value = "单位住房公积金缴存登记";
  worksheet.getCell("B20").value = "归集业务总量";
  worksheet.getCell("B21").value = "占同类型业务比例";
  worksheet.getCell("B52").value = "业务量总计";
  worksheet.getCell("B53").value = "占总业务量比例";

  const data = {
    reviewerItemCounts: new Map([
      ["单位住房公积金缴存登记", new Map([["李俊昌", 2], ["李阳", 1]])]
    ])
  };

  __test__.fillReviewerSheet(worksheet, data);

  assert.equal(worksheet.getCell("C3").value, 2);
  assert.equal(worksheet.getCell("D3").value, 1);
  assert.equal(worksheet.getCell("I3").value, 3);
  assert.equal(worksheet.getCell("C20").value, 2);
  assert.equal(worksheet.getCell("D20").value, 1);
  assert.equal(worksheet.getCell("I20").value, 3);
  assert.equal(worksheet.getCell("C21").value, 2 / 3);
  assert.equal(worksheet.getCell("D21").value, 1 / 3);
  assert.equal(worksheet.getCell("I21").value, 1);
  assert.equal(worksheet.getCell("I52").value, 3);
  assert.equal(worksheet.getCell("C53").value, 2 / 3);
});

test("fill channel sheet populates total row", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("各渠道");
  worksheet.getCell("B5").value = "单位住房公积金缴存登记";
  worksheet.getCell("B6").value = "军转干缴存开户";

  const data = {
    channelItemCounts: new Map([
      ["单位住房公积金缴存登记", new Map([["柜面", 2], ["网厅", 3]])],
      ["军转干缴存开户", new Map([["柜面", 1], ["其他", 4]])]
    ]),
    channelItemTimeCounts: new Map([
      ["单位住房公积金缴存登记", new Map([["工作时间", 4], ["中午时间", 1]])],
      ["军转干缴存开户", new Map([["非工作时间", 5]])]
    ])
  };

  __test__.fillChannelSheet(worksheet, data, { year: 2026, month: 4 });

  assert.equal(worksheet.getCell("K2").value, "查询日期：2026年4月");
  assert.equal(worksheet.getCell("C5").value, 2);
  assert.equal(worksheet.getCell("F5").value, 3);
  assert.equal(worksheet.getCell("J5").value, 5);
  assert.equal(worksheet.getCell("C44").value, 3);
  assert.equal(worksheet.getCell("I44").value, 4);
  assert.equal(worksheet.getCell("J44").value, 10);
  assert.equal(worksheet.getCell("M44").value, 5);
});

test("fill region sheet refreshes subtotal and total cached values", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("东港");
  const rows = [
    [6, "中心", "东港分部"],
    [7, "银行网点", "工行东港支行"],
    [8, "银行网点", "农行东港支行"],
    [9, "银行网点", "建行东港支行"],
    [10, "银行网点", "小计"],
    [11, "合计", null]
  ];
  rows.forEach(([row, a, b]) => {
    worksheet.getCell(`A${row}`).value = a;
    if (b != null) worksheet.getCell(`B${row}`).value = b;
  });
  worksheet.getCell("J12").value = "填报日期：2026年1月5日";
  worksheet.getCell("G10").value = { formula: "SUM(G7:G9)", result: 338 };
  worksheet.getCell("J10").value = { formula: "SUM(J7:J9)", result: 2746 };
  worksheet.getCell("K10").value = { formula: "J10/J11", result: 0.68 };
  worksheet.getCell("C11").value = { formula: "C6+C10", result: 99 };
  worksheet.getCell("J11").value = 3999;

  const toGroupMap = (obj) => new Map(Object.entries(obj));
  const data = {
    regionMonthOutletGroup: new Map([
      ["东港", new Map([
        ["东港分部", toGroupMap({ 归集: 10, 提取: 20, 贷款: 0, 贷后: 0 })],
        ["工行东港支行", toGroupMap({ 归集: 1, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["农行东港支行", toGroupMap({ 归集: 2, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["建行东港支行", toGroupMap({ 归集: 3, 提取: 0, 贷款: 0, 贷后: 0 })]
      ])]
    ]),
    regionYtdOutletGroup: new Map([
      ["东港", new Map([
        ["东港分部", toGroupMap({ 归集: 100, 提取: 100, 贷款: 0, 贷后: 0 })],
        ["工行东港支行", toGroupMap({ 归集: 10, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["农行东港支行", toGroupMap({ 归集: 20, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["建行东港支行", toGroupMap({ 归集: 30, 提取: 0, 贷款: 0, 贷后: 0 })]
      ])]
    ]),
    regionPrevOutletGroup: new Map([
      ["东港", new Map([
        ["东港分部", toGroupMap({ 归集: 8, 提取: 12, 贷款: 0, 贷后: 0 })],
        ["工行东港支行", toGroupMap({ 归集: 1, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["农行东港支行", toGroupMap({ 归集: 1, 提取: 0, 贷款: 0, 贷后: 0 })],
        ["建行东港支行", toGroupMap({ 归集: 1, 提取: 0, 贷款: 0, 贷后: 0 })]
      ])]
    ])
  };

  __test__.fillRegionSheet(worksheet, "东港", data, { year: 2026, month: 4 }, new Date("2026-04-25T00:00:00Z"));

  assert.equal(worksheet.getCell("G10").value.result, 6);
  assert.equal(worksheet.getCell("J10").value.result, 60);
  assert.equal(worksheet.getCell("K10").value.result, 60 / 260);
  assert.equal(worksheet.getCell("C11").value.result, 16);
  assert.equal(worksheet.getCell("J11").value, 260);
  assert.equal(worksheet.getCell("J12").value, "填报日期：2026年4月25日");
});

test("fill ranking sheet updates month column headers", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("排名");
  worksheet.getCell("A2").value = "统计月度：2025年12月";
  worksheet.getCell("B3").value = "12月办理业务数量";
  worksheet.getCell("C3").value = "12月办理业务数量";

  const data = {
    regionMonthOutletGroup: new Map(),
    regionYtdOutletGroup: new Map()
  };

  __test__.fillRankingSheet(worksheet, data, { year: 2026, month: 3 }, new Date("2026-04-25T00:00:00Z"));

  assert.equal(worksheet.getCell("A2").value, "统计月度：2026年3月");
  assert.equal(worksheet.getCell("B3").value, "3月办理业务数量");
  assert.equal(worksheet.getCell("C3").value, "3月办理业务数量");
});

test("wrap formula with IFERROR only once", () => {
  assert.equal(__test__.wrapFormulaWithIfError("A1/B1"), 'IFERROR(A1/B1,"")');
  assert.equal(__test__.wrapFormulaWithIfError('IFERROR(A1/B1,"")'), 'IFERROR(A1/B1,"")');
});

test("sanitize worksheet formula errors wraps formula cells", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("总表");
  worksheet.getCell("A1").value = { formula: "A2/B2", result: { error: "#DIV/0!" } };
  worksheet.getCell("A2").value = 1;
  worksheet.getCell("B2").value = 0;

  __test__.sanitizeWorksheetFormulaErrors(worksheet);

  assert.equal(worksheet.getCell("A1").value.formula, 'IFERROR(A2/B2,"")');
});
