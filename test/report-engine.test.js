const test = require("node:test");
const assert = require("node:assert/strict");
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
