const NON_CHART_SHEETS = new Set(["填写说明", "效率提升", "库存周转分析", "总结"]);
const NON_KPI_SHEETS = new Set(["填写说明", "库存周转分析"]);
const WRAP_COLUMNS = new Set(["本周进展", "填写要求", "备注说明"]);
const OVERVIEW_GROUPS = [
  { sheet: "采购数据统计", title: "采购数据统计", fields: ["采购合同个数", "采购单个数"] },
  { sheet: "Key-SKU断货统计", title: "Key-SKU断货统计", fields: ["货号个数", "断货个数", "断货率"] },
  { sheet: "议价数据统计", title: "议价数据统计", fields: [["议价合计（元）", "议价合计"]] },
  { sheet: "订单数据", title: "订单数据", fields: [["新下单订单金额", "新下单金额"], ["待审核总订单金额", "待审核订单金额", "待审核总订单金额"]] }
];

const demoWorkbook = {
  workbookName: "周维度经营看板上传模板-new.xlsx",
  tables: [
    {
      id: "purchase-stats",
      sheet: "采购数据统计",
      title: "采购数据统计",
      headers: ["周次", "采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）", "金额周环比"],
      weekField: "周次",
      numericFields: ["采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）", "金额周环比"],
      rows: [
        { "周次": "07-23", "采购合同个数": 1039, "采购单个数": 2486, "销售采购单个数": 1169, "直发采购单个数": 521, "直发采购单比例": 0.446, "采购金额（元）": 4141198.82, "人数": 15, "人均采购金额（元）": 276079.92, "人均采购单（个）": 165.7, "金额周环比": null },
        { "周次": "07-30", "采购合同个数": 1169, "采购单个数": 3698, "销售采购单个数": 1138, "直发采购单个数": 515, "直发采购单比例": 0.453, "采购金额（元）": 5760008.07, "人数": 15, "人均采购金额（元）": 384000.54, "人均采购单（个）": 246.5, "金额周环比": 0.391 }
      ]
    },
    {
      id: "sku-break",
      sheet: "Key-SKU断货统计",
      title: "Key-SKU断货统计",
      headers: ["周次", "货号个数", "断货个数", "断货率", "备注说明"],
      weekField: "周次",
      numericFields: ["货号个数", "断货个数", "断货率"],
      rows: [
        { "周次": "07-23", "货号个数": 2264, "断货个数": 48, "断货率": 0.0212, "备注说明": "" },
        { "周次": "07-30", "货号个数": 2248, "断货个数": 69, "断货率": 0.0307, "备注说明": "阈值调整" }
      ]
    },
    {
      id: "bargain-stats",
      sheet: "议价数据统计",
      title: "议价数据统计",
      headers: ["周次", "单笔议价金额（元）", "备货议价金额（元）", "议价合计（元）"],
      weekField: "周次",
      numericFields: ["单笔议价金额（元）", "备货议价金额（元）", "议价合计（元）"],
      rows: [
        { "周次": "07-23", "单笔议价金额（元）": 4971.97, "备货议价金额（元）": 56502.88, "议价合计（元）": 61474.85 },
        { "周次": "07-30", "单笔议价金额（元）": 5210.54, "备货议价金额（元）": 25998.03, "议价合计（元）": 31208.57 }
      ]
    },
    {
      id: "warehouse-output",
      sheet: "仓储单位产出",
      title: "仓储单位产出",
      headers: ["月份", "主仓单位面积销售额", "廊坊仓单位面积销售额", "广州仓单位面积销售额", "成都仓单位面积销售额", "主仓单位面积毛利", "廊坊仓单位面积毛利", "广州仓单位面积毛利", "成都仓单位面积毛利"],
      weekField: null,
      numericFields: ["主仓单位面积销售额", "廊坊仓单位面积销售额", "广州仓单位面积销售额", "成都仓单位面积销售额", "主仓单位面积毛利", "廊坊仓单位面积毛利", "广州仓单位面积毛利", "成都仓单位面积毛利"],
      rows: [
        { "月份": "2026-06", "主仓单位面积销售额": 1814, "廊坊仓单位面积销售额": 1092, "广州仓单位面积销售额": 519, "成都仓单位面积销售额": 558, "主仓单位面积毛利": 544, "廊坊仓单位面积毛利": 290, "广州仓单位面积毛利": 184, "成都仓单位面积毛利": 198 }
      ]
    },
    {
      id: "efficiency",
      sheet: "效率提升",
      title: "效率提升",
      headers: ["周次", "模块", "本周进展"],
      weekField: "周次",
      numericFields: [],
      rows: [
        { "周次": "07-23", "模块": "AI", "本周进展": "知识库：产品资料投喂；发票命名自动化" },
        { "周次": "07-30", "模块": "流程优化", "本周进展": "年度协议管理打通；标签自动流转模块沟通" },
        { "周次": "08-06", "模块": "AI", "本周进展": null }
      ]
    },
    {
      id: "inventory-turnover",
      sheet: "库存周转分析",
      title: "库存周转分析",
      headers: ["仓库/品类", "6月库存金额", "6月年营业成本", "6月周转次数", "6月平均周转天数", "7月库存金额", "7月年营业成本", "7月周转次数", "7月平均周转天数", "库存金额环比变化", "周转天数变化(天）"],
      weekField: null,
      numericFields: ["6月库存金额", "6月年营业成本", "6月周转次数", "6月平均周转天数", "7月库存金额", "7月年营业成本", "7月周转次数", "7月平均周转天数", "库存金额环比变化", "周转天数变化(天）"],
      rows: [
        { "仓库/品类": "主仓有阈值的明星产品", "6月库存金额": 12960913.76, "6月年营业成本": 59780928.78, "6月周转次数": 4.61, "6月平均周转天数": 79.1, "7月库存金额": 14129745.66, "7月年营业成本": 57251892.44, "7月周转次数": 4.05, "7月平均周转天数": 90.1, "库存金额环比变化": 0.09, "周转天数变化(天）": 11 },
        { "仓库/品类": "主仓所有商品", "6月库存金额": 32900585.42, "6月年营业成本": 134744245.49, "6月周转次数": 4.1, "6月平均周转天数": 89.1, "7月库存金额": 34198877.49, "7月年营业成本": 131835427.48, "7月周转次数": 3.85, "7月平均周转天数": 94.7, "库存金额环比变化": 0.039, "周转天数变化(天）": 5.6 }
      ]
    },
    {
      id: "instructions",
      sheet: "填写说明",
      title: "填写说明",
      headers: ["说明项", "填写要求"],
      weekField: null,
      numericFields: [],
      rows: [
        { "说明项": "周次格式", "填写要求": "建议填写 07-30、2026-W31 或 本周日期" },
        { "说明项": "首行", "填写要求": "每张表第一行写表头，不要留空" },
        { "说明项": "金额字段", "填写要求": "填写纯数字，不要手动加逗号" },
        { "说明项": "比例字段", "填写要求": "填写 0.453 或 45.3% 都可以" },
        { "说明项": "备注字段", "填写要求": "可为空，用于展示运营说明" }
      ]
    }
  ]
};

demoWorkbook.workbookName = "周维度经营看板上传模板.xlsx";
demoWorkbook.tables = [
  {
      id: "order-stats",
      sheet: "订单数据",
      title: "订单数据-1",
      headers: ["周次", "新下单订单条数", "新下单订单金额", "待审核总订单条数", "待审核总订单金额", "25年及以前待审核订单条数", "25年及以前待审核订单金额", "26年1-6月待审核订单条数", "26年1-6月待审核订单金额", "环比"],
      weekField: "周次",
    numericFields: ["新下单订单条数", "新下单订单金额", "待审核总订单条数", "待审核总订单金额", "25年及以前待审核订单条数", "25年及以前待审核订单金额", "26年1-6月待审核订单条数", "26年1-6月待审核订单金额", "环比"],
      rows: [
        { "周次": "07-30", "新下单订单条数": 5191, "新下单订单金额": 6571276.82, "待审核总订单条数": 1946, "待审核总订单金额": 10539098.29, "25年及以前待审核订单条数": 102, "25年及以前待审核订单金额": 459589.47, "26年1-6月待审核订单条数": 431, "26年1-6月待审核订单金额": 4389773.64, "环比": null },
      { "周次": "08-06", "新下单订单条数": 4684, "新下单订单金额": 8243621.3, "待审核总订单条数": 1917, "待审核总订单金额": 11528192.33, "25年及以前待审核订单条数": 98, "25年及以前待审核订单金额": 455024.27, "26年1-6月待审核订单条数": 356, "26年1-6月待审核订单金额": 3925227.65, "环比": 0.254493080387382 }
      ]
    },
  {
    id: "purchase-stats",
    sheet: "采购数据统计",
    title: "采购数据统计-1",
    headers: ["周次", "采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）", "金额周环比"],
    weekField: "周次",
    numericFields: ["采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）", "金额周环比"],
    rows: [
      { "周次": "07-23", "采购合同个数": 1039, "采购单个数": 2486, "销售采购单个数": 1169, "直发采购单个数": 521, "直发采购单比例": 0.446, "采购金额（元）": 4141198.82, "人数": 15, "人均采购金额（元）": 276079.92, "人均采购单（个）": 165.7, "金额周环比": null },
      { "周次": "07-30", "采购合同个数": 1169, "采购单个数": 3698, "销售采购单个数": 1138, "直发采购单个数": 515, "直发采购单比例": 0.453, "采购金额（元）": 5760008.07, "人数": 15, "人均采购金额（元）": 384000.54, "人均采购单（个）": 246.5, "金额周环比": 0.391 },
      { "周次": "08-07", "采购合同个数": 845, "采购单个数": 2042, "销售采购单个数": 893, "直发采购单个数": 448, "直发采购单比例": 0.501679731243001, "采购金额（元）": 4556368.56, "人数": 15, "人均采购金额（元）": 303757.904, "人均采购单（个）": 136.133333333333, "金额周环比": -0.208964899939802 }
    ]
  },
  {
    id: "sku-break",
    sheet: "Key-SKU断货统计",
    title: "Key-SKU断货统计-1",
    headers: ["周次", "货号个数", "断货个数", "断货率", "备注说明"],
    weekField: "周次",
    numericFields: ["货号个数", "断货个数", "断货率"],
    rows: [
      { "周次": "07-23", "货号个数": 2264, "断货个数": 48, "断货率": 0.0212, "备注说明": "" },
      { "周次": "07-30", "货号个数": 2248, "断货个数": 69, "断货率": 0.0307, "备注说明": "阈值调整" },
      { "周次": "08-07", "货号个数": 2248, "断货个数": 46, "断货率": 0.0204626334519573, "备注说明": null }
    ]
  },
  {
    id: "bargain-stats",
    sheet: "议价数据统计",
    title: "议价数据统计-1",
    headers: ["周次", "单笔议价金额（元）", "备货议价金额（元）", "议价合计（元）"],
    weekField: "周次",
    numericFields: ["单笔议价金额（元）", "备货议价金额（元）", "议价合计（元）"],
    rows: [
      { "周次": "07-23", "单笔议价金额（元）": 4971.97, "备货议价金额（元）": 56502.88, "议价合计（元）": 61474.85 },
      { "周次": "07-30", "单笔议价金额（元）": 5210.54, "备货议价金额（元）": 25998.03, "议价合计（元）": 31208.57 },
      { "周次": "08-07", "单笔议价金额（元）": 17901.513, "备货议价金额（元）": 5144.33, "议价合计（元）": 23045.843 }
    ]
  },
  {
    id: "warehouse-output",
    sheet: "仓储单位产出",
    title: "仓储单位产出-1",
    headers: ["月份", "主仓面积", "廊坊仓面积", "成都仓面积", "广州仓面积", "主仓每平米销售额", "廊坊每平米销售额", "成都每平米销售额", "广州仓每平米销售额", "主仓每平米毛利", "廊坊每平米毛利", "成都每平米毛利", "广州仓每平米毛利", "主仓租金", "廊坊仓租金", "成都仓租金", "广州仓租金"],
    weekField: "月份",
    numericFields: ["主仓面积", "廊坊仓面积", "成都仓面积", "广州仓面积", "主仓每平米销售额", "廊坊每平米销售额", "成都每平米销售额", "广州仓每平米销售额", "主仓每平米毛利", "廊坊每平米毛利", "成都每平米毛利", "广州仓每平米毛利", "主仓租金", "廊坊仓租金", "成都仓租金", "广州仓租金"],
    rows: [
      { "月份": "2026-06", "主仓面积": 8446, "廊坊仓面积": 1301, "成都仓面积": 1040, "广州仓面积": 1202, "主仓每平米销售额": 1814.43594600995, "廊坊每平米销售额": 1091.67947732513, "成都每平米销售额": 558.466346153846, "广州仓每平米销售额": 519.371048252912, "主仓每平米毛利": 543.79694529955, "廊坊每平米毛利": 289.948501152959, "成都每平米毛利": 197.801923076923, "广州仓每平米毛利": 184.280366056572, "主仓租金": 225323, "廊坊仓租金": 33193, "成都仓租金": 24736, "广州仓租金": 22844 },
      { "月份": "2026-07", "主仓面积": 8446, "廊坊仓面积": 1100, "成都仓面积": 1040, "广州仓面积": 1202, "主仓每平米销售额": 2164.31079801089, "廊坊每平米销售额": 2311.32181818182, "成都每平米销售额": 699.485576923077, "广州仓每平米销售额": 794.636439267887, "主仓每平米毛利": 622.540255742363, "廊坊每平米毛利": 571.837272727273, "成都每平米毛利": 241.398076923077, "广州仓每平米毛利": 228.964226289517, "主仓租金": 225323, "廊坊仓租金": 33193, "成都仓租金": 24736, "广州仓租金": 22844 }
    ]
  },
  {
    id: "efficiency",
    sheet: "效率提升",
    title: "效率提升-1",
    headers: ["周次", "模块", "本周进展"],
    weekField: "周次",
    numericFields: [],
    rows: [
      { "周次": "07-23", "模块": "AI", "本周进展": "知识库：产品资料投喂；发票命名自动化" },
      { "周次": "07-30", "模块": "流程优化", "本周进展": "年度协议管理打通；标签自动流转模块沟通" },
      { "周次": "08-06", "模块": "AI", "本周进展": "知识库:\n①文件 URL 规范化,修复中文文件名和空格导致链接打不开的问题口径。\n②知识库龙虾待IT开通商品详情页端口后再开放给销售。\n报价:\n①耗材搜索排序 V1回归中：框架集成完成，12 个基础自动化测试通过，进入真实业务 Case 回归. \n②AI任务分流内测试运行：子代理分流可执行，日报改为人工触发验证口径\n端口进展\n自动关联：采购单中心、销售成品采购任务两个8月底完成\n9月开展：自动生产转换——联动库存+占用+阈值+大小包装：\n时间待定：自动收票——待收票（成品）" }
    ]
  },
  {
    id: "inventory-turnover",
    sheet: "库存周转分析",
    title: "库存周转分析-1",
    headers: ["仓库/品类", "6月库存金额", "6月年营业成本", "6月周转次数", "6月平均周转天数", "7月库存金额", "7月年营业成本", "7月周转次数", "7月平均周转天数", "库存金额环比变化", "周转天数变化(天）"],
    weekField: null,
    numericFields: ["6月库存金额", "6月年营业成本", "6月周转次数", "6月平均周转天数", "7月库存金额", "7月年营业成本", "7月周转次数", "7月平均周转天数", "库存金额环比变化", "周转天数变化(天）"],
    rows: [
      { "仓库/品类": "主仓有阈值的明星产品", "6月库存金额": 12960913.76, "6月年营业成本": 59780928.78, "6月周转次数": 4.61, "6月平均周转天数": 79.1, "7月库存金额": 14129745.66, "7月年营业成本": 57251892.44, "7月周转次数": 4.05, "7月平均周转天数": 90.1, "库存金额环比变化": 0.09, "周转天数变化(天）": 11 },
      { "仓库/品类": "主仓所有商品", "6月库存金额": 32900585.42, "6月年营业成本": 134744245.49, "6月周转次数": 4.1, "6月平均周转天数": 89.1, "7月库存金额": 34198877.49, "7月年营业成本": 131835427.48, "7月周转次数": 3.85, "7月平均周转天数": 94.7, "库存金额环比变化": 0.039, "周转天数变化(天）": 5.6 },
      { "仓库/品类": "成都所有商品", "6月库存金额": 1979750.45, "6月年营业成本": 14321547.68, "6月周转次数": 7.23, "6月平均周转天数": 50.5, "7月库存金额": 1851979.97, "7月年营业成本": 13356804.36, "7月周转次数": 7.21, "7月平均周转天数": 50.6, "库存金额环比变化": -0.065, "周转天数变化(天）": 0.1 },
      { "仓库/品类": "廊坊所有商品", "6月库存金额": 3506939.44, "6月年营业成本": 31251759.68, "6月周转次数": 8.91, "6月平均周转天数": 41, "7月库存金额": 3502534.97, "7月年营业成本": 32057459.56, "7月周转次数": 9.15, "7月平均周转天数": 39.9, "库存金额环比变化": -0.001, "周转天数变化(天）": -1.1 },
      { "仓库/品类": "广州所有商品", "6月库存金额": 1367089.88, "6月年营业成本": 14437535.33, "6月周转次数": 10.56, "6月平均周转天数": 34.6, "7月库存金额": 1345513.08, "7月年营业成本": 12785473.64, "7月周转次数": 9.5, "7月平均周转天数": 38.4, "库存金额环比变化": -0.016, "周转天数变化(天）": 3.8 },
      { "仓库/品类": "宜昌所有商品", "6月库存金额": 11818000, "6月年营业成本": 0, "6月周转次数": 0, "6月平均周转天数": 0, "7月库存金额": 11747000, "7月年营业成本": 0, "7月周转次数": 0, "7月平均周转天数": 0, "库存金额环比变化": 0, "周转天数变化(天）": 0 },
      { "仓库/品类": "全部在售库库存合计", "6月库存金额": 51572365.19, "6月年营业成本": 194755088.18, "6月周转次数": 3.77634586784016, "6月平均周转天数": 98.6, "7月库存金额": 52645905.51, "7月年营业成本": 190035165.04, "7月周转次数": 3.60968556242048, "7月平均周转天数": 103, "库存金额环比变化": 0.019, "周转天数变化(天）": 4.4 }
    ]
  },
  {
    id: "summary",
    sheet: "总结",
    title: "总结-1",
    headers: ["周次", "本周小结"],
    weekField: "周次",
    numericFields: [],
    rows: [
      { "周次": "08-07", "本周小结": "KEY-SKY断货率进一步降低到2%，新下单金额上涨81%，7月廊坊仓面积从1300平降低到1100平" }
    ]
  }
];

const state = { workbook: demoWorkbook };
const fileInput = document.getElementById("fileInput");
const demoButton = document.getElementById("demoButton");
const screenshotButton = document.getElementById("screenshotButton");
const statusText = document.getElementById("statusText");
const overviewSummary = document.getElementById("overviewSummary");
const overviewGrid = document.getElementById("overviewGrid");
const dashboardContent = document.getElementById("dashboardContent");

function isEmpty(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function formatValue(value, field) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") {
    const lower = String(field || "").toLowerCase();
    if (lower.includes("比例") || lower.includes("率") || lower.includes("环比")) return `${(value * 100).toFixed(1)}%`;
    if (Math.abs(value) >= 1000000) return value.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
    if (Number.isInteger(value)) return value.toLocaleString("zh-CN");
    return value.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
  }
  return String(value);
}

function formatChange(value, field) {
  if (value === null || value === undefined || Number.isNaN(value)) return { text: "无环比", cls: "flat" };
  const lower = String(field || "").toLowerCase();
  const icon = value > 0 ? "↑" : "↓";
  if (lower.includes("比例") || lower.includes("率") || lower.includes("环比")) {
    return { text: `${icon} ${Math.abs(value * 100).toFixed(1)}%`, cls: value > 0 ? "up" : "down" };
  }
  return { text: `${icon} ${Math.abs(value).toLocaleString("zh-CN", { maximumFractionDigits: 2 })}`, cls: value > 0 ? "up" : "down" };
}

function looksLikeWeek(text) {
  const sample = String(text || "").trim().toLowerCase();
  return sample.includes("周") ||
    sample.includes("week") ||
    sample.includes("月份") ||
    sample === "月" ||
    /^\d{1,2}[./-]\d{1,2}$/.test(sample) ||
    /^\d{4}-w\d{1,2}$/.test(sample) ||
    /^\d{4}-\d{1,2}$/.test(sample);
}

function tryNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  const text = String(value).trim();
  if (!text || text === "-" || text === "--") return null;
  const percent = text.endsWith("%");
  const cleaned = text.replace(/[,\s万元天个元pp]/g, "");
  const num = Number(cleaned);
  if (!Number.isFinite(num)) return null;
  return percent ? num / 100 : num;
}

function normalizeFieldName(value) {
  return String(value || "")
    .replace(/[（()）\s_-]/g, "")
    .trim();
}

function findFieldByName(fields, target) {
  const normalizedTarget = normalizeFieldName(target);
  return fields.find((field) => normalizeFieldName(field) === normalizedTarget) || null;
}

function findFieldByKeywords(fields, keywords) {
  return fields.find((field) => {
    const normalized = normalizeFieldName(field);
    return keywords.every((keyword) => normalized.includes(normalizeFieldName(keyword)));
  }) || null;
}

function findFieldByAliases(fields, aliases) {
  const list = Array.isArray(aliases) ? aliases : [aliases];
  for (const alias of list) {
    const field = findFieldByName(fields, alias);
    if (field) return field;
  }
  return null;
}

function isOrderTable(table) {
  return table.sheet === "订单数据" || table.title === "订单数据" || table.title.startsWith("订单数据-");
}

function buildSeries(table, field) {
  const values = table.rows.map((row) => {
    const value = Number(row[field]);
    return Number.isFinite(value) ? value : null;
  });
  if (!values.some((value) => value !== null)) return null;
  return { name: field, values };
}

function findExistingField(table, targets) {
  for (const target of targets) {
    const byName = findFieldByName(table.numericFields, target);
    if (byName) return byName;
  }
  return null;
}

function buildConfiguredChart(table, title, fieldTargets, options = {}) {
  const fields = fieldTargets
    .map((target) => findExistingField(table, [target]))
    .filter(Boolean);
  if (!fields.length) return "";
  const seriesList = fields
    .map((field) => buildSeries(table, field))
    .filter(Boolean);
  if (!seriesList.length) return "";
  const labels = table.rows.map((row) => row[table.weekField]);
  if (seriesList.length === 1) {
    return buildLineChart(labels, seriesList[0].values, title || seriesList[0].name, options);
  }
  return buildMultiLineChart(labels, seriesList, title, options);
}

function splitBlocks(rows) {
  const rawBlocks = [];
  let current = [];
  rows.forEach((row) => {
    const hasValue = row.some((cell) => !isEmpty(cell));
    if (hasValue) {
      current.push(row);
    } else if (current.length) {
      rawBlocks.push(current);
      current = [];
    }
  });
  if (current.length) rawBlocks.push(current);

  const blocks = [];
  rawBlocks.forEach((block) => {
    const firstNonEmptyCount = block[0].filter((cell) => !isEmpty(cell)).length;
    const isSingleTitleRow = block.length === 1 && firstNonEmptyCount <= 1;
    if (isSingleTitleRow && blocks.length < rawBlocks.length) {
      blocks.push(block);
      return;
    }
    const previous = blocks[blocks.length - 1];
    if (previous && previous.length === 1 && previous[0].filter((cell) => !isEmpty(cell)).length <= 1) {
      blocks[blocks.length - 1] = [...previous, ...block];
    } else {
      blocks.push(block);
    }
  });
  return blocks;
}

function buildTableFromBlock(sheetName, title, block, blockIndex) {
  const nonEmptyRows = block.filter((row) => row.some((cell) => !isEmpty(cell)));
  if (nonEmptyRows.length < 2) return null;
  let headerRowIndex = 0;
  let tableTitle = title || `${sheetName}-${blockIndex + 1}`;
  const firstRowFilled = nonEmptyRows[0].filter((cell) => !isEmpty(cell)).length;
  if (firstRowFilled <= 1 && nonEmptyRows.length >= 2) {
    const firstValue = nonEmptyRows[0].find((cell) => !isEmpty(cell));
    if (firstValue) tableTitle = String(firstValue).trim();
    headerRowIndex = 1;
  }

  const headers = nonEmptyRows[headerRowIndex].map((value, index) => {
    const text = value === undefined || value === null ? "" : String(value).trim();
    return text || `字段${index + 1}`;
  });

  const rows = nonEmptyRows.slice(headerRowIndex + 1).map((rawRow) => {
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = index < rawRow.length && !isEmpty(rawRow[index]) ? rawRow[index] : null;
    });
    return rowObj;
  }).filter((row) => Object.values(row).some((value) => !isEmpty(value)));

  if (!rows.length) return null;

  const numericFields = [];
  headers.forEach((header) => {
    const converted = rows.map((row) => tryNumber(row[header]));
    const numericCount = converted.filter((value) => value !== null).length;
    if (numericCount >= Math.max(1, Math.floor(rows.length / 2))) {
      numericFields.push(header);
      rows.forEach((row, index) => {
        if (converted[index] !== null) row[header] = converted[index];
      });
    }
  });

  let weekField = headers.find((header) => looksLikeWeek(header)) || null;
  if (!weekField && headers.length) {
    const firstHeader = headers[0];
    const samples = rows.slice(0, 5).map((row) => row[firstHeader]);
    if (samples.some((sample) => looksLikeWeek(sample))) weekField = firstHeader;
  }

  return {
    id: `${sheetName}-${blockIndex + 1}`.replace(/\s+/g, "-").toLowerCase(),
    sheet: sheetName,
    title: tableTitle,
    headers,
    rows,
    numericFields,
    weekField
  };
}

function parseWorkbook(file, data) {
  const workbook = XLSX.read(data, { type: "array" });
  const tables = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    const blocks = splitBlocks(rows);
    blocks.forEach((block, blockIndex) => {
      const table = buildTableFromBlock(sheetName, `${sheetName}-${blockIndex + 1}`, block, blockIndex);
      if (table) tables.push(table);
    });
  });
  return {
    workbookName: file.name,
    tables
  };
}

function shouldShowChart(table) {
  return !NON_CHART_SHEETS.has(table.sheet) && !!table.weekField && table.numericFields.length > 0;
}

function shouldShowKpi(table) {
  return !NON_KPI_SHEETS.has(table.sheet);
}

function computeWeeklyMetrics(table) {
  if (!table || !table.weekField || table.rows.length < 1) return [];
  const latest = table.rows[table.rows.length - 1];
  const previous = table.rows.length > 1 ? table.rows[table.rows.length - 2] : null;
  return table.numericFields.slice(0, 6).map((field) => {
    const latestValue = Number(latest[field]);
    if (!Number.isFinite(latestValue)) return null;
    const previousValue = previous ? Number(previous[field]) : null;
    let delta = null;
    if (Number.isFinite(previousValue)) {
      if (String(field).includes("比例") || String(field).includes("率")) {
        delta = latestValue - previousValue;
      } else if (previousValue !== 0) {
        delta = (latestValue - previousValue) / previousValue;
      }
    }
    return { field, latestValue, delta, week: latest[table.weekField] };
  }).filter(Boolean);
}

function computeMetricsForFields(table, fields) {
  if (!table || !table.weekField || table.rows.length < 1) return [];
  const latest = table.rows[table.rows.length - 1];
  const previous = table.rows.length > 1 ? table.rows[table.rows.length - 2] : null;
  return fields.map((field) => {
    if (!field) return null;
    const latestValue = Number(latest[field]);
    if (!Number.isFinite(latestValue)) return null;
    const previousValue = previous ? Number(previous[field]) : null;
    let delta = null;
    if (Number.isFinite(previousValue)) {
      if (String(field).includes("比例") || String(field).includes("率")) {
        delta = latestValue - previousValue;
      } else if (previousValue !== 0) {
        delta = (latestValue - previousValue) / previousValue;
      }
    }
    return { field, latestValue, delta, week: latest[table.weekField] };
  }).filter(Boolean);
}

function buildOverviewMetrics(workbook) {
  return OVERVIEW_GROUPS.map((group) => {
    const table = workbook.tables.find((item) => item.sheet === group.sheet || item.title === group.sheet);
    if (!table || !table.weekField || !table.rows.length) return null;
    const latest = table.rows[table.rows.length - 1];
    const previous = table.rows.length > 1 ? table.rows[table.rows.length - 2] : null;
    const metrics = group.fields.map((fieldName) => {
      const field = findFieldByAliases(table.numericFields, fieldName);
      if (!field) return null;
      const latestValue = Number(latest[field]);
      if (!Number.isFinite(latestValue)) return null;
      const previousValue = previous ? Number(previous[field]) : null;
      let delta = null;
      if (Number.isFinite(previousValue)) {
        if (String(field).includes("比例") || String(field).includes("率")) {
          delta = latestValue - previousValue;
        } else if (previousValue !== 0) {
          delta = (latestValue - previousValue) / previousValue;
        }
      }
      return { field, latestValue, delta };
    }).filter(Boolean);
    if (!metrics.length) return null;
    return {
      title: group.title,
      week: latest[table.weekField],
      metrics
    };
  }).filter(Boolean);
}

function extractOverviewSummary(workbook) {
  const summaryTable = workbook.tables.find((item) => item.sheet === "总结" || item.title === "总结" || item.title === "经营总览");
  if (!summaryTable) return "";
  const summaryField =
    findFieldByAliases(summaryTable.headers || [], ["本周小结", "总结", "描述", "描述性文字"]) ||
    findFieldByAliases(Object.keys(summaryTable.rows[0] || {}), ["本周小结", "总结", "描述", "描述性文字"]);
  if (summaryField) {
    const latestRow = summaryTable.rows.filter((row) => !isEmpty(row[summaryField])).slice(-1)[0];
    if (latestRow) return String(latestRow[summaryField]).trim();
  }

  const lines = summaryTable.rows
    .map((row) => Object.values(row).filter((value) => !isEmpty(value)).map((value) => String(value).trim()))
    .filter((values) => values.length === 1)
    .map((values) => values[0])
    .filter((text) => text !== "经营总览" && text !== "总结");

  return lines.slice(-1)[0] || "";
}

function buildLineChart(labels, values, title, options = {}) {
  if (!labels.length || !values.length) {
    return `<div class="chart-card"><h3 class="chart-title">${title}</h3><div class="empty-state">暂无可绘制数据</div></div>`;
  }
  const width = options.width || 420;
  const height = options.height || 220;
  const left = 42;
  const right = 16;
  const top = 18;
  const bottom = 36;
  const innerW = width - left - right;
  const innerH = height - top - bottom;
  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = max - min || 1;
  const points = values.map((value, index) => {
    const x = left + (innerW * index) / Math.max(values.length - 1, 1);
    const y = top + innerH - ((value - min) / range) * innerH;
    return { x, y, label: labels[index] };
  });
  const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
  const yTicks = [0, 0.5, 1].map((ratio) => {
    const y = top + innerH - ratio * innerH;
    const tickValue = min + ratio * range;
    return { y, text: tickValue.toLocaleString("zh-CN", { maximumFractionDigits: 2 }) };
  });
  return `
    <div class="chart-card${options.cardClass ? ` ${options.cardClass}` : ""}">
      <h3 class="chart-title">${title}</h3>
      <svg class="svg-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="${title}">
        ${yTicks.map((tick) => `<line class="grid" x1="${left}" y1="${tick.y}" x2="${width - right}" y2="${tick.y}"></line><text x="4" y="${tick.y + 4}">${tick.text}</text>`).join("")}
        <line class="axis" x1="${left}" y1="${top}" x2="${left}" y2="${top + innerH}"></line>
        <line class="axis" x1="${left}" y1="${top + innerH}" x2="${width - right}" y2="${top + innerH}"></line>
        <path class="line" d="${path}"></path>
        ${points.map((point) => `<circle class="dot" cx="${point.x}" cy="${point.y}" r="4"></circle>`).join("")}
        ${points.map((point) => `<text x="${point.x - 10}" y="${height - 12}">${point.label}</text>`).join("")}
      </svg>
      <div class="chart-meta">共 ${labels.length} 个周节点</div>
    </div>
  `;
}

function buildMultiLineChart(labels, seriesList, title, options = {}) {
  if (!labels.length || !seriesList.length) {
    return `<div class="chart-card"><h3 class="chart-title">${title}</h3><div class="empty-state">暂无可绘制数据</div></div>`;
  }
  const width = options.width || 420;
  const height = options.height || 220;
  const left = 42;
  const right = 16;
  const top = 18;
  const bottom = 36;
  const innerW = width - left - right;
  const innerH = height - top - bottom;
  const palette = ["#2d6fb8", "#d97706", "#0b9a53", "#8b5cf6"];

  const allValues = seriesList.flatMap((series) => series.values).filter((value) => Number.isFinite(value));
  const min = Math.min(...allValues);
  const max = Math.max(...allValues);
  const range = max - min || 1;

  const lines = seriesList.map((series, seriesIndex) => {
    const points = series.values.map((value, index) => {
      const x = left + (innerW * index) / Math.max(series.values.length - 1, 1);
      const y = top + innerH - ((value - min) / range) * innerH;
      return { x, y, label: labels[index] };
    });
    const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
    return {
      color: palette[seriesIndex % palette.length],
      name: series.name,
      points,
      path
    };
  });

  const yTicks = [0, 0.5, 1].map((ratio) => {
    const y = top + innerH - ratio * innerH;
    const tickValue = min + ratio * range;
    return { y, text: tickValue.toLocaleString("zh-CN", { maximumFractionDigits: 2 }) };
  });

  return `
    <div class="chart-card${options.cardClass ? ` ${options.cardClass}` : ""}">
      <h3 class="chart-title">${title}</h3>
      <svg class="svg-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="${title}">
        ${yTicks.map((tick) => `<line class="grid" x1="${left}" y1="${tick.y}" x2="${width - right}" y2="${tick.y}"></line><text x="4" y="${tick.y + 4}">${tick.text}</text>`).join("")}
        <line class="axis" x1="${left}" y1="${top}" x2="${left}" y2="${top + innerH}"></line>
        <line class="axis" x1="${left}" y1="${top + innerH}" x2="${width - right}" y2="${top + innerH}"></line>
        ${lines.map((line) => `<path d="${line.path}" fill="none" stroke="${line.color}" stroke-width="3"></path>`).join("")}
        ${lines.map((line) => line.points.map((point) => `<circle cx="${point.x}" cy="${point.y}" r="4" fill="#fff" stroke="${line.color}" stroke-width="2"></circle>`).join("")).join("")}
        ${labels.map((label, index) => {
          const x = left + (innerW * index) / Math.max(labels.length - 1, 1);
          return `<text x="${x - 14}" y="${height - 12}">${label}</text>`;
        }).join("")}
      </svg>
      <div class="chart-legend">
        ${lines.map((line) => `
          <span class="chart-legend-item">
            <span class="chart-legend-swatch" style="background:${line.color}"></span>
            ${line.name}
          </span>
        `).join("")}
      </div>
      <div class="chart-meta">共 ${labels.length} 个周节点</div>
    </div>
  `;
}

function buildTableHtml(table) {
  const wrapperClassNames = [];
  const tableClassNames = [];
  if (table.sheet === "仓储单位产出" || table.title === "仓储单位产出") {
    tableClassNames.push("warehouse-table");
    wrapperClassNames.push("warehouse-table-wrapper");
  }
  if (table.sheet === "库存周转分析" || table.title === "库存周转分析") tableClassNames.push("inventory-table");
  const headerHtml = table.headers.map((header) => {
    const wrapHeaderSheets = ["仓储单位产出", "库存周转分析", "采购数据统计"];
    const wrapHeader = wrapHeaderSheets.includes(table.sheet) || table.title.startsWith("订单数据-");
    const className = wrapHeader ? "wrap-header" : "";
    return `<th class="${className}">${header}</th>`;
  }).join("");
  const rowsHtml = table.rows.map((row) => `
    <tr>${table.headers.map((header) => {
      const className = WRAP_COLUMNS.has(header) ? "wrap-text" : "";
      return `<td class="${className}">${formatValue(row[header], header)}</td>`;
    }).join("")}</tr>
  `).join("");

  return `
    <div class="table-wrapper ${wrapperClassNames.join(" ")}">
      <table class="${tableClassNames.join(" ")}">
        <thead><tr>${headerHtml}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;
}

function buildChartsHtml(table) {
  if (!shouldShowChart(table)) {
    return `<div class="empty-state">该模块按新模板不展示图表。</div>`;
  }
  const labels = table.rows.map((row) => row[table.weekField]);
  const isPurchaseSheet = table.sheet === "采购数据统计" || table.title === "采购数据统计" || table.title === "采购数据统计-1";
  const isOrderSheet = isOrderTable(table);
  const isWarehouseSheet = table.sheet === "仓储单位产出" || table.title === "仓储单位产出";

  if (isOrderSheet) {
    return [
      buildConfiguredChart(table, "新下单金额与待审核订单金额趋势", [
        findFieldByAliases(table.numericFields, ["新下单订单金额", "新下单金额"]),
        findFieldByAliases(table.numericFields, ["待审核总订单金额", "待审核订单金额", "待审核总订单金额"])
      ].filter(Boolean), { cardClass: "chart-card-featured chart-card-order", width: 760, height: 280 })
    ].filter(Boolean).join("");
  }

  if (isPurchaseSheet) {
    return [
      buildConfiguredChart(table, "采购合同个数与采购单个数趋势", ["采购合同个数", "采购单个数"]),
      buildConfiguredChart(table, "销售采购单、直发采购单与直发采购单比例趋势", ["销售采购单个数", "直发采购单个数", "直发采购单比例"])
    ].filter(Boolean).join("");
  }

  if (isWarehouseSheet) {
    return [
      buildConfiguredChart(table, "各仓每平米销售额趋势", ["主仓每平米销售额", "廊坊每平米销售额", "成都每平米销售额", "广州仓每平米销售额"]),
      buildConfiguredChart(table, "各仓每平米毛利趋势", ["主仓每平米毛利", "廊坊每平米毛利", "成都每平米毛利", "广州仓每平米毛利"])
    ].filter(Boolean).join("");
  }
  const chartFields = table.numericFields.slice(0, 4);
  if (!chartFields.length) {
    return `<div class="empty-state">当前表没有适合绘图的数值字段。</div>`;
  }
  return chartFields.map((field) => {
    const values = table.rows.map((row) => {
      const value = Number(row[field]);
      return Number.isFinite(value) ? value : null;
    });
    return buildLineChart(labels, values, field);
  }).join("");
}

function buildInstructionHtml(table) {
  if (table.sheet !== "填写说明") return "";
  return `
    <div class="note-grid">
      ${table.rows.map((row) => `
        <div class="note-card">
          <div class="note-title">${row["说明项"] || "说明"}</div>
          <div class="note-value">${row["填写要求"] || "-"}</div>
        </div>
      `).join("")}
    </div>
  `;
}

function buildMetricsSection(table) {
  if (!shouldShowKpi(table)) return "";
  const weeklyMetrics = computeWeeklyMetrics(table);
  if (!weeklyMetrics.length) return "";
  if (table.sheet === "仓储单位产出" || table.title === "仓储单位产出-1" || table.title === "仓储单位产出") {
    const grossProfitFields = [
      findFieldByAliases(table.numericFields, ["主仓每平米毛利"]),
      findFieldByAliases(table.numericFields, ["廊坊每平米毛利"]),
      findFieldByAliases(table.numericFields, ["成都每平米毛利"]),
      findFieldByAliases(table.numericFields, ["广州仓每平米毛利"])
    ].filter(Boolean);
    const metrics = computeMetricsForFields(table, grossProfitFields);
    return `
      <div class="dashboard-block">
        <h4 class="dashboard-block-title">指标卡片</h4>
        <div class="warehouse-kpi-layout">
          ${metrics.map((metric) => {
            const change = formatChange(metric.delta, metric.field);
            return `
              <div class="kpi-card warehouse-kpi-card">
                <div class="kpi-title">${metric.field}</div>
                <div class="kpi-value">${formatValue(metric.latestValue, metric.field)}</div>
                <div class="kpi-change ${change.cls}">${metric.week} ${change.text}</div>
              </div>
            `;
          }).join("")}
        </div>
      </div>
    `;
  }
  if (isOrderTable(table)) {
    const amountFields = [
      findFieldByAliases(table.numericFields, ["新下单订单金额", "新下单金额"]),
      findFieldByAliases(table.numericFields, ["待审核总订单金额", "待审核订单金额"])
    ].filter(Boolean);
    const countFields = [
      findFieldByAliases(table.numericFields, ["新下单订单条数", "新下单条数"]),
      findFieldByAliases(table.numericFields, ["待审核总订单条数", "待审核订单条数"])
    ].filter(Boolean);
    const selectedFields = [...amountFields, ...countFields];
    const metrics = selectedFields.map((field) => weeklyMetrics.find((metric) => metric.field === field)).filter(Boolean);
    return `
      <div class="dashboard-block">
        <h4 class="dashboard-block-title">指标卡片</h4>
        <div class="order-kpi-layout">
          ${metrics.map((metric, index) => {
            const change = formatChange(metric.delta, metric.field);
            const cardClass = index < amountFields.length ? "order-kpi-primary" : "order-kpi-secondary";
            return `
              <div class="kpi-card ${cardClass}">
                <div class="kpi-title">${metric.field}</div>
                <div class="kpi-value">${formatValue(metric.latestValue, metric.field)}</div>
                <div class="kpi-change ${change.cls}">${metric.week} ${change.text}</div>
              </div>
            `;
          }).join("")}
        </div>
      </div>
    `;
  }
  return `
    <div class="dashboard-block">
      <h4 class="dashboard-block-title">指标卡片</h4>
      <div class="kpi-grid">${weeklyMetrics.slice(0, 4).map((metric) => {
        const change = formatChange(metric.delta, metric.field);
        return `
          <div class="kpi-card">
            <div class="kpi-title">${metric.field}</div>
            <div class="kpi-value">${formatValue(metric.latestValue, metric.field)}</div>
            <div class="kpi-change ${change.cls}">${metric.week} ${change.text}</div>
          </div>
        `;
      }).join("")}</div>
    </div>
  `;
}

function buildChartsSection(table) {
  if (!shouldShowChart(table)) return "";
  const chartsHtml = buildChartsHtml(table);
  if (!chartsHtml || chartsHtml.includes("empty-state")) return "";
  const chartGridClass = isOrderTable(table) ? "chart-grid chart-grid-featured" : "chart-grid";
  return `
    <div class="dashboard-block">
      <h4 class="dashboard-block-title">趋势图</h4>
      <div class="${chartGridClass}">${chartsHtml}</div>
    </div>
  `;
}

function renderOverview() {
  const summaryText = extractOverviewSummary(state.workbook);
  if (summaryText) {
    overviewSummary.hidden = false;
    overviewSummary.innerHTML = summaryText.split("\n").map((line) => `<p class="overview-summary-line">${line}</p>`).join("");
  } else {
    overviewSummary.hidden = true;
    overviewSummary.innerHTML = "";
  }

  const groups = buildOverviewMetrics(state.workbook);
  if (!groups.length) {
    overviewGrid.innerHTML = `<div class="empty-state">当前没有识别到可汇总的周维度 KPI。</div>`;
    return;
  }
  overviewGrid.className = "kpi-grid overview-grid overview-group-grid";
  overviewGrid.innerHTML = groups.map((group) => {
    return `
      <div class="kpi-card overview-group-card">
        <div class="kpi-title">${group.title}</div>
        <div class="overview-metrics">
          ${group.metrics.map((metric) => {
            const change = formatChange(metric.delta, metric.field);
            return `
              <div class="overview-metric-row">
                <div class="overview-metric-name">${metric.field}</div>
                <div class="overview-metric-value">${formatValue(metric.latestValue, metric.field)}</div>
                <div class="kpi-change ${change.cls}">${group.week} ${change.text}</div>
              </div>
            `;
          }).join("")}
        </div>
      </div>
    `;
  }).join("");
}

function renderDashboardContent() {
  const tables = state.workbook.tables.filter((table) => table.sheet !== "总结" && table.title !== "总结-1" && table.title !== "总结");
  if (!tables.length) {
    dashboardContent.innerHTML = `<div class="empty-state">没有识别到任何数据表。</div>`;
    return;
  }

  dashboardContent.innerHTML = tables.map((table) => {
    const meta = table.weekField
      ? `已识别周期字段：${table.weekField} · 共 ${table.rows.length} 行`
      : `未识别周期字段 · 共 ${table.rows.length} 行`;

    const metricsSection = buildMetricsSection(table);
    const chartsSection = buildChartsSection(table);
    const instructionsHtml = table.sheet === "填写说明"
      ? `
        <div class="dashboard-block">
          <h4 class="dashboard-block-title">填写提示</h4>
          ${buildInstructionHtml(table)}
        </div>
      `
      : "";

    return `
      <article class="dashboard-section">
        <div class="dashboard-section-header">
          <h3 class="dashboard-section-title">${table.title}</h3>
          <p class="dashboard-section-meta">${meta}</p>
        </div>
        ${metricsSection}
        ${chartsSection}
        ${instructionsHtml}
        <div class="dashboard-block">
          <h4 class="dashboard-block-title">数据明细</h4>
          ${buildTableHtml(table)}
        </div>
      </article>
    `;
  }).join("");
}

function renderAll() {
  renderOverview();
  renderDashboardContent();
}

async function uploadWorkbook(file) {
  statusText.textContent = `正在解析：${file.name}`;
  const buffer = await file.arrayBuffer();
  const workbook = parseWorkbook(file, buffer);
  if (!workbook.tables.length) throw new Error("没有识别到可用数据表");
  state.workbook = workbook;
  statusText.textContent = `已加载：${workbook.workbookName}`;
  renderAll();
}

async function downloadScreenshot() {
  screenshotButton.disabled = true;
  statusText.textContent = "正在生成页面截图...";
  try {
    const canvas = await html2canvas(document.getElementById("captureTarget"), {
      backgroundColor: "#f4f8fc",
      scale: 2,
      useCORS: true,
      ignoreElements: (element) => element.classList?.contains("no-capture")
    });
    const link = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    link.download = `周维度经营看板截图-${stamp}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
    statusText.textContent = "页面截图已生成并开始下载";
  } catch (error) {
    statusText.textContent = `截图失败：${error.message}`;
  } finally {
    screenshotButton.disabled = false;
  }
}

fileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    await uploadWorkbook(file);
  } catch (error) {
    statusText.textContent = `上传失败：${error.message}`;
  } finally {
    fileInput.value = "";
  }
});

demoButton.addEventListener("click", () => {
  state.workbook = demoWorkbook;
  statusText.textContent = "已恢复示例数据";
  renderAll();
});

screenshotButton.addEventListener("click", downloadScreenshot);

renderAll();
