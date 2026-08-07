const NON_CHART_SHEETS = new Set(["填写说明", "效率提升", "仓储单位产出", "库存周转分析"]);
const NON_KPI_SHEETS = new Set(["填写说明", "库存周转分析"]);
const OVERVIEW_PRIORITY = ["采购数据统计", "Key-SKU断货统计", "议价数据统计"];
const WRAP_COLUMNS = new Set(["本周进展", "填写要求", "备注说明"]);

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

const state = { workbook: demoWorkbook };
const fileInput = document.getElementById("fileInput");
const demoButton = document.getElementById("demoButton");
const screenshotButton = document.getElementById("screenshotButton");
const statusText = document.getElementById("statusText");
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
  return sample.includes("周") || sample.includes("week") || /^\d{1,2}[./-]\d{1,2}$/.test(sample) || /^\d{4}-w\d{1,2}$/.test(sample);
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

function buildOverviewMetrics(workbook) {
  return OVERVIEW_PRIORITY.flatMap((sheetName) => {
    const table = workbook.tables.find((item) => item.sheet === sheetName);
    if (!table) return [];
    return computeWeeklyMetrics(table).slice(0, 2).map((metric) => ({
      ...metric,
      title: `${table.title} · ${metric.field}`
    }));
  });
}

function buildLineChart(labels, values, title) {
  if (!labels.length || !values.length) {
    return `<div class="chart-card"><h3 class="chart-title">${title}</h3><div class="empty-state">暂无可绘制数据</div></div>`;
  }
  const width = 420;
  const height = 220;
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
    <div class="chart-card">
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

function buildMultiLineChart(labels, seriesList, title) {
  if (!labels.length || !seriesList.length) {
    return `<div class="chart-card"><h3 class="chart-title">${title}</h3><div class="empty-state">暂无可绘制数据</div></div>`;
  }
  const width = 420;
  const height = 220;
  const left = 42;
  const right = 16;
  const top = 18;
  const bottom = 36;
  const innerW = width - left - right;
  const innerH = height - top - bottom;
  const palette = ["#2d6fb8", "#0b9a53", "#d63b3b", "#8b5cf6"];

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
    <div class="chart-card">
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
      <div class="chart-meta">${lines.map((line) => `<span style="margin-right:12px;color:${line.color};font-weight:700">${line.name}</span>`).join("")}</div>
    </div>
  `;
}

function buildTableHtml(table) {
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
    <div class="table-wrapper">
      <table>
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
  if (table.sheet.includes("采购数据统计") || table.title.includes("采购数据统计")) {
    const combinedFields = ["采购合同个数", "采购单个数"];
    const combinedAvailable = combinedFields.every((field) => table.numericFields.includes(field));
    const remainingFields = table.numericFields.filter((field) => !combinedFields.includes(field));
    const charts = [];
    if (combinedAvailable) {
      charts.push(
        buildMultiLineChart(
          labels,
          combinedFields.map((field) => ({
            name: field,
            values: table.rows.map((row) => Number(row[field])).filter((value) => Number.isFinite(value))
          })),
          "采购合同个数 / 采购单个数"
        )
      );
    }
    remainingFields.slice(0, 2).forEach((field) => {
      const values = table.rows.map((row) => Number(row[field])).filter((value) => Number.isFinite(value));
      charts.push(buildLineChart(labels, values, field));
    });
    return charts.join("");
  }
  const chartFields = table.numericFields.slice(0, 4);
  if (!chartFields.length) {
    return `<div class="empty-state">当前表没有适合绘图的数值字段。</div>`;
  }
  return chartFields.map((field) => {
    const values = table.rows.map((row) => Number(row[field])).filter((value) => Number.isFinite(value));
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
  const labels = table.rows.map((row) => row[table.weekField]);
  const chartFields = table.numericFields.slice(0, 4);
  if (!chartFields.length) return "";
  return `
    <div class="dashboard-block">
      <h4 class="dashboard-block-title">趋势图</h4>
      <div class="chart-grid">${chartFields.map((field) => {
        const values = table.rows.map((row) => Number(row[field])).filter((value) => Number.isFinite(value));
        return buildLineChart(labels, values, field);
      }).join("")}</div>
    </div>
  `;
}

function renderOverview() {
  const metrics = buildOverviewMetrics(state.workbook);
  if (!metrics.length) {
    overviewGrid.innerHTML = `<div class="empty-state">当前没有识别到可汇总的周维度 KPI。</div>`;
    return;
  }
  overviewGrid.className = "kpi-grid overview-grid";
  overviewGrid.innerHTML = metrics.map((metric) => {
    const change = formatChange(metric.delta, metric.field);
    return `
      <div class="kpi-card">
        <div class="kpi-title">${metric.title}</div>
        <div class="kpi-value">${formatValue(metric.latestValue, metric.field)}</div>
        <div class="kpi-change ${change.cls}">${metric.week} ${change.text}</div>
      </div>
    `;
  }).join("");
}

function renderDashboardContent() {
  const tables = state.workbook.tables;
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
