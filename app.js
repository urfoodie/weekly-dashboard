const demoWorkbook = {
  workbookName: "示例采购经营周报",
  tables: [
    {
      id: "purchase-stats",
      sheet: "采购数据统计",
      title: "采购数据统计",
      headers: ["周次", "采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）", "金额周环比"],
      weekField: "周次",
      numericFields: ["采购合同个数", "采购单个数", "销售采购单个数", "直发采购单个数", "直发采购单比例", "采购金额（元）", "人数", "人均采购金额（元）", "人均采购单（个）"],
      rows: [
        { "周次": "7.09", "采购合同个数": 1012, "采购单个数": 2320, "销售采购单个数": 1100, "直发采购单个数": 488, "直发采购单比例": 0.443, "采购金额（元）": 3980112.2, "人数": 15, "人均采购金额（元）": 265340.81, "人均采购单（个）": 154.7, "金额周环比": null },
        { "周次": "7.16", "采购合同个数": 1028, "采购单个数": 2408, "销售采购单个数": 1126, "直发采购单个数": 502, "直发采购单比例": 0.447, "采购金额（元）": 4042050.4, "人数": 15, "人均采购金额（元）": 269470.03, "人均采购单（个）": 160.5, "金额周环比": 0.016 },
        { "周次": "7.23", "采购合同个数": 1039, "采购单个数": 2486, "销售采购单个数": 1169, "直发采购单个数": 521, "直发采购单比例": 0.446, "采购金额（元）": 4141198.82, "人数": 15, "人均采购金额（元）": 276079.92, "人均采购单（个）": 165.7, "金额周环比": 0.025 },
        { "周次": "7.30", "采购合同个数": 1169, "采购单个数": 3698, "销售采购单个数": 1138, "直发采购单个数": 515, "直发采购单比例": 0.453, "采购金额（元）": 5760008.07, "人数": 15, "人均采购金额（元）": 384000.54, "人均采购单（个）": 246.5, "金额周环比": 0.391 }
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
        { "周次": "7.09", "货号个数": 2281, "断货个数": 44, "断货率": 0.0193, "备注说明": "" },
        { "周次": "7.16", "货号个数": 2270, "断货个数": 47, "断货率": 0.0207, "备注说明": "" },
        { "周次": "7.23", "货号个数": 2264, "断货个数": 48, "断货率": 0.0212, "备注说明": "" },
        { "周次": "7.30", "货号个数": 2248, "断货个数": 69, "断货率": 0.0307, "备注说明": "阈值调整" }
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
        { "周次": "7.09", "单笔议价金额（元）": 4880.2, "备货议价金额（元）": 50210.4, "议价合计（元）": 55090.6 },
        { "周次": "7.16", "单笔议价金额（元）": 4955.4, "备货议价金额（元）": 54105.2, "议价合计（元）": 59060.6 },
        { "周次": "7.23", "单笔议价金额（元）": 4971.97, "备货议价金额（元）": 56502.88, "议价合计（元）": 61474.85 },
        { "周次": "7.30", "单笔议价金额（元）": 5210.54, "备货议价金额（元）": 25998.03, "议价合计（元）": 31208.57 }
      ]
    },
    {
      id: "inventory-turnover",
      sheet: "库存周转分析",
      title: "库存周转分析（月度对比）",
      headers: ["仓库/品类", "6月库存金额", "6月平均周转天数", "7月库存金额", "7月平均周转天数", "库存金额环比变化", "周转天数变化"],
      weekField: null,
      numericFields: ["6月库存金额", "6月平均周转天数", "7月库存金额", "7月平均周转天数", "库存金额环比变化", "周转天数变化"],
      rows: [
        { "仓库/品类": "主仓有阈值的明星产品", "6月库存金额": 12960913.76, "6月平均周转天数": 79.1, "7月库存金额": 14129745.66, "7月平均周转天数": 90.1, "库存金额环比变化": 0.09, "周转天数变化": 11 },
        { "仓库/品类": "主仓所有商品", "6月库存金额": 32900585.42, "6月平均周转天数": 89.1, "7月库存金额": 34198877.49, "7月平均周转天数": 94.7, "库存金额环比变化": 0.039, "周转天数变化": 5.6 }
      ]
    }
  ]
};

const state = { workbook: demoWorkbook };
const fileInput = document.getElementById("fileInput");
const demoButton = document.getElementById("demoButton");
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
  const blocks = [];
  let current = [];
  rows.forEach((row) => {
    const hasValue = row.some((cell) => !isEmpty(cell));
    if (hasValue) {
      current.push(row);
    } else if (current.length) {
      blocks.push(current);
      current = [];
    }
  });
  if (current.length) blocks.push(current);
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
  return workbook.tables
    .filter((table) => table.weekField && table.rows.length)
    .slice(0, 3)
    .flatMap((table) =>
      computeWeeklyMetrics(table).slice(0, 2).map((metric) => ({
        ...metric,
        title: `${table.title} · ${metric.field}`
      }))
    );
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

function buildTableHtml(table) {
  const headerHtml = table.headers.map((header) => `<th>${header}</th>`).join("");
  const rowsHtml = table.rows.map((row) => `
    <tr>${table.headers.map((header) => `<td>${formatValue(row[header], header)}</td>`).join("")}</tr>
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
  if (!table.weekField) {
    return `<div class="empty-state">当前表没有周字段，跳过趋势图。</div>`;
  }
  const labels = table.rows.map((row) => row[table.weekField]);
  const chartFields = table.numericFields.slice(0, 4);
  if (!chartFields.length) {
    return `<div class="empty-state">当前表没有适合绘图的数值字段。</div>`;
  }
  return chartFields.map((field) => {
    const values = table.rows.map((row) => Number(row[field])).filter((value) => Number.isFinite(value));
    return buildLineChart(labels, values, field);
  }).join("");
}

function renderOverview() {
  const metrics = buildOverviewMetrics(state.workbook);
  if (!metrics.length) {
    overviewGrid.innerHTML = `<div class="empty-state">当前没有识别到可汇总的周维度 KPI。</div>`;
    return;
  }
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
    const weeklyMetrics = computeWeeklyMetrics(table);
    const metricsHtml = weeklyMetrics.length
      ? `<div class="kpi-grid">${weeklyMetrics.slice(0, 4).map((metric) => {
          const change = formatChange(metric.delta, metric.field);
          return `
            <div class="kpi-card">
              <div class="kpi-title">${metric.field}</div>
              <div class="kpi-value">${formatValue(metric.latestValue, metric.field)}</div>
              <div class="kpi-change ${change.cls}">${metric.week} ${change.text}</div>
            </div>
          `;
        }).join("")}</div>`
      : `<div class="empty-state">当前表没有周维度 KPI。</div>`;
    const meta = table.weekField ? `已识别周字段：${table.weekField} · 共 ${table.rows.length} 行` : `未识别周字段 · 共 ${table.rows.length} 行`;
    return `
      <article class="dashboard-section">
        <div class="dashboard-section-header">
          <h3 class="dashboard-section-title">${table.title}</h3>
          <p class="dashboard-section-meta">${meta}</p>
        </div>
        <div class="dashboard-block">
          <h4 class="dashboard-block-title">指标卡片</h4>
          ${metricsHtml}
        </div>
        <div class="dashboard-block">
          <h4 class="dashboard-block-title">趋势图</h4>
          <div class="chart-grid">${buildChartsHtml(table)}</div>
        </div>
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

renderAll();
