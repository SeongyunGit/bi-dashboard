import { parseWorkbook } from "./parser.js";
import { createState } from "./state.js";
import { createUI } from "./ui.js";

const KEYWORDS = {
  progress: "\uC9C4\uD589\uD604\uD669",
  execution: "\uC218\uD589\uD604\uD669",
  quality: "\uD488\uC9C8\uD604\uD669",
  defect: "\uACB0\uD568\uBC1C\uC0DD\uD604\uD669",
  system: "\uC2DC\uC2A4\uD15C",
  part: "\uD30C\uD2B8",
  total: "\uCD1D \uC218\uB7C9",
  planned: "\uACC4\uD68D \uC218\uB7C9",
  completed: "\uC644\uB8CC \uC218\uB7C9",
  planRate: "\uACC4\uD68D\uBAA9\uD45C\uB960",
  actualRate: "\uC2E4\uC81C\uC9C4\uCC99\uB960",
  planVsActual: "\uACC4\uD68D \uB300\uBE44 \uC2E4\uC81C",
  executed: "\uB204\uC801 \uC218\uD589 \uC218",
  successAccum: "\uB204\uC801 \uC131\uACF5\uC218",
  success: "\uC131\uACF5\uC218",
  fail: "\uC2E4\uD328\uC218",
  pending: "\uBBF8\uC218\uD589\uC218",
  successRate: "\uC131\uACF5\uB960",
  defect: "\uACB0\uD568\uC218",
  defectRate: "\uACB0\uD568\uB960",
};

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value === "string") {
    const parsed = Number(value.replace(/,/g, "").trim());
    return Number.isFinite(parsed) ? parsed : 0;
  }

  return 0;
}

function toPercent(value) {
  const numeric = toNumber(value);
  return numeric >= -1 && numeric <= 1 ? numeric * 100 : numeric;
}

function clampPercent(value) {
  return Math.max(0, Math.min(100, value));
}

function formatPercent(value) {
  return `${value.toFixed(1)}%`;
}

function findColumn(columns, keyword) {
  return columns.find((column) => String(column).includes(keyword)) ?? "";
}

function buildLabel(row, systemKey, partKey, index) {
  const system = String(row[systemKey] ?? "").trim();
  const part = String(row[partKey] ?? "").trim();

  if (system && part) {
    return `${system} / ${part}`;
  }
  if (part) {
    return part;
  }
  if (system) {
    return system;
  }
  return `Row ${index + 1}`;
}

function createCard(label, value, tone = "default") {
  return { label, value: String(value), tone };
}

function createBarDataset(label, data, color) {
  return {
    label,
    data,
    backgroundColor: color,
    borderRadius: 8,
    maxBarThickness: 18,
  };
}

function createLineDataset(label, data, color) {
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: color,
    pointRadius: 3,
    pointHoverRadius: 4,
    tension: 0.32,
  };
}

function summarizeRows(sheet, fields) {
  const systemKey = findColumn(sheet.headers, KEYWORDS.system);
  const partKey = findColumn(sheet.headers, KEYWORDS.part);

  return sheet.rows
    .map((row, index) => {
      const summary = {
        label: buildLabel(row, systemKey, partKey, index),
        planRate: toPercent(row[findColumn(sheet.headers, KEYWORDS.planRate)]),
        actualRate: toPercent(row[findColumn(sheet.headers, KEYWORDS.actualRate)]),
        successRate: toPercent(
          row[findColumn(sheet.headers, KEYWORDS.successRate)],
        ),
        defectRate: toPercent(row[findColumn(sheet.headers, KEYWORDS.defectRate)]),
        gap: toPercent(row[findColumn(sheet.headers, KEYWORDS.planVsActual)]),
      };

      Object.entries(fields).forEach(([name, keyword]) => {
        summary[name] = toNumber(row[findColumn(sheet.headers, keyword)]);
      });

      return summary;
    })
    .filter((row) => row.total > 0);
}

function isDetailSheet(sheet) {
  const hasId = sheet.headers.some((header) => String(header).includes("ID"));
  const hasResult = sheet.headers.some((header) =>
    String(header).includes("\uACB0\uACFC"),
  );
  const hasTotal = sheet.headers.some((header) =>
    String(header).includes("\uCD1D \uC218\uB7C9"),
  );

  return hasId && hasResult && hasTotal;
}

function buildDetailSheetFallback(allSheets) {
  return allSheets
    .filter((sheet) => isDetailSheet(sheet))
    .map((sheet) => {
      const totalKey = findColumn(sheet.headers, KEYWORDS.total);
      const completedKey = findColumn(sheet.headers, KEYWORDS.completed);
      const resultKey =
        sheet.headers.find((header) => String(header).includes("\uACB0\uACFC")) ??
        "";

      const summary = sheet.rows.reduce(
        (acc, row) => {
          const total = toNumber(row[totalKey]);
          const completed = toNumber(row[completedKey]);
          const result = String(row[resultKey] ?? "").trim();

          acc.total += total;
          acc.completed += completed;

          if (result.includes("\uC131\uACF5")) {
            acc.success += completed || total;
          } else if (result.includes("\uC2E4\uD328")) {
            acc.fail += completed || total || 1;
          } else {
            acc.pending += Math.max(0, total - completed);
          }

          return acc;
        },
        { total: 0, completed: 0, success: 0, fail: 0, pending: 0 },
      );

      const label = sheet.name
        .replace(/\(.+?\)/g, "")
        .replace(/_/g, " ")
        .trim();

      return {
        label,
        total: summary.total,
        planned: summary.total,
        completed: summary.completed,
        executed: summary.completed,
        success: summary.success,
        fail: summary.fail,
        pending: summary.pending || Math.max(0, summary.total - summary.completed),
        planRate: summary.total > 0 ? (summary.total / summary.total) * 100 : 0,
        actualRate:
          summary.total > 0 ? (summary.completed / summary.total) * 100 : 0,
        successRate:
          summary.total > 0 ? (summary.success / summary.total) * 100 : 0,
        defect: summary.fail,
        defectRate:
          summary.success > 0 ? (summary.fail / summary.success) * 100 : 0,
        gap:
          summary.total > 0
            ? ((summary.total - summary.completed) / summary.total) * 100
            : 0,
      };
    })
    .filter((row) => row.total > 0);
}

function buildProgressDashboard(sheet, allSheets) {
  let rows = summarizeRows(sheet, {
    total: KEYWORDS.total,
    planned: KEYWORDS.planned,
    completed: KEYWORDS.completed,
  });

  if (rows.length === 0 || rows.every((row) => row.total === 0)) {
    rows = buildDetailSheetFallback(allSheets);
  }

  rows = rows.slice(0, 14);

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const planned = rows.reduce((sum, row) => sum + row.planned, 0);
  const completed = rows.reduce((sum, row) => sum + row.completed, 0);
  const remaining = Math.max(0, total - completed);

  return {
    title: "Progress Dashboard",
    subtitle: "Plan, actual progress, completion, and gap in one place",
    cards: [
      createCard("Total Cases", total, "primary"),
      createCard("Planned", planned, "secondary"),
      createCard("Completed", completed, "accent"),
      createCard("Remaining", remaining, "muted"),
    ],
    charts: [
      {
        key: "progress-volume",
        title: "Planned vs Completed Cases",
        type: "bar",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createBarDataset("Planned", rows.map((row) => row.planned), "#60a5fa"),
            createBarDataset(
              "Completed",
              rows.map((row) => row.completed),
              "#22c55e",
            ),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "progress-rate",
        title: "Plan Rate vs Actual Rate",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => row.label),
          datasets: [
            {
              label: "Plan Rate",
              data: rows.slice(0, 8).map((row) => clampPercent(row.planRate)),
              borderColor: "#60a5fa",
              backgroundColor: "rgba(96,165,250,0.18)",
            },
            {
              label: "Actual Rate",
              data: rows.slice(0, 8).map((row) => clampPercent(row.actualRate)),
              borderColor: "#22c55e",
              backgroundColor: "rgba(34,197,94,0.18)",
            },
          ],
        },
        options: radarOptions(),
      },
      {
        key: "progress-donut",
        title: "Completed vs Remaining",
        type: "doughnut",
        data: {
          labels: ["Completed", "Remaining"],
          datasets: [
            {
              data: [completed, remaining],
              backgroundColor: ["#22c55e", "#334155"],
              borderWidth: 0,
            },
          ],
        },
        options: doughnutOptions(),
      },
      {
        key: "progress-gap",
        title: "Progress Gap Trend",
        type: "line",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createLineDataset(
              "Gap",
              rows.map((row) => clampPercent(row.gap)),
              "#f59e0b",
            ),
          ],
        },
        options: lineOptions("%"),
      },
    ],
  };
}

function buildExecutionDashboard(sheet, allSheets) {
  let rows = summarizeRows(sheet, {
    total: KEYWORDS.total,
    executed: KEYWORDS.executed,
    success: KEYWORDS.successAccum,
    fail: KEYWORDS.fail,
    pending: KEYWORDS.pending,
  });

  if (rows.length === 0 || rows.every((row) => row.total === 0)) {
    rows = buildDetailSheetFallback(allSheets);
  }

  rows = rows
    .map((row) => ({
      ...row,
      pending: row.pending || Math.max(0, row.total - row.executed),
    }))
    .slice(0, 14);

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const executed = rows.reduce((sum, row) => sum + row.executed, 0);
  const fail = rows.reduce((sum, row) => sum + row.fail, 0);
  const pending = rows.reduce((sum, row) => sum + row.pending, 0);

  return {
    title: "Execution Dashboard",
    subtitle: "Execution, failures, pending load, and completion rate",
    cards: [
      createCard("Total Cases", total, "primary"),
      createCard("Executed", executed, "secondary"),
      createCard("Failed", fail, "danger"),
      createCard("Pending", pending, "muted"),
    ],
    charts: [
      {
        key: "execution-stack",
        title: "Execution Distribution by Part",
        type: "bar",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createBarDataset(
              "Executed",
              rows.map((row) => row.executed),
              "#60a5fa",
            ),
            createBarDataset("Failed", rows.map((row) => row.fail), "#f87171"),
            createBarDataset("Pending", rows.map((row) => row.pending), "#475569"),
          ],
        },
        options: stackedBarOptions(),
      },
      {
        key: "execution-donut",
        title: "Overall Execution Mix",
        type: "doughnut",
        data: {
          labels: ["Executed", "Failed", "Pending"],
          datasets: [
            {
              data: [executed, fail, pending],
              backgroundColor: ["#60a5fa", "#f87171", "#475569"],
              borderWidth: 0,
            },
          ],
        },
        options: doughnutOptions(),
      },
      {
        key: "execution-line",
        title: "Failure Line",
        type: "line",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createLineDataset("Failed", rows.map((row) => row.fail), "#f87171"),
          ],
        },
        options: lineOptions(""),
      },
      {
        key: "execution-rate",
        title: "Execution Rate",
        type: "bar",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createBarDataset(
              "Execution Rate",
              rows.map((row) => (row.total > 0 ? (row.executed / row.total) * 100 : 0)),
              "#22c55e",
            ),
          ],
        },
        options: percentBarOptions(),
      },
    ],
  };
}

function buildQualityDashboard(sheet, allSheets) {
  let rows = summarizeRows(sheet, {
    total: KEYWORDS.total,
    completed: KEYWORDS.completed,
    success: KEYWORDS.success,
    fail: KEYWORDS.fail,
  });

  if (rows.length === 0 || rows.every((row) => row.total === 0)) {
    rows = buildDetailSheetFallback(allSheets);
  }

  rows = rows.slice(0, 14);

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const completed = rows.reduce((sum, row) => sum + row.completed, 0);
  const success = rows.reduce((sum, row) => sum + row.success, 0);
  const fail = rows.reduce((sum, row) => sum + row.fail, 0);

  return {
    title: "Quality Dashboard",
    subtitle: "Pass/fail quality view with completion and quality rates",
    cards: [
      createCard("Total Cases", total, "primary"),
      createCard("Completed", completed, "secondary"),
      createCard("Quality Pass", success, "accent"),
      createCard("Quality Fail", fail, "danger"),
    ],
    charts: [
      {
        key: "quality-bar",
        title: "Quality Pass vs Fail",
        type: "bar",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createBarDataset("Pass", rows.map((row) => row.success), "#22c55e"),
            createBarDataset("Fail", rows.map((row) => row.fail), "#f87171"),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "quality-donut",
        title: "Quality Result Mix",
        type: "doughnut",
        data: {
          labels: ["Pass", "Fail"],
          datasets: [
            {
              data: [success, fail],
              backgroundColor: ["#22c55e", "#f87171"],
              borderWidth: 0,
            },
          ],
        },
        options: doughnutOptions(),
      },
      {
        key: "quality-radar",
        title: "Quality Rate vs Completion Rate",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => row.label),
          datasets: [
            {
              label: "Quality Rate",
              data: rows.slice(0, 8).map((row) => clampPercent(row.successRate)),
              borderColor: "#f59e0b",
              backgroundColor: "rgba(245,158,11,0.18)",
            },
            {
              label: "Completion Rate",
              data: rows
                .slice(0, 8)
                .map((row) => (row.total > 0 ? (row.completed / row.total) * 100 : 0)),
              borderColor: "#60a5fa",
              backgroundColor: "rgba(96,165,250,0.16)",
            },
          ],
        },
        options: radarOptions(),
      },
      {
        key: "quality-line",
        title: "Quality Rate Trend",
        type: "line",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createLineDataset(
              "Quality Rate",
              rows.map((row) => clampPercent(row.successRate)),
              "#f59e0b",
            ),
          ],
        },
        options: lineOptions("%"),
      },
    ],
  };
}

function buildDefectDashboard(sheet, allSheets) {
  let rows = summarizeRows(sheet, {
    total: KEYWORDS.total,
    success: KEYWORDS.success,
    defect: KEYWORDS.defect,
  });

  if (rows.length === 0 || rows.every((row) => row.total === 0)) {
    rows = buildDetailSheetFallback(allSheets);
  }

  rows = rows.slice(0, 14);

  const total = rows.reduce((sum, row) => sum + row.total, 0);
  const success = rows.reduce((sum, row) => sum + row.success, 0);
  const defect = rows.reduce((sum, row) => sum + row.defect, 0);

  return {
    title: "Defect Dashboard",
    subtitle: "Defects, defect rate, and part-level defect concentration",
    cards: [
      createCard("Total Cases", total, "primary"),
      createCard("Successful Runs", success, "secondary"),
      createCard("Defects", defect, "danger"),
      createCard(
        "Defect Rate",
        success > 0 ? formatPercent((defect / success) * 100) : "0.0%",
        "accent",
      ),
    ],
    charts: [
      {
        key: "defect-bar",
        title: "Defects by Part",
        type: "bar",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createBarDataset("Defects", rows.map((row) => row.defect), "#f87171"),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "defect-line",
        title: "Defect Rate Trend",
        type: "line",
        data: {
          labels: rows.map((row) => row.label),
          datasets: [
            createLineDataset(
              "Defect Rate",
              rows.map((row) => clampPercent(row.defectRate)),
              "#fb7185",
            ),
          ],
        },
        options: lineOptions("%"),
      },
      {
        key: "defect-donut",
        title: "Success vs Defect Mix",
        type: "doughnut",
        data: {
          labels: ["Success", "Defect"],
          datasets: [
            {
              data: [success, defect],
              backgroundColor: ["#22c55e", "#f87171"],
              borderWidth: 0,
            },
          ],
        },
        options: doughnutOptions(),
      },
      {
        key: "defect-radar",
        title: "Defect Rate vs Defect Load",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => row.label),
          datasets: [
            {
              label: "Defect Rate",
              data: rows.slice(0, 8).map((row) => clampPercent(row.defectRate)),
              borderColor: "#fb7185",
              backgroundColor: "rgba(251,113,133,0.18)",
            },
            {
              label: "Defect Load",
              data: rows
                .slice(0, 8)
                .map((row) => (row.total > 0 ? (row.defect / row.total) * 100 : 0)),
              borderColor: "#60a5fa",
              backgroundColor: "rgba(96,165,250,0.14)",
            },
          ],
        },
        options: radarOptions(),
      },
    ],
  };
}

function baseChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: "#cbd5e1",
          boxWidth: 12,
        },
      },
    },
    scales: {
      x: {
        ticks: { color: "#94a3b8" },
        grid: { color: "rgba(148,163,184,0.08)" },
      },
      y: {
        ticks: { color: "#94a3b8" },
        grid: { color: "rgba(148,163,184,0.08)" },
      },
    },
  };
}

function groupedBarOptions() {
  return baseChartOptions();
}

function stackedBarOptions() {
  const options = baseChartOptions();
  options.scales.x.stacked = true;
  options.scales.y.stacked = true;
  return options;
}

function percentBarOptions() {
  const options = baseChartOptions();
  options.scales.y.min = 0;
  options.scales.y.max = 100;
  return options;
}

function lineOptions(suffix) {
  const options = baseChartOptions();
  options.scales.y.beginAtZero = true;
  if (suffix) {
    options.scales.y.ticks.callback = (value) => `${value}${suffix}`;
  }
  return options;
}

function doughnutOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    cutout: "70%",
    plugins: {
      legend: {
        position: "bottom",
        labels: {
          color: "#cbd5e1",
          boxWidth: 12,
        },
      },
    },
  };
}

function radarOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: "#cbd5e1",
        },
      },
    },
    scales: {
      r: {
        min: 0,
        max: 100,
        angleLines: { color: "rgba(148,163,184,0.12)" },
        grid: { color: "rgba(148,163,184,0.12)" },
        pointLabels: { color: "#94a3b8", font: { size: 10 } },
        ticks: {
          color: "#94a3b8",
          backdropColor: "transparent",
        },
      },
    },
  };
}

function buildDashboard(sheet, allSheets) {
  if (sheet.name.includes(KEYWORDS.progress)) {
    return buildProgressDashboard(sheet, allSheets);
  }
  if (sheet.name.includes(KEYWORDS.execution)) {
    return buildExecutionDashboard(sheet, allSheets);
  }
  if (sheet.name.includes(KEYWORDS.quality)) {
    return buildQualityDashboard(sheet, allSheets);
  }
  if (sheet.name.includes(KEYWORDS.defect)) {
    return buildDefectDashboard(sheet, allSheets);
  }
  return null;
}

function downloadJson(fileName, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function selectSheetData(sheets, sheetName) {
  const selectedSheet =
    sheets.find((sheet) => sheet.name === sheetName) ?? sheets[0] ?? null;

  return {
    selectedSheetName: selectedSheet?.name ?? "",
    tableColumns: selectedSheet?.headers ?? [],
    tableRows: selectedSheet?.rows ?? [],
    dashboard: selectedSheet ? buildDashboard(selectedSheet, sheets) : null,
  };
}

function createDashboardApp(options = {}) {
  const state = createState();
  const ui = createUI(options.ui);
  let selectedFile = null;

  async function processFile(file) {
    if (!file) {
      state.setState({
        error: "No file selected",
        status: "error",
      });
      return;
    }

    state.setState((currentState) => ({
      ...currentState,
      status: "loading",
      error: null,
      fileName: file.name,
    }));

    try {
      const workbook = await parseWorkbook(file);
      const selection = selectSheetData(workbook.sheets, workbook.sheets[0]?.name);

      state.setState((currentState) => ({
        ...currentState,
        workbook,
        sheets: workbook.sheets,
        ...selection,
        status: "ready",
        error: null,
      }));
    } catch (error) {
      state.setState((currentState) => ({
        ...currentState,
        status: "error",
        error: error instanceof Error ? error.message : String(error),
      }));
    }
  }

  function handleFileChange(event) {
    const file = event.target?.files?.[0] ?? null;
    selectedFile = file;

    state.setState((currentState) => ({
      ...currentState,
      fileName: file?.name ?? "",
      status: file ? "selected" : "idle",
      error: null,
    }));
  }

  function handleSheetSelect(event) {
    const sheetName = event.target?.value ?? "";

    state.setState((currentState) => ({
      ...currentState,
      ...selectSheetData(currentState.sheets, sheetName),
    }));
  }

  async function handleParse() {
    await processFile(selectedFile);
  }

  function handleDownload() {
    const currentState = state.getState();

    if (currentState.tableRows.length === 0) {
      state.setState((previousState) => ({
        ...previousState,
        error: "No sheet data to download",
      }));
      return;
    }

    const sourceName = currentState.fileName || "dashboard-data";
    const safeSheetName = currentState.selectedSheetName || "sheet";
    const fileName =
      sourceName.replace(/\.[^.]+$/, "") + `-${safeSheetName}.json`;

    downloadJson(fileName, currentState.tableRows);
  }

  function handleReset() {
    selectedFile = null;

    if (ui.elements.fileInput) {
      ui.elements.fileInput.value = "";
    }

    state.resetState();
  }

  ui.bindEvents({
    onFileChange: handleFileChange,
    onParse: handleParse,
    onSheetSelect: handleSheetSelect,
    onDownload: handleDownload,
    onReset: handleReset,
  });

  state.subscribe((nextState) => {
    ui.render(nextState);
  });

  return {
    state,
    processFile,
  };
}

export { createDashboardApp };

export function initDashboardApp(options) {
  return createDashboardApp(options);
}

if (typeof window !== "undefined") {
  window.addEventListener("DOMContentLoaded", () => {
    initDashboardApp();
  });
}
