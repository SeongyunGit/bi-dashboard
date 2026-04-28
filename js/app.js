import { parseWorkbook } from "./parser.js";
import { createState } from "./state.js";
import { createUI } from "./ui.js";

const KO = {
  progressSheet: "\uB2E8\uC704\uD14C\uC2A4\uD2B8 \uC9C4\uD589\uD604\uD669",
  executionSheet: "\uB2E8\uC704\uD14C\uC2A4\uD2B8 \uC218\uD589\uD604\uD669",
  qualitySheet: "\uB2E8\uC704\uD14C\uC2A4\uD2B8 \uD488\uC9C8\uD604\uD669",
  defectSheet:
    "\uB2E8\uC704\uD14C\uC2A4\uD2B8 \uACB0\uD568\uBC1C\uC0DD\uD604\uD669",
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
  fail: "\uC2E4\uD328\uC218",
  pending: "\uBBF8\uC218\uD589\uC218",
  success: "\uC131\uACF5\uC218",
  successRate: "\uC131\uACF5\uB960",
  defect: "\uACB0\uD568\uC218",
  defectRate: "\uACB0\uD568\uB960",
};

const SPECIAL_SHEET_BUILDERS = [
  { match: KO.progressSheet, build: buildProgressVisualization },
  { match: KO.executionSheet, build: buildExecutionVisualization },
  { match: KO.qualitySheet, build: buildQualityVisualization },
  { match: KO.defectSheet, build: buildDefectVisualization },
];

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value === "string") {
    const normalized = value.replace(/,/g, "").trim();
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  return 0;
}

function asPercent(value) {
  const numeric = toNumber(value);
  if (numeric <= 1 && numeric >= -1) {
    return numeric * 100;
  }
  return numeric;
}

function safeLabel(row, systemKey, partKey, index) {
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

function findColumn(columns, keywords) {
  return (
    columns.find((column) =>
      keywords.every((keyword) => String(column).includes(keyword)),
    ) ?? ""
  );
}

function metric(label, value) {
  return { label, value: String(value) };
}

function buildVisualization(sheet) {
  const builder = SPECIAL_SHEET_BUILDERS.find((item) =>
    sheet.name.includes(item.match),
  );
  return builder ? builder.build(sheet) : null;
}

function selectSheetData(sheets, sheetName) {
  const selectedSheet =
    sheets.find((sheet) => sheet.name === sheetName) ?? sheets[0] ?? null;

  return {
    selectedSheetName: selectedSheet?.name ?? "",
    tableColumns: selectedSheet?.headers ?? [],
    tableRows: selectedSheet?.rows ?? [],
    visualization: selectedSheet ? buildVisualization(selectedSheet) : null,
  };
}

function buildProgressVisualization(sheet) {
  const systemKey = findColumn(sheet.headers, [KO.system]);
  const partKey = findColumn(sheet.headers, [KO.part]);
  const totalKey = findColumn(sheet.headers, [KO.total]);
  const plannedKey = findColumn(sheet.headers, [KO.planned]);
  const completedKey = findColumn(sheet.headers, [KO.completed]);
  const planRateKey = findColumn(sheet.headers, [KO.planRate]);
  const actualRateKey = findColumn(sheet.headers, [KO.actualRate]);
  const gapKey = findColumn(sheet.headers, [KO.planVsActual]);

  const rows = sheet.rows
    .map((row, index) => ({
      label: safeLabel(row, systemKey, partKey, index),
      total: toNumber(row[totalKey]),
      planned: toNumber(row[plannedKey]),
      completed: toNumber(row[completedKey]),
      planRate: asPercent(row[planRateKey]),
      actualRate: asPercent(row[actualRateKey]),
      gap: asPercent(row[gapKey]),
    }))
    .filter((row) => row.total > 0);

  const totals = rows.reduce(
    (acc, row) => {
      acc.total += row.total;
      acc.planned += row.planned;
      acc.completed += row.completed;
      return acc;
    },
    { total: 0, planned: 0, completed: 0 },
  );

  return {
    type: "progress",
    title: "Progress Overview",
    accent: "progress",
    metrics: [
      metric("Total Cases", totals.total),
      metric("Planned", totals.planned),
      metric("Completed", totals.completed),
      metric(
        "Completion Rate",
        totals.total > 0
          ? `${((totals.completed / totals.total) * 100).toFixed(1)}%`
          : "0.0%",
      ),
    ],
    chartTitle: "Plan vs Actual Progress",
    rows: rows
      .sort((a, b) => b.total - a.total)
      .slice(0, 10)
      .map((row) => ({
        label: row.label,
        primary: Math.max(0, Math.min(100, row.planRate)),
        secondary: Math.max(0, Math.min(100, row.actualRate)),
        detail: `Gap ${row.gap.toFixed(1)}%`,
      })),
    legend: [
      { label: "Plan", tone: "primary" },
      { label: "Actual", tone: "secondary" },
    ],
  };
}

function buildExecutionVisualization(sheet) {
  const systemKey = findColumn(sheet.headers, [KO.system]);
  const partKey = findColumn(sheet.headers, [KO.part]);
  const totalKey = findColumn(sheet.headers, [KO.total]);
  const executedKey = findColumn(sheet.headers, [KO.executed]);
  const successKey = findColumn(sheet.headers, [KO.successAccum]);
  const failKey = findColumn(sheet.headers, [KO.fail]);
  const pendingKey = findColumn(sheet.headers, [KO.pending]);

  const rows = sheet.rows
    .map((row, index) => {
      const total = toNumber(row[totalKey]);
      const executed = toNumber(row[executedKey]);
      const fail = toNumber(row[failKey]);
      const pending = pendingKey
        ? toNumber(row[pendingKey])
        : Math.max(0, total - executed);

      return {
        label: safeLabel(row, systemKey, partKey, index),
        total,
        executed,
        success: toNumber(row[successKey]),
        fail,
        pending,
      };
    })
    .filter((row) => row.total > 0);

  const totals = rows.reduce(
    (acc, row) => {
      acc.total += row.total;
      acc.executed += row.executed;
      acc.success += row.success;
      acc.fail += row.fail;
      acc.pending += row.pending;
      return acc;
    },
    { total: 0, executed: 0, success: 0, fail: 0, pending: 0 },
  );

  return {
    type: "execution",
    title: "Execution Overview",
    accent: "execution",
    metrics: [
      metric("Total Cases", totals.total),
      metric("Executed", totals.executed),
      metric("Failed", totals.fail),
      metric("Pending", totals.pending),
    ],
    chartTitle: "Execution Distribution",
    rows: rows
      .sort((a, b) => b.total - a.total)
      .slice(0, 10)
      .map((row) => ({
        label: row.label,
        segments: [
          { label: "Executed", value: row.executed, tone: "primary" },
          { label: "Failed", value: row.fail, tone: "danger" },
          { label: "Pending", value: row.pending, tone: "muted" },
        ],
        total: row.total,
      })),
    legend: [
      { label: "Executed", tone: "primary" },
      { label: "Failed", tone: "danger" },
      { label: "Pending", tone: "muted" },
    ],
  };
}

function buildQualityVisualization(sheet) {
  const systemKey = findColumn(sheet.headers, [KO.system]);
  const partKey = findColumn(sheet.headers, [KO.part]);
  const totalKey = findColumn(sheet.headers, [KO.total]);
  const completedKey = findColumn(sheet.headers, [KO.completed]);
  const successKey = findColumn(sheet.headers, [KO.success]);
  const failKey = findColumn(sheet.headers, [KO.fail]);
  const successRateKey = findColumn(sheet.headers, [KO.successRate]);

  const rows = sheet.rows
    .map((row, index) => ({
      label: safeLabel(row, systemKey, partKey, index),
      total: toNumber(row[totalKey]),
      completed: toNumber(row[completedKey]),
      success: toNumber(row[successKey]),
      fail: toNumber(row[failKey]),
      successRate: asPercent(row[successRateKey]),
    }))
    .filter((row) => row.total > 0);

  const totals = rows.reduce(
    (acc, row) => {
      acc.total += row.total;
      acc.completed += row.completed;
      acc.success += row.success;
      acc.fail += row.fail;
      return acc;
    },
    { total: 0, completed: 0, success: 0, fail: 0 },
  );

  return {
    type: "quality",
    title: "Quality Overview",
    accent: "quality",
    metrics: [
      metric("Total Cases", totals.total),
      metric("Completed", totals.completed),
      metric("Quality Pass", totals.success),
      metric("Quality Fail", totals.fail),
    ],
    chartTitle: "Quality Rate by Part",
    rows: rows
      .sort((a, b) => b.total - a.total)
      .slice(0, 10)
      .map((row) => ({
        label: row.label,
        primary: Math.max(0, Math.min(100, row.successRate)),
        secondary:
          row.total > 0
            ? Math.max(0, Math.min(100, (row.completed / row.total) * 100))
            : 0,
        detail: `Pass ${row.success} / Fail ${row.fail}`,
      })),
    legend: [
      { label: "Quality Rate", tone: "primary" },
      { label: "Completion", tone: "secondary" },
    ],
  };
}

function buildDefectVisualization(sheet) {
  const systemKey = findColumn(sheet.headers, [KO.system]);
  const partKey = findColumn(sheet.headers, [KO.part]);
  const totalKey = findColumn(sheet.headers, [KO.total]);
  const successKey = findColumn(sheet.headers, [KO.success]);
  const defectKey = findColumn(sheet.headers, [KO.defect]);
  const defectRateKey = findColumn(sheet.headers, [KO.defectRate]);

  const rows = sheet.rows
    .map((row, index) => ({
      label: safeLabel(row, systemKey, partKey, index),
      total: toNumber(row[totalKey]),
      success: toNumber(row[successKey]),
      defect: toNumber(row[defectKey]),
      defectRate: asPercent(row[defectRateKey]),
    }))
    .filter((row) => row.total > 0);

  const totals = rows.reduce(
    (acc, row) => {
      acc.total += row.total;
      acc.success += row.success;
      acc.defect += row.defect;
      return acc;
    },
    { total: 0, success: 0, defect: 0 },
  );

  return {
    type: "defect",
    title: "Defect Overview",
    accent: "defect",
    metrics: [
      metric("Total Cases", totals.total),
      metric("Successful Runs", totals.success),
      metric("Defects", totals.defect),
      metric(
        "Defect Rate",
        totals.success > 0
          ? `${((totals.defect / totals.success) * 100).toFixed(1)}%`
          : "0.0%",
      ),
    ],
    chartTitle: "Defect Rate by Part",
    rows: rows
      .sort((a, b) => b.defect - a.defect || b.total - a.total)
      .slice(0, 10)
      .map((row) => ({
        label: row.label,
        primary: Math.max(0, Math.min(100, row.defectRate)),
        secondary:
          row.total > 0
            ? Math.max(0, Math.min(100, (row.defect / row.total) * 100))
            : 0,
        detail: `Defects ${row.defect}`,
      })),
    legend: [
      { label: "Defect Rate", tone: "danger" },
      { label: "Defects / Total", tone: "secondary" },
    ],
  };
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
