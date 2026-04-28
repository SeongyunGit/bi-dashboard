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

function cleanDisplayLabel(label) {
  return String(label)
    .replace(/\([^)]*\)/g, "")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .replace(/시스템 구축/g, "시스템")
    .replace(/시스템 업무개발/g, "업무개발")
    .trim();
}

function toChartLabel(label) {
  const cleaned = cleanDisplayLabel(label);
  const words = cleaned.split(" ");
  const lines = [];
  let current = "";

  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length > 12) {
      if (current) {
        lines.push(current);
      }
      current = word;
    } else {
      current = next;
    }
  });

  if (current) {
    lines.push(current);
  }

  return lines.length > 0 ? lines : [cleaned];
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

      const label = cleanDisplayLabel(sheet.name);

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
    subtitle: "계획, 실제 진척, 완료 수량, 격차를 한 화면에서 봅니다",
    cards: [
      createCard("총 케이스", total, "primary"),
      createCard("계획 수량", planned, "secondary"),
      createCard("완료 수량", completed, "accent"),
      createCard("잔여 수량", remaining, "muted"),
    ],
    charts: [
      {
        key: "progress-volume",
        title: "계획 수량 대비 완료 수량",
        type: "bar",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createBarDataset("계획", rows.map((row) => row.planned), "#60a5fa"),
            createBarDataset(
              "완료",
              rows.map((row) => row.completed),
              "#22c55e",
            ),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "progress-rate",
        title: "계획 진척률 대비 실제 진척률",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => cleanDisplayLabel(row.label)),
          datasets: [
            {
              label: "계획 진척률",
              data: rows.slice(0, 8).map((row) => clampPercent(row.planRate)),
              borderColor: "#60a5fa",
              backgroundColor: "rgba(96,165,250,0.18)",
            },
            {
              label: "실제 진척률",
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
        title: "완료 대비 잔여 비중",
        type: "doughnut",
        data: {
          labels: ["완료", "잔여"],
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
        title: "진척 격차 추이",
        type: "line",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createLineDataset(
              "격차",
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
    subtitle: "수행, 실패, 미수행, 수행률을 동시에 확인합니다",
    cards: [
      createCard("총 케이스", total, "primary"),
      createCard("수행 수", executed, "secondary"),
      createCard("실패 수", fail, "danger"),
      createCard("미수행 수", pending, "muted"),
    ],
    charts: [
      {
        key: "execution-stack",
        title: "파트별 수행 분포",
        type: "bar",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createBarDataset(
              "수행",
              rows.map((row) => row.executed),
              "#60a5fa",
            ),
            createBarDataset("실패", rows.map((row) => row.fail), "#f87171"),
            createBarDataset("미수행", rows.map((row) => row.pending), "#475569"),
          ],
        },
        options: stackedBarOptions(),
      },
      {
        key: "execution-donut",
        title: "전체 수행 비중",
        type: "doughnut",
        data: {
          labels: ["수행", "실패", "미수행"],
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
        title: "실패 추이",
        type: "line",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createLineDataset("실패", rows.map((row) => row.fail), "#f87171"),
          ],
        },
        options: lineOptions(""),
      },
      {
        key: "execution-rate",
        title: "수행률",
        type: "bar",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createBarDataset(
              "수행률",
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
    subtitle: "품질 성공, 실패, 완료율, 품질률을 함께 보여줍니다",
    cards: [
      createCard("총 케이스", total, "primary"),
      createCard("완료 수량", completed, "secondary"),
      createCard("품질 성공", success, "accent"),
      createCard("품질 실패", fail, "danger"),
    ],
    charts: [
      {
        key: "quality-bar",
        title: "품질 성공 대비 실패",
        type: "bar",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createBarDataset("성공", rows.map((row) => row.success), "#22c55e"),
            createBarDataset("실패", rows.map((row) => row.fail), "#f87171"),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "quality-donut",
        title: "품질 결과 비중",
        type: "doughnut",
        data: {
          labels: ["성공", "실패"],
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
        title: "품질률 대비 완료율",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => cleanDisplayLabel(row.label)),
          datasets: [
            {
              label: "품질률",
              data: rows.slice(0, 8).map((row) => clampPercent(row.successRate)),
              borderColor: "#f59e0b",
              backgroundColor: "rgba(245,158,11,0.18)",
            },
            {
              label: "완료율",
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
        title: "품질률 추이",
        type: "line",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createLineDataset(
              "품질률",
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
    subtitle: "결함 수, 결함률, 파트별 결함 집중도를 봅니다",
    cards: [
      createCard("총 케이스", total, "primary"),
      createCard("성공 수", success, "secondary"),
      createCard("결함 수", defect, "danger"),
      createCard(
        "결함률",
        success > 0 ? formatPercent((defect / success) * 100) : "0.0%",
        "accent",
      ),
    ],
    charts: [
      {
        key: "defect-bar",
        title: "파트별 결함 수",
        type: "bar",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createBarDataset("결함 수", rows.map((row) => row.defect), "#f87171"),
          ],
        },
        options: groupedBarOptions(),
      },
      {
        key: "defect-line",
        title: "결함률 추이",
        type: "line",
        data: {
          labels: rows.map((row) => toChartLabel(row.label)),
          datasets: [
            createLineDataset(
              "결함률",
              rows.map((row) => clampPercent(row.defectRate)),
              "#fb7185",
            ),
          ],
        },
        options: lineOptions("%"),
      },
      {
        key: "defect-donut",
        title: "성공 대비 결함 비중",
        type: "doughnut",
        data: {
          labels: ["성공", "결함"],
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
        title: "결함률 대비 결함 부하",
        type: "radar",
        data: {
          labels: rows.slice(0, 8).map((row) => cleanDisplayLabel(row.label)),
          datasets: [
            {
              label: "결함률",
              data: rows.slice(0, 8).map((row) => clampPercent(row.defectRate)),
              borderColor: "#fb7185",
              backgroundColor: "rgba(251,113,133,0.18)",
            },
            {
              label: "결함 부하",
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
