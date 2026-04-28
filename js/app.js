import { parseWorkbook } from "./parser.js";
import { createState } from "./state.js";
import { createUI } from "./ui.js";

function selectSheetData(sheets, sheetName) {
  const selectedSheet =
    sheets.find((sheet) => sheet.name === sheetName) ?? sheets[0] ?? null;

  return {
    selectedSheetName: selectedSheet?.name ?? "",
    tableColumns: selectedSheet?.headers ?? [],
    tableRows: selectedSheet?.rows ?? [],
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
