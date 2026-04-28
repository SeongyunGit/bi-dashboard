function resolveElement(target) {
  if (!target) {
    return null;
  }

  if (typeof target === "string") {
    return document.querySelector(target);
  }

  return target;
}

function setText(target, value, fallback = "-") {
  if (!target) {
    return;
  }

  target.textContent = value ?? fallback;
}

function renderRows(target, columns, rows) {
  if (!target) {
    return;
  }

  if (!Array.isArray(columns) || columns.length === 0) {
    target.innerHTML = "<tr><td colspan=\"100%\">No data</td></tr>";
    return;
  }

  const head = `
    <thead>
      <tr>${columns.map((column) => `<th>${column}</th>`).join("")}</tr>
    </thead>
  `;

  const body = rows
    .map(
      (row) => `
        <tr>
          ${columns
            .map((column) => `<td>${row[column] ?? ""}</td>`)
            .join("")}
        </tr>
      `,
    )
    .join("");

  target.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderSheetOptions(target, sheets, selectedSheetName) {
  if (!target) {
    return;
  }

  if (!Array.isArray(sheets) || sheets.length === 0) {
    target.innerHTML = "<option value=\"\">Select a sheet</option>";
    return;
  }

  target.innerHTML = sheets
    .map(
      (sheet) => `
        <option value="${sheet.name}" ${
          sheet.name === selectedSheetName ? "selected" : ""
        }>
          ${sheet.name} (${sheet.rowCount})
        </option>
      `,
    )
    .join("");
}

export function createUI(config = {}) {
  const fileInput = resolveElement(config.fileInput ?? "#file-input");
  const parseButton = resolveElement(config.parseButton ?? "#parse-button");
  const downloadButton = resolveElement(
    config.downloadButton ?? "#download-button",
  );
  const resetButton = resolveElement(config.resetButton ?? "#reset-button");
  const statusText = resolveElement(config.statusText ?? "#status");
  const errorText = resolveElement(config.errorText ?? "#error");
  const rowCountText = resolveElement(config.rowCountText ?? "#row-count");
  const fileNameText = resolveElement(config.fileNameText ?? "#file-name");
  const sheetSelect = resolveElement(config.sheetSelect ?? "#sheet-select");
  const tableContainer = resolveElement(config.tableContainer ?? "#data-table");

  function render(state) {
    setText(statusText, state.status, "idle");
    setText(errorText, state.error, "");
    setText(rowCountText, state.tableRows.length, "0");
    setText(fileNameText, state.fileName, "No file");
    renderSheetOptions(sheetSelect, state.sheets, state.selectedSheetName);
    renderRows(tableContainer, state.tableColumns, state.tableRows);

    if (downloadButton) {
      downloadButton.disabled = state.tableRows.length === 0;
    }
  }

  function bindEvents(handlers) {
    if (fileInput && handlers.onFileChange) {
      fileInput.addEventListener("change", handlers.onFileChange);
    }

    if (parseButton && handlers.onParse) {
      parseButton.addEventListener("click", handlers.onParse);
    }

    if (sheetSelect && handlers.onSheetSelect) {
      sheetSelect.addEventListener("change", handlers.onSheetSelect);
    }

    if (downloadButton && handlers.onDownload) {
      downloadButton.addEventListener("click", handlers.onDownload);
    }

    if (resetButton && handlers.onReset) {
      resetButton.addEventListener("click", handlers.onReset);
    }
  }

  return {
    elements: {
      fileInput,
      parseButton,
      downloadButton,
      resetButton,
      sheetSelect,
    },
    render,
    bindEvents,
  };
}
