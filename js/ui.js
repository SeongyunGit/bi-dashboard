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

function renderVisualization(target, visualization) {
  if (!target) {
    return;
  }

  if (!visualization) {
    target.hidden = true;
    target.innerHTML = "";
    return;
  }

  const metrics = visualization.metrics
    .map(
      (metric) => `
        <article class="viz-metric">
          <span class="viz-metric-label">${metric.label}</span>
          <strong class="viz-metric-value">${metric.value}</strong>
        </article>
      `,
    )
    .join("");

  const legend = visualization.legend
    .map(
      (item) => `
        <span class="viz-legend-item">
          <i class="viz-legend-dot ${item.tone}"></i>
          ${item.label}
        </span>
      `,
    )
    .join("");

  const rows = visualization.rows
    .map((row) => {
      if (row.segments) {
        const segments = row.segments
          .map((segment) => {
            const width =
              row.total > 0 ? Math.max(0, (segment.value / row.total) * 100) : 0;
            return `<span class="viz-segment ${segment.tone}" style="width:${width}%"></span>`;
          })
          .join("");

        return `
          <div class="viz-chart-row">
            <div class="viz-chart-head">
              <span class="viz-chart-label">${row.label}</span>
              <span class="viz-chart-detail">${row.total}</span>
            </div>
            <div class="viz-stack">${segments}</div>
          </div>
        `;
      }

      return `
        <div class="viz-chart-row">
          <div class="viz-chart-head">
            <span class="viz-chart-label">${row.label}</span>
            <span class="viz-chart-detail">${row.detail}</span>
          </div>
          <div class="viz-bars">
            <div class="viz-bar-track">
              <span class="viz-bar primary" style="width:${row.primary}%"></span>
            </div>
            <div class="viz-bar-track thin">
              <span class="viz-bar secondary" style="width:${row.secondary}%"></span>
            </div>
          </div>
        </div>
      `;
    })
    .join("");

  target.hidden = false;
  target.innerHTML = `
    <section class="viz-surface ${visualization.accent}">
      <div class="viz-header">
        <div>
          <p class="viz-eyebrow">${visualization.title}</p>
          <h2 class="viz-title">${visualization.chartTitle}</h2>
        </div>
        <div class="viz-legend">${legend}</div>
      </div>
      <div class="viz-metrics">${metrics}</div>
      <div class="viz-chart">${rows}</div>
    </section>
  `;
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
  const visualizationContainer = resolveElement(
    config.visualizationContainer ?? "#sheet-visualization",
  );
  const tableContainer = resolveElement(config.tableContainer ?? "#data-table");

  function render(state) {
    setText(statusText, state.status, "idle");
    setText(errorText, state.error, "");
    setText(rowCountText, state.tableRows.length, "0");
    setText(fileNameText, state.fileName, "No file");
    renderSheetOptions(sheetSelect, state.sheets, state.selectedSheetName);
    renderVisualization(visualizationContainer, state.visualization);
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
      visualizationContainer,
    },
    render,
    bindEvents,
  };
}
