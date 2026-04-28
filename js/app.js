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
          ${columns.map((column) => `<td>${row[column] ?? ""}</td>`).join("")}
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

function renderDashboard(target, dashboard) {
  if (!target) {
    return [];
  }

  if (!dashboard) {
    target.hidden = true;
    target.innerHTML = "";
    return [];
  }

  const cards = dashboard.cards
    .map(
      (card) => `
        <article class="dash-card ${card.tone}">
          <span class="dash-card-label">${card.label}</span>
          <strong class="dash-card-value">${card.value}</strong>
        </article>
      `,
    )
    .join("");

  const charts = dashboard.charts
    .map(
      (chart) => `
        <section class="chart-card">
          <div class="chart-card-title">${chart.title}</div>
          <div class="chart-canvas-wrap">
            <canvas id="chart-${chart.key}"></canvas>
          </div>
        </section>
      `,
    )
    .join("");

  target.hidden = false;
  target.innerHTML = `
    <section class="dashboard-hero">
      <div class="dashboard-copy">
        <p class="dashboard-kicker">BI VISUALIZATION</p>
        <h2 class="dashboard-title">${dashboard.title}</h2>
        <p class="dashboard-subtitle">${dashboard.subtitle}</p>
      </div>
      <div class="dashboard-cards">${cards}</div>
    </section>
    <section class="dashboard-charts">${charts}</section>
  `;

  return dashboard.charts.map((chart) => ({
    ...chart,
    element: target.querySelector(`#chart-${chart.key}`),
  }));
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
  const dashboardContainer = resolveElement(
    config.dashboardContainer ?? "#sheet-dashboard",
  );
  const tableContainer = resolveElement(config.tableContainer ?? "#data-table");

  const chartInstances = [];

  function destroyCharts() {
    while (chartInstances.length > 0) {
      const chart = chartInstances.pop();
      chart.destroy();
    }
  }

  function mountCharts(charts) {
    destroyCharts();

    if (!window.Chart) {
      return;
    }

    charts.forEach((chart) => {
      if (!chart.element) {
        return;
      }

      const instance = new window.Chart(chart.element, {
        type: chart.type,
        data: chart.data,
        options: chart.options,
      });

      chartInstances.push(instance);
    });
  }

  function render(state) {
    setText(statusText, state.status, "idle");
    setText(errorText, state.error, "");
    setText(rowCountText, state.tableRows.length, "0");
    setText(fileNameText, state.fileName, "No file");
    renderSheetOptions(sheetSelect, state.sheets, state.selectedSheetName);
    const chartDefs = renderDashboard(dashboardContainer, state.dashboard);
    mountCharts(chartDefs);
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
      dashboardContainer,
    },
    render,
    bindEvents,
  };
}
