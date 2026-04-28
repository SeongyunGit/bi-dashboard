function normalizeCell(value) {
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  return value ?? "";
}

function nonEmptyCount(row) {
  return row.filter((cell) => String(cell ?? "").trim() !== "").length;
}

function detectHeaderRow(matrix) {
  const candidates = matrix.slice(0, 20);
  let bestIndex = 0;
  let bestScore = -1;

  candidates.forEach((row, index) => {
    const score = nonEmptyCount(row);
    if (score >= 2 && score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function buildRecords(matrix) {
  const cleanedMatrix = matrix.map((row) => row.map(normalizeCell));
  const headerRowIndex = detectHeaderRow(cleanedMatrix);
  const headers = cleanedMatrix[headerRowIndex].map((cell, index) => {
    const label = String(cell).trim();
    return label || `Column ${index + 1}`;
  });

  const rows = cleanedMatrix
    .slice(headerRowIndex + 1)
    .filter((row) => nonEmptyCount(row) > 0)
    .map((row) =>
      headers.reduce((record, header, index) => {
        record[header] = row[index] ?? "";
        return record;
      }, {}),
    );

  return {
    headerRowIndex,
    headers,
    rows,
  };
}

export async function parseWorkbook(file) {
  if (!window.XLSX) {
    throw new Error("Excel parser is not loaded");
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, {
    type: "array",
    cellDates: true,
  });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const matrix = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    const { headerRowIndex, headers, rows } = buildRecords(matrix);

    return {
      name: sheetName,
      headerRowIndex,
      headers,
      rows,
      rowCount: rows.length,
    };
  });

  return {
    fileName: file.name,
    sheetNames: workbook.SheetNames,
    sheets,
  };
}
