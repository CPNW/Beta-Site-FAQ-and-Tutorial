// Browser-only CSV Doctor (CPNW Passport cleaner) using ExcelJS.
// Processes the uploaded workbook entirely in the browser and returns a cleaned Excel file.

const MODULE_BASES = [
  "Bloodborne Pathogens and Workplace Safety",
  "Chemical Hazard Communication",
  "Compliance",
  "Emergency Procedures",
  "Magnetic Resonance Imaging Safety",
  "Patient Rights",
  "Patient Safety",
  "Infectious Medical Waste",
  "Fall Risk Prevention",
  "Infection Prevention and Standard Precautions",
];
const MODULE_BASES_SET = new Set(MODULE_BASES);
const BASE_FIELDS = ["Name", "Email", "Program"];
const LINE_BREAK = "\n";

const form = document.getElementById("clean-form");
const fileInput = document.getElementById("file-input");
const sheetInput = document.getElementById("sheet-name");
const outputInput = document.getElementById("output-name");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("submit-btn");

form?.addEventListener("submit", async (e) => {
  e.preventDefault();
  clearStatus();

  const file = fileInput.files[0];
  if (!file) {
    setStatus("Please choose an Excel file (.xlsx).", true);
    return;
  }

  const sheetName = sheetInput.value || "Sheet 1";
  let outputName = (outputInput.value || "CPNW_CleanExport.xlsx").trim();
  if (!outputName.toLowerCase().endsWith(".xlsx")) {
    outputName += ".xlsx";
  }

  setBusy(true);
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in the workbook.`);
    }

    const rows = sheetToObjects(sheet);
    if (!rows.length) {
      throw new Error("The sheet is empty or could not be read.");
    }

    const cleanedRows = rows.map((row) => cleanRow(row));
    const outWb = new ExcelJS.Workbook();
    const outSheet = outWb.addWorksheet("Cleaned");
    buildTable(outSheet, cleanedRows);
    applySheetStyling(outSheet);

    const blob = await outWb.xlsx.writeBuffer();
    triggerDownload(blob, outputName);
    setStatus("Done. Downloaded cleaned file. Reloading…", false, true);
    setTimeout(() => location.reload(), 300);
  } catch (err) {
    console.error(err);
    setStatus(err.message || "Failed to clean file.", true);
  } finally {
    setBusy(false);
  }
});

function sheetToObjects(sheet) {
  const header = [];
  sheet.getRow(1).eachCell({ includeEmpty: true }, (cell) => {
    header.push((cell.value ?? "").toString());
  });

  const rows = [];
  sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const key = header[colNumber - 1] || `col${colNumber}`;
      obj[key] = normalizeCellValue(cell);
    });
    if (Object.values(obj).every((v) => v === "")) return; // skip fully empty
    rows.push(obj);
  });
  return rows;
}

function normalizeCellValue(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (v.text) return v.text.trim();
  if (v.result) return v.result;
  return v;
}

function buildTable(sheet, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const dataRows = rows.map((row) => headers.map((h) => row[h] ?? ""));

  sheet.addTable({
    name: "CleanedTable",
    ref: "A1",
    headerRow: true,
    totalsRow: false,
    style: {
      theme: "TableStyleMedium2",
      showRowStripes: true,
      showColumnStripes: false,
    },
    columns: headers.map((h) => ({ name: h, filterButton: true })),
    rows: dataRows,
  });
}

function cleanRow(row) {
  const groupedColumns = {};
  const groupedLabels = {};
  const order = [];

  Object.keys(row).forEach((col) => {
    if (BASE_FIELDS.includes(col)) return;
    const { base, detail } = splitColumnName(col);
    if (!groupedColumns[base]) {
      groupedColumns[base] = [];
      groupedLabels[base] = {};
      order.push(base);
    }
    groupedColumns[base].push(col);
    if (detail) groupedLabels[base][col] = detail;
  });

  const out = {
    Name: row["Name"] ?? "",
    Email: row["Email"] ?? "",
    Program: row["Program"] ?? "",
  };

  order.forEach((base) => {
    if (MODULE_BASES_SET.has(base)) return; // handled later
    const destCol = cleanBaseName(base);
    out[destCol] = summarizeRow(row, groupedColumns[base], groupedLabels[base]);
  });

  out["eLearning Modules"] = buildModuleCell(row, groupedColumns, groupedLabels);
  return out;
}

function buildModuleCell(row, groupedColumns, groupedLabels) {
  const lines = [];
  MODULE_BASES.forEach((moduleBase) => {
    if (!groupedColumns[moduleBase]) {
      lines.push(cleanBaseName(moduleBase));
      return;
    }
    const cols = groupedColumns[moduleBase];
    const labels = groupedLabels[moduleBase];
    const parts = [];
    cols.forEach((col) => {
      const val = formatValue(row[col]);
      if (!val) return;
      const label = labels[col] || col;
      parts.push(`${label}: ${val}`);
    });
    let line = cleanBaseName(moduleBase);
    if (parts.length) line += " - " + parts.join(" - ");
    lines.push(line);
  });
  return lines.join(LINE_BREAK);
}

function summarizeRow(row, columns, labels) {
  const pieces = [];
  columns.forEach((col) => {
    const val = formatValue(row[col]);
    if (!val) return;
    const label = labels[col] || col;
    pieces.push(`${label}: ${val}`);
  });
  return pieces.join(LINE_BREAK);
}

function splitColumnName(columnName) {
  if (columnName.startsWith("CPNW: ")) {
    const [base, ...rest] = columnName.split(" - ");
    return { base: base.trim(), detail: rest.join(" - ").trim() };
  }
  const parts = columnName.split(" - ");
  if (parts.length >= 3 && (parts[0] === "ND" || parts[0] === "WA")) {
    const base = parts.slice(0, 2).join(" - ");
    const detail = parts.slice(2).join(" - ");
    return { base: base.trim(), detail: detail.trim() };
  }
  if (parts.length >= 2) {
    return { base: parts[0].trim(), detail: parts.slice(1).join(" - ").trim() };
  }
  return { base: columnName.trim(), detail: "" };
}

function cleanBaseName(baseName) {
  if (baseName.startsWith("CPNW: ")) {
    return baseName.replace("CPNW: ", "").trim();
  }
  return baseName.replace(/\s+/g, " ").trim();
}

function formatValue(value) {
  if (value === null || value === undefined) return "";

  if (typeof value === "number") {
    const serialDate = parseExcelSerialDate(value);
    if (serialDate) return formatDate(serialDate);
    return String(value).trim();
  }

  if (value instanceof Date && !isNaN(value.valueOf())) {
    return formatDate(value);
  }

  const text = String(value).trim();
  const maybeDate = parseDate(text);
  if (maybeDate) return formatDate(maybeDate);
  return text;
}

function parseExcelSerialDate(num) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  if (num > 59) {
    const ms = (num - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.valueOf()) && d.getFullYear() > 1900) return d;
  }
  return null;
}

function parseDate(text) {
  const parsed = new Date(text);
  if (!isNaN(parsed.valueOf())) return parsed;
  return null;
}

function formatDate(dateObj) {
  const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
  const dd = String(dateObj.getDate()).padStart(2, "0");
  const yyyy = dateObj.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

function triggerDownload(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function setStatus(msg, isError = false, isSuccess = false) {
  statusEl.textContent = msg;
  statusEl.className = [
    "small",
    isError ? "text-danger" : "",
    isSuccess ? "text-success" : "",
    !isError && !isSuccess && msg ? "text-muted" : "",
  ]
    .filter(Boolean)
    .join(" ");
}

function clearStatus() {
  setStatus("");
}

function setBusy(busy) {
  submitBtn.disabled = busy;
  submitBtn.textContent = busy ? "Processing..." : "Clean Export";
}

function applySheetStyling(sheet) {
  // Wrap text for all cells; table styling applies stripes, so just borders and widths.
  const totalRows = sheet.rowCount;
  const totalCols = sheet.columnCount;

  for (let r = 1; r <= totalRows; r++) {
    const row = sheet.getRow(r);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = Object.assign({}, cell.alignment, {
        wrapText: true,
        vertical: "top",
      });
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
  }

  // Approximate column widths based on content length.
  for (let c = 1; c <= totalCols; c++) {
    let maxLen = 10;
    for (let r = 1; r <= totalRows; r++) {
      const cell = sheet.getRow(r).getCell(c);
      const val = cell.value;
      if (!val) continue;
      const lines = String(val).split(/\r?\n/);
      lines.forEach((l) => {
        if (l.length > maxLen) maxLen = l.length;
      });
    }
    sheet.getColumn(c).width = Math.min(Math.max(maxLen + 2, 12), 60);
  }
}
