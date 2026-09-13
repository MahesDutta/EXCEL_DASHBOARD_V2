import * as XLSX from "xlsx";

const MAX_ROWS = 50000;
const SAMPLE_SIZE = 500;

function text(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  try { return JSON.stringify(value); } catch { return String(value); }
}

function isoDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseDate(value) {
  if (!value) return null;
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [y, m, d] = value.split("-").map(Number);
    const date = new Date(y, m - 1, d);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  if (typeof value === "number" && value > 0 && value < 100000) {
    const date = new Date((value - 25569) * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeHeader(value, index) {
  if (typeof value === "string") {
    let normalized = value.trim();
    if (!normalized) normalized = `Column ${index + 1}`;
    return normalized;
  }
  return `Column ${index + 1}`;
}

function uniqueName(name, used) {
  if (!used.has(name)) {
    used.add(name);
    return name;
  }
  for (let i = 1; i <= 1000; i++) {
    const candidate = `${name} ${i}`;
    if (!used.has(candidate)) {
      used.add(candidate);
      return candidate;
    }
  }
  return name;
}

function findHeaderRow(rawRows) {
  if (rawRows.length === 0) return 0;
  for (let i = 0; i < Math.min(10, rawRows.length); i++) {
    const row = rawRows[i];
    if (!Array.isArray(row)) continue;
    let headerCount = 0;
    for (const cell of row) {
      if (typeof cell === "string" && cell.trim().length > 0) headerCount++;
    }
    if (headerCount >= row.length * 0.5) return i;
  }
  return 0;
}

function inferType(values, header) {
  let dateCount = 0;
  let numberCount = 0;
  let validCount = 0;

  for (const val of values) {
    if (val === "" || val === null || val === undefined) continue;
    validCount++;

    const num = parseNumber(val);
    if (num !== null) numberCount++;

    const date = parseDate(val);
    if (date) dateCount++;
  }

  if (validCount === 0) return { isNumeric: false, isDate: false, isDimension: false };

  const dateRatio = dateCount / validCount;
  const numberRatio = numberCount / validCount;

  const isDate = dateRatio > 0.5;
  const isNumeric = numberRatio > 0.5 && !isDate;
  const isDimension = !isNumeric && !isDate;
  const measureHint = isNumeric && (header.toLowerCase().includes("total") || header.toLowerCase().includes("amount") || header.toLowerCase().includes("value") || header.toLowerCase().includes("count"));

  return { isNumeric, isDate, isDimension, measureHint };
}

function detectColumns(rows, headers) {
  const columns = [];
  const sampleRows = rows.slice(0, Math.min(SAMPLE_SIZE, rows.length));

  for (let i = 0; i < headers.length; i++) {
    const values = sampleRows.map((row) => row[headers[i]]).filter((v) => v !== null && v !== undefined);
    const type = inferType(values, headers[i]);
    
    const uniqueValues = new Set(values.map(String));
    columns.push({
      header: headers[i],
      index: i,
      isNumeric: type.isNumeric,
      isDate: type.isDate,
      isDimension: type.isDimension,
      unique: uniqueValues.size,
      measureHint: type.measureHint
    });
  }
  return columns;
}

function safeValue(value, type) {
  if (value === null || value === undefined || value === "") return "";
  if (type.isNumeric) {
    const n = parseNumber(value);
    return n !== null ? n : "";
  }
  if (type.isDate) {
    const d = parseDate(value);
    if (d) return isoDate(d);
    return "";
  }
  return text(value);
}

function cleanSheet(sheet, name) {
  const rawRows = sheet.get_array ? sheet.get_array() : XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  
  if (rawRows.length === 0) return null;

  const headerRowIndex = findHeaderRow(rawRows);
  const headerRow = Array.isArray(rawRows[headerRowIndex]) ? rawRows[headerRowIndex] : [];
  
  const usedNames = new Set();
  const headers = headerRow.map((h, i) => uniqueName(normalizeHeader(h, i), usedNames));
  
  const rows = [];
  for (let i = headerRowIndex + 1; i < Math.min(headerRowIndex + 1 + MAX_ROWS, rawRows.length); i++) {
    const rawRow = rawRows[i];
    if (!Array.isArray(rawRow)) continue;
    
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = rawRow[j] ?? "";
    }
    rows.push(row);
  }

  const columns = detectColumns(rows, headers);

  const cleanedRows = rows.map((row) => {
    const cleaned = {};
    for (const column of columns) {
      cleaned[column.header] = safeValue(row[column.header], column);
    }
    return cleaned;
  });

  return { name, headers, rows: cleanedRows, columns, headerRow: headerRowIndex + 1 };
}

function processWorkbook(buffer) {
  try {
    const workbook = XLSX.read(buffer, { cellDates: true, cellFormulas: false });
    const sheets = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const cleaned = cleanSheet(sheet, sheetName);
      if (cleaned && cleaned.rows.length > 0) {
        sheets.push(cleaned);
      }
    }

    if (sheets.length === 0) {
      return { ok: false, error: "No data found in workbook." };
    }

    return { ok: true, result: { sheets } };
  } catch (error) {
    console.error("Workbook processing error:", error);
    return { ok: false, error: error.message || "Failed to process workbook." };
  }
}

self.onmessage = async (event) => {
  const { id, buffer } = event.data;
  try {
    const result = processWorkbook(buffer);
    self.postMessage({ id, ...result });
  } catch (error) {
    console.error("Worker error:", error);
    self.postMessage({ id, ok: false, error: "Excel processing failed." });
  }
};
