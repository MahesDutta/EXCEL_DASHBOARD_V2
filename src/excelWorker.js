import * as XLSX from "xlsx";

const DATE_WORDS = [
  "date", "day", "month", "year", "time", "timestamp", "period",
  "created", "updated", "invoice date", "order date", "transaction date",
  "delivery", "dispatch", "joining", "entry", "posting"
];

const MEASURE_WORDS = [
  "amount", "value", "sales", "sale", "revenue", "price", "cost", "profit",
  "income", "expense", "total", "qty", "quantity", "units", "volume", "balance",
  "stock", "rate", "count", "target", "budget", "margin"
];

const ID_WORDS = ["id", "code", "no", "number", "phone", "mobile", "zip", "pin"];

const ALIASES = {
  qty: "Quantity", "qty.": "Quantity", qnty: "Quantity", quantity: "Quantity",
  amt: "Amount", "amt.": "Amount", amount: "Amount", value: "Amount",
  sales: "Sales", sale: "Sales", revenue: "Revenue", rev: "Revenue",
  date: "Date", dt: "Date", "dt.": "Date", customer: "Customer", cust: "Customer",
  product: "Product", item: "Product", region: "Region", location: "Location",
  status: "Status", department: "Department"
};

function text(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : isoDate(value);
  if (typeof value === "object") {
    try { return JSON.stringify(value); } catch { return String(value); }
  }
  return String(value).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function isoDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = text(value);
  if (!raw) return null;
  const negative = /^\(.*\)$/.test(raw);
  const cleaned = raw
    .replace(/[₹$€£¥₽,%]/g, "")
    .replace(/\s/g, "")
    .replace(/[^\d.+-]/g, "");
  if (!cleaned || cleaned === "." || cleaned === "-" || cleaned === "+") return null;
  const number = Number(cleaned);
  return Number.isFinite(number) ? (negative ? -number : number) : null;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    const d = XLSX.SSF.parse_date_code(value);
    if (d?.y && d?.m && d?.d) return new Date(d.y, d.m - 1, d.d);
  }
  const raw = text(value);
  if (!raw) return null;

  // ISO / browser-readable dates first.
  const direct = new Date(raw);
  if (!Number.isNaN(direct.getTime()) && /\d{4}/.test(raw)) return direct;

  const parts = raw.split(/[./-]/).map(Number);
  if (parts.length !== 3 || !parts.every(Number.isFinite)) return null;
  let [a, b, c] = parts;
  if (c < 100) c += 2000;
  if (a >= 1000) return new Date(a, b - 1, c);
  if (c >= 1000) {
    // Prefer DD/MM/YYYY when the first part is > 12; otherwise MM/DD/YYYY.
    if (a > 12) return new Date(c, b - 1, a);
    if (b > 12) return new Date(c, a - 1, b);
    return new Date(c, a - 1, b);
  }
  return null;
}

function normalizeHeader(value, index) {
  const raw = text(value);
  if (!raw) return `Column ${index + 1}`;
  const key = raw.toLowerCase();
  if (ALIASES[key]) return ALIASES[key];
  return raw
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (m) => m.toUpperCase());
}

function uniqueName(name, used) {
  const base = text(name) || "Column";
  let candidate = base;
  let counter = 2;
  while (used.has(candidate.toLowerCase())) candidate = `${base} ${counter++}`;
  used.add(candidate.toLowerCase());
  return candidate;
}

function findHeaderRow(rawRows) {
  const limit = Math.min(rawRows.length, 30);
  let bestIndex = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < limit; i += 1) {
    const row = Array.isArray(rawRows[i]) ? rawRows[i] : [];
    const cells = row.map(text);
    const nonEmpty = cells.filter(Boolean);
    if (nonEmpty.length < 2) continue;
    const unique = new Set(nonEmpty.map((x) => x.toLowerCase())).size;
    let labelLike = 0;
    for (const value of nonEmpty) {
      if (/[a-zA-Z]/.test(value) && !/^\d+(?:[.,]\d+)?$/.test(value)) labelLike += 1;
    }
    const score = nonEmpty.length * 2 + unique + labelLike * 1.5 - i * 0.4;
    if (score > bestScore) { bestScore = score; bestIndex = i; }
  }
  return bestIndex;
}

function inferType(values, header) {
  const sample = [];
  for (const value of values) {
    if (text(value) !== "") sample.push(value);
    if (sample.length >= 300) break;
  }
  if (!sample.length) return "text";

  let dates = 0;
  let numbers = 0;
  for (const value of sample) {
    if (parseDate(value)) dates += 1;
    if (parseNumber(value) !== null) numbers += 1;
  }
  const dateScore = dates / sample.length;
  const numberScore = numbers / sample.length;
  const h = text(header).toLowerCase();
  const dateHint = DATE_WORDS.some((word) => h.includes(word));
  const idHint = ID_WORDS.some((word) => h === word || h.endsWith(` ${word}`));

  if (dateHint && dateScore >= 0.55) return "date";
  if (dateScore >= 0.90 && numberScore < 0.80) return "date";
  if (numberScore >= 0.90 && !idHint) return "number";
  return "text";
}

function detectColumns(rows, headers) {
  return headers.map((header, index) => {
    const values = new Array(rows.length);
    const uniqueSet = new Set();
    for (let i = 0; i < rows.length; i += 1) {
      const value = rows[i][index];
      values[i] = value;
      const normalized = text(value);
      if (normalized) uniqueSet.add(normalized);
    }
    const type = inferType(values, header);
    const measureHint = MEASURE_WORDS.some((word) => header.toLowerCase().includes(word));
    return {
      index,
      header,
      type,
      unique: uniqueSet.size,
      measureHint,
      isDate: type === "date",
      isNumeric: type === "number",
      isDimension: type === "text" && uniqueSet.size > 1
    };
  });
}

function safeValue(value, type) {
  if (value === null || value === undefined || value === "") return "";
  if (type === "number") {
    const number = parseNumber(value);
    return number === null ? text(value) : number;
  }
  if (type === "date") {
    const date = parseDate(value);
    return date ? isoDate(date) : text(value);
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return text(value);
}

function cleanSheet(sheet, name) {
  const raw = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: false
  });
  if (!raw.length) return null;

  const headerRowIndex = findHeaderRow(raw);
  const sourceHeaders = Array.isArray(raw[headerRowIndex]) ? raw[headerRowIndex] : [];
  let maxCols = sourceHeaders.length;
  for (const row of raw) if (Array.isArray(row) && row.length > maxCols) maxCols = row.length;
  if (maxCols === 0) return null;

  const used = new Set();
  const headers = new Array(maxCols);
  for (let i = 0; i < maxCols; i += 1) headers[i] = uniqueName(normalizeHeader(sourceHeaders[i], i), used);

  const matrix = [];
  for (let r = headerRowIndex + 1; r < raw.length; r += 1) {
    const source = Array.isArray(raw[r]) ? raw[r] : [];
    const row = new Array(maxCols);
    let hasValue = false;
    for (let c = 0; c < maxCols; c += 1) {
      const value = source[c] ?? "";
      row[c] = value;
      if (text(value) !== "") hasValue = true;
    }
    if (hasValue) matrix.push(row);
  }

  const columns = detectColumns(matrix, headers);
  const rows = new Array(matrix.length);
  for (let r = 0; r < matrix.length; r += 1) {
    const source = matrix[r];
    const output = { __row: r + 1 };
    for (const column of columns) output[column.header] = safeValue(source[column.index], column.type);
    rows[r] = output;
  }

  return {
    name,
    headerRow: headerRowIndex + 1,
    headers,
    columns,
    rows,
    sourceRows: raw.length
  };
}

function processWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellNF: false, cellStyles: false });
  const sheets = [];
  for (const name of workbook.SheetNames) {
    const cleaned = cleanSheet(workbook.Sheets[name], name);
    if (cleaned?.rows?.length) sheets.push(cleaned);
  }
  if (!sheets.length) throw new Error("No usable table data was found in this workbook.");
  return { sheetNames: sheets.map((sheet) => sheet.name), sheets };
}

self.onmessage = async (event) => {
  const { id, buffer } = event.data || {};
  if (!id) return;
  try {
    const result = processWorkbook(buffer);
    self.postMessage({ id, ok: true, result });
  } catch (error) {
    self.postMessage({ id, ok: false, error: error?.message || "The workbook could not be processed." });
  }
};
