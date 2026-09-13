import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  CartesianGrid,
  Cell,
  PieChart,
  Pie,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  ResponsiveContainer,
  ReferenceLine,
} from "recharts";
import {
  Upload,
  FileSpreadsheet,
  Sun,
  Moon,
  Download,
  Printer,
  RefreshCcw,
  Search,
  SlidersHorizontal,
  Sparkles,
  Database,
  CalendarDays,
  BarChart3,
  Filter,
  X,
  ChevronDown,
  Check,
  AlertTriangle,
  FileDown,
  Eye,
  Columns3,
  WandSparkles,
} from "lucide-react";
import "./styles.css";

const DATE_WORDS = [
  "date", "day", "month", "year", "time", "timestamp", "period",
  "created", "updated", "invoice", "order", "transaction", "delivery",
  "dispatch", "joining", "entry", "posting"
];
const MEASURE_WORDS = [
  "amount", "value", "sales", "sale", "revenue", "price", "cost",
  "profit", "income", "expense", "total", "qty", "quantity", "units",
  "volume", "balance", "stock", "rate", "count", "target", "budget"
];
const ID_WORDS = ["id", "code", "no", "number", "invoice", "phone", "mobile", "zip", "pin"];

const cleanText = (value) =>
  String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const titleCase = (value) =>
  cleanText(value)
    .replace(/[_-]+/g, " ")
    .replace(/\b\w/g, (m) => m.toUpperCase());

function uniqueName(name, used) {
  const base = cleanText(name) || "Column";
  let result = base;
  let i = 2;
  while (used.has(result.toLowerCase())) {
    result = `${base} ${i++}`;
  }
  used.add(result.toLowerCase());
  return result;
}

function normalizeHeader(value, index) {
  const raw = cleanText(value);
  if (!raw) return `Column ${index + 1}`;
  const aliases = {
    qty: "Quantity",
    "qty.": "Quantity",
    quantity: "Quantity",
    qnty: "Quantity",
    amt: "Amount",
    "amt.": "Amount",
    amount: "Amount",
    value: "Amount",
    sales: "Sales",
    sale: "Sales",
    revenue: "Revenue",
    rev: "Revenue",
    date: "Date",
    dt: "Date",
    "dt.": "Date",
    customer: "Customer",
    cust: "Customer",
    product: "Product",
    item: "Product",
    region: "Region",
    location: "Location",
    status: "Status",
    department: "Department",
  };
  const key = raw.toLowerCase();
  return aliases[key] || titleCase(raw);
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value === null || value === undefined || value === "") return null;
  let text = cleanText(value);
  if (!text) return null;

  const negative = /^\(.*\)$/.test(text);
  text = text
    .replace(/[₹$€£¥₽,%]/g, "")
    .replace(/\s/g, "")
    .replace(/[^\d.+-]/g, "");

  if (!text || text === "-" || text === "." || text === "+") return null;
  const number = Number(text);
  if (!Number.isFinite(number)) return null;
  return negative ? -number : number;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
  }

  const text = cleanText(value);
  if (!text) return null;

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) return direct;

  const parts = text.split(/[./-]/).map((x) => Number(x));
  if (parts.length === 3 && parts.every(Number.isFinite)) {
    let [a, b, c] = parts;
    if (c < 100) c += 2000;
    if (a > 1000) return new Date(a, b - 1, c);
    if (c > 1000) {
      if (a > 12) return new Date(c, b - 1, a);
      return new Date(c, a - 1, b);
    }
  }
  return null;
}

function dateKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatFullNumber(value) {
  if (value === null || value === undefined || !Number.isFinite(Number(value))) return "—";
  return new Intl.NumberFormat("en-IN", {
    maximumFractionDigits: 2,
  }).format(Number(value));
}

function formatCompact(value) {
  if (!Number.isFinite(Number(value))) return "—";
  const n = Number(value);
  const abs = Math.abs(n);
  if (abs >= 1e9) return `${(n / 1e9).toFixed(1)}B`;
  if (abs >= 1e6) return `${(n / 1e6).toFixed(1)}M`;
  if (abs >= 1e3) return `${(n / 1e3).toFixed(1)}K`;
  return formatFullNumber(n);
}

function formatDate(date) {
  return date
    ? new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short", year: "numeric" }).format(date)
    : "—";
}

function formatDateInput(date) {
  return dateKey(date);
}

function niceNumber(value) {
  return Math.round(Number(value) * 100) / 100;
}

function inferColumnType(values, header) {
  const nonEmpty = values.filter((v) => cleanText(v) !== "");
  if (!nonEmpty.length) return "text";

  const dateScore = nonEmpty.slice(0, 250).reduce((sum, v) => sum + (parseDate(v) ? 1 : 0), 0) / Math.min(nonEmpty.length, 250);
  const numberScore = nonEmpty.slice(0, 250).reduce((sum, v) => sum + (parseNumber(v) !== null ? 1 : 0), 0) / Math.min(nonEmpty.length, 250);

  const h = cleanText(header).toLowerCase();
  if (DATE_WORDS.some((w) => h.includes(w)) && dateScore >= 0.55) return "date";
  if (dateScore >= 0.88 && numberScore < 0.8) return "date";
  if (numberScore >= 0.9) {
    if (ID_WORDS.some((w) => h === w || h.endsWith(` ${w}`))) return "text";
    return "number";
  }
  return "text";
}

function detectColumns(rows, headers) {
  return headers.map((header, index) => {
    const values = rows.map((r) => r[index]);
    const type = inferColumnType(values, header);
    const unique = new Set(values.map(cleanText).filter(Boolean)).size;
    const measureHint = MEASURE_WORDS.some((w) => header.toLowerCase().includes(w));
    return {
      index,
      header,
      type,
      unique,
      measureHint,
      isDate: type === "date",
      isNumeric: type === "number",
      isDimension: type === "text" && unique > 1,
    };
  });
}

function findHeaderRow(rawRows) {
  const scan = rawRows.slice(0, Math.min(rawRows.length, 20));
  let best = { index: 0, score: -Infinity };
  scan.forEach((row, index) => {
    const cells = row.map(cleanText);
    const nonEmpty = cells.filter(Boolean);
    if (nonEmpty.length < 2) return;
    const unique = new Set(nonEmpty.map((x) => x.toLowerCase())).size;
    const likelyHeaders = nonEmpty.filter((x) =>
      /[a-zA-Z]/.test(x) &&
      !/^\d+([.,]\d+)?$/.test(x)
    ).length;
    const score = nonEmpty.length * 2 + unique + likelyHeaders * 1.5 - index * 0.4;
    if (score > best.score) best = { index, score };
  });
  return best.index;
}

function cleanSheet(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: false });
  if (!raw.length) return null;

  const headerRow = findHeaderRow(raw);
  const sourceHeaders = raw[headerRow] || [];
  let maxCols = sourceHeaders.length;
  for (let i = 0; i < raw.length; i += 1) {
    if (raw[i].length > maxCols) maxCols = raw[i].length;
  }
  const used = new Set();
  const headers = Array.from({ length: maxCols }, (_, i) => uniqueName(normalizeHeader(sourceHeaders[i], i), used));

  const rows = raw
    .slice(headerRow + 1)
    .map((row) => headers.map((_, i) => row[i] ?? ""))
    .filter((row) => row.some((v) => cleanText(v) !== ""));

  const columns = detectColumns(rows, headers);

  const objects = rows.map((row, rowIndex) => {
    const obj = { __row: rowIndex + 1 };
    columns.forEach((c) => {
      const rawValue = row[c.index];
      obj[c.header] =
        c.type === "number" ? parseNumber(rawValue) :
        c.type === "date" ? parseDate(rawValue) :
        cleanText(rawValue);
    });
    return obj;
  });

  return {
    name: sheet,
    headerRow: headerRow + 1,
    headers,
    columns,
    rows: objects,
    sourceRows: raw.length,
  };
}

function choosePrimaryMetric(columns) {
  const numeric = columns.filter((c) => c.isNumeric);
  if (!numeric.length) return null;
  const hinted = numeric.find((c) => c.measureHint);
  return hinted || numeric.sort((a, b) => b.unique - a.unique)[0];
}

function chooseSecondaryMetric(columns, primary) {
  return columns.find((c) => c.isNumeric && c.header !== primary?.header) || null;
}

function chooseDimension(columns) {
  const dimensions = columns.filter((c) => c.isDimension);
  if (!dimensions.length) return null;
  return dimensions
    .sort((a, b) => {
      const aScore = Math.abs(Math.log10(Math.max(a.unique, 1)) - 1.5);
      const bScore = Math.abs(Math.log10(Math.max(b.unique, 1)) - 1.5);
      return aScore - bScore;
    })[0];
}

function chooseDate(columns) {
  return columns.find((c) => c.isDate) || null;
}

function aggregateByCategory(rows, dimension, metric) {
  if (!dimension || !metric) return [];
  const map = new Map();
  rows.forEach((row) => {
    const key = cleanText(row[dimension.header]) || "Blank";
    const value = Number(row[metric.header]);
    if (!Number.isFinite(value)) return;
    map.set(key, (map.get(key) || 0) + value);
  });
  return [...map.entries()]
    .map(([name, value]) => ({ name, value: niceNumber(value) }))
    .sort((a, b) => b.value - a.value);
}

function aggregateTrend(rows, dateColumn, metric) {
  if (!dateColumn || !metric) return [];
  const map = new Map();
  rows.forEach((row) => {
    const date = row[dateColumn.header];
    const value = Number(row[metric.header]);
    if (!(date instanceof Date) || !Number.isFinite(value)) return;
    const key = dateKey(date);
    map.set(key, (map.get(key) || 0) + value);
  });
  return [...map.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, value]) => ({
      date: key,
      label: new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short" }).format(new Date(`${key}T00:00:00`)),
      value: niceNumber(value),
    }));
}

function aggregateMonthly(rows, dateColumn, metric) {
  if (!dateColumn || !metric) return [];
  const map = new Map();
  rows.forEach((row) => {
    const date = row[dateColumn.header];
    const value = Number(row[metric.header]);
    if (!(date instanceof Date) || !Number.isFinite(value)) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    map.set(key, (map.get(key) || 0) + value);
  });
  return [...map.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, value]) => ({
      month: new Intl.DateTimeFormat("en-IN", { month: "short", year: "numeric" }).format(new Date(`${key}-01T00:00:00`)),
      value: niceNumber(value),
    }));
}

function getNumericStats(rows, metric) {
  if (!metric) return { total: 0, average: 0, min: 0, max: 0, count: 0 };
  const values = rows.map((r) => Number(r[metric.header])).filter(Number.isFinite);
  if (!values.length) return { total: 0, average: 0, min: 0, max: 0, count: 0 };
  let total = 0;
  let min = values[0];
  let max = values[0];
  for (let i = 0; i < values.length; i += 1) {
    const value = values[i];
    total += value;
    if (value < min) min = value;
    if (value > max) max = value;
  }
  return {
    total: niceNumber(total),
    average: niceNumber(total / values.length),
    min,
    max,
    count: values.length,
  };
}

function getDateBounds(rows, dateColumn) {
  if (!dateColumn) return { min: null, max: null };
  const dates = rows.map((r) => r[dateColumn.header]).filter((d) => d instanceof Date && !Number.isNaN(d.getTime()));
  if (!dates.length) return { min: null, max: null };
  let min = dates[0];
  let max = dates[0];
  for (let i = 1; i < dates.length; i += 1) {
    if (dates[i] < min) min = dates[i];
    if (dates[i] > max) max = dates[i];
  }
  return { min, max };
}

function percentageChange(current, previous) {
  if (!Number.isFinite(previous) || previous === 0) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function generateInsights(rows, columns, metric, dimension, dateColumn) {
  const insights = [];
  if (!rows.length) return insights;

  const stats = getNumericStats(rows, metric);
  if (metric && stats.count) {
    insights.push(`Total ${metric.header} is ${formatFullNumber(stats.total)} across ${formatFullNumber(stats.count)} populated records.`);
    insights.push(`Average ${metric.header} is ${formatFullNumber(stats.average)}, with a range from ${formatFullNumber(stats.min)} to ${formatFullNumber(stats.max)}.`);
  }

  if (dimension && metric) {
    const ranking = aggregateByCategory(rows, dimension, metric);
    if (ranking.length) {
      const top = ranking[0];
      insights.push(`${top.name} is the leading ${dimension.header.toLowerCase()} by ${metric.header.toLowerCase()}, at ${formatFullNumber(top.value)}.`);
      if (ranking.length >= 3) {
        const share = stats.total ? (top.value / stats.total) * 100 : 0;
        insights.push(`The top ${dimension.header.toLowerCase()} contributes ${share.toFixed(1)}% of the total ${metric.header.toLowerCase()}.`);
      }
    }
  }

  if (dateColumn && metric) {
    const monthly = aggregateMonthly(rows, dateColumn, metric);
    if (monthly.length >= 2) {
      const last = monthly[monthly.length - 1].value;
      const previous = monthly[monthly.length - 2].value;
      const change = percentageChange(last, previous);
      if (change !== null) {
        insights.push(`${metric.header} changed ${Math.abs(change).toFixed(1)}% ${change >= 0 ? "up" : "down"} in the latest period versus the previous period.`);
      }
    }
  }

  const missing = columns
    .map((c) => ({
      header: c.header,
      missing: rows.filter((r) => r[c.header] === "" || r[c.header] === null || r[c.header] === undefined).length,
    }))
    .filter((x) => x.missing > 0)
    .sort((a, b) => b.missing - a.missing);

  if (missing.length) {
    const worst = missing[0];
    const pct = (worst.missing / rows.length) * 100;
    if (pct >= 10) insights.push(`${worst.header} has ${formatFullNumber(worst.missing)} missing values (${pct.toFixed(1)}% of records).`);
  }

  return insights.slice(0, 6);
}

function downloadBlob(content, name, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportWorkbook(rows, headers, filename) {
  const data = rows.map((row) =>
    Object.fromEntries(headers.map((h) => [h, row[h] instanceof Date ? dateKey(row[h]) : row[h] ?? ""]))
  );
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cleaned Data");
  XLSX.writeFile(wb, filename);
}

function Empty({ title = "No data available", text = "Upload a workbook to begin." }) {
  return (
    <div className="empty-state">
      <Database size={30} />
      <strong>{title}</strong>
      <span>{text}</span>
    </div>
  );
}

function KPI({ icon, label, value, detail }) {
  return (
    <article className="kpi-card">
      <div className="kpi-icon">{icon}</div>
      <div className="kpi-body">
        <div className="kpi-label">{label}</div>
        <div className="kpi-value" title={String(value)}>{value}</div>
        {detail && <div className="kpi-detail">{detail}</div>}
      </div>
    </article>
  );
}

function ChartCard({ title, subtitle, children, className = "" }) {
  return (
    <section className={`chart-card ${className}`}>
      <div className="chart-head">
        <div>
          <h3>{title}</h3>
          {subtitle && <p>{subtitle}</p>}
        </div>
      </div>
      <div className="chart-area">{children}</div>
    </section>
  );
}

function FilterSelect({ label, value, options, onChange }) {
  return (
    <label className="filter-control">
      <span>{label}</span>
      <div className="select-wrap">
        <select value={value} onChange={(e) => onChange(e.target.value)}>
          <option value="">All</option>
          {options.map((option) => <option key={option} value={option}>{option}</option>)}
        </select>
        <ChevronDown size={16} />
      </div>
    </label>
  );
}

function App() {
  const fileRef = useRef(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [filters, setFilters] = useState({});
  const [search, setSearch] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [theme, setTheme] = useState(() => localStorage.getItem("excel-dashboard-theme") || "light");
  const [showFilters, setShowFilters] = useState(true);
  const [showData, setShowData] = useState(false);
  const [visibleCharts, setVisibleCharts] = useState({
    trend: true,
    ranking: true,
    mix: true,
    monthly: true,
  });
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const sheet = workbook?.sheets?.[sheetIndex] || null;

  const dateColumn = useMemo(() => chooseDate(sheet?.columns || []), [sheet]);
  const primaryMetric = useMemo(() => choosePrimaryMetric(sheet?.columns || []), [sheet]);
  const secondaryMetric = useMemo(() => chooseSecondaryMetric(sheet?.columns || [], primaryMetric), [sheet, primaryMetric]);
  const primaryDimension = useMemo(() => chooseDimension(sheet?.columns || []), [sheet]);

  const dateBounds = useMemo(() => getDateBounds(sheet?.rows || [], dateColumn), [sheet, dateColumn]);

  const filterableColumns = useMemo(() => {
    if (!sheet) return [];
    return sheet.columns.filter((c) => c.isDimension && c.unique <= 100);
  }, [sheet]);

  const filteredRows = useMemo(() => {
    if (!sheet) return [];
    return sheet.rows.filter((row) => {
      if (dateColumn && (fromDate || toDate)) {
        const d = row[dateColumn.header];
        if (!(d instanceof Date)) return false;
        const key = dateKey(d);
        if (fromDate && key < fromDate) return false;
        if (toDate && key > toDate) return false;
      }

      for (const column of filterableColumns) {
        const selected = filters[column.header];
        if (selected && cleanText(row[column.header]) !== selected) return false;
      }

      if (search) {
        const q = search.toLowerCase();
        const found = sheet.headers.some((h) => cleanText(row[h]).toLowerCase().includes(q));
        if (!found) return false;
      }
      return true;
    });
  }, [sheet, dateColumn, fromDate, toDate, filters, filterableColumns, search]);

  const stats = useMemo(() => getNumericStats(filteredRows, primaryMetric), [filteredRows, primaryMetric]);
  const secondaryStats = useMemo(() => getNumericStats(filteredRows, secondaryMetric), [filteredRows, secondaryMetric]);

  const trendData = useMemo(() => aggregateTrend(filteredRows, dateColumn, primaryMetric), [filteredRows, dateColumn, primaryMetric]);
  const monthlyData = useMemo(() => aggregateMonthly(filteredRows, dateColumn, primaryMetric), [filteredRows, dateColumn, primaryMetric]);
  const rankingData = useMemo(() => aggregateByCategory(filteredRows, primaryDimension, primaryMetric).slice(0, 12), [filteredRows, primaryDimension, primaryMetric]);
  const mixData = useMemo(() => aggregateByCategory(filteredRows, primaryDimension, primaryMetric).slice(0, 8), [filteredRows, primaryDimension, primaryMetric]);

  const insights = useMemo(
    () => generateInsights(filteredRows, sheet?.columns || [], primaryMetric, primaryDimension, dateColumn),
    [filteredRows, sheet, primaryMetric, primaryDimension, dateColumn]
  );

  const uniqueOptions = useMemo(() => {
    if (!sheet) return {};
    return Object.fromEntries(
      filterableColumns.map((c) => [
        c.header,
        [...new Set(sheet.rows.map((r) => cleanText(r[c.header])).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 100),
      ])
    );
  }, [sheet, filterableColumns]);

  function resetFilters() {
    setFilters({});
    setFromDate("");
    setToDate("");
    setSearch("");
  }

  function applyQuickRange(type) {
    if (!dateBounds.min || !dateBounds.max) return;
    const max = new Date(dateBounds.max);
    let start = new Date(max);

    if (type === "latest") {
      const key = dateKey(max);
      setFromDate(key);
      setToDate(key);
      return;
    }
    if (type === "month") {
      start = new Date(max.getFullYear(), max.getMonth(), 1);
    } else if (type === "previous-month") {
      start = new Date(max.getFullYear(), max.getMonth() - 1, 1);
      const end = new Date(max.getFullYear(), max.getMonth(), 0);
      setFromDate(dateKey(start));
      setToDate(dateKey(end));
      return;
    } else if (type === "year") {
      start = new Date(max.getFullYear(), 0, 1);
    }
    setFromDate(dateKey(start));
    setToDate(dateKey(max));
  }

  async function handleFile(file) {
    if (!file) return;
    setError("");
    setLoading(true);
    try {
      // Give the browser one paint before starting the synchronous XLSX parser.
      await new Promise((resolve) => setTimeout(resolve, 30));
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheets = wb.SheetNames.map((name) => cleanSheet(wb.Sheets[name])).filter(Boolean);
      if (!sheets.length || sheets.every((s) => !s.rows.length)) {
        throw new Error("No usable table data was found in this workbook.");
      }
      setWorkbook({ fileName: file.name, sheets });
      setSheetIndex(0);
      resetFilters();
      setShowData(false);
    } catch (e) {
      console.error("Excel import failed:", e);
      setError(e?.message || "The file could not be read. Please check that it is a valid Excel or CSV file.");
      setWorkbook(null);
    } finally {
      setLoading(false);
    }
  }

  function onFileChange(e) {
    handleFile(e.target.files?.[0]);
    e.target.value = "";
  }

  function exportCSV() {
    if (!sheet) return;
    const data = filteredRows.map((row) =>
      Object.fromEntries(sheet.headers.map((h) => [h, row[h] instanceof Date ? dateKey(row[h]) : row[h] ?? ""]))
    );
    const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(data, { header: sheet.headers }));
    downloadBlob(csv, "dashboard-filtered-data.csv", "text/csv;charset=utf-8");
  }

  function toggleChart(key) {
    setVisibleCharts((v) => ({ ...v, [key]: !v[key] }));
  }

  function toggleTheme() {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    localStorage.setItem("excel-dashboard-theme", next);
  }

  const hasData = Boolean(sheet?.rows?.length);
  const exactTotal = formatFullNumber(stats.total);

  return (
    <div className={`app ${theme}`}>
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><BarChart3 size={19} /></div>
          <div>
            <strong>Excel Intelligence</strong>
            <span>Business Dashboard Studio</span>
          </div>
        </div>

        <div className="top-actions">
          {workbook && (
            <>
              <button className="btn secondary" onClick={() => fileRef.current?.click()}>
                <RefreshCcw size={16} /> Replace file
              </button>
              <button className="btn secondary" onClick={() => window.print()}>
                <Printer size={16} /> Print / PDF
              </button>
            </>
          )}
          <button className="icon-btn" title="Toggle theme" onClick={toggleTheme}>
            {theme === "light" ? <Moon size={18} /> : <Sun size={18} />}
          </button>
        </div>
      </header>

      {loading && (
        <div className="loading-overlay" role="status" aria-live="polite">
          <div className="loading-card">
            <div className="loading-spinner" />
            <strong>Reading your Excel file…</strong>
            <span>Detecting headers, dates, values and categories.</span>
          </div>
        </div>
      )}

      <main>
        {!workbook ? (
          <section className="hero">
            <div className="eyebrow"><Sparkles size={14} /> 100% browser-based · no API key</div>
            <h1>Turn messy Excel data into a <span>decision-ready dashboard.</span></h1>
            <p>
              Upload a workbook and the app detects headers, dates, metrics and dimensions,
              cleans the table, creates useful filters, builds large visualizations and writes automatic business insights.
            </p>

            <button className="upload-card" onClick={() => fileRef.current?.click()}>
              <div className="upload-icon"><Upload size={25} /></div>
              <strong>Upload your Excel file</strong>
              <span>Choose .xlsx, .xls or .csv from your device</span>
              <small>Data stays in your browser. Your workbook is not uploaded to a server.</small>
            </button>

            <div className="feature-grid">
              <div><WandSparkles size={18} /><strong>Automatic cleaning</strong><span>Detects headers, types and messy values.</span></div>
              <div><Filter size={18} /><strong>Dynamic filters</strong><span>Filters are generated from your actual columns.</span></div>
              <div><BarChart3 size={18} /><strong>Smart visuals</strong><span>Large charts are selected from the data.</span></div>
              <div><Sparkles size={18} /><strong>Automatic insights</strong><span>Highlights trends, leaders and data quality.</span></div>
            </div>

            {error && <div className="error-box"><AlertTriangle size={18} /> {error}</div>}
          </section>
        ) : (
          <>
            <section className="dashboard-title">
              <div>
                <div className="eyebrow"><FileSpreadsheet size={14} /> {workbook.fileName}</div>
                <h1>{primaryMetric ? `${primaryMetric.header} Performance Dashboard` : "Data Intelligence Dashboard"}</h1>
                <p>{sheet?.name} · {formatFullNumber(filteredRows.length)} of {formatFullNumber(sheet.rows.length)} rows currently included</p>
              </div>
              <div className="title-actions">
                <button className="btn primary" onClick={() => fileRef.current?.click()}><Upload size={16} /> New file</button>
              </div>
            </section>

            {error && <div className="error-box"><AlertTriangle size={18} /> {error}</div>}

            <section className="toolbar">
              <div className="toolbar-left">
                <button className={`btn ${showFilters ? "active" : "secondary"}`} onClick={() => setShowFilters((v) => !v)}>
                  <SlidersHorizontal size={16} /> Filters
                </button>
                <label className="search-box">
                  <Search size={16} />
                  <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search all columns..." />
                </label>
              </div>
              <div className="toolbar-right">
                <button className="btn secondary" onClick={resetFilters}><X size={15} /> Clear</button>
                <button className="btn secondary" onClick={exportCSV}><Download size={15} /> CSV</button>
                <button className="btn secondary" onClick={() => exportWorkbook(filteredRows, sheet.headers, "cleaned-dashboard-data.xlsx")}><FileDown size={15} /> Clean Excel</button>
              </div>
            </section>

            {showFilters && (
              <section className="filter-panel">
                <div className="filter-panel-head">
                  <div><strong>Control your dashboard</strong><span>Every available filter comes from the uploaded workbook.</span></div>
                  <button className="icon-btn" onClick={() => setShowFilters(false)}><X size={17} /></button>
                </div>

                <div className="filter-grid">
                  {dateColumn && (
                    <>
                      <label className="filter-control">
                        <span><CalendarDays size={14} /> From date</span>
                        <input type="date" value={fromDate} min={formatDateInput(dateBounds.min)} max={formatDateInput(dateBounds.max)} onChange={(e) => setFromDate(e.target.value)} />
                      </label>
                      <label className="filter-control">
                        <span><CalendarDays size={14} /> To date</span>
                        <input type="date" value={toDate} min={formatDateInput(dateBounds.min)} max={formatDateInput(dateBounds.max)} onChange={(e) => setToDate(e.target.value)} />
                      </label>
                    </>
                  )}

                  {filterableColumns.map((column) => (
                    <FilterSelect
                      key={column.header}
                      label={column.header}
                      value={filters[column.header] || ""}
                      options={uniqueOptions[column.header] || []}
                      onChange={(value) => setFilters((old) => ({ ...old, [column.header]: value }))}
                    />
                  ))}
                </div>

                {dateColumn && (
                  <div className="quick-row">
                    <span>Quick range:</span>
                    <button onClick={() => applyQuickRange("latest")}>Latest day</button>
                    <button onClick={() => applyQuickRange("month")}>Latest month</button>
                    <button onClick={() => applyQuickRange("previous-month")}>Previous month</button>
                    <button onClick={() => applyQuickRange("year")}>Latest year</button>
                  </div>
                )}
              </section>
            )}

            <section className="status-strip">
              <div><Database size={16} /><strong>{formatFullNumber(sheet.rows.length)}</strong><span>Total rows</span></div>
              <div><Columns3 size={16} /><strong>{formatFullNumber(sheet.columns.length)}</strong><span>Columns detected</span></div>
              <div><CalendarDays size={16} /><strong>{dateColumn ? "Detected" : "None"}</strong><span>Date field</span></div>
              <div><BarChart3 size={16} /><strong>{primaryMetric?.header || "None"}</strong><span>Primary metric</span></div>
            </section>

            <section className="kpi-grid">
              <KPI icon={<BarChart3 size={18} />} label={`Total ${primaryMetric?.header || "Value"}`} value={exactTotal} detail="Exact filtered total" />
              <KPI icon={<Database size={18} />} label="Records included" value={formatFullNumber(filteredRows.length)} detail={`of ${formatFullNumber(sheet.rows.length)} source rows`} />
              <KPI icon={<Sparkles size={18} />} label={`Average ${primaryMetric?.header || "Value"}`} value={formatFullNumber(stats.average)} detail="Exact average" />
              <KPI icon={<Eye size={18} />} label={`Highest ${primaryMetric?.header || "Value"}`} value={formatFullNumber(stats.max)} detail="Highest populated value" />
            </section>

            <section className="insight-card">
              <div className="insight-title"><Sparkles size={18} /><div><strong>Automatic business insights</strong><span>Rule-based analysis generated locally from your current filters.</span></div></div>
              {insights.length ? (
                <div className="insight-list">{insights.map((item, i) => <div key={i}><Check size={15} />{item}</div>)}</div>
              ) : (
                <div className="muted">Not enough structured data to generate insights yet.</div>
              )}
            </section>

            <section className="chart-controls">
              <strong>Dashboard sections</strong>
              <div>
                {[
                  ["trend", "Daily trend"],
                  ["monthly", "Monthly trend"],
                  ["ranking", "Category ranking"],
                  ["mix", "Category mix"],
                ].map(([key, label]) => (
                  <button key={key} className={visibleCharts[key] ? "chip active" : "chip"} onClick={() => toggleChart(key)}>
                    {visibleCharts[key] && <Check size={13} />} {label}
                  </button>
                ))}
              </div>
            </section>

            <section className="chart-grid">
              {visibleCharts.trend && (
                <ChartCard
                  title={`${primaryMetric?.header || "Value"} trend over time`}
                  subtitle={dateColumn ? `Based on ${dateColumn.header}` : "A date field was not detected in this sheet."}
                  className="chart-wide"
                >
                  {trendData.length ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={trendData} margin={{ top: 15, right: 20, left: 15, bottom: 10 }}>
                        <defs>
                          <linearGradient id="trendFill" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="var(--accent)" stopOpacity={0.35} />
                            <stop offset="100%" stopColor="var(--accent)" stopOpacity={0.02} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
                        <XAxis dataKey="label" tick={{ fill: "var(--muted)", fontSize: 12 }} />
                        <YAxis tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} width={70} />
                        <Tooltip formatter={(v) => [formatFullNumber(v), primaryMetric?.header || "Value"]} />
                        <Area type="monotone" dataKey="value" stroke="var(--accent)" fill="url(#trendFill)" strokeWidth={3} />
                      </AreaChart>
                    </ResponsiveContainer>
                  ) : <Empty title="Trend not available" text="A usable date and numeric field are required for this chart." />}
                </ChartCard>
              )}

              {visibleCharts.monthly && (
                <ChartCard
                  title={`Monthly ${primaryMetric?.header || "value"}`}
                  subtitle="Full values are available in the tooltip."
                >
                  {monthlyData.length ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={monthlyData} margin={{ top: 15, right: 20, left: 10, bottom: 10 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" />
                        <XAxis dataKey="month" tick={{ fill: "var(--muted)", fontSize: 12 }} />
                        <YAxis tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} width={70} />
                        <Tooltip formatter={(v) => [formatFullNumber(v), primaryMetric?.header || "Value"]} />
                        <Bar dataKey="value" fill="var(--accent)" radius={[8, 8, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  ) : <Empty title="Monthly view not available" text="A date field and numeric field are needed." />}
                </ChartCard>
              )}

              {visibleCharts.ranking && (
                <ChartCard
                  title={`Top ${primaryDimension?.header || "categories"} by ${primaryMetric?.header || "value"}`}
                  subtitle="Ranked from the currently filtered data."
                  className="chart-wide"
                >
                  {rankingData.length ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={rankingData} layout="vertical" margin={{ top: 10, right: 25, left: 30, bottom: 10 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" horizontal={false} />
                        <XAxis type="number" tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} />
                        <YAxis type="category" dataKey="name" width={130} tick={{ fill: "var(--text)", fontSize: 12 }} />
                        <Tooltip formatter={(v) => [formatFullNumber(v), primaryMetric?.header || "Value"]} />
                        <Bar dataKey="value" fill="var(--accent-2)" radius={[0, 8, 8, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  ) : <Empty title="Category ranking not available" text="A category-like column and numeric metric are required." />}
                </ChartCard>
              )}

              {visibleCharts.mix && (
                <ChartCard
                  title={`${primaryMetric?.header || "Value"} mix by ${primaryDimension?.header || "category"}`}
                  subtitle="Top categories shown individually; the remainder is grouped."
                >
                  {mixData.length ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie data={mixData} dataKey="value" nameKey="name" cx="50%" cy="48%" outerRadius="72%" innerRadius="42%" paddingAngle={2}>
                          {mixData.map((entry, index) => <Cell key={entry.name} fill={`hsl(${(index * 43) % 360} 70% 55%)`} />)}
                        </Pie>
                        <Tooltip formatter={(v) => [formatFullNumber(v), primaryMetric?.header || "Value"]} />
                        <Legend />
                      </PieChart>
                    </ResponsiveContainer>
                  ) : <Empty title="Category mix not available" text="A category-like column and numeric metric are required." />}
                </ChartCard>
              )}
            </section>

            <section className="data-section">
              <div className="data-head">
                <div><h2>Cleaned data preview</h2><p>Detected headers are normalized for readability. Your original workbook is not modified.</p></div>
                <button className="btn secondary" onClick={() => setShowData((v) => !v)}>{showData ? "Hide data" : "Show data"}</button>
              </div>

              {showData && (
                <div className="table-wrap">
                  <table>
                    <thead><tr>{sheet.headers.map((h) => <th key={h}>{h}</th>)}</tr></thead>
                    <tbody>
                      {filteredRows.slice(0, 100).map((row, i) => (
                        <tr key={i}>{sheet.headers.map((h) => (
                          <td key={h}>{row[h] instanceof Date ? formatDate(row[h]) : typeof row[h] === "number" ? formatFullNumber(row[h]) : row[h] || "—"}</td>
                        ))}</tr>
                      ))}
                    </tbody>
                  </table>
                  {!filteredRows.length && <Empty title="No rows match these filters" text="Clear or adjust the filters to see data." />}
                </div>
              )}
            </section>

            <footer className="footer-note">
              <span>Browser-only processing · No external API · Exact values preserved</span>
              <span>Detected header row: {sheet.headerRow} · Source rows: {formatFullNumber(sheet.sourceRows)}</span>
            </footer>
          </>
        )}
      </main>

      <input
        ref={fileRef}
        className="hidden-input"
        type="file"
        accept=".xlsx,.xls,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        onChange={onFileChange}
      />
    </div>
  );
}

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, message: "" };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, message: error?.message || "Unexpected dashboard error." };
  }

  componentDidCatch(error) {
    console.error("Dashboard render error:", error);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="fatal-error">
          <div className="fatal-card">
            <AlertTriangle size={28} />
            <h1>Dashboard could not display this workbook</h1>
            <p>{this.state.message}</p>
            <button className="btn primary" onClick={() => window.location.reload()}>Reload dashboard</button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

createRoot(document.getElementById("root")).render(
  <AppErrorBoundary>
    <App />
  </AppErrorBoundary>
);
