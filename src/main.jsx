import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AreaChart, Area, BarChart, Bar, CartesianGrid, Cell, PieChart, Pie,
  XAxis, YAxis, Tooltip, Legend, ResponsiveContainer
} from "recharts";
import {
  AlertTriangle, BarChart3, CalendarDays, Check, ChevronDown, Columns3,
  Database, Download, Eye, FileDown, FileSpreadsheet, Filter, Moon,
  Printer, RefreshCcw, Search, SlidersHorizontal, Sparkles, Sun, Upload,
  WandSparkles, X
} from "lucide-react";
import * as XLSX from "xlsx";
import "./styles.css";

const MAX_FILE_MB = 150;
const PREVIEW_ROWS = 100;

function isFiniteNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function safeText(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  try { return JSON.stringify(value); } catch { return String(value); }
}

function formatFullNumber(value) {
  if (!isFiniteNumber(Number(value))) return "—";
  return new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 }).format(Number(value));
}

function formatCompact(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  if (abs >= 1e9) return `${(n / 1e9).toFixed(1)}B`;
  if (abs >= 1e6) return `${(n / 1e6).toFixed(1)}M`;
  if (abs >= 1e3) return `${(n / 1e3).toFixed(1)}K`;
  return formatFullNumber(n);
}

function parseDate(value) {
  if (!value) return null;
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [y, m, d] = value.split("-").map(Number);
    const date = new Date(y, m - 1, d);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function dateKey(value) {
  const date = value instanceof Date ? value : parseDate(value);
  if (!date) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDate(value) {
  const date = parseDate(value);
  return date ? new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short", year: "numeric" }).format(date) : "—";
}

function formatCell(value) {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "number") return formatFullNumber(value);
  return safeText(value) || "—";
}

function aggregateCategory(rows, dimension, metric) {
  if (!dimension || !metric) return [];
  const map = new Map();
  for (const row of rows) {
    const key = safeText(row[dimension.header]) || "Blank";
    const value = Number(row[metric.header]);
    if (!Number.isFinite(value)) continue;
    map.set(key, (map.get(key) || 0) + value);
  }
  return [...map.entries()]
    .map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }))
    .sort((a, b) => b.value - a.value);
}

function aggregateTrend(rows, dateColumn, metric) {
  if (!dateColumn || !metric) return [];
  const map = new Map();
  for (const row of rows) {
    const key = dateKey(row[dateColumn.header]);
    const value = Number(row[metric.header]);
    if (!key || !Number.isFinite(value)) continue;
    map.set(key, (map.get(key) || 0) + value);
  }
  return [...map.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, value]) => ({
    date: key,
    label: new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short" }).format(parseDate(key)),
    value: Math.round(value * 100) / 100
  }));
}

function aggregateMonthly(rows, dateColumn, metric) {
  if (!dateColumn || !metric) return [];
  const map = new Map();
  for (const row of rows) {
    const date = parseDate(row[dateColumn.header]);
    const value = Number(row[metric.header]);
    if (!date || !Number.isFinite(value)) continue;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    map.set(key, (map.get(key) || 0) + value);
  }
  return [...map.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, value]) => ({
    month: new Intl.DateTimeFormat("en-IN", { month: "short", year: "numeric" }).format(new Date(`${key}-01T00:00:00`)),
    value: Math.round(value * 100) / 100
  }));
}

function getStats(rows, metric) {
  if (!metric) return { total: 0, average: 0, min: 0, max: 0, count: 0 };
  let total = 0;
  let count = 0;
  let min = Infinity;
  let max = -Infinity;
  for (const row of rows) {
    const value = Number(row[metric.header]);
    if (!Number.isFinite(value)) continue;
    total += value;
    count += 1;
    if (value < min) min = value;
    if (value > max) max = value;
  }
  if (!count) return { total: 0, average: 0, min: 0, max: 0, count: 0 };
  return { total, average: total / count, min, max, count };
}

function getDateBounds(rows, dateColumn) {
  if (!dateColumn) return { min: null, max: null };
  let min = null;
  let max = null;
  for (const row of rows) {
    const date = parseDate(row[dateColumn.header]);
    if (!date) continue;
    if (!min || date < min) min = date;
    if (!max || date > max) max = date;
  }
  return { min, max };
}

function chooseMetric(columns) {
  const numeric = columns.filter((column) => column.isNumeric);
  if (!numeric.length) return null;
  return numeric.find((column) => column.measureHint) || numeric[0];
}

function chooseSecondaryMetric(columns, primary) {
  return columns.find((column) => column.isNumeric && column.header !== primary?.header) || null;
}

function chooseDimension(columns) {
  const candidates = columns.filter((column) => column.isDimension && column.unique <= 200);
  if (!candidates.length) return null;
  return [...candidates].sort((a, b) => {
    const aDistance = Math.abs(Math.log10(Math.max(a.unique, 1)) - 1.5);
    const bDistance = Math.abs(Math.log10(Math.max(b.unique, 1)) - 1.5);
    return aDistance - bDistance;
  })[0];
}

function chooseDate(columns) {
  return columns.find((column) => column.isDate) || null;
}

function changePercent(current, previous) {
  if (!Number.isFinite(previous) || previous === 0) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function generateInsights(rows, columns, metric, dimension, dateColumn) {
  const insights = [];
  if (!rows.length) return insights;
  const stats = getStats(rows, metric);

  if (metric && stats.count) {
    insights.push(`Total ${metric.header} is ${formatFullNumber(stats.total)} across ${formatFullNumber(stats.count)} populated records.`);
    insights.push(`Average ${metric.header} is ${formatFullNumber(stats.average)}, ranging from ${formatFullNumber(stats.min)} to ${formatFullNumber(stats.max)}.`);
  }

  if (dimension && metric) {
    const ranking = aggregateCategory(rows, dimension, metric);
    if (ranking.length) {
      const top = ranking[0];
      insights.push(`${top.name} is the leading ${dimension.header.toLowerCase()} by ${metric.header.toLowerCase()}, at ${formatFullNumber(top.value)}.`);
      if (stats.total !== 0) insights.push(`The leading ${dimension.header.toLowerCase()} represents ${((top.value / stats.total) * 100).toFixed(1)}% of total ${metric.header.toLowerCase()}.`);
    }
  }

  if (dateColumn && metric) {
    const monthly = aggregateMonthly(rows, dateColumn, metric);
    if (monthly.length >= 2) {
      const latest = monthly[monthly.length - 1].value;
      const previous = monthly[monthly.length - 2].value;
      const change = changePercent(latest, previous);
      if (change !== null) insights.push(`${metric.header} moved ${Math.abs(change).toFixed(1)}% ${change >= 0 ? "up" : "down"} in the latest month versus the previous month.`);
    }
  }

  let worst = null;
  for (const column of columns) {
    let missing = 0;
    for (const row of rows) if (row[column.header] === "" || row[column.header] === null || row[column.header] === undefined) missing += 1;
    if (!worst || missing > worst.missing) worst = { header: column.header, missing };
  }
  if (worst?.missing) {
    const percent = (worst.missing / rows.length) * 100;
    if (percent >= 10) insights.push(`${worst.header} has ${formatFullNumber(worst.missing)} missing values (${percent.toFixed(1)}% of the filtered records).`);
  }
  return insights.slice(0, 6);
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportExcel(rows, headers) {
  const data = rows.map((row) => Object.fromEntries(headers.map((header) => [header, row[header] ?? ""])));
  const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Cleaned Data");
  XLSX.writeFile(workbook, "cleaned-dashboard-data.xlsx");
}

function exportCSV(rows, headers) {
  const data = rows.map((row) => Object.fromEntries(headers.map((header) => [header, row[header] ?? ""])));
  const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  downloadBlob(csv, "dashboard-filtered-data.csv", "text/csv;charset=utf-8");
}

function Empty({ title, text }) {
  return <div className="empty-state"><Database size={30} /><strong>{title}</strong><span>{text}</span></div>;
}

function KPI({ icon, label, value, detail }) {
  return <article className="kpi-card"><div className="kpi-icon">{icon}</div><div className="kpi-body"><div className="kpi-label">{label}</div><div className="kpi-value" title={String(value)}>{value}</div><div className="kpi-detail">{detail}</div></div></article>;
}

function ChartCard({ title, subtitle, children, wide = false }) {
  return <section className={`chart-card${wide ? " chart-wide" : ""}`}><div className="chart-head"><h3>{title}</h3>{subtitle && <p>{subtitle}</p>}</div><div className="chart-area">{children}</div></section>;
}

function FilterSelect({ label, value, options, onChange }) {
  return <label className="filter-control"><span>{label}</span><div className="select-wrap"><select value={value} onChange={(event) => onChange(event.target.value)}><option value="">All</option>{options.map((option) => <option key={option} value={option}>{option}</option>)}</select><ChevronDown size={16} /></div></label>;
}

function App() {
  const fileInputRef = useRef(null);
  const workerRef = useRef(null);
  const requestIdRef = useRef(0);

  const [workbook, setWorkbook] = useState(null);
  const [sheetIndex, setSheetIndex] = useState(0);
  const [filters, setFilters] = useState({});
  const [search, setSearch] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [theme, setTheme] = useState(() => {
    try { return localStorage.getItem("excel-dashboard-theme") || "light"; } catch { return "light"; }
  });
  const [showFilters, setShowFilters] = useState(true);
  const [showData, setShowData] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [charts, setCharts] = useState({ trend: true, monthly: true, ranking: true, mix: true });

  useEffect(() => {
    const worker = new Worker(new URL("./excelWorker.js", import.meta.url), { type: "module" });
    workerRef.current = worker;
    worker.onerror = (event) => {
      console.error("Excel worker error:", event);
      setLoading(false);
      setError("The Excel processing engine stopped unexpectedly. Please reload the dashboard and try again.");
    };
    return () => { worker.terminate(); workerRef.current = null; };
  }, []);

  useEffect(() => {
    try { localStorage.setItem("excel-dashboard-theme", theme); } catch { /* storage may be unavailable */ }
  }, [theme]);

  const sheet = workbook?.sheets?.[sheetIndex] || null;
  const dateColumn = useMemo(() => chooseDate(sheet?.columns || []), [sheet]);
  const primaryMetric = useMemo(() => chooseMetric(sheet?.columns || []), [sheet]);
  const secondaryMetric = useMemo(() => chooseSecondaryMetric(sheet?.columns || [], primaryMetric), [sheet, primaryMetric]);
  const primaryDimension = useMemo(() => chooseDimension(sheet?.columns || []), [sheet]);
  const dateBounds = useMemo(() => getDateBounds(sheet?.rows || [], dateColumn), [sheet, dateColumn]);

  const filterableColumns = useMemo(() => (sheet ? sheet.columns.filter((column) => column.isDimension && column.unique <= 100) : []), [sheet]);
  const uniqueOptions = useMemo(() => {
    if (!sheet) return {};
    const result = {};
    for (const column of filterableColumns) {
      const set = new Set();
      for (const row of sheet.rows) {
        const value = safeText(row[column.header]);
        if (value) set.add(value);
        if (set.size > 100) break;
      }
      result[column.header] = [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 100);
    }
    return result;
  }, [sheet, filterableColumns]);

  const filteredRows = useMemo(() => {
    if (!sheet) return [];
    const query = search.trim().toLowerCase();
    return sheet.rows.filter((row) => {
      if (dateColumn && (fromDate || toDate)) {
        const key = dateKey(row[dateColumn.header]);
        if (!key || (fromDate && key < fromDate) || (toDate && key > toDate)) return false;
      }
      for (const column of filterableColumns) {
        const selected = filters[column.header];
        if (selected && safeText(row[column.header]) !== selected) return false;
      }
      if (query) {
        let found = false;
        for (const header of sheet.headers) {
          if (safeText(row[header]).toLowerCase().includes(query)) { found = true; break; }
        }
        if (!found) return false;
      }
      return true;
    });
  }, [sheet, dateColumn, fromDate, toDate, filters, filterableColumns, search]);

  const stats = useMemo(() => getStats(filteredRows, primaryMetric), [filteredRows, primaryMetric]);
  const trendData = useMemo(() => aggregateTrend(filteredRows, dateColumn, primaryMetric), [filteredRows, dateColumn, primaryMetric]);
  const monthlyData = useMemo(() => aggregateMonthly(filteredRows, dateColumn, primaryMetric), [filteredRows, dateColumn, primaryMetric]);
  const rankingData = useMemo(() => aggregateCategory(filteredRows, primaryDimension, primaryMetric).slice(0, 12), [filteredRows, primaryDimension, primaryMetric]);
  const mixData = useMemo(() => aggregateCategory(filteredRows, primaryDimension, primaryMetric).slice(0, 8), [filteredRows, primaryDimension, primaryMetric]);
  const insights = useMemo(() => generateInsights(filteredRows, sheet?.columns || [], primaryMetric, primaryDimension, dateColumn), [filteredRows, sheet, primaryMetric, primaryDimension, dateColumn]);

  function resetFilters() {
    setFilters({}); setSearch(""); setFromDate(""); setToDate("");
  }

  function applyQuickRange(type) {
    if (!dateBounds.max) return;
    const max = new Date(dateBounds.max);
    let start = new Date(max);
    if (type === "latest") { setFromDate(dateKey(max)); setToDate(dateKey(max)); return; }
    if (type === "month") start = new Date(max.getFullYear(), max.getMonth(), 1);
    if (type === "previous-month") {
      start = new Date(max.getFullYear(), max.getMonth() - 1, 1);
      const end = new Date(max.getFullYear(), max.getMonth(), 0);
      setFromDate(dateKey(start)); setToDate(dateKey(end)); return;
    }
    if (type === "year") start = new Date(max.getFullYear(), 0, 1);
    setFromDate(dateKey(start)); setToDate(dateKey(max));
  }

  function handleFile(file) {
    if (!file || !workerRef.current) return;
    setLoading(true); setError(""); setWorkbook(null); setShowData(false);
    if (file.size > MAX_FILE_MB * 1024 * 1024) {
      setLoading(false);
      setError(`This file is ${Math.ceil(file.size / 1024 / 1024)} MB. For reliable browser processing, please use a file smaller than ${MAX_FILE_MB} MB.`);
      return;
    }

    const id = ++requestIdRef.current;
    file.arrayBuffer().then((buffer) => {
      if (id !== requestIdRef.current || !workerRef.current) return;
      workerRef.current.postMessage({ id, buffer }, [buffer]);
    }).catch((cause) => {
      console.error("File read failed:", cause);
      if (id === requestIdRef.current) { setLoading(false); setError("The selected file could not be read from the device."); }
    });

    const worker = workerRef.current;
    const onMessage = (event) => {
      const message = event.data || {};
      if (message.id !== id) return;
      worker.removeEventListener("message", onMessage);
      if (id !== requestIdRef.current) return;
      if (!message.ok) {
        setLoading(false); setError(message.error || "The workbook could not be processed."); return;
      }
      setWorkbook({ fileName: file.name, sheets: message.result.sheets });
      setSheetIndex(0); resetFilters(); setLoading(false);
    };
    worker.addEventListener("message", onMessage);
  }

  function onFileChange(event) {
    const file = event.target.files?.[0];
    event.target.value = "";
    handleFile(file);
  }

  function toggleChart(name) { setCharts((current) => ({ ...current, [name]: !current[name] })); }
  function toggleTheme() { setTheme((current) => current === "light" ? "dark" : "light"); }

  return (
    <div className={`app ${theme}`}>
      <header className="topbar">
        <div className="brand"><div className="brand-mark"><BarChart3 size={19} /></div><div><strong>Excel Intelligence</strong><span>Business Dashboard Studio</span></div></div>
        <div className="top-actions">
          {workbook && <><button className="btn secondary" onClick={() => fileInputRef.current?.click()}><RefreshCcw size={16} /> Replace file</button><button className="btn secondary" onClick={() => window.print()}><Printer size={16} /> Print / PDF</button></>}
          <button className="icon-btn" title="Toggle theme" onClick={toggleTheme}>{theme === "light" ? <Moon size={18} /> : <Sun size={18} />}</button>
        </div>
      </header>

      {loading && <div className="loading-overlay" role="status" aria-live="polite"><div className="loading-card"><div className="loading-spinner" /><strong>Processing your Excel file…</strong><span>Reading sheets and detecting headers, dates, values and categories.</span></div></div>}

      <main>
        {!workbook ? (
          <section className="hero">
            <div className="eyebrow"><Sparkles size={14} /> 100% browser-based · no API key</div>
            <h1>Turn messy Excel data into a <span>decision-ready dashboard.</span></h1>
            <p>Upload a workbook and the app automatically detects structure, cleans values, creates dynamic filters, builds large charts and generates local business insights.</p>
            <button className="upload-card" onClick={() => fileInputRef.current?.click()}>
              <div className="upload-icon"><Upload size={25} /></div><strong>Upload your Excel file</strong><span>Choose .xlsx, .xls or .csv from your device</span><small>Data is processed locally in your browser and is not sent to a server.</small>
            </button>
            <div className="feature-grid"><div><WandSparkles size={18} /><strong>Automatic cleaning</strong><span>Detects headers, types and messy values.</span></div><div><Filter size={18} /><strong>Dynamic filters</strong><span>Filters are generated from your actual columns.</span></div><div><BarChart3 size={18} /><strong>Smart visuals</strong><span>Charts adapt to the detected data structure.</span></div><div><Sparkles size={18} /><strong>Automatic insights</strong><span>Highlights trends, leaders and data quality.</span></div></div>
            {error && <div className="error-box"><AlertTriangle size={18} /> <span>{error}</span></div>}
          </section>
        ) : (
          <>
            <section className="dashboard-title"><div><div className="eyebrow"><FileSpreadsheet size={14} /> {workbook.fileName}</div><h1>{primaryMetric ? `${primaryMetric.header} Performance Dashboard` : "Data Intelligence Dashboard"}</h1><p>{sheet?.name} · {formatFullNumber(filteredRows.length)} of {formatFullNumber(sheet?.rows.length || 0)} rows currently included</p></div><button className="btn primary" onClick={() => fileInputRef.current?.click()}><Upload size={16} /> New file</button></section>
            {error && <div className="error-box"><AlertTriangle size={18} /> <span>{error}</span></div>}

            {workbook.sheets.length > 1 && <section className="sheet-tabs"><strong>Sheets</strong>{workbook.sheets.map((item, index) => <button key={item.name} className={index === sheetIndex ? "active" : ""} onClick={() => { setSheetIndex(index); resetFilters(); }}>{item.name}</button>)}</section>}

            <section className="toolbar"><div className="toolbar-left"><button className={`btn ${showFilters ? "active" : "secondary"}`} onClick={() => setShowFilters((value) => !value)}><SlidersHorizontal size={16} /> Filters</button><label className="search-box"><Search size={16} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search all columns..." /></label></div><div className="toolbar-right"><button className="btn secondary" onClick={resetFilters}><X size={15} /> Clear</button><button className="btn secondary" onClick={() => exportCSV(filteredRows, sheet.headers)}><Download size={15} /> CSV</button><button className="btn secondary" onClick={() => exportExcel(filteredRows, sheet.headers)}><FileDown size={15} /> Clean Excel</button></div></section>

            {showFilters && <section className="filter-panel"><div className="filter-panel-head"><div><strong>Dashboard controls</strong><span>Every filter is derived from the current worksheet.</span></div><button className="icon-btn" onClick={() => setShowFilters(false)}><X size={17} /></button></div><div className="filter-grid">
              {dateColumn && <><label className="filter-control"><span><CalendarDays size={14} /> From date</span><input type="date" value={fromDate} min={dateKey(dateBounds.min)} max={dateKey(dateBounds.max)} onChange={(event) => setFromDate(event.target.value)} /></label><label className="filter-control"><span><CalendarDays size={14} /> To date</span><input type="date" value={toDate} min={dateKey(dateBounds.min)} max={dateKey(dateBounds.max)} onChange={(event) => setToDate(event.target.value)} /></label></>}
              {filterableColumns.map((column) => <FilterSelect key={column.header} label={column.header} value={filters[column.header] || ""} options={uniqueOptions[column.header] || []} onChange={(value) => setFilters((current) => ({ ...current, [column.header]: value }))} />)}
            </div>{dateColumn && <div className="quick-row"><span>Quick range:</span><button onClick={() => applyQuickRange("latest")}>Latest day</button><button onClick={() => applyQuickRange("month")}>Latest month</button><button onClick={() => applyQuickRange("previous-month")}>Previous month</button><button onClick={() => applyQuickRange("year")}>Latest year</button></div>}</section>}

            <section className="status-strip"><div><Database size={16} /><strong>{formatFullNumber(sheet.rows.length)}</strong><span>Total rows</span></div><div><Columns3 size={16} /><strong>{formatFullNumber(sheet.columns.length)}</strong><span>Columns detected</span></div><div><CalendarDays size={16} /><strong>{dateColumn ? "Detected" : "None"}</strong><span>Date field</span></div><div><BarChart3 size={16} /><strong>{primaryMetric?.header || "None"}</strong><span>Primary metric</span></div></section>

            <section className="kpi-grid"><KPI icon={<BarChart3 size={18} />} label={`Total ${primaryMetric?.header || "Value"}`} value={formatFullNumber(stats.total)} detail="Exact filtered total" /><KPI icon={<Database size={18} />} label="Records included" value={formatFullNumber(filteredRows.length)} detail={`of ${formatFullNumber(sheet.rows.length)} source rows`} /><KPI icon={<Sparkles size={18} />} label={`Average ${primaryMetric?.header || "Value"}`} value={formatFullNumber(stats.average)} detail="Exact average" /><KPI icon={<Eye size={18} />} label={`Highest ${primaryMetric?.header || "Value"}`} value={formatFullNumber(stats.max)} detail="Highest populated value" /></section>

            <section className="insight-card"><div className="insight-title"><Sparkles size={18} /><div><strong>Automatic business insights</strong><span>Rule-based analysis generated locally from the current filtered data.</span></div></div>{insights.length ? <div className="insight-list">{insights.map((item) => <div key={item}><Check size={15} /> <span>{item}</span></div>)}</div> : <div className="muted">Not enough structured data to generate insights yet.</div>}</section>

            <section className="chart-controls"><strong>Dashboard sections</strong><div>{[["trend", "Daily trend"], ["monthly", "Monthly trend"], ["ranking", "Category ranking"], ["mix", "Category mix"]].map(([key, label]) => <button key={key} className={charts[key] ? "chip active" : "chip"} onClick={() => toggleChart(key)}>{charts[key] && <Check size={13} />} {label}</button>)}</div></section>

            <section className="chart-grid">
              {charts.trend && <ChartCard wide title={`${primaryMetric?.header || "Value"} trend over time`} subtitle={dateColumn ? `Based on ${dateColumn.header}` : "A date field was not detected in this sheet."}>{trendData.length ? <ResponsiveContainer width="100%" height="100%"><AreaChart data={trendData} margin={{ top: 15, right: 20, left: 15, bottom: 10 }}><defs><linearGradient id="trendFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="var(--accent)" stopOpacity={0.35} /><stop offset="100%" stopColor="var(--accent)" stopOpacity={0.03} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" /><XAxis dataKey="label" tick={{ fill: "var(--muted)", fontSize: 12 }} /><YAxis tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} width={70} /><Tooltip formatter={(value) => [formatFullNumber(value), primaryMetric?.header || "Value"]} /><Area type="monotone" dataKey="value" stroke="var(--accent)" fill="url(#trendFill)" strokeWidth={3} /></AreaChart></ResponsiveContainer> : <Empty title="Trend not available" text="A usable date and numeric field are required for this chart." />}</ChartCard>}

              {charts.monthly && <ChartCard title={`Monthly ${primaryMetric?.header || "value"}`} subtitle="Full values are available in the tooltip.">{monthlyData.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={monthlyData} margin={{ top: 15, right: 20, left: 10, bottom: 10 }}><CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" /><XAxis dataKey="month" tick={{ fill: "var(--muted)", fontSize: 12 }} /><YAxis tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} width={70} /><Tooltip formatter={(value) => [formatFullNumber(value), primaryMetric?.header || "Value"]} /><Bar dataKey="value" fill="var(--accent)" radius={[8, 8, 0, 0]} /></BarChart></ResponsiveContainer> : <Empty title="Monthly view not available" text="A date field and numeric field are needed." />}</ChartCard>}

              {charts.ranking && <ChartCard wide title={`Top ${primaryDimension?.header || "categories"} by ${primaryMetric?.header || "value"}`} subtitle="Ranked from the currently filtered data.">{rankingData.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={rankingData} layout="vertical" margin={{ top: 10, right: 25, left: 30, bottom: 10 }}><CartesianGrid strokeDasharray="3 3" stroke="var(--grid)" horizontal={false} /><XAxis type="number" tick={{ fill: "var(--muted)", fontSize: 12 }} tickFormatter={formatCompact} /><YAxis type="category" dataKey="name" width={140} tick={{ fill: "var(--text)", fontSize: 12 }} /><Tooltip formatter={(value) => [formatFullNumber(value), primaryMetric?.header || "Value"]} /><Bar dataKey="value" fill="var(--accent-2)" radius={[0, 8, 8, 0]} /></BarChart></ResponsiveContainer> : <Empty title="Category ranking not available" text="A category-like column and numeric metric are required." />}</ChartCard>}

              {charts.mix && <ChartCard title={`${primaryMetric?.header || "Value"} mix by ${primaryDimension?.header || "category"}`} subtitle="Top categories shown individually.">{mixData.length ? <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={mixData} dataKey="value" nameKey="name" cx="50%" cy="48%" outerRadius="72%" innerRadius="42%" paddingAngle={2}>{mixData.map((entry, index) => <Cell key={`${entry.name}-${index}`} fill={`hsl(${(index * 43) % 360} 70% 55%)`} />)}</Pie><Tooltip formatter={(value) => [formatFullNumber(value), primaryMetric?.header || "Value"]} /><Legend /></PieChart></ResponsiveContainer> : <Empty title="Category mix not available" text="A category-like column and numeric metric are required." />}</ChartCard>}
            </section>

            <section className="data-section"><div className="data-head"><div><h2>Cleaned data preview</h2><p>Values are normalized to safe browser data types. The original workbook is not modified.</p></div><button className="btn secondary" onClick={() => setShowData((value) => !value)}>{showData ? "Hide data" : "Show data"}</button></div>{showData && <div className="table-wrap">{filteredRows.length ? <table><thead><tr>{sheet.headers.map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{filteredRows.slice(0, PREVIEW_ROWS).map((row, index) => <tr key={`${row.__row}-${index}`}>{sheet.headers.map((header) => <td key={header}>{formatCell(row[header])}</td>)}</tr>)}</tbody></table> : <Empty title="No rows match these filters" text="Clear or adjust the filters to see data." />}</div>}</section>

            <footer className="footer-note"><span>Browser-only processing · No external API · Exact values shown in cards, tables and tooltips</span><span>Detected header row: {sheet.headerRow} · Source rows: {formatFullNumber(sheet.sourceRows)}</span></footer>
          </>
        )}
      </main>

      <input ref={fileInputRef} className="hidden-input" type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" onChange={onFileChange} />
    </div>
  );
}

class AppErrorBoundary extends React.Component {
  state = { hasError: false, message: "" };
  static getDerivedStateFromError(error) { return { hasError: true, message: error?.message || "Unexpected dashboard error." }; }
  componentDidCatch(error) { console.error("Dashboard render error:", error); }
  render() {
    if (!this.state.hasError) return this.props.children;
    return <div className="fatal-error"><div className="fatal-card"><AlertTriangle size={28} /><h1>Dashboard could not display this workbook</h1><p>{this.state.message}</p><button className="btn primary" onClick={() => window.location.reload()}>Reload dashboard</button></div></div>;
  }
}

createRoot(document.getElementById("root")).render(<AppErrorBoundary><App /></AppErrorBoundary>);
