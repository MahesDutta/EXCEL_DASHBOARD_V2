  
  }

      const message = event.data || {};
      if (message.id !== requestIdRef.current) return;

      if (message.type === "progress") {
        setProgress(Math.max(0, Math.min(100, Number(message.percent) || 0)));
        setProgressMessage(message.message || "Processing workbook…");
        return;
      }

      if (message.type !== "complete") return;
      setLoading(false);

      if (!message.ok) {
        setError(message.error || "The workbook could not be processed.");
        return;
      }

      setProgress(100);
      setProgressMessage("Workbook ready");
      setWorkbook({ fileName: pendingFileNameRef.current || "Workbook", sheets: message.result.sheets });
      setSheetIndex(0);
      resetFilters();
    };

    worker.onerror = (event) => {
      console.error("Excel worker error:", event);
      setLoading(false);
      setError(event?.message ? `Excel processor error: ${event.message}` : "The Excel processing engine stopped unexpectedly. Please reload the dashboard and try again.");
    };

    worker.onmessageerror = (event) => {
      console.error("Excel worker message error:", event);
      setLoading(false);
      setError("The Excel processor could not transfer the workbook data. Please reload the dashboard and try again.");
    };

    return () => {
      worker.terminate();
      workerRef.current = null;
    };
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

  async function handleFile(file) {
    if (!file) return;

    setLoading(true);
    setProgress(0);
    setProgressMessage("Preparing your Excel file…");
    setError("");
    setWorkbook(null);
    setShowData(false);

    if (file.size > MAX_FILE_MB * 1024 * 1024) {
      setLoading(false);
      setError(`This file is ${Math.ceil(file.size / 1024 / 1024)} MB. For reliable browser processing, please use a file smaller than ${MAX_FILE_MB} MB.`);
      return;
    }

    const id = ++requestIdRef.current;
    pendingFileNameRef.current = file.name;

    try {
      const buffer = await file.arrayBuffer();
      if (id !== requestIdRef.current) return;

      // Small and medium workbooks are processed directly. This removes the
      // most common module-worker failure mode on mobile browsers. Larger
      // workbooks stay off the UI thread through the existing worker.
      const DIRECT_LIMIT = 10 * 1024 * 1024;

      if (file.size <= DIRECT_LIMIT) {
        const result = await processWorkbookDirect(buffer, (percent, message) => {
          if (id === requestIdRef.current) {
            setProgress(percent);
            setProgressMessage(message);
          }
        });

        if (id !== requestIdRef.current) return;
        setWorkbook({ fileName: file.name || "Workbook", sheets: result.sheets });
        setSheetIndex(0);
        resetFilters();
        setProgress(100);
        setProgressMessage("Workbook ready");
        setLoading(false);
        return;
      }

      if (!workerRef.current) throw new Error("The Excel processing engine is unavailable. Please reload the dashboard.");

      setProgress(3);
      setProgressMessage("Sending workbook to the Excel processor…");
      workerRef.current.postMessage({ id, buffer }, [buffer]);
    } catch (cause) {
      console.error("Excel import failed:", cause);
      if (id === requestIdRef.current) {
        setLoading(false);
        setError(cause?.message || "The selected file could not be processed. Please try another Excel or CSV file.");
      }
    }
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

      {loading && <div className="loading-overlay" role="status" aria-live="polite"><div className="loading-card"><div className="loading-spinner" /><strong>Processing your Excel file…</strong><div className="progress-track"><div className="progress-fill" style={{ width: `${progress}%` }} /></div><div className="progress-meta"><span>{progressMessage}</span><strong>{progress}%</strong></div><span>This runs locally in your browser. Large workbooks may take longer.</span></div></div>}

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
