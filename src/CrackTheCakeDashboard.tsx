import { useState, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, RadarChart, Radar,
  PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer,
} from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// ── Types ─────────────────────────────────────────────────────────────────────
type Row = Record<string, string>;
type ChartType = "bar" | "horizontalBar" | "line" | "area" | "pie" | "radar";
type DataMode = "cat" | "multi";

interface ChartCfg {
  id: number;
  type: ChartType;
  mode: DataMode;
  col: string;
  title: string;
  filterCol: string;
  filterVals: string[];
}

// ── Colors ────────────────────────────────────────────────────────────────────
const COLORS = [
  "#FF6B9D","#FF9F43","#FECA57","#48DBFB",
  "#54A0FF","#5F27CD","#4BC99A","#F97B6B",
  "#B47FEB","#E06EBD","#F7934C","#7DC97E",
];

// ── File Parser ───────────────────────────────────────────────────────────────
function parseFile(file: File): Promise<Row[]> {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext === "csv") {
      // Try multiple encodings — handles Thai UTF-8 and Windows TIS-620
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target?.result as string;
        Papa.parse(text, {
          header: true,
          skipEmptyLines: true,
          dynamicTyping: false,
          transformHeader: (h) => h.trim(),
          transform: (v) => v.trim(),
          complete: (r) => resolve(r.data as Row[]),
          error: reject,
        });
      };
      reader.onerror = reject;
      // Use UTF-8 BOM-aware reading
      reader.readAsText(file, "UTF-8");
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target?.result, { type: "binary", codepage: 874 });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });
          // Trim keys and values
          const clean = json.map((row) =>
            Object.fromEntries(
              Object.entries(row).map(([k, v]) => [k.trim(), String(v).trim()])
            )
          );
          resolve(clean);
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    } else {
      reject(new Error("รองรับเฉพาะ .csv .xlsx .xls"));
    }
  });
}

// ── Auto-detect helpers ───────────────────────────────────────────────────────
const GEN_KEYWORDS = ["gen z","gen alpha","gen x","gen y","millennial","boomer","generation"];
const MULTI_KEYWORDS = ["select all","up to","choose up","you can answer","select all that apply"];

function detectGenCol(cols: string[], data: Row[]): string {
  const byName = cols.find((c) => /gen(eration)?|age.?group|cohort|ช่วงอายุ/i.test(c));
  if (byName) return byName;
  return cols.find((c) =>
    data.slice(0, 20).some((r) =>
      GEN_KEYWORDS.some((g) => String(r[c] ?? "").toLowerCase().includes(g))
    )
  ) ?? "";
}

function detectMultiCols(cols: string[]): string[] {
  return cols.filter((c) =>
    MULTI_KEYWORDS.some((k) => c.toLowerCase().includes(k))
  );
}

// ── Data helpers ──────────────────────────────────────────────────────────────
const shorten = (k: string, max = 44): string =>
  k.length > max ? k.slice(0, max) + "…" : k;

const allVals = (data: Row[], col: string): string[] =>
  [...new Set(data.map((r) => r[col] ?? ""))].filter(Boolean).sort();

function countCol(data: Row[], col: string): [string, number][] {
  const c: Record<string, number> = {};
  data.forEach((r) => {
    const v = r[col] ?? "";
    if (v) c[v] = (c[v] ?? 0) + 1;
  });
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}

function countMulti(data: Row[], col: string): [string, number][] {
  const c: Record<string, number> = {};
  data.forEach((r) =>
    (r[col] ?? "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((v) => { c[v] = (c[v] ?? 0) + 1; })
  );
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}

function applyFilters(
  data: Row[],
  genCol: string,
  gFilter: string[],
  lcol: string,
  lvals: string[]
): Row[] {
  let d = data;
  if (genCol && gFilter.length) d = d.filter((r) => gFilter.includes(r[genCol] ?? ""));
  if (lcol && lvals.length) d = d.filter((r) => lvals.includes(r[lcol] ?? ""));
  return d;
}

// ── Custom Tooltip ────────────────────────────────────────────────────────────
function CustomTooltip({
  active, payload, label, total,
}: {
  active?: boolean;
  payload?: { value: number }[];
  label?: string;
  total: number;
}) {
  if (!active || !payload?.length) return null;
  const val = payload[0].value;
  return (
    <div style={{ background: "#12122a", border: "1px solid #ffffff22", borderRadius: 10, padding: "8px 12px", color: "#fff", fontSize: 13 }}>
      <p style={{ color: "#aaa", marginBottom: 2, fontSize: 11 }}>{label}</p>
      <p style={{ fontWeight: 700 }}>
        {val} <span style={{ color: "#FF9F43", fontWeight: 400 }}>({((val / (total || 1)) * 100).toFixed(1)}%)</span>
      </p>
    </div>
  );
}

// ── Chart Card ────────────────────────────────────────────────────────────────
function ChartCard({
  cfg, data, genCol, gFilter, onEdit, onRemove,
}: {
  cfg: ChartCfg;
  data: Row[];
  genCol: string;
  gFilter: string[];
  onEdit: () => void;
  onRemove: () => void;
}) {
  const filtered = useMemo(
    () => applyFilters(data, genCol, gFilter, cfg.filterCol, cfg.filterVals),
    [data, genCol, gFilter, cfg]
  );
  const entries = useMemo(
    () => cfg.mode === "multi" ? countMulti(filtered, cfg.col) : countCol(filtered, cfg.col),
    [filtered, cfg]
  );
  const total = filtered.length;
  const chartData = entries.map(([name, value]) => ({ name, value }));

  const fLabel = [
    gFilter.length ? `Gen: ${gFilter.join(", ")}` : "",
    cfg.filterCol && cfg.filterVals.length ? `${shorten(cfg.filterCol, 22)}: ${cfg.filterVals.join(", ")}` : "",
  ].filter(Boolean).join(" | ") || "ทั้งหมด";

  const AXIS = { tick: { fill: "#888", fontSize: 10 }, grid: { stroke: "#ffffff0d" } };

  const renderInner = () => {
    if (cfg.type === "pie") {
      return (
        <PieChart>
          <Pie data={chartData} dataKey="value" cx="50%" cy="50%" innerRadius={55} outerRadius={100} paddingAngle={2}>
            {chartData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} stroke="transparent" />)}
          </Pie>
          <Tooltip content={<CustomTooltip total={total} />} />
          <Legend wrapperStyle={{ color: "#aaa", fontSize: 11 }} formatter={(v) => shorten(v, 26)} />
        </PieChart>
      );
    }
    if (cfg.type === "radar") {
      return (
        <RadarChart data={chartData.slice(0, 10)}>
          <PolarGrid stroke="#ffffff18" />
          <PolarAngleAxis dataKey="name" tick={{ fill: "#aaa", fontSize: 10 }} />
          <PolarRadiusAxis tick={{ fill: "#555", fontSize: 9 }} />
          <Radar name="จำนวน" dataKey="value" stroke={COLORS[0]} fill={COLORS[0]} fillOpacity={0.3} />
          <Tooltip content={<CustomTooltip total={total} />} />
        </RadarChart>
      );
    }
    const isH = cfg.type === "horizontalBar";
    const isLine = cfg.type === "line";
    const isArea = cfg.type === "area";
    const xAxis = isH
      ? <XAxis type="number" tick={AXIS.tick} />
      : <XAxis dataKey="name" tick={{ fill: "#888", fontSize: 10 }} angle={-35} textAnchor="end" interval={0} />;
    const yAxis = isH
      ? <YAxis type="category" dataKey="name" tick={{ fill: "#888", fontSize: 10 }} width={140} />
      : <YAxis tick={AXIS.tick} />;
    const margin = isH
      ? { top: 8, right: 40, left: 8, bottom: 8 }
      : { top: 8, right: 16, left: 0, bottom: 65 };
    if (isLine)
      return (
        <LineChart data={chartData} margin={margin}>
          <CartesianGrid strokeDasharray="3 3" stroke="#ffffff0d" />
          {xAxis}{yAxis}
          <Tooltip content={<CustomTooltip total={total} />} />
          <Line type="monotone" dataKey="value" stroke={COLORS[0]} strokeWidth={2.5} dot={{ fill: COLORS[0], r: 4 }} />
        </LineChart>
      );
    if (isArea)
      return (
        <AreaChart data={chartData} margin={margin}>
          <defs>
            <linearGradient id={`ag-${cfg.id}`} x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={COLORS[0]} stopOpacity={0.45} />
              <stop offset="95%" stopColor={COLORS[0]} stopOpacity={0} />
            </linearGradient>
          </defs>
          <CartesianGrid strokeDasharray="3 3" stroke="#ffffff0d" />
          {xAxis}{yAxis}
          <Tooltip content={<CustomTooltip total={total} />} />
          <Area type="monotone" dataKey="value" stroke={COLORS[0]} fill={`url(#ag-${cfg.id})`} strokeWidth={2} />
        </AreaChart>
      );
    return (
      <BarChart data={chartData} layout={isH ? "vertical" : "horizontal"} margin={margin}>
        <CartesianGrid strokeDasharray="3 3" stroke="#ffffff0d" />
        {xAxis}{yAxis}
        <Tooltip content={<CustomTooltip total={total} />} />
        <Bar dataKey="value" radius={isH ? [0, 5, 5, 0] : [5, 5, 0, 0]}>
          {chartData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
        </Bar>
      </BarChart>
    );
  };

  const h = cfg.type === "horizontalBar" ? Math.max(280, chartData.length * 30) : 300;

  return (
    <div style={{ background: "#12122a", border: "1px solid #ffffff10", borderRadius: 16, padding: "1.1rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: ".75rem", gap: 8 }}>
        <div>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 15, color: "#fff" }}>{cfg.title}</p>
          <p style={{ margin: "3px 0 0", fontSize: 11, color: "#666" }}>{fLabel} · {total} แถว</p>
        </div>
        <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
          <button onClick={onEdit} style={ICON_BTN_STYLE}>✏️</button>
          <button onClick={onRemove} style={{ ...ICON_BTN_STYLE, color: "#FF6B6B", borderColor: "#FF6B6B33" }}>✕</button>
        </div>
      </div>
      <div style={{ height: h }}>
        <ResponsiveContainer width="100%" height="100%">{renderInner()}</ResponsiveContainer>
      </div>
    </div>
  );
}

const ICON_BTN_STYLE: React.CSSProperties = {
  padding: "4px 10px", borderRadius: 7, fontSize: 13, cursor: "pointer",
  border: "1px solid #ffffff18", background: "#ffffff08", color: "#aaa",
};

// ── Builder Modal ─────────────────────────────────────────────────────────────
function BuilderModal({
  cols, data, multiCols, genCol, initial, onSave, onClose,
}: {
  cols: string[];
  data: Row[];
  multiCols: string[];
  genCol: string;
  initial?: ChartCfg;
  onSave: (cfg: Omit<ChartCfg, "id">) => void;
  onClose: () => void;
}) {
  const [type,       setType]       = useState<ChartType>(initial?.type ?? "horizontalBar");
  const [mode,       setMode]       = useState<DataMode>(initial?.mode ?? "cat");
  const [col,        setCol]        = useState(initial?.col ?? cols[0] ?? "");
  const [title,      setTitle]      = useState(initial?.title ?? "");
  const [filterCol,  setFilterCol]  = useState(initial?.filterCol ?? genCol);
  const [filterVals, setFilterVals] = useState<string[]>(initial?.filterVals ?? []);

  const filterOptions = filterCol ? allVals(data, filterCol) : [];
  const preview = useMemo(() => {
    if (!col) return "";
    const entries = mode === "multi" ? countMulti(data, col) : countCol(data, col);
    return entries.slice(0, 5).map(([k, v]) => `${k} (${v})`).join("  ·  ");
  }, [col, mode, data]);

  const toggleFVal = (v: string) =>
    setFilterVals((p) => p.includes(v) ? p.filter((x) => x !== v) : [...p, v]);

  const CHART_TYPES: { t: ChartType; l: string }[] = [
    { t: "bar", l: "Bar" }, { t: "horizontalBar", l: "H.Bar" },
    { t: "line", l: "Line" }, { t: "area", l: "Area" },
    { t: "pie", l: "Pie" }, { t: "radar", l: "Radar" },
  ];

  return (
    <div
      onClick={(e) => e.target === e.currentTarget && onClose()}
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.78)", zIndex: 999, display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "1.5rem", overflowY: "auto" }}
    >
      <div style={{ background: "#0f0f22", border: "1px solid #ffffff18", borderRadius: 18, padding: "1.35rem", width: "100%", maxWidth: 540 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
          <h3 style={{ margin: 0, color: "#fff", fontSize: 16, fontWeight: 700 }}>
            {initial ? "✏️ แก้ไขกราฟ" : "🎂 เพิ่มกราฟ"}
          </h3>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#aaa", fontSize: 20, cursor: "pointer" }}>✕</button>
        </div>

        {/* Chart type */}
        <FieldLabel>ประเภทกราฟ</FieldLabel>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 6, marginBottom: "1rem" }}>
          {CHART_TYPES.map(({ t, l }) => (
            <button key={t} onClick={() => setType(t)}
              style={{ padding: "8px 4px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer", border: "1px solid", transition: "all .15s",
                background: type === t ? "#FF6B9D22" : "#ffffff08",
                borderColor: type === t ? "#FF6B9D" : "#ffffff15",
                color: type === t ? "#FF6B9D" : "#888" }}>
              {l}
            </button>
          ))}
        </div>

        {/* Title */}
        <FieldLabel>ชื่อกราฟ</FieldLabel>
        <input value={title} onChange={(e) => setTitle(e.target.value)}
          placeholder="เช่น สิ่งที่ให้ความสำคัญเวลาซื้อเค้ก"
          style={INPUT_STYLE} />

        {/* Mode */}
        <FieldLabel>รูปแบบข้อมูล</FieldLabel>
        <div style={{ display: "flex", gap: 8, marginBottom: "1rem" }}>
          {(["cat", "multi"] as DataMode[]).map((m) => (
            <button key={m} onClick={() => setMode(m)}
              style={{ flex: 1, padding: "8px", borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: "pointer", border: "1px solid", transition: "all .15s",
                background: mode === m ? "#5F27CD22" : "#ffffff08",
                borderColor: mode === m ? "#5F27CD" : "#ffffff15",
                color: mode === m ? "#B47FEB" : "#888" }}>
              {m === "cat" ? "📊 หมวดหมู่ (นับจำนวน)" : "📋 Multi-select (คั่น ,)"}
            </button>
          ))}
        </div>

        {/* Column picker */}
        <FieldLabel>คอลัมน์ข้อมูล {mode === "multi" && <span style={{ color: "#B47FEB", fontWeight: 400 }}>— ตัวที่ตรวจพบว่า multi-select จะมี ★</span>}</FieldLabel>
        <select value={col} onChange={(e) => setCol(e.target.value)} style={INPUT_STYLE}>
          {cols.map((c) => (
            <option key={c} value={c} style={{ background: "#0f0f22" }}>
              {multiCols.includes(c) ? "★ " : ""}{shorten(c)}
            </option>
          ))}
        </select>
        {preview && (
          <p style={{ fontSize: 11, color: "#666", margin: "-8px 0 1rem", padding: "6px 10px", background: "#ffffff06", borderRadius: 8 }}>
            {preview}
          </p>
        )}

        {/* Filter */}
        <div style={{ background: "#ffffff06", border: "1px solid #FF9F4333", borderRadius: 12, padding: "12px", marginBottom: "1rem" }}>
          <p style={{ fontSize: 11, color: "#FF9F43", textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 8px", fontWeight: 700 }}>
            🎯 Filter เฉพาะกราฟนี้
          </p>
          <select value={filterCol} onChange={(e) => { setFilterCol(e.target.value); setFilterVals([]); }} style={{ ...INPUT_STYLE, marginBottom: 8 }}>
            <option value="">— ไม่ใช้ Filter —</option>
            {cols.map((c) => <option key={c} value={c} style={{ background: "#0f0f22" }}>{shorten(c)}</option>)}
          </select>
          {filterCol && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <span style={{ fontSize: 11, color: "#888" }}>เลือกค่า ({filterOptions.length} ตัวเลือก)</span>
                <div style={{ display: "flex", gap: 4 }}>
                  <MiniBtn onClick={() => setFilterVals([])}>ทั้งหมด</MiniBtn>
                  <MiniBtn onClick={() => setFilterVals([...filterOptions])}>เลือกทุกตัว</MiniBtn>
                </div>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {filterOptions.map((v, i) => (
                  <button key={v} onClick={() => toggleFVal(v)}
                    style={{ padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 600, cursor: "pointer", border: "1px solid", transition: "all .15s",
                      background: filterVals.includes(v) ? COLORS[i % COLORS.length] + "33" : "#ffffff08",
                      borderColor: filterVals.includes(v) ? COLORS[i % COLORS.length] : "#ffffff15",
                      color: filterVals.includes(v) ? "#fff" : "#888" }}>
                    {v}
                  </button>
                ))}
              </div>
              <p style={{ fontSize: 11, color: "#555", marginTop: 6 }}>
                {filterVals.length === 0 ? "ไม่เลือก = แสดงทั้งหมด" : `แสดงเฉพาะ: ${filterVals.join(", ")}`}
              </p>
            </>
          )}
        </div>

        <button
          onClick={() => {
            if (!col) { alert("กรุณาเลือกคอลัมน์"); return; }
            onSave({ type, mode, col, title: title || shorten(col), filterCol, filterVals });
            onClose();
          }}
          style={{ width: "100%", padding: "12px", borderRadius: 10, fontSize: 15, fontWeight: 700, cursor: "pointer", background: "linear-gradient(135deg,#FF6B9D,#FF9F43)", border: "none", color: "#fff" }}>
          ✦ {initial ? "บันทึกการแก้ไข" : "สร้างกราฟ"}
        </button>
      </div>
    </div>
  );
}

const FieldLabel = ({ children }: { children: React.ReactNode }) => (
  <p style={{ fontSize: 11, color: "#888", textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 6px", fontWeight: 600 }}>{children}</p>
);
const MiniBtn = ({ children, onClick }: { children: React.ReactNode; onClick: () => void }) => (
  <button onClick={onClick} style={{ padding: "2px 8px", borderRadius: 5, fontSize: 11, cursor: "pointer", border: "1px solid #ffffff15", background: "#ffffff08", color: "#aaa" }}>
    {children}
  </button>
);
const INPUT_STYLE: React.CSSProperties = {
  width: "100%", background: "#0d0d1a", border: "1px solid #ffffff18", borderRadius: 8,
  color: "#e0e0e0", padding: "8px 12px", fontSize: 13, outline: "none",
  marginBottom: "1rem", boxSizing: "border-box",
};

// ── Upload Zone ───────────────────────────────────────────────────────────────
function UploadZone({ onLoad }: { onLoad: (rows: Row[], fname: string) => void }) {
  const [dragging, setDragging] = useState(false);
  const [loading,  setLoading]  = useState(false);
  const [error,    setError]    = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const handle = useCallback(async (file: File) => {
    setLoading(true); setError("");
    try {
      const rows = await parseFile(file);
      if (!rows.length) throw new Error("ไฟล์ว่างเปล่า หรือไม่มีข้อมูล");
      onLoad(rows, file.name);
    } catch (e) {
      setError(e instanceof Error ? e.message : "โหลดไฟล์ไม่ได้");
    } finally { setLoading(false); }
  }, [onLoad]);

  return (
    <div style={{ padding: "3rem 1rem" }}>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handle(f); }}
        onClick={() => inputRef.current?.click()}
        style={{
          border: `2px dashed ${dragging ? "#FF6B9D" : "#ffffff25"}`,
          borderRadius: 20, padding: "3.5rem 2rem", textAlign: "center", cursor: "pointer",
          background: dragging ? "rgba(255,107,157,.07)" : "rgba(255,255,255,.02)",
          transition: "all .2s",
        }}
      >
        <div style={{ fontSize: 48, marginBottom: 12 }}>{loading ? "⏳" : "📂"}</div>
        <p style={{ margin: 0, fontWeight: 700, fontSize: 17, color: "#e0e0e0" }}>
          {loading ? "กำลังโหลด..." : "วางไฟล์ที่นี่ หรือคลิกเพื่อเลือก"}
        </p>
        <p style={{ margin: "6px 0 0", fontSize: 13, color: "#666" }}>
          รองรับ .csv · .xlsx · .xls (รองรับ Thai UTF-8 และ TIS-620)
        </p>
        {error && (
          <p style={{ margin: "12px auto 0", color: "#FF6B6B", fontSize: 13, background: "#FF6B6B15", borderRadius: 8, padding: "6px 14px", display: "inline-block" }}>
            ⚠️ {error}
          </p>
        )}
      </div>
      <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (f) handle(f); }} />
    </div>
  );
}

// ── Stat Card ─────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, color = "#54A0FF" }: { label: string; value: string; sub: string; color?: string }) {
  return (
    <div style={{ background: `${color}12`, border: `1px solid ${color}33`, borderRadius: 12, padding: "12px 14px" }}>
      <p style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: ".06em", color: "#888", margin: "0 0 4px", fontWeight: 600 }}>{label}</p>
      <p style={{ fontSize: 20, fontWeight: 800, color: "#fff", margin: 0 }}>{value}</p>
      <p style={{ fontSize: 11, color: "#666", margin: "2px 0 0" }}>{sub}</p>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function CrackTheCakeDashboard() {
  const [data,        setData]        = useState<Row[]>([]);
  const [cols,        setCols]        = useState<string[]>([]);
  const [fileName,    setFileName]    = useState("");
  const [genCol,      setGenCol]      = useState("");
  const [multiCols,   setMultiCols]   = useState<string[]>([]);
  const [gFilter,     setGFilter]     = useState<string[]>([]);
  const [charts,      setCharts]      = useState<ChartCfg[]>([]);
  const [showBuilder, setShowBuilder] = useState(false);
  const [editChart,   setEditChart]   = useState<ChartCfg | null>(null);
  const idRef = useRef(1);

  const genVals   = useMemo(() => genCol ? allVals(data, genCol) : [], [data, genCol]);
  const genCounts = useMemo(() => genCol ? countCol(data, genCol) : [], [data, genCol]);

  const handleLoad = useCallback((rows: Row[], fname: string) => {
    const c = Object.keys(rows[0]);
    setData(rows);
    setCols(c);
    setFileName(fname);
    setCharts([]);
    setGFilter([]);
    const gCol = detectGenCol(c, rows);
    const mCols = detectMultiCols(c);
    setGenCol(gCol);
    setMultiCols(mCols);
  }, []);

  const toggleGen = useCallback((v: string) => {
    if (v === "ทั้งหมด") { setGFilter([]); return; }
    setGFilter((p) => p.includes(v) ? p.filter((x) => x !== v) : [...p, v]);
  }, []);

  const addChart = (cfg: Omit<ChartCfg, "id">) =>
    setCharts((p) => [...p, { ...cfg, id: idRef.current++ }]);

  const saveEdit = (cfg: Omit<ChartCfg, "id">) =>
    setCharts((p) => p.map((c) => c.id === editChart!.id ? { ...cfg, id: c.id } : c));

  // Auto-suggest default charts after load
  const loadDefaults = useCallback(() => {
    if (!cols.length) return;
    const suggestions: Omit<ChartCfg, "id">[] = [];
    // 1. Gen distribution if exists
    if (genCol) suggestions.push({ type: "pie", mode: "cat", col: genCol, title: `สัดส่วน ${genCol}`, filterCol: "", filterVals: [] });
    // 2. First multi-select col overall
    if (multiCols[0]) suggestions.push({ type: "horizontalBar", mode: "multi", col: multiCols[0], title: shorten(multiCols[0], 38) + " (ทั้งหมด)", filterCol: "", filterVals: [] });
    // 3. First multi-select filtered by each gen val
    if (multiCols[0] && genCol) {
      genVals.slice(0, 3).forEach((gv) =>
        suggestions.push({ type: "horizontalBar", mode: "multi", col: multiCols[0], title: `${shorten(multiCols[0], 28)} — ${gv}`, filterCol: genCol, filterVals: [gv] })
      );
    }
    // 4. Other cat columns
    const catCols = cols.filter((c) => c !== genCol && !multiCols.includes(c) && allVals(data, c).length <= 20).slice(0, 3);
    catCols.forEach((c) => suggestions.push({ type: "horizontalBar", mode: "cat", col: c, title: shorten(c, 38), filterCol: "", filterVals: [] }));

    idRef.current = 1;
    setCharts(suggestions.map((s, i) => ({ ...s, id: i + 1 })));
    idRef.current = suggestions.length + 1;
  }, [cols, genCol, multiCols, genVals, data]);

  return (
    <div style={{ fontFamily: "system-ui,sans-serif", background: "#0d0d1a", minHeight: "100vh", color: "#e0e0e0" }}>
      {/* bg blobs */}
      <div style={{ position: "fixed", top: 0, left: "20%", width: 400, height: 400, borderRadius: "50%", background: "radial-gradient(circle,rgba(255,107,157,.09),transparent 70%)", filter: "blur(40px)", pointerEvents: "none" }} />
      <div style={{ position: "fixed", bottom: "20%", right: 0, width: 320, height: 320, borderRadius: "50%", background: "radial-gradient(circle,rgba(255,159,67,.07),transparent 70%)", filter: "blur(40px)", pointerEvents: "none" }} />

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "1.5rem 1rem", position: "relative" }}>

        {/* Header */}
        <div style={{ marginBottom: "1.5rem" }}>
          <h1 style={{ fontSize: 30, fontWeight: 900, margin: 0, background: "linear-gradient(135deg,#FF6B9D,#FF9F43,#FECA57)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
            🎂 Crack the Cake Dashboard
          </h1>
          <p style={{ color: "#555", fontSize: 13, margin: "4px 0 0" }}>
            อัปโหลดไฟล์ข้อมูลของคุณ · รองรับ CSV / XLSX · Multi-select · Filter ตาม Generation
          </p>
        </div>

        {/* Upload zone — show when no data */}
        {!data.length ? (
          <UploadZone onLoad={handleLoad} />
        ) : (
          <>
            {/* File bar */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#12122a", border: "1px solid #ffffff10", borderRadius: 12, padding: "10px 16px", marginBottom: "1rem", flexWrap: "wrap", gap: 8 }}>
              <div>
                <span style={{ fontWeight: 600, fontSize: 14, color: "#fff" }}>📄 {fileName}</span>
                <span style={{ fontSize: 12, color: "#555", marginLeft: 10 }}>
                  {data.length.toLocaleString()} แถว · {cols.length} คอลัมน์
                  {genCol && <span style={{ color: "#FF9F43", marginLeft: 8 }}>· Gen column: <b>{genCol}</b></span>}
                  {multiCols.length > 0 && <span style={{ color: "#B47FEB", marginLeft: 8 }}>· Multi-select: {multiCols.length} คอลัมน์</span>}
                </span>
              </div>
              <button
                onClick={() => { setData([]); setCols([]); setCharts([]); setFileName(""); setGenCol(""); setMultiCols([]); setGFilter([]); }}
                style={{ padding: "5px 14px", borderRadius: 8, fontSize: 12, cursor: "pointer", border: "1px solid #ffffff18", background: "#ffffff08", color: "#aaa" }}>
                เปลี่ยนไฟล์
              </button>
            </div>

            {/* Stat cards */}
            {genCounts.length > 0 && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 8, marginBottom: "1rem" }}>
                <StatCard label="ทั้งหมด" value={String(data.length)} sub="แถว" />
                {genCounts.map(([g, n], i) => (
                  <StatCard key={g} label={g} value={String(n)} sub={`${((n / data.length) * 100).toFixed(1)}%`} color={COLORS[i]} />
                ))}
              </div>
            )}

            {/* Gen filter bar */}
            {genCol && (
              <div style={{ background: "#12122a", border: "1px solid #ffffff10", borderRadius: 14, padding: "12px 16px", marginBottom: "1.25rem" }}>
                <p style={{ fontSize: 11, color: "#888", textTransform: "uppercase", letterSpacing: ".06em", margin: "0 0 8px", fontWeight: 600 }}>
                  Filter ตาม {genCol} (ทุกกราฟ)
                </p>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {["ทั้งหมด", ...genVals].map((v, i) => {
                    const isAll = v === "ทั้งหมด";
                    const active = isAll ? gFilter.length === 0 : gFilter.includes(v);
                    const col = isAll ? "#54A0FF" : COLORS[i - 1];
                    return (
                      <button key={v} onClick={() => toggleGen(v)}
                        style={{ padding: "5px 14px", borderRadius: 20, fontSize: 13, fontWeight: 600, cursor: "pointer", border: "1px solid", transition: "all .15s",
                          background: active ? col : "#ffffff0a",
                          borderColor: active ? col : "#ffffff18",
                          color: active ? "#fff" : "#aaa",
                          boxShadow: active ? `0 0 10px ${col}44` : "none" }}>
                        {v}
                      </button>
                    );
                  })}
                </div>
                {gFilter.length > 0 && (
                  <p style={{ fontSize: 11, color: "#FF9F43", margin: "8px 0 0" }}>🎯 Filter: {gFilter.join(", ")}</p>
                )}
              </div>
            )}

            {/* Toolbar */}
            <div style={{ display: "flex", gap: 10, marginBottom: "1.25rem", flexWrap: "wrap", alignItems: "center" }}>
              <button onClick={() => { setEditChart(null); setShowBuilder(true); }}
                style={{ padding: "9px 18px", borderRadius: 10, fontWeight: 700, fontSize: 14, cursor: "pointer", background: "linear-gradient(135deg,#FF6B9D,#FF9F43)", border: "none", color: "#fff" }}>
                + เพิ่มกราฟ
              </button>
              <button onClick={loadDefaults}
                style={{ padding: "9px 18px", borderRadius: 10, fontWeight: 600, fontSize: 13, cursor: "pointer", background: "#ffffff0a", border: "1px solid #ffffff18", color: "#ccc" }}>
                ⚡ โหลดกราฟแนะนำ
              </button>
              {charts.length > 0 && (
                <button onClick={() => setCharts([])}
                  style={{ padding: "9px 18px", borderRadius: 10, fontWeight: 600, fontSize: 13, cursor: "pointer", background: "#ff6b6b0a", border: "1px solid #ff6b6b33", color: "#ff6b6b" }}>
                  🗑 ล้างทั้งหมด
                </button>
              )}
            </div>

            {/* Charts */}
            {charts.length === 0 ? (
              <div style={{ textAlign: "center", padding: "4rem 1rem", color: "#444", fontSize: 14 }}>
                <div style={{ fontSize: 40, marginBottom: 8 }}>📊</div>
                กด "+ เพิ่มกราฟ" หรือ "⚡ โหลดกราฟแนะนำ" เพื่อเริ่มต้น
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
                {charts.map((c) => (
                  <ChartCard
                    key={c.id}
                    cfg={c}
                    data={data}
                    genCol={genCol}
                    gFilter={gFilter}
                    onEdit={() => { setEditChart(c); setShowBuilder(true); }}
                    onRemove={() => setCharts((p) => p.filter((x) => x.id !== c.id))}
                  />
                ))}
              </div>
            )}
          </>
        )}
      </div>

      {/* Builder modal */}
      {showBuilder && (
        <BuilderModal
          cols={cols}
          data={data}
          multiCols={multiCols}
          genCol={genCol}
          initial={editChart ?? undefined}
          onSave={editChart ? saveEdit : addChart}
          onClose={() => { setShowBuilder(false); setEditChart(null); }}
        />
      )}
    </div>
  );
}