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

// ─── Types ────────────────────────────────────────────────────────────────────
type Row       = Record<string, string>;
type ChartType = "bar" | "horizontalBar" | "line" | "area" | "pie" | "radar";
type DataMode  = "cat" | "multi";

interface ChartCfg {
  id: number;
  type: ChartType;
  mode: DataMode;
  col: string;
  title: string;
  filterCol: string;
  filterVals: string[];
}

// ─── Dark design tokens ───────────────────────────────────────────────────────
const D = {
  bg:         "#0C0C0E",
  surface:    "#111114",
  surfaceHov: "#1C1C1F",
  border:     "#2A2A2E",
  borderMd:   "#3F3F46",
  text:       "#E8E8EA",
  textSub:    "#adadb8",
  textMuted:  "#adadb8",
  accent:     "#E8E8EA",
  positive:   "#4ADE80",
  positiveBg: "#0D2B1E",
  negative:   "#F87171",
  negativeBg: "#2B0D0D",
  sidebarW:   220,
  topH:       52,
};

// chart colours — desaturated, professional on dark
const CC = ["#FFB7C5","#B2F2BB","#D0BFFF","#A5D8FF","#FFF4B2","#FFD8A8","#FCC2D7","#96F2F7"];
// accent colours for pie/donut only
const PC = ["#FFB7C5","#B2F2BB","#D0BFFF","#A5D8FF","#FFF4B2","#FFD8A8","#FCC2D7","#96F2F7"];

// ─── File parser ──────────────────────────────────────────────────────────────
function parseFile(file: File): Promise<Row[]> {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = (e) => {
        Papa.parse(e.target?.result as string, {
          header: true, skipEmptyLines: true, dynamicTyping: false,
          transformHeader: (h) => h.trim(), transform: (v) => v.trim(),
          complete: (r) => resolve(r.data as Row[]), error: reject,
        });
      };
      reader.onerror = reject;
      reader.readAsText(file, "UTF-8");
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const wb   = XLSX.read(e.target?.result, { type: "binary" });
        const json = XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        resolve(json.map((r) =>
          Object.fromEntries(Object.entries(r).map(([k, v]) => [k.trim(), String(v).trim()]))
        ));
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    } else reject(new Error("รองรับเฉพาะ .csv .xlsx .xls"));
  });
}

// ─── Auto-detect helpers ──────────────────────────────────────────────────────
const GEN_VALS  = ["gen z","gen alpha","gen x","gen y","millennial","boomer"];
const MULTI_KW  = ["select all","up to","you can answer","select all that apply","choose up"];

function detectGenCol(cols: string[], data: Row[]): string {
  return (
    cols.find((c) => /gen(eration)?|age.?group|cohort|ช่วงอายุ/i.test(c)) ??
    cols.find((c) => data.slice(0,20).some((r) =>
      GEN_VALS.some((g) => String(r[c]??"").toLowerCase().includes(g))
    )) ?? ""
  );
}
const detectMultiCols = (cols: string[]) =>
  cols.filter((c) => MULTI_KW.some((k) => c.toLowerCase().includes(k)));

// ─── Data helpers ─────────────────────────────────────────────────────────────
const shorten  = (s: string, n = 46) => s.length > n ? s.slice(0, n) + "…" : s;
const allVals  = (data: Row[], col: string): string[] =>
  [...new Set(data.map((r) => r[col] ?? ""))].filter(Boolean).sort();

function countCol(data: Row[], col: string): [string, number][] {
  const c: Record<string, number> = {};
  data.forEach((r) => { const v = r[col] ?? ""; if (v) c[v] = (c[v] ?? 0) + 1; });
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}
function countMulti(data: Row[], col: string): [string, number][] {
  const c: Record<string, number> = {};
  data.forEach((r) =>
    (r[col] ?? "").split(",").map((s) => s.trim()).filter(Boolean)
      .forEach((v) => { c[v] = (c[v] ?? 0) + 1; })
  );
  return Object.entries(c).sort((a, b) => b[1] - a[1]);
}
function applyFilters(data: Row[], genCol: string, gf: string[], lcol: string, lv: string[]): Row[] {
  let d = data;
  if (genCol && gf.length) d = d.filter((r) => gf.includes(r[genCol] ?? ""));
  if (lcol  && lv.length)  d = d.filter((r) => lv.includes(r[lcol]   ?? ""));
  return d;
}

// ─── Shared tooltip ───────────────────────────────────────────────────────────
function DarkTip({ active, payload, label, total }:
  { active?: boolean; payload?: {value:number}[]; label?: string; total: number }) {
  if (!active || !payload?.length) return null;
  const v = payload[0].value;
  return (
    <div style={{ background: "#1C1C1F", border: `1px solid ${D.border}`, borderRadius: 8,
      padding: "8px 12px", color: D.text, fontSize: 12, boxShadow: "0 8px 24px rgba(0,0,0,.5)" }}>
      <p style={{ margin: "0 0 2px", fontSize: 11, color: D.textSub }}>{label}</p>
      <p style={{ margin: 0, fontWeight: 500 }}>
        {v}&nbsp;
        <span style={{ color: D.textMuted, fontWeight: 400 }}>({((v / (total||1)) * 100).toFixed(1)}%)</span>
      </p>
    </div>
  );
}

// ─── Icon set (inline SVG, zero deps) ────────────────────────────────────────
const Ico = {
  grid:    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>,
  chart:   <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6"  y1="20" x2="6"  y2="14"/></svg>,
  upload:  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0018 9h-1.26A8 8 0 103 16.3"/></svg>,
  plus:    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  bolt:    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>,
  edit:    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4z"/></svg>,
  trash:   <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6M9 6V4h6v2"/></svg>,
  x:       <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  filter:  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>,
  logo:    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"><circle cx="12" cy="12" r="10"/><path d="M8 12L12 8L16 12"/><path d="M12 8v8"/></svg>,
};

// ─── Primitives ───────────────────────────────────────────────────────────────
function NavItem({ icon, label, active, onClick }:
  { icon: React.ReactNode; label: string; active?: boolean; onClick?: () => void }) {
  return (
    <button onClick={onClick} style={{
      display: "flex", alignItems: "center", gap: 9, width: "100%",
      padding: "7px 10px", borderRadius: 6, border: "none", cursor: "pointer",
      background: active ? D.surfaceHov : "transparent",
      color: active ? D.text : D.textSub,
      fontSize: 12, fontWeight: 400, textAlign: "left", transition: "all .12s",
    }}>
      {icon} {label}
    </button>
  );
}

function Btn({ children, onClick, variant = "ghost", style: s }: {
  children: React.ReactNode; onClick?: () => void;
  variant?: "ghost" | "solid" | "danger"; style?: React.CSSProperties;
}) {
  const base: React.CSSProperties = {
    display: "flex", alignItems: "center", gap: 6,
    padding: "6px 12px", borderRadius: 6, fontSize: 12, fontWeight: 500,
    cursor: "pointer", border: "none", transition: "all .12s", ...s,
  };
  const v: Record<string, React.CSSProperties> = {
    ghost:  { background: D.surfaceHov, color: D.textSub,  border: `0.5px solid ${D.border}` },
    solid:  { background: D.text,       color: D.bg,        border: "none" },
    danger: { background: D.negativeBg, color: D.negative,  border: `0.5px solid #3F1010` },
  };
  return <button onClick={onClick} style={{ ...base, ...v[variant] }}>{children}</button>;
}

function IBtn({ children, onClick, danger }: { children: React.ReactNode; onClick: () => void; danger?: boolean }) {
  return (
    <button onClick={onClick} style={{
      width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center",
      borderRadius: 5, cursor: "pointer",
      border: `0.5px solid ${danger ? "#3F1010" : D.border}`,
      background: danger ? D.negativeBg : D.surfaceHov,
      color: danger ? D.negative : D.textMuted, flexShrink: 0,
    }}>{children}</button>
  );
}

// ─── Stat card ────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, up }:
  { label: string; value: string; sub: string; up?: boolean }) {
  return (
    <div style={{ background: D.surface, border: `0.5px solid ${D.border}`, borderRadius: 8, padding: "14px 16px" }}>
      <p style={{ margin: 0, fontSize: 10, textTransform: "uppercase", letterSpacing: ".07em", color: D.textMuted, fontWeight: 500 }}>{label}</p>
      <p style={{ margin: "6px 0 0", fontSize: 22, fontWeight: 500, color: D.text, letterSpacing: "-0.5px" }}>{value}</p>
      <p style={{ margin: "4px 0 0", display: "inline-flex", alignItems: "center", gap: 3, fontSize: 10, padding: "2px 6px", borderRadius: 4,
        background: up === undefined ? "transparent" : up ? D.positiveBg : D.negativeBg,
        color: up === undefined ? D.textMuted : up ? D.positive : D.negative }}>
        {up !== undefined && (up ? "↑" : "↓")} {sub}
      </p>
    </div>
  );
}

// ─── Chart card ───────────────────────────────────────────────────────────────
function ChartCard({ cfg, data, genCol, gFilter, onEdit, onRemove }:
  { cfg: ChartCfg; data: Row[]; genCol: string; gFilter: string[];
    onEdit: () => void; onRemove: () => void }) {

  const filtered = useMemo(
    () => applyFilters(data, genCol, gFilter, cfg.filterCol, cfg.filterVals),
    [data, genCol, gFilter, cfg]
  );
  const entries = useMemo(
    () => cfg.mode === "multi" ? countMulti(filtered, cfg.col) : countCol(filtered, cfg.col),
    [filtered, cfg]
  );
  const total = filtered.length;
  const cd    = entries.map(([name, value]) => ({ name, value }));

  const fLabel = [
    gFilter.length ? gFilter.join(", ") : "",
    cfg.filterCol && cfg.filterVals.length ? `${shorten(cfg.filterCol, 20)}: ${cfg.filterVals.join(", ")}` : "",
  ].filter(Boolean).join(" · ") || "ทั้งหมด";

  const axTick = { fill: D.textMuted, fontSize: 10 };

  const inner = () => {
    const tip = <DarkTip total={total} />;
    if (cfg.type === "pie")
      return (
        <PieChart>
          <Pie data={cd} dataKey="value" cx="50%" cy="50%" innerRadius={55} outerRadius={100} paddingAngle={1.5} stroke="none">
            {cd.map((_, i) => <Cell key={i} fill={PC[i % PC.length]} />)}
          </Pie>
          <Tooltip content={tip} />
          <Legend iconType="circle" iconSize={7} wrapperStyle={{ fontSize: 11, color: D.textSub }}
            formatter={(v) => <span style={{ color: D.textSub }}>{shorten(v, 22)}</span>} />
        </PieChart>
      );
    if (cfg.type === "radar")
      return (
        <RadarChart data={cd.slice(0,10)}>
          <PolarGrid stroke={D.border} />
          <PolarAngleAxis dataKey="name" tick={axTick} />
          <PolarRadiusAxis tick={{ fill: D.textMuted, fontSize: 9 }} />
          <Radar dataKey="value" stroke={D.text} fill={D.text} fillOpacity={0.1} />
          <Tooltip content={tip} />
        </RadarChart>
      );

    const isH = cfg.type === "horizontalBar";
    const isL = cfg.type === "line";
    const isA = cfg.type === "area";
    const margin = isH
      ? { top: 8, right: 36, left: 8, bottom: 8 }
      : { top: 8, right: 12, left: 0, bottom: 60 };
    const xAxis = isH
      ? <XAxis type="number" tick={axTick} axisLine={false} tickLine={false} />
      : <XAxis dataKey="name" tick={{ fill: D.textMuted, fontSize: 10 }} angle={-35}
               textAnchor="end" interval={0} axisLine={false} tickLine={false} />;
    const yAxis = isH
      ? <YAxis type="category" dataKey="name" tick={{ fill: D.textMuted, fontSize: 10 }}
               width={148} axisLine={false} tickLine={false} />
      : <YAxis tick={axTick} axisLine={false} tickLine={false} />;
    const grid  = <CartesianGrid stroke={D.border} strokeDasharray="0" opacity={0.5} />;

    if (isL) return (
      <LineChart data={cd} margin={margin}>
        {grid}{xAxis}{yAxis}
        <Tooltip content={tip} />
        <Line type="monotone" dataKey="value" stroke={D.text} strokeWidth={1.8}
          dot={{ fill: D.text, r: 3, strokeWidth: 0 }} activeDot={{ r: 5, strokeWidth: 0 }} />
      </LineChart>
    );
    if (isA) return (
      <AreaChart data={cd} margin={margin}>
        <defs>
          <linearGradient id={`ag${cfg.id}`} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%"   stopColor={D.text} stopOpacity={0.15} />
            <stop offset="100%" stopColor={D.text} stopOpacity={0} />
          </linearGradient>
        </defs>
        {grid}{xAxis}{yAxis}
        <Tooltip content={tip} />
        <Area type="monotone" dataKey="value" stroke={D.text} fill={`url(#ag${cfg.id})`} strokeWidth={1.8} dot={false} />
      </AreaChart>
    );
    return (
      <BarChart data={cd} layout={isH ? "vertical" : "horizontal"} margin={margin}>
        {grid}{xAxis}{yAxis}
        <Tooltip content={tip} />
        <Bar dataKey="value" radius={isH ? [0, 3, 3, 0] : [3, 3, 0, 0]} maxBarSize={28}>
          {cd.map((_, i) => <Cell key={i} fill={CC[i % CC.length]} />)}
        </Bar>
      </BarChart>
    );
  };

  const h = cfg.type === "horizontalBar" ? Math.max(260, cd.length * 32) : 280;

  return (
    <div style={{ background: D.surface, border: `0.5px solid ${D.border}`, borderRadius: 8, padding: "18px 20px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div>
          <p style={{ margin: 0, fontSize: 12, fontWeight: 500, color: D.text }}>{cfg.title}</p>
          <p style={{ margin: "3px 0 0", fontSize: 10, color: D.textMuted }}>{fLabel} · {total} รายการ</p>
        </div>
        <div style={{ display: "flex", gap: 4 }}>
          <IBtn onClick={onEdit}>{Ico.edit}</IBtn>
          <IBtn onClick={onRemove} danger>{Ico.trash}</IBtn>
        </div>
      </div>
      <div style={{ height: h }}>
        <ResponsiveContainer width="100%" height="100%">{inner()}</ResponsiveContainer>
      </div>
    </div>
  );
}

// ─── Builder modal ────────────────────────────────────────────────────────────
function BuilderModal({ cols, data, multiCols, genCol, initial, onSave, onClose }: {
  cols: string[]; data: Row[]; multiCols: string[]; genCol: string;
  initial?: ChartCfg;
  onSave: (c: Omit<ChartCfg,"id">) => void;
  onClose: () => void;
}) {
  const [type,       setType]       = useState<ChartType>(initial?.type ?? "horizontalBar");
  const [mode,       setMode]       = useState<DataMode>(initial?.mode ?? "cat");
  const [col,        setCol]        = useState(initial?.col ?? cols[0] ?? "");
  const [title,      setTitle]      = useState(initial?.title ?? "");
  const [filterCol,  setFilterCol]  = useState(initial?.filterCol ?? genCol);
  const [filterVals, setFilterVals] = useState<string[]>(initial?.filterVals ?? []);
  const fOpts = filterCol ? allVals(data, filterCol) : [];

  const preview = useMemo(() => {
    if (!col) return "";
    return (mode === "multi" ? countMulti(data, col) : countCol(data, col))
      .slice(0, 4).map(([k, v]) => `${k} (${v})`).join("  ·  ");
  }, [col, mode, data]);

  const toggleFV = (v: string) =>
    setFilterVals((p) => p.includes(v) ? p.filter((x) => x !== v) : [...p, v]);

  const TYPES: { t: ChartType; l: string }[] = [
    { t:"bar",l:"Bar" }, { t:"horizontalBar",l:"H.Bar" },
    { t:"line",l:"Line" }, { t:"area",l:"Area" },
    { t:"pie",l:"Pie" }, { t:"radar",l:"Radar" },
  ];

  const inp: React.CSSProperties = {
    width: "100%", background: D.bg, border: `0.5px solid ${D.border}`,
    borderRadius: 6, color: D.text, padding: "7px 10px", fontSize: 12,
    outline: "none", marginBottom: 14, boxSizing: "border-box",
  };

  return (
    <div onClick={(e) => e.target === e.currentTarget && onClose()}
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", zIndex: 999,
        display: "flex", alignItems: "flex-start", justifyContent: "center",
        padding: "2rem 1rem", overflowY: "auto" }}>
      <div style={{ background: D.surface, border: `0.5px solid ${D.border}`, borderRadius: 10,
        padding: "22px 24px", width: "100%", maxWidth: 500,
        boxShadow: "0 24px 64px rgba(0,0,0,.6)" }}>

        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <p style={{ margin: 0, fontSize: 14, fontWeight: 500, color: D.text }}>
            {initial ? "แก้ไขกราฟ" : "เพิ่มกราฟ"}
          </p>
          <IBtn onClick={onClose}>{Ico.x}</IBtn>
        </div>

        {/* Type */}
        <FL>ประเภทกราฟ</FL>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 5, marginBottom: 16 }}>
          {TYPES.map(({ t, l }) => (
            <button key={t} onClick={() => setType(t)} style={{
              padding: "7px 2px", borderRadius: 5, fontSize: 11, fontWeight: 500, cursor: "pointer",
              border: `0.5px solid ${type === t ? D.text : D.border}`,
              background: type === t ? D.text : D.bg,
              color: type === t ? D.bg : D.textSub, transition: "all .1s",
            }}>{l}</button>
          ))}
        </div>

        {/* Title */}
        <FL>ชื่อกราฟ</FL>
        <input value={title} onChange={(e) => setTitle(e.target.value)}
          placeholder="ชื่อที่จะแสดงบนการ์ด" style={inp} />

        {/* Mode */}
        <FL>รูปแบบข้อมูล</FL>
        <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
          {(["cat", "multi"] as DataMode[]).map((m) => (
            <button key={m} onClick={() => setMode(m)} style={{
              flex: 1, padding: "7px", borderRadius: 5, fontSize: 12, fontWeight: 500,
              cursor: "pointer", border: `0.5px solid ${mode === m ? D.text : D.border}`,
              background: mode === m ? D.text : D.bg,
              color: mode === m ? D.bg : D.textSub, transition: "all .1s",
            }}>
              {m === "cat" ? "หมวดหมู่ (นับจำนวน)" : "Multi-select (คั่น ,)"}
            </button>
          ))}
        </div>

        {/* Column */}
        <FL>คอลัมน์ข้อมูล</FL>
        <select value={col} onChange={(e) => setCol(e.target.value)}
          style={{ ...inp, colorScheme: "dark" }}>
          {cols.map((c) => (
            <option key={c} value={c} style={{ background: D.surface }}>
              {multiCols.includes(c) ? "★ " : ""}{shorten(c)}
            </option>
          ))}
        </select>
        {preview && (
          <p style={{ fontSize: 10, color: D.textMuted, margin: "-10px 0 14px",
            padding: "5px 10px", background: D.bg, borderRadius: 5 }}>
            {preview}
          </p>
        )}

        {/* Filter */}
        <div style={{ background: D.bg, border: `0.5px solid ${D.border}`, borderRadius: 7, padding: 14, marginBottom: 16 }}>
          <p style={{ margin: "0 0 8px", fontSize: 10, fontWeight: 500, color: D.textSub,
            textTransform: "uppercase", letterSpacing: ".07em" }}>Filter เฉพาะกราฟนี้</p>
          <select value={filterCol} onChange={(e) => { setFilterCol(e.target.value); setFilterVals([]); }}
            style={{ ...inp, marginBottom: 8, colorScheme: "dark" }}>
            <option value="">— ไม่ใช้ Filter —</option>
            {cols.map((c) => <option key={c} value={c} style={{ background: D.surface }}>{shorten(c)}</option>)}
          </select>
          {filterCol && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <span style={{ fontSize: 10, color: D.textMuted }}>{fOpts.length} ค่า</span>
                <div style={{ display: "flex", gap: 4 }}>
                  {[["ทั้งหมด", () => setFilterVals([])], ["ทุกตัว", () => setFilterVals([...fOpts])]].map(([l, fn]) => (
                    <button key={l as string} onClick={fn as () => void}
                      style={{ padding: "2px 8px", borderRadius: 4, fontSize: 10, cursor: "pointer",
                        border: `0.5px solid ${D.border}`, background: D.surfaceHov, color: D.textSub }}>
                      {l as string}
                    </button>
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                {fOpts.map((v) => {
                  const sel = filterVals.includes(v);
                  return (
                    <button key={v} onClick={() => toggleFV(v)} style={{
                      padding: "3px 10px", borderRadius: 4, fontSize: 11, fontWeight: 500, cursor: "pointer",
                      border: `0.5px solid ${sel ? D.text : D.border}`,
                      background: sel ? D.text : "transparent",
                      color: sel ? D.bg : D.textSub, transition: "all .1s",
                    }}>{v}</button>
                  );
                })}
              </div>
              <p style={{ fontSize: 10, color: D.textMuted, marginTop: 6 }}>
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
          style={{ width: "100%", padding: "10px", borderRadius: 7, fontSize: 13, fontWeight: 500,
            cursor: "pointer", background: D.text, border: "none", color: D.bg }}>
          {initial ? "บันทึก" : "สร้างกราฟ"}
        </button>
      </div>
    </div>
  );
}

function FL({ children }: { children: React.ReactNode }) {
  return (
    <p style={{ margin: "0 0 5px", fontSize: 10, fontWeight: 500, color: D.textSub,
      textTransform: "uppercase", letterSpacing: ".07em" }}>{children}</p>
  );
}

// ─── Upload zone ──────────────────────────────────────────────────────────────
function UploadZone({ onLoad }: { onLoad: (rows: Row[], fname: string) => void }) {
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const [err,  setErr]  = useState("");
  const ref = useRef<HTMLInputElement>(null);

  const handle = useCallback(async (f: File) => {
    setBusy(true); setErr("");
    try {
      const rows = await parseFile(f);
      if (!rows.length) throw new Error("ไฟล์ไม่มีข้อมูล");
      onLoad(rows, f.name);
    } catch (e) { setErr(e instanceof Error ? e.message : "โหลดไม่ได้"); }
    finally { setBusy(false); }
  }, [onLoad]);

  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", flex: 1 }}>
      <div style={{ maxWidth: 420, width: "100%", textAlign: "center" }}>
        <div
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); const f = e.dataTransfer.files[0]; if (f) handle(f); }}
          onClick={() => ref.current?.click()}
          style={{
            border: `1px dashed ${drag ? D.text : D.border}`,
            borderRadius: 10, padding: "40px 24px", cursor: "pointer",
            background: drag ? D.surfaceHov : D.surface, transition: "all .15s",
          }}>
          <div style={{ width: 42, height: 42, borderRadius: 9, background: D.surfaceHov,
            border: `0.5px solid ${D.border}`, display: "flex", alignItems: "center",
            justifyContent: "center", margin: "0 auto 14px", color: D.textSub }}>
            {Ico.upload}
          </div>
          <p style={{ margin: "0 0 4px", fontSize: 13, fontWeight: 500, color: D.text }}>
            {busy ? "กำลังโหลด..." : "อัปโหลดไฟล์ข้อมูล"}
          </p>
          <p style={{ margin: 0, fontSize: 11, color: D.textMuted }}>
            CSV · XLSX · XLS — ลากวางหรือคลิก
          </p>
          {err && <p style={{ margin: "10px 0 0", fontSize: 11, color: D.negative }}>{err}</p>}
        </div>
        <input ref={ref} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
          onChange={(e) => { const f = e.target.files?.[0]; if (f) handle(f); }} />
      </div>
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────────
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
  const [activeNav,   setActiveNav]   = useState<"dashboard"|"analytics"|"data">("dashboard");
  const idRef = useRef(1);

  const genVals   = useMemo(() => genCol ? allVals(data, genCol)    : [],  [data, genCol]);
  const genCounts = useMemo(() => genCol ? countCol(data, genCol)   : [],  [data, genCol]);
  const hasData   = data.length > 0;

  const handleLoad = useCallback((rows: Row[], fname: string) => {
    const c = Object.keys(rows[0]);
    setData(rows); setCols(c); setFileName(fname);
    setCharts([]); setGFilter([]);
    setGenCol(detectGenCol(c, rows));
    setMultiCols(detectMultiCols(c));
  }, []);

  const toggleGen = useCallback((v: string) => {
    if (v === "ทั้งหมด") { setGFilter([]); return; }
    setGFilter((p) => p.includes(v) ? p.filter((x) => x !== v) : [...p, v]);
  }, []);

  const addChart  = (cfg: Omit<ChartCfg,"id">) =>
    setCharts((p) => [...p, { ...cfg, id: idRef.current++ }]);
  const saveEdit  = (cfg: Omit<ChartCfg,"id">) =>
    setCharts((p) => p.map((c) => c.id === editChart!.id ? { ...cfg, id: c.id } : c));

  const loadDefaults = useCallback(() => {
    if (!cols.length) return;
    const list: Omit<ChartCfg,"id">[] = [];
    if (genCol)
      list.push({ type:"pie", mode:"cat", col:genCol,
        title:`สัดส่วน ${genCol}`, filterCol:"", filterVals:[] });
    if (multiCols[0]) {
      list.push({ type:"horizontalBar", mode:"multi", col:multiCols[0],
        title: shorten(multiCols[0],36)+" (ทั้งหมด)", filterCol:"", filterVals:[] });
      genVals.slice(0,3).forEach((gv) =>
        list.push({ type:"horizontalBar", mode:"multi", col:multiCols[0],
          title:`${shorten(multiCols[0],26)} — ${gv}`, filterCol:genCol, filterVals:[gv] })
      );
    }
    cols.filter((c) => c!==genCol && !multiCols.includes(c) && allVals(data,c).length<=20)
      .slice(0,3)
      .forEach((c) => list.push({ type:"horizontalBar", mode:"cat", col:c,
        title:shorten(c,36), filterCol:"", filterVals:[] }));
    idRef.current = 1;
    setCharts(list.map((s, i) => ({ ...s, id: i+1 })));
    idRef.current = list.length + 1;
  }, [cols, genCol, multiCols, genVals, data]);

  const reset = useCallback(() => {
    setData([]); setCols([]); setCharts([]); setFileName("");
    setGenCol(""); setMultiCols([]); setGFilter([]);
  }, []);

  // ── Sidebar ───────────────────────────────────────────────────────────────
  const NAV = [
    { key: "dashboard" as const, icon: Ico.grid,   label: "Dashboard"  },
    { key: "analytics" as const, icon: Ico.chart,  label: "Analytics"  },
    { key: "data"      as const, icon: Ico.upload, label: "Data"       },
  ];

  return (
    <div style={{ display:"flex", height:"100vh", fontFamily:"'Geist','Inter',system-ui,sans-serif",
      background:D.bg, color:D.text, overflow:"hidden" }}>

      {/* ── Sidebar ──────────────────────────────────────────────────────── */}
      <aside style={{ width:D.sidebarW, background:D.surface,
        borderRight:`0.5px solid ${D.border}`, display:"flex",
        flexDirection:"column", flexShrink:0 }}>

        {/* Logo */}
        <div style={{ padding:"18px 16px 14px", borderBottom:`0.5px solid ${D.border}`,
          display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:28, height:28, background:D.text, borderRadius:6,
            display:"flex", alignItems:"center", justifyContent:"center", color:D.bg, flexShrink:0 }}>
            {Ico.logo}
          </div>
          <span style={{ fontSize:13, fontWeight:500, color:D.text, letterSpacing:"-0.2px" }}>CakeDash</span>
        </div>

        {/* Nav */}
        <nav style={{ flex:1, padding:"10px 8px", display:"flex", flexDirection:"column", gap:1 }}>
          <p style={{ fontSize:10, color:D.textMuted, textTransform:"uppercase",
            letterSpacing:".08em", padding:"8px 8px 4px", fontWeight:500 }}>Main</p>
          {NAV.map(({ key, icon, label }) => (
            <NavItem key={key} icon={icon} label={label}
              active={activeNav===key} onClick={()=>setActiveNav(key)} />
          ))}

          <div style={{ height:"0.5px", background:D.border, margin:"10px 4px" }} />

          <p style={{ fontSize:10, color:D.textMuted, textTransform:"uppercase",
            letterSpacing:".08em", padding:"4px 8px", fontWeight:500 }}>Settings</p>
          <NavItem icon={Ico.filter} label="Filters" />
        </nav>

        {/* File info */}
        {hasData && (
          <div style={{ borderTop:`0.5px solid ${D.border}`, padding:"12px 10px" }}>
            <div style={{ background:D.surfaceHov, borderRadius:6, padding:"8px 10px" }}>
              <p style={{ margin:0, fontSize:11, fontWeight:500, color:D.text,
                whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{fileName}</p>
              <p style={{ margin:"2px 0 0", fontSize:10, color:D.textMuted }}>
                {data.length.toLocaleString()} rows · {cols.length} cols
              </p>
            </div>
          </div>
        )}
      </aside>

      {/* ── Main ─────────────────────────────────────────────────────────── */}
      <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

        {/* Topbar */}
        <header style={{ height:D.topH, background:D.surface,
          borderBottom:`0.5px solid ${D.border}`, display:"flex",
          alignItems:"center", justifyContent:"space-between",
          padding:"0 22px", flexShrink:0 }}>
          <p style={{ margin:0, fontSize:13, fontWeight:500, color:D.text }}>Dashboard</p>

          {hasData && (
            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
              <Btn onClick={()=>{setEditChart(null);setShowBuilder(true);}}>
                {Ico.plus} เพิ่มกราฟ
              </Btn>
              <Btn onClick={loadDefaults}>{Ico.bolt} กราฟแนะนำ</Btn>
              <Btn onClick={reset} variant="ghost">{Ico.upload} เปลี่ยนไฟล์</Btn>
            </div>
          )}
        </header>

        {/* Content */}
        <main style={{ flex:1, overflowY:"auto", padding:"20px 22px",
          display:"flex", flexDirection:"column", gap:16 }}>

          {!hasData ? (
            <UploadZone onLoad={handleLoad} />
          ) : (
            <>
              {/* Stat row */}
              {genCounts.length > 0 && (
                <div style={{ display:"grid",
                  gridTemplateColumns:`repeat(${Math.min(genCounts.length+1,5)},1fr)`,
                  gap:10 }}>
                  <StatCard label="ผู้ตอบทั้งหมด"
                    value={data.length.toLocaleString()} sub={`${cols.length} คอลัมน์`} />
                  {genCounts.map(([g, n], i) => (
                    <StatCard key={g} label={g} value={String(n)}
                      sub={`${((n/data.length)*100).toFixed(1)}%`}
                      up={i < Math.ceil(genCounts.length/2)} />
                  ))}
                </div>
              )}

              {/* Gen filter */}
              {genCol && (
                <div style={{ background:D.surface, border:`0.5px solid ${D.border}`,
                  borderRadius:8, padding:"12px 16px" }}>
                  <div style={{ display:"flex", alignItems:"center",
                    justifyContent:"space-between", marginBottom:10 }}>
                    <p style={{ margin:0, fontSize:11, fontWeight:500, color:D.text }}>
                      Filter — {genCol}
                    </p>
                    {gFilter.length > 0 && (
                      <button onClick={()=>setGFilter([])}
                        style={{ fontSize:10, color:D.textMuted, background:"none",
                          border:"none", cursor:"pointer", padding:0 }}>
                        ล้าง
                      </button>
                    )}
                  </div>
                  <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                    {["ทั้งหมด",...genVals].map((v) => {
                      const isAll = v === "ทั้งหมด";
                      const on    = isAll ? gFilter.length===0 : gFilter.includes(v);
                      return (
                        <button key={v} onClick={()=>toggleGen(v)} style={{
                          padding:"5px 13px", borderRadius:5, fontSize:11, fontWeight:500,
                          cursor:"pointer", transition:"all .12s",
                          border:`0.5px solid ${on ? D.text : D.border}`,
                          background: on ? D.text : "transparent",
                          color: on ? D.bg : D.textSub,
                        }}>{v}</button>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Charts */}
              {charts.length === 0 ? (
                <div style={{ background:D.surface, border:`1px dashed ${D.border}`,
                  borderRadius:8, padding:"48px 24px", textAlign:"center", flex:1 }}>
                  <div style={{ width:38, height:38, borderRadius:8, background:D.surfaceHov,
                    border:`0.5px solid ${D.border}`, display:"flex", alignItems:"center",
                    justifyContent:"center", margin:"0 auto 12px", color:D.textMuted }}>
                    {Ico.chart}
                  </div>
                  <p style={{ margin:"0 0 4px", fontSize:13, fontWeight:500, color:D.text }}>
                    ยังไม่มีกราฟ
                  </p>
                  <p style={{ margin:"0 0 18px", fontSize:11, color:D.textMuted }}>
                    กด "เพิ่มกราฟ" หรือ "กราฟแนะนำ" ในแถบด้านบน
                  </p>
                  <div style={{ display:"flex", gap:8, justifyContent:"center" }}>
                    <Btn onClick={()=>{setEditChart(null);setShowBuilder(true);}}>
                      {Ico.plus} เพิ่มกราฟ
                    </Btn>
                    <Btn onClick={loadDefaults} variant="solid">
                      {Ico.bolt} โหลดกราฟแนะนำ
                    </Btn>
                  </div>
                </div>
              ) : (
                <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
                  {charts.map((c) => (
                    <ChartCard key={c.id} cfg={c} data={data} genCol={genCol} gFilter={gFilter}
                      onEdit={()=>{setEditChart(c);setShowBuilder(true);}}
                      onRemove={()=>setCharts((p)=>p.filter((x)=>x.id!==c.id))} />
                  ))}
                </div>
              )}
            </>
          )}
        </main>
      </div>

      {/* Builder modal */}
      {showBuilder && (
        <BuilderModal
          cols={cols} data={data} multiCols={multiCols} genCol={genCol}
          initial={editChart ?? undefined}
          onSave={editChart ? saveEdit : addChart}
          onClose={()=>{setShowBuilder(false);setEditChart(null);}} />
      )}
    </div>
  );
}