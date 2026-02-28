"use client";

import { useState, useEffect, useRef, useCallback } from "react";
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  ReferenceLine,
  TooltipProps,
} from "recharts";

/* ══════════════════════════════════════════════════════
   FORMULA (matches calculateMetrics() from JS reference)
   ──────────────────────────────────────────────────────
   Availability = Σ(totalBeban)  /  (avgMesin × 24 × nDays)
   Performance  = Σ(esKeluar)    /  (Σ(totalBeban) × avgProdJam)
   Quality      = (ΣesKeluar - ΣtotalRusak) / ΣesKeluar
   OEE          = Availability × Performance × Quality

   prodJam per row = esKeluar / totalBeban  (bal/jam)
   — taken from col[16] in Excel
══════════════════════════════════════════════════════ */

/* ─── CONSTANTS ─── */
const DEFAULT_CAPACITY = 3896;

/* ─── TYPES ─── */
interface DataRow {
  date: Date;
  dateStr: string;
  dateISO: string;
  esKeluar: number;
  bak1: number;
  bak2: number;
  bak3: number;
  totalRusak: number;
  bebanNormal: number;
  bebanPuncak: number;
  totalBeban: number;
  jumlahMesin: number;
  prodJam: number; // bal/jam — key for Performance calculation
  kapasitas: number;
  unsold: number;
  // per-row OEE components (for sparkline / table display)
  rowQuality: number;
  rowAvailability: number;
  rowPerformance: number;
  rowOEE: number;
}

/** Aggregate OEE metrics calculated over an array of rows (JS-faithful formula) */
interface OEEMetrics {
  availability: number; // 0–1
  performance: number; // 0–1
  quality: number; // 0–1
  oee: number; // 0–1
}

type TabId = "overview" | "produksi" | "tabel";

/* ─── HELPERS ─── */
const pct = (v: number | null, d = 1): string =>
  v != null ? (v * 100).toFixed(d) + "%" : "–";
const num = (v: number | null | undefined): string =>
  v != null ? Number(v).toLocaleString("id-ID") : "–";
const sumK = (arr: DataRow[], k: keyof DataRow): number =>
  arr.reduce((s, d) => s + ((d[k] as number) || 0), 0);

const statusColor = (v: number | null): string => {
  if (v == null) return "#6b7280";
  if (v >= 0.85) return "#10b981";
  if (v >= 0.65) return "#f59e0b";
  return "#ef4444";
};
const statusLabel = (v: number | null): string => {
  if (v == null) return "–";
  if (v >= 0.85) return "World Class";
  if (v >= 0.65) return "Typical";
  return "Low";
};

/**
 * Core aggregate OEE formula — matches calculateMetrics() from the JS reference.
 *
 * Availability = totalJam / (avgMesin × 24 × nDays)
 * Performance  = totalProd / (totalJam × avgProdJam)
 * Quality      = (totalProd − totalRusak) / totalProd
 */
function calcOEEMetrics(rows: DataRow[]): OEEMetrics | null {
  if (!rows.length) return null;
  const n = rows.length;
  const totalProd = sumK(rows, "esKeluar");
  const totalRusak = sumK(rows, "totalRusak");
  const totalJam = sumK(rows, "totalBeban");
  const totalMesin = sumK(rows, "jumlahMesin");
  const totalProdJam = sumK(rows, "prodJam");

  if (totalProd === 0 || totalJam === 0) return null;

  const avgMesin = totalMesin / n;
  const avgProdJam = totalProdJam / n;

  const availability = Math.min(1, totalJam / (avgMesin * 24 * n));
  const performance = Math.min(1, totalProd / (totalJam * avgProdJam));
  const quality = Math.min(
    1,
    Math.max(0, (totalProd - totalRusak) / totalProd),
  );
  const oee = availability * performance * quality;

  return { availability, performance, quality, oee };
}

/* ─── DEMO DATA (Januari 2026 — real data from Excel) ─── */
function buildDemo(): DataRow[] {
  // [esKeluar, bak1, bak2, bak3, bebanNormal, bebanPuncak, totalBeban, mesin, prodJam, kapasitas]
  const rows: [
    number,
    number,
    number,
    number,
    number,
    number,
    number,
    number,
    number,
    number,
  ][] = [
    [601, 0, 0, 0, 16, 0, 16, 2, 37.56, 3896],
    [875, 0, 0, 1, 21, 0, 21, 2, 41.62, 3896],
    [1336, 0, 0, 1, 30, 0, 30, 2, 44.5, 3896],
    [2544, 0, 3, 1, 51, 5, 56, 3, 45.36, 3896],
    [1015, 0, 0, 1, 32, 11, 43, 2, 23.58, 3896],
    [1577, 0, 2, 1, 29, 0, 29, 3, 54.28, 3896],
    [2035, 0, 3, 0, 46, 2, 48, 3, 42.33, 3896],
    [1408, 0, 0, 1, 37, 0, 37, 2, 38.03, 3896],
    [820, 0, 0, 0, 19, 0, 19, 2, 43.16, 3896],
    [921, 0, 0, 1, 21, 0, 21, 2, 43.81, 3896],
    [870, 0, 2, 1, 17, 0, 17, 2, 51.0, 3896],
    [901, 0, 0, 1, 21, 0, 21, 2, 42.86, 3896],
    [1397, 0, 1, 0, 29, 0, 29, 2, 48.14, 3896],
    [1063, 0, 0, 0, 26, 0, 26, 2, 40.88, 3896],
    [898, 0, 0, 1, 18, 0, 18, 2, 49.83, 3896],
    [759, 1, 0, 1, 17, 0, 17, 2, 44.53, 3896],
    [886, 0, 0, 1, 19, 0, 19, 2, 46.58, 3896],
    [1141, 0, 0, 1, 26, 0, 26, 2, 43.85, 3896],
    [804, 0, 0, 1, 21, 0, 21, 2, 38.24, 3896],
    [1099, 0, 0, 1, 23, 0, 23, 2, 47.65, 3896],
    [1562, 0, 2, 2, 35, 0, 35, 2, 44.51, 3896],
    [2395, 3, 5, 0, 52, 0, 52, 2, 45.9, 3896],
    [1464, 3, 2, 0, 27, 15, 42, 2, 34.74, 3896],
    [1339, 2, 0, 1, 28, 0, 28, 2, 47.71, 3896],
    [1255, 2, 0, 1, 26, 1, 27, 2, 46.37, 3896],
    [1380, 0, 4, 0, 26, 0, 26, 2, 52.92, 3896],
    [1034, 1, 0, 1, 25, 0, 25, 2, 41.28, 3896],
    [1103, 0, 0, 1, 26, 0, 26, 2, 42.38, 3896],
    [1310, 4, 0, 1, 30, 0, 30, 2, 43.5, 3896],
    [987, 0, 1, 2, 17, 0, 17, 2, 57.88, 3896],
    [1076, 0, 0, 1, 27, 1, 28, 2, 38.39, 3896],
  ];

  return rows.map(([esK, b1, b2, b3, bN, bP, tB, mesin, pJ, kap], i) => {
    const date = new Date(2026, 0, i + 1);
    const totalRusak = b1 + b2 + b3;
    // per-row components (for table/sparkline)
    const rowAvailability = tB > 0 ? Math.min(1, tB / (mesin * 24)) : 0;
    const rowPerformance = tB > 0 && pJ > 0 ? Math.min(1, esK / (tB * pJ)) : 0;
    const rowQuality =
      esK > 0 ? Math.min(1, Math.max(0, (esK - totalRusak) / esK)) : 0;
    const rowOEE = rowAvailability * rowPerformance * rowQuality;
    return {
      date,
      dateStr: date.toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "short",
      }),
      dateISO: date.toISOString().split("T")[0],
      esKeluar: esK,
      bak1: b1,
      bak2: b2,
      bak3: b3,
      totalRusak,
      bebanNormal: bN,
      bebanPuncak: bP,
      totalBeban: tB,
      jumlahMesin: mesin,
      prodJam: pJ,
      kapasitas: kap,
      unsold: Math.max(0, kap - esK),
      rowQuality,
      rowAvailability,
      rowPerformance,
      rowOEE,
    };
  });
}

/* ─── EXCEL PARSER ─── */
// Sheet "Konsol" column map (0-based):
// col[5]  = Tanggal
// col[6]  = Es Keluar
// col[7]  = Es Rusak Bak 1
// col[8]  = Es Rusak Bak 2
// col[9]  = Es Rusak Bak 3
// col[10] = Total Es Rusak
// col[12] = Beban Normal
// col[13] = Beban Puncak
// col[14] = Total Beban
// col[15] = Jumlah Mesin
// col[16] = Produktivitas Mesin Per Jam (prodJam)
// col[18] = Kapasitas Produksi
// col[19] = Es Tidak Terjual
//
// Rows: idx 0 = plant info, idx 1 = col headers, idx 2 = sub-headers (NOT data), idx 3+ = daily data
function parseExcel(wb: any): DataRow[] {
  const ws = wb.Sheets["Konsol"];
  if (!ws) throw new Error("Sheet 'Konsol' tidak ditemukan.");
  const XLSX = (window as any).XLSX;
  const raw: any[][] = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: false,
  });

  const toN = (v: any): number | null => {
    if (v == null || v === "" || (typeof v === "string" && v.startsWith("#")))
      return null;
    const n = parseFloat(String(v).replace(",", "."));
    return isNaN(n) ? null : n;
  };
  const parseDate = (v: any): Date | null => {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  };

  const records: DataRow[] = [];

  for (let i = 3; i < raw.length; i++) {
    const r = raw[i];
    if (!r || r.length < 7) continue;

    const dateRaw = r[5];
    const esKeluar = toN(r[6]);

    if (!dateRaw) continue;
    if (typeof dateRaw === "string" && dateRaw.toUpperCase().includes("TOTAL"))
      break;
    if (esKeluar == null) continue;

    const date = parseDate(dateRaw);
    if (!date) continue;

    const bak1 = toN(r[7]) ?? 0;
    const bak2 = toN(r[8]) ?? 0;
    const bak3 = toN(r[9]) ?? 0;
    const totalRusak = toN(r[10]) ?? bak1 + bak2 + bak3;
    const bebanNormal = toN(r[12]) ?? 0;
    const bebanPuncak = toN(r[13]) ?? 0;
    const totalBeban = toN(r[14]) ?? bebanNormal + bebanPuncak;
    const jumlahMesin = toN(r[15]) ?? 2;
    // prodJam from col[16]; fallback: esKeluar/totalBeban
    const prodJam = toN(r[16]) ?? (totalBeban > 0 ? esKeluar / totalBeban : 0);
    const kapasitas = toN(r[18]) ?? DEFAULT_CAPACITY;
    const unsold = toN(r[19]) ?? Math.max(0, kapasitas - esKeluar);

    // per-row components (for table display & sparklines)
    const rowAvailability =
      totalBeban > 0 ? Math.min(1, totalBeban / (jumlahMesin * 24)) : 0;
    const rowPerformance =
      totalBeban > 0 && prodJam > 0
        ? Math.min(1, esKeluar / (totalBeban * prodJam))
        : 0;
    const rowQuality =
      esKeluar > 0
        ? Math.min(1, Math.max(0, (esKeluar - totalRusak) / esKeluar))
        : 0;
    const rowOEE = rowAvailability * rowPerformance * rowQuality;

    records.push({
      date,
      dateStr: date.toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "short",
      }),
      dateISO: date.toISOString().split("T")[0],
      esKeluar,
      bak1,
      bak2,
      bak3,
      totalRusak,
      bebanNormal,
      bebanPuncak,
      totalBeban,
      jumlahMesin,
      prodJam,
      kapasitas,
      unsold,
      rowQuality,
      rowAvailability,
      rowPerformance,
      rowOEE,
    });
  }

  if (!records.length)
    throw new Error("Tidak ada data terbaca dari sheet Konsol.");
  return records.sort((a, b) => a.date.getTime() - b.date.getTime());
}

/* ─── SVG DONUT GAUGE ─── */
interface DonutGaugeProps {
  value: number | null;
  color: string;
  size?: number;
}
function DonutGauge({ value, color, size = 100 }: DonutGaugeProps) {
  const r = 38,
    cx = size / 2,
    cy = size / 2,
    circ = 2 * Math.PI * r;
  const safe = value != null ? Math.min(1, Math.max(0, value)) : 0;
  return (
    <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
      <circle
        cx={cx}
        cy={cy}
        r={r}
        fill="none"
        stroke="#e5e7eb"
        strokeWidth={9}
      />
      <circle
        cx={cx}
        cy={cy}
        r={r}
        fill="none"
        stroke={color}
        strokeWidth={9}
        strokeDasharray={`${safe * circ} ${circ}`}
        strokeLinecap="round"
        style={{ transition: "stroke-dasharray 0.7s ease" }}
      />
    </svg>
  );
}

/* ─── GAUGE CARD ─── */
interface GaugeCardProps {
  label: string;
  value: number | null;
  sparkData?: (number | null)[];
  color: string;
  sublabel?: string;
}
function GaugeCard({
  label,
  value,
  sparkData = [],
  color,
  sublabel,
}: GaugeCardProps) {
  const sc = statusColor(value);
  const pill: Record<string, [string, string]> = {
    "#10b981": ["#d1fae5", "#065f46"],
    "#f59e0b": ["#fef3c7", "#78350f"],
    "#ef4444": ["#fee2e2", "#7f1d1d"],
    "#6b7280": ["#f3f4f6", "#374151"],
  };
  const [bg, fg] = pill[sc] ?? ["#f3f4f6", "#374151"];
  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 12,
        padding: "20px 16px",
        boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
        border: "1px solid #e5e7eb",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: 8,
      }}
    >
      <div style={{ fontSize: 13, fontWeight: 600, color: "#1a1a1a" }}>
        {label}
      </div>
      <div style={{ position: "relative", width: 100, height: 100 }}>
        <DonutGauge value={value} color={color} size={100} />
        <div
          style={{
            position: "absolute",
            inset: 0,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <span
            style={{
              fontSize: 20,
              fontWeight: 700,
              color: "#1a1a1a",
              lineHeight: 1,
            }}
          >
            {pct(value, 1)}
          </span>
        </div>
      </div>
      <span
        style={{
          fontSize: 11,
          fontWeight: 600,
          padding: "3px 10px",
          borderRadius: 20,
          background: bg,
          color: fg,
        }}
      >
        {statusLabel(value)}
      </span>
      {sublabel && (
        <div style={{ fontSize: 11, color: "#6b7280" }}>{sublabel}</div>
      )}
      {sparkData.length > 1 && (
        <div style={{ width: "100%", height: 32 }}>
          <ResponsiveContainer width="100%" height={32}>
            <AreaChart
              data={sparkData.map((v, i) => ({
                i,
                v: v != null ? +(v * 100).toFixed(1) : null,
              }))}
            >
              <defs>
                <linearGradient
                  id={`sg-${label.replace(/\s/g, "")}`}
                  x1="0"
                  y1="0"
                  x2="0"
                  y2="1"
                >
                  <stop offset="0%" stopColor={color} stopOpacity={0.2} />
                  <stop offset="100%" stopColor={color} stopOpacity={0} />
                </linearGradient>
              </defs>
              <Area
                type="monotone"
                dataKey="v"
                stroke={color}
                strokeWidth={1.5}
                fill={`url(#sg-${label.replace(/\s/g, "")})`}
                dot={false}
                connectNulls
                isAnimationActive={false}
              />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

/* ─── KPI CARD ─── */
interface KPICardProps {
  label: string;
  value: string | number;
  sub?: string;
  accent?: string;
}
function KPICard({ label, value, sub, accent = "#0066ff" }: KPICardProps) {
  return (
    <div
      style={{ background: "#f5f7fa", borderRadius: 8, padding: "14px 16px" }}
    >
      <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>
        {label}
      </div>
      <div
        style={{ fontSize: 20, fontWeight: 600, color: accent, lineHeight: 1 }}
      >
        {value}
      </div>
      {sub && (
        <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
          {sub}
        </div>
      )}
    </div>
  );
}

/* ─── CHART TOOLTIP ─── */
function ChartTip(props: TooltipProps<number, string> & { payload?: any[] }) {
  const { active, payload } = props;
  if (!active || !payload?.length) return null;
  const label = (props as any).label;
  return (
    <div
      style={{
        background: "#fff",
        border: "1px solid #e5e7eb",
        borderRadius: 8,
        padding: "10px 14px",
        fontSize: 12,
        boxShadow: "0 4px 12px rgba(0,0,0,0.08)",
      }}
    >
      <div style={{ fontWeight: 600, marginBottom: 6, color: "#1a1a1a" }}>
        {label}
      </div>
      {payload.map((p, i) => (
        <div
          key={i}
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: 16,
            marginBottom: 2,
          }}
        >
          <span style={{ color: "#6b7280" }}>{p.name}</span>
          <span
            style={{ color: (p.color as string) || "#1a1a1a", fontWeight: 600 }}
          >
            {typeof p.value === "number" ? p.value.toFixed(1) + "%" : p.value}
          </span>
        </div>
      ))}
    </div>
  );
}

/* ─── DAY STATUS CELL ─── */
interface StatusCell {
  bg: string;
  label: string;
}
function getDayStatus(d: DataRow): StatusCell {
  if (!d.esKeluar) return { bg: "#9ca3af", label: "Idle" };
  if (d.bak3 > 2) return { bg: "#ef4444", label: "Defect Bak 3" };
  if (d.bak2 > 2) return { bg: "#f97316", label: "Defect Bak 2" };
  if (d.bak1 > 2) return { bg: "#f59e0b", label: "Defect Bak 1" };
  if (d.bebanPuncak > 0) return { bg: "#3b82f6", label: "Peak Load" };
  return { bg: "#10b981", label: "Normal" };
}

/* ─── PILL HELPERS ─── */
const pillBg: Record<string, string> = {
  "#10b981": "#d1fae5",
  "#f59e0b": "#fef3c7",
  "#ef4444": "#fee2e2",
  "#6b7280": "#f3f4f6",
};
const pillFg: Record<string, string> = {
  "#10b981": "#065f46",
  "#f59e0b": "#78350f",
  "#ef4444": "#7f1d1d",
  "#6b7280": "#374151",
};

/* ════════════════════════════════════════
   MAIN COMPONENT
════════════════════════════════════════ */
export default function OEEDashboard() {
  const [xlsxReady, setXlsxReady] = useState(false);
  const [allData, setAllData] = useState<DataRow[]>(buildDemo());
  const [fileName, setFileName] = useState("Demo – Januari 2026");
  const [startDate, setStartDate] = useState("2026-01-01");
  const [endDate, setEndDate] = useState("2026-01-31");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState<false | "png" | "jpg">(false);
  const [tab, setTab] = useState<TabId>("overview");
  const fileRef = useRef<HTMLInputElement>(null);
  const dashRef = useRef<HTMLDivElement>(null);

  /* Load SheetJS */
  useEffect(() => {
    if ((window as any).XLSX) {
      setXlsxReady(true);
      return;
    }
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    s.onload = () => setXlsxReady(true);
    document.head.appendChild(s);
  }, []);

  /* Filtered data */
  const data = allData.filter(
    (d) =>
      (!startDate || d.dateISO >= startDate) &&
      (!endDate || d.dateISO <= endDate),
  );

  /* ── AGGREGATE OEE METRICS (the correct formula) ── */
  const metrics = calcOEEMetrics(data);
  const aggOEE = metrics?.oee ?? null;
  const aggAvail = metrics?.availability ?? null;
  const aggPerf = metrics?.performance ?? null;
  const aggQual = metrics?.quality ?? null;

  /* ── PRODUCTION STATS ── */
  const totProd = sumK(data, "esKeluar");
  const totRusak = sumK(data, "totalRusak");
  const totBak1 = sumK(data, "bak1");
  const totBak2 = sumK(data, "bak2");
  const totBak3 = sumK(data, "bak3");
  const totUnsold = sumK(data, "unsold");
  const defRate = totProd > 0 ? totRusak / totProd : 0;
  const avgProdPerDay = data.length ? Math.round(totProd / data.length) : 0;

  /* ── AVERAGE prodJam ── */
  const avgProdJam = data.length ? sumK(data, "prodJam") / data.length : 0;

  /* Upload */
  const handleFile = useCallback(
    async (file: File | null | undefined) => {
      if (!file || !xlsxReady) return;
      setLoading(true);
      setError(null);
      try {
        const buf = await file.arrayBuffer();
        const wb = (window as any).XLSX.read(buf, {
          type: "array",
          cellDates: true,
        });
        const parsed = parseExcel(wb);
        setAllData(parsed);
        setFileName(file.name);
        setStartDate(parsed[0].dateISO);
        setEndDate(parsed[parsed.length - 1].dateISO);
      } catch (e) {
        setError((e as Error).message);
      }
      setLoading(false);
    },
    [xlsxReady],
  );

  /* Export CSV */
  const exportCSV = () => {
    if (!data.length) return;
    const hdr =
      "Tanggal,Es Keluar,Bak 1,Bak 2,Bak 3,Total Rusak,Bbn Normal,Bbn Puncak,Total Beban,Mesin,Prod/Jam,Availability,Performance,Quality,OEE";
    const rows = data.map((d) =>
      [
        d.dateStr,
        d.esKeluar,
        d.bak1,
        d.bak2,
        d.bak3,
        d.totalRusak,
        d.bebanNormal,
        d.bebanPuncak,
        d.totalBeban,
        d.jumlahMesin,
        d.prodJam.toFixed(2),
        (d.rowAvailability * 100).toFixed(2),
        (d.rowPerformance * 100).toFixed(2),
        (d.rowQuality * 100).toFixed(2),
        (d.rowOEE * 100).toFixed(2),
      ].join(","),
    );
    const blob = new Blob([[hdr, ...rows].join("\n")], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `OEE_PMP_${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
  };

  /* Export Image (PNG / JPG) via html2canvas */
  const exportImage = async (fmt: "png" | "jpg") => {
    if (!dashRef.current) return;
    setExporting(fmt);
    try {
      if (!(window as any).html2canvas) {
        await new Promise<void>((res, rej) => {
          const s = document.createElement("script");
          s.src =
            "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
          s.onload = () => res();
          s.onerror = () => rej(new Error("Gagal memuat html2canvas"));
          document.head.appendChild(s);
        });
      }
      const canvas: HTMLCanvasElement = await (window as any).html2canvas(
        dashRef.current,
        {
          backgroundColor: "#f5f7fa",
          scale: 2,
          useCORS: true,
          logging: false,
          windowWidth: dashRef.current.scrollWidth,
          windowHeight: dashRef.current.scrollHeight,
          scrollX: 0,
          scrollY: -window.scrollY,
        },
      );
      const mimeType = fmt === "jpg" ? "image/jpeg" : "image/png";
      const quality = fmt === "jpg" ? 0.92 : undefined;
      const dataUrl = canvas.toDataURL(mimeType, quality);
      const dateStr = new Date().toISOString().split("T")[0];
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = `OEE_Dashboard_${dateStr}.${fmt}`;
      a.click();
    } catch (e) {
      setError("Export gambar gagal: " + (e as Error).message);
    }
    setExporting(false);
  };

  /* Chart data — use per-row OEE components for trend lines */
  const lineData = data.map((d) => ({
    date: d.dateStr,
    OEE: +(d.rowOEE * 100).toFixed(1),
    Availability: +(d.rowAvailability * 100).toFixed(1),
    Performance: +(d.rowPerformance * 100).toFixed(1),
    Quality: +(d.rowQuality * 100).toFixed(1),
  }));

  const barData = data.map((d) => ({
    date: d.dateStr,
    "Es Keluar": d.esKeluar,
    "Bak 1": d.bak1,
    "Bak 2": d.bak2,
    "Bak 3": d.bak3,
  }));

  /* Group by month */
  const months: Record<string, DataRow[]> = {};
  data.forEach((d) => {
    const k = d.date.toLocaleDateString("id-ID", {
      month: "long",
      year: "numeric",
    });
    if (!months[k]) months[k] = [];
    months[k].push(d);
  });

  const tabs: [TabId, string][] = [
    ["overview", "Overview"],
    ["produksi", "Produksi & Defect"],
    ["tabel", "Data Harian"],
  ];

  /* ────────────────────────────────────────────── */
  return (
    <div
      ref={dashRef}
      style={{
        fontFamily: "'Inter', sans-serif",
        background: "#f5f7fa",
        color: "#1a1a1a",
        padding: 20,
        minHeight: "100vh",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 2px; }
        input[type=date] { color-scheme: light; }
        @keyframes fadeUp { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .day-cell:hover { opacity: 1 !important; transform: scale(1.1) !important; }
        .tbl-row:hover td { background: #f9fafb !important; }
        .tab-btn:not(.active):hover { border-color: #0066ff !important; color: #0066ff !important; }
      `}</style>

      {/* ══ HEADER ══ */}
      <div
        style={{
          background: "#fff",
          padding: "14px 24px",
          borderRadius: 12,
          marginBottom: 20,
          boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
          border: "1px solid #e5e7eb",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          flexWrap: "wrap",
          gap: 12,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 24 }}>🧊</span>
          <div>
            <div style={{ fontSize: 18, fontWeight: 700, lineHeight: 1.1 }}>
              OEE Dashboard
            </div>
            <div style={{ fontSize: 11, color: "#6b7280", marginTop: 1 }}>
              PMP Ice Plant · Monitoring
            </div>
          </div>
        </div>

        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            flexWrap: "wrap",
          }}
        >
          {/* Tabs */}
          <div style={{ display: "flex", gap: 6 }}>
            {tabs.map(([id, lbl]) => (
              <button
                key={id}
                onClick={() => setTab(id)}
                className={`tab-btn${tab === id ? " active" : ""}`}
                style={{
                  padding: "7px 16px",
                  borderRadius: 6,
                  fontSize: 13,
                  fontWeight: 500,
                  cursor: "pointer",
                  border: "1px solid",
                  borderColor: tab === id ? "#0066ff" : "#e5e7eb",
                  background: tab === id ? "#0066ff" : "#fff",
                  color: tab === id ? "#fff" : "#6b7280",
                  transition: "all 0.15s",
                }}
              >
                {lbl}
              </button>
            ))}
          </div>

          {/* Date range */}
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              style={{
                padding: "7px 10px",
                border: "1px solid #e5e7eb",
                borderRadius: 6,
                fontSize: 13,
                fontFamily: "Inter,sans-serif",
              }}
            />
            <span style={{ color: "#9ca3af" }}>–</span>
            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              style={{
                padding: "7px 10px",
                border: "1px solid #e5e7eb",
                borderRadius: 6,
                fontSize: 13,
                fontFamily: "Inter,sans-serif",
              }}
            />
          </div>

          <button
            onClick={() => fileRef.current?.click()}
            style={{
              padding: "7px 14px",
              borderRadius: 6,
              fontSize: 13,
              fontWeight: 500,
              cursor: "pointer",
              border: "none",
              background: loading ? "#e5e7eb" : "#0066ff",
              color: loading ? "#6b7280" : "#fff",
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            {loading ? "⏳ Memproses..." : "📂 Upload XLSX"}
          </button>
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={(e) => {
              handleFile(e.target.files?.[0]);
              e.target.value = "";
            }}
          />

          {/* Export group */}
          <div style={{ display: "flex", gap: 4 }}>
            <button
              onClick={exportCSV}
              style={{
                padding: "7px 12px",
                borderRadius: "6px 0 0 6px",
                fontSize: 13,
                cursor: "pointer",
                background: "#fff",
                border: "1px solid #e5e7eb",
                borderRight: "none",
                color: "#6b7280",
                transition: "all .15s",
              }}
            >
              ⬇ CSV
            </button>
            <button
              onClick={() => exportImage("png")}
              disabled={exporting !== false}
              style={{
                padding: "7px 12px",
                borderRadius: 0,
                fontSize: 13,
                cursor: exporting ? "wait" : "pointer",
                background: exporting === "png" ? "#e0f2fe" : "#fff",
                border: "1px solid #e5e7eb",
                borderRight: "none",
                color: exporting === "png" ? "#0284c7" : "#6b7280",
                transition: "all .15s",
              }}
            >
              {exporting === "png" ? "⏳ PNG…" : "🖼 PNG"}
            </button>
            <button
              onClick={() => exportImage("jpg")}
              disabled={exporting !== false}
              style={{
                padding: "7px 12px",
                borderRadius: "0 6px 6px 0",
                fontSize: 13,
                cursor: exporting ? "wait" : "pointer",
                background: exporting === "jpg" ? "#e0f2fe" : "#fff",
                border: "1px solid #e5e7eb",
                color: exporting === "jpg" ? "#0284c7" : "#6b7280",
                transition: "all .15s",
              }}
            >
              {exporting === "jpg" ? "⏳ JPG…" : "📷 JPG"}
            </button>
          </div>
        </div>
      </div>

      {error && (
        <div
          style={{
            background: "#fee2e2",
            border: "1px solid #fca5a5",
            borderRadius: 8,
            padding: "10px 16px",
            marginBottom: 16,
            color: "#7f1d1d",
            fontSize: 13,
            display: "flex",
            justifyContent: "space-between",
          }}
        >
          <span>⚠ {error}</span>
          <span style={{ cursor: "pointer" }} onClick={() => setError(null)}>
            ✕
          </span>
        </div>
      )}

      {fileName.startsWith("Demo") && (
        <div
          onClick={() => fileRef.current?.click()}
          style={{
            border: "1.5px dashed #93c5fd",
            borderRadius: 10,
            padding: "10px 20px",
            cursor: "pointer",
            color: "#6b7280",
            fontSize: 12,
            textAlign: "center",
            background: "#eff6ff",
            marginBottom: 16,
          }}
        >
          📂 Upload{" "}
          <strong style={{ color: "#0066ff" }}>file XLSX PMP Plant</strong>{" "}
          untuk data real &nbsp;·&nbsp; Menampilkan demo data Januari 2026
        </div>
      )}

      {!fileName.startsWith("Demo") && (
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            padding: "4px 12px",
            background: "#dcfce7",
            borderRadius: 20,
            fontSize: 11,
            color: "#166534",
            marginBottom: 16,
          }}
        >
          ✓ {fileName} · {data.length} hari
        </div>
      )}

      {/* ══ OEE FORMULA NOTE ══ */}
      <div
        style={{
          background: "#fff",
          border: "1px solid #e5e7eb",
          borderRadius: 10,
          padding: "10px 18px",
          marginBottom: 20,
          fontSize: 11,
          color: "#6b7280",
          display: "flex",
          gap: 20,
          flexWrap: "wrap",
        }}
      >
        <span style={{ fontWeight: 600, color: "#374151" }}>Formula OEE:</span>
        <span>
          📐 <b>Availability</b> = ΣJam Kerja ÷ (Avg Mesin × 24 × n Hari)
        </span>
        <span>
          📐 <b>Performance</b> = ΣEs Keluar ÷ (ΣJam Kerja × Avg Prod/Jam)
        </span>
        <span>
          📐 <b>Quality</b> = (ΣEs Keluar − ΣRusak) ÷ ΣEs Keluar
        </span>
        <span>
          📐 <b>OEE</b> = A × P × Q
        </span>
      </div>

      {/* ══ TAB: OVERVIEW ══ */}
      {tab === "overview" && (
        <div style={{ animation: "fadeUp .25s ease" }}>
          {/* Gauge row */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(4,1fr)",
              gap: 16,
              marginBottom: 20,
            }}
          >
            <GaugeCard
              label="OEE"
              value={aggOEE}
              sparkData={data.map((d) => d.rowOEE)}
              color={statusColor(aggOEE)}
              sublabel={`${data.length} hari`}
            />
            <GaugeCard
              label="Availability"
              value={aggAvail}
              sparkData={data.map((d) => d.rowAvailability)}
              color="#3b82f6"
              sublabel="ΣJam ÷ (AvgMesin×24×n)"
            />
            <GaugeCard
              label="Performance"
              value={aggPerf}
              sparkData={data.map((d) => d.rowPerformance)}
              color="#f59e0b"
              sublabel="ΣProd ÷ (ΣJam×AvgPJ)"
            />
            <GaugeCard
              label="Quality"
              value={aggQual}
              sparkData={data.map((d) => d.rowQuality)}
              color="#10b981"
              sublabel="(ΣProd−ΣRusak) ÷ ΣProd"
            />
          </div>

          {/* KPI + Trend */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "340px 1fr",
              gap: 16,
              marginBottom: 20,
            }}
          >
            {/* KPI panel */}
            <div
              style={{
                background: "#fff",
                borderRadius: 12,
                padding: 24,
                boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                border: "1px solid #e5e7eb",
              }}
            >
              <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 16 }}>
                Ringkasan Produksi
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 12,
                }}
              >
                <KPICard
                  label="Total Produksi"
                  value={num(totProd)}
                  sub="balok es keluar"
                  accent="#0066ff"
                />
                <KPICard
                  label="Total Defect"
                  value={num(totRusak)}
                  sub={`B1:${totBak1} B2:${totBak2} B3:${totBak3}`}
                  accent="#ef4444"
                />
                <KPICard
                  label="Defect Rate"
                  value={pct(defRate, 2)}
                  sub="dari total prod."
                  accent={statusColor(1 - defRate)}
                />
                <KPICard
                  label="Avg Prod/Jam"
                  value={avgProdJam.toFixed(1)}
                  sub="bal/jam"
                  accent="#f59e0b"
                />
                <KPICard
                  label="Rata-rata/Hari"
                  value={num(avgProdPerDay)}
                  sub="bal/hari"
                  accent="#0066ff"
                />
                <KPICard
                  label="Es Tdk Terjual"
                  value={num(totUnsold)}
                  sub="bal"
                  accent="#6b7280"
                />
                <KPICard
                  label="Gap World Class"
                  value={aggOEE != null ? pct(Math.max(0, 0.85 - aggOEE)) : "–"}
                  sub={
                    aggOEE != null && aggOEE >= 0.85
                      ? "✅ Tercapai"
                      : "target ≥ 85%"
                  }
                  accent={
                    aggOEE != null && aggOEE >= 0.85 ? "#10b981" : "#ef4444"
                  }
                />
                <KPICard
                  label="Hari Dianalisa"
                  value={data.length}
                  sub={`${startDate} → ${endDate}`}
                  accent="#6b7280"
                />
              </div>
            </div>

            {/* Trend chart */}
            <div
              style={{
                background: "#fff",
                borderRadius: 12,
                padding: 24,
                boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                border: "1px solid #e5e7eb",
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginBottom: 16,
                }}
              >
                <div style={{ fontSize: 15, fontWeight: 600 }}>
                  Tren OEE Harian
                </div>
                <div style={{ display: "flex", gap: 14 }}>
                  {(
                    [
                      ["OEE", statusColor(aggOEE)],
                      ["Availability", "#3b82f6"],
                      ["Performance", "#f59e0b"],
                      ["Quality", "#10b981"],
                    ] as [string, string][]
                  ).map(([l, c]) => (
                    <div
                      key={l}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 5,
                        fontSize: 11,
                        color: "#6b7280",
                      }}
                    >
                      <div
                        style={{
                          width: 14,
                          height: 2,
                          background: c,
                          borderRadius: 1,
                        }}
                      />
                      {l}
                    </div>
                  ))}
                </div>
              </div>
              <ResponsiveContainer width="100%" height={240}>
                <LineChart data={lineData}>
                  <CartesianGrid
                    strokeDasharray="3 6"
                    stroke="#f3f4f6"
                    vertical={false}
                  />
                  <XAxis
                    dataKey="date"
                    stroke="#d1d5db"
                    tick={{ fontSize: 10, fill: "#9ca3af" }}
                    tickLine={false}
                  />
                  <YAxis
                    domain={[0, 100]}
                    tickFormatter={(v) => v + "%"}
                    stroke="#d1d5db"
                    tick={{ fontSize: 10, fill: "#9ca3af" }}
                    tickLine={false}
                    axisLine={false}
                  />
                  <Tooltip content={<ChartTip />} />
                  <ReferenceLine
                    y={85}
                    stroke="#10b981"
                    strokeDasharray="4 4"
                    strokeOpacity={0.6}
                    label={{
                      value: "85% WC",
                      fill: "#10b981",
                      fontSize: 10,
                      position: "insideTopRight",
                    }}
                  />
                  <Line
                    dataKey="OEE"
                    name="OEE"
                    stroke={statusColor(aggOEE)}
                    strokeWidth={2.5}
                    dot={false}
                    connectNulls
                    activeDot={{ r: 4 }}
                  />
                  <Line
                    dataKey="Availability"
                    name="Availability"
                    stroke="#3b82f6"
                    strokeWidth={1.5}
                    strokeDasharray="4 2"
                    dot={false}
                    connectNulls
                  />
                  <Line
                    dataKey="Performance"
                    name="Performance"
                    stroke="#f59e0b"
                    strokeWidth={1.5}
                    strokeDasharray="4 2"
                    dot={false}
                    connectNulls
                  />
                  <Line
                    dataKey="Quality"
                    name="Quality"
                    stroke="#10b981"
                    strokeWidth={1.5}
                    strokeDasharray="4 2"
                    dot={false}
                    connectNulls
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Status calendar */}
          <div
            style={{
              background: "#fff",
              borderRadius: 12,
              padding: 24,
              boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
              border: "1px solid #e5e7eb",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 16,
                flexWrap: "wrap",
                gap: 8,
              }}
            >
              <div style={{ fontSize: 15, fontWeight: 600 }}>
                Status Mesin Harian
              </div>
              <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                {(
                  [
                    { c: "#10b981", l: "Normal" },
                    { c: "#3b82f6", l: "Peak Load" },
                    { c: "#f59e0b", l: "Defect Bak 1" },
                    { c: "#f97316", l: "Defect Bak 2" },
                    { c: "#ef4444", l: "Defect Bak 3" },
                    { c: "#9ca3af", l: "Idle" },
                  ] as { c: string; l: string }[]
                ).map(({ c, l }) => (
                  <div
                    key={l}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 5,
                      fontSize: 11,
                      color: "#6b7280",
                    }}
                  >
                    <div
                      style={{
                        width: 10,
                        height: 10,
                        borderRadius: 2,
                        background: c,
                      }}
                    />
                    {l}
                  </div>
                ))}
              </div>
            </div>
            {Object.entries(months).map(([month, days]) => (
              <div key={month} style={{ marginBottom: 14 }}>
                <div
                  style={{
                    fontSize: 11,
                    color: "#9ca3af",
                    fontWeight: 600,
                    letterSpacing: 1,
                    marginBottom: 8,
                    textTransform: "uppercase",
                  }}
                >
                  {month}
                </div>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                  {days.map((d, i) => {
                    const { bg } = getDayStatus(d);
                    return (
                      <div
                        key={i}
                        className="day-cell"
                        title={`${d.dateStr}\nOEE: ${pct(d.rowOEE)}\nEs Keluar: ${num(d.esKeluar)} bal\nBak1:${d.bak1} Bak2:${d.bak2} Bak3:${d.bak3}`}
                        style={{
                          flex: "1 0 30px",
                          maxWidth: 46,
                          height: 42,
                          borderRadius: 6,
                          background: bg,
                          opacity: 0.85,
                          cursor: "default",
                          display: "flex",
                          flexDirection: "column",
                          alignItems: "center",
                          justifyContent: "center",
                          gap: 2,
                          transition: "transform 0.12s, opacity 0.12s",
                        }}
                      >
                        <span
                          style={{
                            fontSize: 9,
                            color: "rgba(255,255,255,.7)",
                            lineHeight: 1,
                          }}
                        >
                          {d.date.getDate()}
                        </span>
                        <span
                          style={{
                            fontSize: 9,
                            color: "rgba(255,255,255,.95)",
                            lineHeight: 1,
                            fontWeight: 700,
                          }}
                        >
                          {(d.rowOEE * 100).toFixed(0)}%
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ══ TAB: PRODUKSI ══ */}
      {tab === "produksi" && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 340px",
            gap: 16,
            animation: "fadeUp .25s ease",
          }}
        >
          {/* Bar chart */}
          <div
            style={{
              background: "#fff",
              borderRadius: 12,
              padding: 24,
              boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
              border: "1px solid #e5e7eb",
            }}
          >
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 16 }}>
              Produksi Harian & Defect
            </div>
            <ResponsiveContainer width="100%" height={320}>
              <BarChart data={barData} barGap={2}>
                <CartesianGrid
                  strokeDasharray="3 6"
                  stroke="#f3f4f6"
                  vertical={false}
                />
                <XAxis
                  dataKey="date"
                  stroke="#d1d5db"
                  tick={{ fontSize: 10, fill: "#9ca3af" }}
                  tickLine={false}
                />
                <YAxis
                  stroke="#d1d5db"
                  tick={{ fontSize: 10, fill: "#9ca3af" }}
                  tickLine={false}
                  axisLine={false}
                />
                <Tooltip content={<ChartTip />} />
                <ReferenceLine
                  y={DEFAULT_CAPACITY}
                  stroke="#0066ff"
                  strokeDasharray="3 3"
                  strokeOpacity={0.5}
                  label={{
                    value: "Kapasitas",
                    fill: "#0066ff",
                    fontSize: 10,
                    position: "insideTopRight",
                  }}
                />
                <Bar dataKey="Es Keluar" name="Es Keluar" radius={[3, 3, 0, 0]}>
                  {barData.map((d, i) => (
                    <Cell
                      key={i}
                      fill={d["Es Keluar"] < 1000 ? "#9ca3af" : "#0066ff"}
                      fillOpacity={0.7}
                    />
                  ))}
                </Bar>
                <Bar
                  dataKey="Bak 1"
                  name="Bak 1"
                  fill="#f59e0b"
                  radius={[3, 3, 0, 0]}
                  fillOpacity={0.9}
                />
                <Bar
                  dataKey="Bak 2"
                  name="Bak 2"
                  fill="#f97316"
                  radius={[3, 3, 0, 0]}
                  fillOpacity={0.9}
                />
                <Bar
                  dataKey="Bak 3"
                  name="Bak 3"
                  fill="#ef4444"
                  radius={[3, 3, 0, 0]}
                  fillOpacity={0.9}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Right panel */}
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {/* Defect breakdown */}
            <div
              style={{
                background: "#fff",
                borderRadius: 12,
                padding: 24,
                boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                border: "1px solid #e5e7eb",
              }}
            >
              <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 16 }}>
                Breakdown Defect
              </div>
              {(
                [
                  { label: "Bak 1", val: totBak1, color: "#f59e0b" },
                  { label: "Bak 2", val: totBak2, color: "#f97316" },
                  { label: "Bak 3", val: totBak3, color: "#ef4444" },
                ] as { label: string; val: number; color: string }[]
              ).map((item) => (
                <div key={item.label} style={{ marginBottom: 14 }}>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      marginBottom: 6,
                    }}
                  >
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 7 }}
                    >
                      <div
                        style={{
                          width: 8,
                          height: 8,
                          borderRadius: 2,
                          background: item.color,
                        }}
                      />
                      <span style={{ fontSize: 12, color: "#6b7280" }}>
                        Rusak {item.label}
                      </span>
                    </div>
                    <span
                      style={{
                        fontSize: 13,
                        fontWeight: 600,
                        color: item.color,
                      }}
                    >
                      {num(item.val)} bal
                    </span>
                  </div>
                  <div
                    style={{
                      background: "#f3f4f6",
                      borderRadius: 4,
                      height: 6,
                      overflow: "hidden",
                    }}
                  >
                    <div
                      style={{
                        height: "100%",
                        borderRadius: 4,
                        background: item.color,
                        width: `${totRusak ? (item.val / totRusak) * 100 : 0}%`,
                        transition: "width .6s ease",
                      }}
                    />
                  </div>
                </div>
              ))}
              <div
                style={{
                  paddingTop: 14,
                  borderTop: "1px solid #e5e7eb",
                  display: "flex",
                  justifyContent: "space-between",
                }}
              >
                <span style={{ fontSize: 12, color: "#6b7280" }}>
                  Total Defect
                </span>
                <span
                  style={{ fontSize: 16, fontWeight: 700, color: "#ef4444" }}
                >
                  {num(totRusak)} bal
                </span>
              </div>
            </div>

            {/* OEE summary */}
            <div
              style={{
                background: "#fff",
                borderRadius: 12,
                padding: 24,
                boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
                border: "1px solid #e5e7eb",
              }}
            >
              <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 16 }}>
                Ringkasan OEE
              </div>
              {(
                [
                  ["Total Produksi", `${num(totProd)} bal`, "#0066ff"],
                  ["Total Defect", `${num(totRusak)} bal`, "#ef4444"],
                  ["Defect Rate", pct(defRate, 2), statusColor(1 - defRate)],
                  ["Rata-rata/Hari", `${num(avgProdPerDay)} bal`, "#0066ff"],
                  ["Avg Prod/Jam", `${avgProdJam.toFixed(1)} bal`, "#f59e0b"],
                  ["Es Tdk Terjual", `${num(totUnsold)} bal`, "#6b7280"],
                  ["OEE", pct(aggOEE), statusColor(aggOEE)],
                  ["Availability", pct(aggAvail), "#3b82f6"],
                  ["Performance", pct(aggPerf), "#f59e0b"],
                  ["Quality", pct(aggQual), "#10b981"],
                ] as [string, string, string][]
              ).map(([l, v, c]) => (
                <div
                  key={l}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    paddingBottom: 9,
                    marginBottom: 9,
                    borderBottom: "1px solid #f3f4f6",
                  }}
                >
                  <span style={{ fontSize: 12, color: "#6b7280" }}>{l}</span>
                  <span style={{ fontSize: 14, fontWeight: 600, color: c }}>
                    {v}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ══ TAB: TABEL ══ */}
      {tab === "tabel" && (
        <div
          style={{
            background: "#fff",
            borderRadius: 12,
            overflow: "hidden",
            boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
            border: "1px solid #e5e7eb",
            animation: "fadeUp .25s ease",
          }}
        >
          <div
            style={{
              padding: "14px 24px",
              borderBottom: "1px solid #e5e7eb",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <span style={{ fontSize: 15, fontWeight: 600 }}>
              Data OEE Per Hari
            </span>
            <span style={{ fontSize: 11, color: "#6b7280" }}>
              {data.length} hari &nbsp;·&nbsp; OEE agg:{" "}
              <strong style={{ color: statusColor(aggOEE) }}>
                {pct(aggOEE)}
              </strong>
            </span>
          </div>
          <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: 12,
              }}
            >
              <thead>
                <tr
                  style={{
                    background: "#f9fafb",
                    position: "sticky",
                    top: 0,
                    zIndex: 1,
                  }}
                >
                  {[
                    "Tanggal",
                    "Es Keluar",
                    "Bak 1",
                    "Bak 2",
                    "Bak 3",
                    "Tot. Rusak",
                    "Bbn Normal",
                    "Bbn Puncak",
                    "Mesin",
                    "Prod/Jam",
                    "Availability",
                    "Performance",
                    "Quality",
                    "OEE",
                    "Status",
                  ].map((h) => (
                    <th
                      key={h}
                      style={{
                        padding: "10px 12px",
                        textAlign: "left",
                        color: "#6b7280",
                        fontWeight: 600,
                        fontSize: 11,
                        letterSpacing: 0.3,
                        whiteSpace: "nowrap",
                        borderBottom: "1px solid #e5e7eb",
                      }}
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.map((d, i) => {
                  const sc = statusColor(d.rowOEE);
                  return (
                    <tr
                      key={i}
                      className="tbl-row"
                      style={{ borderTop: "1px solid #f3f4f6" }}
                    >
                      <td
                        style={{
                          padding: "9px 12px",
                          whiteSpace: "nowrap",
                          fontWeight: 500,
                        }}
                      >
                        {d.dateStr}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          fontWeight: 600,
                          color: "#1a1a1a",
                        }}
                      >
                        {num(d.esKeluar)}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: d.bak1 > 0 ? "#f59e0b" : "#d1d5db",
                        }}
                      >
                        {d.bak1}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: d.bak2 > 0 ? "#f97316" : "#d1d5db",
                        }}
                      >
                        {d.bak2}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: d.bak3 > 0 ? "#ef4444" : "#d1d5db",
                        }}
                      >
                        {d.bak3}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: d.totalRusak > 0 ? "#ef4444" : "#d1d5db",
                          fontWeight: d.totalRusak > 0 ? 600 : 400,
                        }}
                      >
                        {d.totalRusak}
                      </td>
                      <td style={{ padding: "9px 12px", color: "#6b7280" }}>
                        {d.bebanNormal}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: d.bebanPuncak > 0 ? "#3b82f6" : "#6b7280",
                        }}
                      >
                        {d.bebanPuncak}
                      </td>
                      <td style={{ padding: "9px 12px", color: "#6b7280" }}>
                        {d.jumlahMesin}
                      </td>
                      <td style={{ padding: "9px 12px", color: "#6b7280" }}>
                        {d.prodJam.toFixed(1)}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: statusColor(d.rowAvailability),
                          fontWeight: 600,
                        }}
                      >
                        {pct(d.rowAvailability)}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: statusColor(d.rowPerformance),
                          fontWeight: 600,
                        }}
                      >
                        {pct(d.rowPerformance)}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: statusColor(d.rowQuality),
                          fontWeight: 600,
                        }}
                      >
                        {pct(d.rowQuality)}
                      </td>
                      <td
                        style={{
                          padding: "9px 12px",
                          color: sc,
                          fontWeight: 700,
                        }}
                      >
                        {pct(d.rowOEE)}
                      </td>
                      <td style={{ padding: "9px 12px" }}>
                        <span
                          style={{
                            fontSize: 10,
                            padding: "3px 9px",
                            borderRadius: 12,
                            fontWeight: 600,
                            background: pillBg[sc] ?? "#f3f4f6",
                            color: pillFg[sc] ?? "#374151",
                          }}
                        >
                          {statusLabel(d.rowOEE)}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              {data.length > 0 && (
                <tfoot>
                  <tr
                    style={{
                      borderTop: "2px solid #e5e7eb",
                      background: "#f9fafb",
                    }}
                  >
                    <td
                      style={{
                        padding: "10px 12px",
                        fontWeight: 700,
                        fontSize: 12,
                      }}
                    >
                      AGREGAT
                    </td>
                    <td style={{ padding: "10px 12px", fontWeight: 600 }}>
                      {num(totProd)}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: "#f59e0b",
                        fontWeight: 600,
                      }}
                    >
                      {totBak1}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: "#f97316",
                        fontWeight: 600,
                      }}
                    >
                      {totBak2}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: "#ef4444",
                        fontWeight: 600,
                      }}
                    >
                      {totBak3}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: "#ef4444",
                        fontWeight: 600,
                      }}
                    >
                      {totRusak}
                    </td>
                    <td
                      colSpan={3}
                      style={{ padding: "10px 12px", color: "#9ca3af" }}
                    >
                      –
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: "#f59e0b",
                        fontWeight: 600,
                      }}
                    >
                      {avgProdJam.toFixed(1)}
                    </td>
                    {/* Aggregate OEE components */}
                    <td
                      style={{
                        padding: "10px 12px",
                        color: statusColor(aggAvail),
                        fontWeight: 700,
                      }}
                    >
                      {pct(aggAvail)}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: statusColor(aggPerf),
                        fontWeight: 700,
                      }}
                    >
                      {pct(aggPerf)}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: statusColor(aggQual),
                        fontWeight: 700,
                      }}
                    >
                      {pct(aggQual)}
                    </td>
                    <td
                      style={{
                        padding: "10px 12px",
                        color: statusColor(aggOEE),
                        fontWeight: 700,
                      }}
                    >
                      {pct(aggOEE)}
                    </td>
                    <td style={{ padding: "10px 12px" }}>
                      <span
                        style={{
                          fontSize: 10,
                          padding: "3px 9px",
                          borderRadius: 12,
                          fontWeight: 600,
                          background: pillBg[statusColor(aggOEE)] ?? "#f3f4f6",
                          color: pillFg[statusColor(aggOEE)] ?? "#374151",
                        }}
                      >
                        {statusLabel(aggOEE)}
                      </span>
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      )}

      <div
        style={{
          textAlign: "center",
          fontSize: 10,
          color: "#9ca3af",
          paddingTop: 20,
          paddingBottom: 8,
        }}
      >
        OEE = Availability × Performance × Quality &nbsp;·&nbsp; Target World
        Class ≥ 85% &nbsp;·&nbsp; PMP Ice Plant
      </div>
    </div>
  );
}
