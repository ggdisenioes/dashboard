// --- IndexedDB persistence & types ---
type SupplierGroupOverride = { merge: boolean; chosenDisplay: string };

type PersistedStateV1 = {
  version: 1;
  savedAt: number;
  rows: Row[];
  headers: string[];
  colTypes: Record<string, ColType>;
  filesLoaded: number;
  mapping: FieldMapping;
  fixedFilters: {
    Buyer: string[];
    Supplier: string[];
    Country: string[];
    Year: string[];
    Sector: string[];
    Turnover: { from: number | null; to: number | null };
  };
  dynamicFilters: DynamicFilter[];
  supplierUnifyEnabled: boolean;
  supplierGroupOverrides: Record<string, SupplierGroupOverride>;
  fuzzyApprovedAliases: Record<string, string>;
};

const IDB_DB = "dashboard_suppliers";
const IDB_STORE = "kv";
const IDB_KEY = "snapshot_v1";

function idbOpen(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_DB, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbGet<T>(key: string): Promise<T | null> {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const store = tx.objectStore(IDB_STORE);
    const req = store.get(key);
    req.onsuccess = () => resolve((req.result ?? null) as T | null);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

async function idbSet<T>(key: string, value: T): Promise<void> {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    const store = tx.objectStore(IDB_STORE);
    const req = store.put(value as any, key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

async function idbDel(key: string): Promise<void> {
  const db = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    const store = tx.objectStore(IDB_STORE);
    const req = store.delete(key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}
import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  LineChart,
  Line,
  CartesianGrid,
} from "recharts";

// Twinco branding colors and logo
const BRAND = {
  navy: "#030F33",
  neon: "#74F9C0",
  softGrey: "#C5C6C7",
  lightGrey: "#F7F7F7",
  white: "#FFFFFF",
  softBlue: "#78A8E4",
  skyBlue: "#4580E0",
  electricBlue: "#002196",
  darkBlue: "#001469",
};

// Put the logo file in /public with this exact name.
const HEADER_LOGO_SRC = "/LogotipoTwinco_Negativo.png";

const EMPTY_OPTION = "(Sin asignar)";

/**
 * Dashboard Suppliers
 * - Carga múltiples CSV/XLSX (se unifican en un solo dataset)
 * - Filtros fijos: Buyer, Supplier (con unificación), Country, Year, Sector, Turnover
 * - Filtros dinámicos: Add filter
 * - Unificación de Suppliers:
 *   - Normaliza y detecta variantes exactas por canonicalKey
 *   - Vista previa para autorizar merge o separar
 *   - UPGRADE 1: buscador de grupos
 *   - UPGRADE 2: Apply all suggested merges
 *   - UPGRADE 3: fuzzy suggestions (probables duplicados aun si no cae en mismo canonicalKey)
 * - Export Excel: exporta el dataset filtrado (con Supplier unificado si ON)
 * - Todo local, sin red.
 */

type Row = Record<string, any>;
type ColType = "number" | "string" | "boolean" | "date" | "unknown";
type FixedFieldKey = "Buyer" | "Supplier" | "Country" | "Year" | "Turnover" | "Sector";
type FieldMapping = Record<FixedFieldKey, string | null>;

type FilterKind = "multi" | "range" | "boolean" | "text";
type DynamicFilter = { id: string; column: string; kind: FilterKind; value: any };

type SupplierVariant = { name: string; count: number };

type SupplierGroup = {
  key: string; // canonical key
  variants: SupplierVariant[];
  suggestedDisplay: string; // variante más frecuente
  chosenDisplay: string; // elegida por usuario
  merge: boolean; // autoriza unificación
};

type FuzzySuggestion = {
  a: string; // supplier raw
  b: string; // supplier raw
  score: number; // 0..1
  chosenDisplay: string; // display a aplicar
  approved: boolean | null; // null=pendiente
};

function normalizeHeader(h: any) {
  return String(h ?? "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero-width
    .replace(/\u00A0/g, " ") // NBSP
    .replace(/\s+/g, " ")
    .trim();
}
function safeStr(v: any) {
  return String(v ?? "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .trim();
}
function isEmptyish(v: any) {
  const s = safeStr(v).toLowerCase();
  return (
    s === "" ||
    s === "-" ||
    s === "—" ||
    s === "–" ||
    s === "n/a" ||
    s === "na" ||
    s === "nan" ||
    s === "null" ||
    s === "undefined"
  );
}
// --- Multi-value tokenization and matching helpers ---
function splitMultiTokens(v: any) {
  const s = normStr(v);
  if (!s) return [] as string[];
  // HubSpot/CRM exports often separate multi-values with ; or new lines
  const parts = s.split(/[;|\n]+/g).map((x) => x.trim()).filter(Boolean);
  return parts.length ? parts : [s];
}

function multiMatches(v: any, setSel: Set<string>, includeEmpty: boolean) {
  const s = normStr(v);
  if (s === "") return includeEmpty;
  if (setSel.size === 0) return false;

  const tokens = splitMultiTokens(v);
  for (const t of tokens) {
    if (setSel.has(t.toLowerCase())) return true;
  }
  // fallback: match the whole string too
  return setSel.has(s.toLowerCase());
}
function normStr(v: any) {
  return isEmptyish(v) ? "" : safeStr(v);
}
function uniq<T>(arr: T[]) {
  return Array.from(new Set(arr));
}
function pickColumn(headers: string[], candidates: string[]) {
  const lower = headers.map((h) => h.toLowerCase());
  for (const c of candidates) {
    const idx = lower.indexOf(c.toLowerCase());
    if (idx >= 0) return headers[idx];
  }
  return null;
}
function toNumber(v: any): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" && isFinite(v)) return v;
  if (typeof v === "string") {
    const s = v.replace(/\./g, "").replace(",", ".").trim();
    if (s === "") return null;
    const n = Number(s);
    return isFinite(n) ? n : null;
  }
  return null;
}
function formatCompact(n: number) {
  try {
    return new Intl.NumberFormat(undefined, { notation: "compact", maximumFractionDigits: 1 }).format(n);
  } catch {
    return String(n);
  }
}
function classNames(...xs: Array<string | false | null | undefined>) {
  return xs.filter(Boolean).join(" ");
}
function inferColType(rows: Row[], col: string): ColType {
  const sample = rows
    .map((r) => r[col])
    .filter((v) => v !== null && v !== undefined && String(v).trim() !== "")
    .slice(0, 80);
  if (sample.length === 0) return "unknown";

  const boolCount = sample.filter((v) => typeof v === "boolean").length;
  if (boolCount / sample.length > 0.85) return "boolean";

  const numCount = sample.filter((v) => typeof v === "number" && isFinite(v)).length;
  const numStrCount = sample.filter((v) => {
    if (typeof v !== "string") return false;
    const s = v.replace(/\./g, "").replace(",", ".").trim();
    return s !== "" && isFinite(Number(s));
  }).length;

  if ((numCount + numStrCount) / sample.length > 0.8) return "number";
  const strCount = sample.filter((v) => typeof v === "string").length;
  if (strCount / sample.length > 0.5) return "string";
  return "unknown";
}

/** ---------- Supplier unification ---------- **/
const LEGAL_SUFFIXES = new Set([
  "sa", "s.a", "s.a.", "sl", "s.l", "s.l.", "srl", "s.r.l", "s.r.l.",
  "ltd", "ltd.", "limited", "inc", "inc.", "llc", "gmbh", "bv", "b.v", "b.v.",
  "ag", "spa", "s.p.a", "s.p.a.", "co", "company", "corp", "corporation",
  // common variants
  "private", "pvt", "pvt.", "pte", "pte.", "plc", "group", "holding", "holdings",
]);
/** ---------- Smart Excel/CSV parsing (detect header row) ---------- */
function makeUniqueHeaders(headers: string[]) {
  const seen = new Map<string, number>();
  return headers.map((h, i) => {
    const base = normalizeHeader(h) || `Column ${i + 1}`;
    const key = base.toLowerCase();
    const n = (seen.get(key) ?? 0) + 1;
    seen.set(key, n);
    return n === 1 ? base : `${base} (${n})`;
  });
}

function scoreHeaderRow(row: any[]) {
  const tokens = [
    "buyer", "brand", "customer",
    "supplier", "vendor", "factory", "name",
    "country", "year", "turnover", "revenue", "sales", "amount",
    "sector", "industry",
    "status", "program",
  ];
  let score = 0;
  for (const cell of row) {
    if (!cell) continue;
    const s = String(cell).trim().toLowerCase();
    if (!s) continue;
    for (const t of tokens) {
      if (s === t || s.includes(t)) {
        score += 1;
        break;
      }
    }
  }
  return score;
}

function sheetToRowsSmart(ws: XLSX.WorkSheet) {
  // Get a matrix of rows (arrays). Keeps blank rows=false for a cleaner scan.
  const matrix: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });

  // Find best header row among first ~30 lines.
  let bestIdx = 0;
  let bestScore = -1;
  const scanLimit = Math.min(matrix.length, 30);
  for (let i = 0; i < scanLimit; i++) {
    const row = matrix[i] ?? [];
    // Skip rows that are basically empty (use safeStr/isEmptyish)
    const nonEmpty = row.some((v) => {
      const s = safeStr(v);
      return s !== "" && !isEmptyish(s);
    });
    if (!nonEmpty) continue;
    const s = scoreHeaderRow(row);
    if (s > bestScore) {
      bestScore = s;
      bestIdx = i;
    }
  }

  // If we didn't find a convincing header, fall back to first non-empty row.
  if (bestScore < 2) {
    bestIdx = 0;
    for (let i = 0; i < scanLimit; i++) {
      const row = matrix[i] ?? [];
      const nonEmpty = row.some((v) => {
        const s = safeStr(v);
        return s !== "" && !isEmptyish(s);
      });
      if (nonEmpty) {
        bestIdx = i;
        break;
      }
    }
  }

  const rawHeaders = (matrix[bestIdx] ?? []).map((v) => (v == null ? "" : String(v)));
  const headers = makeUniqueHeaders(rawHeaders);

  const outRows: Row[] = [];
  for (let r = bestIdx + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const nonEmpty = row.some((v) => {
      const s = safeStr(v);
      return s !== "" && !isEmptyish(s);
    });
    if (!nonEmpty) continue;

    const obj: Row = {};
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      obj[h] = row[c] ?? null;
    }

    // ✅ Definitive: skip rows that are visually empty (incl. NBSP/zero-width/NaN/-/N/A)
    const filledHeaders = headers.filter((h) => normStr(obj[h]) !== "");
    if (filledHeaders.length === 0) continue;

    // ✅ Skip garbage rows that only contain IDs (Record ID, *ID, *IDs, etc.)
    const meaningful = filledHeaders.filter((h) => {
      const hn = h.toLowerCase().replace(/\s+/g, "");
      if (hn === "id") return false;
      if (hn.includes("recordid")) return false;
      if (hn.endsWith("id")) return false;
      if (hn.endsWith("ids")) return false;
      return true;
    });
    if (meaningful.length === 0) continue;

    outRows.push(obj);
  }

  return { headers, rows: outRows };
}

function stripDiacritics(s: string) {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/** Para agrupar exacto (Denim-e / Denime / denime -> denime) */
function canonicalSupplierKey(name: string) {
  const raw = stripDiacritics(name.toLowerCase());
  const cleaned = raw
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const tokens = cleaned
    .split(" ")
    .filter(Boolean)
    .filter((t) => !LEGAL_SUFFIXES.has(t));

  return tokens.join(""); // sin espacios
}

/** Para fuzzy: string "suave" con espacios (mejora similitud) */
function canonicalSoft(name: string) {
  const raw = stripDiacritics(name.toLowerCase());
  const cleaned = raw
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const tokens = cleaned
    .split(" ")
    .filter(Boolean)
    .filter((t) => !LEGAL_SUFFIXES.has(t));

  return tokens.join(" ");
}

/** Levenshtein similarity 0..1 */
function levenshtein(a: string, b: string) {
  const al = a.length, bl = b.length;
  if (a === b) return 0;
  if (al === 0) return bl;
  if (bl === 0) return al;

  const v0 = new Array(bl + 1).fill(0);
  const v1 = new Array(bl + 1).fill(0);

  for (let i = 0; i <= bl; i++) v0[i] = i;

  for (let i = 0; i < al; i++) {
    v1[0] = i + 1;
    for (let j = 0; j < bl; j++) {
      const cost = a[i] === b[j] ? 0 : 1;
      v1[j + 1] = Math.min(
        v1[j] + 1,      // insertion
        v0[j + 1] + 1,  // deletion
        v0[j] + cost    // substitution
      );
    }
    for (let j = 0; j <= bl; j++) v0[j] = v1[j];
  }
  return v1[bl];
}
function similarity01(a: string, b: string) {
  const A = a.trim();
  const B = b.trim();
  const maxLen = Math.max(A.length, B.length);
  if (maxLen === 0) return 1;
  const dist = levenshtein(A, B);
  return 1 - dist / maxLen;
}

function buildSupplierGroups(rows: Row[], supplierCol: string | null): SupplierGroup[] {
  if (!supplierCol) return [];

  const counts = new Map<string, Map<string, number>>(); // key -> (variant -> count)
  for (const r of rows) {
    const raw = safeStr(r[supplierCol]);
    if (!raw) continue;
    const key = canonicalSupplierKey(raw);
    if (!key) continue;
    if (!counts.has(key)) counts.set(key, new Map());
    const m = counts.get(key)!;
    m.set(raw, (m.get(raw) ?? 0) + 1);
  }

  const groups: SupplierGroup[] = [];
  for (const [key, variantMap] of counts.entries()) {
    const variants: SupplierVariant[] = Array.from(variantMap.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);

    if (variants.length <= 1) continue;

    const suggestedDisplay = variants[0]?.name ?? key;
    groups.push({
      key,
      variants,
      suggestedDisplay,
      chosenDisplay: suggestedDisplay,
      merge: false,
    });
  }

  groups.sort((a, b) => {
    const ta = a.variants.reduce((s, v) => s + v.count, 0);
    const tb = b.variants.reduce((s, v) => s + v.count, 0);
    return tb - ta;
  });

  return groups;
}

/** AliasMap: variant -> display elegido (solo grupos con merge=true) */
function buildSupplierAliasMap(groups: SupplierGroup[], fuzzyApproved: Record<string, string>) {
  const map: Record<string, string> = { ...fuzzyApproved };
  for (const g of groups) {
    if (!g.merge) continue;
    for (const v of g.variants) map[v.name] = g.chosenDisplay;
  }
  return map;
}

/**
 * Fuzzy suggestions:
 * - Tomamos un set de nombres distintos
 * - Comparamos soft strings
 * - Si similitud >= umbral y no están ya en misma canonicalKey, sugerimos
 * - Limitamos por performance (top N)
 */
function buildFuzzySuggestions(rows: Row[], supplierCol: string | null, maxPairs = 40) {
  if (!supplierCol) return [];

  const rawNames = uniq(rows.map((r) => safeStr(r[supplierCol])).filter(Boolean));
  // para performance: nos quedamos con los más frecuentes primero
  const freq = new Map<string, number>();
  for (const r of rows) {
    const s = safeStr(r[supplierCol]);
    if (!s) continue;
    freq.set(s, (freq.get(s) ?? 0) + 1);
  }
  rawNames.sort((a, b) => (freq.get(b) ?? 0) - (freq.get(a) ?? 0));

  const limitNames = rawNames.slice(0, 250); // cap por performance
  const soft = new Map<string, string>();
  const key = new Map<string, string>();
  for (const n of limitNames) {
    soft.set(n, canonicalSoft(n));
    key.set(n, canonicalSupplierKey(n));
  }

  const suggestions: FuzzySuggestion[] = [];
  const threshold = 0.92;

  for (let i = 0; i < limitNames.length; i++) {
    for (let j = i + 1; j < limitNames.length; j++) {
      const a = limitNames[i];
      const b = limitNames[j];
      // si ya caen en la misma key exacta, no lo sugerimos como fuzzy (ya lo maneja grupo)
      if (key.get(a) === key.get(b)) continue;

      const sa = soft.get(a) ?? "";
      const sb = soft.get(b) ?? "";
      if (!sa || !sb) continue;

      // quick pruning por longitud
      const lenRatio = Math.min(sa.length, sb.length) / Math.max(sa.length, sb.length);
      if (lenRatio < 0.7) continue;

      const score = similarity01(sa, sb);
      if (score >= threshold) {
        // display recomendado: el más frecuente entre a y b
        const chosenDisplay = (freq.get(a) ?? 0) >= (freq.get(b) ?? 0) ? a : b;
        suggestions.push({ a, b, score, chosenDisplay, approved: null });
      }
    }
  }

  suggestions.sort((x, y) => y.score - x.score);
  return suggestions.slice(0, maxPairs);
}

/** ---------- UI bits ---------- */
function Pill({ children }: { children: React.ReactNode }) {
  return (
    <span className="inline-flex items-center rounded-full border border-slate-200 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 shadow-sm">
      {children}
    </span>
  );
}
function Card({ title, right, children }: { title: string; right?: React.ReactNode; children: React.ReactNode }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white shadow-sm">
      <div className="flex items-center justify-between border-b border-slate-100 px-5 py-4">
        <div className="text-sm font-semibold text-slate-800">{title}</div>
        {right ? <div className="text-xs text-slate-500">{right}</div> : null}
      </div>
      <div className="p-5">{children}</div>
    </div>
  );
}
function EmptyState({ text }: { text: string }) {
  return (
    <div className="rounded-2xl border border-dashed border-slate-300 bg-white p-10 text-center text-sm text-slate-600">
      {text}
    </div>
  );
}

function MultiSelect({
  label,
  options,
  selected,
  onChange,
  placeholder = "Seleccionar…",
  searchable = true,
}: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (v: string[]) => void;
  placeholder?: string;
  searchable?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const ref = useRef<HTMLDivElement | null>(null);

  const filteredOptions = useMemo(() => {
    const qq = q.trim().toLowerCase();
    if (!searchable || !qq) return options;
    return options.filter((o) => o.toLowerCase().includes(qq));
  }, [options, q, searchable]);

  useEffect(() => {
    if (!open) setQ("");
  }, [open]);

  useEffect(() => {
    const onDoc = (e: MouseEvent) => {
      if (!ref.current) return;
      if (!ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  const toggle = (o: string) => {
    if (selected.includes(o)) onChange(selected.filter((x) => x !== o));
    else onChange([...selected, o]);
  };

  return (
    <div className="relative" ref={ref}>
      {label !== "" ? <div className="mb-1 text-xs font-medium text-slate-700">{label}</div> : null}
      <button
        type="button"
        onClick={() => setOpen((s) => !s)}
        className="flex w-full items-center justify-between rounded-xl border border-slate-200 bg-white px-3 py-2 text-left text-sm text-slate-800 shadow-sm hover:border-slate-300"
      >
        <span className={classNames(selected.length ? "text-slate-900" : "text-slate-400")}>
          {selected.length ? `${selected.length} seleccionado(s)` : placeholder}
        </span>
        <span className="ml-3 text-slate-400">▾</span>
      </button>

      {open ? (
        <div className="absolute z-50 mt-2 w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
          {searchable ? (
            <div className="border-b border-slate-100 p-2">
              <input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder="Buscar…"
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-300 focus:outline-none"
              />
            </div>
          ) : null}

          <div className="max-h-64 overflow-auto p-2">
            {filteredOptions.length === 0 ? (
              <div className="p-3 text-sm text-slate-500">Sin resultados</div>
            ) : (
              filteredOptions.map((o) => (
                <label key={o} className="flex cursor-pointer items-center gap-2 rounded-lg px-2 py-2 text-sm hover:bg-slate-50">
                  <input
                    type="checkbox"
                    checked={selected.includes(o)}
                    onChange={() => toggle(o)}
                    className="h-4 w-4 rounded border-slate-300"
                  />
                  <span className="truncate text-slate-800">{o}</span>
                </label>
              ))
            )}
          </div>
          <div className="flex items-center justify-between border-t border-slate-100 p-2">
            <button
              type="button"
              className="rounded-lg px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
              onClick={() => onChange([])}
            >
              Limpiar
            </button>
            <button
              type="button"
              className="rounded-lg bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-800"
              onClick={() => setOpen(false)}
            >
              Listo
            </button>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function RangeFilter({
  label,
  min,
  max,
  value,
  onChange,
}: {
  label: string;
  min: number;
  max: number;
  value: { from: number | null; to: number | null };
  onChange: (v: { from: number | null; to: number | null }) => void;
}) {
  return (
    <div>
      {label !== "" ? <div className="mb-1 text-xs font-medium text-slate-700">{label}</div> : null}
      <div className="grid grid-cols-2 gap-2">
        <input
          type="number"
          placeholder={String(min)}
          value={value.from ?? ""}
          onChange={(e) => onChange({ ...value, from: e.target.value === "" ? null : Number(e.target.value) })}
          className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-300 focus:outline-none"
        />
        <input
          type="number"
          placeholder={String(max)}
          value={value.to ?? ""}
          onChange={(e) => onChange({ ...value, to: e.target.value === "" ? null : Number(e.target.value) })}
          className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-300 focus:outline-none"
        />
      </div>
      <div className="mt-1 text-[11px] text-slate-500">
        Min: {formatCompact(min)} · Max: {formatCompact(max)}
      </div>
    </div>
  );
}

/** ---------- Main ---------- */
export default function DashboardSuppliers() {
  const [rows, setRows] = useState<Row[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [colTypes, setColTypes] = useState<Record<string, ColType>>({});
  const [filesLoaded, setFilesLoaded] = useState<number>(0);

  const [mapping, setMapping] = useState<FieldMapping>({
    Buyer: null,
    Supplier: null,
    Country: null,
    Year: null,
    Turnover: null,
    Sector: null,
  });

  const [fixedFilters, setFixedFilters] = useState<{
    Buyer: string[];
    Supplier: string[]; // sobre Supplier UNIFICADO si está ON
    Country: string[];
    Year: string[];
    Sector: string[];
    Turnover: { from: number | null; to: number | null };
  }>({
    Buyer: [],
    Supplier: [],
    Country: [],
    Year: [],
    Sector: [],
    Turnover: { from: null, to: null },
  });

  const [dynamicFilters, setDynamicFilters] = useState<DynamicFilter[]>([]);

  // Unificación supplier
  const [supplierGroups, setSupplierGroups] = useState<SupplierGroup[]>([]);
  const [supplierAliasMap, setSupplierAliasMap] = useState<Record<string, string>>({});
  const [supplierUnifyEnabled, setSupplierUnifyEnabled] = useState(true);

  // UPGRADE 1: search
  const [supplierGroupSearch, setSupplierGroupSearch] = useState("");

  // UX: hide resolved merges by default
  const [showResolvedReconciliation, setShowResolvedReconciliation] = useState(false);
  // UX: reconciliation compact/expand
  const [reconcileOpen, setReconcileOpen] = useState(false);

  // UPGRADE 3: fuzzy suggestions + aprobaciones (manual aliases)
  const [fuzzySuggestions, setFuzzySuggestions] = useState<FuzzySuggestion[]>([]);
  const [fuzzyApprovedAliases, setFuzzyApprovedAliases] = useState<Record<string, string>>({});

  // Persisted user decisions
  const [supplierGroupOverrides, setSupplierGroupOverrides] = useState<Record<string, SupplierGroupOverride>>({});

  // Hydration guard
  const [hydrated, setHydrated] = useState(false);
  const [hasSavedSession, setHasSavedSession] = useState(false);

  // UI feedback
  const [toast, setToast] = useState<string | null>(null);
  const toastTimerRef = useRef<number | null>(null);

  // Preview pagination (table)
  const [previewPage, setPreviewPage] = useState(1);
  const [previewPageSize, setPreviewPageSize] = useState(30);
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);

  function showToast(msg: string) {
    setToast(msg);
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current);
    toastTimerRef.current = window.setTimeout(() => setToast(null), 1800);
  }

  // Keep latest values to avoid stale closures in click handlers
  const supplierGroupsRef = useRef<SupplierGroup[]>([]);
  const fuzzySuggestionsRef = useRef<FuzzySuggestion[]>([]);
  const fuzzyApprovedAliasesRef = useRef<Record<string, string>>({});
  const supplierGroupOverridesRef = useRef<Record<string, SupplierGroupOverride>>({});

  // IMPORTANT: keep refs in sync during render so buttons work immediately after loading data
  supplierGroupsRef.current = supplierGroups;
  fuzzySuggestionsRef.current = fuzzySuggestions;
  fuzzyApprovedAliasesRef.current = fuzzyApprovedAliases;
  supplierGroupOverridesRef.current = supplierGroupOverrides;

  // Auto-apply supplier alias mapping whenever the user changes merge toggles or chosen display
  useEffect(() => {
    setSupplierAliasMap(buildSupplierAliasMap(supplierGroups, fuzzyApprovedAliases));
  }, [supplierGroups, fuzzyApprovedAliases]);


  // --- IndexedDB: Save/reset helpers ---
  const saveTimerRef = useRef<number | null>(null);

  async function clearSavedSession() {
    await idbDel(IDB_KEY);
    setHasSavedSession(false);

    // Reset app state
    setRows([]);
    setHeaders([]);
    setSelectedColumns([]);
    setColTypes({});
    setFilesLoaded(0);
    setMapping({ Buyer: null, Supplier: null, Country: null, Year: null, Turnover: null, Sector: null });
    setFixedFilters({ Buyer: [], Supplier: [], Country: [], Year: [], Sector: [], Turnover: { from: null, to: null } });
    setDynamicFilters([]);
    setSupplierGroups([]);
    setSupplierAliasMap({});
    setSupplierUnifyEnabled(true);
    setSupplierGroupOverrides({});
    setFuzzySuggestions([]);
    setFuzzyApprovedAliases({});

    showToast("Sesión borrada ✅");
  }

  
  function applyOverridesToGroups(
    groups: SupplierGroup[],
    overrides?: Record<string, SupplierGroupOverride>
  ) {
    const ovs = overrides ?? supplierGroupOverridesRef.current;
    return groups.map((g) => {
      const ov = ovs[g.key];
      if (!ov) return g;
      return { ...g, merge: ov.merge, chosenDisplay: ov.chosenDisplay };
    });
  }

  function recomputeSupplierStuff(
    allRows: Row[],
    supplierCol: string | null,
    overrides?: Record<string, SupplierGroupOverride>,
    fuzzyApproved?: Record<string, string>
  ) {
    const rawGroups = buildSupplierGroups(allRows, supplierCol);
    const groups = applyOverridesToGroups(rawGroups, overrides);
    setSupplierGroups(groups);

    const fuzzy = buildFuzzySuggestions(allRows, supplierCol, 40);
    setFuzzySuggestions(fuzzy);

    const fuzzyMap = fuzzyApproved ?? fuzzyApprovedAliasesRef.current;

    // Alias final (grupos + fuzzy aprobados)
    setSupplierAliasMap(buildSupplierAliasMap(groups, fuzzyMap));
    setSupplierUnifyEnabled(true);
  }

  function supplierDisplay(raw: string) {
    if (!supplierUnifyEnabled) return raw;
    return supplierAliasMap[raw] ?? raw;
  }

  async function loadFiles(fileList: FileList) {
    const incomingRows: Row[] = [];
    let incomingHeaders: string[] = [];

    for (const file of Array.from(fileList)) {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];

      const parsed = sheetToRowsSmart(ws);
      const hs = parsed.headers.map(normalizeHeader).filter(Boolean);
      incomingHeaders = uniq([...incomingHeaders, ...hs]).filter(Boolean);

      const normalizedRows: Row[] = parsed.rows.map((r) => {
        const nr: Row = {};
        for (const k of Object.keys(r)) nr[normalizeHeader(k)] = r[k];
        return nr;
      });

      incomingRows.push(...normalizedRows);
    }

    const merged = [...rows, ...incomingRows];
    const headerUnion = uniq([...headers, ...incomingHeaders]).filter((h) => h);

    const types: Record<string, ColType> = {};
    for (const h of headerUnion) types[h] = inferColType(merged, h);

    // Only keep previous mapping if the column exists in the new headerUnion, otherwise re-detect
    const keepIfExists = (col: string | null) => (col && headerUnion.includes(col) ? col : null);

    const detected: FieldMapping = {
      Buyer:
        keepIfExists(mapping.Buyer) ??
        pickColumn(headerUnion, ["Buyer", "buyer", "Brand", "brand", "Customer", "customer"]),
      Supplier:
        keepIfExists(mapping.Supplier) ??
        pickColumn(headerUnion, [
          "Supplier",
          "supplier",
          "Vendor",
          "vendor",
          "Factory",
          "factory",
          "name",
          "Name",
        ]),
      Country:
        keepIfExists(mapping.Country) ??
        pickColumn(headerUnion, ["Country", "country", "CountryName", "country_name", "country name"]),
      Year:
        keepIfExists(mapping.Year) ??
        pickColumn(headerUnion, ["Year", "year", "Año", "anio", "FY", "fiscal_year"]),
      Turnover:
        keepIfExists(mapping.Turnover) ??
        pickColumn(headerUnion, ["Turnover", "turnover", "Revenue", "revenue", "Sales", "sales", "Amount", "amount"]),
      Sector:
        keepIfExists(mapping.Sector) ??
        pickColumn(headerUnion, ["Sector", "sector", "Industry", "industry"]),
    };

    setRows(merged);
    setHeaders(headerUnion);
    setColTypes(types);
    // Safety reset: if the supplier column changed, clear supplier-related filters
    const supplierChanged = mapping.Supplier !== detected.Supplier;
    setMapping(detected);
    setFilesLoaded((n) => n + fileList.length);
    setSelectedColumns([]);

    // Si cambió la columna Supplier (p.ej. de "Supplier" a "Name"), limpiamos selección para evitar filtros inválidos
    if (supplierChanged) {
      setFixedFilters((s) => ({ ...s, Supplier: [] }));
    }

    // recalcular conciliación total
    recomputeSupplierStuff(merged, detected.Supplier);

    // Reset filtros (dejamos Buyer si único)
    setDynamicFilters([]);
    setFixedFilters((s) => ({
      ...s,
      Supplier: [],
      Country: [],
      Year: [],
      Sector: [],
      Turnover: { from: null, to: null },
    }));
    showToast(`Cargado ✅ ${fileList.length} archivo(s)`);
  }

  const turnoverRange = useMemo(() => {
    const col = mapping.Turnover;
    if (!col) return { min: 0, max: 0, ok: false };
    const nums = rows.map((r) => toNumber(r[col])).filter((n): n is number => n !== null && isFinite(n));
    if (nums.length === 0) return { min: 0, max: 0, ok: false };
    return { min: Math.min(...nums), max: Math.max(...nums), ok: true };
  }, [rows, mapping]);

  function makeDynamicFilter(column: string): DynamicFilter {
    const t = colTypes[column] ?? "unknown";

    // numeric range
    if (t === "number") {
      const nums = rows.map((r) => toNumber(r[column])).filter((n): n is number => n !== null);
      const min = nums.length ? Math.min(...nums) : 0;
      const max = nums.length ? Math.max(...nums) : 0;
      return { id: column, column, kind: "range", value: { from: null, to: null, min, max } };
    }

    // boolean
    if (t === "boolean") {
      return { id: column, column, kind: "boolean", value: null };
    }

    // string/unknown: choose multi vs text depending on cardinality
    // Use token cardinality for multi-value cells
    const uniqVals = uniq(rows.flatMap((r) => splitMultiTokens(r[column])));
    if (uniqVals.length > 300) {
      return { id: column, column, kind: "text", value: "" };
    }

    return { id: column, column, kind: "multi", value: [] };
  }

  function syncDynamicFilters(selectedCols: string[]) {
    setDynamicFilters((prev) => {
      const prevMap = new Map(prev.map((f) => [f.column, f] as const));
      return selectedCols.map((c) => prevMap.get(c) ?? makeDynamicFilter(c));
    });
  }

  function removeDynamicFilter(col: string) {
    setDynamicFilters((fs) => fs.filter((f) => f.column !== col));
  }

  // Dynamic options
  const dynamicOptions = useMemo(() => {
    const out: Record<string, string[]> = {};
    for (const f of dynamicFilters) {
      if (f.kind !== "multi") continue;
      const hasEmpty = rows.some((r) => normStr(r[f.column]) === "");
      const tokens = rows.flatMap((r) => splitMultiTokens(r[f.column]));
      const base = uniq(tokens)
        .sort((a, b) => a.localeCompare(b))
        .slice(0, 5000);
      out[f.column] = hasEmpty ? [EMPTY_OPTION, ...base] : base;
    }
    return out;
  }, [dynamicFilters, rows]);

  // Apply filters
  const filteredRows = useMemo(() => {
    let out = rows;

    (["Buyer", "Country", "Year", "Sector"] as const).forEach((k) => {
      const col = mapping[k];
      const selected = fixedFilters[k];
      if (!col || selected.length === 0) return;
      const includeEmpty = selected.includes(EMPTY_OPTION);
      const setSel = new Set(selected.filter((x) => x !== EMPTY_OPTION).map((x) => x.toLowerCase()));
      out = out.filter((r) => multiMatches(r[col], setSel, includeEmpty));
    });

    // Supplier (unificado)
    if (mapping.Supplier && fixedFilters.Supplier.length > 0) {
      const col = mapping.Supplier;
      const includeEmpty = fixedFilters.Supplier.includes(EMPTY_OPTION);
      const setSel = new Set(
        fixedFilters.Supplier.filter((x) => x !== EMPTY_OPTION).map((x) => x.toLowerCase())
      );
      out = out.filter((r) => {
        const v = supplierDisplay(normStr(r[col]));
        if (v === "") return includeEmpty;
        if (setSel.size === 0) return false;
        return setSel.has(v.toLowerCase());
      });
    }

    // Turnover
    const turnCol = mapping.Turnover;
    if (turnCol && turnoverRange.ok) {
      const { from, to } = fixedFilters.Turnover;
      if (from !== null || to !== null) {
        out = out.filter((r) => {
          const n = toNumber(r[turnCol]);
          if (n === null) return false;
          if (from !== null && n < from) return false;
          if (to !== null && n > to) return false;
          return true;
        });
      }
    }

    // Dynamic filters
    for (const f of dynamicFilters) {
      const col = f.column;

      if (f.kind === "multi") {
        const selected: string[] = f.value ?? [];
        if (!selected.length) continue;
        const includeEmpty = selected.includes(EMPTY_OPTION);
        const setSel = new Set(selected.filter((x) => x !== EMPTY_OPTION).map((x) => x.toLowerCase()));
        out = out.filter((r) => multiMatches(r[col], setSel, includeEmpty));
      }

      if (f.kind === "range") {
        const { from, to } = f.value ?? {};
        if (from == null && to == null) continue;
        out = out.filter((r) => {
          const n = toNumber(r[col]);
          if (n === null) return false;
          if (from != null && n < from) return false;
          if (to != null && n > to) return false;
          return true;
        });
      }

      if (f.kind === "boolean") {
        const v = f.value;
        if (v === null) continue;
        out = out.filter((r) => Boolean(r[col]) === v);
      }

      if (f.kind === "text") {
        const q = String(f.value ?? "").trim().toLowerCase();
        if (!q) continue;
        out = out.filter((r) => normStr(r[col]).toLowerCase().includes(q));
      }
    }

    return out;
  }, [rows, mapping, fixedFilters, dynamicFilters, turnoverRange, supplierAliasMap, supplierUnifyEnabled]);

  // Reset page when filters change
  useEffect(() => {
    setPreviewPage(1);
  }, [fixedFilters, dynamicFilters]);

  const previewTotalPages = useMemo(() => {
    return Math.max(1, Math.ceil(filteredRows.length / previewPageSize));
  }, [filteredRows.length, previewPageSize]);

  useEffect(() => {
  setPreviewPage(1);
}, [fixedFilters, dynamicFilters, selectedColumns]);

  const previewFrom = (previewPage - 1) * previewPageSize;
  const previewTo = Math.min(filteredRows.length, previewFrom + previewPageSize);
  const previewHeaders = useMemo(
  () => selectedColumns.filter((c) => headers.includes(c)),
  [selectedColumns, headers]
  );
  const previewRows = filteredRows.slice(previewFrom, previewTo);

  // KPIs + aggregations
  const kpis = useMemo(() => {
    const total = filteredRows.length;
    const supplierCol = mapping.Supplier;
    const countryCol = mapping.Country;
    const yearCol = mapping.Year;
    const turnCol = mapping.Turnover;

    const uniqueSuppliers = supplierCol
      ? uniq(filteredRows.map((r) => supplierDisplay(safeStr(r[supplierCol])).toLowerCase()).filter(Boolean)).length
      : null;

    const uniqueCountries = countryCol
      ? uniq(filteredRows.map((r) => safeStr(r[countryCol]).toLowerCase()).filter(Boolean)).length
      : null;

    let turnoverSum: number | null = null;
    let turnoverAvg: number | null = null;
    if (turnCol) {
      const nums = filteredRows.map((r) => toNumber(r[turnCol])).filter((n): n is number => n !== null);
      if (nums.length) {
        turnoverSum = nums.reduce((a, b) => a + b, 0);
        turnoverAvg = turnoverSum / nums.length;
      }
    }

    let byCountry: Array<{ name: string; value: number }> = [];
    if (countryCol) {
      const m = new Map<string, number>();
      for (const r of filteredRows) {
        const c = safeStr(r[countryCol]) || "N/A";
        m.set(c, (m.get(c) ?? 0) + 1);
      }
      byCountry = Array.from(m.entries())
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 12);
    }

    let byYear: Array<{ year: string; value: number; turnover?: number }> = [];
    if (yearCol) {
      const m = new Map<string, { c: number; t: number }>();
      for (const r of filteredRows) {
        const y = safeStr(r[yearCol]) || "N/A";
        const cur = m.get(y) ?? { c: 0, t: 0 };
        cur.c += 1;
        if (turnCol) {
          const n = toNumber(r[turnCol]);
          if (n !== null) cur.t += n;
        }
        m.set(y, cur);
      }
      byYear = Array.from(m.entries())
        .map(([year, o]) => ({ year, value: o.c, turnover: turnCol ? o.t : undefined }))
        .sort((a, b) => a.year.localeCompare(b.year))
        .slice(-20);
    }

    return { total, uniqueSuppliers, uniqueCountries, turnoverSum, turnoverAvg, byCountry, byYear };
  }, [filteredRows, mapping, supplierAliasMap, supplierUnifyEnabled]);

  // Export filtered results to Excel
  function exportFilteredToExcel() {
    if (!filteredRows.length) return;

    const supplierCol = mapping.Supplier;

    // Armamos exportRows: si unify ON y hay Supplier, agregamos columna "Supplier (Unified)"
    const exportRows = filteredRows.map((r) => {
      const obj: Row = { ...r };
      if (supplierCol) {
        const raw = safeStr(r[supplierCol]);
        obj["Supplier (Unified)"] = supplierDisplay(raw);
      }
      return obj;
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Filtered");

    const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    XLSX.writeFile(wb, `filtered-results-${ts}.xlsx`);
  }

  const hasData = rows.length > 0;

  // Conciliación: apply helpers
  function applyAliasesNow() {
    // Ensure user sees the effect immediately
    setSupplierUnifyEnabled(true);

    // Avoid getting stuck with old Supplier selections
    setFixedFilters((s) => ({ ...s, Supplier: [] }));

    showToast("Cambios aplicados ✅");
  }

  // UPGRADE 2: apply all suggested merges (and fuzzy suggestions)
  function applyAllSuggestedMerges() {
    // (A) Exact groups
    const base = supplierGroupsRef.current;
    const updated = base.map((g) => ({
      ...g,
      merge: true,
      chosenDisplay: g.suggestedDisplay,
    }));

    setSupplierGroups(updated);
    setSupplierGroupOverrides(() => {
      const next: Record<string, SupplierGroupOverride> = {};
      for (const g of updated) next[g.key] = { merge: true, chosenDisplay: g.chosenDisplay };
      return next;
    });

    // (B) Fuzzy suggestions (auto-approve all pending)
    const fuzzies = fuzzySuggestionsRef.current;
    const pending = fuzzies.filter((s) => s.approved === null);
    if (pending.length) {
      const aliasUpdates: Record<string, string> = { ...fuzzyApprovedAliasesRef.current };
      for (const s of pending) {
        aliasUpdates[s.a] = s.chosenDisplay;
        aliasUpdates[s.b] = s.chosenDisplay;
      }
      setFuzzyApprovedAliases(aliasUpdates);
      setFuzzySuggestions((prev) => prev.map((s) => (s.approved === null ? { ...s, approved: true } : s)));
    }

    setSupplierUnifyEnabled(true);
    setFixedFilters((s) => ({ ...s, Supplier: [] }));

    showToast(`Aplicado ✅ ${updated.length} grupos + ${pending.length} fuzzy`);
  }
  // --- IndexedDB hydration on mount ---
  useEffect(() => {
    (async () => {
      try {
        const snap = await idbGet<PersistedStateV1>(IDB_KEY);
        if (!snap || snap.version !== 1) {
          setHasSavedSession(false);
          setHydrated(true);
          return;
        }

        setHasSavedSession(true);
        setRows(snap.rows ?? []);
        setHeaders(snap.headers ?? []);
        setColTypes(snap.colTypes ?? {});
        setFilesLoaded(snap.filesLoaded ?? 0);
        setMapping(snap.mapping ?? { Buyer: null, Supplier: null, Country: null, Year: null, Turnover: null, Sector: null });
        setFixedFilters(snap.fixedFilters ?? { Buyer: [], Supplier: [], Country: [], Year: [], Sector: [], Turnover: { from: null, to: null } });
        setDynamicFilters(snap.dynamicFilters ?? []);
        setSupplierUnifyEnabled(snap.supplierUnifyEnabled ?? true);
        setSupplierGroupOverrides(snap.supplierGroupOverrides ?? {});
        setFuzzyApprovedAliases(snap.fuzzyApprovedAliases ?? {});

        // Rebuild groups + fuzzy from rows and apply overrides
        const supplierCol = (snap.mapping?.Supplier ?? null) as string | null;
        if (snap.rows && snap.rows.length) {
          setTimeout(() => {
            recomputeSupplierStuff(
              snap.rows,
              supplierCol,
              snap.supplierGroupOverrides ?? {},
              snap.fuzzyApprovedAliases ?? {}
            );
          }, 0);
        }

        showToast("Sesión restaurada ✅");
      } catch {
        // ignore restore errors
      } finally {
        setHydrated(true);
      }
    })();
  }, []);

  // --- IndexedDB autosave (debounced) ---
  useEffect(() => {
    if (!hydrated) return;

    // Only save if we have data
    if (!rows.length) {
      // If user cleared data, also clear saved snapshot
      setHasSavedSession(false);
      return;
    }

    const snapshot: PersistedStateV1 = {
      version: 1,
      savedAt: Date.now(),
      rows,
      headers,
      colTypes,
      filesLoaded,
      mapping,
      fixedFilters,
      dynamicFilters,
      supplierUnifyEnabled,
      supplierGroupOverrides,
      fuzzyApprovedAliases,
    };

    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(() => {
      idbSet(IDB_KEY, snapshot)
        .then(() => setHasSavedSession(true))
        .catch(() => {});
    }, 350);
  }, [
    hydrated,
    rows,
    headers,
    colTypes,
    filesLoaded,
    mapping,
    fixedFilters,
    dynamicFilters,
    supplierUnifyEnabled,
    supplierGroupOverrides,
    fuzzyApprovedAliases,
  ]);

  // UPGRADE 3: approve fuzzy suggestion -> crea alias manual (a y b -> chosenDisplay)
  function approveFuzzy(s: FuzzySuggestion, approved: boolean) {
    // Update UI state
    setFuzzySuggestions((prev) =>
      prev.map((x) => (x.a === s.a && x.b === s.b && x.score === s.score ? { ...x, approved } : x))
    );

    if (!approved) {
      showToast("Marcado como separado ✅");
      return;
    }

    const aliasUpdates: Record<string, string> = {
      ...fuzzyApprovedAliasesRef.current,
      [s.a]: s.chosenDisplay,
      [s.b]: s.chosenDisplay,
    };

    setFuzzyApprovedAliases(aliasUpdates);
    setSupplierUnifyEnabled(true);
    setFixedFilters((s0) => ({ ...s0, Supplier: [] }));
    showToast("Unificación fuzzy aplicada ✅");
  }

  // Recalcular conciliación
  function recalcAll() {
    if (!hasData) return;
    const supplierCol = mapping.Supplier;
    recomputeSupplierStuff(rows, supplierCol);
    setSupplierUnifyEnabled(true);
    setFixedFilters((s) => ({ ...s, Supplier: [] }));
    showToast("Recalculado ✅");
  }

  const reconcileCounts = useMemo(() => {
    const pendingGroups = supplierGroups.filter((g) => !g.merge).length;
    const resolvedGroups = supplierGroups.filter((g) => g.merge).length;
    const pendingFuzzy = fuzzySuggestions.filter((s) => s.approved === null).length;
    const resolvedFuzzy = fuzzySuggestions.filter((s) => s.approved !== null).length;
    return { pendingGroups, resolvedGroups, pendingFuzzy, resolvedFuzzy };
  }, [supplierGroups, fuzzySuggestions]);

  useEffect(() => {
    // Auto-minimize when everything is resolved
    if (reconcileCounts.pendingGroups === 0 && reconcileCounts.pendingFuzzy === 0) {
      setReconcileOpen(false);
    }
  }, [reconcileCounts.pendingGroups, reconcileCounts.pendingFuzzy]);

  const supplierGroupsFiltered = useMemo(() => {
    const q = supplierGroupSearch.trim().toLowerCase();

    const base = showResolvedReconciliation ? supplierGroups : supplierGroups.filter((g) => !g.merge);

    if (!q) return base;

    return base.filter((g) => {
      if (g.chosenDisplay.toLowerCase().includes(q)) return true;
      if (g.suggestedDisplay.toLowerCase().includes(q)) return true;
      return g.variants.some((v) => v.name.toLowerCase().includes(q));
    });
  }, [supplierGroups, supplierGroupSearch, showResolvedReconciliation]);

  return (
    <div className="min-h-screen" style={{ background: BRAND.lightGrey }}>
      {toast ? (
        <div className="pointer-events-none fixed left-1/2 top-4 z-[999] -translate-x-1/2 rounded-xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-lg">
          {toast}
        </div>
      ) : null}
      {/* Top bar */}
      <div className="border-b" style={{ borderColor: BRAND.darkBlue, background: BRAND.navy }}>
        <div className="mx-auto w-full max-w-none px-4 py-12 sm:px-6 lg:px-10 2xl:px-14">
          <div className="grid grid-cols-1 gap-3 md:grid-cols-3 md:items-center">
            <div>
              <div className="text-sm font-semibold text-white/90 tracking-wide">SUPPLIERS DASHBOARD</div>
              <div className="mt-1 text-xs text-white/70">
                Cargá uno o varios CSV/XLSX. Supplier se reconcilia (variantes + fuzzy). Exporta resultados filtrados.
              </div>
            </div>

            {/* Center logo */}
            <div className="flex justify-center">
              <div className="flex items-center justify-center">
                <img
                  src={HEADER_LOGO_SRC}
                  alt="Twinco Capital"
                  className="h-20 md:h-28 lg:h-32 w-auto object-contain"
                />
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-start gap-3 md:justify-end">
              <label
                className="inline-flex cursor-pointer items-center gap-2 rounded-xl border px-3 py-2 text-sm shadow-sm"
                style={{ borderColor: "rgba(255,255,255,0.18)", background: "rgba(255,255,255,0.06)" }}
              >
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  className="hidden"
                  onChange={(e) => {
                    const fl = e.target.files;
                    if (fl && fl.length) loadFiles(fl);
                  }}
                />
                <span className="text-white/90">📁 Cargar archivo(s)</span>
              </label>

              {hasSavedSession ? (
                <button
                  type="button"
                  className="rounded-xl border px-3 py-2 text-sm font-semibold shadow-sm"
                  style={{ borderColor: "rgba(255,255,255,0.18)", background: "rgba(255,255,255,0.06)", color: "rgba(255,255,255,0.9)" }}
                  onClick={clearSavedSession}
                  title="Borra la sesión guardada (datos + decisiones)"
                >
                  🧹 Borrar sesión
                </button>
              ) : null}

              <button
                type="button"
                className="rounded-xl px-3 py-2 text-sm font-semibold disabled:opacity-50"
                style={{ background: BRAND.neon, color: BRAND.navy }}
                disabled={!hasData || filteredRows.length === 0}
                onClick={exportFilteredToExcel}
                title={!hasData ? "Cargá archivos" : filteredRows.length === 0 ? "No hay filas para exportar" : "Exportar filtrado"}
              >
                ⬇ Export Excel (filtrado)
              </button>

              {hasData ? (
                <>
                  <span className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium shadow-sm" style={{ background: "rgba(255,255,255,0.08)", color: "rgba(255,255,255,0.85)", border: "1px solid rgba(255,255,255,0.14)" }}>
                    {formatCompact(filteredRows.length)} / {formatCompact(rows.length)} filas
                  </span>
                  <span className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium shadow-sm" style={{ background: "rgba(255,255,255,0.08)", color: "rgba(255,255,255,0.85)", border: "1px solid rgba(255,255,255,0.14)" }}>
                    {filesLoaded} archivo(s)
                  </span>
                  {hasSavedSession ? (
                    <span className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium shadow-sm" style={{ background: "rgba(255,255,255,0.08)", color: "rgba(255,255,255,0.85)", border: "1px solid rgba(255,255,255,0.14)" }}>
                      Sesión guardada
                    </span>
                  ) : null}
                </>
              ) : (
                <span className="inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium shadow-sm" style={{ background: "rgba(255,255,255,0.08)", color: "rgba(255,255,255,0.85)", border: "1px solid rgba(255,255,255,0.14)" }}>
                  Sin archivo
                </span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Spacer to keep filters lower under the large header */}
      <div className="h-12" style={{ background: BRAND.lightGrey }} />

      {/* Filters */}
      <div className="border-b" style={{ borderColor: BRAND.softGrey, background: "rgba(255,255,255,0.92)", backdropFilter: "blur(8px)" }}>
        <div className="mx-auto w-full max-w-none px-4 py-9 sm:px-6 lg:px-10 2xl:px-14">
          

          <div className="mt-5">
            {!hasData ? (
              <div className="rounded-xl border border-slate-200 bg-white px-3 py-3 text-sm text-slate-500 shadow-sm">
                Cargá un archivo para habilitar los filtros extra por cabecera.
              </div>
            ) : (
              <div className="grid grid-cols-1 gap-3">
                <MultiSelect
                  label="Columnas (y filtros)"
                  options={headers}
                  selected={selectedColumns}
                  onChange={(cols) => {
                    setSelectedColumns(cols);
                    syncDynamicFilters(cols);
                  }}
                  placeholder={headers.length ? "Elegí columnas para ver y filtrar…" : "Cargá un Excel para ver columnas"}
                  searchable
                />
                <div className="text-xs text-slate-600">
                  Elegí una o más cabeceras (ej: “Nombre”). La tabla mostrará SOLO esas columnas y debajo se generan los filtros.
                </div>
              </div>
            )}
          </div>
          {/* Dynamic filters row */}
          <div className="mt-4 flex flex-wrap items-center gap-2">
            {dynamicFilters.map((f) => (
              <div key={f.id} className="w-full md:w-[340px]">
                <div className="flex items-center justify-between">
                  <div className="text-xs font-medium text-slate-700">{f.column}</div>
                  <button
                    type="button"
                    className="rounded-lg px-2 py-1 text-xs font-semibold text-slate-500 hover:bg-slate-100"
                    onClick={() => removeDynamicFilter(f.column)}
                    title="Quitar filtro"
                  >
                    ✕
                  </button>
                </div>

                {f.kind === "multi" ? (
                  <MultiSelect
                    label=""
                    options={dynamicOptions[f.column] ?? []}
                    selected={f.value ?? []}
                    onChange={(v) => setDynamicFilters((fs) => fs.map((x) => (x.column === f.column ? { ...x, value: v } : x)))}
                    placeholder="Seleccionar…"
                  />
                ) : null}

                {f.kind === "range" ? (
                  <RangeFilter
                    label=""
                    min={f.value?.min ?? 0}
                    max={f.value?.max ?? 0}
                    value={{ from: f.value?.from ?? null, to: f.value?.to ?? null }}
                    onChange={(v) =>
                      setDynamicFilters((fs) => fs.map((x) => (x.column === f.column ? { ...x, value: { ...x.value, ...v } } : x)))
                    }
                  />
                ) : null}

                {f.kind === "boolean" ? (
                  <div className="mt-1 flex gap-2">
                    {[
                      { label: "Cualquiera", val: null as null | boolean },
                      { label: "True", val: true },
                      { label: "False", val: false },
                    ].map((o) => (
                      <button
                        key={o.label}
                        type="button"
                        onClick={() => setDynamicFilters((fs) => fs.map((x) => (x.column === f.column ? { ...x, value: o.val } : x)))}
                        className={classNames(
                          "rounded-xl border px-3 py-2 text-sm shadow-sm",
                          f.value === o.val ? "border-slate-900 bg-slate-900 text-white" : "border-slate-200 bg-white text-slate-700 hover:border-slate-300"
                        )}
                      >
                        {o.label}
                      </button>
                    ))}
                  </div>
                ) : null}

                {f.kind === "text" ? (
                  <input
                    value={String(f.value ?? "")}
                    onChange={(e) =>
                      setDynamicFilters((fs) =>
                        fs.map((x) => (x.column === f.column ? { ...x, value: e.target.value } : x))
                      )
                    }
                    placeholder="Buscar… (contiene)"
                    className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-slate-300 focus:outline-none"
                  />
                ) : null}
              </div>
            ))}

            {hasData ? (
              <button
                type="button"
                className="ml-auto rounded-xl bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-800"
                onClick={() => {
  setSelectedColumns([]);
  setDynamicFilters([]);
  setFixedFilters({
    Buyer: [],
    Supplier: [],
    Country: [],
    Year: [],
    Sector: [],
    Turnover: { from: null, to: null },
  });
}}
              >
                Limpiar filtros
              </button>
            ) : null}
          </div>
        </div>
      </div>

      {/* Main */}
      <div className="mx-auto w-full max-w-none px-4 py-6 sm:px-6 lg:px-10 2xl:px-14">
        {!hasData ? (
          <div className="flex min-h-[60vh] items-center justify-center">
            <EmptyState text="Cargá tus Excel/CSV para generar el dashboard automáticamente (todo se procesa localmente)." />
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-6">
            {/* Reconciliación (compacto) */}
            <div className="rounded-2xl border border-slate-200 bg-white shadow-sm">
              <div className="flex flex-col gap-3 border-b border-slate-100 px-5 py-4 md:flex-row md:items-center md:justify-between">
                <div className="flex flex-wrap items-center gap-2">
                  <div className="text-sm font-semibold text-slate-800">Unificación de Suppliers</div>
                  <Pill>Pendientes: {reconcileCounts.pendingGroups} grupos</Pill>
                  <Pill>Pendientes fuzzy: {reconcileCounts.pendingFuzzy}</Pill>
                  {reconcileCounts.pendingGroups === 0 && reconcileCounts.pendingFuzzy === 0 ? <Pill>✅ Todo conciliado</Pill> : null}
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  <label className="flex items-center gap-2 text-xs text-slate-600">
                    <input
                      type="checkbox"
                      checked={supplierUnifyEnabled}
                      onChange={(e) => setSupplierUnifyEnabled(e.target.checked)}
                      className="h-4 w-4 rounded border-slate-300"
                    />
                    Unificación en dashboard
                  </label>

                  <button
                    type="button"
                    className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 hover:border-slate-300"
                    onClick={recalcAll}
                  >
                    Recalcular
                  </button>

                  <button
                    type="button"
                    className="rounded-xl bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-800"
                    onClick={applyAllSuggestedMerges}
                    title="Aplica grupos exactos + fuzzy pendientes"
                  >
                    Aplicar sugeridos
                  </button>

                  {(reconcileCounts.pendingGroups > 0 || reconcileCounts.pendingFuzzy > 0 || showResolvedReconciliation) ? (
                    <button
                      type="button"
                      className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 hover:border-slate-300"
                      onClick={() => setReconcileOpen((s) => !s)}
                    >
                      {reconcileOpen ? "Minimizar" : "Revisar"}
                    </button>
                  ) : null}
                </div>
              </div>

              {reconcileOpen ? (
                <div className="p-5">
                  {!mapping.Supplier ? (
                    <div className="text-sm text-slate-600">No se encontró la cabecera “Supplier”.</div>
                  ) : (
                    <>
                      <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                        <div className="text-xs text-slate-600">Revisá y autorizá merges exactos y sugerencias fuzzy.</div>
                        <div className="flex flex-wrap items-center gap-3">
                          <label className="flex items-center gap-2 text-xs text-slate-600">
                            <input
                              type="checkbox"
                              checked={showResolvedReconciliation}
                              onChange={(e) => setShowResolvedReconciliation(e.target.checked)}
                              className="h-4 w-4 rounded border-slate-300"
                            />
                            Mostrar resueltos
                          </label>
                          <input
                            value={supplierGroupSearch}
                            onChange={(e) => setSupplierGroupSearch(e.target.value)}
                            placeholder="Buscar supplier / variante…"
                            className="w-full md:w-80 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm"
                          />
                        </div>
                      </div>

                      <div className="mb-3 flex flex-wrap items-center gap-2 rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600">
                        <Pill>Pendientes: {reconcileCounts.pendingGroups} grupos</Pill>
                        <Pill>Resueltos: {reconcileCounts.resolvedGroups} grupos</Pill>
                        <Pill>Pendientes fuzzy: {reconcileCounts.pendingFuzzy}</Pill>
                        <Pill>Resueltos fuzzy: {reconcileCounts.resolvedFuzzy}</Pill>
                      </div>

                      {/* Grupos exactos */}
                      <div className="space-y-3">
                        {supplierGroupsFiltered.length === 0 ? (
                          <div className="text-sm text-slate-600">No hay grupos que coincidan con la búsqueda.</div>
                        ) : (
                          supplierGroupsFiltered.slice(0, 8).map((g) => {
                            const total = g.variants.reduce((s, v) => s + v.count, 0);
                            return (
                              <div key={g.key} className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                                <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                                  <div>
                                    <div className="text-sm font-semibold text-slate-900">
                                      {g.chosenDisplay}
                                      <span className="ml-2 text-xs font-medium text-slate-500">({total} registros)</span>
                                    </div>
                                    <div className="mt-1 flex flex-wrap gap-2">
                                      {g.variants.slice(0, 6).map((v) => (
                                        <Pill key={v.name}>
                                          {v.name} · {v.count}
                                        </Pill>
                                      ))}
                                      {g.variants.length > 6 ? <Pill>+{g.variants.length - 6} más</Pill> : null}
                                    </div>
                                  </div>

                                  <div className="flex flex-col gap-2 md:items-end">
                                    <label className="flex items-center gap-2 text-sm text-slate-700">
                                      <input
                                        type="checkbox"
                                        checked={g.merge}
                                        onChange={(e) => {
                                          const merge = e.target.checked;
                                          const updated = supplierGroups.map((x) => (x.key === g.key ? { ...x, merge } : x));
                                          setSupplierGroups(updated);
                                          setSupplierGroupOverrides((prev) => ({
                                            ...prev,
                                            [g.key]: { merge, chosenDisplay: (prev[g.key]?.chosenDisplay ?? g.chosenDisplay) },
                                          }));
                                          showToast("Actualizado en tiempo real ✅");
                                        }}
                                        className="h-4 w-4 rounded border-slate-300"
                                      />
                                      Unificar este grupo
                                    </label>

                                    <div className="flex items-center gap-2">
                                      <span className="text-xs text-slate-600">Nombre:</span>
                                      <select
                                        value={g.chosenDisplay}
                                        onChange={(e) => {
                                          const chosenDisplay = e.target.value;
                                          const updated = supplierGroups.map((x) => (x.key === g.key ? { ...x, chosenDisplay } : x));
                                          setSupplierGroups(updated);
                                          setSupplierGroupOverrides((prev) => ({
                                            ...prev,
                                            [g.key]: { merge: (prev[g.key]?.merge ?? g.merge), chosenDisplay },
                                          }));
                                          showToast("Actualizado en tiempo real ✅");
                                        }}
                                        className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm"
                                      >
                                        {g.variants.map((v) => (
                                          <option key={v.name} value={v.name}>
                                            {v.name}
                                          </option>
                                        ))}
                                      </select>
                                    </div>

                                    <button
                                      type="button"
                                      className="rounded-xl bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-800"
                                      onClick={() => applyAliasesNow()}
                                    >
                                      Guardar (en vivo)
                                    </button>
                                  </div>
                                </div>
                              </div>
                            );
                          })
                        )}

                        {supplierGroupsFiltered.length > 8 ? (
                          <div className="text-xs text-slate-500">Mostrando 8 grupos (buscador reduce lista).</div>
                        ) : null}
                      </div>

                      {/* Fuzzy */}
                      <div className="mt-6">
                        <div className="mb-2 text-sm font-semibold text-slate-900">Sugerencias fuzzy</div>
                        <div className="text-xs text-slate-600 mb-3">Pares similares fuera de grupos exactos.</div>

                        {(showResolvedReconciliation ? fuzzySuggestions : fuzzySuggestions.filter((s) => s.approved === null)).length === 0 ? (
                          <div className="text-sm text-slate-600">No hay sugerencias fuzzy pendientes.</div>
                        ) : (
                          <div className="space-y-2">
                            {(showResolvedReconciliation ? fuzzySuggestions : fuzzySuggestions.filter((s) => s.approved === null))
                              .slice(0, 8)
                              .map((s, idx) => (
                                <div key={`${s.a}-${s.b}-${idx}`} className="rounded-xl border border-slate-200 bg-white p-4">
                                  <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                                    <div>
                                      <div className="text-sm font-semibold text-slate-900">
                                        {s.a} <span className="text-slate-400">≈</span> {s.b}
                                      </div>
                                      <div className="mt-1 text-xs text-slate-600">
                                        Score: <span className="font-semibold">{(s.score * 100).toFixed(1)}%</span>
                                      </div>
                                    </div>

                                    <div className="flex flex-wrap items-center gap-2">
                                      <span className="text-xs text-slate-600">Nombre:</span>
                                      <select
                                        value={s.chosenDisplay}
                                        onChange={(e) => {
                                          const chosenDisplay = e.target.value;
                                          setFuzzySuggestions((prev) =>
                                            prev.map((x) =>
                                              x.a === s.a && x.b === s.b && x.score === s.score ? { ...x, chosenDisplay } : x
                                            )
                                          );
                                        }}
                                        className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm"
                                      >
                                        {[s.a, s.b].map((v) => (
                                          <option key={v} value={v}>
                                            {v}
                                          </option>
                                        ))}
                                      </select>

                                      <button
                                        type="button"
                                        className={classNames(
                                          "rounded-xl px-3 py-2 text-xs font-semibold",
                                          s.approved === true
                                            ? "bg-slate-900 text-white"
                                            : "border border-slate-200 bg-white text-slate-800 hover:border-slate-300"
                                        )}
                                        onClick={() => approveFuzzy(s, true)}
                                      >
                                        Unificar
                                      </button>

                                      <button
                                        type="button"
                                        className={classNames(
                                          "rounded-xl px-3 py-2 text-xs font-semibold",
                                          s.approved === false
                                            ? "bg-slate-200 text-slate-900"
                                            : "border border-slate-200 bg-white text-slate-800 hover:border-slate-300"
                                        )}
                                        onClick={() => approveFuzzy(s, false)}
                                      >
                                        Separar
                                      </button>
                                    </div>
                                  </div>
                                </div>
                              ))}
                          </div>
                        )}

                        {Object.keys(fuzzyApprovedAliases).length > 0 ? (
                          <div className="mt-3 text-xs text-slate-500">
                            Aprobaciones fuzzy activas: {Object.keys(fuzzyApprovedAliases).length} alias(es).
                          </div>
                        ) : null}
                      </div>
                    </>
                  )}
                </div>
              ) : null}
            </div>

            {/* KPIs */}
            <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-3 inline-flex h-9 w-9 items-center justify-center rounded-xl text-sm" style={{ background: BRAND.neon, color: BRAND.navy }}>▦</div>
              <div className="text-xs font-medium text-slate-500">Registros</div>
              <div className="mt-1 text-2xl font-semibold text-slate-900">{formatCompact(kpis.total)}</div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-3 inline-flex h-9 w-9 items-center justify-center rounded-xl text-sm" style={{ background: BRAND.neon, color: BRAND.navy }}>◎</div>
              <div className="text-xs font-medium text-slate-500">Suppliers únicos (unificados)</div>
              <div className="mt-1 text-2xl font-semibold text-slate-900">
                {kpis.uniqueSuppliers === null ? "—" : formatCompact(kpis.uniqueSuppliers)}
              </div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-3 inline-flex h-9 w-9 items-center justify-center rounded-xl text-sm" style={{ background: BRAND.neon, color: BRAND.navy }}>⌁</div>
              <div className="text-xs font-medium text-slate-500">Countries únicos</div>
              <div className="mt-1 text-2xl font-semibold text-slate-900">
                {kpis.uniqueCountries === null ? "—" : formatCompact(kpis.uniqueCountries)}
              </div>
            </div>
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-3 inline-flex h-9 w-9 items-center justify-center rounded-xl text-sm" style={{ background: BRAND.neon, color: BRAND.navy }}>€</div>
              <div className="text-xs font-medium text-slate-500">Turnover (sum)</div>
              <div className="mt-1 text-2xl font-semibold text-slate-900">
                {kpis.turnoverSum === null ? "—" : formatCompact(kpis.turnoverSum)}
              </div>
              <div className="mt-1 text-xs text-slate-500">
                Avg: {kpis.turnoverAvg === null ? "—" : formatCompact(kpis.turnoverAvg)}
              </div>
            </div>
            </div>

            {/* Charts */}
            <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
              <Card title="Top Countries (count)" right={<span>{mapping.Country ? "por Country" : "sin Country"}</span>}>
                {mapping.Country ? (
                  <div className="h-80">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={kpis.byCountry}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="name" tick={{ fontSize: 12 }} interval={0} angle={-20} textAnchor="end" height={70} />
                        <YAxis />
                        <Tooltip />
                        <Bar dataKey="value" fill={BRAND.neon} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="text-sm text-slate-600">No se encontró la cabecera “Country”.</div>
                )}
              </Card>

              <Card
                title={mapping.Year ? "Evolución por Year" : "Evolución (sin Year)"}
                right={mapping.Year && mapping.Turnover ? <span>líneas: count y turnover</span> : mapping.Year ? <span>línea: count</span> : null}
              >
                {mapping.Year ? (
                  <div className="h-80">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={kpis.byYear}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="year" />
                        <YAxis />
                        <Tooltip />
                        <Line type="monotone" dataKey="value" stroke={BRAND.skyBlue} strokeWidth={2} dot={false} />
                        {mapping.Turnover ? (
                          <Line type="monotone" dataKey="turnover" stroke={BRAND.electricBlue} strokeWidth={2} dot={false} />
                        ) : null}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="text-sm text-slate-600">No se encontró la cabecera “Year”.</div>
                )}
              </Card>
            </div>

            {/* Preview */}
            <Card
              title="Vista previa"
              right={
                <span>
                  {headers.length} columnas · {filteredRows.length} filas · pág {previewPage}/{previewTotalPages}
                </span>
              }
            >
              {previewHeaders.length === 0 ? (
                <div className="rounded-xl border border-dashed border-slate-300 bg-white p-6 text-sm text-slate-600">
                  Elegí al menos una columna en “Columnas (y filtros)” para ver la vista previa.
                </div>
              ) : (
                <div className="overflow-auto min-h-[360px] rounded-xl border border-slate-200">
                  <table className="min-w-full text-left text-sm">
                    <thead className="bg-slate-50 text-xs text-slate-600">
                      <tr>
                        {previewHeaders.map((h) => (
                          <th key={h} className="whitespace-nowrap px-3 py-2 font-semibold">
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows
                        .filter((r) => !previewHeaders.every((h) => normStr(r?.[h]) === ""))
                        .map((r, idx) => (
                          <tr key={String(r?.["Record ID"] ?? idx)} className="border-b border-slate-100 last:border-0">
                            {previewHeaders.map((h) => {
                              const v = r?.[h];
                              const s = v === null || v === undefined ? "" : String(v);
                              return (
                                <td key={h} className="px-3 py-2 text-xs text-slate-700">
                                  <span className="block truncate">{s}</span>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              )}

              <div className="mt-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div className="text-xs text-slate-600">
                  Mostrando <span className="font-semibold">{filteredRows.length ? previewFrom + 1 : 0}</span>–{" "}
                  <span className="font-semibold">{previewTo}</span> de{" "}
                  <span className="font-semibold">{filteredRows.length}</span>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <select
                    value={previewPageSize}
                    onChange={(e) => {
                      setPreviewPageSize(Number(e.target.value));
                      setPreviewPage(1);
                    }}
                    className="rounded-xl border border-slate-200 bg-white px-2 py-2 text-xs shadow-sm"
                    title="Filas por página"
                  >
                    {[10, 20, 30, 50, 100].map((n) => (
                      <option key={n} value={n}>
                        {n} / pág
                      </option>
                    ))}
                  </select>

                  <button
                    type="button"
                    className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 disabled:opacity-40"
                    disabled={previewPage <= 1}
                    onClick={() => setPreviewPage(1)}
                  >
                    « Primera
                  </button>

                  <button
                    type="button"
                    className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 disabled:opacity-40"
                    disabled={previewPage <= 1}
                    onClick={() => setPreviewPage((p) => Math.max(1, p - 1))}
                  >
                    ‹ Anterior
                  </button>

                  <div className="flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs shadow-sm">
                    <span className="text-slate-600">Pág</span>
                    <input
                      type="number"
                      min={1}
                      max={previewTotalPages}
                      value={previewPage}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        if (!Number.isFinite(v)) return;
                        setPreviewPage(Math.min(Math.max(1, v), previewTotalPages));
                      }}
                      className="w-16 rounded-lg border border-slate-200 bg-white px-2 py-1 text-xs focus:outline-none"
                    />
                    <span className="text-slate-600">de</span>
                    <span className="font-semibold text-slate-800">{previewTotalPages}</span>
                  </div>

                  <button
                    type="button"
                    className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 disabled:opacity-40"
                    disabled={previewPage >= previewTotalPages}
                    onClick={() => setPreviewPage((p) => Math.min(previewTotalPages, p + 1))}
                  >
                    Siguiente › 
                  </button>

                  <button
                    type="button"
                    className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-800 disabled:opacity-40"
                    disabled={previewPage >= previewTotalPages}
                    onClick={() => setPreviewPage(previewTotalPages)}
                  >
                    Última »
                  </button>
                </div>
              </div>

              <div className="mt-3 flex flex-wrap gap-2">
                {Object.entries(mapping)
                  .filter(([, col]) => col)
                  .map(([k, col]) => (
                    <Pill key={k}>
                      {k}: {col}
                    </Pill>
                  ))}
                <Pill>Supplier unificación: {supplierUnifyEnabled ? "ON" : "OFF"}</Pill>
              </div>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}