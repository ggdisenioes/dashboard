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

type FilterKind = "multi" | "range" | "boolean";
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
  return String(h ?? "").trim();
}
function safeStr(v: any) {
  return String(v ?? "").trim();
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
]);

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
}: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (v: string[]) => void;
  placeholder?: string;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

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
          <div className="max-h-64 overflow-auto p-2">
            {options.length === 0 ? (
              <div className="p-3 text-sm text-slate-500">Sin opciones</div>
            ) : (
              options.map((o) => (
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
  const [addOpen, setAddOpen] = useState(false);
  const addRef = useRef<HTMLDivElement | null>(null);

  // Unificación supplier
  const [supplierGroups, setSupplierGroups] = useState<SupplierGroup[]>([]);
  const [supplierAliasMap, setSupplierAliasMap] = useState<Record<string, string>>({});
  const [supplierUnifyEnabled, setSupplierUnifyEnabled] = useState(true);

  // UPGRADE 1: search
  const [supplierGroupSearch, setSupplierGroupSearch] = useState("");

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

  function showToast(msg: string) {
    setToast(msg);
    if (toastTimerRef.current) window.clearTimeout(toastTimerRef.current);
    toastTimerRef.current = window.setTimeout(() => setToast(null), 1800);
  }

  // Keep latest values to avoid stale closures in click handlers
  const supplierGroupsRef = useRef<SupplierGroup[]>([]);
  const fuzzyApprovedAliasesRef = useRef<Record<string, string>>({});

  // IMPORTANT: keep refs in sync during render so buttons work immediately after loading data
  supplierGroupsRef.current = supplierGroups;
  fuzzyApprovedAliasesRef.current = fuzzyApprovedAliases;

  // Auto-apply supplier alias mapping whenever the user changes merge toggles or chosen display
  useEffect(() => {
    setSupplierAliasMap(buildSupplierAliasMap(supplierGroups, fuzzyApprovedAliases));
  }, [supplierGroups, fuzzyApprovedAliases]);

  useEffect(() => {
    const onDoc = (e: MouseEvent) => {
      if (!addRef.current) return;
      if (!addRef.current.contains(e.target as Node)) setAddOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  // --- IndexedDB: Save/reset helpers ---
  const saveTimerRef = useRef<number | null>(null);

  async function clearSavedSession() {
    await idbDel(IDB_KEY);
    setHasSavedSession(false);

    // Reset app state
    setRows([]);
    setHeaders([]);
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

  function applyOverridesToGroups(groups: SupplierGroup[]) {
    return groups.map((g) => {
      const ov = supplierGroupOverrides[g.key];
      if (!ov) return g;
      return { ...g, merge: ov.merge, chosenDisplay: ov.chosenDisplay };
    });
  }

  function recomputeSupplierStuff(allRows: Row[], supplierCol: string | null) {
    const rawGroups = buildSupplierGroups(allRows, supplierCol);
    const groups = applyOverridesToGroups(rawGroups);
    setSupplierGroups(groups);

    const fuzzy = buildFuzzySuggestions(allRows, supplierCol, 40);
    setFuzzySuggestions(fuzzy);

    // Alias final (grupos + fuzzy aprobados)
    setSupplierAliasMap(buildSupplierAliasMap(groups, fuzzyApprovedAliasesRef.current));
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
      const json: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null });

      const hs = Object.keys(json[0] ?? {}).map(normalizeHeader);
      if (incomingHeaders.length === 0) incomingHeaders = hs;

      const normalizedRows: Row[] = json.map((r) => {
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

    const detected: FieldMapping = {
      Buyer: mapping.Buyer ?? pickColumn(headerUnion, ["Buyer", "buyer", "Brand", "brand", "Customer", "customer"]),
      Supplier: mapping.Supplier ?? pickColumn(headerUnion, ["Supplier", "supplier", "Vendor", "vendor", "Factory", "factory", "name"]),
      Country: mapping.Country ?? pickColumn(headerUnion, ["Country", "country", "CountryName", "country_name", "country name"]),
      Year: mapping.Year ?? pickColumn(headerUnion, ["Year", "year", "Año", "anio", "FY", "fiscal_year"]),
      Turnover: mapping.Turnover ?? pickColumn(headerUnion, ["Turnover", "turnover", "Revenue", "revenue", "Sales", "sales", "Amount", "amount"]),
      Sector: mapping.Sector ?? pickColumn(headerUnion, ["Sector", "sector", "Industry", "industry"]),
    };

    setRows(merged);
    setHeaders(headerUnion);
    setColTypes(types);
    setMapping(detected);
    setFilesLoaded((n) => n + fileList.length);

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

    if (detected.Buyer) {
      const vals = uniq(merged.map((r) => safeStr(r[detected.Buyer!])).filter((v) => v !== ""));
      if (vals.length === 1) setFixedFilters((s) => ({ ...s, Buyer: vals }));
    }
  }

  // Opciones fijas
  const fixedOptions = useMemo(() => {
    const opts: Record<string, string[]> = {};
    for (const key of ["Buyer", "Country", "Year", "Sector"] as const) {
      const col = mapping[key];
      if (!col) {
        opts[key] = [];
        continue;
      }
      opts[key] = uniq(rows.map((r) => safeStr(r[col])).filter(Boolean)).sort((a, b) => a.localeCompare(b));
    }

    if (mapping.Supplier) {
      const col = mapping.Supplier;
      opts.Supplier = uniq(rows.map((r) => supplierDisplay(safeStr(r[col]))).filter(Boolean)).sort((a, b) => a.localeCompare(b));
    } else {
      opts.Supplier = [];
    }
    return opts;
  }, [rows, mapping, supplierAliasMap, supplierUnifyEnabled]);

  const turnoverRange = useMemo(() => {
    const col = mapping.Turnover;
    if (!col) return { min: 0, max: 0, ok: false };
    const nums = rows.map((r) => toNumber(r[col])).filter((n): n is number => n !== null && isFinite(n));
    if (nums.length === 0) return { min: 0, max: 0, ok: false };
    return { min: Math.min(...nums), max: Math.max(...nums), ok: true };
  }, [rows, mapping]);

  // Add-filter candidates
  const addCandidates = useMemo(() => {
    const fixedCols = new Set(Object.values(mapping).filter(Boolean) as string[]);
    const addedCols = new Set(dynamicFilters.map((f) => f.column));
    return headers
      .filter((h) => !fixedCols.has(h))
      .filter((h) => !addedCols.has(h))
      .filter((h) => !["id", "os_id", "lat", "lng", "address"].includes(h.toLowerCase()))
      .sort((a, b) => a.localeCompare(b));
  }, [headers, mapping, dynamicFilters]);

  function createDynamicFilter(column: string) {
    const t = colTypes[column] ?? "unknown";
    let kind: FilterKind = "multi";
    let value: any = [];
    if (t === "number") {
      const nums = rows.map((r) => toNumber(r[column])).filter((n): n is number => n !== null);
      const min = nums.length ? Math.min(...nums) : 0;
      const max = nums.length ? Math.max(...nums) : 0;
      kind = "range";
      value = { from: null, to: null, min, max };
    } else if (t === "boolean") {
      kind = "boolean";
      value = null;
    } else {
      kind = "multi";
      value = [];
    }
    setDynamicFilters((fs) => [...fs, { id: column, column, kind, value }]);
    setAddOpen(false);
  }
  function removeDynamicFilter(col: string) {
    setDynamicFilters((fs) => fs.filter((f) => f.column !== col));
  }

  // Dynamic options
  const dynamicOptions = useMemo(() => {
    const out: Record<string, string[]> = {};
    for (const f of dynamicFilters) {
      if (f.kind !== "multi") continue;
      out[f.column] = uniq(rows.map((r) => safeStr(r[f.column])).filter(Boolean))
        .sort((a, b) => a.localeCompare(b))
        .slice(0, 5000);
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
      const setSel = new Set(selected.map((x) => x.toLowerCase()));
      out = out.filter((r) => setSel.has(safeStr(r[col]).toLowerCase()));
    });

    // Supplier (unificado)
    if (mapping.Supplier && fixedFilters.Supplier.length > 0) {
      const col = mapping.Supplier;
      const setSel = new Set(fixedFilters.Supplier.map((x) => x.toLowerCase()));
      out = out.filter((r) => setSel.has(supplierDisplay(safeStr(r[col])).toLowerCase()));
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
        const setSel = new Set(selected.map((x) => x.toLowerCase()));
        out = out.filter((r) => setSel.has(safeStr(r[col]).toLowerCase()));
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
    }

    return out;
  }, [rows, mapping, fixedFilters, dynamicFilters, turnoverRange, supplierAliasMap, supplierUnifyEnabled]);

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

  // UPGRADE 2: apply all suggested merges
  function applyAllSuggestedMerges() {
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

    setSupplierUnifyEnabled(true);
    setFixedFilters((s) => ({ ...s, Supplier: [] }));
    showToast("Merges sugeridos aplicados ✅");
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
          // recompute will apply overrides and build alias map
          // defer to next tick to ensure overrides are in state
          setTimeout(() => {
            recomputeSupplierStuff(snap.rows, supplierCol);
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
    const groups = buildSupplierGroups(rows, supplierCol);
    const fuzzy = buildFuzzySuggestions(rows, supplierCol, 40);
    setSupplierGroups(groups);
    setFuzzySuggestions(fuzzy);
    setSupplierUnifyEnabled(true);
    setFixedFilters((s) => ({ ...s, Supplier: [] }));
    showToast("Recalculado ✅");
  }

  const supplierGroupsFiltered = useMemo(() => {
    const q = supplierGroupSearch.trim().toLowerCase();
    if (!q) return supplierGroups;
    return supplierGroups.filter((g) => {
      if (g.chosenDisplay.toLowerCase().includes(q)) return true;
      if (g.suggestedDisplay.toLowerCase().includes(q)) return true;
      return g.variants.some((v) => v.name.toLowerCase().includes(q));
    });
  }, [supplierGroups, supplierGroupSearch]);

  return (
    <div className="min-h-screen bg-slate-50">
      {toast ? (
        <div className="pointer-events-none fixed left-1/2 top-4 z-[999] -translate-x-1/2 rounded-xl bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-lg">
          {toast}
        </div>
      ) : null}
      {/* Top bar */}
      <div className="border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto w-full max-w-none px-4 py-4 sm:px-6 lg:px-10 2xl:px-14">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="text-lg font-semibold text-slate-900">Dashboard de Suppliers</div>
              <div className="text-xs text-slate-600">
                Cargá uno o varios CSV/XLSX. Supplier se reconcilia (variantes + fuzzy). Exporta resultados filtrados.
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <label className="inline-flex cursor-pointer items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm hover:border-slate-300">
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
                <span className="text-slate-700">📁 Cargar archivo(s)</span>
              </label>

              {hasSavedSession ? (
                <button
                  type="button"
                  className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm hover:border-slate-300"
                  onClick={clearSavedSession}
                  title="Borra la sesión guardada (datos + decisiones)"
                >
                  🧹 Borrar sesión
                </button>
              ) : null}

              <button
                type="button"
                className="rounded-xl bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-800 disabled:opacity-50"
                disabled={!hasData || filteredRows.length === 0}
                onClick={exportFilteredToExcel}
                title={!hasData ? "Cargá archivos" : filteredRows.length === 0 ? "No hay filas para exportar" : "Exportar filtrado"}
              >
                ⬇ Export Excel (filtrado)
              </button>

              {hasData ? (
                <>
                  <Pill>{formatCompact(filteredRows.length)} / {formatCompact(rows.length)} filas</Pill>
                  <Pill>{filesLoaded} archivo(s)</Pill>
                  {hasSavedSession ? <Pill>Sesión guardada</Pill> : null}
                </>
              ) : (
                <Pill>Sin archivo</Pill>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Filters (sticky) */}
      <div className="sticky top-0 z-40 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto w-full max-w-none px-4 py-4 sm:px-6 lg:px-10 2xl:px-14">
          <div className="grid grid-cols-1 gap-3 md:grid-cols-6">
            <MultiSelect
              label="Buyer"
              options={fixedOptions.Buyer ?? []}
              selected={fixedFilters.Buyer}
              onChange={(v) => setFixedFilters((s) => ({ ...s, Buyer: v }))}
              placeholder={mapping.Buyer ? "Seleccionar Buyer…" : "Columna no encontrada"}
            />
            <MultiSelect
              label="Supplier"
              options={fixedOptions.Supplier ?? []}
              selected={fixedFilters.Supplier}
              onChange={(v) => setFixedFilters((s) => ({ ...s, Supplier: v }))}
              placeholder={mapping.Supplier ? "Seleccionar Supplier…" : "Columna no encontrada"}
            />
            <MultiSelect
              label="Country"
              options={fixedOptions.Country ?? []}
              selected={fixedFilters.Country}
              onChange={(v) => setFixedFilters((s) => ({ ...s, Country: v }))}
              placeholder={mapping.Country ? "Seleccionar Country…" : "Columna no encontrada"}
            />
            <MultiSelect
              label="Year"
              options={fixedOptions.Year ?? []}
              selected={fixedFilters.Year}
              onChange={(v) => setFixedFilters((s) => ({ ...s, Year: v }))}
              placeholder={mapping.Year ? "Seleccionar Year…" : "Columna no encontrada"}
            />
            <MultiSelect
              label="Sector"
              options={fixedOptions.Sector ?? []}
              selected={fixedFilters.Sector}
              onChange={(v) => setFixedFilters((s) => ({ ...s, Sector: v }))}
              placeholder={mapping.Sector ? "Seleccionar Sector…" : "Columna no encontrada"}
            />
            <div>
              {mapping.Turnover && turnoverRange.ok ? (
                <RangeFilter
                  label="Turnover"
                  min={turnoverRange.min}
                  max={turnoverRange.max}
                  value={fixedFilters.Turnover}
                  onChange={(v) => setFixedFilters((s) => ({ ...s, Turnover: v }))}
                />
              ) : (
                <div>
                  <div className="mb-1 text-xs font-medium text-slate-700">Turnover</div>
                  <div className="rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm text-slate-400 shadow-sm">
                    Columna no encontrada / sin valores numéricos
                  </div>
                </div>
              )}
            </div>
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
              </div>
            ))}

            <div className="relative" ref={addRef}>
              <button
                type="button"
                className="inline-flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm hover:border-slate-300"
                onClick={() => setAddOpen((s) => !s)}
                disabled={!hasData}
              >
                ➕ Add filter <span className="text-slate-400">▾</span>
              </button>

              {addOpen ? (
                <div className="absolute z-50 mt-2 w-72 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
                  <div className="max-h-72 overflow-auto p-2">
                    {addCandidates.length === 0 ? (
                      <div className="p-3 text-sm text-slate-500">No hay más columnas relevantes para filtrar.</div>
                    ) : (
                      addCandidates.map((c) => (
                        <button
                          key={c}
                          type="button"
                          className="flex w-full items-center justify-between rounded-lg px-3 py-2 text-left text-sm text-slate-800 hover:bg-slate-50"
                          onClick={() => createDynamicFilter(c)}
                        >
                          <span className="truncate">{c}</span>
                          <span className="text-xs text-slate-400">{colTypes[c] ?? "unknown"}</span>
                        </button>
                      ))
                    )}
                  </div>
                </div>
              ) : null}
            </div>

            {hasData ? (
              <button
                type="button"
                className="ml-auto rounded-xl bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-800"
                onClick={() => {
                  setFixedFilters({
                    Buyer: [],
                    Supplier: [],
                    Country: [],
                    Year: [],
                    Sector: [],
                    Turnover: { from: null, to: null },
                  });
                  setDynamicFilters([]);
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
            {/* Reconciliación */}
            <Card
              title="Unificación de Suppliers (reconciliación)"
              right={
                <div className="flex flex-wrap items-center gap-3">
                  <label className="flex items-center gap-2 text-xs text-slate-600">
                    <input
                      type="checkbox"
                      checked={supplierUnifyEnabled}
                      onChange={(e) => setSupplierUnifyEnabled(e.target.checked)}
                      className="h-4 w-4 rounded border-slate-300"
                    />
                    Activar unificación en dashboard
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
                    title="Marca todos los grupos como merge=true y aplica el nombre sugerido"
                  >
                    Apply all suggested merges
                  </button>
                </div>
              }
            >
              {!mapping.Supplier ? (
                <div className="text-sm text-slate-600">No se encontró la cabecera “Supplier”.</div>
              ) : (
                <>
                  {/* UPGRADE 1: search */}
                  <div className="mb-3 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
                    <div className="text-xs text-slate-600">
                      Revisá y autorizá merges exactos (canonicalKey) y sugerencias fuzzy.
                    </div>
                    <input
                      value={supplierGroupSearch}
                      onChange={(e) => setSupplierGroupSearch(e.target.value)}
                      placeholder="Buscar supplier / variante…"
                      className="w-full md:w-80 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm"
                    />
                  </div>

                  {/* Grupos exactos */}
                  <div className="space-y-3">
                    {supplierGroupsFiltered.length === 0 ? (
                      <div className="text-sm text-slate-600">No hay grupos que coincidan con la búsqueda.</div>
                    ) : (
                      supplierGroupsFiltered.slice(0, 12).map((g) => {
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
                                  {g.variants.slice(0, 8).map((v) => (
                                    <Pill key={v.name}>
                                      {v.name} · {v.count}
                                    </Pill>
                                  ))}
                                  {g.variants.length > 8 ? <Pill>+{g.variants.length - 8} más</Pill> : null}
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
                                  <span className="text-xs text-slate-600">Nombre unificado:</span>
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
                                  Guardar (ya es en vivo)
                                </button>
                              </div>
                            </div>
                          </div>
                        );
                      })
                    )}

                    {supplierGroupsFiltered.length > 12 ? (
                      <div className="text-xs text-slate-500">Mostrando 12 grupos (búsqueda reduce lista).</div>
                    ) : null}
                  </div>

                  {/* UPGRADE 3: fuzzy suggestions */}
                  <div className="mt-6">
                    <div className="mb-2 text-sm font-semibold text-slate-900">Sugerencias fuzzy (posibles duplicados)</div>
                    <div className="text-xs text-slate-600 mb-3">
                      Estos pares no cayeron en el mismo grupo exacto, pero son muy similares. Aprobá “Unificar” o “Separar”.
                    </div>

                    {fuzzySuggestions.length === 0 ? (
                      <div className="text-sm text-slate-600">No detecté duplicados probables con el umbral actual.</div>
                    ) : (
                      <div className="space-y-2">
                        {fuzzySuggestions.map((s, idx) => (
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
                                <span className="text-xs text-slate-600">Nombre unificado:</span>
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
                                    s.approved === true ? "bg-slate-900 text-white" : "border border-slate-200 bg-white text-slate-800 hover:border-slate-300"
                                  )}
                                  onClick={() => approveFuzzy(s, true)}
                                >
                                  Unificar
                                </button>

                                <button
                                  type="button"
                                  className={classNames(
                                    "rounded-xl px-3 py-2 text-xs font-semibold",
                                    s.approved === false ? "bg-slate-200 text-slate-900" : "border border-slate-200 bg-white text-slate-800 hover:border-slate-300"
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
            </Card>

            {/* KPIs */}
            <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
              <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="text-xs font-medium text-slate-500">Registros</div>
                <div className="mt-1 text-2xl font-semibold text-slate-900">{formatCompact(kpis.total)}</div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="text-xs font-medium text-slate-500">Suppliers únicos (unificados)</div>
                <div className="mt-1 text-2xl font-semibold text-slate-900">
                  {kpis.uniqueSuppliers === null ? "—" : formatCompact(kpis.uniqueSuppliers)}
                </div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <div className="text-xs font-medium text-slate-500">Countries únicos</div>
                <div className="mt-1 text-2xl font-semibold text-slate-900">
                  {kpis.uniqueCountries === null ? "—" : formatCompact(kpis.uniqueCountries)}
                </div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
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
                        <Bar dataKey="value" />
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
                        <Line type="monotone" dataKey="value" />
                        {mapping.Turnover ? <Line type="monotone" dataKey="turnover" /> : null}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="text-sm text-slate-600">No se encontró la cabecera “Year”.</div>
                )}
              </Card>
            </div>

            {/* Preview */}
            <Card title="Vista previa (primeras 30 filas)" right={<span>{headers.length} columnas</span>}>
              <div className="overflow-auto rounded-xl border border-slate-200">
                <table className="min-w-full text-left text-sm">
                  <thead className="bg-slate-50 text-xs text-slate-600">
                    <tr>
                      {headers.slice(0, 10).map((h) => (
                        <th key={h} className="whitespace-nowrap px-3 py-2 font-semibold">{h}</th>
                      ))}
                      {headers.length > 10 ? <th className="px-3 py-2 font-semibold text-slate-400">…</th> : null}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.slice(0, 30).map((r, idx) => (
                      <tr key={idx} className="border-t border-slate-100">
                        {headers.slice(0, 10).map((h) => (
                          <td key={h} className="max-w-[260px] truncate px-3 py-2 text-slate-800">{safeStr(r[h])}</td>
                        ))}
                        {headers.length > 10 ? <td className="px-3 py-2 text-slate-400">…</td> : null}
                      </tr>
                    ))}
                  </tbody>
                </table>
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