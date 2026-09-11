import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload, CheckCircle2, AlertTriangle, XCircle, Download, ListChecks,
  LayoutGrid, Settings2, Trash2, FileSpreadsheet, Search, Info, RefreshCw
} from "lucide-react";

/* ============================================================
   DATE DE REFERINȚĂ  — extrase din
   "LOCURI_DE_CAZARE_2026-2027_Criterii_Academice_Sociale.xlsx"
   (toate totalurile verificate)
   ============================================================ */

const DORMS = ["C12", "C13", "C15", "C16", "C17", "DREPT", "CAMELIA", "BRD", "Renașterii 2"];

const FACULTIES = ["FAD", "FCBG", "FD", "FEFS", "FEAA", "FFM", "FLIFT", "FI", "FMT", "FSAS", "FPSE", "FSGC"];

// Locuri pe criterii ACADEMICE: facultate -> { cămin: număr locuri }
const ACADEMIC = {
  FAD:  { C12: 20, C15: 27, C16: 15, CAMELIA: 6, BRD: 44, "Renașterii 2": 25 },
  FCBG: { C13: 39, C15: 5, C16: 23, C17: 7, BRD: 50, "Renașterii 2": 34 },
  FD:   { C15: 25, C16: 26, DREPT: 74, BRD: 21, "Renașterii 2": 50 },
  FEFS: { C13: 26, C15: 5, C16: 24, C17: 7, BRD: 30, "Renașterii 2": 35 },
  FEAA: { C12: 179, C15: 50, C16: 120, BRD: 87, "Renașterii 2": 142 },
  FFM:  { C13: 17, C15: 5, C16: 14, BRD: 16, "Renașterii 2": 15 },
  FLIFT:{ C12: 51, C16: 44, C17: 14, CAMELIA: 15, BRD: 62, "Renașterii 2": 53 },
  FI:   { C13: 56, C15: 20, C16: 36, C17: 5, CAMELIA: 5, BRD: 36, "Renașterii 2": 48 },
  FMT:  { C13: 8, C16: 10, C17: 2, CAMELIA: 6, BRD: 20, "Renașterii 2": 11 },
  FSAS: { C13: 16, C15: 40, C16: 11, CAMELIA: 2, BRD: 14, "Renașterii 2": 32 },
  FPSE: { C13: 47, C15: 64, C16: 30, C17: 6, CAMELIA: 5, BRD: 31, "Renașterii 2": 80 },
  FSGC: { C12: 26, C13: 47, C15: 16, C16: 28, C17: 4, CAMELIA: 14, BRD: 43, "Renașterii 2": 49 },
};

// Locuri pe criterii SOCIALE (PNRR): facultate -> { cămin: număr locuri }
const SOCIAL = {
  FAD:  { C15: 13, CAMELIA: 6 },
  FCBG: { C17: 24 },
  FD:   { C15: 4, DREPT: 56 },
  FEFS: { C17: 24 },
  FEAA: { C15: 105 },
  FFM:  { C15: 23 },
  FLIFT:{ C17: 37, CAMELIA: 15 },
  FI:   { C17: 35, CAMELIA: 9 },
  FMT:  { C15: 5, C17: 6, CAMELIA: 7 },
  FSAS: { C15: 22, C17: 3, CAMELIA: 6 },
  FPSE: { C15: 30, C17: 6, CAMELIA: 9 },
  FSGC: { C15: 6, C17: 15, CAMELIA: 17 },
};

const TOTAL_ACADEMIC = Object.values(ACADEMIC).reduce((s, d) => s + Object.values(d).reduce((a, b) => a + b, 0), 0);
const TOTAL_SOCIAL = Object.values(SOCIAL).reduce((s, d) => s + Object.values(d).reduce((a, b) => a + b, 0), 0);

/* ============================================================
   HELPERE
   ============================================================ */

const norm = (s) =>
  (s == null ? "" : String(s)).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();

// tabel de căutare pentru cămine (nume canonic din fișierul de locuri)
const DORM_LOOKUP = (() => {
  const m = {};
  DORMS.forEach((d) => { m[norm(d)] = d; });
  m["CAMELIEI"] = "CAMELIA"; // alias implicit (macheta scrie "Cameliei")
  return m;
})();

function canonicalDorm(raw, aliases) {
  if (raw == null || String(raw).trim() === "") return null;
  const n = norm(raw);
  if (DORM_LOOKUP[n]) return DORM_LOOKUP[n];
  if (aliases && aliases[n]) return aliases[n];
  return null; // cămin necunoscut
}

// Regulă naționalitate: român dacă "Tip loc curent" începe cu R (RB, RT, RPCB, ...)
function isRomanian(tip) {
  return norm(tip).startsWith("R");
}

function normCriteriu(c) {
  const n = norm(c);
  if (n.startsWith("SOC")) return "social";
  if (n.startsWith("ACAD")) return "academic";
  return null;
}

// găsește indexul unei coloane după potriviri de cuvinte-cheie
function findCol(headerRow, keywords) {
  for (let i = 0; i < headerRow.length; i++) {
    const h = norm(headerRow[i]);
    if (keywords.some((k) => h.includes(k))) return i;
  }
  return -1;
}

function parseMacheta(aoa) {
  // caută rândul de antet (conține "NUME" și "CRITERIU" sau "CAMIN")
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const row = aoa[i].map((x) => norm(x));
    const joined = row.join("|");
    if (joined.includes("NUME") && (joined.includes("CRITERIU") || joined.includes("CAMIN"))) {
      hdrIdx = i; break;
    }
  }
  if (hdrIdx === -1) throw new Error("Nu am găsit antetul machetei (rând cu «Nume» și «Criteriu»/«Cămin»).");

  const header = aoa[hdrIdx];
  const cNr = findCol(header, ["CRT"]);
  const cNume = findCol(header, ["NUME"]);
  const cEmail = findCol(header, ["EMAIL", "MAIL"]);
  const cCamin = findCol(header, ["CAMIN"]);
  const cTip = findCol(header, ["TIP LOC", "TIP"]);
  const cCrit = findCol(header, ["CRITERIU"]);

  const missing = [];
  if (cNume === -1) missing.push("Nume&Prenume");
  if (cCamin === -1) missing.push("Cămin Atribuit");
  if (cCrit === -1) missing.push("Criteriu loc cazare");
  if (missing.length) throw new Error("Lipsesc coloane obligatorii: " + missing.join(", "));

  const rows = [];
  for (let i = hdrIdx + 1; i < aoa.length; i++) {
    const r = aoa[i];
    const nume = cNume > -1 ? String(r[cNume] ?? "").trim() : "";
    if (!nume) continue; // sar peste rânduri goale
    rows.push({
      nr: cNr > -1 ? String(r[cNr] ?? "").trim() : "",
      nume,
      email: cEmail > -1 ? String(r[cEmail] ?? "").trim() : "",
      camin: cCamin > -1 ? String(r[cCamin] ?? "").trim() : "",
      tip: cTip > -1 ? String(r[cTip] ?? "").trim() : "",
      criteriu: cCrit > -1 ? String(r[cCrit] ?? "").trim() : "",
    });
  }
  return rows;
}

// validează un rând pentru o facultate dată -> { errors:[], warnings:[] }
function validateRow(row, faculty, aliases) {
  const errors = [];
  const warnings = [];
  const crit = normCriteriu(row.criteriu);

  if (!crit) warnings.push(`Criteriu necompletat sau necunoscut ("${row.criteriu || "gol"}")`);
  if (!row.email) warnings.push("Lipsește adresa de email");
  if (!row.camin) warnings.push("Lipsește căminul atribuit");

  const dorm = canonicalDorm(row.camin, aliases);
  if (row.camin && dorm === null) warnings.push(`Cămin necunoscut: "${row.camin}" — mapează-l în Setări`);

  if (crit === "social") {
    // 1) trebuie să fie român
    if (!isRomanian(row.tip)) {
      errors.push(`Caz social: trebuie să fie student ROMÂN (tip loc: "${row.tip || "gol"}")`);
    }
    // 2) căminul trebuie să fie printre locurile sociale ale facultății
    const allowed = SOCIAL[faculty] || {};
    const allowedList = Object.keys(allowed);
    if (dorm && !allowed[dorm]) {
      errors.push(
        `Caz social în ${dorm}: ${faculty} are locuri sociale doar în ${allowedList.length ? allowedList.join(", ") : "(niciun cămin)"}`
      );
    }
  }
  return { errors, warnings };
}

function downloadWorkbook(wb, filename) {
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ============================================================
   PALETĂ / STIL
   ============================================================ */

const C = {
  bg: "#F4F6F5",
  surface: "#FFFFFF",
  ink: "#16333F",
  inkSoft: "#3D5760",
  muted: "#6A7B80",
  border: "#D8E0E1",
  accent: "#2E7D8A",
  accentDark: "#1F5B66",
  okBg: "#E6F4EC", okFg: "#1E7A46",
  warnBg: "#FBF0D8", warnFg: "#8A5A00",
  errBg: "#FCE9E7", errFg: "#B42318",
};

const CSS = `
* { box-sizing: border-box; }
.vc-root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: ${C.ink}; background: ${C.bg}; min-height: 100%; }
.vc-num { font-variant-numeric: tabular-nums; }
.vc-btn { border: 1px solid ${C.border}; background: ${C.surface}; color: ${C.ink};
  padding: 8px 14px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer;
  display: inline-flex; align-items: center; gap: 8px; transition: background .12s, border-color .12s; }
.vc-btn:hover { background: #EEF3F3; }
.vc-btn:disabled { opacity: .45; cursor: not-allowed; }
.vc-btn-primary { background: ${C.accent}; border-color: ${C.accent}; color: #fff; }
.vc-btn-primary:hover { background: ${C.accentDark}; }
.vc-btn-danger { color: ${C.errFg}; border-color: #EEC9C4; background: #FFF7F6; }
.vc-btn-danger:hover { background: ${C.errBg}; }
.vc-tab { border: none; background: transparent; padding: 10px 4px; margin-right: 22px; font-size: 15px;
  font-weight: 600; color: ${C.muted}; cursor: pointer; border-bottom: 2px solid transparent; }
.vc-tab.active { color: ${C.ink}; border-bottom-color: ${C.accent}; }
.vc-card { background: ${C.surface}; border: 1px solid ${C.border}; border-radius: 12px; }
.vc-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.vc-table th { text-align: left; font-weight: 700; color: ${C.inkSoft}; padding: 9px 10px;
  border-bottom: 2px solid ${C.border}; position: sticky; top: 0; background: ${C.surface}; z-index: 1; }
.vc-table td { padding: 8px 10px; border-bottom: 1px solid #EDF1F1; vertical-align: top; }
.vc-table tr:hover td { background: #F7FAFA; }
.vc-badge { display: inline-flex; align-items: center; gap: 5px; padding: 2px 9px; border-radius: 999px;
  font-size: 12px; font-weight: 700; white-space: nowrap; }
.vc-input { border: 1px solid ${C.border}; border-radius: 8px; padding: 8px 11px; font-size: 14px;
  color: ${C.ink}; background: #fff; outline: none; }
.vc-input:focus { border-color: ${C.accent}; box-shadow: 0 0 0 3px rgba(46,125,138,.14); }
.vc-select { appearance: none; }
.vc-matrix td, .vc-matrix th { text-align: center; padding: 6px 8px; border: 1px solid ${C.border}; }
.vc-matrix th:first-child, .vc-matrix td:first-child { text-align: left; font-weight: 700; }
.vc-scroll { overflow: auto; }
`;

/* ============================================================
   COMPONENTE MICI
   ============================================================ */

function Badge({ kind, children }) {
  const map = {
    ok: { bg: C.okBg, fg: C.okFg, Icon: CheckCircle2 },
    warn: { bg: C.warnBg, fg: C.warnFg, Icon: AlertTriangle },
    err: { bg: C.errBg, fg: C.errFg, Icon: XCircle },
    neutral: { bg: "#EEF3F3", fg: C.inkSoft, Icon: null },
  };
  const s = map[kind] || map.neutral;
  const Icon = s.Icon;
  return (
    <span className="vc-badge" style={{ background: s.bg, color: s.fg }}>
      {Icon ? <Icon size={13} /> : null}
      {children}
    </span>
  );
}

function Kpi({ label, value, sub, accent }) {
  return (
    <div className="vc-card" style={{ padding: "16px 18px", flex: "1 1 0", minWidth: 150 }}>
      <div style={{ fontSize: 12.5, color: C.muted, fontWeight: 600, marginBottom: 6 }}>{label}</div>
      <div className="vc-num" style={{ fontSize: 28, fontWeight: 800, color: accent || C.ink, lineHeight: 1 }}>
        {value}
      </div>
      {sub ? <div className="vc-num" style={{ fontSize: 12.5, color: C.muted, marginTop: 5 }}>{sub}</div> : null}
    </div>
  );
}

/* ============================================================
   APP
   ============================================================ */

export default function App() {
  const [machete, setMachete] = useState({}); // faculty -> { faculty, filename, uploadedAt, rows }
  const [aliases, setAliases] = useState({}); // norm(raw) -> canonical dorm
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("upload");

  // persistență (window.storage, cu fallback în memorie)
  useEffect(() => {
    (async () => {
      try {
        if (typeof window !== "undefined" && window.storage) {
          try { const m = await window.storage.get("cazare_machete"); if (m?.value) setMachete(JSON.parse(m.value)); } catch (e) {}
          try { const a = await window.storage.get("cazare_aliases"); if (a?.value) setAliases(JSON.parse(a.value)); } catch (e) {}
        }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);
  useEffect(() => {
    if (!loaded) return;
    (async () => { try { if (window.storage) await window.storage.set("cazare_machete", JSON.stringify(machete)); } catch (e) {} })();
  }, [machete, loaded]);
  useEffect(() => {
    if (!loaded) return;
    (async () => { try { if (window.storage) await window.storage.set("cazare_aliases", JSON.stringify(aliases)); } catch (e) {} })();
  }, [aliases, loaded]);

  // toate rândurile centralizate + validare
  const allRows = useMemo(() => {
    const out = [];
    Object.values(machete).forEach((m) => {
      m.rows.forEach((r) => {
        const v = validateRow(r, m.faculty, aliases);
        out.push({ ...r, faculty: m.faculty, ...v, status: v.errors.length ? "err" : v.warnings.length ? "warn" : "ok" });
      });
    });
    return out;
  }, [machete, aliases]);

  const problems = allRows.filter((r) => r.status === "err").length;
  const warns = allRows.filter((r) => r.status === "warn").length;

  // ocupare
  const occ = useMemo(() => {
    const acc = { academic: {}, social: {}, unknown: {} };
    Object.values(machete).forEach((m) => {
      m.rows.forEach((r) => {
        const crit = normCriteriu(r.criteriu);
        if (crit !== "academic" && crit !== "social") return;
        const dorm = canonicalDorm(r.camin, aliases);
        const fac = m.faculty;
        if (!dorm) {
          acc.unknown[r.camin || "(gol)"] = (acc.unknown[r.camin || "(gol)"] || 0) + 1;
          return;
        }
        acc[crit][fac] = acc[crit][fac] || {};
        acc[crit][fac][dorm] = (acc[crit][fac][dorm] || 0) + 1;
      });
    });
    return acc;
  }, [machete, aliases]);

  const occTotal = (kind) =>
    Object.values(occ[kind]).reduce((s, d) => s + Object.values(d).reduce((a, b) => a + b, 0), 0);

  // cămine necunoscute găsite (pentru Setări)
  const unknownDorms = useMemo(() => {
    const set = new Map();
    Object.values(machete).forEach((m) =>
      m.rows.forEach((r) => {
        if (r.camin && canonicalDorm(r.camin, aliases) === null) {
          set.set(norm(r.camin), r.camin);
        }
      })
    );
    return Array.from(set.entries()); // [normKey, rawExample]
  }, [machete, aliases]);

  return (
    <div className="vc-root">
      <style>{CSS}</style>
      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "24px 20px 60px" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <FileSpreadsheet size={26} color={C.accent} />
              <h1 style={{ fontSize: 24, fontWeight: 800, margin: 0, letterSpacing: "-0.01em" }}>
                Verificare & centralizare cazare
              </h1>
            </div>
            <p style={{ margin: "6px 0 0", color: C.muted, fontSize: 14 }}>
              An universitar 2026–2027 · încarci macheta fiecărei facultăți, o verific, o adaug în centralizator.
            </p>
          </div>
        </div>

        {/* KPI */}
        <div style={{ display: "flex", gap: 12, marginTop: 20, flexWrap: "wrap" }}>
          <Kpi label="Facultăți încărcate" value={`${Object.keys(machete).length}/12`} />
          <Kpi label="Studenți centralizați" value={allRows.length} />
          <Kpi label="Locuri academice ocupate" value={occTotal("academic")} sub={`din ${TOTAL_ACADEMIC} alocate`} accent={C.accent} />
          <Kpi label="Locuri sociale ocupate" value={occTotal("social")} sub={`din ${TOTAL_SOCIAL} alocate`} accent={C.accent} />
          <Kpi label="Probleme detectate" value={problems} sub={warns ? `${warns} avertismente` : "0 avertismente"} accent={problems ? C.errFg : C.okFg} />
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", marginTop: 24, borderBottom: `1px solid ${C.border}`, flexWrap: "wrap" }}>
          <button className={`vc-tab ${tab === "upload" ? "active" : ""}`} onClick={() => setTab("upload")}>
            <span style={{ display: "inline-flex", gap: 7, alignItems: "center" }}><Upload size={16} /> Încărcare machetă</span>
          </button>
          <button className={`vc-tab ${tab === "centralizator" ? "active" : ""}`} onClick={() => setTab("centralizator")}>
            <span style={{ display: "inline-flex", gap: 7, alignItems: "center" }}><ListChecks size={16} /> Centralizator</span>
          </button>
          <button className={`vc-tab ${tab === "ocupare" ? "active" : ""}`} onClick={() => setTab("ocupare")}>
            <span style={{ display: "inline-flex", gap: 7, alignItems: "center" }}><LayoutGrid size={16} /> Ocupare locuri</span>
          </button>
          <button className={`vc-tab ${tab === "setari" ? "active" : ""}`} onClick={() => setTab("setari")}>
            <span style={{ display: "inline-flex", gap: 7, alignItems: "center" }}><Settings2 size={16} /> Setări</span>
          </button>
        </div>

        <div style={{ marginTop: 22 }}>
          {tab === "upload" && <UploadTab machete={machete} setMachete={setMachete} aliases={aliases} />}
          {tab === "centralizator" && <CentralizatorTab allRows={allRows} machete={machete} aliases={aliases} />}
          {tab === "ocupare" && <OcupareTab occ={occ} />}
          {tab === "setari" && (
            <SetariTab aliases={aliases} setAliases={setAliases} unknownDorms={unknownDorms} machete={machete} setMachete={setMachete} />
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   TAB: ÎNCĂRCARE
   ============================================================ */

function UploadTab({ machete, setMachete, aliases }) {
  const [faculty, setFaculty] = useState("");
  const [report, setReport] = useState(null); // { faculty, filename, rows, summary }
  const [err, setErr] = useState("");
  const fileRef = useRef(null);

  async function onFile(e) {
    setErr("");
    const file = e.target.files?.[0];
    if (!file) return;
    if (!faculty) { setErr("Alege întâi facultatea, apoi încarcă macheta."); e.target.value = ""; return; }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
      const rows = parseMacheta(aoa);
      if (!rows.length) throw new Error("Macheta nu conține niciun student.");

      const validated = rows.map((r) => ({ ...r, ...validateRow(r, faculty, aliases) }));
      const nErr = validated.filter((r) => r.errors.length).length;
      const nWarn = validated.filter((r) => !r.errors.length && r.warnings.length).length;
      const social = rows.filter((r) => normCriteriu(r.criteriu) === "social").length;
      const academic = rows.filter((r) => normCriteriu(r.criteriu) === "academic").length;

      setReport({
        faculty, filename: file.name, rows: validated,
        summary: { total: rows.length, academic, social, nErr, nWarn },
      });
    } catch (ex) {
      setErr(ex.message || "Nu am putut citi fișierul.");
      setReport(null);
    }
    e.target.value = "";
  }

  function commit() {
    if (!report) return;
    setMachete((prev) => ({
      ...prev,
      [report.faculty]: {
        faculty: report.faculty,
        filename: report.filename,
        uploadedAt: new Date().toISOString(),
        rows: report.rows.map(({ errors, warnings, ...r }) => r), // stocăm datele brute; validăm live
      },
    }));
    setReport(null);
    setFaculty("");
  }

  const already = faculty && machete[faculty];

  return (
    <div>
      <div className="vc-card" style={{ padding: 20 }}>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div>
            <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: C.inkSoft, marginBottom: 6 }}>
              Facultate
            </label>
            <select
              className="vc-input vc-select"
              style={{ minWidth: 180 }}
              value={faculty}
              onChange={(e) => setFaculty(e.target.value)}
            >
              <option value="">— alege facultatea —</option>
              {FACULTIES.map((f) => (
                <option key={f} value={f}>
                  {f}{machete[f] ? "  ✓ (încărcată)" : ""}
                </option>
              ))}
            </select>
          </div>
          <div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={onFile} style={{ display: "none" }} />
            <button className="vc-btn vc-btn-primary" disabled={!faculty} onClick={() => fileRef.current?.click()}>
              <Upload size={16} /> Alege macheta (.xlsx)
            </button>
          </div>
          {already ? (
            <div style={{ fontSize: 13, color: C.warnFg, display: "inline-flex", alignItems: "center", gap: 6 }}>
              <AlertTriangle size={15} /> {faculty} are deja o machetă — reîncărcarea o va înlocui.
            </div>
          ) : null}
        </div>
        {err ? (
          <div style={{ marginTop: 14, padding: "10px 12px", background: C.errBg, color: C.errFg, borderRadius: 8, fontSize: 13.5, fontWeight: 600 }}>
            <XCircle size={15} style={{ verticalAlign: "-2px", marginRight: 6 }} />{err}
          </div>
        ) : null}
      </div>

      {report ? <UploadReport report={report} onCommit={commit} onCancel={() => setReport(null)} /> : (
        <UploadedList machete={machete} />
      )}
    </div>
  );
}

function UploadReport({ report, onCommit, onCancel }) {
  const { summary } = report;
  const clean = summary.nErr === 0;
  const problematic = report.rows.filter((r) => r.errors.length || r.warnings.length);

  return (
    <div className="vc-card" style={{ padding: 20, marginTop: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontSize: 17, fontWeight: 800 }}>
            Verificare machetă · {report.faculty}
          </div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 3 }}>{report.filename}</div>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="vc-btn" onClick={onCancel}>Renunță</button>
          <button className="vc-btn vc-btn-primary" onClick={onCommit}>
            <CheckCircle2 size={16} /> Adaugă în centralizator
          </button>
        </div>
      </div>

      {/* banner notificare */}
      <div
        style={{
          marginTop: 16, padding: "12px 14px", borderRadius: 9, fontSize: 14, fontWeight: 600,
          background: clean ? C.okBg : C.errBg, color: clean ? C.okFg : C.errFg,
          display: "flex", alignItems: "center", gap: 9,
        }}
      >
        {clean ? <CheckCircle2 size={18} /> : <XCircle size={18} />}
        {clean
          ? `Totul e în regulă: ${summary.total} studenți, fără erori.`
          : `Atenție: ${summary.nErr} ${summary.nErr === 1 ? "problemă" : "probleme"} de rezolvat în această machetă.`}
      </div>

      <div style={{ display: "flex", gap: 12, marginTop: 14, flexWrap: "wrap" }}>
        <MiniStat label="Total" value={summary.total} />
        <MiniStat label="Academic" value={summary.academic} />
        <MiniStat label="Social" value={summary.social} />
        <MiniStat label="Erori" value={summary.nErr} tone={summary.nErr ? "err" : "ok"} />
        <MiniStat label="Avertismente" value={summary.nWarn} tone={summary.nWarn ? "warn" : "ok"} />
      </div>

      {problematic.length ? (
        <div style={{ marginTop: 18 }}>
          <div style={{ fontSize: 13.5, fontWeight: 700, color: C.inkSoft, marginBottom: 8 }}>
            Rânduri de verificat ({problematic.length})
          </div>
          <div className="vc-scroll" style={{ maxHeight: 380, border: `1px solid ${C.border}`, borderRadius: 9 }}>
            <table className="vc-table">
              <thead>
                <tr>
                  <th>Nume</th><th>Cămin</th><th>Tip</th><th>Criteriu</th><th>Ce e în neregulă</th>
                </tr>
              </thead>
              <tbody>
                {problematic.map((r, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 600 }}>{r.nume}</td>
                    <td>{r.camin}</td>
                    <td className="vc-num">{r.tip}</td>
                    <td>{r.criteriu}</td>
                    <td>
                      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        {r.errors.map((e, j) => <Badge key={"e" + j} kind="err">{e}</Badge>)}
                        {r.warnings.map((w, j) => <Badge key={"w" + j} kind="warn">{w}</Badge>)}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div style={{ marginTop: 16, color: C.muted, fontSize: 13.5 }}>Niciun rând problematic. </div>
      )}
    </div>
  );
}

function MiniStat({ label, value, tone }) {
  const color = tone === "err" ? C.errFg : tone === "warn" ? C.warnFg : tone === "ok" ? C.okFg : C.ink;
  return (
    <div style={{ border: `1px solid ${C.border}`, borderRadius: 9, padding: "9px 14px", minWidth: 92 }}>
      <div style={{ fontSize: 12, color: C.muted, fontWeight: 600 }}>{label}</div>
      <div className="vc-num" style={{ fontSize: 20, fontWeight: 800, color }}>{value}</div>
    </div>
  );
}

function UploadedList({ machete }) {
  const list = Object.values(machete);
  if (!list.length) {
    return (
      <div className="vc-card" style={{ padding: 32, marginTop: 18, textAlign: "center", color: C.muted }}>
        <Info size={22} style={{ opacity: 0.6 }} />
        <div style={{ marginTop: 8, fontSize: 14.5 }}>Nicio machetă încărcată încă. Alege o facultate și încarcă fișierul ei.</div>
      </div>
    );
  }
  return (
    <div className="vc-card" style={{ padding: 20, marginTop: 18 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.inkSoft, marginBottom: 10 }}>Machete încărcate</div>
      <table className="vc-table">
        <thead><tr><th>Facultate</th><th>Fișier</th><th>Studenți</th><th>Încărcat</th></tr></thead>
        <tbody>
          {list.map((m) => (
            <tr key={m.faculty}>
              <td style={{ fontWeight: 700 }}>{m.faculty}</td>
              <td style={{ color: C.muted }}>{m.filename}</td>
              <td className="vc-num">{m.rows.length}</td>
              <td style={{ color: C.muted }}>{new Date(m.uploadedAt).toLocaleString("ro-RO")}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ============================================================
   TAB: CENTRALIZATOR
   ============================================================ */

function CentralizatorTab({ allRows, machete, aliases }) {
  const [q, setQ] = useState("");
  const [fFac, setFFac] = useState("");
  const [fStatus, setFStatus] = useState("");
  const [showAll, setShowAll] = useState(false);

  const filtered = useMemo(() => {
    const qq = norm(q);
    return allRows.filter((r) => {
      if (fFac && r.faculty !== fFac) return false;
      if (fStatus && r.status !== fStatus) return false;
      if (qq) {
        const hay = norm(`${r.nume} ${r.email} ${r.camin} ${r.tip} ${r.criteriu} ${r.faculty}`);
        if (!hay.includes(qq)) return false;
      }
      return true;
    });
  }, [allRows, q, fFac, fStatus]);

  const shown = showAll ? filtered : filtered.slice(0, 500);

  function exportAll() {
    const wb = XLSX.utils.book_new();

    // Sheet 1 — Centralizator (toate machetele combinate)
    const head = ["Facultate", "Nr.", "Nume & Prenume", "Adresă email", "Cămin Atribuit", "Tip loc curent", "Criteriu", "Status", "Probleme"];
    const data = allRows.map((r, i) => [
      r.faculty, i + 1, r.nume, r.email, r.camin, r.tip, r.criteriu,
      r.status === "err" ? "EROARE" : r.status === "warn" ? "Avertisment" : "OK",
      [...r.errors, ...r.warnings].join(" | "),
    ]);
    const ws1 = XLSX.utils.aoa_to_sheet([head, ...data]);
    ws1["!cols"] = [{ wch: 9 }, { wch: 5 }, { wch: 34 }, { wch: 30 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, ws1, "Centralizator");

    // Sheet 2 & 3 — Ocupare (cu aliasurile de cămine aplicate)
    const buildMatrix = (kind, alloc) => {
      const rows = [["Facultate", ...DORMS, "Total ocupat", "Total alocat"]];
      FACULTIES.forEach((f) => {
        const line = [f];
        let occSum = 0, allocSum = 0;
        DORMS.forEach((d) => {
          const a = (alloc[f] && alloc[f][d]) || 0;
          const o = allRows.filter(
            (r) => r.faculty === f && normCriteriu(r.criteriu) === kind && canonicalDorm(r.camin, aliases) === d
          ).length;
          occSum += o; allocSum += a;
          line.push(a || o ? `${o}/${a}` : "");
        });
        line.push(occSum); line.push(allocSum);
        rows.push(line);
      });
      return XLSX.utils.aoa_to_sheet(rows);
    };
    XLSX.utils.book_append_sheet(wb, buildMatrix("academic", ACADEMIC), "Ocupare academic");
    XLSX.utils.book_append_sheet(wb, buildMatrix("social", SOCIAL), "Ocupare social");

    downloadWorkbook(wb, "Centralizator_cazare_2026-2027.xlsx");
  }

  return (
    <div>
      <div className="vc-card" style={{ padding: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ position: "relative" }}>
            <Search size={15} color={C.muted} style={{ position: "absolute", left: 10, top: 10 }} />
            <input className="vc-input" style={{ paddingLeft: 30, width: 220 }} placeholder="Caută nume, email, cămin…" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
          <select className="vc-input vc-select" value={fFac} onChange={(e) => setFFac(e.target.value)}>
            <option value="">Toate facultățile</option>
            {FACULTIES.filter((f) => machete[f]).map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
          <select className="vc-input vc-select" value={fStatus} onChange={(e) => setFStatus(e.target.value)}>
            <option value="">Toate stările</option>
            <option value="ok">Doar OK</option>
            <option value="warn">Doar avertismente</option>
            <option value="err">Doar erori</option>
          </select>
        </div>
        <button className="vc-btn vc-btn-primary" disabled={!allRows.length} onClick={exportAll}>
          <Download size={16} /> Export combinat (.xlsx)
        </button>
      </div>

      <div style={{ margin: "10px 2px", fontSize: 13, color: C.muted }}>
        {filtered.length} rezultate{filtered.length !== allRows.length ? ` (din ${allRows.length})` : ""}
      </div>

      {allRows.length === 0 ? (
        <div className="vc-card" style={{ padding: 32, textAlign: "center", color: C.muted }}>
          Încă nu e nimic centralizat. Adaugă cel puțin o machetă din tabul „Încărcare".
        </div>
      ) : (
        <div className="vc-card vc-scroll" style={{ maxHeight: 560 }}>
          <table className="vc-table">
            <thead>
              <tr>
                <th style={{ width: 34 }}></th>
                <th>Fac.</th><th>Nume & Prenume</th><th>Email</th><th>Cămin</th><th>Tip</th><th>Criteriu</th><th>Probleme</th>
              </tr>
            </thead>
            <tbody>
              {shown.map((r, i) => (
                <tr key={i}>
                  <td>
                    {r.status === "err" ? <XCircle size={16} color={C.errFg} />
                      : r.status === "warn" ? <AlertTriangle size={16} color={C.warnFg} />
                      : <CheckCircle2 size={16} color={C.okFg} />}
                  </td>
                  <td style={{ fontWeight: 700 }}>{r.faculty}</td>
                  <td style={{ fontWeight: 600 }}>{r.nume}</td>
                  <td style={{ color: C.muted }}>{r.email}</td>
                  <td>{r.camin}</td>
                  <td className="vc-num">{r.tip}</td>
                  <td>{r.criteriu}</td>
                  <td>
                    <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                      {r.errors.map((e, j) => <Badge key={"e" + j} kind="err">{e}</Badge>)}
                      {r.warnings.map((w, j) => <Badge key={"w" + j} kind="warn">{w}</Badge>)}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!showAll && filtered.length > 500 ? (
            <div style={{ padding: 14, textAlign: "center" }}>
              <button className="vc-btn" onClick={() => setShowAll(true)}>Afișează toate cele {filtered.length}</button>
            </div>
          ) : null}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   TAB: OCUPARE
   ============================================================ */

function Matrix({ title, alloc, occKind }) {
  // occKind: obiect faculty -> {dorm: count}
  const facRows = FACULTIES.filter((f) => alloc[f] || occKind[f]);
  return (
    <div className="vc-card" style={{ padding: 18, marginBottom: 20 }}>
      <div style={{ fontSize: 15, fontWeight: 800, marginBottom: 12 }}>{title}</div>
      <div className="vc-scroll">
        <table className="vc-matrix vc-num" style={{ fontSize: 12.5, minWidth: 760 }}>
          <thead>
            <tr>
              <th>Facultate</th>
              {DORMS.map((d) => <th key={d}>{d}</th>)}
              <th>Ocupat</th><th>Alocat</th><th>Liber</th>
            </tr>
          </thead>
          <tbody>
            {facRows.map((f) => {
              const a = alloc[f] || {};
              const o = occKind[f] || {};
              let occSum = 0, allocSum = 0;
              const cells = DORMS.map((d) => {
                const av = a[d] || 0;
                const ov = o[d] || 0;
                occSum += ov; allocSum += av;
                let bg = "transparent", fg = C.ink;
                if (av === 0 && ov === 0) return <td key={d} style={{ color: "#C6D0D1" }}>–</td>;
                if (ov > av) { bg = C.errBg; fg = C.errFg; }            // supra-ocupat
                else if (av > 0 && ov === av) { bg = C.okBg; fg = C.okFg; } // plin
                else if (av === 0 && ov > 0) { bg = C.warnBg; fg = C.warnFg; } // ocupat unde nu-s locuri
                return (
                  <td key={d} style={{ background: bg, color: fg, fontWeight: 600 }}>
                    {ov}/{av}
                  </td>
                );
              });
              const free = allocSum - occSum;
              return (
                <tr key={f}>
                  <td>{f}</td>
                  {cells}
                  <td style={{ fontWeight: 700 }}>{occSum}</td>
                  <td>{allocSum}</td>
                  <td style={{ color: free < 0 ? C.errFg : C.inkSoft, fontWeight: 700 }}>{free}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", gap: 16, marginTop: 10, flexWrap: "wrap", fontSize: 12, color: C.muted }}>
        <span><span style={{ background: C.okBg, color: C.okFg, padding: "1px 7px", borderRadius: 5, fontWeight: 700 }}>plin</span> ocupat = alocat</span>
        <span><span style={{ background: C.errBg, color: C.errFg, padding: "1px 7px", borderRadius: 5, fontWeight: 700 }}>depășit</span> ocupat &gt; alocat</span>
        <span><span style={{ background: C.warnBg, color: C.warnFg, padding: "1px 7px", borderRadius: 5, fontWeight: 700 }}>neprevăzut</span> ocupat unde nu-s locuri</span>
      </div>
    </div>
  );
}

function OcupareTab({ occ }) {
  const unk = Object.entries(occ.unknown);
  return (
    <div>
      <Matrix title="Ocupare pe criterii ACADEMICE (ocupat / alocat)" alloc={ACADEMIC} occKind={occ.academic} />
      <Matrix title="Ocupare pe criterii SOCIALE — PNRR (ocupat / alocat)" alloc={SOCIAL} occKind={occ.social} />
      {unk.length ? (
        <div className="vc-card" style={{ padding: 16, borderColor: "#EEDBB0", background: "#FFFBF2" }}>
          <div style={{ fontWeight: 700, color: C.warnFg, display: "flex", alignItems: "center", gap: 8 }}>
            <AlertTriangle size={16} /> Studenți în cămine neînregistrate (necontorizați mai sus)
          </div>
          <div style={{ marginTop: 8, fontSize: 13.5, color: C.inkSoft }}>
            {unk.map(([k, v]) => `${k}: ${v}`).join(" · ")}
          </div>
          <div style={{ marginTop: 6, fontSize: 12.5, color: C.muted }}>
            Mapează aceste cămine în tabul <b>Setări</b> ca să fie incluse în calcul.
          </div>
        </div>
      ) : null}
    </div>
  );
}

/* ============================================================
   TAB: SETĂRI
   ============================================================ */

function SetariTab({ aliases, setAliases, unknownDorms, machete, setMachete }) {
  function setAlias(key, dorm) {
    setAliases((prev) => {
      const next = { ...prev };
      if (dorm) next[key] = dorm; else delete next[key];
      return next;
    });
  }
  function removeMacheta(fac) {
    if (!window.confirm(`Ștergi macheta pentru ${fac}?`)) return;
    setMachete((prev) => { const n = { ...prev }; delete n[fac]; return n; });
  }
  function resetAll() {
    if (!window.confirm("Ștergi TOATE machetele centralizate? Acțiunea nu se poate anula.")) return;
    setMachete({});
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Aliasuri cămine */}
      <div className="vc-card" style={{ padding: 20 }}>
        <div style={{ fontSize: 15, fontWeight: 800 }}>Mapare cămine</div>
        <p style={{ fontSize: 13.5, color: C.muted, marginTop: 4 }}>
          Când în machetă apare un cămin scris altfel decât în fișierul de locuri, mapează-l aici la căminul corect.
          („Cameliei" → „CAMELIA" e deja mapat automat.)
        </p>
        {unknownDorms.length === 0 ? (
          <div style={{ fontSize: 13.5, color: C.okFg, display: "flex", alignItems: "center", gap: 7, marginTop: 8 }}>
            <CheckCircle2 size={16} /> Toate căminele din machetele încărcate sunt recunoscute.
          </div>
        ) : (
          <table className="vc-table" style={{ marginTop: 8 }}>
            <thead><tr><th>Cămin din machetă</th><th>Mapează la</th></tr></thead>
            <tbody>
              {unknownDorms.map(([key, raw]) => (
                <tr key={key}>
                  <td style={{ fontWeight: 600 }}>{raw}</td>
                  <td>
                    <select className="vc-input vc-select" value={aliases[key] || ""} onChange={(e) => setAlias(key, e.target.value)}>
                      <option value="">— nemapat —</option>
                      {DORMS.map((d) => <option key={d} value={d}>{d}</option>)}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Regula de naționalitate */}
      <div className="vc-card" style={{ padding: 20 }}>
        <div style={{ fontSize: 15, fontWeight: 800 }}>Regula pentru cazurile sociale</div>
        <ul style={{ fontSize: 13.5, color: C.inkSoft, marginTop: 8, paddingLeft: 20, lineHeight: 1.6 }}>
          <li><b>Naționalitate:</b> un caz social e considerat român dacă „Tip loc curent" începe cu <b>R</b> (RB, RT, RPCB, RPFB, RBRU, RBDIZ, RPT). Altfel → eroare.</li>
          <li><b>Cămin:</b> un caz social trebuie să fie într-un cămin unde facultatea are locuri sociale alocate (conform fișierului de locuri). Altfel → eroare.</li>
        </ul>
        <div style={{ fontSize: 12.5, color: C.muted, marginTop: 6 }}>
          Vrei altă regulă de naționalitate? Se poate ajusta ușor.
        </div>
      </div>

      {/* Gestionare machete */}
      <div className="vc-card" style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 800 }}>Machete stocate</div>
          <button className="vc-btn vc-btn-danger" disabled={!Object.keys(machete).length} onClick={resetAll}>
            <Trash2 size={15} /> Șterge tot
          </button>
        </div>
        {Object.keys(machete).length === 0 ? (
          <div style={{ fontSize: 13.5, color: C.muted, marginTop: 8 }}>Nicio machetă stocată.</div>
        ) : (
          <table className="vc-table" style={{ marginTop: 10 }}>
            <thead><tr><th>Facultate</th><th>Fișier</th><th>Studenți</th><th></th></tr></thead>
            <tbody>
              {Object.values(machete).map((m) => (
                <tr key={m.faculty}>
                  <td style={{ fontWeight: 700 }}>{m.faculty}</td>
                  <td style={{ color: C.muted }}>{m.filename}</td>
                  <td className="vc-num">{m.rows.length}</td>
                  <td style={{ textAlign: "right" }}>
                    <button className="vc-btn vc-btn-danger" onClick={() => removeMacheta(m.faculty)}>
                      <Trash2 size={14} /> Șterge
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
