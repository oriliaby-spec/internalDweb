import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  CalendarDays,
  Users,
  ClipboardList,
  Plus,
  Trash2,
  Check,
  ChevronRight,
  ChevronLeft,
  Upload,
  X,
  ArrowLeftRight,
  Search,
  StickyNote,
  Image as ImageIcon,
  FileSpreadsheet,
  Link as LinkIcon,
  AlertCircle,
} from "lucide-react";

/* ============================================================
   Design tokens
   Ink:      #1C2B33 (deep slate)
   Canvas:   #F5F6F4 (cool paper)
   Card:     #FFFFFF
   Side A:   #2F6F63 (clinical teal)
   Side A bg:#E7F1EE
   Side B:   #92552E (warm clay)
   Side B bg:#F5EAE1
   Amber(today): #C68A2E / bg #FCEFD8
   Weekend bg: #ECEDEA
   Danger: #B3492F
   ============================================================ */

const INK = "#1C2B33";
const CANVAS = "#F5F6F4";
const SIDE_A = "#2F6F63";
const SIDE_A_BG = "#E7F1EE";
const SIDE_B = "#92552E";
const SIDE_B_BG = "#F5EAE1";
const TODAY = "#C68A2E";
const TODAY_BG = "#FCEFD8";
const WEEKEND_BG = "#ECEDEA";
const DANGER = "#B3492F";
const LINE = "#DCDED9";
const PRESENT = "#5B6EA6";
const PRESENT_BG = "#E9EBF3";

const HEB_MONTHS = [
  "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
  "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר",
];
const HEB_DAYS_SHORT = ["א", "ב", "ג", "ד", "ה", "ו", "ש"];

const STORAGE_KEY = "ward-d-app-data-v1";

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function pad(n) {
  return String(n).length < 2 ? "0" + n : String(n);
}

function monthKey(year, month) {
  return `${year}-${pad(month + 1)}`;
}

function daysInMonth(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function emptyDayEntry() {
  return {
    sideA: { lead: "", residents: "", rotation: "", interns: "" },
    sideB: { lead: "", residents: "", rotation: "", interns: "" },
    shared: {
      ward: "",
      miku: "",
      onCall: "",
      paAssistant: "",
      erOnCall1: "",
      erOnCall2: "",
      erOnCall3: "",
      shortPresentation: "",
      longPresentation: "",
    },
  };
}

// Splits a multi-name cell like "רוני, דרור" or "רוני / דרור" into separate names
function splitNames(str) {
  if (!str) return [];
  return str
    .split(/[,\/\n\\،]+/)
    .map((s) => s.trim())
    .filter(Boolean);
}

// The residents actually scheduled to work on a given day (both sides,
// "מתמחה אחראי" + "מתמחים"), used to auto-populate who's eligible for
// ER on-call duty ("תורן מיון") that day - pulled straight from the table
// instead of a separate manually-maintained list.
function dayResidentPool(entry) {
  const names = [
    ...splitNames(entry?.sideA?.lead),
    ...splitNames(entry?.sideA?.residents),
    ...splitNames(entry?.sideB?.lead),
    ...splitNames(entry?.sideB?.residents),
  ];
  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b, "he"));
}

// Collects every resident name that has ever been typed into "מתמחה אחראי" /
// "מתמחים" fields (both sides, all months) - this is the department's known
// resident list, used to validate/constrain the ER on-call selection.
function getKnownResidents(schedule) {
  const set = new Set();
  Object.values(schedule || {}).forEach((monthData) => {
    Object.values(monthData || {}).forEach((entry) => {
      ["sideA", "sideB"].forEach((side) => {
        splitNames(entry?.[side]?.lead).forEach((n) => set.add(n));
        splitNames(entry?.[side]?.residents).forEach((n) => set.add(n));
      });
    });
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "he"));
}

// 0-indexed -> Excel-style column letters (0 -> A, 26 -> AA ...)
function colLetter(i) {
  let s = "";
  i += 1;
  while (i > 0) {
    const rem = (i - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

// Detects a Google Sheets share/edit link and extracts the spreadsheet ID
function parseGoogleSheetId(url) {
  const m = String(url || "").match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : null;
}

function googleXlsxExportUrl(id) {
  return `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
}

// Strips prefixes like "צד 1 -" and "ד"ר" from a detected leader-name cell,
// leaving just the doctor's name.
function extractLeaderName(raw) {
  let t = String(raw || "").trim();
  if (!t) return "";
  t = t.replace(/^צד\s*[0-9א-ת]+\s*[-–:]?\s*/, "");
  t = t.replace(/^ד["'׳״`]?ר\.?\s*/, "");
  return t.trim();
}

// Classifies a single header-cell's text into a known field category.
// Order matters: more specific keywords are checked before generic ones
// (e.g. "מיון 1"/"תורן מיון" must be caught before the generic "מחלקה"/
// "אחראי" rules, and a digit after "מיון" pins the exact ER on-call slot).
function classifyHeaderCell(raw) {
  const t = String(raw || "").trim();
  if (!t) return null;
  if (t.includes("תאריך")) return "date";
  const miyunDigit = t.match(/מיון\s*([1-3])\b/);
  if (miyunDigit) return "eroncall" + miyunDigit[1];
  if (t.includes("תורן") && t.includes("מיון")) return "eroncall";
  if (t.includes("עוזר")) return "paAssistant";
  if (t.includes("כונן") || (t.includes("תורן") && t.includes("מחלקה"))) return "oncall";
  if (t.includes("מחלקה")) return "ward";
  if (t.includes("מיק")) return "miku";
  if (t.includes("מיון")) return "ignore";
  if (t.includes("הצגה")) {
    if (t.includes("קצר")) return "presentShort";
    if (t.includes("ארוכ")) return "presentLong";
    return "ignore";
  }
  if (t.includes("אחראי")) return "lead";
  if (t.includes("סטאז")) return "intern";
  if (t.includes("רוטצי")) return "rotation";
  if (t.includes("מתמחה") || t.includes("מתמחים")) return "resident";
  return null;
}

// Scans the first N rows to find the most likely header row (most
// recognized keyword cells), then builds a column mapping automatically,
// merging repeated per-side columns (e.g. 3x "מתמחה") into single fields.
// Returns null if no convincing header row was found.
function autoDetectImport(rows) {
  const SCAN_LIMIT = Math.min(rows.length, 15);
  let bestRowIdx = -1;
  let bestScore = 0;
  for (let i = 0; i < SCAN_LIMIT; i++) {
    const row = rows[i] || [];
    const score = row.reduce((acc, cell) => (classifyHeaderCell(cell) ? acc + 1 : acc), 0);
    if (score > bestScore) {
      bestScore = score;
      bestRowIdx = i;
    }
  }
  if (bestRowIdx === -1 || bestScore < 3) return null;

  const headerRow = rows[bestRowIdx] || [];
  const classified = headerRow.map((cell, idx) => ({ idx, cat: classifyHeaderCell(cell) }));

  const leadCols = classified.filter((c) => c.cat === "lead").map((c) => c.idx);
  // Fall back to the horizontal midpoint of the header row when we can't
  // find two "מתמחה אחראי" columns to split on directly - this keeps side
  // A/B (and therefore leader-name) detection working even on sheets whose
  // wording differs slightly from the expected keywords.
  const fallbackSplit = headerRow.length ? Math.floor(headerRow.length / 2) : Infinity;
  const splitPoint = leadCols.length >= 2 ? leadCols[1] : fallbackSplit;
  const sideOf = (idx) => (idx < splitPoint ? "A" : "B");

  const pick = (cat) => classified.find((c) => c.cat === cat)?.idx;
  const pickAll = (cat) => classified.filter((c) => c.cat === cat).map((c) => c.idx);

  const dateCol = pick("date");
  const wardCol = pick("ward");
  const mikuCol = pick("miku");
  const oncallCol = pick("oncall");
  const paAssistantCol = pick("paAssistant");
  const eroncallGeneric = pickAll("eroncall");
  const eroncallCols = [pick("eroncall1"), pick("eroncall2"), pick("eroncall3")];
  let genericIdx = 0;
  for (let i = 0; i < 3; i++) {
    if (eroncallCols[i] === undefined && genericIdx < eroncallGeneric.length) {
      eroncallCols[i] = eroncallGeneric[genericIdx++];
    }
  }
  const presentShortCol = pick("presentShort");
  const presentLongCol = pick("presentLong");

  const rotationCols = pickAll("rotation");
  const rotationA = rotationCols.find((i) => sideOf(i) === "A");
  const rotationB = rotationCols.find((i) => sideOf(i) === "B");

  const residentCols = pickAll("resident");
  const residentsA = residentCols.filter((i) => sideOf(i) === "A");
  const residentsB = residentCols.filter((i) => sideOf(i) === "B");

  const internCols = pickAll("intern");
  const internsA = internCols.filter((i) => sideOf(i) === "A");
  const internsB = internCols.filter((i) => sideOf(i) === "B");

  // The senior doctor's name is usually written above the day-by-day
  // table itself (e.g. `צד 1 - ד"ר יעל מליגרום`), so scan the rows
  // preceding the detected header row for a "ד"ר ..." style cell.
  let leaderA;
  let leaderB;
  const DOCTOR_RE = /ד["'׳״`]?ר/;
  for (let i = 0; i < bestRowIdx; i++) {
    const row = rows[i] || [];
    row.forEach((cell, idx) => {
      const text = String(cell || "").trim();
      if (!text || !DOCTOR_RE.test(text) || text.length > 60) return;
      const side = sideOf(idx);
      if (side === "A" && leaderA === undefined) leaderA = extractLeaderName(text);
      if (side === "B" && leaderB === undefined) leaderB = extractLeaderName(text);
    });
  }

  const origColCount = rows.reduce((max, r) => Math.max(max, r.length), 0);

  // Build merged "virtual" columns for the multi-cell fields
  const mergeSpecs = [];
  if (residentsA.length) mergeSpecs.push({ field: "sideA.residents", cols: residentsA, label: "(מיזוג) צד א׳ · מתמחים" });
  if (internsA.length) mergeSpecs.push({ field: "sideA.interns", cols: internsA, label: "(מיזוג) צד א׳ · סטאז׳רים" });
  if (residentsB.length) mergeSpecs.push({ field: "sideB.residents", cols: residentsB, label: "(מיזוג) צד ב׳ · מתמחים" });
  if (internsB.length) mergeSpecs.push({ field: "sideB.interns", cols: internsB, label: "(מיזוג) צד ב׳ · סטאז׳רים" });

  const virtualLabels = {};
  mergeSpecs.forEach((spec, i) => {
    virtualLabels[origColCount + i] = spec.label;
  });

  const augmentedRows = rows.map((r) => [
    ...r,
    ...mergeSpecs.map((spec) => spec.cols.map((c) => (r[c] || "").toString().trim()).filter(Boolean).join(", ")),
  ]);

  const mapping = {};
  if (dateCol !== undefined) mapping.date = dateCol;
  if (leadCols[0] !== undefined) mapping["sideA.lead"] = leadCols[0];
  if (leadCols[1] !== undefined) mapping["sideB.lead"] = leadCols[1];
  if (rotationA !== undefined) mapping["sideA.rotation"] = rotationA;
  if (rotationB !== undefined) mapping["sideB.rotation"] = rotationB;
  if (wardCol !== undefined) mapping["shared.ward"] = wardCol;
  if (mikuCol !== undefined) mapping["shared.miku"] = mikuCol;
  if (oncallCol !== undefined) mapping["shared.onCall"] = oncallCol;
  if (paAssistantCol !== undefined) mapping["shared.paAssistant"] = paAssistantCol;
  if (eroncallCols[0] !== undefined) mapping["shared.erOnCall1"] = eroncallCols[0];
  if (eroncallCols[1] !== undefined) mapping["shared.erOnCall2"] = eroncallCols[1];
  if (eroncallCols[2] !== undefined) mapping["shared.erOnCall3"] = eroncallCols[2];
  if (presentShortCol !== undefined) mapping["shared.shortPresentation"] = presentShortCol;
  if (presentLongCol !== undefined) mapping["shared.longPresentation"] = presentLongCol;
  mergeSpecs.forEach((spec, i) => {
    mapping[spec.field] = origColCount + i;
  });

  return {
    headerRowIndex: bestRowIdx,
    augmentedRows,
    mapping,
    virtualLabels,
    origColCount,
    leaders: { a: leaderA, b: leaderB },
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.error("Failed to save", e);
  }
}

/* ---------------- Editable cell ---------------- */
function EditableField({ value, onChange, placeholder, multiline, accent, bold }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value || "");
  const ref = useRef(null);

  useEffect(() => {
    setDraft(value || "");
  }, [value]);

  useEffect(() => {
    if (editing && ref.current) {
      ref.current.focus();
      ref.current.select?.();
    }
  }, [editing]);

  const commit = () => {
    setEditing(false);
    if (draft !== value) onChange(draft);
  };

  if (editing) {
    return (
      <input
        ref={ref}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit();
          if (e.key === "Escape") {
            setDraft(value || "");
            setEditing(false);
          }
        }}
        placeholder={placeholder}
        className={`w-full min-w-[64px] rounded px-1.5 py-1 text-[13px] outline-none ${bold ? "font-bold" : ""}`}
        style={{
          border: `1.5px solid ${accent || INK}`,
          background: "#fff",
          color: INK,
        }}
      />
    );
  }

  return (
    <button
      onClick={() => setEditing(true)}
      className={`w-full min-w-[64px] rounded px-1.5 py-1 text-right text-[13px] transition-colors hover:bg-black/5 ${bold && value ? "font-bold" : ""}`}
      style={{ color: value ? INK : "#B5B8B2", minHeight: "28px" }}
      title="לחיצה לעריכה"
    >
      {value || placeholder || "—"}
    </button>
  );
}

/* ---------------- Schedule Tab ---------------- */
function ScheduleTab({
  schedule,
  setSchedule,
  leaders,
  setLeaders,
  notes,
  setNotes,
  year,
  month,
  setYear,
  setMonth,
}) {
  const mk = monthKey(year, month);
  const nDays = daysInMonth(year, month);
  const today = todayISO();
  const [showImport, setShowImport] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);

  const knownResidents = useMemo(() => getKnownResidents(schedule), [schedule]);

  const monthData = schedule[mk] || {};

  const clearMonth = () => {
    setSchedule((prev) => {
      const next = { ...prev };
      delete next[mk];
      return next;
    });
    setLeaders({ a: "", b: "" });
    setConfirmClear(false);
  };

  const updateDay = (day, section, field, val) => {
    setSchedule((prev) => {
      const cur = { ...(prev[mk] || {}) };
      const dayEntry = cur[day] ? JSON.parse(JSON.stringify(cur[day])) : emptyDayEntry();
      dayEntry[section][field] = val;
      cur[day] = dayEntry;
      return { ...prev, [mk]: cur };
    });
  };

  const shiftMonth = (delta) => {
    let m = month + delta;
    let y = year;
    if (m < 0) {
      m = 11;
      y -= 1;
    } else if (m > 11) {
      m = 0;
      y += 1;
    }
    setMonth(m);
    setYear(y);
  };

  const [newNote, setNewNote] = useState("");

  const days = Array.from({ length: nDays }, (_, i) => i + 1);

  return (
    <div className="flex flex-col gap-4">
      {/* Leaders + month nav */}
      <div className="flex flex-col gap-3 rounded-xl bg-white p-3 shadow-sm" style={{ border: `1px solid ${LINE}` }}>
        <div className="flex items-center justify-between">
          <button
            onClick={() => shiftMonth(1)}
            className="flex h-8 w-8 items-center justify-center rounded-full transition-colors hover:bg-black/5"
            style={{ color: INK }}
            aria-label="חודש הבא"
          >
            <ChevronRight size={18} />
          </button>
          <div className="text-center">
            <div className="text-[15px] font-bold" style={{ color: INK }}>
              {HEB_MONTHS[month]} {year}
            </div>
          </div>
          <button
            onClick={() => shiftMonth(-1)}
            className="flex h-8 w-8 items-center justify-center rounded-full transition-colors hover:bg-black/5"
            style={{ color: INK }}
            aria-label="חודש קודם"
          >
            <ChevronLeft size={18} />
          </button>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setShowImport(true)}
            className="flex items-center justify-center gap-1.5 rounded-lg px-3 py-1.5 text-[12px] font-semibold text-white"
            style={{ background: INK }}
          >
            <FileSpreadsheet size={14} /> ייבוא לוח מקובץ Excel
          </button>
          <button
            onClick={() => setConfirmClear(true)}
            className="flex items-center justify-center gap-1.5 rounded-lg px-3 py-1.5 text-[12px] font-semibold"
            style={{ background: "#FBEAE5", color: DANGER }}
          >
            <Trash2 size={14} /> נקה חודש זה
          </button>
        </div>
        {confirmClear && (
          <div
            className="flex flex-wrap items-center gap-2 rounded-lg px-3 py-2"
            style={{ background: "#FBEAE5" }}
          >
            <span className="text-[12px] font-semibold" style={{ color: DANGER }}>
              למחוק את כל נתוני {HEB_MONTHS[month]} {year}? לא ניתן לבטל.
            </span>
            <button
              onClick={clearMonth}
              className="rounded-md px-3 py-1 text-[11.5px] font-bold text-white"
              style={{ background: DANGER }}
            >
              כן, מחק
            </button>
            <button
              onClick={() => setConfirmClear(false)}
              className="rounded-md px-3 py-1 text-[11.5px] font-semibold"
              style={{ background: "#fff", color: INK, border: `1px solid ${LINE}` }}
            >
              ביטול
            </button>
          </div>
        )}
        <div className="grid grid-cols-2 gap-2">
          <div
            className="flex items-center gap-2 rounded-lg px-2.5 py-1.5"
            style={{ background: SIDE_A_BG }}
          >
            <span className="text-[11px] font-bold shrink-0" style={{ color: SIDE_A }}>
              צד א׳
            </span>
            <input
              value={leaders.a}
              onChange={(e) => setLeaders((p) => ({ ...p, a: e.target.value }))}
              className="w-full bg-transparent text-[12px] font-semibold outline-none"
              style={{ color: SIDE_A }}
              placeholder="שם רופא/ה בכיר/ה"
            />
          </div>
          <div
            className="flex items-center gap-2 rounded-lg px-2.5 py-1.5"
            style={{ background: SIDE_B_BG }}
          >
            <span className="text-[11px] font-bold shrink-0" style={{ color: SIDE_B }}>
              צד ב׳
            </span>
            <input
              value={leaders.b}
              onChange={(e) => setLeaders((p) => ({ ...p, b: e.target.value }))}
              className="w-full bg-transparent text-[12px] font-semibold outline-none"
              style={{ color: SIDE_B }}
              placeholder="שם רופא/ה בכיר/ה"
            />
          </div>
        </div>
      </div>

      {/* Table */}
      <div
        className="overflow-x-auto rounded-xl bg-white shadow-sm"
        style={{ border: `1px solid ${LINE}` }}
      >
        <table className="border-collapse text-right" style={{ minWidth: "980px" }}>
          <thead>
            <tr>
              <th
                className="sticky right-0 z-10 whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                style={{ background: "#fff", color: INK, borderBottom: `2px solid ${INK}` }}
              >
                תאריך
              </th>
              {["מתמחה אחראי", "מתמחים", "רוטציה", "סטאז׳רים"].map((h) => (
                <th
                  key={"a" + h}
                  className="whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                  style={{ background: SIDE_A_BG, color: SIDE_A, borderBottom: `2px solid ${SIDE_A}` }}
                >
                  {h}
                </th>
              ))}
              {["מתמחה אחראי", "מתמחים", "רוטציה", "סטאז׳רים"].map((h) => (
                <th
                  key={"b" + h}
                  className="whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                  style={{ background: SIDE_B_BG, color: SIDE_B, borderBottom: `2px solid ${SIDE_B}` }}
                >
                  {h}
                </th>
              ))}
              {["מחלקה", "מיק״ו", "כונן מחלקה", "עוזר רופא"].map((h) => (
                <th
                  key={"s" + h}
                  className="whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                  style={{ background: "#F0F0EE", color: INK, borderBottom: `2px solid ${INK}` }}
                >
                  {h}
                </th>
              ))}
              {["תורן מיון 1", "תורן מיון 2", "תורן מיון 3"].map((h) => (
                <th
                  key={"er" + h}
                  className="whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                  style={{ background: TODAY_BG, color: TODAY, borderBottom: `2px solid ${TODAY}` }}
                >
                  {h}
                </th>
              ))}
              {["הצגה קצרה", "הצגה ארוכה"].map((h) => (
                <th
                  key={"pr" + h}
                  className="whitespace-nowrap px-2 py-2 text-[11px] font-bold"
                  style={{ background: PRESENT_BG, color: PRESENT, borderBottom: `2px solid ${PRESENT}` }}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {days.map((day) => {
              const dateObj = new Date(year, month, day);
              const iso = `${year}-${pad(month + 1)}-${pad(day)}`;
              const isToday = iso === today;
              const dow = dateObj.getDay();
              const isWeekend = dow === 5 || dow === 6;
              const entry = monthData[day] || emptyDayEntry();
              const rowBg = isToday ? TODAY_BG : isWeekend ? WEEKEND_BG : "#fff";

              return (
                <tr key={day} style={{ background: rowBg }}>
                  <td
                    className="sticky right-0 z-10 whitespace-nowrap px-2 py-1 text-[12px] font-bold"
                    style={{
                      background: rowBg,
                      color: isToday ? TODAY : INK,
                      borderBottom: `1px solid ${LINE}`,
                      boxShadow: "2px 0 0 0 rgba(0,0,0,0.02)",
                    }}
                  >
                    {day} {HEB_DAYS_SHORT[dow]}
                    {isToday && (
                      <span
                        className="mr-1 inline-block rounded-full px-1.5 py-0.5 text-[9px] font-bold"
                        style={{ background: TODAY, color: "#fff" }}
                      >
                        היום
                      </span>
                    )}
                  </td>
                  {["lead", "residents", "rotation", "interns"].map((f) => (
                    <td key={"a" + f} className="px-1 py-0.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                      <EditableField
                        value={entry.sideA[f]}
                        accent={SIDE_A}
                        bold={f === "lead" || f === "residents"}
                        placeholder={f === "interns" ? "שם פרטי" : undefined}
                        onChange={(v) => updateDay(day, "sideA", f, v)}
                      />
                    </td>
                  ))}
                  {["lead", "residents", "rotation", "interns"].map((f) => (
                    <td key={"b" + f} className="px-1 py-0.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                      <EditableField
                        value={entry.sideB[f]}
                        accent={SIDE_B}
                        bold={f === "lead" || f === "residents"}
                        placeholder={f === "interns" ? "שם פרטי" : undefined}
                        onChange={(v) => updateDay(day, "sideB", f, v)}
                      />
                    </td>
                  ))}
                  {["ward", "miku", "onCall", "paAssistant"].map((f) => (
                    <td key={"s" + f} className="px-1 py-0.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                      <EditableField
                        value={entry.shared[f]}
                        accent={INK}
                        onChange={(v) => updateDay(day, "shared", f, v)}
                      />
                    </td>
                  ))}
                  {["erOnCall1", "erOnCall2", "erOnCall3"].map((f) => (
                    <td key={"er" + f} className="px-1 py-0.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                      <EditableField
                        value={entry.shared[f]}
                        accent={TODAY}
                        onChange={(v) => updateDay(day, "shared", f, v)}
                      />
                    </td>
                  ))}
                  {["shortPresentation", "longPresentation"].map((f) =>
                    dow === 4 ? (
                      <td key={"pr" + f} className="px-1 py-0.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                        <select
                          value={entry.shared[f]}
                          onChange={(e) => updateDay(day, "shared", f, e.target.value)}
                          className="w-full min-w-[86px] rounded px-1 py-1 text-[12px] outline-none"
                          style={{
                            border: `1px solid ${entry.shared[f] ? PRESENT : LINE}`,
                            color: entry.shared[f] ? INK : "#B5B8B2",
                            background: "#fff",
                          }}
                        >
                          <option value="">—</option>
                          {knownResidents.map((n) => (
                            <option key={n} value={n}>
                              {n}
                            </option>
                          ))}
                        </select>
                      </td>
                    ) : (
                      <td
                        key={"pr" + f}
                        className="px-1 py-0.5 text-center text-[12px]"
                        style={{ borderBottom: `1px solid ${LINE}`, color: "#D5D7D1", background: "#FAFAF9" }}
                        title="רלוונטי רק ביום חמישי"
                      >
                        —
                      </td>
                    )
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {knownResidents.length === 0 && (
        <div
          className="flex items-center gap-2 rounded-lg px-3 py-2 text-[11.5px]"
          style={{ background: TODAY_BG, color: TODAY }}
        >
          <AlertCircle size={14} />
          <span>
            רשימת הבחירה ל"הצגה קצרה/ארוכה" (ביום חמישי) מוצגת לפי שמות המתמחים שהוזנו
            בטבלה (מתמחה אחראי / מתמחים). הזינו לפחות מתמחה אחד כדי שהרשימה תתמלא.
          </span>
        </div>
      )}

      {/* Notes */}
      <div className="rounded-xl bg-white p-3 shadow-sm" style={{ border: `1px solid ${LINE}` }}>
        <div className="mb-2 flex items-center gap-1.5">
          <StickyNote size={15} style={{ color: INK }} />
          <h3 className="text-[13px] font-bold" style={{ color: INK }}>
            הערות כלליות
          </h3>
        </div>
        <div className="flex flex-col gap-1.5">
          {notes.length === 0 && (
            <p className="text-[12px]" style={{ color: "#9A9D96" }}>
              אין הערות עדיין
            </p>
          )}
          {notes.map((n) => (
            <div
              key={n.id}
              className="flex items-center justify-between gap-2 rounded-lg px-2.5 py-1.5"
              style={{ background: CANVAS }}
            >
              <span className="text-[12.5px]" style={{ color: INK }}>
                {n.text}
              </span>
              <button
                onClick={() => setNotes((p) => p.filter((x) => x.id !== n.id))}
                style={{ color: DANGER }}
                aria-label="מחק הערה"
              >
                <X size={14} />
              </button>
            </div>
          ))}
        </div>
        <form
          className="mt-2 flex gap-2"
          onSubmit={(e) => {
            e.preventDefault();
            if (!newNote.trim()) return;
            setNotes((p) => [...p, { id: uid(), text: newNote.trim() }]);
            setNewNote("");
          }}
        >
          <input
            value={newNote}
            onChange={(e) => setNewNote(e.target.value)}
            placeholder="הוסף הערה..."
            className="flex-1 rounded-lg px-2.5 py-1.5 text-[12.5px] outline-none"
            style={{ border: `1px solid ${LINE}`, color: INK }}
          />
          <button
            type="submit"
            className="flex items-center gap-1 rounded-lg px-3 py-1.5 text-[12.5px] font-semibold text-white"
            style={{ background: INK }}
          >
            <Plus size={14} /> הוסף
          </button>
        </form>
      </div>

      {showImport && (
        <ImportModal
          onClose={() => setShowImport(false)}
          year={year}
          month={month}
          setYear={setYear}
          setMonth={setMonth}
          setSchedule={setSchedule}
          mk={mk}
          leaders={leaders}
          setLeaders={setLeaders}
        />
      )}
    </div>
  );
}

/* ---------------- Import from Excel ---------------- */
const FIELD_DEFS = [
  { key: "date", label: "תאריך (יום בחודש)", required: true },
  { key: "sideA.lead", label: "צד א׳ · מתמחה אחראי" },
  { key: "sideA.residents", label: "צד א׳ · מתמחים" },
  { key: "sideA.rotation", label: "צד א׳ · רוטציה" },
  { key: "sideA.interns", label: "צד א׳ · סטאז׳רים" },
  { key: "sideB.lead", label: "צד ב׳ · מתמחה אחראי" },
  { key: "sideB.residents", label: "צד ב׳ · מתמחים" },
  { key: "sideB.rotation", label: "צד ב׳ · רוטציה" },
  { key: "sideB.interns", label: "צד ב׳ · סטאז׳רים" },
  { key: "shared.ward", label: "מחלקה" },
  { key: "shared.miku", label: "מיק״ו" },
  { key: "shared.onCall", label: "כונן מחלקה" },
  { key: "shared.paAssistant", label: "עוזר רופא" },
  { key: "shared.erOnCall1", label: "תורן מיון 1" },
  { key: "shared.erOnCall2", label: "תורן מיון 2" },
  { key: "shared.erOnCall3", label: "תורן מיון 3" },
  { key: "shared.shortPresentation", label: "הצגה קצרה" },
  { key: "shared.longPresentation", label: "הצגה ארוכה" },
];

// Picks the sheet whose name best matches the target month/year (e.g. a
// tab literally called "אוגוסט 2026" or "אוגוסט 26"), falling back to the
// first sheet if nothing matches.
function findMatchingSheet(sheetNames, year, month) {
  const monthName = HEB_MONTHS[month];
  const yearFull = String(year);
  const yearShort = yearFull.slice(-2);
  let best = sheetNames[0];
  let bestScore = 0;
  sheetNames.forEach((name) => {
    let score = 0;
    if (name.includes(monthName)) score += 2;
    if (name.includes(yearFull)) score += 2;
    else if (new RegExp(`(^|\\D)${yearShort}(\\D|$)`).test(name)) score += 1;
    if (score > bestScore) {
      bestScore = score;
      best = name;
    }
  });
  return best;
}

function ImportModal({ onClose, year, month, setYear, setMonth, setSchedule, mk, leaders, setLeaders }) {
  const [mode, setMode] = useState("upload"); // "upload" | "link"
  const [link, setLink] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [sheetNames, setSheetNames] = useState([]);
  const [activeSheet, setActiveSheet] = useState("");
  const [workbook, setWorkbook] = useState(null);
  const [rows, setRows] = useState([]); // array of arrays, raw preview (+ merged virtual cols)
  const [origColCount, setOrigColCount] = useState(0); // real spreadsheet columns, excludes merged virtual ones
  const [virtualLabels, setVirtualLabels] = useState({});
  const [headerRowIndex, setHeaderRowIndex] = useState(-1);
  const [autoDetected, setAutoDetected] = useState(false);
  const [detectedLeaders, setDetectedLeaders] = useState(null);
  const [startRow, setStartRow] = useState(1); // 1-indexed
  const [mapping, setMapping] = useState({});
  const fileRef = useRef(null);

  // Default the import target to the real current month whenever the
  // import dialog opens, regardless of which month the schedule tab
  // happens to be showing.
  useEffect(() => {
    const now = new Date();
    setYear(now.getFullYear());
    setMonth(now.getMonth());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const colCount = useMemo(
    () => rows.reduce((max, r) => Math.max(max, r.length), 0),
    [rows]
  );
  const colOptions = useMemo(() => Array.from({ length: colCount }, (_, i) => i), [colCount]);
  const previewColCount = origColCount || colCount;

  const colLabel = (idx) => virtualLabels[idx] || `עמודה ${colLetter(idx)}`;

  const loadWorkbook = (wb) => {
    setWorkbook(wb);
    setSheetNames(wb.SheetNames);
    const matched = findMatchingSheet(wb.SheetNames, year, month);
    setActiveSheet(matched);
    readSheet(wb, matched);
  };

  const readSheet = (wb, sheetName) => {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
    const detected = autoDetectImport(data);
    if (detected) {
      setRows(detected.augmentedRows);
      setOrigColCount(detected.origColCount);
      setVirtualLabels(detected.virtualLabels);
      setHeaderRowIndex(detected.headerRowIndex);
      setMapping(detected.mapping);
      setStartRow(detected.headerRowIndex + 2);
      setAutoDetected(true);
      if (detected.leaders && (detected.leaders.a || detected.leaders.b)) {
        setDetectedLeaders(detected.leaders);
        setLeaders((prev) => ({
          a: detected.leaders.a || prev.a,
          b: detected.leaders.b || prev.b,
        }));
      } else {
        setDetectedLeaders(null);
      }
    } else {
      setRows(data);
      setOrigColCount(0);
      setVirtualLabels({});
      setHeaderRowIndex(-1);
      setMapping({});
      setStartRow(1);
      setAutoDetected(false);
      setDetectedLeaders(null);
    }
  };

  const onFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    setLoading(true);
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const wb = XLSX.read(data, { type: "array" });
        loadWorkbook(wb);
      } catch (err) {
        setError("לא הצלחתי לקרוא את הקובץ. ודאו שמדובר בקובץ Excel תקין (.xlsx).");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const fetchFromLink = async () => {
    if (!link.trim()) return;
    setError("");
    setLoading(true);
    const sheetId = parseGoogleSheetId(link);
    const targetUrl = sheetId ? googleXlsxExportUrl(sheetId) : link.trim();
    try {
      const res = await fetch(targetUrl);
      if (!res.ok) throw new Error("bad response");
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      loadWorkbook(wb);
    } catch (err) {
      setError(
        sheetId
          ? "לא הצלחתי לטעון את הגיליון ישירות (גוגל חוסם לרוב גישה חיצונית לקבצים פרטיים). " +
            "לחצו על \"הורדה כקובץ Excel\" למטה, ואז העלו את הקובץ שהורד דרך \"העלאת קובץ\"."
          : "לא הצלחתי לטעון את הקובץ מהקישור (ייתכן שהאתר חוסם גישה חיצונית לקובץ). " +
            "מומלץ להוריד את קובץ ה-Excel למכשיר ולהעלות אותו ישירות באמצעות \"העלאת קובץ\"."
      );
    } finally {
      setLoading(false);
    }
  };

  const setMap = (fieldKey, colIdx) => setMapping((m) => ({ ...m, [fieldKey]: colIdx }));

  const parseDay = (raw) => {
    if (raw === undefined || raw === null || raw === "") return null;
    const str = String(raw).trim();
    // plain integer day-of-month, e.g. "14"
    if (/^\d{1,2}$/.test(str)) {
      const n = parseInt(str, 10);
      return n >= 1 && n <= 31 ? n : null;
    }
    // dd/mm or dd/mm/yyyy or dd.mm.yyyy
    const m = str.match(/^(\d{1,2})[./](\d{1,2})/);
    if (m) return parseInt(m[1], 10);
    // fallback: try Date parsing
    const d = new Date(str);
    if (!isNaN(d.getTime())) return d.getDate();
    return null;
  };

  const runImport = () => {
    if (mapping.date === undefined) {
      setError("יש לבחור עמודת תאריך לפני הייבוא.");
      return;
    }
    const nDays = daysInMonth(year, month);
    const dataRows = rows.slice(startRow - 1);
    const updates = {};

    dataRows.forEach((r) => {
      const day = parseDay(r[mapping.date]);
      if (!day || day < 1 || day > nDays) return;
      const entry = emptyDayEntry();
      FIELD_DEFS.forEach(({ key }) => {
        if (key === "date") return;
        const idx = mapping[key];
        if (idx === undefined || idx === "") return;
        const [section, field] = key.split(".");
        const val = r[idx];
        entry[section][field] = val !== undefined && val !== null ? String(val).trim() : "";
      });
      // "תורן מיון" columns rarely exist as their own column in the source
      // file - auto-assign them from that day's own resident roster
      // (same pool the dropdown itself offers), just like every other
      // column gets filled automatically from the sheet.
      const pool = dayResidentPool(entry);
      if (!entry.shared.erOnCall1 && pool[0]) entry.shared.erOnCall1 = pool[0];
      if (!entry.shared.erOnCall2 && pool[1]) entry.shared.erOnCall2 = pool[1];
      if (!entry.shared.erOnCall3 && pool[2]) entry.shared.erOnCall3 = pool[2];
      // הצגה קצרה/ארוכה relevant only on Thursdays
      const dow = new Date(year, month, day).getDay();
      if (dow !== 4) {
        entry.shared.shortPresentation = "";
        entry.shared.longPresentation = "";
      }
      updates[day] = entry;
    });

    if (Object.keys(updates).length === 0) {
      setError("לא נמצאו שורות תקינות לייבוא. בדקו את עמודת התאריך ואת שורת ההתחלה.");
      return;
    }

    setSchedule((prev) => ({
      ...prev,
      [mk]: { ...(prev[mk] || {}), ...updates },
    }));
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/40 sm:items-center" onClick={onClose}>
      <div
        className="max-h-[92vh] w-full max-w-2xl overflow-y-auto rounded-t-2xl bg-white p-4 sm:rounded-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex items-center justify-between">
          <h3 className="text-[15px] font-bold" style={{ color: INK }}>
            ייבוא לוח שיבוצים מקובץ Excel
          </h3>
          <button onClick={onClose} style={{ color: "#8A8D86" }}>
            <X size={18} />
          </button>
        </div>

        {!workbook && (
          <>
            <div className="mb-3 flex overflow-hidden self-start rounded-md w-fit" style={{ border: `1px solid ${LINE}` }}>
              <button
                onClick={() => setMode("upload")}
                className="flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold"
                style={{ background: mode === "upload" ? INK : "#fff", color: mode === "upload" ? "#fff" : INK }}
              >
                <Upload size={13} /> העלאת קובץ
              </button>
              <button
                onClick={() => setMode("link")}
                className="flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold"
                style={{ background: mode === "link" ? INK : "#fff", color: mode === "link" ? "#fff" : INK }}
              >
                <LinkIcon size={13} /> קישור לקובץ
              </button>
            </div>

            {mode === "upload" ? (
              <div
                className="flex flex-col items-center justify-center gap-2 rounded-xl p-6"
                style={{ border: `1.5px dashed ${LINE}`, background: CANVAS }}
              >
                <FileSpreadsheet size={26} style={{ color: "#B5B8B2" }} />
                <button
                  onClick={() => fileRef.current?.click()}
                  className="rounded-lg px-3 py-1.5 text-[12px] font-semibold text-white"
                  style={{ background: INK }}
                >
                  בחר קובץ Excel (.xlsx)
                </button>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onFile} />
              </div>
            ) : (
              <div className="flex flex-col gap-2">
                <p className="text-[11.5px] leading-relaxed" style={{ color: "#8A8D86" }}>
                  אפשר להדביק כאן קישור לגיליון Google Sheets (הקישור הרגיל שמשתפים, עם
                  ‎/edit) או קישור ישיר להורדת קובץ Excel. עבור קישורי Google Sheets פרטיים,
                  לרוב יהיה צורך להוריד ולהעלות את הקובץ (ראו כפתור ההורדה שיופיע למטה).
                </p>
                <div className="flex gap-1.5">
                  <input
                    value={link}
                    onChange={(e) => setLink(e.target.value)}
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                    className="flex-1 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
                    style={{ border: `1px solid ${LINE}`, color: INK }}
                  />
                  <button
                    onClick={fetchFromLink}
                    disabled={loading}
                    className="rounded-md px-3 py-1.5 text-[12px] font-semibold text-white"
                    style={{ background: INK, opacity: loading ? 0.6 : 1 }}
                  >
                    {loading ? "טוען..." : "טען"}
                  </button>
                </div>

                {parseGoogleSheetId(link) && (
                  <div
                    className="flex flex-col gap-1.5 rounded-lg px-3 py-2"
                    style={{ background: TODAY_BG }}
                  >
                    <span className="text-[11.5px] font-semibold" style={{ color: TODAY }}>
                      זוהה קישור לגיליון Google Sheets
                    </span>
                    <span className="text-[11px] leading-relaxed" style={{ color: "#8A6A2B" }}>
                      אם הטעינה האוטומטית לא הצליחה - לחצו להורדה, ואז העלו את הקובץ שהתקבל
                      דרך "העלאת קובץ" למעלה.
                    </span>
                    <a
                      href={googleXlsxExportUrl(parseGoogleSheetId(link))}
                      target="_blank"
                      rel="noreferrer"
                      className="flex w-fit items-center gap-1.5 rounded-md px-3 py-1.5 text-[11.5px] font-bold text-white"
                      style={{ background: TODAY }}
                    >
                      <FileSpreadsheet size={13} /> הורדה כקובץ Excel
                    </a>
                  </div>
                )}
              </div>
            )}
          </>
        )}

        {error && (
          <div
            className="mt-3 flex items-start gap-2 rounded-lg px-3 py-2 text-[11.5px]"
            style={{ background: "#FBEAE5", color: DANGER }}
          >
            <AlertCircle size={14} className="mt-0.5 shrink-0" />
            <span>{error}</span>
          </div>
        )}

        {workbook && (
          <div className="mt-1 flex flex-col gap-3">
            {autoDetected ? (
              <div
                className="flex items-start gap-2 rounded-lg px-3 py-2 text-[11.5px]"
                style={{ background: SIDE_A_BG, color: SIDE_A }}
              >
                <Check size={14} className="mt-0.5 shrink-0" />
                <span>
                  זוהו כותרות העמודות אוטומטית לפי הקובץ, ושיוך העמודות מולא בהתאם. אפשר
                  לבדוק ולתקן ידנית ברשימת "שיוך עמודות" שלמטה במידת הצורך.
                  {" "}
                  עמודות "תורן מיון" ימולאו אוטומטית מתוך המתמחים שמופיעים באותו יום, אלא אם
                  קיימת עמודה מתאימה בקובץ עצמו.
                  {detectedLeaders && (detectedLeaders.a || detectedLeaders.b) && (
                    <>
                      {" "}
                      זוהו גם שמות הרופאים הבכירים
                      {detectedLeaders.a && <> - צד א׳: <b>{detectedLeaders.a}</b></>}
                      {detectedLeaders.b && <> · צד ב׳: <b>{detectedLeaders.b}</b></>}
                      , ומולאו בראש הלוח.
                    </>
                  )}
                </span>
              </div>
            ) : (
              <div
                className="flex items-start gap-2 rounded-lg px-3 py-2 text-[11.5px]"
                style={{ background: "#FBEAE5", color: DANGER }}
              >
                <AlertCircle size={14} className="mt-0.5 shrink-0" />
                <span>לא זוהו כותרות עמודות מוכרות באופן אוטומטי. יש לשייך את העמודות ידנית למטה.</span>
              </div>
            )}

            {sheetNames.length > 1 && (
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-[11.5px] font-semibold" style={{ color: INK }}>
                  גיליון:
                </span>
                <select
                  value={activeSheet}
                  onChange={(e) => {
                    setActiveSheet(e.target.value);
                    readSheet(workbook, e.target.value);
                  }}
                  className="rounded-md px-2 py-1 text-[12px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                >
                  {sheetNames.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <span className="text-[10.5px]" style={{ color: "#8A8D86" }}>
                  נבחר אוטומטית לפי חודש היעד ({HEB_MONTHS[month]} {year}) - ניתן להחליף גיליון ידנית.
                </span>
              </div>
            )}

            <div className="flex flex-wrap items-center gap-3">
              <div className="flex items-center gap-1.5">
                <span className="text-[11.5px] font-semibold" style={{ color: INK }}>
                  חודש יעד:
                </span>
                <select
                  value={month}
                  onChange={(e) => {
                    const newMonth = parseInt(e.target.value, 10);
                    setMonth(newMonth);
                    if (workbook && sheetNames.length > 1) {
                      const matched = findMatchingSheet(sheetNames, year, newMonth);
                      setActiveSheet(matched);
                      readSheet(workbook, matched);
                    }
                  }}
                  className="rounded-md px-2 py-1 text-[12px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                >
                  {HEB_MONTHS.map((m, i) => (
                    <option key={m} value={i}>
                      {m}
                    </option>
                  ))}
                </select>
                <input
                  type="number"
                  value={year}
                  onChange={(e) => {
                    const newYear = parseInt(e.target.value, 10) || year;
                    setYear(newYear);
                    if (workbook && sheetNames.length > 1) {
                      const matched = findMatchingSheet(sheetNames, newYear, month);
                      setActiveSheet(matched);
                      readSheet(workbook, matched);
                    }
                  }}
                  className="w-20 rounded-md px-2 py-1 text-[12px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                />
              </div>
              <div className="flex items-center gap-1.5">
                <span className="text-[11.5px] font-semibold" style={{ color: INK }}>
                  שורת התחלת נתונים:
                </span>
                <input
                  type="number"
                  min={1}
                  value={startRow}
                  onChange={(e) => setStartRow(parseInt(e.target.value, 10) || 1)}
                  className="w-16 rounded-md px-2 py-1 text-[12px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                />
              </div>
            </div>

            {/* preview */}
            <div className="overflow-x-auto rounded-lg" style={{ border: `1px solid ${LINE}` }}>
              <table className="border-collapse text-right text-[11px]" style={{ minWidth: "500px" }}>
                <thead>
                  <tr>
                    {Array.from({ length: previewColCount }, (_, c) => c).map((c) => (
                      <th key={c} className="whitespace-nowrap px-2 py-1 font-bold" style={{ background: CANVAS, color: INK }}>
                        {colLetter(c)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 8).map((r, ri) => (
                    <tr
                      key={ri}
                      style={{
                        background: ri === startRow - 1 ? TODAY_BG : ri === headerRowIndex ? SIDE_A_BG : "#fff",
                      }}
                    >
                      {Array.from({ length: previewColCount }, (_, c) => c).map((c) => (
                        <td key={c} className="whitespace-nowrap px-2 py-1" style={{ borderTop: `1px solid ${LINE}`, color: INK }}>
                          {r[c] ?? ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-[10.5px]" style={{ color: "#8A8D86" }}>
              {autoDetected && "שורת הכותרות שזוהתה מסומנת בירוק. "}
              השורה הכתומה היא שורת ההתחלה שנבחרה. מציגים 8 שורות ראשונות לתצוגה מקדימה בלבד.
            </p>

            {/* mapping */}
            <div className="flex flex-col gap-1.5">
              <div className="text-[12px] font-bold" style={{ color: INK }}>
                שיוך עמודות {autoDetected && "(מולא אוטומטית, ניתן לערוך)"}
              </div>
              <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-2">
                {FIELD_DEFS.map(({ key, label, required }) => (
                  <div key={key} className="flex items-center justify-between gap-2 rounded-md px-2 py-1.5" style={{ background: CANVAS }}>
                    <span className="text-[11.5px]" style={{ color: INK }}>
                      {label}
                      {required && <span style={{ color: DANGER }}> *</span>}
                    </span>
                    <select
                      value={mapping[key] === undefined ? "" : mapping[key]}
                      onChange={(e) => setMap(key, e.target.value === "" ? undefined : parseInt(e.target.value, 10))}
                      className="rounded-md px-1.5 py-1 text-[11.5px] outline-none"
                      style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
                    >
                      <option value="">לא בשימוש</option>
                      {colOptions.map((c) => (
                        <option key={c} value={c}>
                          {colLabel(c)}
                        </option>
                      ))}
                    </select>
                  </div>
                ))}
              </div>
            </div>

            <button
              onClick={runImport}
              className="rounded-lg py-2 text-[13px] font-bold text-white"
              style={{ background: INK }}
            >
              ייבא לתוך {HEB_MONTHS[month]} {year}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------- Patients Tab ---------------- */
function PatientColumn({ side, label, accent, accentBg, patients, setPatients, otherSideKey }) {
  const [newName, setNewName] = useState("");
  const [newId, setNewId] = useState("");
  const [newRoom, setNewRoom] = useState("");
  const [taskDrafts, setTaskDrafts] = useState({});

  const addPatient = (e) => {
    e.preventDefault();
    if (!newName.trim()) return;
    setPatients((p) => [
      ...p,
      { id: uid(), name: newName.trim(), idNumber: newId.trim(), room: newRoom.trim(), tasks: [] },
    ]);
    setNewName("");
    setNewId("");
    setNewRoom("");
  };

  const removePatient = (pid) => setPatients((p) => p.filter((x) => x.id !== pid));

  const addTask = (pid) => {
    const text = (taskDrafts[pid] || "").trim();
    if (!text) return;
    setPatients((p) =>
      p.map((pt) => (pt.id === pid ? { ...pt, tasks: [...pt.tasks, { id: uid(), text, done: false }] } : pt))
    );
    setTaskDrafts((d) => ({ ...d, [pid]: "" }));
  };

  const toggleTask = (pid, tid) =>
    setPatients((p) =>
      p.map((pt) =>
        pt.id === pid
          ? { ...pt, tasks: pt.tasks.map((t) => (t.id === tid ? { ...t, done: !t.done } : t)) }
          : pt
      )
    );

  const removeTask = (pid, tid) =>
    setPatients((p) =>
      p.map((pt) => (pt.id === pid ? { ...pt, tasks: pt.tasks.filter((t) => t.id !== tid) } : pt))
    );

  const moveOut = (patient) => {
    setPatients((p) => p.filter((x) => x.id !== patient.id));
    otherSideKey(patient);
  };

  return (
    <div className="flex-1 min-w-[280px]">
      <div
        className="mb-2 flex items-center justify-between rounded-lg px-3 py-2"
        style={{ background: accentBg }}
      >
        <h3 className="text-[14px] font-bold" style={{ color: accent }}>
          {label}
        </h3>
        <span className="text-[11px] font-semibold" style={{ color: accent }}>
          {patients.length} חולים
        </span>
      </div>

      <div className="flex flex-col gap-2">
        {patients.map((pt) => (
          <div
            key={pt.id}
            className="rounded-xl bg-white p-3 shadow-sm"
            style={{ border: `1px solid ${LINE}`, borderRight: `4px solid ${accent}` }}
          >
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-[13.5px] font-bold" style={{ color: INK }}>
                  {pt.name}
                </div>
                <div className="text-[11px]" style={{ color: "#8A8D86" }}>
                  {pt.idNumber && <span>ת.ז. {pt.idNumber}</span>}
                  {pt.idNumber && pt.room && <span> · </span>}
                  {pt.room && <span>חדר {pt.room}</span>}
                </div>
              </div>
              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  onClick={() => moveOut(pt)}
                  title={`העבר ל${side === "a" ? "צד ב׳" : "צד א׳"}`}
                  className="flex h-7 w-7 items-center justify-center rounded-full transition-colors hover:bg-black/5"
                  style={{ color: "#8A8D86" }}
                >
                  <ArrowLeftRight size={14} />
                </button>
                <button
                  onClick={() => removePatient(pt.id)}
                  title="מחק חולה"
                  className="flex h-7 w-7 items-center justify-center rounded-full transition-colors hover:bg-black/5"
                  style={{ color: DANGER }}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>

            <div className="mt-2 flex flex-col gap-1">
              {pt.tasks.map((t) => (
                <div key={t.id} className="flex items-center gap-2">
                  <button
                    onClick={() => toggleTask(pt.id, t.id)}
                    className="flex h-5 w-5 shrink-0 items-center justify-center rounded"
                    style={{
                      border: `1.5px solid ${t.done ? accent : "#C7C9C2"}`,
                      background: t.done ? accent : "transparent",
                    }}
                  >
                    {t.done && <Check size={12} color="#fff" />}
                  </button>
                  <span
                    className="flex-1 text-[12.5px]"
                    style={{
                      color: t.done ? "#A9ABA4" : INK,
                      textDecoration: t.done ? "line-through" : "none",
                    }}
                  >
                    {t.text}
                  </span>
                  <button onClick={() => removeTask(pt.id, t.id)} style={{ color: "#C7C9C2" }}>
                    <X size={12} />
                  </button>
                </div>
              ))}
              <form
                className="mt-1 flex gap-1.5"
                onSubmit={(e) => {
                  e.preventDefault();
                  addTask(pt.id);
                }}
              >
                <input
                  value={taskDrafts[pt.id] || ""}
                  onChange={(e) => setTaskDrafts((d) => ({ ...d, [pt.id]: e.target.value }))}
                  placeholder="הוסף מטלה..."
                  className="flex-1 rounded-md px-2 py-1 text-[12px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                />
                <button
                  type="submit"
                  className="flex items-center justify-center rounded-md px-2"
                  style={{ background: accentBg, color: accent }}
                >
                  <Plus size={14} />
                </button>
              </form>
            </div>
          </div>
        ))}
      </div>

      <form
        onSubmit={addPatient}
        className="mt-2 flex flex-col gap-1.5 rounded-xl p-3"
        style={{ border: `1.5px dashed ${LINE}` }}
      >
        <div className="text-[11px] font-semibold" style={{ color: "#8A8D86" }}>
          הוספת חולה חדש
        </div>
        <div className="flex flex-wrap gap-1.5">
          <input
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            placeholder="שם החולה"
            className="min-w-[110px] flex-1 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
            style={{ border: `1px solid ${LINE}`, color: INK }}
          />
          <input
            value={newId}
            onChange={(e) => setNewId(e.target.value)}
            placeholder="ת.ז."
            className="w-24 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
            style={{ border: `1px solid ${LINE}`, color: INK }}
          />
          <input
            value={newRoom}
            onChange={(e) => setNewRoom(e.target.value)}
            placeholder="חדר"
            className="w-16 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
            style={{ border: `1px solid ${LINE}`, color: INK }}
          />
        </div>
        <button
          type="submit"
          className="flex items-center justify-center gap-1 rounded-md py-1.5 text-[12.5px] font-semibold text-white"
          style={{ background: accent }}
        >
          <Plus size={14} /> הוסף חולה
        </button>
      </form>
    </div>
  );
}

function IntakeModal({ onClose, addPatientToSide }) {
  const [image, setImage] = useState(null);
  const [rows, setRows] = useState([{ id: uid(), name: "", idNumber: "", room: "", side: "a" }]);
  const fileRef = useRef(null);

  const onFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setImage(reader.result);
    reader.readAsDataURL(file);
  };

  const updateRow = (id, field, val) =>
    setRows((r) => r.map((row) => (row.id === id ? { ...row, [field]: val } : row)));

  const addRow = () =>
    setRows((r) => [...r, { id: uid(), name: "", idNumber: "", room: "", side: "a" }]);

  const removeRow = (id) => setRows((r) => r.filter((row) => row.id !== id));

  const submit = () => {
    rows.forEach((row) => {
      if (row.name.trim()) {
        addPatientToSide(row.side, {
          id: uid(),
          name: row.name.trim(),
          idNumber: row.idNumber.trim(),
          room: row.room.trim(),
          tasks: [],
        });
      }
    });
    onClose();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-end justify-center bg-black/40 sm:items-center"
      onClick={onClose}
    >
      <div
        className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-t-2xl bg-white p-4 sm:rounded-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex items-center justify-between">
          <h3 className="text-[15px] font-bold" style={{ color: INK }}>
            קליטת דף נוכחות
          </h3>
          <button onClick={onClose} style={{ color: "#8A8D86" }}>
            <X size={18} />
          </button>
        </div>

        <p className="mb-3 text-[12px] leading-relaxed" style={{ color: "#8A8D86" }}>
          העלו צילום של דף הנוכחות לצפייה תוך כדי הקלדה. זיהוי טקסט אוטומטי מכתב יד אינו זמין
          באפליקציה, לכן יש להקליד את השמות ולשייך כל חולה לצד המתאים בטבלה שמתחת לתמונה.
        </p>

        <div
          className="mb-3 flex flex-col items-center justify-center gap-2 rounded-xl p-4"
          style={{ border: `1.5px dashed ${LINE}`, background: CANVAS }}
        >
          {image ? (
            <img src={image} alt="דף נוכחות" className="max-h-56 rounded-lg object-contain" />
          ) : (
            <ImageIcon size={28} style={{ color: "#B5B8B2" }} />
          )}
          <button
            onClick={() => fileRef.current?.click()}
            className="flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-[12px] font-semibold text-white"
            style={{ background: INK }}
          >
            <Upload size={13} /> {image ? "החלף תמונה" : "העלה תמונה"}
          </button>
          <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={onFile} />
        </div>

        <div className="flex flex-col gap-2">
          {rows.map((row) => (
            <div key={row.id} className="flex flex-wrap items-center gap-1.5">
              <input
                value={row.name}
                onChange={(e) => updateRow(row.id, "name", e.target.value)}
                placeholder="שם החולה"
                className="min-w-[100px] flex-1 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
                style={{ border: `1px solid ${LINE}`, color: INK }}
              />
              <input
                value={row.idNumber}
                onChange={(e) => updateRow(row.id, "idNumber", e.target.value)}
                placeholder="ת.ז."
                className="w-20 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
                style={{ border: `1px solid ${LINE}`, color: INK }}
              />
              <input
                value={row.room}
                onChange={(e) => updateRow(row.id, "room", e.target.value)}
                placeholder="חדר"
                className="w-14 rounded-md px-2 py-1.5 text-[12.5px] outline-none"
                style={{ border: `1px solid ${LINE}`, color: INK }}
              />
              <div className="flex overflow-hidden rounded-md" style={{ border: `1px solid ${LINE}` }}>
                <button
                  onClick={() => updateRow(row.id, "side", "a")}
                  className="px-2 py-1.5 text-[11px] font-bold"
                  style={{
                    background: row.side === "a" ? SIDE_A : "#fff",
                    color: row.side === "a" ? "#fff" : SIDE_A,
                  }}
                >
                  א׳
                </button>
                <button
                  onClick={() => updateRow(row.id, "side", "b")}
                  className="px-2 py-1.5 text-[11px] font-bold"
                  style={{
                    background: row.side === "b" ? SIDE_B : "#fff",
                    color: row.side === "b" ? "#fff" : SIDE_B,
                  }}
                >
                  ב׳
                </button>
              </div>
              <button onClick={() => removeRow(row.id)} style={{ color: DANGER }}>
                <X size={14} />
              </button>
            </div>
          ))}
        </div>

        <button
          onClick={addRow}
          className="mt-2 flex items-center gap-1 text-[12px] font-semibold"
          style={{ color: INK }}
        >
          <Plus size={14} /> הוסף שורה
        </button>

        <button
          onClick={submit}
          className="mt-4 w-full rounded-lg py-2 text-[13px] font-bold text-white"
          style={{ background: INK }}
        >
          הוסף לרשימות החולים
        </button>
      </div>
    </div>
  );
}

function PatientsTab({ patientsA, setPatientsA, patientsB, setPatientsB }) {
  const [showIntake, setShowIntake] = useState(false);

  const addPatientToSide = (side, patient) => {
    if (side === "a") setPatientsA((p) => [...p, patient]);
    else setPatientsB((p) => [...p, patient]);
  };

  return (
    <div className="flex flex-col gap-3">
      <button
        onClick={() => setShowIntake(true)}
        className="flex items-center justify-center gap-2 self-start rounded-lg px-3.5 py-2 text-[12.5px] font-semibold text-white shadow-sm"
        style={{ background: INK }}
      >
        <Upload size={14} /> קליטת דף נוכחות מצולם
      </button>

      <div className="flex flex-col gap-3 sm:flex-row">
        <PatientColumn
          side="a"
          label="חולים · צד א׳"
          accent={SIDE_A}
          accentBg={SIDE_A_BG}
          patients={patientsA}
          setPatients={setPatientsA}
          otherSideKey={(pt) => setPatientsB((p) => [...p, pt])}
        />
        <PatientColumn
          side="b"
          label="חולים · צד ב׳"
          accent={SIDE_B}
          accentBg={SIDE_B_BG}
          patients={patientsB}
          setPatients={setPatientsB}
          otherSideKey={(pt) => setPatientsA((p) => [...p, pt])}
        />
      </div>

      {showIntake && (
        <IntakeModal onClose={() => setShowIntake(false)} addPatientToSide={addPatientToSide} />
      )}
    </div>
  );
}

/* ---------------- Appointments Tab ---------------- */
function AppointmentsTab({ appointments, setAppointments, allPatients }) {
  const [patientChoice, setPatientChoice] = useState("");
  const [manualId, setManualId] = useState("");
  const [useManual, setUseManual] = useState(false);
  const [type, setType] = useState("");
  const [notes, setNotes] = useState("");
  const [search, setSearch] = useState("");

  const filteredPatients = useMemo(() => {
    if (!search.trim()) return allPatients;
    return allPatients.filter((p) => p.name.includes(search.trim()));
  }, [allPatients, search]);

  const addAppointment = (e) => {
    e.preventDefault();
    const patientLabel = useManual
      ? manualId
        ? `ת.ז. ${manualId}`
        : ""
      : (() => {
          const p = allPatients.find((x) => x.id === patientChoice);
          return p ? p.name : "";
        })();
    if (!patientLabel) return;

    setAppointments((prev) => [
      ...prev,
      {
        id: uid(),
        scheduledDate: "",
        patientLabel,
        type: type.trim(),
        notes: notes.trim(),
        done: false,
      },
    ]);
    setPatientChoice("");
    setManualId("");
    setType("");
    setNotes("");
  };

  const toggleDone = (id) =>
    setAppointments((prev) => prev.map((a) => (a.id === id ? { ...a, done: !a.done } : a)));

  const setScheduledDate = (id, val) =>
    setAppointments((prev) => prev.map((a) => (a.id === id ? { ...a, scheduledDate: val } : a)));

  const removeAppt = (id) => setAppointments((prev) => prev.filter((a) => a.id !== id));

  const sorted = useMemo(() => {
    return [...appointments].sort((a, b) => {
      if (a.done !== b.done) return a.done ? 1 : -1;
      if (a.scheduledDate && b.scheduledDate) return new Date(a.scheduledDate) - new Date(b.scheduledDate);
      if (a.scheduledDate) return -1;
      if (b.scheduledDate) return 1;
      return 0;
    });
  }, [appointments]);

  return (
    <div className="flex flex-col gap-3">
      <form
        onSubmit={addAppointment}
        className="flex flex-col gap-2 rounded-xl bg-white p-3 shadow-sm"
        style={{ border: `1px solid ${LINE}` }}
      >
        <div className="text-[13px] font-bold" style={{ color: INK }}>
          הוספת בירור לתיאום תור
        </div>
        <p className="text-[11.5px] leading-relaxed" style={{ color: "#8A8D86" }}>
          הרשימה מיועדת למזכירה לקביעת תורים. אין צורך למלא תאריך מראש - ניתן לסמן ✓ כשהתור
          נקבע, ואופציונלי לציין לאיזה תאריך.
        </p>

        <div className="flex overflow-hidden self-start rounded-md" style={{ border: `1px solid ${LINE}` }}>
          <button
            type="button"
            onClick={() => setUseManual(false)}
            className="px-2.5 py-1.5 text-[11px] font-semibold"
            style={{ background: !useManual ? INK : "#fff", color: !useManual ? "#fff" : INK }}
          >
            מרשימת חולים
          </button>
          <button
            type="button"
            onClick={() => setUseManual(true)}
            className="px-2.5 py-1.5 text-[11px] font-semibold"
            style={{ background: useManual ? INK : "#fff", color: useManual ? "#fff" : INK }}
          >
            הזנת ת.ז.
          </button>
        </div>

        {useManual ? (
          <input
            value={manualId}
            onChange={(e) => setManualId(e.target.value)}
            placeholder="תעודת זהות של החולה"
            className="rounded-md px-2 py-1.5 text-[12.5px] outline-none"
            style={{ border: `1px solid ${LINE}`, color: INK }}
          />
        ) : (
          <div className="flex flex-col gap-1.5">
            <div className="relative">
              <Search
                size={13}
                className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2"
                style={{ color: "#B5B8B2" }}
              />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="חיפוש חולה..."
                className="w-full rounded-md py-1.5 pl-2 pr-7 text-[12.5px] outline-none"
                style={{ border: `1px solid ${LINE}`, color: INK }}
              />
            </div>
            <select
              value={patientChoice}
              onChange={(e) => setPatientChoice(e.target.value)}
              className="rounded-md px-2 py-1.5 text-[12.5px] outline-none"
              style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
            >
              <option value="">בחר חולה...</option>
              {filteredPatients.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.name} · צד {p.side === "a" ? "א׳" : "ב׳"}
                </option>
              ))}
            </select>
          </div>
        )}

        <input
          value={type}
          onChange={(e) => setType(e.target.value)}
          placeholder="סוג הבירור / התור (בדיקה, מעקב, תוצאות...)"
          className="rounded-md px-2 py-1.5 text-[12.5px] outline-none"
          style={{ border: `1px solid ${LINE}`, color: INK }}
        />
        <input
          value={notes}
          onChange={(e) => setNotes(e.target.value)}
          placeholder="הערות"
          className="rounded-md px-2 py-1.5 text-[12.5px] outline-none"
          style={{ border: `1px solid ${LINE}`, color: INK }}
        />

        <button
          type="submit"
          className="flex items-center justify-center gap-1.5 rounded-md py-2 text-[12.5px] font-bold text-white"
          style={{ background: INK }}
        >
          <Plus size={14} /> הוסף תור
        </button>
      </form>

      <div className="flex flex-col gap-2">
        {sorted.length === 0 && (
          <p className="text-center text-[12.5px]" style={{ color: "#9A9D96" }}>
            אין תורים עדיין
          </p>
        )}
        {sorted.map((a) => (
          <div
            key={a.id}
            className="flex items-start gap-2.5 rounded-xl bg-white p-3 shadow-sm"
            style={{
              border: `1px solid ${LINE}`,
              borderRight: `4px solid ${a.done ? "#5C8A6A" : TODAY}`,
              opacity: a.done ? 0.7 : 1,
            }}
          >
            <button
              onClick={() => toggleDone(a.id)}
              title="סמן שהתור נקבע"
              className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded"
              style={{
                border: `1.5px solid ${a.done ? "#5C8A6A" : TODAY}`,
                background: a.done ? "#5C8A6A" : "transparent",
              }}
            >
              {a.done && <Check size={12} color="#fff" />}
            </button>
            <div className="flex-1">
              <div className="flex flex-wrap items-baseline gap-x-2">
                <span className="text-[13px] font-bold" style={{ color: INK }}>
                  {a.patientLabel}
                </span>
                <span
                  className="rounded-full px-1.5 py-0.5 text-[10px] font-bold"
                  style={{
                    background: a.done ? "#E7F1EE" : TODAY_BG,
                    color: a.done ? "#5C8A6A" : TODAY,
                  }}
                >
                  {a.done ? "נקבע תור" : "ממתין לקביעה"}
                </span>
              </div>
              {a.type && (
                <div className="mt-0.5 text-[12px]" style={{ color: INK }}>
                  {a.type}
                </div>
              )}
              {a.notes && (
                <div className="text-[11.5px]" style={{ color: "#8A8D86" }}>
                  {a.notes}
                </div>
              )}
              <div className="mt-1.5 flex items-center gap-1.5">
                <span className="text-[11px] shrink-0" style={{ color: "#8A8D86" }}>
                  תאריך התור (אם נקבע):
                </span>
                <input
                  type="date"
                  value={a.scheduledDate || ""}
                  onChange={(e) => setScheduledDate(a.id, e.target.value)}
                  className="rounded-md px-1.5 py-1 text-[11.5px] outline-none"
                  style={{ border: `1px solid ${LINE}`, color: INK }}
                />
              </div>
            </div>
            <button onClick={() => removeAppt(a.id)} style={{ color: DANGER }}>
              <Trash2 size={14} />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------- Main App ---------------- */
export default function WardApp() {
  const [tab, setTab] = useState("schedule");

  const initial = loadState();

  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth());

  const [schedule, setSchedule] = useState(initial?.schedule || {});
  const [leaders, setLeaders] = useState(initial?.leaders || { a: "", b: "" });
  const [notes, setNotes] = useState(initial?.notes || []);
  const [patientsA, setPatientsA] = useState(initial?.patientsA || []);
  const [patientsB, setPatientsB] = useState(initial?.patientsB || []);
  const [appointments, setAppointments] = useState(initial?.appointments || []);

  useEffect(() => {
    saveState({ schedule, leaders, notes, patientsA, patientsB, appointments });
  }, [schedule, leaders, notes, patientsA, patientsB, appointments]);

  const allPatients = useMemo(
    () => [
      ...patientsA.map((p) => ({ ...p, side: "a" })),
      ...patientsB.map((p) => ({ ...p, side: "b" })),
    ],
    [patientsA, patientsB]
  );

  const tabs = [
    { key: "schedule", label: "לוח שיבוצים", icon: CalendarDays },
    { key: "patients", label: "חולים", icon: Users },
    { key: "appointments", label: "תורים", icon: ClipboardList },
  ];

  return (
    <div dir="rtl" style={{ background: CANVAS, minHeight: "100vh", fontFamily: "'Assistant', 'Heebo', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700;800&family=Heebo:wght@500;700;800&display=swap');
        * { font-family: 'Assistant', 'Heebo', sans-serif; }
        h1, h2, .heebo { font-family: 'Heebo', 'Assistant', sans-serif; }
        input, select { font-family: 'Assistant', sans-serif; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: #D5D7D1; border-radius: 4px; }
      `}</style>

      {/* Header */}
      <header
        className="sticky top-0 z-20 px-4 pb-3 pt-4 shadow-sm"
        style={{ background: INK }}
      >
        <div className="mx-auto flex max-w-3xl items-center justify-between">
          <div>
            <h1 className="heebo text-[16px] font-extrabold text-white">פנימית ד׳</h1>
            <p className="text-[11px]" style={{ color: "#AEB6B4" }}>
              לוח תורנויות · חולים · תורים
            </p>
          </div>
          <div
            className="flex h-9 w-9 items-center justify-center rounded-full text-[13px] font-extrabold"
            style={{ background: TODAY, color: INK }}
          >
            ד׳
          </div>
        </div>

        <nav className="mx-auto mt-3 flex max-w-3xl gap-1.5">
          {tabs.map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              onClick={() => setTab(key)}
              className="flex flex-1 items-center justify-center gap-1.5 rounded-lg py-2 text-[12.5px] font-bold transition-colors"
              style={{
                background: tab === key ? "#fff" : "rgba(255,255,255,0.08)",
                color: tab === key ? INK : "#D7DBD9",
              }}
            >
              <Icon size={14} />
              {label}
            </button>
          ))}
        </nav>
      </header>

      {/* Content */}
      <main className="mx-auto max-w-3xl px-3 py-4 sm:px-4">
        {tab === "schedule" && (
          <ScheduleTab
            schedule={schedule}
            setSchedule={setSchedule}
            leaders={leaders}
            setLeaders={setLeaders}
            notes={notes}
            setNotes={setNotes}
            year={year}
            month={month}
            setYear={setYear}
            setMonth={setMonth}
          />
        )}
        {tab === "patients" && (
          <PatientsTab
            patientsA={patientsA}
            setPatientsA={setPatientsA}
            patientsB={patientsB}
            setPatientsB={setPatientsB}
          />
        )}
        {tab === "appointments" && (
          <AppointmentsTab
            appointments={appointments}
            setAppointments={setAppointments}
            allPatients={allPatients}
          />
        )}
      </main>
    </div>
  );
}
