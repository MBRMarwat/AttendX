import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import "./App.css";

// ─── utils ────────────────────────────────────────────────────────────────────
const todayStr = () => new Date().toISOString().split("T")[0];
const fmtDate  = (iso) => { const [y,m,d]=iso.split("-"); return `${d}/${m}/${String(y).slice(2)}`; };
const isP = (v) => v?.trim().toLowerCase() === "p";
const isA = (v) => v?.trim().toLowerCase() === "a";
const norm = (v) => { if(isP(v)) return "P"; if(isA(v)) return "A"; return v??""; };
const uid  = () => Math.random().toString(36).slice(2,9);
const clamp = (v,mn,mx) => Math.min(mx, Math.max(mn, v));
let XLSX = null;
const loadXlsx = async () => {
  if (!XLSX) XLSX = await import("xlsx");
  return XLSX;
};

const _canvas = typeof document !== "undefined" ? document.createElement("canvas") : null;
const measureText = (text, font="500 13px IBM Plex Sans, sans-serif") => {
  if (!_canvas) return text.length * 8;
  const ctx = _canvas.getContext("2d");
  ctx.font = font;
  return ctx.measureText(text).width;
};
const autoFitWidth = (values, minW, maxW, pad=24) => {
  if (!values.length) return minW;
  const widest = Math.max(...values.map(v => measureText(String(v ?? ""))));
  return Math.min(maxW, Math.max(minW, Math.ceil(widest) + pad));
};

const CELL_W=60, MARK_W=82;
const ROW_H=44, VIRTUAL_ROW_THRESHOLD=120, VIRTUAL_OVERSCAN=14;
const MIN_ROLL_W=86, MAX_ROLL_W=140;
const MIN_SYS_W=90, MAX_SYS_W=150;
const MIN_NAME_W=160, MAX_NAME_W=320;
const CLASS_COLORS=["#e04e20","#2563eb","#16a34a","#9333ea","#ea580c","#0891b2","#be185d","#ca8a04"];
const STORAGE_KEY = "attendx_v3";
const LOW_ATT_THRESHOLD = 25;
const DEFAULTER_THRESHOLD = 75;
const LOW_WEIGHTED_THRESHOLD = 50;
const MAX_UNDO = 30;

const DEFAULT_GRADES = [
  { label: "A+", min: 90 },
  { label: "A",  min: 85 },
  { label: "A-", min: 80 },
  { label: "B+", min: 75 },
  { label: "B",  min: 70 },
  { label: "B-", min: 65 },
  { label: "C+", min: 60 },
  { label: "C",  min: 55 },
  { label: "C-", min: 50 },
  { label: "D",  min: 45 },
  { label: "F",  min: 0 },
];

const emptyInfo = { faculty:"", session:"", teacher:"", subject:"", course:"" };

const sanitizeClasses = (input) => {
  if (!Array.isArray(input)) return [];
  return input.filter(Boolean).map((c, classIndex) => ({
    id: c.id || uid(),
    color: c.color || CLASS_COLORS[classIndex % CLASS_COLORS.length],
    type: c.type === "marks" ? "marks" : "attendance",
    info: { ...emptyInfo, ...(c.info || {}) },
    students: Array.isArray(c.students) ? c.students.filter(Boolean).map((st, studentIndex) => ({
      id: st.id ?? (1000 + studentIndex),
      roll: String(st.roll ?? st.studentId ?? String(studentIndex + 1).padStart(2, "0")),
      systemId: String(st.systemId ?? st.systemID ?? st.sysId ?? ""),
      name: String(st.name ?? st.studentName ?? "Unnamed Student"),
    })) : [],
    dates: Array.isArray(c.dates) ? c.dates : [],
    att: c.att && typeof c.att === "object" ? c.att : {},
    assessments: Array.isArray(c.assessments) ? c.assessments.filter(Boolean).map(a => ({
      id: a.id || uid(),
      name: String(a.name ?? "Assessment"),
      outOf: Number(a.outOf) || 10,
      type: String(a.type ?? "quiz"),
      weight: Number(a.weight) || 0,
    })) : [],
    marks: c.marks && typeof c.marks === "object" ? c.marks : {},
    grades: Array.isArray(c.grades) && c.grades.length ? c.grades.map(g => ({
      label: String(g.label ?? ""),
      min: Number(g.min) ?? 0,
    })) : [...DEFAULT_GRADES],
  }));
};

// ─── CSS ─────────────────────────────────────────────────────────────────────
const makeStyle = (rollW, sysW, nameW) => `:root{--roll-w:${rollW}px;--sys-w:${sysW}px;--name-w:${nameW}px;}`;
// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────
export default function App() {
  const [classes,          setClasses]         = useState([]);
  const [activeId,         setActiveId]        = useState(null);
  const [showModal,        setShowModal]        = useState(false);
  const [showCopyModal,    setShowCopyModal]    = useState(false);
  const [showAssessModal,  setShowAssessModal]  = useState(false);
  const [showImportModal,  setShowImportModal]  = useState(false);
  const [showEditAssess,   setShowEditAssess]   = useState(false); // edit single assessment
  const [editAssessId,     setEditAssessId]     = useState(null);
  const [copySourceId,     setCopySourceId]     = useState(null);
  const [saveStatus,       setSaveStatus]       = useState("saved");
  const [newRoll,          setNewRoll]          = useState("");
  const [newSystemId,      setNewSystemId]      = useState("");
  const [newName,          setNewName]          = useState("");
  const [newDate,          setNewDate]          = useState(todayStr());
  const [searchQ,          setSearchQ]          = useState("");
  const [lowAttOnly,       setLowAttOnly]       = useState(false);
  const [todayMode,        setTodayMode]        = useState(false);
  const [infoCollapsed,    setInfoCollapsed]    = useState(false);
  const [insightsOpen,     setInsightsOpen]     = useState(true);
  const [insightsMaximized,setInsightsMaximized]= useState(false);
  const [draft,            setDraft]            = useState({faculty:"",session:"",teacher:"",subject:"",course:"",type:"attendance"});
  const [assessDraft,      setAssessDraft]      = useState({name:"Quiz 1",outOf:"10",weight:"",type:"quiz"});
  const [importData,       setImportData]       = useState(null); // {rows:[{roll,name}], fileName, warnings:[]}
  const [importDragOver,   setImportDragOver]   = useState(false);
  const [darkMode,         setDarkMode]         = useState(() => {
    try { return localStorage.getItem("attendx_dark") === "1"; } catch { return false; }
  });
  const [showGradeModal,   setShowGradeModal]   = useState(false);
  const [gradeDraft,       setGradeDraft]       = useState([]);
  const [tableScrollTop,   setTableScrollTop]   = useState(0);
  const [tableViewportH,   setTableViewportH]   = useState(640);

  // ── undo / redo ──
  const undoStack = useRef([]);
  const redoStack = useRef([]);
  const [undoLen, setUndoLen] = useState(0);
  const [redoLen, setRedoLen] = useState(0);
  const skipHistory = useRef(false);

  const nextSid   = useRef(1000);
  const refs      = useRef({});
  const tableScrollRef = useRef(null);
  const initDone  = useRef(false);
  const saveTimer = useRef(null);
  const fileRef   = useRef(null);
  const marksTemplateFileRef = useRef(null);
  const isElectron = typeof window !== "undefined" && !!window.electronAPI;

  // ── LOAD ─────────────────────────────────────────────────────────────────────
  useEffect(() => {
    async function load() {
      try {
        let saved = null;
        if (isElectron) { saved = await window.electronAPI.loadData(); }
        else { const raw = localStorage.getItem(STORAGE_KEY); if (raw) saved = JSON.parse(raw); }
        if (saved?.classes?.length) {
          const cleanClasses = sanitizeClasses(saved.classes);
          setClasses(cleanClasses);
          setActiveId(saved.activeId && cleanClasses.some(c => c.id === saved.activeId) ? saved.activeId : cleanClasses[0]?.id ?? null);
          const maxId = cleanClasses
            .flatMap(c => c.students ?? [])
            .reduce((m, s) => Math.max(m, Number(s.id) || 0), 999);
          nextSid.current = maxId + 1;
        }
      } catch(e) { console.warn("Load failed:", e); }
      finally    { initDone.current = true; }
    }
    load();
  }, []); // eslint-disable-line

  // ── AUTO-SAVE ─────────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!initDone.current) return;
    setSaveStatus("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      const payload = { classes, activeId };
      try {
        if (isElectron) { await window.electronAPI.saveData(payload); }
        else { localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); }
        setSaveStatus("saved");
      } catch(e) { setSaveStatus("saved"); }
    }, 600);
    return () => clearTimeout(saveTimer.current);
  }, [classes, activeId]); // eslint-disable-line

  useEffect(() => { setSearchQ(""); setLowAttOnly(false); setTodayMode(false); }, [activeId]);
  useEffect(() => {
    setTableScrollTop(0);
    if (tableScrollRef.current) tableScrollRef.current.scrollTop = 0;
  }, [activeId, searchQ, lowAttOnly, todayMode]);

  // ── Dark mode toggle ─────────────────────────────────────────────────────────
  useEffect(() => {
    document.documentElement.classList.toggle("dark", darkMode);
    try { localStorage.setItem("attendx_dark", darkMode ? "1" : "0"); } catch {}
  }, [darkMode]);

  // ── Undo / Redo history ───────────────────────────────────────────────────────
  const pushUndo = useCallback((snapshot) => {
    undoStack.current.push(snapshot);
    if (undoStack.current.length > MAX_UNDO) undoStack.current.shift();
    redoStack.current = [];
    setUndoLen(undoStack.current.length);
    setRedoLen(0);
  }, []);

  const setClassesWithHistory = useCallback((updater) => {
    setClasses(prev => {
      const snapshot = JSON.stringify(prev);
      const next = typeof updater === "function" ? updater(prev) : updater;
      if (JSON.stringify(next) !== snapshot) {
        pushUndo(snapshot);
      }
      return next;
    });
  }, [pushUndo]);

  const undo = useCallback(() => {
    if (!undoStack.current.length) return;
    const snapshot = undoStack.current.pop();
    redoStack.current.push(JSON.stringify(classes));
    skipHistory.current = true;
    setClasses(JSON.parse(snapshot));
    setUndoLen(undoStack.current.length);
    setRedoLen(redoStack.current.length);
  }, [classes]);

  const redo = useCallback(() => {
    if (!redoStack.current.length) return;
    const snapshot = redoStack.current.pop();
    undoStack.current.push(JSON.stringify(classes));
    skipHistory.current = true;
    setClasses(JSON.parse(snapshot));
    setUndoLen(undoStack.current.length);
    setRedoLen(redoStack.current.length);
  }, [classes]);

  // Ctrl+Z / Ctrl+Y global shortcuts
  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
      if ((e.ctrlKey || e.metaKey) && (e.key === "y" || (e.key === "z" && e.shiftKey))) { e.preventDefault(); redo(); }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [undo, redo]);

  // ── derived ───────────────────────────────────────────────────────────────────
  const active      = classes.find(c => c.id === activeId) ?? null;
  const isMarks     = active?.type === "marks";
  const isAtt       = !isMarks;
  const sortedDates = useMemo(() => active && isAtt ? [...(active.dates ?? [])].sort() : [], [active, isAtt]);
  const todayIso    = todayStr();
  const visibleDates = useMemo(() => todayMode && isAtt ? sortedDates.filter(d => d === todayIso) : sortedDates, [todayMode, isAtt, sortedDates, todayIso]);
  const assessments = useMemo(() => active?.assessments ?? [], [active]);

  // Safe cell reader must be initialized before attendance summaries are computed.
  const getCell = (c, sid, d) => c?.att?.[`${sid}_${d}`] ?? "";

  // Weight helpers
  const totalWeight = assessments.reduce((s, a) => s + (parseFloat(a.weight) || 0), 0);
  const hasWeightedAssessments = assessments.some(a => parseFloat(a.weight) > 0);
  const weightsOk   = Math.abs(totalWeight - 100) < 0.01;

  // ── AUTO-FIT ──────────────────────────────────────────────────────────────────
  const { rollW, sysW, nameW } = useMemo(() => {
    if (!active?.students?.length) return { rollW: MIN_ROLL_W, sysW: MIN_SYS_W, nameW: MIN_NAME_W };
    const rollVals = ["Student ID", ...active.students.map(s => s.roll ?? "")];
    const sysVals  = ["System ID", ...active.students.map(s => s.systemId ?? "")];
    const nameVals = ["Student Name", ...active.students.map(s => s.name ?? "")];
    return {
      rollW: autoFitWidth(rollVals, MIN_ROLL_W, MAX_ROLL_W, 28),
      sysW:  autoFitWidth(sysVals, MIN_SYS_W, MAX_SYS_W, 28),
      nameW: autoFitWidth(nameVals, MIN_NAME_W, MAX_NAME_W, 32),
    };
  }, [active?.students]);

  const updActive = fn => {
    if (active) setClassesWithHistory(cs => cs.map(c => c.id === activeId ? fn(c) : c));
  };

  // ── create class ─────────────────────────────────────────────────────────────
  const createClass = () => {
    const cls = {
      id: uid(), color: CLASS_COLORS[classes.length % CLASS_COLORS.length],
      type: draft.type || "attendance",
      info: { faculty:draft.faculty, session:draft.session, teacher:draft.teacher, subject:draft.subject, course:draft.course },
      students:[], dates:[todayStr()], att:{}, assessments:[], marks:{},
      grades: [...DEFAULT_GRADES],
    };
    setClassesWithHistory(cs => [...cs, cls]);
    setActiveId(cls.id);
    setShowModal(false);
    setDraft({faculty:"",session:"",teacher:"",subject:"",course:"",type:"attendance"});
  };

  const deleteClass = (id, e) => {
    e.stopPropagation();
    if (!window.confirm("Delete this class and all its data?")) return;
    const remaining = classes.filter(c => c.id !== id);
    setClassesWithHistory(remaining);
    if (activeId === id) setActiveId(remaining[0]?.id ?? null);
  };

  // ── add student ───────────────────────────────────────────────────────────────
  const addStudent = () => {
    if (!active) return;
    const rollParts = newRoll.split(",").map(r => r.trim()).filter(Boolean);
    const systemParts = newSystemId.split(",").map(r => r.trim()).filter(Boolean);
    const nameParts = newName.split(",").map(n => n.trim()).filter(Boolean);
    if (!nameParts.length) return;
    const base = active.students.length;
    const newStudents = nameParts.map((name, i) => ({
      id:   nextSid.current++,
      roll: rollParts[i] || String(base + i + 1).padStart(2, "0"),
      systemId: systemParts[i] || "",
      name,
    }));
    updActive(c => ({ ...c, students: [...c.students, ...newStudents] }));
    setNewName(""); setNewRoll(""); setNewSystemId("");
  };

  // ── EXCEL IMPORT ──────────────────────────────────────────────────────────────
  const parseImportFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        await loadXlsx();
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        // Find which columns are Student ID, System ID and Name
        let rollCol = -1, systemCol = -1, nameCol = -1;
        // Try to detect header row (look in first 3 rows)
        let dataStart = 0;
        for (let r = 0; r < Math.min(3, raw.length); r++) {
          const row = raw[r].map(c => String(c).toLowerCase().trim());
          const systemIdx = row.findIndex(c => c.includes("system") || c.includes("cms") || c.includes("registration"));
          const rollIdx = row.findIndex(c =>
            c.includes("student id") || c.includes("student_id") || c.includes("roll") || c === "id" || c === "no" || c === "#"
          );
          const nameIdx = row.findIndex(c => c.includes("name") || c.includes("student"));
          if (rollIdx >= 0 || systemIdx >= 0 || nameIdx >= 0) {
            rollCol = rollIdx;
            systemCol = systemIdx;
            nameCol = nameIdx >= 0 ? nameIdx : (rollIdx === 0 ? 1 : 0);
            dataStart = r + 1;
            break;
          }
        }
        // fallback: assume col 0 = Student ID, col 1 = System ID, col 2 = Name
        if (rollCol < 0 && systemCol < 0 && nameCol < 0) { rollCol = 0; systemCol = 1; nameCol = 2; dataStart = 0; }
        // if file has only two columns, assume Student ID + Name
        if (nameCol < 0) { nameCol = rollCol === 0 ? 1 : 0; }
        if (rollCol < 0) { rollCol = nameCol === 0 ? 1 : 0; }

        const rows = [];
        const warnings = [];
        for (let r = dataStart; r < raw.length; r++) {
          const row = raw[r];
          const rollRaw = String(row[rollCol] ?? "").trim();
          const systemRaw = systemCol >= 0 ? String(row[systemCol] ?? "").trim() : "";
          const nameRaw = String(row[nameCol] ?? "").trim();
          if (!nameRaw && !rollRaw && !systemRaw) continue; // skip empty rows
          if (!nameRaw) { warnings.push(`Row ${r+1}: missing name (student ID: ${rollRaw || "—"})`); continue; }
          rows.push({ roll: rollRaw || String(r - dataStart + 1).padStart(2,"0"), systemId: systemRaw, name: nameRaw });
        }
        setImportData({ rows, fileName: file.name, warnings });
      } catch(err) {
        setImportData({ rows: [], fileName: file.name, warnings: [], error: "Could not read file. Make sure it's a valid .xlsx or .csv file." });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleImportDrop = (e) => {
    e.preventDefault(); setImportDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) parseImportFile(file);
  };

  const confirmImport = () => {
    if (!importData?.rows?.length || !active) return;
    const existingRolls = new Set(active.students.map(s => s.roll).filter(Boolean));
    const existingSystemIds = new Set(active.students.map(s => s.systemId).filter(Boolean));
    const existingNames = new Set(active.students.map(s => String(s.name ?? "").trim().toLowerCase()).filter(Boolean));
    let added = 0;
    const toAdd = importData.rows.filter(r => {
      if (r.roll && existingRolls.has(r.roll)) return false;
      if (r.systemId && existingSystemIds.has(r.systemId)) return false;
      if (existingNames.has(r.name.trim().toLowerCase())) return false;
      return true;
    }).map(r => { added++; return { id: nextSid.current++, roll: r.roll, systemId: r.systemId || "", name: r.name }; });
    updActive(c => ({ ...c, students: [...c.students, ...toAdd] }));
    setShowImportModal(false);
    setImportData(null);
  };

  // ── copy roster ───────────────────────────────────────────────────────────────
  const openCopyModal = () => { setCopySourceId(null); setShowCopyModal(true); };
  const confirmCopyRoster = () => {
    if (!copySourceId || !active) return;
    const src = classes.find(c => c.id === copySourceId);
    if (!src?.students?.length) return;
    const existingRolls = new Set(active.students.map(s => s.roll).filter(Boolean));
    const existingSystemIds = new Set(active.students.map(s => s.systemId).filter(Boolean));
    const existingNames = new Set(active.students.map(s => String(s.name ?? "").trim().toLowerCase()).filter(Boolean));
    const toAdd = src.students
      .filter(s => {
        if (s.roll && existingRolls.has(s.roll)) return false;
        if (s.systemId && existingSystemIds.has(s.systemId)) return false;
        if (existingNames.has(String(s.name ?? "").trim().toLowerCase())) return false;
        return true;
      })
      .map(s => ({ ...s, id: nextSid.current++ }));
    if (!toAdd.length) { alert("All students from that class are already in this class."); return; }
    updActive(c => ({ ...c, students: [...c.students, ...toAdd] }));
    setShowCopyModal(false); setCopySourceId(null);
  };

  // ── attendance ops ────────────────────────────────────────────────────────────
  const addDate    = () => { if (!active || !newDate || (active.dates??[]).includes(newDate)) return; updActive(c=>({...c,dates:[...(c.dates??[]),newDate]})); };
  const removeDate = d  => updActive(c=>({...c,dates:(c.dates??[]).filter(x=>x!==d),att:Object.fromEntries(Object.entries(c.att??{}).filter(([k])=>!k.endsWith("_"+d)))}));
  const removeStudent = sid => updActive(c=>({...c,students:c.students.filter(s=>s.id!==sid),att:Object.fromEntries(Object.entries(c.att??{}).filter(([k])=>!k.startsWith(sid+"_"))),marks:Object.fromEntries(Object.entries(c.marks??{}).filter(([k])=>!k.startsWith(sid+"_")))}));
  const setCell    = (sid,d,v) => updActive(c=>({...c,att:{...c.att,[`${sid}_${d}`]:norm(v)}}));
  const scopedDates = (c) => todayMode ? [todayIso] : [...(c.dates??[])].sort();
  const markRestP  = () => updActive(c=>{ const dates=scopedDates(c); const clsDates=todayMode && !(c.dates??[]).includes(todayIso) ? [...(c.dates??[]),todayIso] : c.dates; const a={...c.att}; c.students.forEach(s=>dates.forEach(d=>{const k=`${s.id}_${d}`;if(!a[k])a[k]="P";})); return {...c,dates:clsDates,att:a}; });
  const markRestA  = () => updActive(c=>{ const dates=scopedDates(c); const clsDates=todayMode && !(c.dates??[]).includes(todayIso) ? [...(c.dates??[]),todayIso] : c.dates; const a={...c.att}; c.students.forEach(s=>dates.forEach(d=>{const k=`${s.id}_${d}`;if(!a[k])a[k]="A";})); return {...c,dates:clsDates,att:a}; });
  const toggleTodayMode = () => {
    if (!active || !isAtt) return;
    if (!todayMode && !(active.dates??[]).includes(todayIso)) {
      updActive(c => ({ ...c, dates: [...(c.dates??[]), todayIso] }));
    }
    setNewDate(todayIso);
    setTodayMode(v => !v);
  };

  // ── marks ops ─────────────────────────────────────────────────────────────────
  const addAssessment = () => {
    const name  = assessDraft.name.trim() || "Assessment";
    const outOf = parseFloat(assessDraft.outOf) || 10;
    const type  = assessDraft.type || "quiz";
    const weight = parseFloat(assessDraft.weight) || 0;
    if (weight > 0) {
      const currentTotal = assessments.reduce((s, a) => s + (parseFloat(a.weight) || 0), 0);
      if (currentTotal + weight > 100) {
        alert(`Cannot add: total weight would be ${(currentTotal + weight).toFixed(1)}% (max 100%). Current total is ${currentTotal.toFixed(1)}%.`);
        return;
      }
    }
    updActive(c => ({ ...c, assessments: [...(c.assessments??[]), { id:uid(), name, outOf, type, weight }] }));
    setShowAssessModal(false);
    setAssessDraft({name:"Quiz 1",outOf:"10",weight:"",type:"quiz"});
  };

  // Edit assessment (name, outOf, type, weight) after creation
  const openEditAssess = (aid) => {
    if (!active) return;
    const a = (active.assessments??[]).find(x => x.id === aid);
    if (!a) return;
    setEditAssessId(aid);
    setAssessDraft({ name: a.name, outOf: String(a.outOf), weight: String(a.weight||""), type: a.type||"quiz" });
    setShowEditAssess(true);
  };
  const saveEditAssess = () => {
    if (!editAssessId) return;
    const newWeight = parseFloat(assessDraft.weight) || 0;
    if (newWeight > 0) {
      const otherTotal = assessments.filter(a => a.id !== editAssessId).reduce((s, a) => s + (parseFloat(a.weight) || 0), 0);
      if (otherTotal + newWeight > 100) {
        alert(`Cannot save: total weight would be ${(otherTotal + newWeight).toFixed(1)}% (max 100%). Other assessments total ${otherTotal.toFixed(1)}%.`);
        return;
      }
    }
    updActive(c => ({
      ...c,
      assessments: (c.assessments??[]).map(a => a.id === editAssessId
        ? { ...a, name: assessDraft.name.trim()||a.name, outOf: parseFloat(assessDraft.outOf)||a.outOf, type: assessDraft.type||a.type, weight: newWeight }
        : a)
    }));
    setShowEditAssess(false); setEditAssessId(null);
  };

  const removeAssessment = aid => updActive(c=>({...c,assessments:(c.assessments??[]).filter(a=>a.id!==aid),marks:Object.fromEntries(Object.entries(c.marks??{}).filter(([k])=>!k.endsWith("_"+aid)))}));

  const setMark  = (sid,aid,v) => { const raw=v==="A"||v==="a"?"A":v.replace(/[^0-9.]/g,""); updActive(c=>({...c,marks:{...c.marks,[`${sid}_${aid}`]:raw}})); };
  const getMark  = (c,sid,aid) => c.marks?.[`${sid}_${aid}`] ?? "";
  const markCls  = (c,sid,aid) => {
    const v=getMark(c,sid,aid); if(v==="")return ""; if(v==="A")return "absent-m";
    const assess=(c.assessments??[]).find(a=>a.id===aid); if(!assess)return "";
    const n=parseFloat(v); if(isNaN(n))return "";
    const pct=n/assess.outOf*100;
    if(pct>=100)return "full"; if(pct>=60)return "good"; return "low-m";
  };

  const setAssessWeight = (aid, val) => {
    const w = val === "" ? 0 : Math.max(0, parseFloat(val) || 0);
    const otherTotal = assessments.filter(a => a.id !== aid).reduce((s, a) => s + (parseFloat(a.weight) || 0), 0);
    const clamped = Math.min(w, 100 - otherTotal);
    updActive(c => ({ ...c, assessments: (c.assessments??[]).map(a => a.id===aid ? {...a, weight: Math.max(0, clamped)} : a) }));
  };

  // ── summaries ─────────────────────────────────────────────────────────────────
  function stuAttSum(c,sid) {
    const sd=[...(c.dates??[])].sort(); let p=0,a=0,tot=0;
    sd.forEach(d=>{ const v=getCell(c,sid,d); if(isP(v)){p++;tot++;} else if(isA(v)){a++;tot++;} });
    if(!tot)return null;
    return {p,a,pct:Math.round(p/tot*100)};
  }

  // Raw total (sum of marks)
  const stuMarksSum = (c, sid) => {
    if (!(c.assessments?.length)) return null;
    let earned=0,total=0,count=0;
    c.assessments.forEach(a => {
      const v=getMark(c,sid,a.id);
      if(v!==""&&v!=="A"){const n=parseFloat(v);if(!isNaN(n)){earned+=clamp(n,0,a.outOf);total+=a.outOf;count++;}}
    });
    if(!count)return null;
    return {earned:earned.toFixed(1),total:total.toFixed(1),pct:Math.round(earned/total*100)};
  };

  // Weighted total: Σ (mark/outOf * weight)  → out of totalWeight
  const stuWeightedSum = (c, sid) => {
    if (!(c.assessments?.length)) return null;
    let weightedEarned = 0, weightedTotal = 0, count = 0;
    c.assessments.forEach(a => {
      const w = parseFloat(a.weight) || 0;
      if (w <= 0) return;
      const v = getMark(c, sid, a.id);
      if (v !== "" && v !== "A") {
        const n = parseFloat(v);
        if (!isNaN(n)) {
          weightedEarned += (clamp(n,0,a.outOf) / a.outOf) * w;
          weightedTotal  += w;
          count++;
        }
      }
    });
    if (!count) return null;
    // pct = earned out of total assigned weight
    const pct = weightedTotal > 0 ? Math.round(weightedEarned / weightedTotal * 100) : 0;
    return { earned: weightedEarned.toFixed(1), total: weightedTotal.toFixed(1), pct };
  };

  const dateAttSum = (c,d) => { let p=0,a=0; c.students.forEach(s=>{const v=getCell(c,s.id,d);if(isP(v))p++;else if(isA(v))a++;}); return {p,a}; };
  const assessColAvg = (c,aid) => {
    const assess=(c.assessments??[]).find(a=>a.id===aid); if(!assess)return null;
    let sum=0,cnt=0;
    c.students.forEach(s=>{const v=getMark(c,s.id,aid);if(v!==""&&v!=="A"){const n=parseFloat(v);if(!isNaN(n)){sum+=n;cnt++;}}});
    if(!cnt)return null;
    return (sum/cnt).toFixed(1);
  };
  const classAttendancePct = (cls) => {
    if (!cls || cls.type === "marks") return null;
    let p = 0, a = 0;
    (cls.students ?? []).forEach(s => {
      (cls.dates ?? []).forEach(d => {
        const v = getCell(cls, s.id, d);
        if (isP(v)) p++;
        else if (isA(v)) a++;
      });
    });
    const total = p + a;
    return total ? Math.round((p / total) * 100) : null;
  };
  const classMarksPct = (cls) => {
    if (!cls || cls.type !== "marks" || !(cls.students?.length) || !(cls.assessments?.some(a => parseFloat(a.weight) > 0))) return null;
    let sum = 0, count = 0;
    cls.students.forEach(s => {
      const wm = stuWeightedSum(cls, s.id);
      if (wm) { sum += wm.pct; count++; }
    });
    return count ? Math.round(sum / count) : null;
  };
  const classHealth = (cls) => {
    const pct = cls?.type === "marks" ? classMarksPct(cls) : classAttendancePct(cls);
    const tone = pct === null ? "#475569" : pct >= 75 ? "#22c55e" : pct >= 50 ? "#f59e0b" : "#ef4444";
    return { pct, tone };
  };
  const heatForStudent = (cls, sid) => [...(cls?.dates ?? [])]
    .sort()
    .slice(-18)
    .map(d => {
      const v = getCell(cls, sid, d);
      return isP(v) ? "p" : isA(v) ? "a" : "blank";
    });
  const attendanceSummaryByStudent = useMemo(() => {
    const map = new Map();
    if (!active || !isAtt) return map;
    (active.students ?? []).forEach(s => {
      let p = 0, a = 0, tot = 0;
      sortedDates.forEach(d => {
        const v = getCell(active, s.id, d);
        if (isP(v)) { p++; tot++; }
        else if (isA(v)) { a++; tot++; }
      });
      if (tot) map.set(s.id, { p, a, pct: Math.round(p / tot * 100) });
    });
    return map;
  }, [active, isAtt, sortedDates]);
  const visibleAttendanceSummaryByStudent = useMemo(() => {
    const map = new Map();
    if (!active || !isAtt) return map;
    (active.students ?? []).forEach(s => {
      let p = 0, a = 0, tot = 0;
      visibleDates.forEach(d => {
        const v = getCell(active, s.id, d);
        if (isP(v)) { p++; tot++; }
        else if (isA(v)) { a++; tot++; }
      });
      if (tot) map.set(s.id, { p, a, pct: Math.round(p / tot * 100) });
    });
    return map;
  }, [active, isAtt, visibleDates]);
  const weightedSummaryByStudent = useMemo(() => {
    const map = new Map();
    if (!active || !isMarks || !(active.assessments?.length)) return map;
    (active.students ?? []).forEach(s => {
      let weightedEarned = 0, weightedTotal = 0, count = 0;
      (active.assessments ?? []).forEach(a => {
        const w = parseFloat(a.weight) || 0;
        if (w <= 0) return;
        const v = getMark(active, s.id, a.id);
        if (v !== "" && v !== "A") {
          const n = parseFloat(v);
          if (!isNaN(n)) {
            weightedEarned += (clamp(n, 0, a.outOf) / a.outOf) * w;
            weightedTotal += w;
            count++;
          }
        }
      });
      if (count) {
        const pct = weightedTotal > 0 ? Math.round(weightedEarned / weightedTotal * 100) : 0;
        map.set(s.id, { earned: weightedEarned.toFixed(1), total: weightedTotal.toFixed(1), pct });
      }
    });
    return map;
  }, [active, isMarks]);
  const activeWeightedAverage = useMemo(() => {
    if (!isMarks || !hasWeightedAssessments || !weightedSummaryByStudent.size) return null;
    let sum = 0, count = 0;
    weightedSummaryByStudent.forEach(wm => { sum += wm.pct; count++; });
    return count ? Math.round(sum / count) : null;
  }, [isMarks, hasWeightedAssessments, weightedSummaryByStudent]);
  const dateSummaryByDate = useMemo(() => {
    const map = new Map();
    if (!active || !isAtt) return map;
    visibleDates.forEach(d => {
      let p = 0, a = 0;
      (active.students ?? []).forEach(s => {
        const v = getCell(active, s.id, d);
        if (isP(v)) p++;
        else if (isA(v)) a++;
      });
      map.set(d, { p, a });
    });
    return map;
  }, [active, isAtt, visibleDates]);
  const assessAvgByAssessment = useMemo(() => {
    const map = new Map();
    if (!active || !isMarks) return map;
    (active.assessments ?? []).forEach(a => {
      let sum = 0, cnt = 0;
      (active.students ?? []).forEach(s => {
        const v = getMark(active, s.id, a.id);
        if (v !== "" && v !== "A") {
          const n = parseFloat(v);
          if (!isNaN(n)) { sum += n; cnt++; }
        }
      });
      if (cnt) map.set(a.id, (sum / cnt).toFixed(1));
    });
    return map;
  }, [active, isMarks]);
  const lowAttendanceStudents = useMemo(() => active && isAtt
    ? (active.students ?? []).filter(s => {
        const sm = attendanceSummaryByStudent.get(s.id);
        return sm && sm.pct < LOW_ATT_THRESHOLD;
      })
    : [], [active, isAtt, attendanceSummaryByStudent]);
  const filteredStudents = useMemo(() => active
    ? (active.students ?? []).filter(s => {
        const sm = isAtt ? attendanceSummaryByStudent.get(s.id) : null;
        if (lowAttOnly && (!sm || sm.pct >= LOW_ATT_THRESHOLD)) return false;
        if (!searchQ.trim()) return true;
        const q = searchQ.trim().toLowerCase();
        return String(s.name ?? "").toLowerCase().includes(q)
          || String(s.roll ?? "").toLowerCase().includes(q)
          || String(s.systemId ?? "").toLowerCase().includes(q);
      })
    : [], [active, isAtt, attendanceSummaryByStudent, lowAttOnly, searchQ]);
  const rowVirtual = useMemo(() => {
    const enabled = filteredStudents.length > VIRTUAL_ROW_THRESHOLD;
    if (!enabled) {
      return { enabled, start: 0, end: filteredStudents.length, topPad: 0, bottomPad: 0 };
    }
    const viewportRows = Math.max(1, Math.ceil(tableViewportH / ROW_H));
    const start = Math.max(0, Math.floor(tableScrollTop / ROW_H) - VIRTUAL_OVERSCAN);
    const end = Math.min(filteredStudents.length, start + viewportRows + VIRTUAL_OVERSCAN * 2);
    return {
      enabled,
      start,
      end,
      topPad: start * ROW_H,
      bottomPad: Math.max(0, (filteredStudents.length - end) * ROW_H)
    };
  }, [filteredStudents.length, tableScrollTop, tableViewportH]);
  const renderedStudents = useMemo(
    () => rowVirtual.enabled ? filteredStudents.slice(rowVirtual.start, rowVirtual.end) : filteredStudents,
    [filteredStudents, rowVirtual]
  );
  const attTableColSpan = visibleDates.length + 4;
  const marksTableColSpan = assessments.length + 4;
  const handleTableScroll = useCallback((e) => {
    const el = e.currentTarget;
    setTableScrollTop(el.scrollTop);
    setTableViewportH(el.clientHeight || 640);
  }, []);
  const focusCellOrScroll = useCallback((key, rowIndex) => {
    const current = refs.current[key];
    if (current) { current.focus(); return; }
    const scrollEl = tableScrollRef.current;
    if (!scrollEl) return;
    scrollEl.scrollTop = Math.max(0, rowIndex * ROW_H - ROW_H * 2);
    window.setTimeout(() => refs.current[key]?.focus(), 0);
  }, []);
  const overall = useMemo(() => {
    if (!active || !isAtt) return { p: 0, a: 0 };
    return [...attendanceSummaryByStudent.values()].reduce((acc, sm) => ({ p: acc.p + sm.p, a: acc.a + sm.a }), { p: 0, a: 0 });
  }, [active, isAtt, attendanceSummaryByStudent]);
  const visibleOverall = useMemo(() => {
    if (!active || !isAtt) return { p: 0, a: 0 };
    return [...visibleAttendanceSummaryByStudent.values()].reduce((acc, sm) => ({ p: acc.p + sm.p, a: acc.a + sm.a }), { p: 0, a: 0 });
  }, [active, isAtt, visibleAttendanceSummaryByStudent]);
  const lowAttCount = lowAttendanceStudents.length;
  const overallPct = overall.p + overall.a ? Math.round(overall.p / (overall.p + overall.a) * 100) : null;
  const visiblePct = visibleOverall.p + visibleOverall.a ? Math.round(visibleOverall.p / (visibleOverall.p + visibleOverall.a) * 100) : null;
  const recentDates = sortedDates.slice(-5).reverse();
  const topRisks = useMemo(() => [...lowAttendanceStudents]
    .sort((a, b) => (attendanceSummaryByStudent.get(a.id)?.pct ?? 101) - (attendanceSummaryByStudent.get(b.id)?.pct ?? 101))
    .slice(0, 5), [lowAttendanceStudents, attendanceSummaryByStudent]);
  const attendanceDefaulterStudents = useMemo(() => active && isAtt
    ? (active.students ?? [])
        .map(s => ({ student: s, summary: attendanceSummaryByStudent.get(s.id) }))
        .filter(({ summary }) => summary && summary.pct < DEFAULTER_THRESHOLD)
        .sort((a, b) =>
          a.summary.pct - b.summary.pct ||
          String(a.student.roll ?? "").localeCompare(String(b.student.roll ?? "")) ||
          String(a.student.name ?? "").localeCompare(String(b.student.name ?? ""))
        )
    : [], [active, isAtt, attendanceSummaryByStudent]);
  const lowWeightedStudents = useMemo(() => active && isMarks && hasWeightedAssessments
    ? (active.students ?? [])
        .map(s => ({ student: s, weighted: weightedSummaryByStudent.get(s.id) }))
        .filter(({ weighted }) => weighted && weighted.pct < LOW_WEIGHTED_THRESHOLD)
        .sort((a, b) =>
          a.weighted.pct - b.weighted.pct ||
          String(a.student.roll ?? "").localeCompare(String(b.student.roll ?? "")) ||
          String(a.student.name ?? "").localeCompare(String(b.student.name ?? ""))
        )
    : [], [active, isMarks, hasWeightedAssessments, weightedSummaryByStudent]);

  // ── Grading helpers ──────────────────────────────────────────────────────────
  const gradeScale = active?.grades ?? DEFAULT_GRADES;
  const getGrade = (pct) => {
    if (pct === null || pct === undefined) return null;
    const sorted = [...gradeScale].sort((a, b) => b.min - a.min);
    for (const g of sorted) { if (pct >= g.min) return g.label; }
    return sorted[sorted.length - 1]?.label ?? "F";
  };
  const gradeColorClass = (label) => {
    if (!label) return "";
    const l = label.toUpperCase();
    if (l.startsWith("A")) return "grade-a";
    if (l.startsWith("B")) return "grade-b";
    if (l.startsWith("C")) return "grade-c";
    if (l.startsWith("D")) return "grade-d";
    return "grade-f";
  };

  // ── Defaulter warning: how many absences until below threshold ────────────────
  const absencesUntilDefaulter = (c, sid) => {
    const sd = [...(c.dates ?? [])].sort();
    let p = 0, total = 0;
    sd.forEach(d => { const v = getCell(c, sid, d); if (isP(v)) { p++; total++; } else if (isA(v)) { total++; } });
    if (!total) return null;
    const pct = Math.round(p / total * 100);
    if (pct < DEFAULTER_THRESHOLD) return 0;
    let extraAbsences = 0;
    while (true) {
      extraAbsences++;
      const newTotal = total + extraAbsences;
      const newPct = Math.round(p / newTotal * 100);
      if (newPct < DEFAULTER_THRESHOLD) return extraAbsences;
      if (extraAbsences > 999) return null;
    }
  };

  // ── Grading modal ────────────────────────────────────────────────────────────
  const openGradeModal = () => {
    setGradeDraft((active?.grades ?? DEFAULT_GRADES).map(g => ({ ...g })));
    setShowGradeModal(true);
  };
  const saveGrades = () => {
    const cleaned = gradeDraft
      .filter(g => g.label.trim())
      .map(g => ({ label: g.label.trim(), min: Math.max(0, Math.min(100, Number(g.min) || 0)) }))
      .sort((a, b) => b.min - a.min);
    if (!cleaned.length) { alert("Add at least one grade."); return; }
    updActive(c => ({ ...c, grades: cleaned }));
    setShowGradeModal(false);
  };

  // ── keyboard nav ──────────────────────────────────────────────────────────────
  const navKeyAtt = (e,sid,dIdx) => {
    if(!["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter","Tab"].includes(e.key))return;
    e.preventDefault();
    const sts=filteredStudents; const sIdx=sts.findIndex(s=>s.id===sid);
    let ns=sIdx,nd=dIdx;
    if(e.key==="ArrowDown"||e.key==="Enter")ns=Math.min(sIdx+1,sts.length-1);
    else if(e.key==="ArrowUp")ns=Math.max(sIdx-1,0);
    else if(e.key==="ArrowRight"||(e.key==="Tab"&&!e.shiftKey))nd=Math.min(dIdx+1,visibleDates.length-1);
    else if(e.key==="ArrowLeft"||(e.key==="Tab"&&e.shiftKey))nd=Math.max(dIdx-1,0);
    focusCellOrScroll(`att_${sts[ns].id}_${visibleDates[nd]}`, ns);
  };
  const navKeyMarks = (e,sid,aIdx) => {
    if(!["ArrowUp","ArrowDown","ArrowLeft","ArrowRight","Enter","Tab"].includes(e.key))return;
    e.preventDefault();
    const sts=filteredStudents; const sIdx=sts.findIndex(s=>s.id===sid);
    let ns=sIdx,nd=aIdx;
    if(e.key==="ArrowDown"||e.key==="Enter")ns=Math.min(sIdx+1,sts.length-1);
    else if(e.key==="ArrowUp")ns=Math.max(sIdx-1,0);
    else if(e.key==="ArrowRight"||(e.key==="Tab"&&!e.shiftKey))nd=Math.min(aIdx+1,assessments.length-1);
    else if(e.key==="ArrowLeft"||(e.key==="Tab"&&e.shiftKey))nd=Math.max(aIdx-1,0);
    focusCellOrScroll(`mark_${sts[ns].id}_${assessments[nd]?.id}`, ns);
  };

  const setInfo  = (f,v) => updActive(c=>({...c,info:{...c.info,[f]:v}}));
  const onAddKey = e => { if(e.key==="Enter")addStudent(); };

  // ── Excel export ──────────────────────────────────────────────────────────────
  const buildAttWs = cls => {
    const sd=[...(cls.dates??[])].sort(); const inf=cls.info;
    const infoRows=[[(inf?.faculty)||""],[`Attendance Sheet (${(inf?.session)||""})`],[`Teacher's Name: ${(inf?.teacher)||""}`],[`Subject: ${(inf?.subject)||""}`],[`Course: ${(inf?.course)||""}`],[]];
    const hdr=["#","Student ID","System ID","Student Name",...sd.map(fmtDate),"Present","Absent","% Present"];
    const rows=(cls.students??[]).map((s,i)=>{ const sm=stuAttSum(cls,s.id); return [i+1,s.roll,s.systemId||"",s.name,...sd.map(d=>{const v=getCell(cls,s.id,d);return isP(v)?"P":isA(v)?"A":"";}),sm?.p??0,sm?.a??0,sm?sm.pct+"%":"-"]; });
    const sumP=["","","","── PRESENT ──",...sd.map(d=>dateAttSum(cls,d).p),"","",""];
    const sumA=["","","","── ABSENT ──",...sd.map(d=>dateAttSum(cls,d).a),"","",""];
    const ws=XLSX.utils.aoa_to_sheet([...infoRows,hdr,...rows,[],sumP,sumA]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:14},{wch:28},...sd.map(()=>({wch:8})),{wch:9},{wch:9},{wch:10}];
    return ws;
  };
  const buildMarksWs = cls => {
    const inf=cls.info; const asm=cls.assessments??[];
    const hasWeights = asm.some(a=>parseFloat(a.weight)>0);
    const gs = cls.grades ?? DEFAULT_GRADES;
    const getG = (pct) => { if(pct===null||pct===undefined)return "-"; const sorted=[...gs].sort((a,b)=>b.min-a.min); for(const g of sorted){if(pct>=g.min)return g.label;} return sorted[sorted.length-1]?.label??"F"; };
    const infoRows=[[(inf?.faculty)||""],[`Marks Register (${(inf?.session)||""})`],[`Teacher's Name: ${(inf?.teacher)||""}`],[`Subject: ${(inf?.subject)||""}`],[`Course: ${(inf?.course)||""}`],[]];
    const outOfRow=["","","","Out Of",...asm.map(a=>a.outOf),"Weighted %","Grade"];
    const weightRow=["","","","Weight %",...asm.map(a=>a.weight?(a.weight+"%"):"—"),"","",""];
    const hdr=["#","Student ID","System ID","Student Name",...asm.map(a=>`${a.name} (${a.type})`),"Weighted %","Grade"];
    const rows=(cls.students??[]).map((s,i)=>{
      const wm=hasWeights?stuWeightedSum(cls,s.id):null;
      const gradePct = wm ? wm.pct : null;
      return [i+1,s.roll,s.systemId||"",s.name,...asm.map(a=>getMark(cls,s.id,a.id)||""),wm?wm.pct+"%":"-",getG(gradePct)];
    });
    const avgRow=["","","","Average",...asm.map(a=>assessColAvg(cls,a.id)??"-"),"",""];
    const ws=XLSX.utils.aoa_to_sheet([...infoRows,outOfRow,weightRow.slice(0, hdr.length),hdr,...rows,[],avgRow]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:14},{wch:28},...asm.map(()=>({wch:14})),{wch:12},{wch:8}];
    return ws;
  };

  const cleanExportName = (value, fallback) => {
    const cleaned = String(value || fallback).replace(/[<>:"\/\\|?*\x00-\x1F]/g, "").trim();
    return (cleaned || fallback).slice(0, 60);
  };

  const buildMarksTemplateWs = cls => {
    const inf = cls.info ?? {};
    const asm = cls.assessments ?? [];
    const metaRows = [
      ["AttendX Marks Upload Template"],
      [`Subject: ${inf.subject || ""}`],
      [`Course: ${inf.course || ""}`],
      ["Fill only the assessment cells. Keep student rows, ID columns, and headers unchanged."],
      []
    ];
    const outOfRow = ["", "", "", "", "Out Of", ...asm.map(a => a.outOf)];
    const weightRow = ["", "", "", "", "Weight %", ...asm.map(a => a.weight ? `${a.weight}%` : "")];
    const assessmentIdRow = ["", "", "", "", "AttendX Assessment ID", ...asm.map(a => a.id)];
    const header = ["AttendX Student Key", "#", "Student ID", "System ID", "Student Name", ...asm.map(a => `${a.name} (${a.type})`)];
    const rows = (cls.students ?? []).map((s, i) => [s.id, i + 1, s.roll || "", s.systemId || "", s.name || "", ...asm.map(() => "")]);
    const sheetRows = [...metaRows, outOfRow, weightRow, assessmentIdRow, header, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(sheetRows);
    ws["!cols"] = [
      { wch: 18, hidden: true },
      { wch: 4 },
      { wch: 12 },
      { wch: 14 },
      { wch: 28 },
      ...asm.map(a => ({ wch: Math.max(14, Math.min(26, String(a.name || "Assessment").length + 6)) }))
    ];
    ws["!rows"] = sheetRows.map((_, i) => i === metaRows.length + 2 ? { hidden: true } : {});
    return ws;
  };

  const downloadMarksTemplate = async () => {
    if (!active || !isMarks) return;
    if (!(active.students?.length)) { alert("Add students before downloading a marks template."); return; }
    if (!(active.assessments?.length)) { alert("Add at least one assessment before downloading a marks template."); return; }
    await loadXlsx();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, buildMarksTemplateWs(active), "Marks Template");
    const fileName = cleanExportName(active.info?.subject || active.info?.course, "marks_template");
    XLSX.writeFile(wb, `attendx_${fileName}_marks_template.xlsx`);
  };

  const normalizeTemplateMark = (value) => {
    if (value === null || value === undefined) return { empty: true };
    if (typeof value === "number") {
      if (!Number.isFinite(value)) return { empty: true };
      return { empty: false, value: Number.isInteger(value) ? String(value) : String(Number(value.toFixed(2))) };
    }
    const text = String(value).trim();
    if (!text || text === "-") return { empty: true };
    const lower = text.toLowerCase();
    if (lower === "a" || lower === "absent" || lower === "ab") return { empty: false, value: "A" };
    const compact = text.replace(/,/g, "");
    if (!/^\d+(\.\d+)?$/.test(compact)) return { empty: false, error: `Invalid mark "${text}"` };
    const n = Number(compact);
    return { empty: false, value: Number.isInteger(n) ? String(n) : String(Number(n.toFixed(2))) };
  };

  const normalizeTemplateHeader = (value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s*\([^)]*\)\s*$/, "")
      .replace(/\s+/g, " ");

  const importMarksTemplateFile = (file) => {
    if (!file || !active || !isMarks) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        await loadXlsx();
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const lowerRows = raw.map(row => row.map(c => String(c).trim().toLowerCase()));
        const headerRowIdx = lowerRows.findIndex(row =>
          row.some(c => c === "student name" || c.includes("student name")) &&
          row.some(c => c === "student id" || c === "attendx student key")
        );

        if (headerRowIdx < 0) {
          alert("Could not find the marks template header. Download a fresh template and try again.");
          return;
        }

        const header = lowerRows[headerRowIdx];
        const originalHeader = raw[headerRowIdx] ?? [];
        const assessmentIdRow = raw[headerRowIdx - 1] ?? [];
        const keyCol = header.findIndex(c => c === "attendx student key");
        const rollCol = header.findIndex(c => c === "student id");
        const systemCol = header.findIndex(c => c === "system id");
        const nameCol = header.findIndex(c => c === "student name");

        const assessmentById = new Map((active.assessments ?? []).map(a => [String(a.id), a]));
        const assessmentByHeader = new Map();
        (active.assessments ?? []).forEach(a => {
          if (!assessmentByHeader.has(normalizeTemplateHeader(a.name))) {
            assessmentByHeader.set(normalizeTemplateHeader(a.name), a);
          }
          if (!assessmentByHeader.has(normalizeTemplateHeader(`${a.name} (${a.type})`))) {
            assessmentByHeader.set(normalizeTemplateHeader(`${a.name} (${a.type})`), a);
          }
        });

        const assessmentColumns = [];
        for (let col = 0; col < originalHeader.length; col++) {
          const id = String(assessmentIdRow[col] ?? "").trim();
          const byId = id ? assessmentById.get(id) : null;
          const byHeader = assessmentByHeader.get(normalizeTemplateHeader(originalHeader[col]));
          const assessment = byId || byHeader;
          if (assessment) assessmentColumns.push({ col, assessment });
        }

        if (!assessmentColumns.length) {
          alert("No matching assessment columns were found in this file.");
          return;
        }

        const studentById = new Map((active.students ?? []).map(s => [String(s.id), s]));
        const studentByRoll = new Map((active.students ?? []).filter(s => s.roll).map(s => [String(s.roll).trim(), s]));
        const studentBySystem = new Map((active.students ?? []).filter(s => s.systemId).map(s => [String(s.systemId).trim(), s]));
        const studentByName = new Map((active.students ?? []).filter(s => s.name).map(s => [String(s.name).trim().toLowerCase(), s]));
        const findStudent = (row) => {
          const key = keyCol >= 0 ? String(row[keyCol] ?? "").trim() : "";
          const systemId = systemCol >= 0 ? String(row[systemCol] ?? "").trim() : "";
          const roll = rollCol >= 0 ? String(row[rollCol] ?? "").trim() : "";
          const name = nameCol >= 0 ? String(row[nameCol] ?? "").trim().toLowerCase() : "";
          return (key && studentById.get(key)) ||
            (systemId && studentBySystem.get(systemId)) ||
            (roll && studentByRoll.get(roll)) ||
            (name && studentByName.get(name)) ||
            null;
        };

        const marksToSet = {};
        let filled = 0;
        let skippedExisting = 0;
        let invalid = 0;
        let unmatchedRows = 0;

        for (let r = headerRowIdx + 1; r < raw.length; r++) {
          const row = raw[r] ?? [];
          if (!row.some(c => String(c ?? "").trim())) continue;
          const student = findStudent(row);
          if (!student) {
            if (assessmentColumns.some(({ col }) => String(row[col] ?? "").trim())) unmatchedRows++;
            continue;
          }
          assessmentColumns.forEach(({ col, assessment }) => {
            const parsed = normalizeTemplateMark(row[col]);
            if (parsed.empty) return;
            if (parsed.error) { invalid++; return; }
            const key = `${student.id}_${assessment.id}`;
            if ((active.marks?.[key] ?? "") !== "") { skippedExisting++; return; }
            if (marksToSet[key] !== undefined) return;
            marksToSet[key] = parsed.value;
            filled++;
          });
        }

        if (filled) {
          updActive(c => ({ ...c, marks: { ...(c.marks ?? {}), ...marksToSet } }));
        }

        const summary = [`Filled ${filled} blank mark cell${filled === 1 ? "" : "s"} from ${file.name}.`];
        if (skippedExisting) summary.push(`Skipped ${skippedExisting} cell${skippedExisting === 1 ? "" : "s"} that already had marks in the app.`);
        if (invalid) summary.push(`Ignored ${invalid} invalid value${invalid === 1 ? "" : "s"}. Use numbers or A.`);
        if (unmatchedRows) summary.push(`Skipped ${unmatchedRows} row${unmatchedRows === 1 ? "" : "s"} that did not match an existing student.`);
        alert(summary.join("\n"));
      } catch(err) {
        alert("Could not read this marks template. Make sure it is a valid .xlsx, .xls, or .csv file.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleMarksTemplateUpload = (e) => {
    const file = e.target.files?.[0];
    if (file) importMarksTemplateFile(file);
    e.target.value = "";
  };

  const exportAll = async () => {
    if(!classes.length)return;
    await loadXlsx();
    const wb=XLSX.utils.book_new(); const used={};
    classes.forEach(cls=>{ let sn=((cls.info?.subject)||(cls.info?.course)||"Class").replace(/[:\\/?*\[\]]/g,"").slice(0,24)||"Sheet"; const sfx=cls.type==="marks"?" Marks":" Att"; let fn=(sn+sfx).slice(0,31); let n=2; while(used[fn]){fn=`${sn}_${n++}`;} used[fn]=1; XLSX.utils.book_append_sheet(wb, cls.type==="marks"?buildMarksWs(cls):buildAttWs(cls), fn); });
    XLSX.writeFile(wb,"attendx_all_classes.xlsx");
  };

  const exportDefaultersAll = async () => {
    await loadXlsx();
    const threshold = 75;
    const rows = [];

    classes.forEach(cls => {
      if (cls.type === "marks") return;

      const dates = [...(cls.dates ?? [])].sort();
      const inf = cls.info ?? {};

      (cls.students ?? []).forEach(s => {
        let p = 0;
        let a = 0;
        let total = 0;

        dates.forEach(d => {
          const v = getCell(cls, s.id, d);
          if (isP(v)) { p++; total++; }
          else if (isA(v)) { a++; total++; }
        });

        if (!total) return;

        const pct = Math.round((p / total) * 100);

        if (pct < threshold) {
          let level = "Below 75%";
          if (pct < 25) level = "Below 25%";
          else if (pct < 60) level = "Below 60%";

          rows.push({
            faculty: inf.faculty || "",
            session: inf.session || "",
            teacher: inf.teacher || "",
            subject: inf.subject || "",
            course: inf.course || "",
            studentId: s.roll || "",
            systemId: s.systemId || "",
            name: s.name || "",
            totalClasses: total,
            present: p,
            absent: a,
            attendancePct: pct,
            level,
          });
        }
      });
    });

    if (!rows.length) {
      alert(`No defaulters found below ${threshold}%.`);
      return;
    }

    rows.sort((x, y) =>
      String(x.course).localeCompare(String(y.course)) ||
      String(x.subject).localeCompare(String(y.subject)) ||
      x.attendancePct - y.attendancePct ||
      String(x.name).localeCompare(String(y.name))
    );

    const header = [
      "Sr. No.",
      "Faculty / Department",
      "Session",
      "Teacher",
      "Subject",
      "Course / Class / Section",
      "Student ID",
      "System ID",
      "Student Name",
      "Total Classes",
      "Present",
      "Absent",
      "Attendance %",
      "Defaulter Level"
    ];

    const data = rows.map((r, i) => [
      i + 1,
      r.faculty,
      r.session,
      r.teacher,
      r.subject,
      r.course,
      r.studentId,
      r.systemId,
      r.name,
      r.totalClasses,
      r.present,
      r.absent,
      `${r.attendancePct}%`,
      r.level
    ]);

    const titleRows = [
      ["AttendX - Combined Defaulter List"],
      [`Generated On: ${new Date().toLocaleDateString()}`],
      [`Threshold: Below ${threshold}%`],
      []
    ];

    const ws = XLSX.utils.aoa_to_sheet([...titleRows, header, ...data]);

    ws["!cols"] = [
      { wch: 8 },
      { wch: 28 },
      { wch: 18 },
      { wch: 22 },
      { wch: 22 },
      { wch: 24 },
      { wch: 14 },
      { wch: 16 },
      { wch: 28 },
      { wch: 14 },
      { wch: 10 },
      { wch: 10 },
      { wch: 14 },
      { wch: 16 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Defaulter List");
    XLSX.writeFile(wb, "attendx_combined_defaulter_list.xlsx");
  };
  const exportOne = async () => {
    if(!active)return;
    await loadXlsx();
    const wb=XLSX.utils.book_new();
    const sn=((active.info?.subject)||(active.info?.course)||"Sheet").replace(/[:\\/?*\[\]]/g,"").slice(0,31);
    XLSX.utils.book_append_sheet(wb, isMarks?buildMarksWs(active):buildAttWs(active), sn||"Sheet1");
    XLSX.writeFile(wb,`attendx_${sn||"export"}.xlsx`);
  };

  const exportInsights = async () => {
    if (!active) return;
    await loadXlsx();
    const inf = active.info ?? {};
    const sn = cleanExportName(inf.subject || inf.course || "class", "class");
    const wb = XLSX.utils.book_new();
    const title = isMarks ? "AttendX - Marks Insights" : "AttendX - Attendance Insights";
    const threshold = isMarks ? `Weighted score below ${LOW_WEIGHTED_THRESHOLD}%` : `Attendance below ${DEFAULTER_THRESHOLD}%`;
    const metaRows = [
      [title],
      [`Generated On: ${new Date().toLocaleDateString()}`],
      [`Faculty / Department: ${inf.faculty || ""}`],
      [`Session: ${inf.session || ""}`],
      [`Teacher: ${inf.teacher || ""}`],
      [`Subject: ${inf.subject || ""}`],
      [`Course: ${inf.course || ""}`],
      [`Threshold: ${threshold}`],
      []
    ];

    if (isMarks) {
      const header = ["#", "Student ID", "System ID", "Student Name", "Weighted Earned", "Weight Total", "Weighted %", "Grade"];
      const rows = lowWeightedStudents.map(({ student, weighted }, i) => [
        i + 1,
        student.roll || "",
        student.systemId || "",
        student.name || "",
        weighted.earned,
        weighted.total,
        `${weighted.pct}%`,
        getGrade(weighted.pct) || ""
      ]);
      const body = rows.length ? rows : [["", "", "", "No weighted scores below 50.", "", "", "", ""]];
      const ws = XLSX.utils.aoa_to_sheet([...metaRows, header, ...body]);
      ws["!cols"] = [{wch:4},{wch:14},{wch:16},{wch:28},{wch:16},{wch:14},{wch:12},{wch:10}];
      XLSX.utils.book_append_sheet(wb, ws, "Weighted Below 50");
    } else {
      const header = ["#", "Student ID", "System ID", "Student Name", "Present", "Absent", "Attendance %"];
      const rows = attendanceDefaulterStudents.map(({ student, summary }, i) => [
        i + 1,
        student.roll || "",
        student.systemId || "",
        student.name || "",
        summary.p,
        summary.a,
        `${summary.pct}%`
      ]);
      const body = rows.length ? rows : [["", "", "", "No students below 75% attendance.", "", "", ""]];
      const ws = XLSX.utils.aoa_to_sheet([...metaRows, header, ...body]);
      ws["!cols"] = [{wch:4},{wch:14},{wch:16},{wch:28},{wch:10},{wch:10},{wch:14}];
      XLSX.utils.book_append_sheet(wb, ws, "Attendance Below 75");
    }

    XLSX.writeFile(wb, `attendx_${sn}_insights.xlsx`);
  };

  // ── render ────────────────────────────────────────────────────────────────────
  const STYLE = makeStyle(rollW, sysW, nameW);

  return (
    <>
      <style>{STYLE}</style>
      <input ref={marksTemplateFileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleMarksTemplateUpload}/>
      <div className="shell">

        {/* ── sidebar ── */}
        <div className="sb">
          <div className="sb-logo">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div className="logo">Attend<span>X</span></div>
              <button className="dm-toggle" onClick={()=>setDarkMode(v=>!v)} title={darkMode?"Switch to light mode":"Switch to dark mode"}>
                {darkMode ? "☀ Light" : "🌙 Dark"}
              </button>
            </div>
            <div className="logo-sub">Attendance · Marks</div>
          </div>
          <div className="sb-sec"><span>Classes ({classes.length})</span></div>
          <div className="cl-list">
            {classes.length===0&&(
              <div style={{fontFamily:"var(--mono)",fontSize:"0.62rem",color:"#2e2c24",padding:"10px 6px",textAlign:"center",lineHeight:1.7}}>
                No classes yet.<br/>Add one below ↓
              </div>
            )}
            {classes.map(cls=>(
              <div key={cls.id} className={`ci${cls.id===activeId?" act":""}`} onClick={()=>setActiveId(cls.id)}>
                <div className="ci-dot" style={{background:cls.color||"#e04e20"}}/>
                <div
                  className={`ci-ring${classHealth(cls).pct===null?" empty":""}`}
                  style={{background:`conic-gradient(${classHealth(cls).tone} ${classHealth(cls).pct ?? 0}%, rgba(148,163,184,.18) 0)`}}
                  title={classHealth(cls).pct===null?"No data yet":`${classHealth(cls).pct}% health`}
                >
                  <span>{classHealth(cls).pct===null?"--":`${classHealth(cls).pct}%`}</span>
                </div>
                <div className="ci-info">
                  <div className="ci-name">{cls.info?.subject||"Untitled"}</div>
                  <div className="ci-sub">{cls.info?.course||cls.info?.teacher||"—"}</div>
                </div>
                <span className={`ci-type-badge ${cls.type==="marks"?"ci-type-marks":"ci-type-att"}`}>{cls.type==="marks"?"MARKS":"ATT"}</span>
                <button className="ci-del" onClick={e=>deleteClass(cls.id,e)}>×</button>
              </div>
            ))}
          </div>
          <div className="sb-foot">
            <div className={`save-badge ${saveStatus}`}>{saveStatus==="saving"?"⟳ saving…":"✓ all changes saved"}</div>
            <button className="add-cls-btn" onClick={()=>setShowModal(true)}>+ New Class</button>
            <button className="exp-all-btn" onClick={exportAll}>↓ Export All Classes</button>
            <button className="def-btn" onClick={exportDefaultersAll}>⚠ Generate Defaulter List</button>
          </div>
        </div>

        {/* ── main ── */}
        <div className="main">
          {!active ? (
            <div className="no-class">
              <div className="no-class-icon">📋</div>
              <div>Select a class or create a new one.</div>
              <button className="btn btn-accent" style={{marginTop:6}} onClick={()=>setShowModal(true)}>+ New Class</button>
            </div>
          ) : (<>

            {/* info bar */}
            <div className={`ibar${isMarks?" marks-mode":""}`}>
              {infoCollapsed ? (
                <div className="ibar-collapsed">
                  <span className={`mode-pill ${isMarks?"marks":"att"}`}>{isMarks?"📊 MARKS":"✅ ATTENDANCE"}</span>
                  <span className="ibar-summary">
                    {[active.info?.faculty, active.info?.course, active.info?.session, active.info?.subject, active.info?.teacher].filter(Boolean).join("  ·  ") || "No class details set"}
                  </span>
                  <button className="ibar-toggle" onClick={()=>setInfoCollapsed(false)}>▾ Show details</button>
                </div>
              ) : (
                <div className="ibar-row">
                  <span className={`mode-pill ${isMarks?"marks":"att"}`}>{isMarks?"📊 MARKS":"✅ ATTENDANCE"}</span>
                  {[["faculty","Faculty / Department","Faculty of Engineering & Computing","w1"],["session","Session","Feb 2026 – Jun 2026","w2"],["teacher","Teacher","Full name","w3"],["subject","Subject","Subject","w3"],["course","Course","BSSE 39 M 1 (B)","w3"]].map(([f,lbl,ph,w])=>(
                    <div key={f} className="ibar-f" style={{flex:w==="w1"?"2 1 180px":"1 1 120px"}}>
                      <span className="ibar-lbl">{lbl}</span>
                      <input className={`ibar-inp ${w}`} placeholder={ph} value={active.info?.[f] ?? ""} onChange={e=>setInfo(f,e.target.value)}/>
                    </div>
                  ))}
                  <button className="ibar-toggle" onClick={()=>setInfoCollapsed(true)}>▴ Hide</button>
                </div>
              )}
            </div>

            {/* topbar */}
            <div className="topbar">
              {isAtt&&<>
                <button className="btn btn-ghost" onClick={()=>{if(window.confirm("Clear all attendance?"))updActive(c=>({...c,att:{}}));}}>Clear</button>
                <button className={`btn today-toggle${todayMode?" on":""}`} onClick={toggleTodayMode}>{todayMode?"Today On":"Today Mode"}</button>
                <button className="btn btn-green" onClick={markRestP}>✓ Mark Rest Present</button>
                <button className="btn btn-ghost" onClick={markRestA}>✗ Mark Rest Absent</button>
                <button className={`btn ${lowAttOnly?"btn-amber":"btn-ghost"}`} onClick={()=>setLowAttOnly(v=>!v)}>⚠ Below 25%</button>
              </>}
              {isMarks&&<>
                <button className="btn btn-ghost" onClick={()=>{if(window.confirm("Clear all marks?"))updActive(c=>({...c,marks:{}}));}}>Clear Marks</button>
                <button className="btn btn-purple" onClick={openGradeModal}>🎓 Grading Criteria</button>
                <button className="btn btn-blue" onClick={downloadMarksTemplate}>↓ Marks Template</button>
                <button className="btn btn-teal" onClick={()=>marksTemplateFileRef.current?.click()}>↑ Upload Marks</button>
              </>}
              {classes.length>1&&<button className="btn btn-amber" onClick={openCopyModal}>⇄ Copy Roster</button>}
              <button className="btn btn-teal" onClick={()=>{setImportData(null);setShowImportModal(true);}}>↑ Import Students</button>
              <button
                className="btn btn-ghost"
                onClick={()=>{
                  if (insightsOpen) setInsightsMaximized(false);
                  setInsightsOpen(v=>!v);
                }}
              >
                {insightsOpen?"Hide Insights":"Show Insights"}
              </button>
              <div className="undo-bar">
                <button className="undo-btn" disabled={undoLen===0} onClick={undo} title="Undo (Ctrl+Z)">↩</button>
                <button className="undo-btn" disabled={redoLen===0} onClick={redo} title="Redo (Ctrl+Y)">↪</button>
              </div>
              <button className="btn btn-accent" onClick={exportOne}>↓ Export This Class</button>
            </div>

            {/* subbar */}
            <div className="subbar">
              <span className="slabel">Add Student</span>
              <input className="inp inp-r" placeholder="Student ID" value={newRoll} onChange={e=>setNewRoll(e.target.value)} onKeyDown={onAddKey}/>
              <input className="inp inp-s" placeholder="System ID" value={newSystemId} onChange={e=>setNewSystemId(e.target.value)} onKeyDown={onAddKey}/>
              <input className="inp inp-n" placeholder="Name (or Ali, Sara, Ahmed…)" value={newName} onChange={e=>setNewName(e.target.value)} onKeyDown={onAddKey}/>
              {isAtt&&<>
                <span className="slabel">Add Date</span>
                <input className="inp inp-d" type="date" value={newDate} onChange={e=>setNewDate(e.target.value)}/>
                <button className="btn btn-blue" onClick={addDate}>+ Add Column</button>
              </>}
              {isMarks&&<button className="btn btn-purple" onClick={()=>setShowAssessModal(true)}>+ Add Assessment</button>}
            </div>

            {/* search */}
            <div className="searchbar">
              <span className="search-icon">🔍</span>
              <input className="search-inp" placeholder="Search by name, student ID, or system ID…" value={searchQ} onChange={e=>setSearchQ(e.target.value)}/>
              {(searchQ || lowAttOnly)&&<>
                <span className="search-count">{filteredStudents.length}/{active.students.length} shown{lowAttOnly?` · below ${LOW_ATT_THRESHOLD}%`:""}</span>
                <button className="search-clear" onClick={()=>{setSearchQ("");setLowAttOnly(false);}}>✕ Clear</button>
              </>}
            </div>

            {/* stats */}
            <div className="statsbar">
              <div className="stat"><span className="slbl">Students:</span><span className="sv y">{active.students.length}</span></div>
              {isAtt&&<>
                <div className="stat"><span className="slbl">Dates:</span><span className="sv y">{sortedDates.length}</span></div>
                <div className="stat"><span className="slbl">Present:</span><span className="sv g">{overall.p}</span></div>
                <div className="stat"><span className="slbl">Absent:</span><span className="sv r">{overall.a}</span></div>
                {(overall.p+overall.a)>0&&<div className="stat"><span className="slbl">Overall:</span><span className="sv g">{Math.round(overall.p/(overall.p+overall.a)*100)}% present</span></div>}
                {lowAttCount>0&&<div className="stat"><span className="slbl">Below 25%:</span><span className="sv am">⚠ {lowAttCount}</span></div>}
              </>}
              {isMarks&&<>
                <div className="stat"><span className="slbl">Assessments:</span><span className="sv pu">{assessments.length}</span></div>
                <div className="stat"><span className="slbl">Total Marks:</span><span className="sv pu">{assessments.reduce((s,a)=>s+a.outOf,0)}</span></div>
              </>}
              <span className="hint">{isAtt?"P = Present · A = Absent · Arrow keys to navigate":"Enter marks · A = Absent · Click column header to edit weight"}</span>
            </div>

            {/* WEIGHT BAR — only for marks mode */}
            {todayMode && isAtt && (
              <div className="today-banner">
                <span>Today Mode: showing {fmtDate(todayIso)} only</span>
                <span>{visiblePct === null ? "No marks yet" : `${visibleOverall.p} present / ${visibleOverall.a} absent (${visiblePct}%)`}</span>
              </div>
            )}

            {isMarks && assessments.length > 0 && (
              <div className="wbar">
                <span className="wbar-label">Weights</span>
                <div className="wbar-pills">
                  {assessments.map(a => (
                    <div key={a.id} className="wpill">
                      <span className="wpill-name" title={a.name}>{a.name}</span>
                      <input
                        className="wpill-inp"
                        type="number" min="0" max="100" step="1"
                        placeholder="0"
                        value={a.weight || ""}
                        onChange={e => setAssessWeight(a.id, e.target.value)}
                        title={`Weight for ${a.name}`}
                      />
                      <span className="wpill-pct">%</span>
                    </div>
                  ))}
                </div>
                <div className={`wbar-total ${totalWeight>100?"warn":weightsOk?"ok":"warn"}`}>
                  Σ {totalWeight.toFixed(1)}% {totalWeight>100?"⚠ Over!":weightsOk?"✓":"≠ 100"}
                </div>
                {totalWeight>100&&<div className="weight-error">Weights exceed 100% — reduce to proceed</div>}
              </div>
            )}

            {/* SPREADSHEET */}
            <div className={`workspace${insightsOpen ? "" : " no-insights"}`}>
              <div className="workbench">
            <div className="sh-outer">
              <div className="sh-scroll" ref={tableScrollRef} onScroll={handleTableScroll}>
                {active.students.length===0
                  ? <div className="empty">Add students above or import from Excel to begin.</div>
                  : filteredStudents.length===0
                  ? <div className="empty">No students match "{searchQ}".</div>
                  : isAtt ? (

                  /* ATTENDANCE TABLE */
                  <table className="sh-tbl">
                    <thead>
                      <tr className="chr">
                        <th className="fr">Student ID</th>
                        <th className="fs">System ID</th>
                        <th className="fn">Student Name</th>
                        {visibleDates.map(d=>{
                          const [y,m,day]=d.split("-");
                          const dow=new Date(d).toLocaleDateString("en-US",{weekday:"short"});
                          return (
                            <th key={d} className="dc">
                              <div className="dch">
                                <span className="dch-t">{dow} {m}/{y.slice(2)}</span>
                                <span className="dch-m">{day}</span>
                                <button className="xbtn" onClick={()=>removeDate(d)}>×</button>
                              </div>
                            </th>
                          );
                        })}
                        <th className="ft">Summary</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rowVirtual.enabled && rowVirtual.topPad > 0 && (
                        <tr className="virtual-spacer"><td colSpan={attTableColSpan} style={{height:rowVirtual.topPad}} /></tr>
                      )}
                      {renderedStudents.map(s=>{
                        const sm=attendanceSummaryByStudent.get(s.id);
                        const isLow=sm&&sm.pct<LOW_ATT_THRESHOLD;
                        return (
                          <tr key={s.id} className={`dr${isLow?" low-att":""}`}>
                            <td className="tdr">
                              <span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:2,height:"34px"}}>
                                <input className="rinp" value={s.roll}
                                  onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,roll:e.target.value}:x)}))}/>
                                <button className="xrow" onClick={()=>removeStudent(s.id)}>×</button>
                              </span>
                            </td>
                            <td className="tds">
                              <input className="sinp" value={s.systemId || ""}
                                onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,systemId:e.target.value}:x)}))}/>
                            </td>
                            <td className="tdn">
                              <input className="ninp" value={s.name}
                                onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,name:e.target.value}:x)}))}/>
                              {isLow&&<span className="low-badge">⚠ LOW</span>}
                              <div className="heat-strip" title="Recent attendance history">
                                {heatForStudent(active,s.id).map((h,i)=><span key={i} className={`heat-dot ${h}`}/>)}
                              </div>
                            </td>
                            {visibleDates.map((d,dIdx)=>{
                              const v=getCell(active,s.id,d);
                              const cl=isP(v)?"ip":isA(v)?"ia":"";
                              const rk=`att_${s.id}_${d}`;
                              return (
                                <td key={d} className="tda">
                                  <input ref={el=>{refs.current[rk]=el;}}
                                    className={`ainp ${cl}`} value={v} maxLength={1}
                                    onChange={e=>setCell(s.id,d,e.target.value)}
                                    onKeyDown={e=>navKeyAtt(e,s.id,dIdx)}/>
                                </td>
                              );
                            })}
                            <td className="tdt">
                              {sm
                                ? <span>
                                    <span className="sp">{sm.p}P</span>/<span className="sa">{sm.a}A</span>{" "}
                                    <span style={{color:isLow?"#f59e0b":"#aaa",fontSize:"0.61rem",fontWeight:isLow?700:400}}>({sm.pct}%)</span>
                                    {(() => { const left = absencesUntilDefaulter(active, s.id); return left !== null && left > 0 && left <= 5 ? <span className="warn-badge">{left} left</span> : null; })()}
                                  </span>
                                : <span style={{color:"#ccc"}}>—</span>}
                            </td>
                          </tr>
                        );
                      })}
                      {rowVirtual.enabled && rowVirtual.bottomPad > 0 && (
                        <tr className="virtual-spacer"><td colSpan={attTableColSpan} style={{height:rowVirtual.bottomPad}} /></tr>
                      )}
                    </tbody>
                    <tfoot>
                      <tr className="sr">
                        <td className="sfr">Σ</td>
                        <td className="sfs"></td>
                        <td className="sfn">{active.students.length} students / {visibleDates.length} shown</td>
                        {visibleDates.map(d=>{const{p,a}=dateSummaryByDate.get(d) ?? {p:0,a:0};return <td key={d}><span className="sp">{p}P</span> <span className="sa">{a}A</span></td>;})}
                        <td className="sft"><span className="sp">{visibleOverall.p}P</span>/<span className="sa">{visibleOverall.a}A</span></td>
                      </tr>
                    </tfoot>
                  </table>

                ) : assessments.length===0 ? (
                  <div className="empty">Click <strong>+ Add Assessment</strong> above to add quizzes, assignments, or exams.</div>
                ) : (

                  /* MARKS TABLE */
                  <table className="sh-tbl">
                    <thead>
                      <tr className="chr">
                        <th className="fr">Student ID</th>
                        <th className="fs">System ID</th>
                        <th className="fn marks-mode">Student Name</th>
                        {assessments.map(a=>(
                          <th key={a.id} className="mc" style={{cursor:"pointer"}} onClick={()=>openEditAssess(a.id)} title="Click to edit this assessment">
                            <div className="dch">
                              <span className="dch-t" style={{color:"#a78bfa"}}>{a.type.toUpperCase()}</span>
                              <span className="dch-m">{a.name}</span>
                              <span className="dch-sub">/{a.outOf}</span>
                              {a.weight>0 && <span className="dch-w">{a.weight}%</span>}
                              <button className="xbtn" onClick={e=>{e.stopPropagation();removeAssessment(a.id);}}>×</button>
                            </div>
                          </th>
                        ))}
                        <th className="fw">Weighted %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rowVirtual.enabled && rowVirtual.topPad > 0 && (
                        <tr className="virtual-spacer"><td colSpan={marksTableColSpan} style={{height:rowVirtual.topPad}} /></tr>
                      )}
                      {renderedStudents.map(s=>{
                        const wm=weightedSummaryByStudent.get(s.id);
                        return (
                          <tr key={s.id} className="dr">
                            <td className="tdr">
                              <span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:2,height:"34px"}}>
                                <input className="rinp" value={s.roll}
                                  onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,roll:e.target.value}:x)}))}/>
                                <button className="xrow" onClick={()=>removeStudent(s.id)}>×</button>
                              </span>
                            </td>
                            <td className="tds">
                              <input className="sinp" value={s.systemId || ""}
                                onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,systemId:e.target.value}:x)}))}/>
                            </td>
                            <td className="tdn marks-mode">
                              <input className="ninp" value={s.name}
                                onChange={e=>updActive(c=>({...c,students:c.students.map(x=>x.id===s.id?{...x,name:e.target.value}:x)}))}/>
                            </td>
                            {assessments.map((a,aIdx)=>{
                              const v=getMark(active,s.id,a.id);
                              const cl=markCls(active,s.id,a.id);
                              const rk=`mark_${s.id}_${a.id}`;
                              return (
                                <td key={a.id} className="tmc">
                                  <input ref={el=>{refs.current[rk]=el;}}
                                    className={`minp ${cl}`} value={v} placeholder="—"
                                    onChange={e=>setMark(s.id,a.id,e.target.value)}
                                    onKeyDown={e=>navKeyMarks(e,s.id,aIdx)}
                                    title={`${a.name} / ${a.outOf}`}/>
                                </td>
                              );
                            })}
                            <td className="tdw">
                              {wm
                                ? <span style={{color: wm.pct>=90?"#14633a":wm.pct>=60?"#1d4ed8":"#be2c0a", fontWeight:900}}>
                                    {wm.pct}%
                                    {(() => { const g = getGrade(wm.pct); return g ? <span className={`grade-badge ${gradeColorClass(g)}`}>{g}</span> : null; })()}
                                  </span>
                                : <span style={{color:"#ccc"}}>—</span>}
                            </td>
                          </tr>
                        );
                      })}
                      {rowVirtual.enabled && rowVirtual.bottomPad > 0 && (
                        <tr className="virtual-spacer"><td colSpan={marksTableColSpan} style={{height:rowVirtual.bottomPad}} /></tr>
                      )}
                    </tbody>
                    <tfoot>
                      <tr className="sr">
                        <td className="sfr">Avg</td>
                        <td className="sfs"></td>
                        <td className={`sfn marks-mode`}>{active.students.length} students · {assessments.length} assessments</td>
                        {assessments.map(a=>{
                          const avg=assessAvgByAssessment.get(a.id);
                          return <td key={a.id} style={{color:"#a78bfa"}}>{avg?`${avg}/${a.outOf}`:"—"}</td>;
                        })}
                        <td className="sfw">{activeWeightedAverage === null ? "Weighted" : `${activeWeightedAverage}% avg`}</td>
                      </tr>
                    </tfoot>
                  </table>
                )}
              </div>
            </div>
              </div>

              {insightsOpen && (
                <aside className={`insight-drawer${insightsMaximized ? " max" : ""}`}>
                  <div className="insight-head">
                    <div>
                      <div className="insight-title">Class Insights</div>
                      <div className="insight-kicker">{isAtt ? "Attendance health" : "Marks overview"}</div>
                    </div>
                    <div className="insight-actions">
                      <button className="insight-hide" onClick={exportInsights}>Export Insights</button>
                      <button
                        className="insight-hide"
                        onClick={()=>setInsightsMaximized(v=>!v)}
                        title={insightsMaximized ? "Minimize insights" : "Maximize insights"}
                      >
                        {insightsMaximized ? "Minimize" : "Maximize"}
                      </button>
                      <button className="insight-hide" onClick={()=>{setInsightsMaximized(false);setInsightsOpen(false);}}>Hide</button>
                    </div>
                  </div>

                  {isAtt ? (
                    <>
                      <div className="metric-grid">
                        <div className="metric-card">
                          <div className="metric-value">{overallPct === null ? "--" : `${overallPct}%`}</div>
                          <div className="metric-label">Overall present</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{attendanceDefaulterStudents.length}</div>
                          <div className="metric-label">Below 75%</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{visibleDates.length}</div>
                          <div className="metric-label">{todayMode ? "Today view" : "Dates shown"}</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{overall.p + overall.a}</div>
                          <div className="metric-label">Marked cells</div>
                        </div>
                      </div>

                      <div className="insight-card">
                        <div className="insight-card-title">Students needing attention</div>
                        <div className="risk-list">
                          {topRisks.length ? topRisks.map(s => {
                            const sm = attendanceSummaryByStudent.get(s.id);
                            return (
                              <div key={s.id} className="risk-item">
                                <span className="risk-name">{s.name}</span>
                                <span className="risk-pct">{sm?.pct ?? 0}%</span>
                              </div>
                            );
                          }) : <div className="insight-empty">No critical attendance yet.</div>}
                        </div>
                      </div>

                      <div className="insight-card">
                        <div className="insight-card-title">Recent dates</div>
                        <div className="date-chip-row">
                          {recentDates.length ? recentDates.map(d => <span key={d} className="date-chip">{fmtDate(d)}</span>) : <span className="date-chip">No dates</span>}
                        </div>
                      </div>

                      <div className="insight-card">
                        <div className="insight-card-title">Below 75% attendance</div>
                        {attendanceDefaulterStudents.length ? (
                          <div className="insight-student-list">
                            <div className="insight-list-head">
                              <span>Roll No</span>
                              <span>Name</span>
                              <span>Attendance</span>
                            </div>
                            {attendanceDefaulterStudents.map(({ student, summary }) => (
                              <div key={student.id} className="insight-student-row">
                                <span className="insight-student-roll">{student.roll || "N/A"}</span>
                                <span className="insight-student-name">{student.name}</span>
                                <span className="insight-student-score att">{summary.pct}%</span>
                              </div>
                            ))}
                          </div>
                        ) : <div className="insight-empty">No students below 75%.</div>}
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="metric-grid">
                        <div className="metric-card">
                          <div className="metric-value">{assessments.length}</div>
                          <div className="metric-label">Assessments</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{lowWeightedStudents.length}</div>
                          <div className="metric-label">Below 50</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{totalWeight.toFixed(0)}%</div>
                          <div className="metric-label">Weight total</div>
                        </div>
                        <div className="metric-card">
                          <div className="metric-value">{activeWeightedAverage === null ? "--" : `${activeWeightedAverage}%`}</div>
                          <div className="metric-label">Weighted average</div>
                        </div>
                      </div>
                      <div className="insight-card">
                        <div className="insight-card-title">Weighted score below 50</div>
                        {lowWeightedStudents.length ? (
                          <div className="insight-student-list">
                            <div className="insight-list-head">
                              <span>Roll No</span>
                              <span>Name</span>
                              <span>Weighted Score</span>
                            </div>
                            {lowWeightedStudents.map(({ student, weighted }) => (
                              <div key={student.id} className="insight-student-row">
                                <span className="insight-student-roll">{student.roll || "N/A"}</span>
                                <span className="insight-student-name">{student.name}</span>
                                <span className="insight-student-score marks">{weighted.pct}%</span>
                              </div>
                            ))}
                          </div>
                        ) : <div className="insight-empty">{hasWeightedAssessments ? "No weighted scores below 50." : "No weighted scores yet."}</div>}
                      </div>
                    </>
                  )}
                </aside>
              )}
            </div>
          </>)}
        </div>
      </div>

      {/* ── New Class modal ── */}
      {showModal&&(
        <div className="mbg" onClick={()=>setShowModal(false)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <h2>New Class</h2>
            <div className="modal-sub">Choose type, then fill in details</div>
            <div className="type-pick">
              <div className={`type-opt${draft.type==="attendance"?" sel-att":""}`} onClick={()=>setDraft(d=>({...d,type:"attendance"}))}>
                <div className="type-opt-icon">✅</div>
                <div className="type-opt-label">Attendance</div>
                <div className="type-opt-desc">Track P/A by date</div>
              </div>
              <div className={`type-opt${draft.type==="marks"?" sel-marks":""}`} onClick={()=>setDraft(d=>({...d,type:"marks"}))}>
                <div className="type-opt-icon">📊</div>
                <div className="type-opt-label">Marks</div>
                <div className="type-opt-desc">Quizzes · Assignments · Exams</div>
              </div>
            </div>
            {[["faculty","Faculty / Department","Faculty of Engineering & Computing [Dept. of SE]"],["session","Session","February 2026 – June 2026"],["teacher","Teacher's Name","Muhammad Bilal Rehman"],["subject","Subject","Programming Fundamentals"],["course","Course","BSSE 39 M 1 (B)"]].map(([f,lbl,ph])=>(
              <div key={f} className="mf">
                <label className="ml">{lbl}</label>
                <input className="mi" placeholder={ph} value={draft[f]} onChange={e=>setDraft(d=>({...d,[f]:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")createClass();}}/>
              </div>
            ))}
            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowModal(false)}>Cancel</button>
              <button className={draft.type==="marks"?"mbtn-purple":"mbtn-ok"} onClick={createClass}>Create Class →</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Import from Excel modal ── */}
      {showImportModal&&(
        <div className="mbg" onClick={()=>setShowImportModal(false)}>
          <div className="modal wide" onClick={e=>e.stopPropagation()}>
            <h2>Import Students from Excel</h2>
            <div className="modal-sub">
              Upload an .xlsx or .csv file with Student ID + System ID + Name columns.<br/>
              Headers are auto-detected. Duplicates will be skipped.
            </div>

            {/* hidden file input */}
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
              onChange={e=>{ if(e.target.files[0]) parseImportFile(e.target.files[0]); e.target.value=""; }}/>

            {/* Drop zone */}
            <div
              className={`drop-zone${importDragOver?" drag-over":""}`}
              onClick={()=>fileRef.current?.click()}
              onDragOver={e=>{e.preventDefault();setImportDragOver(true);}}
              onDragLeave={()=>setImportDragOver(false)}
              onDrop={handleImportDrop}
            >
              <div className="drop-zone-icon">📂</div>
              <div className="drop-zone-text">
                {importData?.fileName
                  ? <>File loaded: <strong style={{color:"#a78bfa"}}>{importData.fileName}</strong></>
                  : <>Click to choose file or drag & drop here</>
                }
              </div>
              <div className="drop-zone-hint">Supported: .xlsx · .xls · .csv</div>
            </div>

            {/* Status */}
            {importData && (
              <>
                {importData.error
                  ? <div className="import-status err">⚠ {importData.error}</div>
                  : importData.rows.length === 0
                  ? <div className="import-status warn">No valid rows found. Make sure the file has Student ID, System ID, and Name columns.</div>
                  : <div className="import-status ok">✓ Found {importData.rows.length} student{importData.rows.length!==1?"s":""} ready to import</div>
                }
                {importData.warnings?.length > 0 && (
                  <div className="import-status warn" style={{marginTop:6,textAlign:"left"}}>
                    {importData.warnings.slice(0,5).map((w,i)=><div key={i}>⚠ {w}</div>)}
                    {importData.warnings.length>5&&<div>…and {importData.warnings.length-5} more</div>}
                  </div>
                )}
              </>
            )}

            {/* Preview */}
            {importData?.rows?.length > 0 && (
              <div className="import-preview">
                <div className="import-preview-row" style={{borderBottom:"1px solid #2a2820",marginBottom:4}}>
                  <span className="ipr-roll" style={{color:"#484440"}}>STUDENT ID</span>
                  <span className="ipr-roll" style={{color:"#484440"}}>SYSTEM ID</span>
                  <span className="ipr-name" style={{color:"#484440"}}>NAME</span>
                </div>
                {importData.rows.slice(0,12).map((r,i)=>(
                  <div key={i} className="import-preview-row">
                    <span className="ipr-roll">{r.roll}</span>
                    <span className="ipr-roll">{r.systemId || "—"}</span>
                    <span className="ipr-name">{r.name}</span>
                  </div>
                ))}
                {importData.rows.length>12&&(
                  <div style={{fontFamily:"var(--mono)",fontSize:"0.62rem",color:"#484440",padding:"4px 0",textAlign:"center"}}>
                    …and {importData.rows.length-12} more
                  </div>
                )}
              </div>
            )}

            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowImportModal(false)}>Cancel</button>
              {importData?.rows?.length>0 && (
                <button className="mbtn-green" onClick={confirmImport}>
                  ↑ Import {importData.rows.length} Students →
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Add Assessment modal ── */}
      {showAssessModal&&(
        <div className="mbg" onClick={()=>setShowAssessModal(false)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <h2>Add Assessment</h2>
            <div className="modal-sub">Add a quiz, assignment, or exam column</div>
            <div className="mf">
              <label className="ml">Type</label>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {["quiz","assignment","midterm","final","lab","other"].map(t=>(
                  <button key={t} onClick={()=>setAssessDraft(d=>({...d,type:t}))}
                    style={{fontFamily:"var(--mono)",fontSize:"0.68rem",padding:"5px 12px",borderRadius:"5px",
                      border:`1px solid ${assessDraft.type===t?"#a78bfa":"#2a2820"}`,
                      background:assessDraft.type===t?"#2e1a5e":"transparent",
                      color:assessDraft.type===t?"#a78bfa":"#666",cursor:"pointer",transition:"all 0.13s"}}>
                    {t}
                  </button>
                ))}
              </div>
            </div>
            <div className="mrow">
              <div className="mf" style={{flex:2}}>
                <label className="ml">Name</label>
                <input className="mi" placeholder="e.g. Quiz 1 / Assignment 2 / Midterm" value={assessDraft.name}
                  onChange={e=>setAssessDraft(d=>({...d,name:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")addAssessment();}}/>
              </div>
              <div className="mf" style={{flex:1}}>
                <label className="ml">Out Of</label>
                <input className="mi mi-sm" type="number" min="1" max="1000" placeholder="10" value={assessDraft.outOf}
                  onChange={e=>setAssessDraft(d=>({...d,outOf:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")addAssessment();}}/>
              </div>
              <div className="mf" style={{flex:1}}>
                <label className="ml">Weight %</label>
                <input className="mi mi-sm" type="number" min="0" max="100" placeholder="e.g. 15" value={assessDraft.weight}
                  onChange={e=>setAssessDraft(d=>({...d,weight:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")addAssessment();}}/>
              </div>
            </div>
            <div style={{fontFamily:"var(--mono)",fontSize:"0.6rem",color:"#484440",marginTop:-4,marginBottom:8,lineHeight:1.7}}>
              Weight % is optional. Set it now or adjust later via the Weights bar. All weights should add up to 100%.
            </div>
            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowAssessModal(false)}>Cancel</button>
              <button className="mbtn-purple" onClick={addAssessment}>+ Add Column →</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Edit Assessment modal ── */}
      {showEditAssess&&(
        <div className="mbg" onClick={()=>setShowEditAssess(false)}>
          <div className="modal narrow" onClick={e=>e.stopPropagation()}>
            <h2>Edit Assessment</h2>
            <div className="modal-sub">Update name, marks, type, or weight</div>
            <div className="mf">
              <label className="ml">Type</label>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {["quiz","assignment","midterm","final","lab","other"].map(t=>(
                  <button key={t} onClick={()=>setAssessDraft(d=>({...d,type:t}))}
                    style={{fontFamily:"var(--mono)",fontSize:"0.68rem",padding:"5px 12px",borderRadius:"5px",
                      border:`1px solid ${assessDraft.type===t?"#a78bfa":"#2a2820"}`,
                      background:assessDraft.type===t?"#2e1a5e":"transparent",
                      color:assessDraft.type===t?"#a78bfa":"#666",cursor:"pointer",transition:"all 0.13s"}}>
                    {t}
                  </button>
                ))}
              </div>
            </div>
            <div className="mrow">
              <div className="mf" style={{flex:2}}>
                <label className="ml">Name</label>
                <input className="mi" value={assessDraft.name} onChange={e=>setAssessDraft(d=>({...d,name:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")saveEditAssess();}}/>
              </div>
              <div className="mf" style={{flex:1}}>
                <label className="ml">Out Of</label>
                <input className="mi mi-sm" type="number" min="1" max="1000" value={assessDraft.outOf} onChange={e=>setAssessDraft(d=>({...d,outOf:e.target.value}))} onKeyDown={e=>{if(e.key==="Enter")saveEditAssess();}}/>
              </div>
            </div>
            <div className="mf">
              <label className="ml">Weight %</label>
              <input className="mi mi-sm" type="number" min="0" max="100" placeholder="e.g. 25" value={assessDraft.weight}
                onChange={e=>setAssessDraft(d=>({...d,weight:e.target.value}))}
                onKeyDown={e=>{if(e.key==="Enter")saveEditAssess();}}/>
            </div>
            <div style={{fontFamily:"var(--mono)",fontSize:"0.6rem",color:"#484440",lineHeight:1.7,marginBottom:4}}>
              You can also update weights directly in the Weights bar below the stats.
            </div>
            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowEditAssess(false)}>Cancel</button>
              <button className="mbtn-purple" onClick={saveEditAssess}>Save Changes →</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Copy Roster modal ── */}
      {showCopyModal&&(
        <div className="mbg" onClick={()=>setShowCopyModal(false)}>
          <div className="modal wide" onClick={e=>e.stopPropagation()}>
            <h2>Copy Roster From…</h2>
            <div className="modal-sub">
              Copy student list into <strong style={{color:"#aaa"}}>{active?.info?.subject||"this class"}</strong>.
              Duplicates (same roll or same name) will be skipped.
            </div>
            <div className="cls-pick-list">
              {classes.filter(c=>c.id!==activeId).map(cls=>(
                <div key={cls.id} className={`cls-pick-item${copySourceId===cls.id?" selected":""}`} onClick={()=>setCopySourceId(cls.id)}>
                  <div className="cls-pick-dot" style={{background:cls.color||"#e04e20"}}/>
                  <div className="cls-pick-info">
                    <div className="cls-pick-name">{cls.info?.subject||"Untitled"}</div>
                    <div className="cls-pick-sub">{cls.info?.course||"—"} · {cls.students?.length??0} students · <span style={{color:cls.type==="marks"?"#818cf8":"#6fcf97",marginLeft:4}}>{cls.type==="marks"?"MARKS":"ATT"}</span></div>
                  </div>
                  {copySourceId===cls.id&&<span style={{color:"var(--accent2)",fontFamily:"var(--mono)",fontSize:"0.7rem"}}>✓</span>}
                </div>
              ))}
            </div>
            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowCopyModal(false)}>Cancel</button>
              <button className="mbtn-blue" onClick={confirmCopyRoster} disabled={!copySourceId} style={{opacity:copySourceId?1:0.4}}>Copy Roster →</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Grading Criteria modal ── */}
      {showGradeModal&&(
        <div className="mbg" onClick={()=>setShowGradeModal(false)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <h2>Grading Criteria</h2>
            <div className="modal-sub">
              Define grade labels and minimum weighted percentage required for each.
              Grades are matched top-down — a student gets the highest grade whose minimum they meet.
            </div>
            <div style={{display:"flex",gap:10,padding:"8px 0 6px",fontFamily:"var(--mono)",fontSize:".64rem",color:"var(--muted)",textTransform:"uppercase",letterSpacing:".1em",fontWeight:800}}>
              <span style={{width:60,textAlign:"center"}}>Grade</span>
              <span style={{width:70,textAlign:"center"}}>Min %</span>
              <span style={{flex:1}}>Preview</span>
            </div>
            <div style={{maxHeight:320,overflowY:"auto"}}>
              {gradeDraft.map((g, i) => (
                <div key={i} className="grade-row">
                  <input className="grade-lbl-inp" value={g.label} placeholder="A+"
                    onChange={e => { const d=[...gradeDraft]; d[i]={...d[i],label:e.target.value}; setGradeDraft(d); }}/>
                  <input className="grade-min-inp" type="number" min="0" max="100" value={g.min}
                    onChange={e => { const d=[...gradeDraft]; d[i]={...d[i],min:parseFloat(e.target.value)||0}; setGradeDraft(d); }}/>
                  <span style={{fontFamily:"var(--mono)",fontSize:".68rem",color:"var(--muted)",flex:1}}>
                    {g.min}% — {i===0?"100%":gradeDraft.filter(x=>x.min>g.min).length?`${Math.min(...gradeDraft.filter(x=>x.min>g.min).map(x=>x.min))-1}%`:"100%"}
                  </span>
                  <span className={`grade-badge ${gradeColorClass(g.label)}`} style={{fontSize:".72rem"}}>{g.label||"?"}</span>
                  <button className="grade-del" onClick={() => { const d=[...gradeDraft]; d.splice(i,1); setGradeDraft(d); }} title="Remove grade">×</button>
                </div>
              ))}
            </div>
            <button className="grade-add-btn" onClick={() => setGradeDraft(d => [...d, { label: "", min: 0 }])}>+ Add Grade</button>
            <div className="mact">
              <button className="mbtn-c" onClick={()=>setShowGradeModal(false)}>Cancel</button>
              <button className="mbtn-c" onClick={()=>setGradeDraft(DEFAULT_GRADES.map(g=>({...g})))}>Reset to Default</button>
              <button className="mbtn-purple" onClick={saveGrades}>Save Grades →</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
