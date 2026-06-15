import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── utils ────────────────────────────────────────────────────────────────────
const todayStr = () => new Date().toISOString().split("T")[0];
const fmtDate  = (iso) => { const [y,m,d]=iso.split("-"); return `${d}/${m}/${String(y).slice(2)}`; };
const isP = (v) => v?.trim().toLowerCase() === "p";
const isA = (v) => v?.trim().toLowerCase() === "a";
const norm = (v) => { if(isP(v)) return "P"; if(isA(v)) return "A"; return v??""; };
const uid  = () => Math.random().toString(36).slice(2,9);
const clamp = (v,mn,mx) => Math.min(mx, Math.max(mn, v));

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
const MIN_ROLL_W=86, MAX_ROLL_W=140;
const MIN_SYS_W=90, MAX_SYS_W=150;
const MIN_NAME_W=160, MAX_NAME_W=320;
const CLASS_COLORS=["#e04e20","#2563eb","#16a34a","#9333ea","#ea580c","#0891b2","#be185d","#ca8a04"];
const STORAGE_KEY = "attendx_v3";
const LOW_ATT_THRESHOLD = 25;
const DEFAULTER_THRESHOLD = 75;
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
const makeStyle = (rollW, sysW, nameW) => `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body,#root{height:100%;}
html{font-size:15px;}
body{
  min-height:100%;
  background:var(--bg);
  color:var(--text);
  font-family:'Inter',system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  -webkit-font-smoothing:antialiased;
  text-rendering:geometricPrecision;
}
:root{
  --bg:#eef2f7;
  --surface:#ffffff;
  --surface-2:#f8fafc;
  --surface-3:#f1f5f9;
  --border:#e2e8f0;
  --border-s:#cbd5e1;
  --text:#0f172a;
  --muted:#64748b;
  --muted-2:#94a3b8;
  --hdr:#ffffff;
  --hdr2:#f8fafc;
  --fg:#0f172a;
  --accent:#4f46e5;
  --accent2:#0ea5e9;
  --accent-soft:#eef2ff;
  --gold:#f59e0b;
  --purple:#7c3aed;
  --green:#16a34a;
  --red:#dc2626;
  --orange:#f97316;
  --p-bg:#dcfce7;
  --p-fg:#166534;
  --a-bg:#fee2e2;
  --a-fg:#991b1b;
  --frozen:#f8fafc;
  --row-alt:#fbfdff;
  --mono:'JetBrains Mono','IBM Plex Mono',ui-monospace,SFMono-Regular,monospace;
  --sans:'Inter','IBM Plex Sans',system-ui,sans-serif;
  --disp:'Inter','Syne',system-ui,sans-serif;
  --sw:292px;
  --roll-w:${rollW}px;
  --sys-w:${sysW}px;
  --name-w:${nameW}px;
  --shadow-sm:0 1px 2px rgba(15,23,42,.06);
  --shadow:0 12px 32px rgba(15,23,42,.10);
  --shadow-lg:0 24px 64px rgba(15,23,42,.22);
  --radius:18px;
}
/* ── dark mode ── */
.dark{
  --bg:#0b0f1a;
  --surface:#141926;
  --surface-2:#1a2030;
  --surface-3:#1e2538;
  --border:#2a3146;
  --border-s:#3b4563;
  --text:#e2e8f0;
  --muted:#8892a8;
  --muted-2:#64748b;
  --hdr:#141926;
  --hdr2:#1a2030;
  --fg:#e2e8f0;
  --accent:#818cf8;
  --accent2:#38bdf8;
  --accent-soft:rgba(99,102,241,.15);
  --gold:#fbbf24;
  --purple:#a78bfa;
  --green:#4ade80;
  --red:#f87171;
  --orange:#fb923c;
  --p-bg:rgba(34,197,94,.18);
  --p-fg:#86efac;
  --a-bg:rgba(239,68,68,.18);
  --a-fg:#fca5a5;
  --frozen:#141926;
  --row-alt:#161d2c;
  --shadow-sm:0 1px 4px rgba(0,0,0,.3);
  --shadow:0 12px 32px rgba(0,0,0,.4);
  --shadow-lg:0 24px 64px rgba(0,0,0,.55);
}
.dark .shell{background:linear-gradient(135deg,#0b0f1a 0%,#111827 42%,#0f172a 100%);}
.dark .ibar{background:rgba(20,25,38,.88);border-color:#2a3146;}
.dark .topbar{background:rgba(20,25,38,.78);border-color:#2a3146;}
.dark .subbar{background:rgba(20,25,38,.78);border-color:#2a3146;}
.dark .searchbar{background:rgba(20,25,38,.94);border-color:#2a3146;}
.dark .search-inp{background:#1a2030;border-color:#2a3146;color:#e2e8f0;}
.dark .search-clear{background:#1a2030;border-color:#2a3146;color:#94a3b8;}
.dark .ibar-inp{background:#1a2030;border-color:#2a3146;color:#e2e8f0;}
.dark .ibar-toggle{background:#1a2030;border-color:#2a3146;color:#94a3b8;}
.dark .inp{background:#1a2030;border-color:#2a3146;color:#e2e8f0;}
.dark .btn-ghost{background:#1a2030;border-color:#2a3146;color:#cbd5e1;}
.dark .btn-ghost:hover{background:#1e2538;border-color:#3b4563;color:#e2e8f0;}
.dark .sh-outer{background:#141926;border-color:#2a3146;}
.dark .sh-scroll{background:#141926;}
.dark .sh-scroll::-webkit-scrollbar-track{background:#141926;}
.dark .sh-scroll::-webkit-scrollbar-thumb{background:#3b4563;border-color:#141926;}
.dark .dr td{border-color:#1e2538;background:#141926;}
.dark .dr:nth-child(even) td{background:#161d2c;}
.dark .dr:nth-child(odd) td{background:#141926;}
.dark .dr:hover td{background:#1e2840!important;}
.dark .tdr{background:#111827!important;color:#94a3b8;border-color:#2a3146!important;}
.dark .dr:hover .tdr{background:#1a2542!important;}
.dark .tds{background:#111827!important;border-color:#2a3146!important;}
.dark .dr:hover .tds{background:#1a2542!important;}
.dark .sinp{color:#e2e8f0;}
.dark .sinp:focus{background:#1a2030;box-shadow:inset 0 0 0 2px #818cf8;}
.dark .tdn{background:#111827!important;border-color:#2a3146!important;}
.dark .dr:hover .tdn{background:#1a2542!important;}
.dark .ninp{color:#e2e8f0;}
.dark .ninp:focus{background:#1a2030;box-shadow:inset 0 0 0 2px #818cf8;}
.dark .tdt{background:#1a1a2e!important;border-color:#2a2040!important;color:#fbbf24;}
.dark .dr:hover .tdt{background:#201e3a!important;}
.dark .tdw{background:#1a1530!important;border-color:#2a2050!important;color:#a78bfa;}
.dark .ainp{color:#e2e8f0;}
.dark .ainp.ip{background:rgba(34,197,94,.18)!important;color:#86efac;}
.dark .ainp.ia{background:rgba(239,68,68,.18)!important;color:#fca5a5;}
.dark .ainp:focus{background:#1a2030!important;box-shadow:inset 0 0 0 2px #38bdf8;}
.dark .minp{color:#e2e8f0;}
.dark .minp:focus{background:#1a2030!important;box-shadow:inset 0 0 0 2px #a78bfa;}
.dark .minp.full{color:#86efac;background:rgba(34,197,94,.18)!important;}
.dark .minp.good{color:#93c5fd;background:rgba(59,130,246,.15)!important;}
.dark .minp.low-m{color:#fca5a5;background:rgba(239,68,68,.15)!important;}
.dark .minp.absent-m{color:#64748b;background:rgba(100,116,139,.12)!important;}
.dark .empty{background:#141926;color:#64748b;}
.dark .no-class{background:rgba(20,25,38,.75);border-color:#2a3146;color:#64748b;}
.dark .mbg{background:rgba(0,0,0,.7);}
.dark .modal{background:#1a2030;border-color:#2a3146;color:#e2e8f0;}
.dark .modal h2{color:#f1f5f9;}
.dark .modal-sub{color:#8892a8;}
.dark .mi{background:#111827;border-color:#2a3146;color:#e2e8f0;}
.dark .mi:focus{border-color:#818cf8;}
.dark .mbtn-c{background:#111827;border-color:#2a3146;color:#94a3b8;}
.dark .mbtn-c:hover{background:#1e2538;color:#e2e8f0;}
.dark .type-opt{background:#111827;border-color:#2a3146;}
.dark .type-opt:hover{background:#1a2030;border-color:#3b4563;}
.dark .type-opt-label{color:#e2e8f0;}
.dark .type-opt-desc{color:#64748b;}
.dark .drop-zone{background:#111827;border-color:#3b4563;}
.dark .drop-zone:hover,.dark .drop-zone.drag-over{background:rgba(99,102,241,.1);border-color:#818cf8;}
.dark .drop-zone-text{color:#cbd5e1;}
.dark .import-preview{background:#111827;border-color:#2a3146;}
.dark .import-preview-row{border-color:#2a3146;}
.dark .ipr-name{color:#e2e8f0;}
.dark .cls-pick-item{background:#111827;border-color:#2a3146;}
.dark .cls-pick-item:hover{background:#1a2030;border-color:#3b4563;}
.dark .cls-pick-item.selected{background:rgba(99,102,241,.12);border-color:#818cf8;}
.dark .cls-pick-name{color:#f1f5f9;}
.dark .cls-pick-sub{color:#64748b;}
.dark .dr.low-att td{background:rgba(245,158,11,.1)!important;}
.dark .dr.low-att:hover td{background:rgba(245,158,11,.16)!important;}
.dark .dr.low-att .tdr,.dark .dr.low-att .tds{background:rgba(245,158,11,.12)!important;color:#fbbf24!important;}
.dark .dr.low-att .tdn{background:rgba(245,158,11,.12)!important;border-right-color:#f59e0b!important;}
.dark .dr.low-att .tdt{background:rgba(245,158,11,.08)!important;}
/* dark mode toggle */
.dm-toggle{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.06);border:1px solid rgba(148,163,184,.18);color:#94a3b8;cursor:pointer;font-family:var(--mono);font-size:.68rem;font-weight:800;padding:7px 11px;border-radius:999px;transition:all .15s;letter-spacing:.04em;}
.dm-toggle:hover{background:rgba(255,255,255,.1);color:#e2e8f0;border-color:rgba(148,163,184,.32);}
/* undo/redo */
.undo-bar{display:flex;align-items:center;gap:4px;margin-left:auto;}
.undo-btn{background:#fff;border:1px solid var(--border);color:#475569;cursor:pointer;font-size:.88rem;width:30px;height:30px;border-radius:10px;display:flex;align-items:center;justify-content:center;transition:all .15s;padding:0;box-shadow:var(--shadow-sm);}
.undo-btn:hover:not(:disabled){border-color:#818cf8;color:#4f46e5;background:#eef2ff;transform:translateY(-1px);}
.undo-btn:disabled{opacity:.3;cursor:not-allowed;}
.dark .undo-btn{background:#1a2030;border-color:#2a3146;color:#94a3b8;}
.dark .undo-btn:hover:not(:disabled){background:#1e2538;border-color:#818cf8;color:#818cf8;}
/* warning badge */
.warn-badge{display:inline-flex;align-items:center;gap:4px;background:linear-gradient(135deg,#fbbf24,#f59e0b);color:#78350f;font-family:var(--mono);font-size:.5rem;font-weight:900;padding:2px 6px;border-radius:999px;margin-left:6px;vertical-align:middle;letter-spacing:.03em;white-space:nowrap;}
.dark .warn-badge{background:linear-gradient(135deg,#f59e0b,#d97706);color:#fffbeb;}
/* grade badge */
.grade-badge{display:inline-flex;align-items:center;font-family:var(--mono);font-size:.58rem;font-weight:900;padding:2px 7px;border-radius:999px;margin-left:6px;letter-spacing:.03em;}
.grade-a{background:#dcfce7;color:#166534;}.dark .grade-a{background:rgba(34,197,94,.18);color:#86efac;}
.grade-b{background:#dbeafe;color:#1e40af;}.dark .grade-b{background:rgba(59,130,246,.15);color:#93c5fd;}
.grade-c{background:#fef3c7;color:#92400e;}.dark .grade-c{background:rgba(245,158,11,.15);color:#fde68a;}
.grade-d{background:#fed7aa;color:#9a3412;}.dark .grade-d{background:rgba(249,115,22,.15);color:#fdba74;}
.grade-f{background:#fee2e2;color:#991b1b;}.dark .grade-f{background:rgba(239,68,68,.15);color:#fca5a5;}
/* grading modal */
.grade-row{display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid var(--border);}
.grade-row:last-child{border-bottom:none;}
.grade-lbl-inp{width:60px;font-family:var(--mono);font-size:.82rem;font-weight:800;text-align:center;background:#fff;border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:10px;outline:none;}
.grade-lbl-inp:focus{border-color:#818cf8;box-shadow:0 0 0 3px rgba(99,102,241,.14);}
.dark .grade-lbl-inp{background:#111827;border-color:#2a3146;color:#e2e8f0;}
.grade-min-inp{width:70px;font-family:var(--mono);font-size:.82rem;text-align:center;background:#fff;border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:10px;outline:none;}
.grade-min-inp:focus{border-color:#818cf8;box-shadow:0 0 0 3px rgba(99,102,241,.14);}
.dark .grade-min-inp{background:#111827;border-color:#2a3146;color:#e2e8f0;}
.grade-del{background:none;border:none;color:#dc2626;cursor:pointer;font-size:.9rem;padding:2px 6px;border-radius:999px;}
.grade-del:hover{background:rgba(239,68,68,.1);}
.grade-add-btn{font-family:var(--mono);font-size:.72rem;font-weight:800;background:transparent;border:1px dashed var(--border);color:var(--muted);padding:8px 14px;border-radius:12px;cursor:pointer;transition:all .15s;margin-top:8px;width:100%;}
.grade-add-btn:hover{border-color:#818cf8;color:var(--accent);background:var(--accent-soft);}
/* weight error */
.weight-error{font-family:var(--mono);font-size:.58rem;color:#f87171;font-weight:800;padding:3px 9px;background:rgba(239,68,68,.08);border:1px solid rgba(248,113,113,.22);border-radius:999px;white-space:nowrap;}
.dark .weight-error{background:rgba(239,68,68,.12);border-color:rgba(248,113,113,.18);}
.shell{
  display:flex;
  height:100vh;
  overflow:hidden;
  background:
    radial-gradient(circle at top left,rgba(79,70,229,.14),transparent 34%),
    linear-gradient(135deg,#eef2ff 0%,#f8fafc 42%,#ecfeff 100%);
}
button,input{font:inherit;}
button:disabled{cursor:not-allowed;}

/* sidebar */
.sb{
  width:var(--sw);
  flex-shrink:0;
  background:rgba(15,23,42,.96);
  color:#e5e7eb;
  border-right:1px solid rgba(148,163,184,.18);
  display:flex;
  flex-direction:column;
  overflow:hidden;
  box-shadow:14px 0 40px rgba(15,23,42,.16);
}
.sb-logo{
  padding:22px 18px 18px;
  border-bottom:1px solid rgba(148,163,184,.15);
  flex-shrink:0;
  background:linear-gradient(135deg,rgba(79,70,229,.22),rgba(14,165,233,.08));
}
.logo{
  font-family:var(--disp);
  font-weight:800;
  font-size:1.42rem;
  color:#fff;
  letter-spacing:-.045em;
  line-height:1;
}
.logo span{
  color:#8b5cf6;
  text-shadow:0 0 22px rgba(139,92,246,.65);
}
.logo-sub{
  font-family:var(--mono);
  font-size:.62rem;
  color:#94a3b8;
  text-transform:uppercase;
  letter-spacing:.16em;
  margin-top:8px;
}
.sb-sec{
  padding:16px 18px 8px;
  font-family:var(--mono);
  font-size:.64rem;
  color:#94a3b8;
  text-transform:uppercase;
  letter-spacing:.14em;
  display:flex;
  align-items:center;
  justify-content:space-between;
}
.cl-list{flex:1;overflow-y:auto;padding:8px 12px 12px;}
.cl-list::-webkit-scrollbar{width:7px;}
.cl-list::-webkit-scrollbar-track{background:transparent;}
.cl-list::-webkit-scrollbar-thumb{background:rgba(148,163,184,.25);border-radius:999px;}
.ci{
  display:flex;
  align-items:center;
  gap:10px;
  padding:11px 12px;
  border-radius:14px;
  cursor:pointer;
  transition:all .18s ease;
  margin-bottom:8px;
  position:relative;
  border:1px solid rgba(148,163,184,.12);
  background:rgba(255,255,255,.035);
}
.ci:hover{background:rgba(255,255,255,.075);transform:translateY(-1px);border-color:rgba(148,163,184,.22);}
.ci.act{background:linear-gradient(135deg,rgba(79,70,229,.28),rgba(14,165,233,.12));border-color:rgba(129,140,248,.45);box-shadow:0 12px 24px rgba(0,0,0,.18);}
.ci-dot{width:10px;height:34px;border-radius:999px;flex-shrink:0;box-shadow:0 0 18px currentColor;}
.ci-type-badge{font-family:var(--mono);font-size:.56rem;font-weight:800;padding:4px 7px;border-radius:999px;letter-spacing:.05em;flex-shrink:0;}
.ci-type-att{background:rgba(34,197,94,.15);color:#86efac;border:1px solid rgba(134,239,172,.24);}
.ci-type-marks{background:rgba(139,92,246,.17);color:#c4b5fd;border:1px solid rgba(196,181,253,.24);}
.ci-info{flex:1;min-width:0;}
.ci-name{font-family:var(--sans);font-size:.9rem;font-weight:700;color:#f8fafc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.ci-sub{font-family:var(--mono);font-size:.64rem;color:#94a3b8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:3px;}
.ci-del{background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.12);color:#64748b;cursor:pointer;font-size:1rem;width:24px;height:24px;border-radius:999px;line-height:1;transition:all .15s;flex-shrink:0;}
.ci-del:hover{color:#fecaca;background:rgba(239,68,68,.22);border-color:rgba(248,113,113,.35);}
.sb-foot{padding:14px 12px 16px;border-top:1px solid rgba(148,163,184,.15);flex-shrink:0;display:flex;flex-direction:column;gap:9px;background:rgba(2,6,23,.22);}
.add-cls-btn,.exp-all-btn,.def-btn{width:100%;font-family:var(--sans);font-size:.82rem;font-weight:800;border:none;padding:11px 12px;border-radius:13px;cursor:pointer;transition:all .16s ease;letter-spacing:.01em;box-shadow:var(--shadow-sm);}
.add-cls-btn{background:rgba(255,255,255,.06);border:1px dashed rgba(148,163,184,.36);color:#e2e8f0;}
.add-cls-btn:hover{background:rgba(79,70,229,.18);border-color:#818cf8;color:#fff;transform:translateY(-1px);}
.exp-all-btn{background:linear-gradient(135deg,#16a34a,#059669);color:#ecfdf5;}
.exp-all-btn:hover{filter:brightness(1.05);transform:translateY(-1px);box-shadow:0 14px 26px rgba(5,150,105,.26);}
.def-btn{background:linear-gradient(135deg,#f97316,#ea580c);color:#fff7ed;}
.def-btn:hover{filter:brightness(1.05);transform:translateY(-1px);box-shadow:0 14px 26px rgba(234,88,12,.24);}
.save-badge{font-family:var(--mono);font-size:.63rem;color:#64748b;text-align:center;padding:2px 0 5px;letter-spacing:.04em;}
.save-badge.saved{color:#86efac;}.save-badge.saving{color:#fde68a;}

/* main */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;background:transparent;padding:10px 14px 10px 0;min-height:0;}
.no-class{
  display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:16px;
  font-family:var(--sans);color:var(--muted);font-size:.95rem;background:rgba(255,255,255,.75);
  border:1px solid rgba(226,232,240,.85);border-radius:24px;box-shadow:var(--shadow);
}
.no-class-icon{font-size:3.8rem;opacity:.45;}

/* info bar */
.ibar{
  background:rgba(255,255,255,.82);
  backdrop-filter:blur(18px);
  border:1px solid rgba(226,232,240,.9);
  border-bottom:1px solid rgba(226,232,240,.9);
  border-radius:22px 22px 0 0;
  flex-shrink:0;
  box-shadow:var(--shadow-sm);
}
.ibar.marks-mode{border-top:4px solid var(--purple);} 
.ibar:not(.marks-mode){border-top:4px solid var(--accent);} 
.ibar-row{display:flex;align-items:flex-end;gap:10px;padding:9px 16px;flex-wrap:wrap;}
.ibar-f{display:flex;flex-direction:column;gap:3px;}
.ibar-lbl{font-family:var(--mono);font-size:.62rem;color:var(--muted);text-transform:uppercase;letter-spacing:.12em;font-weight:700;}
.ibar-inp{font-family:var(--sans);font-size:.84rem;font-weight:700;color:var(--text);background:#fff;border:1px solid var(--border);outline:none;border-radius:10px;padding:6px 10px;transition:all .15s;box-shadow:0 1px 0 rgba(15,23,42,.03);}
.ibar-inp:focus{border-color:#818cf8;box-shadow:0 0 0 4px rgba(99,102,241,.14);}
.ibar-inp::placeholder{color:#cbd5e1;font-weight:500;}
.ibar-inp.w1{min-width:240px;}.ibar-inp.w2{min-width:170px;}.ibar-inp.w3{min-width:130px;}
.mode-pill{font-family:var(--mono);font-size:.66rem;font-weight:900;padding:8px 12px;border-radius:999px;letter-spacing:.08em;align-self:center;flex-shrink:0;}
.mode-pill.att{background:#dcfce7;color:#166534;border:1px solid #bbf7d0;}
.mode-pill.marks{background:#f3e8ff;color:#6d28d9;border:1px solid #ddd6fe;}
.ibar-collapsed{display:flex;align-items:center;gap:12px;padding:8px 16px;}
.ibar-summary{flex:1;min-width:0;font-family:var(--sans);font-size:.82rem;font-weight:700;color:#334155;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.ibar-toggle{flex-shrink:0;font-family:var(--mono);font-size:.64rem;font-weight:800;letter-spacing:.04em;text-transform:uppercase;background:#fff;border:1px solid var(--border);color:#475569;cursor:pointer;padding:6px 11px;border-radius:999px;transition:all .15s;box-shadow:var(--shadow-sm);align-self:center;}
.ibar-toggle:hover{border-color:#818cf8;color:#4f46e5;background:#eef2ff;}

/* topbar */
.topbar{background:rgba(255,255,255,.72);display:flex;align-items:center;gap:9px;padding:7px 16px;flex-shrink:0;flex-wrap:wrap;border-left:1px solid rgba(226,232,240,.9);border-right:1px solid rgba(226,232,240,.9);}
.btn{font-family:var(--sans);font-size:.82rem;font-weight:800;border:none;cursor:pointer;border-radius:11px;padding:7px 13px;transition:all .15s ease;white-space:nowrap;letter-spacing:.01em;box-shadow:var(--shadow-sm);}
.btn:hover{transform:translateY(-1px);}
.btn-ghost{background:#fff;border:1px solid var(--border);color:#475569;}
.btn-ghost:hover{border-color:#cbd5e1;color:#0f172a;background:#f8fafc;box-shadow:0 8px 20px rgba(15,23,42,.08);}
.btn-accent{background:linear-gradient(135deg,#6366f1,#4f46e5);color:#fff;}
.btn-accent:hover{box-shadow:0 12px 24px rgba(79,70,229,.25);}
.btn-blue{background:linear-gradient(135deg,#0ea5e9,#0284c7);color:#fff;}
.btn-blue:hover{box-shadow:0 12px 24px rgba(2,132,199,.23);}
.btn-green{background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;}
.btn-green:hover{box-shadow:0 12px 24px rgba(22,163,74,.22);}
.btn-amber{background:linear-gradient(135deg,#f59e0b,#d97706);color:#fff;}
.btn-amber:hover{box-shadow:0 12px 24px rgba(217,119,6,.22);}
.btn-purple{background:linear-gradient(135deg,#8b5cf6,#7c3aed);color:#fff;}
.btn-purple:hover{box-shadow:0 12px 24px rgba(124,58,237,.22);}
.btn-teal{background:linear-gradient(135deg,#14b8a6,#0d9488);color:#fff;}
.btn-teal:hover{box-shadow:0 12px 24px rgba(13,148,136,.22);}

/* subbar */
.subbar{background:rgba(255,255,255,.78);border-left:1px solid rgba(226,232,240,.9);border-right:1px solid rgba(226,232,240,.9);border-top:1px solid rgba(226,232,240,.75);display:flex;align-items:center;gap:10px;padding:7px 16px;flex-shrink:0;flex-wrap:wrap;}
.slabel{font-family:var(--mono);font-size:.66rem;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;white-space:nowrap;font-weight:800;}
.inp{font-family:var(--sans);font-size:.86rem;background:#fff;border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:10px;outline:none;transition:all .15s;box-shadow:var(--shadow-sm);}
.inp:focus{border-color:#818cf8;box-shadow:0 0 0 4px rgba(99,102,241,.14);}
.inp::placeholder{color:#94a3b8;}
.inp-r{width:130px;}.inp-s{width:145px;}.inp-n{width:260px;}.inp-d{width:150px;}
.sdiv{width:1px;height:28px;background:var(--border);margin:0 3px;}

/* searchbar */
.searchbar{background:rgba(248,250,252,.94);border-left:1px solid rgba(226,232,240,.9);border-right:1px solid rgba(226,232,240,.9);border-top:1px solid rgba(226,232,240,.75);display:flex;align-items:center;gap:11px;padding:7px 16px;flex-shrink:0;}
.search-icon{font-size:1rem;color:#64748b;flex-shrink:0;}
.search-inp{flex:1;font-family:var(--sans);font-size:.9rem;background:#fff;border:1px solid var(--border);outline:none;color:var(--text);min-width:0;padding:7px 12px;border-radius:999px;box-shadow:var(--shadow-sm);}
.search-inp:focus{border-color:#818cf8;box-shadow:0 0 0 4px rgba(99,102,241,.14);}
.search-inp::placeholder{color:#94a3b8;}
.search-clear{font-family:var(--sans);font-size:.78rem;color:#64748b;background:#fff;border:1px solid var(--border);cursor:pointer;padding:7px 10px;border-radius:999px;font-weight:700;}
.search-clear:hover{color:#dc2626;border-color:#fecaca;background:#fff5f5;}
.search-count{font-family:var(--mono);font-size:.68rem;color:var(--muted);white-space:nowrap;background:#fff;border:1px solid var(--border);padding:7px 10px;border-radius:999px;}

/* statsbar */
.statsbar{background:#0f172a;display:flex;align-items:center;gap:10px;padding:6px 16px;flex-shrink:0;font-family:var(--mono);font-size:.72rem;flex-wrap:wrap;border-left:1px solid #0f172a;border-right:1px solid #0f172a;}
.stat{display:flex;gap:6px;align-items:center;background:rgba(255,255,255,.06);border:1px solid rgba(148,163,184,.14);padding:4px 10px;border-radius:999px;}
.slbl{color:#94a3b8;}.sv{font-weight:900;}
.sv.g{color:#86efac;}.sv.r{color:#fca5a5;}.sv.y{color:#fde68a;}.sv.am{color:#fdba74;}.sv.pu{color:#c4b5fd;}
.hint{margin-left:auto;font-size:.62rem;color:#64748b;}

/* weight bar */
.wbar{background:linear-gradient(135deg,#211044,#17072f);border-bottom:1px solid rgba(196,181,253,.18);display:flex;align-items:center;gap:8px;padding:5px 14px;flex-shrink:0;flex-wrap:wrap;}
.wbar-label{font-family:var(--mono);font-size:.58rem;color:#c4b5fd;text-transform:uppercase;letter-spacing:.08em;white-space:nowrap;flex-shrink:0;font-weight:900;}
.wbar-pills{display:flex;gap:5px;flex-wrap:wrap;flex:1;}
.wpill{display:flex;align-items:center;gap:5px;background:rgba(255,255,255,.08);border:1px solid rgba(196,181,253,.18);border-radius:999px;padding:3px 7px;font-family:var(--mono);font-size:.6rem;}
.wpill-name{color:#ede9fe;max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.wpill-inp{width:38px;background:rgba(15,23,42,.55);border:1px solid rgba(196,181,253,.22);border-radius:8px;color:#fff;font-family:var(--mono);font-size:.6rem;padding:3px 4px;outline:none;text-align:center;}
.wpill-inp:focus{border-color:#c4b5fd;box-shadow:0 0 0 2px rgba(196,181,253,.16);}
.wpill-pct{color:#a78bfa;font-size:.58rem;}
.wbar-total{font-family:var(--mono);font-size:.62rem;padding:4px 8px;border-radius:999px;white-space:nowrap;font-weight:900;}
.wbar-total.ok{background:rgba(34,197,94,.16);color:#bbf7d0;border:1px solid rgba(187,247,208,.26);}
.wbar-total.warn{background:rgba(245,158,11,.15);color:#fde68a;border:1px solid rgba(253,230,138,.26);}
.wbar-mode{font-family:var(--mono);font-size:.58rem;color:#a78bfa;text-transform:uppercase;letter-spacing:.08em;white-space:nowrap;font-weight:800;}
.wbar-toggle{font-family:var(--mono);font-size:.6rem;background:rgba(255,255,255,.08);color:#ddd6fe;border:1px solid rgba(196,181,253,.25);border-radius:999px;padding:4px 9px;cursor:pointer;transition:all .15s;white-space:nowrap;font-weight:800;}
.wbar-toggle:hover,.wbar-toggle.on{background:rgba(139,92,246,.26);border-color:#a78bfa;color:#fff;}

/* sheet */
.sh-outer{flex:1;overflow:hidden;background:#fff;border-left:1px solid rgba(226,232,240,.9);border-right:1px solid rgba(226,232,240,.9);border-bottom:1px solid rgba(226,232,240,.9);border-radius:0 0 22px 22px;box-shadow:var(--shadow);}
.sh-scroll{width:100%;height:100%;overflow:auto;background:#fff;}
.sh-scroll::-webkit-scrollbar{width:12px;height:12px;}
.sh-scroll::-webkit-scrollbar-track{background:#f8fafc;}
.sh-scroll::-webkit-scrollbar-thumb{background:#cbd5e1;border:3px solid #f8fafc;border-radius:999px;}
.sh-scroll::-webkit-scrollbar-thumb:hover{background:#94a3b8;}
.sh-scroll::-webkit-scrollbar-corner{background:#f8fafc;}
.sh-tbl{border-collapse:separate;border-spacing:0;table-layout:fixed;min-width:max-content;}

/* col headers */
.chr th{
  background:#111827;color:#cbd5e1;font-family:var(--mono);font-size:.63rem;font-weight:800;
  text-transform:uppercase;letter-spacing:.055em;border-right:1px solid rgba(148,163,184,.18);
  border-bottom:1px solid rgba(148,163,184,.22);padding:0 8px;height:38px;vertical-align:middle;
  position:sticky;top:0;z-index:10;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;user-select:none;
}
.chr th.fr{position:sticky;left:0;z-index:22;background:#0f172a;text-align:center;width:var(--roll-w);min-width:var(--roll-w);max-width:var(--roll-w);}
.chr th.fs{position:sticky;left:var(--roll-w);z-index:22;background:#0f172a;text-align:center;width:var(--sys-w);min-width:var(--sys-w);max-width:var(--sys-w);}
.chr th.fn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:22;background:#0f172a;border-right:3px solid var(--accent);width:var(--name-w);min-width:var(--name-w);max-width:var(--name-w);}
.chr th.fn.marks-mode{border-right-color:var(--purple);}
.chr th.ft{background:#1e293b;text-align:center;color:#fde68a;width:112px;min-width:112px;max-width:112px;}
.chr th.fw{background:#25164d;text-align:center;color:#ddd6fe;width:86px;min-width:86px;max-width:86px;}
.chr th.dc{text-align:center;width:${CELL_W}px;min-width:${CELL_W}px;max-width:${CELL_W}px;background:#111827;}
.chr th.mc{text-align:center;width:${MARK_W}px;min-width:${MARK_W}px;max-width:${MARK_W}px;background:#24144d;}
.chr th.mc:hover .xbtn,.chr th.dc:hover .xbtn{opacity:1;}
.dch{display:flex;flex-direction:column;align-items:center;gap:2px;position:relative;}
.dch-t{font-size:.58rem;color:#94a3b8;}.dch-m{font-size:.82rem;font-weight:900;color:#f8fafc;}
.dch-sub{font-size:.57rem;color:#c4b5fd;}
.dch-w{font-size:.55rem;color:#a78bfa;margin-top:1px;}
.xbtn{position:absolute;top:-7px;right:-12px;background:#ef4444;color:#fff;border:none;border-radius:999px;width:18px;height:18px;font-size:.64rem;line-height:18px;text-align:center;cursor:pointer;opacity:0;transition:all .13s;padding:0;box-shadow:0 8px 16px rgba(239,68,68,.24);}
.xbtn:hover{transform:scale(1.08);}

/* data rows */
.dr{height:34px;}
.dr:nth-child(even) td{background:var(--row-alt);}
.dr:nth-child(odd) td{background:var(--surface);}
.dr:hover td{background:#eef2ff!important;}
.dr td{border-right:1px solid #edf2f7;border-bottom:1px solid #edf2f7;padding:0;height:34px;vertical-align:middle;transition:background .12s;}

/* low-att highlight */
.dr.low-att td{background:#fffbeb!important;}
.dr.low-att:hover td{background:#fef3c7!important;}
.dr.low-att .tdr,.dr.low-att .tds{background:#fef3c7!important;color:#92400e!important;}
.dr.low-att .tdn{background:#fef3c7!important;border-right-color:#f59e0b!important;}
.dr.low-att .tdt{background:#fff7ed!important;}
.low-badge{display:inline-block;background:linear-gradient(135deg,#f59e0b,#f97316);color:#fff;font-family:var(--mono);font-size:.53rem;font-weight:900;padding:3px 6px;border-radius:999px;margin-left:7px;vertical-align:middle;letter-spacing:.04em;}

/* frozen roll */
.tdr{position:sticky;left:0;z-index:5;background:var(--frozen)!important;border-right:1px solid var(--border)!important;font-family:var(--mono);font-size:.75rem;color:#64748b;text-align:center;width:var(--roll-w);min-width:var(--roll-w);max-width:var(--roll-w);padding:0;height:34px;line-height:34px;font-weight:700;overflow:hidden;}
.dr:hover .tdr{background:#e0e7ff!important;}
.rinp{flex:1;min-width:0;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:.75rem;color:#0f172a;text-align:center;padding:0 4px;font-weight:700;}
.rinp:focus{background:#fff;box-shadow:inset 0 0 0 2px #818cf8;}
.dark .rinp{color:#e2e8f0;}
.dark .rinp:focus{background:#1a2030;box-shadow:inset 0 0 0 2px #818cf8;}
.xrow{background:#fee2e2;border:1px solid #fecaca;color:#dc2626;cursor:pointer;font-size:.85rem;width:20px;height:20px;border-radius:999px;padding:0;display:none;line-height:1;flex-shrink:0;}
.dr:hover .xrow{display:inline-flex;align-items:center;justify-content:center;}
.dark .xrow{background:rgba(239,68,68,.18);border-color:rgba(248,113,113,.3);color:#fca5a5;}

/* frozen system id */
.tds{position:sticky;left:var(--roll-w);z-index:5;background:var(--frozen)!important;border-right:1px solid var(--border)!important;font-family:var(--mono);font-size:.75rem;color:#475569;text-align:center;width:var(--sys-w);min-width:var(--sys-w);max-width:var(--sys-w);padding:0;height:34px;overflow:hidden;}
.dr:hover .tds{background:#e0e7ff!important;}
.sinp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:.75rem;color:#0f172a;text-align:center;padding:0 8px;display:block;font-weight:600;}
.sinp:focus{background:#fff;box-shadow:inset 0 0 0 2px #818cf8;}

/* frozen name */
.tdn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:5;background:var(--frozen)!important;border-right:3px solid var(--accent)!important;width:var(--name-w);min-width:var(--name-w);max-width:var(--name-w);padding:0;height:34px;display:flex;align-items:center;overflow:hidden;}
.tdn.marks-mode{border-right-color:var(--purple)!important;}
.dr:hover .tdn{background:#e0e7ff!important;}
.ninp{flex:1;height:34px;min-width:0;background:transparent;border:none;outline:none;font-family:var(--sans);font-size:.9rem;font-weight:700;color:#0f172a;padding:0 12px;display:block;}
.ninp:focus{background:#fff;box-shadow:inset 0 0 0 2px #818cf8;}

/* totals / weighted */
.tdt{text-align:center;font-family:var(--mono);font-size:.76rem;font-weight:900;background:#fff7ed!important;border-left:1px solid #fed7aa!important;width:112px;min-width:112px;max-width:112px;height:34px;color:#9a3412;}
.dr:hover .tdt{background:#ffedd5!important;}
.tdw{text-align:center;font-family:var(--mono);font-size:.76rem;font-weight:900;background:#f5f3ff!important;border-left:1px solid #ddd6fe!important;width:86px;min-width:86px;max-width:86px;height:34px;color:#6d28d9;}
.dr:hover .tdw{background:#ede9fe!important;}

.tda{text-align:center;width:${CELL_W}px;min-width:${CELL_W}px;max-width:${CELL_W}px;padding:0;height:34px;}
.ainp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:.92rem;font-weight:900;text-align:center;text-transform:uppercase;cursor:text;transition:all .12s;display:block;}
.ainp.ip{background:#dcfce7!important;color:#166534;}
.ainp.ia{background:#fee2e2!important;color:#991b1b;}
.ainp:focus{box-shadow:inset 0 0 0 2px #0ea5e9;position:relative;z-index:2;background:#fff!important;}

/* marks cells */
.tmc{text-align:center;width:${MARK_W}px;min-width:${MARK_W}px;max-width:${MARK_W}px;padding:0;height:34px;}
.minp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:.86rem;font-weight:900;text-align:center;cursor:text;color:#0f172a;transition:all .12s;display:block;}
.minp:focus{box-shadow:inset 0 0 0 2px var(--purple);position:relative;z-index:2;background:#fff!important;}
.minp.full{color:#166534;background:#dcfce7!important;}
.minp.good{color:#1d4ed8;background:#eff6ff!important;}
.minp.low-m{color:#991b1b;background:#fee2e2!important;}
.minp.absent-m{color:#64748b;background:#f1f5f9!important;}

/* summary row */
.sr td{background:#0f172a!important;color:#cbd5e1;font-family:var(--mono);font-size:.67rem;border-top:1px solid rgba(148,163,184,.22);border-right:1px solid rgba(148,163,184,.14);border-bottom:none;height:34px;text-align:center;vertical-align:middle;position:sticky;bottom:0;z-index:5;font-weight:800;}
.sr td.sfr{position:sticky;left:0;z-index:16;background:#020617!important;color:#64748b;font-size:.6rem;text-transform:uppercase;letter-spacing:.1em;width:var(--roll-w);min-width:var(--roll-w);}
.sr td.sfs{position:sticky;left:var(--roll-w);z-index:16;background:#020617!important;color:#64748b;font-size:.6rem;text-transform:uppercase;letter-spacing:.1em;width:var(--sys-w);min-width:var(--sys-w);}
.sr td.sfn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:16;background:#020617!important;color:#94a3b8;border-right:3px solid var(--accent)!important;text-align:left;padding:0 12px;font-size:.68rem;width:var(--name-w);min-width:var(--name-w);}
.sr td.sfn.marks-mode{border-right-color:var(--purple)!important;}
.sr td.sft{background:#1e293b!important;color:#fde68a;}
.sr td.sfw{background:#25164d!important;color:#ddd6fe;}
.sp{color:#86efac;font-weight:900;}.sa{color:#fca5a5;font-weight:900;}
.empty{padding:90px;text-align:center;font-family:var(--sans);font-size:.9rem;color:var(--muted);background:#fff;}

/* modals */
.mbg{position:fixed;inset:0;background:rgba(15,23,42,.64);backdrop-filter:blur(10px);z-index:100;display:flex;align-items:center;justify-content:center;padding:18px;}
.modal{background:#fff;border:1px solid rgba(226,232,240,.95);border-radius:24px;padding:28px;width:500px;max-width:96vw;box-shadow:var(--shadow-lg);}
.modal.wide{width:600px;}.modal.narrow{width:440px;}
.modal h2{font-family:var(--disp);font-size:1.35rem;color:#0f172a;margin-bottom:6px;letter-spacing:-.035em;}
.modal-sub{font-family:var(--sans);font-size:.82rem;color:#64748b;margin-bottom:22px;line-height:1.6;}
.mf{margin-bottom:14px;}
.ml{font-family:var(--mono);font-size:.66rem;color:#64748b;text-transform:uppercase;letter-spacing:.1em;margin-bottom:7px;display:block;font-weight:800;}
.mi{width:100%;font-family:var(--sans);font-size:.9rem;background:#fff;border:1px solid var(--border);color:#0f172a;padding:11px 13px;border-radius:13px;outline:none;transition:all .15s;box-shadow:var(--shadow-sm);}
.mi:focus{border-color:#818cf8;box-shadow:0 0 0 4px rgba(99,102,241,.14);}
.mi::placeholder{color:#94a3b8;}.mi-sm{width:110px;}
.mrow{display:flex;gap:13px;}
.mact{display:flex;gap:10px;justify-content:flex-end;margin-top:24px;}
.mbtn-c,.mbtn-ok,.mbtn-blue,.mbtn-purple,.mbtn-green{border:none;padding:10px 17px;border-radius:13px;font-family:var(--sans);font-size:.84rem;font-weight:800;cursor:pointer;transition:all .15s;box-shadow:var(--shadow-sm);}
.mbtn-c{background:#f8fafc;border:1px solid var(--border);color:#475569;}
.mbtn-c:hover{background:#f1f5f9;color:#0f172a;}
.mbtn-ok{background:linear-gradient(135deg,#6366f1,#4f46e5);color:#fff;}
.mbtn-blue{background:linear-gradient(135deg,#0ea5e9,#0284c7);color:#fff;}
.mbtn-purple{background:linear-gradient(135deg,#8b5cf6,#7c3aed);color:#fff;}
.mbtn-green{background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;}
.mbtn-ok:hover,.mbtn-blue:hover,.mbtn-purple:hover,.mbtn-green:hover{transform:translateY(-1px);filter:brightness(1.03);}
.type-pick{display:flex;gap:12px;margin-bottom:20px;}
.type-opt{flex:1;padding:18px 12px;border-radius:18px;border:1px solid var(--border);cursor:pointer;text-align:center;transition:all .15s;background:#f8fafc;}
.type-opt:hover{border-color:#cbd5e1;background:#fff;transform:translateY(-1px);}
.type-opt.sel-att{border-color:#86efac;background:#f0fdf4;box-shadow:0 0 0 4px rgba(34,197,94,.10);}
.type-opt.sel-marks{border-color:#c4b5fd;background:#faf5ff;box-shadow:0 0 0 4px rgba(139,92,246,.10);}
.type-opt-icon{font-size:1.75rem;margin-bottom:8px;}
.type-opt-label{font-family:var(--sans);font-size:.95rem;color:#0f172a;font-weight:900;}
.type-opt-desc{font-family:var(--sans);font-size:.74rem;color:#64748b;margin-top:4px;}
.cls-pick-list{display:flex;flex-direction:column;gap:8px;max-height:260px;overflow-y:auto;margin-top:6px;}
.cls-pick-item{display:flex;align-items:center;gap:12px;padding:13px 14px;background:#f8fafc;border:1px solid var(--border);border-radius:16px;cursor:pointer;transition:all .15s;}
.cls-pick-item:hover{border-color:#cbd5e1;background:#fff;transform:translateY(-1px);}
.cls-pick-item.selected{border-color:#818cf8;background:#eef2ff;box-shadow:0 0 0 4px rgba(99,102,241,.10);}
.cls-pick-dot{width:10px;height:34px;border-radius:999px;flex-shrink:0;}
.cls-pick-info{flex:1;min-width:0;}
.cls-pick-name{font-family:var(--sans);font-size:.9rem;font-weight:900;color:#0f172a;}
.cls-pick-sub{font-family:var(--mono);font-size:.66rem;color:#64748b;margin-top:3px;}

/* import modal */
.drop-zone{border:2px dashed #cbd5e1;border-radius:18px;padding:36px 22px;text-align:center;cursor:pointer;transition:all .15s;background:#f8fafc;}
.drop-zone:hover,.drop-zone.drag-over{border-color:#818cf8;background:#eef2ff;box-shadow:0 0 0 4px rgba(99,102,241,.10);}
.drop-zone-icon{font-size:2.4rem;margin-bottom:10px;opacity:.75;}
.drop-zone-text{font-family:var(--sans);font-size:.88rem;color:#475569;line-height:1.8;font-weight:700;}
.drop-zone-hint{font-size:.72rem;color:#94a3b8;margin-top:8px;font-weight:500;}
.import-preview{background:#f8fafc;border:1px solid var(--border);border-radius:16px;padding:12px 14px;margin-top:14px;max-height:190px;overflow-y:auto;}
.import-preview-row{display:flex;gap:12px;font-family:var(--mono);font-size:.72rem;padding:6px 0;border-bottom:1px solid #e2e8f0;}
.import-preview-row:last-child{border-bottom:none;}
.ipr-roll{color:#7c3aed;width:80px;flex-shrink:0;font-weight:900;}
.ipr-name{color:#0f172a;flex:1;font-weight:700;}
.import-status{font-family:var(--sans);font-size:.8rem;padding:10px 12px;border-radius:14px;margin-top:12px;text-align:center;font-weight:800;}
.import-status.ok{background:#dcfce7;color:#166534;border:1px solid #bbf7d0;}
.import-status.warn{background:#fef3c7;color:#92400e;border:1px solid #fde68a;}
.import-status.err{background:#fee2e2;color:#991b1b;border:1px solid #fecaca;}
`;
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
  const [showWeighted,     setShowWeighted]     = useState(false);
  const [infoCollapsed,    setInfoCollapsed]    = useState(false);
  const [draft,            setDraft]            = useState({faculty:"",session:"",teacher:"",subject:"",course:"",type:"attendance"});
  const [assessDraft,      setAssessDraft]      = useState({name:"Quiz 1",outOf:"10",weight:"",type:"quiz"});
  const [importData,       setImportData]       = useState(null); // {rows:[{roll,name}], fileName, warnings:[]}
  const [importDragOver,   setImportDragOver]   = useState(false);
  const [darkMode,         setDarkMode]         = useState(() => {
    try { return localStorage.getItem("attendx_dark") === "1"; } catch { return false; }
  });
  const [showGradeModal,   setShowGradeModal]   = useState(false);
  const [gradeDraft,       setGradeDraft]       = useState([]);

  // ── undo / redo ──
  const undoStack = useRef([]);
  const redoStack = useRef([]);
  const [undoLen, setUndoLen] = useState(0);
  const [redoLen, setRedoLen] = useState(0);
  const skipHistory = useRef(false);

  const nextSid   = useRef(1000);
  const refs      = useRef({});
  const initDone  = useRef(false);
  const saveTimer = useRef(null);
  const fileRef   = useRef(null);
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

  useEffect(() => { setSearchQ(""); setLowAttOnly(false); }, [activeId]);

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
  const sortedDates = active && isAtt ? [...(active.dates ?? [])].sort() : [];
  const assessments = active?.assessments ?? [];

  // Safe cell reader must be initialized before attendance summaries are computed.
  const getCell = (c, sid, d) => c?.att?.[`${sid}_${d}`] ?? "";

  // Weight helpers
  const totalWeight = assessments.reduce((s, a) => s + (parseFloat(a.weight) || 0), 0);
  const weightingOn = showWeighted && isMarks && assessments.some(a => parseFloat(a.weight) > 0);
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

  const lowAttendanceStudents = active && isAtt
    ? active.students.filter(s => { const sm = stuAttSum(active, s.id); return sm && sm.pct < LOW_ATT_THRESHOLD; })
    : [];

  const filteredStudents = active
    ? active.students.filter(s => {
        const sm = isAtt ? stuAttSum(active, s.id) : null;
        if (lowAttOnly && (!sm || sm.pct >= LOW_ATT_THRESHOLD)) return false;
        if (!searchQ.trim()) return true;
        const q = searchQ.trim().toLowerCase();
        return String(s.name ?? "").toLowerCase().includes(q)
          || String(s.roll ?? "").toLowerCase().includes(q)
          || String(s.systemId ?? "").toLowerCase().includes(q);
      })
    : [];

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
    reader.onload = (e) => {
      try {
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
  const markRestP  = () => updActive(c=>{ const a={...c.att}; c.students.forEach(s=>sortedDates.forEach(d=>{const k=`${s.id}_${d}`;if(!a[k])a[k]="P";})); return {...c,att:a}; });
  const markRestA  = () => updActive(c=>{ const a={...c.att}; c.students.forEach(s=>sortedDates.forEach(d=>{const k=`${s.id}_${d}`;if(!a[k])a[k]="A";})); return {...c,att:a}; });

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
  const overall = active&&isAtt ? active.students.reduce((acc,s)=>{ sortedDates.forEach(d=>{const v=getCell(active,s.id,d);if(isP(v))acc.p++;else if(isA(v))acc.a++;}); return acc; },{p:0,a:0}) : {p:0,a:0};
  const lowAttCount = lowAttendanceStudents.length;

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
    else if(e.key==="ArrowRight"||(e.key==="Tab"&&!e.shiftKey))nd=Math.min(dIdx+1,sortedDates.length-1);
    else if(e.key==="ArrowLeft"||(e.key==="Tab"&&e.shiftKey))nd=Math.max(dIdx-1,0);
    refs.current[`att_${sts[ns].id}_${sortedDates[nd]}`]?.focus();
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
    refs.current[`mark_${sts[ns].id}_${assessments[nd]?.id}`]?.focus();
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
    const outOfRow=["","","","Out Of",...asm.map(a=>a.outOf),"Total %","Weighted %","Grade"];
    const weightRow=["","","","Weight %",...asm.map(a=>a.weight?(a.weight+"%"):"—"),"","",""];
    const hdr=["#","Student ID","System ID","Student Name",...asm.map(a=>`${a.name} (${a.type})`),"Total %","Weighted %","Grade"];
    const rows=(cls.students??[]).map((s,i)=>{
      const sm=stuMarksSum(cls,s.id);
      const wm=hasWeights?stuWeightedSum(cls,s.id):null;
      const gradePct = wm ? wm.pct : (sm ? sm.pct : null);
      return [i+1,s.roll,s.systemId||"",s.name,...asm.map(a=>getMark(cls,s.id,a.id)||""),sm?sm.pct+"%":"-",wm?wm.pct+"%":"-",getG(gradePct)];
    });
    const avgRow=["","","","Average",...asm.map(a=>assessColAvg(cls,a.id)??"-"),"","",""];
    const ws=XLSX.utils.aoa_to_sheet([...infoRows,outOfRow,weightRow,hdr,...rows,[],avgRow]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:14},{wch:28},...asm.map(()=>({wch:14})),{wch:10},{wch:12},{wch:8}];
    return ws;
  };
  const exportAll = () => {
    if(!classes.length)return;
    const wb=XLSX.utils.book_new(); const used={};
    classes.forEach(cls=>{ let sn=((cls.info?.subject)||(cls.info?.course)||"Class").replace(/[:\\/?*\[\]]/g,"").slice(0,24)||"Sheet"; const sfx=cls.type==="marks"?" Marks":" Att"; let fn=(sn+sfx).slice(0,31); let n=2; while(used[fn]){fn=`${sn}_${n++}`;} used[fn]=1; XLSX.utils.book_append_sheet(wb, cls.type==="marks"?buildMarksWs(cls):buildAttWs(cls), fn); });
    XLSX.writeFile(wb,"attendx_all_classes.xlsx");
  };

  const exportDefaultersAll = () => {
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
  const exportOne = () => {
    if(!active)return;
    const wb=XLSX.utils.book_new();
    const sn=((active.info?.subject)||(active.info?.course)||"Sheet").replace(/[:\\/?*\[\]]/g,"").slice(0,31);
    XLSX.utils.book_append_sheet(wb, isMarks?buildMarksWs(active):buildAttWs(active), sn||"Sheet1");
    XLSX.writeFile(wb,`attendx_${sn||"export"}.xlsx`);
  };

  // ── render ────────────────────────────────────────────────────────────────────
  const STYLE = makeStyle(rollW, sysW, nameW);

  // grade cell color for weighted
  const wGradeCls = (pct) => {
    if (pct === null) return "";
    if (pct >= 90) return "full";
    if (pct >= 60) return "good";
    return "low-m";
  };

  return (
    <>
      <style>{STYLE}</style>
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
                <button className="btn btn-green" onClick={markRestP}>✓ Mark Rest Present</button>
                <button className="btn btn-ghost" onClick={markRestA}>✗ Mark Rest Absent</button>
                <button className={`btn ${lowAttOnly?"btn-amber":"btn-ghost"}`} onClick={()=>setLowAttOnly(v=>!v)}>⚠ Below 25%</button>
              </>}
              {isMarks&&<>
                <button className="btn btn-ghost" onClick={()=>{if(window.confirm("Clear all marks?"))updActive(c=>({...c,marks:{}}));}}>Clear Marks</button>
                <button className="btn btn-purple" onClick={openGradeModal}>🎓 Grading Criteria</button>
              </>}
              {classes.length>1&&<button className="btn btn-amber" onClick={openCopyModal}>⇄ Copy Roster</button>}
              <button className="btn btn-teal" onClick={()=>{setImportData(null);setShowImportModal(true);}}>↑ Import from Excel</button>
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
                {weightingOn&&<div className="stat"><span className="slbl">Weighted Mode:</span><span className="sv pu">ON</span></div>}
              </>}
              <span className="hint">{isAtt?"P = Present · A = Absent · Arrow keys to navigate":"Enter marks · A = Absent · Click column header to edit weight"}</span>
            </div>

            {/* WEIGHT BAR — only for marks mode */}
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
                <button
                  className={`wbar-toggle${showWeighted?" on":""}`}
                  onClick={()=>setShowWeighted(v=>!v)}
                  title="Toggle weighted total column"
                >
                  {showWeighted?"Weighted ON":"Show Weighted"}
                </button>
              </div>
            )}

            {/* SPREADSHEET */}
            <div className="sh-outer">
              <div className="sh-scroll">
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
                        {sortedDates.map(d=>{
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
                      {filteredStudents.map(s=>{
                        const sm=stuAttSum(active,s.id);
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
                            </td>
                            {sortedDates.map((d,dIdx)=>{
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
                    </tbody>
                    <tfoot>
                      <tr className="sr">
                        <td className="sfr">Σ</td>
                        <td className="sfs"></td>
                        <td className="sfn">{active.students.length} students · {sortedDates.length} dates</td>
                        {sortedDates.map(d=>{const{p,a}=dateAttSum(active,d);return <td key={d}><span className="sp">{p}P</span> <span className="sa">{a}A</span></td>;})}
                        <td className="sft"><span className="sp">{overall.p}P</span>/<span className="sa">{overall.a}A</span></td>
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
                        <th className="ft" style={{color:"#a78bfa"}}>Total %</th>
                        {weightingOn && <th className="fw">Weighted</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredStudents.map(s=>{
                        const sm=stuMarksSum(active,s.id);
                        const wm=weightingOn?stuWeightedSum(active,s.id):null;
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
                            <td className="tdt" style={{background:sm?"#f3f0ff":""}}>
                              {sm
                                ? <span style={{color:"#7c3aed",fontWeight:600}}>
                                    {sm.earned}/{sm.total} <span style={{color:"#aaa",fontSize:"0.61rem"}}>({sm.pct}%)</span>
                                    {!weightingOn && (() => { const g = getGrade(sm.pct); return g ? <span className={`grade-badge ${gradeColorClass(g)}`}>{g}</span> : null; })()}
                                  </span>
                                : <span style={{color:"#ccc"}}>—</span>}
                            </td>
                            {weightingOn && (
                              <td className="tdw">
                                {wm
                                  ? <span style={{color: wm.pct>=90?"#14633a":wm.pct>=60?"#1d4ed8":"#be2c0a", fontWeight:600}}>
                                      {wm.earned}<span style={{color:"#aaa",fontSize:"0.58rem",fontWeight:400}}>/{wm.total}</span>
                                      <span style={{display:"block",fontSize:"0.58rem",color:"#aaa",fontWeight:400}}>{wm.pct}%</span>
                                      {(() => { const g = getGrade(wm.pct); return g ? <span className={`grade-badge ${gradeColorClass(g)}`}>{g}</span> : null; })()}
                                    </span>
                                  : <span style={{color:"#ccc"}}>—</span>}
                              </td>
                            )}
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr className="sr">
                        <td className="sfr">Avg</td>
                        <td className="sfs"></td>
                        <td className={`sfn marks-mode`}>{active.students.length} students · {assessments.length} assessments</td>
                        {assessments.map(a=>{
                          const avg=assessColAvg(active,a.id);
                          return <td key={a.id} style={{color:"#a78bfa"}}>{avg?`${avg}/${a.outOf}`:"—"}</td>;
                        })}
                        <td className="sft" style={{color:"#a78bfa"}}>Class Avg</td>
                        {weightingOn && <td className="sfw">Weighted</td>}
                      </tr>
                    </tfoot>
                  </table>
                )}
              </div>
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