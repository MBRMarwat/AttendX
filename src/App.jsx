import { useState, useRef, useEffect, useMemo } from "react";
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
  }));
};

// ─── CSS ─────────────────────────────────────────────────────────────────────
const makeStyle = (rollW, sysW, nameW) => `
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body{height:100%;background:#0d0c09;}
:root{
  --bg:#f0ede6;--surface:#faf9f6;--border:#d4d0c8;--border-s:#b0aca4;
  --hdr:#18160f;--hdr2:#1e1c14;--fg:#f0ede6;
  --accent:#e04e20;--accent2:#2563eb;--gold:#f2c94c;--purple:#7c3aed;
  --p-bg:#d1f0e0;--p-fg:#14633a;--a-bg:#fde0db;--a-fg:#be2c0a;
  --frozen:#e8e4dc;--row-alt:#f5f2ec;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;--disp:'Syne',sans-serif;
  --sw:252px;
  --roll-w:${rollW}px;
  --sys-w:${sysW}px;
  --name-w:${nameW}px;
}
.shell{display:flex;height:100vh;overflow:hidden;background:#0d0c09;}

/* sidebar */
.sb{width:var(--sw);flex-shrink:0;background:#13120d;border-right:1px solid #222018;display:flex;flex-direction:column;overflow:hidden;}
.sb-logo{padding:15px 15px 11px;border-bottom:1px solid #222018;flex-shrink:0;}
.logo{font-family:var(--disp);font-weight:800;font-size:1.08rem;color:#fff;letter-spacing:-0.02em;}
.logo span{color:var(--accent);}
.logo-sub{font-family:var(--mono);font-size:0.56rem;color:#3a3830;text-transform:uppercase;letter-spacing:0.14em;margin-top:3px;}
.sb-sec{padding:10px 14px 5px;font-family:var(--mono);font-size:0.56rem;color:#3a3830;text-transform:uppercase;letter-spacing:0.14em;display:flex;align-items:center;justify-content:space-between;}
.cl-list{flex:1;overflow-y:auto;padding:4px 8px;}
.cl-list::-webkit-scrollbar{width:3px;}
.cl-list::-webkit-scrollbar-thumb{background:#222;border-radius:2px;}
.ci{display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:7px;cursor:pointer;transition:background 0.13s;margin-bottom:2px;position:relative;border:1px solid transparent;}
.ci:hover{background:#1c1a14;}
.ci.act{background:#221f0d;border-color:#332e0e;}
.ci-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.ci-type-badge{font-family:var(--mono);font-size:0.5rem;font-weight:700;padding:1px 5px;border-radius:3px;letter-spacing:0.06em;flex-shrink:0;}
.ci-type-att{background:#1a3a20;color:#6fcf97;}
.ci-type-marks{background:#1a1a3a;color:#818cf8;}
.ci-info{flex:1;min-width:0;}
.ci-name{font-family:var(--sans);font-size:0.78rem;font-weight:500;color:#ddd;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.ci-sub{font-family:var(--mono);font-size:0.58rem;color:#484440;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px;}
.ci-del{background:none;border:none;color:#2e2c24;cursor:pointer;font-size:0.9rem;padding:0 3px;line-height:1;transition:color 0.13s;flex-shrink:0;}
.ci-del:hover{color:var(--accent);}
.sb-foot{padding:10px 8px;border-top:1px solid #222018;flex-shrink:0;display:flex;flex-direction:column;gap:6px;}
.add-cls-btn{width:100%;font-family:var(--mono);font-size:0.7rem;font-weight:500;background:transparent;border:1px dashed #2e2c24;color:#555;padding:8px;border-radius:7px;cursor:pointer;transition:all 0.15s;letter-spacing:0.04em;}
.add-cls-btn:hover{border-color:var(--accent);color:var(--accent);background:#1a0e08;}
.exp-all-btn{width:100%;font-family:var(--mono);font-size:0.7rem;font-weight:500;background:#14633a;color:#a8ffc8;border:none;padding:8px;border-radius:7px;cursor:pointer;transition:background 0.15s;letter-spacing:0.04em;}
.exp-all-btn:hover{background:#197e4a;}
.save-badge{font-family:var(--mono);font-size:0.58rem;color:#3a3830;text-align:center;padding:4px 0 2px;letter-spacing:0.06em;}
.save-badge.saved{color:#2a6b46;}.save-badge.saving{color:#7a6a20;}

/* main */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;background:var(--bg);}
.no-class{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:14px;font-family:var(--mono);color:#888;font-size:0.82rem;}
.no-class-icon{font-size:3rem;opacity:0.25;}

/* info bar */
.ibar{background:var(--hdr);border-bottom:2px solid var(--accent);flex-shrink:0;}
.ibar.marks-mode{border-bottom-color:var(--purple);}
.ibar-row{display:flex;align-items:flex-end;gap:16px;padding:10px 16px 8px;flex-wrap:wrap;}
.ibar-f{display:flex;flex-direction:column;gap:3px;}
.ibar-lbl{font-family:var(--mono);font-size:0.55rem;color:#484440;text-transform:uppercase;letter-spacing:0.12em;}
.ibar-inp{font-family:var(--sans);font-size:0.78rem;font-weight:500;color:#f0ede6;background:transparent;border:none;outline:none;border-bottom:1px solid #2e2c26;padding:2px 0;transition:border-color 0.15s;}
.ibar-inp:focus{border-bottom-color:var(--accent2);}
.ibar-inp::placeholder{color:#2e2c26;}
.ibar-inp.w1{min-width:240px;}.ibar-inp.w2{min-width:170px;}.ibar-inp.w3{min-width:130px;}
.mode-pill{font-family:var(--mono);font-size:0.6rem;font-weight:700;padding:3px 10px;border-radius:20px;letter-spacing:0.08em;align-self:center;flex-shrink:0;}
.mode-pill.att{background:#1a3a20;color:#6fcf97;}
.mode-pill.marks{background:#2e1a5e;color:#a78bfa;}

/* topbar */
.topbar{background:var(--hdr2);display:flex;align-items:center;gap:8px;padding:7px 16px;flex-shrink:0;flex-wrap:wrap;}
.btn{font-family:var(--mono);font-size:0.7rem;font-weight:500;border:none;cursor:pointer;border-radius:5px;padding:6px 14px;transition:all 0.15s;white-space:nowrap;letter-spacing:0.02em;}
.btn-ghost{background:transparent;border:1px solid #3a3830;color:#888;}
.btn-ghost:hover{border-color:#666;color:#ddd;background:#24221c;}
.btn-accent{background:var(--accent);color:#fff;}
.btn-accent:hover{background:#bf4018;}
.btn-blue{background:var(--accent2);color:#fff;}
.btn-blue:hover{background:#1d52c4;}
.btn-green{background:#14633a;color:#a8ffc8;}
.btn-green:hover{background:#197e4a;}
.btn-amber{background:#d97706;color:#fff;}
.btn-amber:hover{background:#b45309;}
.btn-purple{background:var(--purple);color:#fff;}
.btn-purple:hover{background:#6d28d9;}
.btn-teal{background:#0891b2;color:#fff;}
.btn-teal:hover{background:#0e7490;}

/* subbar */
.subbar{background:var(--surface);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:9px;padding:7px 16px;flex-shrink:0;flex-wrap:wrap;}
.slabel{font-family:var(--mono);font-size:0.62rem;color:#999;text-transform:uppercase;letter-spacing:0.1em;white-space:nowrap;}
.inp{font-family:var(--mono);font-size:0.79rem;background:var(--bg);border:1px solid var(--border);color:#18160f;padding:5px 10px;border-radius:5px;outline:none;transition:border-color 0.15s;}
.inp:focus{border-color:var(--accent2);}
.inp::placeholder{color:#c0bdb5;}
.inp-r{width:130px;}.inp-s{width:145px;}.inp-n{width:240px;}.inp-d{width:146px;}
.sdiv{width:1px;height:22px;background:var(--border);margin:0 2px;}

/* searchbar */
.searchbar{background:#f5f2ec;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;padding:6px 16px;flex-shrink:0;}
.search-icon{font-size:0.85rem;color:#aaa;flex-shrink:0;}
.search-inp{flex:1;font-family:var(--mono);font-size:0.79rem;background:transparent;border:none;outline:none;color:#18160f;min-width:0;}
.search-inp::placeholder{color:#c0bdb5;}
.search-clear{font-family:var(--mono);font-size:0.68rem;color:#aaa;background:none;border:none;cursor:pointer;padding:2px 6px;border-radius:4px;}
.search-clear:hover{color:var(--accent);}
.search-count{font-family:var(--mono);font-size:0.62rem;color:#aaa;white-space:nowrap;}

/* statsbar */
.statsbar{background:var(--hdr);display:flex;align-items:center;gap:20px;padding:5px 16px;flex-shrink:0;font-family:var(--mono);font-size:0.66rem;flex-wrap:wrap;}
.stat{display:flex;gap:5px;align-items:center;}
.slbl{color:#484440;}.sv{font-weight:500;}
.sv.g{color:#6fcf97;}.sv.r{color:#ff7b6b;}.sv.y{color:#f2c94c;}.sv.am{color:#f59e0b;}.sv.pu{color:#a78bfa;}
.hint{margin-left:auto;font-size:0.57rem;color:#333;}

/* weight bar */
.wbar{background:#140f2a;border-bottom:1px solid #2a1a5e;display:flex;align-items:center;gap:12px;padding:6px 16px;flex-shrink:0;flex-wrap:wrap;}
.wbar-label{font-family:var(--mono);font-size:0.6rem;color:#6b5fa0;text-transform:uppercase;letter-spacing:0.1em;white-space:nowrap;flex-shrink:0;}
.wbar-pills{display:flex;gap:6px;flex-wrap:wrap;flex:1;}
.wpill{display:flex;align-items:center;gap:5px;background:#1e1240;border:1px solid #3a2a7e;border-radius:5px;padding:3px 8px;font-family:var(--mono);font-size:0.65rem;}
.wpill-name{color:#a78bfa;max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.wpill-inp{width:38px;background:#120d2a;border:1px solid #3a2a7e;border-radius:3px;color:#e0d4ff;font-family:var(--mono);font-size:0.65rem;padding:2px 4px;outline:none;text-align:center;}
.wpill-inp:focus{border-color:#a78bfa;}
.wpill-pct{color:#6b5fa0;font-size:0.6rem;}
.wbar-total{font-family:var(--mono);font-size:0.65rem;padding:3px 10px;border-radius:5px;white-space:nowrap;}
.wbar-total.ok{background:#0d2a18;color:#6fcf97;border:1px solid #14633a;}
.wbar-total.warn{background:#2a1a08;color:#f59e0b;border:1px solid #854F0B;}
.wbar-mode{font-family:var(--mono);font-size:0.6rem;color:#484440;text-transform:uppercase;letter-spacing:0.08em;white-space:nowrap;}
.wbar-toggle{font-family:var(--mono);font-size:0.62rem;background:#2e1a5e;color:#a78bfa;border:1px solid #534AB7;border-radius:4px;padding:3px 10px;cursor:pointer;transition:all 0.13s;white-space:nowrap;}
.wbar-toggle:hover{background:#3d2270;border-color:#7c3aed;}
.wbar-toggle.on{background:#3d2270;border-color:#7c3aed;color:#c4b5fd;}

/* sheet */
.sh-outer{flex:1;overflow:hidden;}
.sh-scroll{width:100%;height:100%;overflow:auto;}
.sh-scroll::-webkit-scrollbar{width:10px;height:10px;}
.sh-scroll::-webkit-scrollbar-track{background:var(--bg);}
.sh-scroll::-webkit-scrollbar-thumb{background:var(--border-s);border-radius:5px;}
.sh-scroll::-webkit-scrollbar-corner{background:var(--bg);}
.sh-tbl{border-collapse:collapse;table-layout:fixed;}

/* col headers */
.chr th{
  background:var(--hdr);color:#bbb;font-family:var(--mono);font-size:0.61rem;font-weight:500;
  text-transform:uppercase;letter-spacing:0.06em;border-right:1px solid #28261e;
  border-bottom:2px solid #34322a;padding:0 6px;height:40px;vertical-align:middle;
  position:sticky;top:0;z-index:10;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;user-select:none;
}
.chr th.fr{position:sticky;left:0;z-index:20;background:#121008;border-right:1px solid #34322a;text-align:center;width:var(--roll-w);min-width:var(--roll-w);max-width:var(--roll-w);}
.chr th.fs{position:sticky;left:var(--roll-w);z-index:20;background:#121008;border-right:1px solid #34322a;text-align:center;width:var(--sys-w);min-width:var(--sys-w);max-width:var(--sys-w);}
.chr th.fn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:20;background:#121008;border-right:2px solid var(--accent);width:var(--name-w);min-width:var(--name-w);max-width:var(--name-w);}
.chr th.fn.marks-mode{border-right-color:var(--purple);}
.chr th.ft{background:#1a1812;border-left:1px solid #34322a;text-align:center;color:var(--gold);width:110px;min-width:110px;max-width:110px;}
.chr th.fw{background:#110d22;border-left:1px solid #2a1a5e;text-align:center;color:#a78bfa;width:80px;min-width:80px;max-width:80px;}
.chr th.dc{text-align:center;width:${CELL_W}px;min-width:${CELL_W}px;max-width:${CELL_W}px;background:#1c1a14;}
.chr th.mc{text-align:center;width:${MARK_W}px;min-width:${MARK_W}px;max-width:${MARK_W}px;background:#1a1426;}
.chr th.mc:hover .xbtn{opacity:1;}
.chr th.dc:hover .xbtn{opacity:1;}
.dch{display:flex;flex-direction:column;align-items:center;gap:1px;position:relative;}
.dch-t{font-size:0.57rem;color:#888;}.dch-m{font-size:0.78rem;font-weight:500;color:#f0ede6;}
.dch-sub{font-size:0.55rem;color:#a78bfa;}
.dch-w{font-size:0.54rem;color:#7c5fa0;margin-top:1px;}
.xbtn{position:absolute;top:-4px;right:-10px;background:var(--accent);color:#fff;border:none;border-radius:50%;width:14px;height:14px;font-size:0.55rem;line-height:14px;text-align:center;cursor:pointer;opacity:0;transition:opacity 0.13s;padding:0;}

/* data rows */
.dr{height:34px;}
.dr:nth-child(even) td{background:var(--row-alt);}
.dr:nth-child(odd) td{background:var(--surface);}
.dr:hover td{background:#e8e4da!important;}
.dr td{border-right:1px solid var(--border);border-bottom:1px solid var(--border);padding:0;height:34px;vertical-align:middle;transition:background 0.1s;}

/* low-att highlight */
.dr.low-att td{background:#fff8e6!important;}
.dr.low-att:hover td{background:#fef0c0!important;}
.dr.low-att .tdr{background:#fde68a!important;color:#92400e!important;}
.dr.low-att .tds{background:#fde68a!important;color:#92400e!important;}
.dr.low-att .tdn{background:#fde68a!important;border-right-color:#f59e0b!important;}
.dr.low-att .tdt{background:#fef3c7!important;}
.low-badge{display:inline-block;background:#f59e0b;color:#fff;font-family:var(--mono);font-size:0.49rem;font-weight:700;padding:1px 5px;border-radius:3px;margin-left:5px;vertical-align:middle;letter-spacing:0.05em;}

/* frozen roll */
.tdr{position:sticky;left:0;z-index:5;background:var(--frozen)!important;border-right:1px solid var(--border-s)!important;font-family:var(--mono);font-size:0.72rem;color:#888;text-align:center;width:var(--roll-w);min-width:var(--roll-w);max-width:var(--roll-w);padding:0 4px;height:34px;line-height:34px;}
.dr:hover .tdr{background:#dedad2!important;}
.xrow{background:none;border:none;color:var(--accent);cursor:pointer;font-size:0.85rem;padding:0 2px;display:none;line-height:1;}
.dr:hover .xrow{display:inline;}

/* frozen system id */
.tds{position:sticky;left:var(--roll-w);z-index:5;background:var(--frozen)!important;border-right:1px solid var(--border-s)!important;font-family:var(--mono);font-size:0.72rem;color:#666;text-align:center;width:var(--sys-w);min-width:var(--sys-w);max-width:var(--sys-w);padding:0;height:34px;overflow:hidden;}
.dr:hover .tds{background:#dedad2!important;}
.sinp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:0.72rem;color:#18160f;text-align:center;padding:0 6px;display:block;}
.sinp:focus{background:#fff;box-shadow:inset 0 0 0 2px var(--accent2);}

/* frozen name */
.tdn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:5;background:var(--frozen)!important;border-right:2px solid var(--accent)!important;width:var(--name-w);min-width:var(--name-w);max-width:var(--name-w);padding:0;height:34px;display:flex;align-items:center;overflow:hidden;}
.tdn.marks-mode{border-right-color:var(--purple)!important;}
.dr:hover .tdn{background:#dedad2!important;}
.ninp{flex:1;height:34px;min-width:0;background:transparent;border:none;outline:none;font-family:var(--sans);font-size:0.81rem;font-weight:500;color:#18160f;padding:0 10px;display:block;}
.ninp:focus{background:#fff;box-shadow:inset 0 0 0 2px var(--accent2);}

/* totals / weighted */
.tdt{text-align:center;font-family:var(--mono);font-size:0.71rem;font-weight:500;background:#fdf8ec!important;border-left:1px solid var(--border-s)!important;width:110px;min-width:110px;max-width:110px;height:34px;}
.dr:hover .tdt{background:#f5edd4!important;}
.tdw{text-align:center;font-family:var(--mono);font-size:0.71rem;font-weight:500;background:#f3f0ff!important;border-left:1px solid #c4b5fd!important;width:80px;min-width:80px;max-width:80px;height:34px;}
.dr:hover .tdw{background:#ede9ff!important;}

.tda{text-align:center;width:${CELL_W}px;min-width:${CELL_W}px;max-width:${CELL_W}px;padding:0;height:34px;}
.ainp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:0.88rem;font-weight:500;text-align:center;text-transform:uppercase;cursor:text;transition:background 0.12s,color 0.12s;display:block;}
.ainp.ip{background:var(--p-bg)!important;color:var(--p-fg);}
.ainp.ia{background:var(--a-bg)!important;color:var(--a-fg);}
.ainp:focus{box-shadow:inset 0 0 0 2px var(--accent2);position:relative;z-index:2;}

/* marks cells */
.tmc{text-align:center;width:${MARK_W}px;min-width:${MARK_W}px;max-width:${MARK_W}px;padding:0;height:34px;}
.minp{width:100%;height:34px;background:transparent;border:none;outline:none;font-family:var(--mono);font-size:0.82rem;font-weight:500;text-align:center;cursor:text;color:#18160f;transition:background 0.12s;display:block;}
.minp:focus{box-shadow:inset 0 0 0 2px var(--purple);position:relative;z-index:2;background:#f3f0ff;}
.minp.full{color:#14633a;background:#d1f0e0!important;}
.minp.good{color:#1d4ed8;}
.minp.low-m{color:#be2c0a;background:#fde0db!important;}
.minp.absent-m{color:#9ca3af;background:#f3f4f6!important;}

/* summary row */
.sr td{background:#18160f!important;color:#ccc;font-family:var(--mono);font-size:0.64rem;border-top:2px solid #34322a;border-right:1px solid #28261e;border-bottom:none;height:36px;text-align:center;vertical-align:middle;position:sticky;bottom:0;z-index:5;}
.sr td.sfr{position:sticky;left:0;z-index:15;background:#0e0d08!important;color:#444;font-size:0.55rem;text-transform:uppercase;letter-spacing:0.1em;border-right:1px solid #34322a!important;width:var(--roll-w);min-width:var(--roll-w);}
.sr td.sfs{position:sticky;left:var(--roll-w);z-index:15;background:#0e0d08!important;color:#444;font-size:0.55rem;text-transform:uppercase;letter-spacing:0.1em;border-right:1px solid #34322a!important;width:var(--sys-w);min-width:var(--sys-w);}
.sr td.sfn{position:sticky;left:calc(var(--roll-w) + var(--sys-w));z-index:15;background:#0e0d08!important;color:#555;border-right:2px solid var(--accent)!important;text-align:left;padding:0 10px;font-size:0.62rem;width:var(--name-w);min-width:var(--name-w);}
.sr td.sfn.marks-mode{border-right-color:var(--purple)!important;}
.sr td.sft{background:#1a1812!important;border-left:1px solid #34322a!important;color:var(--gold);}
.sr td.sfw{background:#110d22!important;border-left:1px solid #2a1a5e!important;color:#a78bfa;}
.sp{color:#6fcf97;font-weight:500;}.sa{color:#ff7b6b;font-weight:500;}
.empty{padding:80px;text-align:center;font-family:var(--mono);font-size:0.78rem;color:#aaa;}

/* modals */
.mbg{position:fixed;inset:0;background:rgba(0,0,0,0.75);z-index:100;display:flex;align-items:center;justify-content:center;}
.modal{background:#1a1812;border:1px solid #2a2820;border-radius:12px;padding:28px;width:480px;max-width:96vw;box-shadow:0 24px 64px rgba(0,0,0,0.6);}
.modal.wide{width:560px;}
.modal.narrow{width:420px;}
.modal h2{font-family:var(--disp);font-size:1.12rem;color:#fff;margin-bottom:3px;}
.modal-sub{font-family:var(--mono);font-size:0.62rem;color:#484440;margin-bottom:20px;text-transform:uppercase;letter-spacing:0.1em;line-height:1.6;}
.mf{margin-bottom:13px;}
.ml{font-family:var(--mono);font-size:0.62rem;color:#666;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:5px;display:block;}
.mi{width:100%;font-family:var(--sans);font-size:0.85rem;background:#111009;border:1px solid #2a2820;color:#e8e2d9;padding:8px 12px;border-radius:6px;outline:none;transition:border-color 0.15s;}
.mi:focus{border-color:var(--accent2);}
.mi::placeholder{color:#2e2c24;}
.mi-sm{width:100px;}
.mrow{display:flex;gap:12px;}
.mact{display:flex;gap:10px;justify-content:flex-end;margin-top:22px;}
.mbtn-c{background:transparent;border:1px solid #2a2820;color:#555;padding:8px 18px;border-radius:6px;font-family:var(--mono);font-size:0.73rem;cursor:pointer;}
.mbtn-c:hover{border-color:#555;color:#aaa;}
.mbtn-ok{background:var(--accent);color:#fff;border:none;padding:8px 20px;border-radius:6px;font-family:var(--mono);font-size:0.73rem;font-weight:500;cursor:pointer;}
.mbtn-ok:hover{background:#c0421a;}
.mbtn-blue{background:var(--accent2);color:#fff;border:none;padding:8px 20px;border-radius:6px;font-family:var(--mono);font-size:0.73rem;font-weight:500;cursor:pointer;}
.mbtn-blue:hover{background:#1d52c4;}
.mbtn-purple{background:var(--purple);color:#fff;border:none;padding:8px 20px;border-radius:6px;font-family:var(--mono);font-size:0.73rem;font-weight:500;cursor:pointer;}
.mbtn-purple:hover{background:#6d28d9;}
.mbtn-green{background:#14633a;color:#a8ffc8;border:none;padding:8px 20px;border-radius:6px;font-family:var(--mono);font-size:0.73rem;font-weight:500;cursor:pointer;}
.mbtn-green:hover{background:#197e4a;}
.type-pick{display:flex;gap:10px;margin-bottom:18px;}
.type-opt{flex:1;padding:14px 10px;border-radius:8px;border:2px solid #2a2820;cursor:pointer;text-align:center;transition:all 0.15s;}
.type-opt:hover{border-color:#444;}
.type-opt.sel-att{border-color:#14633a;background:#0d2a18;}
.type-opt.sel-marks{border-color:var(--purple);background:#1a0f30;}
.type-opt-icon{font-size:1.5rem;margin-bottom:6px;}
.type-opt-label{font-family:var(--disp);font-size:0.9rem;color:#ddd;font-weight:700;}
.type-opt-desc{font-family:var(--mono);font-size:0.6rem;color:#555;margin-top:3px;}
.cls-pick-list{display:flex;flex-direction:column;gap:6px;max-height:240px;overflow-y:auto;margin-top:4px;}
.cls-pick-item{display:flex;align-items:center;gap:10px;padding:10px 12px;background:#111009;border:1px solid #2a2820;border-radius:7px;cursor:pointer;transition:border-color 0.13s;}
.cls-pick-item:hover{border-color:var(--accent2);}
.cls-pick-item.selected{border-color:var(--accent2);background:#0d1a2e;}
.cls-pick-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.cls-pick-info{flex:1;min-width:0;}
.cls-pick-name{font-family:var(--sans);font-size:0.82rem;font-weight:500;color:#ddd;}
.cls-pick-sub{font-family:var(--mono);font-size:0.6rem;color:#484440;margin-top:2px;}

/* import modal */
.drop-zone{border:2px dashed #2a2820;border-radius:8px;padding:32px 20px;text-align:center;cursor:pointer;transition:all 0.15s;background:#111009;}
.drop-zone:hover,.drop-zone.drag-over{border-color:var(--accent2);background:#0d1a2e;}
.drop-zone-icon{font-size:2rem;margin-bottom:8px;opacity:0.5;}
.drop-zone-text{font-family:var(--mono);font-size:0.75rem;color:#555;line-height:1.8;}
.drop-zone-hint{font-size:0.6rem;color:#333;margin-top:6px;}
.import-preview{background:#0e0d08;border:1px solid #2a2820;border-radius:6px;padding:10px 14px;margin-top:12px;max-height:180px;overflow-y:auto;}
.import-preview-row{display:flex;gap:12px;font-family:var(--mono);font-size:0.68rem;padding:3px 0;border-bottom:1px solid #1a1812;}
.import-preview-row:last-child{border-bottom:none;}
.ipr-roll{color:#a78bfa;width:70px;flex-shrink:0;}
.ipr-name{color:#e8e2d9;flex:1;}
.import-status{font-family:var(--mono);font-size:0.68rem;padding:8px 12px;border-radius:5px;margin-top:10px;text-align:center;}
.import-status.ok{background:#0d2a18;color:#6fcf97;border:1px solid #14633a;}
.import-status.warn{background:#2a1a08;color:#f59e0b;border:1px solid #854F0B;}
.import-status.err{background:#2a0e0e;color:#ff7b6b;border:1px solid #7a1a1a;}
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
  const [draft,            setDraft]            = useState({faculty:"",session:"",teacher:"",subject:"",course:"",type:"attendance"});
  const [assessDraft,      setAssessDraft]      = useState({name:"Quiz 1",outOf:"10",weight:"",type:"quiz"});
  const [importData,       setImportData]       = useState(null); // {rows:[{roll,name}], fileName, warnings:[]}
  const [importDragOver,   setImportDragOver]   = useState(false);

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
    if (active) setClasses(cs => cs.map(c => c.id === activeId ? fn(c) : c));
  };

  // ── create class ─────────────────────────────────────────────────────────────
  const createClass = () => {
    const cls = {
      id: uid(), color: CLASS_COLORS[classes.length % CLASS_COLORS.length],
      type: draft.type || "attendance",
      info: { faculty:draft.faculty, session:draft.session, teacher:draft.teacher, subject:draft.subject, course:draft.course },
      students:[], dates:[todayStr()], att:{}, assessments:[], marks:{},
    };
    setClasses(cs => [...cs, cls]);
    setActiveId(cls.id);
    setShowModal(false);
    setDraft({faculty:"",session:"",teacher:"",subject:"",course:"",type:"attendance"});
  };

  const deleteClass = (id, e) => {
    e.stopPropagation();
    if (!window.confirm("Delete this class and all its data?")) return;
    const remaining = classes.filter(c => c.id !== id);
    setClasses(remaining);
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
    updActive(c => ({
      ...c,
      assessments: (c.assessments??[]).map(a => a.id === editAssessId
        ? { ...a, name: assessDraft.name.trim()||a.name, outOf: parseFloat(assessDraft.outOf)||a.outOf, type: assessDraft.type||a.type, weight: parseFloat(assessDraft.weight)||0 }
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

  // Update weight inline from wbar
  const setAssessWeight = (aid, val) => {
    const w = val === "" ? 0 : parseFloat(val) || 0;
    updActive(c => ({ ...c, assessments: (c.assessments??[]).map(a => a.id===aid ? {...a, weight: w} : a) }));
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
    const infoRows=[[(inf?.faculty)||""],[`Marks Register (${(inf?.session)||""})`],[`Teacher's Name: ${(inf?.teacher)||""}`],[`Subject: ${(inf?.subject)||""}`],[`Course: ${(inf?.course)||""}`],[]];
    const outOfRow=["","","","Out Of",...asm.map(a=>a.outOf),"Total %","Weighted %"];
    const weightRow=["","","","Weight %",...asm.map(a=>a.weight?(a.weight+"%"):"—"),"",""];
    const hdr=["#","Student ID","System ID","Student Name",...asm.map(a=>`${a.name} (${a.type})`),"Total %","Weighted %"];
    const rows=(cls.students??[]).map((s,i)=>{ 
      const sm=stuMarksSum(cls,s.id); 
      const wm=hasWeights?stuWeightedSum(cls,s.id):null;
      return [i+1,s.roll,s.systemId||"",s.name,...asm.map(a=>getMark(cls,s.id,a.id)||""),sm?sm.pct+"%":"-",wm?wm.pct+"%":"-"]; 
    });
    const avgRow=["","","","Average",...asm.map(a=>assessColAvg(cls,a.id)??"-"),"",""];
    const ws=XLSX.utils.aoa_to_sheet([...infoRows,outOfRow,weightRow,hdr,...rows,[],avgRow]);
    ws["!cols"]=[{wch:4},{wch:12},{wch:14},{wch:28},...asm.map(()=>({wch:14})),{wch:10},{wch:12}];
    return ws;
  };
  const exportAll = () => {
    if(!classes.length)return;
    const wb=XLSX.utils.book_new(); const used={};
    classes.forEach(cls=>{ let sn=((cls.info?.subject)||(cls.info?.course)||"Class").replace(/[:\\/?*\[\]]/g,"").slice(0,24)||"Sheet"; const sfx=cls.type==="marks"?" Marks":" Att"; let fn=(sn+sfx).slice(0,31); let n=2; while(used[fn]){fn=`${sn}_${n++}`;} used[fn]=1; XLSX.utils.book_append_sheet(wb, cls.type==="marks"?buildMarksWs(cls):buildAttWs(cls), fn); });
    XLSX.writeFile(wb,"attendx_all_classes.xlsx");
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
            <div className="logo">Attend<span>X</span></div>
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
              <div className="ibar-row">
                <span className={`mode-pill ${isMarks?"marks":"att"}`}>{isMarks?"📊 MARKS":"✅ ATTENDANCE"}</span>
                {[["faculty","Faculty / Department","Faculty of Engineering & Computing","w1"],["session","Session","Feb 2026 – Jun 2026","w2"],["teacher","Teacher","Full name","w3"],["subject","Subject","Subject","w3"],["course","Course","BSSE 39 M 1 (B)","w3"]].map(([f,lbl,ph,w])=>(
                  <div key={f} className="ibar-f" style={{flex:w==="w1"?"2 1 180px":"1 1 120px"}}>
                    <span className="ibar-lbl">{lbl}</span>
                    <input className={`ibar-inp ${w}`} placeholder={ph} value={active.info?.[f] ?? ""} onChange={e=>setInfo(f,e.target.value)}/>
                  </div>
                ))}
              </div>
            </div>

            {/* topbar */}
            <div className="topbar">
              {isAtt&&<>
                <button className="btn btn-ghost" onClick={()=>{if(window.confirm("Clear all attendance?"))updActive(c=>({...c,att:{}}));}}>Clear</button>
                <button className="btn btn-green" onClick={markRestP}>✓ Mark Rest Present</button>
                <button className="btn btn-ghost" onClick={markRestA}>✗ Mark Rest Absent</button>
                <button className={`btn ${lowAttOnly?"btn-amber":"btn-ghost"}`} onClick={()=>setLowAttOnly(v=>!v)}>⚠ Below 25%</button>
              </>}
              {isMarks&&<button className="btn btn-ghost" onClick={()=>{if(window.confirm("Clear all marks?"))updActive(c=>({...c,marks:{}}));}}>Clear Marks</button>}
              {classes.length>1&&<button className="btn btn-amber" onClick={openCopyModal}>⇄ Copy Roster</button>}
              <div style={{flex:1}}/>
              <button className="btn btn-accent" onClick={exportOne}>↓ Export This Class</button>
            </div>

            {/* subbar */}
            <div className="subbar">
              <span className="slabel">Add Student</span>
              <input className="inp inp-r" placeholder="Student ID" value={newRoll} onChange={e=>setNewRoll(e.target.value)} onKeyDown={onAddKey}/>
              <input className="inp inp-s" placeholder="System ID" value={newSystemId} onChange={e=>setNewSystemId(e.target.value)} onKeyDown={onAddKey}/>
              <input className="inp inp-n" placeholder="Name (or Ali, Sara, Ahmed…)" value={newName} onChange={e=>setNewName(e.target.value)} onKeyDown={onAddKey}/>
              <button className="btn btn-accent" onClick={addStudent}>+ Add Row</button>
              <div className="sdiv"/>
              {/* IMPORT BUTTON */}
              <button className="btn btn-teal" onClick={()=>{setImportData(null);setShowImportModal(true);}}>↑ Import from Excel</button>
              <div className="sdiv"/>
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
                <div className={`wbar-total ${weightsOk?"ok":"warn"}`}>
                  Σ {totalWeight.toFixed(1)}% {weightsOk?"✓":"≠ 100"}
                </div>
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
                              <span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:3}}>
                                <span>{s.roll}</span>
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
                                ? <span><span className="sp">{sm.p}P</span>/<span className="sa">{sm.a}A</span> <span style={{color:isLow?"#f59e0b":"#aaa",fontSize:"0.61rem",fontWeight:isLow?700:400}}>({sm.pct}%)</span></span>
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
                              <span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:3}}>
                                <span>{s.roll}</span>
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
                                ? <span style={{color:"#7c3aed",fontWeight:600}}>{sm.earned}/{sm.total} <span style={{color:"#aaa",fontSize:"0.61rem"}}>({sm.pct}%)</span></span>
                                : <span style={{color:"#ccc"}}>—</span>}
                            </td>
                            {weightingOn && (
                              <td className="tdw">
                                {wm
                                  ? <span style={{color: wm.pct>=90?"#14633a":wm.pct>=60?"#1d4ed8":"#be2c0a", fontWeight:600}}>
                                      {wm.earned}<span style={{color:"#aaa",fontSize:"0.58rem",fontWeight:400}}>/{wm.total}</span>
                                      <span style={{display:"block",fontSize:"0.58rem",color:"#aaa",fontWeight:400}}>{wm.pct}%</span>
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
    </>
  );
}