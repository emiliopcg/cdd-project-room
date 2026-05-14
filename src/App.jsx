import { useState, useEffect, useRef } from “react”;
import * as XLSX from “xlsx”;
import Papa from “papaparse”;
import { RadarChart, Radar, PolarGrid, PolarAngleAxis, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, Cell, LineChart, Line } from “recharts”;

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const APP_KEY = “cdd_projects_v1”;
const COLORS = [”#E0294A”,”#3b82f6”,”#22c55e”,”#f59e0b”,”#a855f7”,”#06b6d4”,”#f97316”];
const SECTORS = [“Technology / SaaS”,“Healthcare / MedTech”,“Financial Services / Fintech”,“Consumer / Retail”,“Industrials / Manufacturing”,“Real Estate / PropTech”,“Dental / Healthcare Services”,“Education / EdTech”,“Logistics / Supply Chain”,“Food & Beverage”,“Business Services / BPO”,“Other”];
const GEOGRAPHIES = [“España”,“Europa Occidental”,“Europa (total)”,“UK”,“DACH”,“Nórdicos”,“LATAM”,“USA”,“Global”];
const DEAL_TYPES = [“Buyout (LBO)”,“Growth Equity”,“Minority Stake”,“Add-on / Bolt-on”,“Carve-out”,“IPO Preparation”];
const SIGNAL_COLORS = { Positive:”#22c55e”, Neutral:”#888”, Negative:”#E0294A”, Mixed:”#f59e0b”, “Too Early”:”#888” };
const HYP_STATUS = { pending: { color:”#f59e0b”, label:“Pendiente”, icon:“◌” }, validated: { color:”#22c55e”, label:“Validada”, icon:“✓” }, refuted: { color:”#E0294A”, label:“Refutada”, icon:“✗” }, investigating: { color:”#3b82f6”, label:“En análisis”, icon:“◎” } };

// ─── CLAUDE ───────────────────────────────────────────────────────────────────
const claude = async (system, user) => {
const r = await fetch(“https://urldefense.com/v3/__https://api.anthropic.com/v1/messages__;!!Nyu6ZXf5!sq6YgluEaR9FbFa9FRK8gCQWMRMQf3HSsfUvJtRzU5ULLbCdd7uwjSj9QnSY2cbWnKP2Dj7Y4mzgFZxzaiaXeBca3tx6GnCd$ ”, {
method: “POST”, headers: { “Content-Type”: “application/json” },
body: JSON.stringify({ model: “claude-sonnet-4-20250514”, max_tokens: 1000, system, messages: [{ role: “user”, content: user }] }),
});
const d = await r.json();
const t = d.content?.map(i => i.text || “”).join(””) || “”;
return JSON.parse(t.replace(/`json|`/g, “”).trim());
};

// ─── STORAGE ──────────────────────────────────────────────────────────────────
const loadProjects = async () => {
try {
const r = await window.storage.get(APP_KEY);
return r ? JSON.parse(r.value) : {};
} catch { return {}; }
};
const saveProjects = async (projects) => {
try { await window.storage.set(APP_KEY, JSON.stringify(projects)); } catch {}
};

// ─── STYLES ───────────────────────────────────────────────────────────────────
const css = {
inp: { width:“100%”, background:”#161616”, border:“1px solid #2a2a2a”, borderRadius:“3px”, color:”#c8c4be”, fontFamily:“monospace”, fontSize:“13px”, padding:“9px 13px”, outline:“none”, transition:“border-color 0.2s”, boxSizing:“border-box” },
lbl: { fontSize:“10px”, letterSpacing:“3px”, color:”#555”, textTransform:“uppercase”, marginBottom:“6px”, fontFamily:“monospace”, display:“block” },
card: (extra={}) => ({ background:”#141414”, border:“1px solid #1e1e1e”, borderRadius:“4px”, padding:“18px 22px”, …extra }),
fr: e => (e.target.style.borderColor=”#E0294A”),
bg: e => (e.target.style.borderColor=”#2a2a2a”),
secLabel: { fontSize:“10px”, letterSpacing:“3px”, color:”#555”, textTransform:“uppercase”, marginBottom:“14px”, fontFamily:“monospace” },
};

// ─── SHARED COMPONENTS ────────────────────────────────────────────────────────
const Btn = ({ onClick, disabled, loading, children, variant=“primary”, style={} }) => {
const bg = variant === “ghost” ? “none” : (disabled||loading ? “#2a2a2a” : “#E0294A”);
const col = variant === “ghost” ? “#888” : (disabled||loading ? “#555” : “#fff”);
const bdr = variant === “ghost” ? “1px solid #2a2a2a” : “none”;
return <button onClick={onClick} disabled={disabled||loading} style={{ background:bg, color:col, border:bdr, padding:“10px 28px”, fontSize:“11px”, letterSpacing:“3px”, textTransform:“uppercase”, fontFamily:“monospace”, cursor:(disabled||loading)?“not-allowed”:“pointer”, borderRadius:“2px”, display:“flex”, alignItems:“center”, gap:“8px”, transition:“all 0.2s”, …style }}>
{loading ? <><span style={{animation:“spin 1s linear infinite”,display:“inline-block”}}>◌</span>Procesando…</> : children}
</button>;
};

const Tabs = ({ tabs, active, onChange, accentColor=”#E0294A” }) => (

  <div style={{ display:"flex", gap:"2px", borderBottom:"1px solid #1e1e1e", overflowX:"auto" }}>
    {tabs.map(t => <button key={t.id} onClick={() => onChange(t.id)} style={{ background:active===t.id?"#141414":"none", border:"none", borderBottom:active===t.id?`2px solid ${accentColor}`:"2px solid transparent", color:active===t.id?"#f0ede8":"#555", padding:"10px 18px", fontSize:"11px", letterSpacing:"1.5px", cursor:"pointer", fontFamily:"monospace", whiteSpace:"nowrap" }}>{t.label}</button>)}
  </div>
);

const TipBox = ({ content }) => <div style={{ background:”#1a1a1a”, border:“1px solid #333”, padding:“9px 13px”, borderRadius:“3px”, fontFamily:“monospace”, fontSize:“12px”, color:”#888” }}>{content}</div>;

// ─── MODULE PROMPTS ───────────────────────────────────────────────────────────
const PROMPTS = {
kickoff: (p) => ({
sys: `You are a senior Strategy& CDD expert. Generate a full project kickoff package. Respond ONLY in JSON: {"deal_signal":"Positive|Neutral|Negative|Too Early","investment_thesis":"string","hypotheses":[{"id":number,"workstream":"Market|Competitive|Customer|BP|Risk","text":"string","why_critical":"string","status":"pending"}],"data_needed":["string"],"expert_profiles":[{"profile":"string","why":"string","questions":["string"]}],"week_plan":[{"week":"string","deliverables":["string"]}],"red_flags":["string"],"initial_market_size":"string"}`,
usr: `Generate CDD kickoff for: ${p.company} | ${p.sector} | ${p.dealType} | ${p.geography}${p.revenue?` | Revenue: ${p.revenue}`:""}${p.description?`\nDescription: ${p.description}`:""}`
}),
market: (p, ctx) => ({
sys: `You are a senior Strategy& consultant. Perform dual market sizing (top-down + bottom-up). Respond ONLY in JSON: {"market_title":"string","tam":{"value":number,"definition":"string"},"sam":{"value":number,"definition":"string"},"som":{"value":number,"definition":"string"},"cagr_historical":"string","cagr_forward":"string","top_down":{"result":number,"steps":[{"step":number,"label":"string","value":"string","assumption":"string"}]},"bottom_up":{"result":number,"steps":[{"step":number,"label":"string","value":"string","assumption":"string"}]},"recommended_estimate":number,"gap_pct":"string","growth_drivers":[{"driver":"string","impact":"High|Medium|Low"}],"risks":[{"risk":"string","impact":"High|Medium|Low"}]}`,
usr: `Market sizing for ${p.company} in ${p.sector}, ${p.geography}.${ctx?` Context: ${ctx}`:""}`
}),
benchmark: (companies, sector) => ({
sys: `You are a senior Strategy& CDD consultant analyzing competitive landscape. Respond ONLY in JSON: {"landscape_narrative":"string","competitive_intensity":number,"insights":["string"],"positioning":{"leaders":["string"],"challengers":["string"],"niche":["string"]},"target_strengths":["string"],"target_vulnerabilities":["string"],"strategic_implications":[{"implication":"string","detail":"string"}],"radar_scores":[{"company":"string","scale":number,"growth":number,"profitability":number,"market_position":number,"customer_value":number}]}`,
usr: `Analyze competitive landscape for sector: ${sector}\n\nCompanies:\n${companies.map(c=>`${c.name}${c.isTarget?” [TARGET]”:””}: Revenue ${c.revenue||“N/A”}, EBITDA margin ${c.ebitda_margin||“N/A”}%, Growth ${c.growth||“N/A”}%`).join("\n")}`
}),
bp: (bpText, company, sector) => ({
sys: `You are a senior Strategy& CDD consultant stress-testing a business plan. Respond ONLY in JSON: {"credibility_score":number,"credibility_label":"string","credibility_summary":"string","make_or_break":[{"assumption":"string","why":"string","ebitda_impact":"string"}],"assumptions":[{"id":number,"name":"string","category":"string","management_claim":"string","benchmark":"string","verdict":"Aggressive|Reasonable|Conservative|Unsubstantiated","challenge_question":"string","deal_impact":"High|Medium|Low"}],"scenarios":{"bear":{"revenue":"string","ebitda":"string","rationale":"string"},"base":{"revenue":"string","ebitda":"string","rationale":"string"},"bull":{"revenue":"string","ebitda":"string","rationale":"string"}},"challenge_memo":"string"}`,
usr: `Stress-test this business plan for ${company} (${sector}):\n\n${bpText}`
}),
expert: (transcript) => ({
sys: `You are a senior Strategy& CDD consultant analyzing an expert call. Extract structured insights by workstream. Respond ONLY in JSON: {"expert_name":"string","expert_profile":"string","credibility_score":number,"credibility_rationale":"string","deal_signal":"Positive|Neutral|Negative|Mixed","deal_signal_rationale":"string","workstreams":{"market":{"insights":["string"],"so_what":"string"},"competitive":{"insights":["string"],"so_what":"string"},"customer":{"insights":["string"],"so_what":"string"},"bp":{"insights":["string"],"so_what":"string"},"risk":{"insights":["string"],"so_what":"string"},"quotes":{"insights":["string"],"so_what":"string"}},"follow_up_questions":["string","string","string"]}`,
usr: `Analyze this expert call transcript:\n\n${transcript}`
}),
storyboard: (p, analyses) => ({
sys: `You are a senior Strategy& CDD consultant building a slide deck storyboard. Create a narrative-driven structure. Respond ONLY in JSON: {"deal_verdict":"string","overall_signal":"Positive|Neutral|Negative|Mixed","exec_summary":{"investment_thesis":"string","top_findings":["string"],"key_risks":["string"],"recommendation":"string"},"sections":[{"id":"string","title":"string","narrative_role":"string","slides":[{"title":"string","message":"string","chart_type":"string","data_hint":"string"}]}],"slide_count":number}`,
usr: `Build CDD storyboard for ${p.company} (${p.sector}, ${p.dealType}).\n\nAvailable analyses: ${analyses.join(", ")}\n${p.description?`Company: ${p.description}`:""}`
}),
};

// ─── PARSE FILE ───────────────────────────────────────────────────────────────
const parseFile = (file) => new Promise((resolve, reject) => {
const ext = file.name.split(”.”).pop().toLowerCase();
if (ext === “csv”) {
Papa.parse(file, { header:true, skipEmptyLines:true, complete: r => resolve({ rows:r.data, columns:r.meta.fields||[] }), error:reject });
} else {
const reader = new FileReader();
reader.onload = e => {
const wb = XLSX.read(e.target.result, { type:“array” });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { defval:”” });
resolve({ rows, columns: rows.length>0 ? Object.keys(rows[0]) : [] });
};
reader.onerror = reject;
reader.readAsArrayBuffer(file);
}
});

// ─── HOME SCREEN ──────────────────────────────────────────────────────────────
const HomeScreen = ({ projects, onCreate, onOpen, onDelete }) => {
const [newName, setNewName] = useState(””);
return (
<div style={{ minHeight:“100vh”, background:”#0a0a0a”, fontFamily:”‘Georgia’,‘Times New Roman’,serif”, color:”#f0ede8” }}>
<div style={{ borderBottom:“1px solid #1a1a1a”, padding:“32px 48px”, background:”#0e0e0e” }}>
<div style={{ fontSize:“10px”, letterSpacing:“5px”, color:”#E0294A”, fontFamily:“monospace”, marginBottom:“10px” }}>STRATEGY& · DEALS</div>
<h1 style={{ margin:0, fontSize:“34px”, fontWeight:“400”, letterSpacing:”-1px” }}>CDD Project Room</h1>
<p style={{ margin:“10px 0 0”, fontSize:“13px”, color:”#666”, maxWidth:“500px” }}>Suite completa de análisis para Deals Strategy. Cada proyecto guarda todo tu trabajo automáticamente.</p>
</div>
<div style={{ padding:“40px 48px”, maxWidth:“1100px” }}>
{/* Create new */}
<div style={css.card({ marginBottom:“32px”, borderColor:”#E0294A30” })}>
<div style={css.secLabel}>Nuevo proyecto</div>
<div style={{ display:“flex”, gap:“12px”, alignItems:“center” }}>
<input value={newName} onChange={e=>setNewName(e.target.value)} onKeyDown={e=>e.key===“Enter”&&newName.trim()&&onCreate(newName.trim())} placeholder=“Nombre del proyecto — Ej: Project Dental Premium, Project SaaS Alpha…” style={{ …css.inp, flex:1, fontSize:“14px”, padding:“12px 16px” }} onFocus={css.fr} onBlur={css.bg} />
<Btn onClick={() => { if(newName.trim()) { onCreate(newName.trim()); setNewName(””); } }} disabled={!newName.trim()}>+ Crear proyecto</Btn>
</div>
</div>

```
    {/* Projects list */}
    {Object.keys(projects).length === 0 ? (
      <div style={{ textAlign:"center", padding:"60px 0", color:"#444" }}>
        <div style={{ fontSize:"40px", marginBottom:"16px" }}>◈</div>
        <div style={{ fontSize:"14px", fontFamily:"monospace" }}>No hay proyectos todavía. Crea el primero.</div>
      </div>
    ) : (
      <div>
        <div style={css.secLabel}>Proyectos activos — {Object.keys(projects).length}</div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"12px" }}>
          {Object.entries(projects).map(([id, p]) => {
            const modules = ["kickoff","market","benchmark","bp","experts","storyboard"].filter(m => p[m]);
            const completion = Math.round((modules.length / 6) * 100);
            return (
              <div key={id} style={css.card({ cursor:"pointer", transition:"all 0.2s", position:"relative" })}
                onMouseEnter={e=>{e.currentTarget.style.borderColor="#E0294A";e.currentTarget.style.background="#181414";}}
                onMouseLeave={e=>{e.currentTarget.style.borderColor="#1e1e1e";e.currentTarget.style.background="#141414";}}>
                <div onClick={()=>onOpen(id)} style={{ paddingBottom:"10px" }}>
                  <div style={{ fontSize:"10px", color:"#E0294A", fontFamily:"monospace", letterSpacing:"2px", marginBottom:"6px" }}>{p.dealType || "DEAL TYPE TBD"}</div>
                  <div style={{ fontSize:"17px", color:"#f0ede8", marginBottom:"4px" }}>{p.name}</div>
                  <div style={{ fontSize:"12px", color:"#666", marginBottom:"12px" }}>{p.company ? `${p.company} · ` : ""}{p.sector || "Sector TBD"}</div>
                  <div style={{ display:"flex", flexWrap:"wrap", gap:"5px", marginBottom:"12px" }}>
                    {["kickoff","market","benchmark","bp","experts","storyboard"].map(m => <span key={m} style={{ fontSize:"9px", fontFamily:"monospace", color:p[m]?"#22c55e":"#333", background:p[m]?"rgba(34,197,94,0.1)":"#111", padding:"2px 7px", borderRadius:"2px", border:`1px solid ${p[m]?"#22c55e30":"#2a2a2a"}` }}>{m}</span>)}
                  </div>
                  <div style={{ height:"3px", background:"#1a1a1a", borderRadius:"2px" }}>
                    <div style={{ height:"100%", width:`${completion}%`, background:"#E0294A", borderRadius:"2px", transition:"width 0.5s" }} />
                  </div>
                  <div style={{ fontSize:"10px", color:"#555", fontFamily:"monospace", marginTop:"5px" }}>{completion}% completo</div>
                </div>
                <button onClick={e=>{e.stopPropagation();onDelete(id);}} style={{ position:"absolute", top:"14px", right:"14px", background:"none", border:"none", color:"#333", cursor:"pointer", fontSize:"14px", padding:"2px 6px" }} onMouseEnter={e=>e.target.style.color="#E0294A"} onMouseLeave={e=>e.target.style.color="#333"}>✕</button>
              </div>
            );
          })}
        </div>
      </div>
    )}

    {/* Module legend */}
    <div style={{ marginTop:"40px", borderTop:"1px solid #1a1a1a", paddingTop:"24px" }}>
      <div style={css.secLabel}>Módulos disponibles</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"8px" }}>
        {[["◈","Kickoff","Hipótesis, work plan, expertos a llamar"],["◎","Market Sizing","Top-down + bottom-up + sensitivity"],["⊞","Benchmark","KPI table, radar, positioning"],["▲","BP Stress-Test","Challenge memo, scenarios Bear/Base/Bull"],["◉","Expert Calls","Síntesis por workstream CDD"],["◆","Storyboard","Narrativa del deck + exec summary"]].map(([icon,name,desc])=>(
          <div key={name} style={{ display:"flex", gap:"12px", padding:"12px", background:"#111", borderRadius:"3px", border:"1px solid #1a1a1a" }}>
            <span style={{ color:"#E0294A", fontSize:"16px", flexShrink:0 }}>{icon}</span>
            <div><div style={{ fontSize:"12px", color:"#f0ede8", marginBottom:"2px" }}>{name}</div><div style={{ fontSize:"11px", color:"#555" }}>{desc}</div></div>
          </div>
        ))}
      </div>
    </div>
  </div>
</div>
```

);
};

// ─── PROJECT ROOM ─────────────────────────────────────────────────────────────
const ProjectRoom = ({ project, onUpdate, onBack }) => {
const [activeModule, setActiveModule] = useState(“setup”);
const p = project;

const save = (key, data) => onUpdate({ …p, [key]: data, updatedAt: Date.now() });

const MODULES = [
{ id:“setup”, icon:“⚙”, label:“Setup” },
{ id:“kickoff”, icon:“◈”, label:“Kickoff” },
{ id:“data”, icon:“⬆”, label:“Data” },
{ id:“market”, icon:“◎”, label:“Market Sizing” },
{ id:“benchmark”, icon:“⊞”, label:“Benchmark” },
{ id:“bp”, icon:“▲”, label:“BP Stress-Test” },
{ id:“experts”, icon:“◉”, label:“Expert Calls” },
{ id:“hypotheses”, icon:“◌”, label:“Hypothesis Log” },
{ id:“storyboard”, icon:“◆”, label:“Storyboard” },
{ id:“export”, icon:“⬇”, label:“Export” },
];

return (
<div style={{ minHeight:“100vh”, background:”#0a0a0a”, fontFamily:”‘Georgia’,‘Times New Roman’,serif”, color:”#f0ede8”, display:“flex”, flexDirection:“column” }}>
{/* Top bar */}
<div style={{ background:”#0e0e0e”, borderBottom:“1px solid #1a1a1a”, padding:“14px 24px”, display:“flex”, alignItems:“center”, gap:“16px”, flexShrink:0 }}>
<button onClick={onBack} style={{ background:“none”, border:“none”, color:”#555”, cursor:“pointer”, fontFamily:“monospace”, fontSize:“13px”, padding:“2px 0” }}>← Proyectos</button>
<div style={{ width:“1px”, height:“16px”, background:”#2a2a2a” }} />
<div style={{ fontSize:“10px”, letterSpacing:“3px”, color:”#E0294A”, fontFamily:“monospace” }}>STRATEGY& · CDD</div>
<div style={{ width:“1px”, height:“16px”, background:”#2a2a2a” }} />
<div style={{ fontSize:“14px”, color:”#f0ede8” }}>{p.name}</div>
{p.company && <div style={{ fontSize:“12px”, color:”#666” }}>{p.company} · {p.sector}</div>}
<div style={{ marginLeft:“auto”, display:“flex”, gap:“6px” }}>
{MODULES.filter(m=>m.id!==“setup”).map(m => <button key={m.id} onClick={()=>setActiveModule(m.id)} title={m.label} style={{ background:activeModule===m.id?“rgba(224,41,74,0.15)”:“none”, border:`1px solid ${activeModule===m.id?"#E0294A40":"#1a1a1a"}`, color:activeModule===m.id?”#E0294A”:”#444”, padding:“5px 8px”, fontSize:“12px”, cursor:“pointer”, borderRadius:“2px”, fontFamily:“monospace” }}>
{p[m.id] ? <span style={{ color:”#22c55e” }}>✓</span> : m.icon}
</button>)}
</div>
</div>

```
  <div style={{ display:"flex", flex:1, overflow:"hidden" }}>
    {/* Sidebar */}
    <div style={{ width:"200px", background:"#0e0e0e", borderRight:"1px solid #1a1a1a", flexShrink:0, overflowY:"auto" }}>
      {MODULES.map(m => (
        <button key={m.id} onClick={()=>setActiveModule(m.id)} style={{ width:"100%", background:activeModule===m.id?"#141414":"none", border:"none", borderLeft:activeModule===m.id?"3px solid #E0294A":"3px solid transparent", borderBottom:"1px solid #111", color:activeModule===m.id?"#f0ede8":"#555", padding:"12px 14px", cursor:"pointer", textAlign:"left", transition:"all 0.15s" }}
          onMouseEnter={e=>{if(activeModule!==m.id){e.currentTarget.style.background="#111";e.currentTarget.style.color="#888";}}}
          onMouseLeave={e=>{if(activeModule!==m.id){e.currentTarget.style.background="none";e.currentTarget.style.color="#555";}}}>
          <div style={{ display:"flex", gap:"8px", alignItems:"center" }}>
            <span style={{ color:p[m.id]?"#22c55e":activeModule===m.id?"#E0294A":"#444", fontSize:"13px" }}>{p[m.id]&&m.id!=="setup"?"✓":m.icon}</span>
            <span style={{ fontSize:"12px" }}>{m.label}</span>
          </div>
        </button>
      ))}
    </div>

    {/* Main */}
    <div style={{ flex:1, overflowY:"auto", padding:"28px 36px" }}>
      {activeModule === "setup" && <SetupModule project={p} onSave={data=>onUpdate({...p,...data})} />}
      {activeModule === "kickoff" && <KickoffModule project={p} onSave={r=>save("kickoff",r)} />}
      {activeModule === "data" && <DataModule project={p} onSave={r=>save("data",r)} />}
      {activeModule === "market" && <MarketModule project={p} onSave={r=>save("market",r)} />}
      {activeModule === "benchmark" && <BenchmarkModule project={p} onSave={r=>save("benchmark",r)} />}
      {activeModule === "bp" && <BPModule project={p} onSave={r=>save("bp",r)} />}
      {activeModule === "experts" && <ExpertsModule project={p} onSave={r=>save("experts",r)} />}
      {activeModule === "hypotheses" && <HypothesesModule project={p} onSave={r=>save("hypotheses",r)} />}
      {activeModule === "storyboard" && <StoryboardModule project={p} onSave={r=>save("storyboard",r)} />}
      {activeModule === "export" && <ExportModule project={p} />}
    </div>
  </div>
</div>
```

);
};

// ─── SETUP MODULE ─────────────────────────────────────────────────────────────
const SetupModule = ({ project:p, onSave }) => {
const [form, setForm] = useState({ company:p.company||””, description:p.description||””, sector:p.sector||””, dealType:p.dealType||””, geography:p.geography||””, revenue:p.revenue||””, ebitda:p.ebitda||””, notes:p.notes||”” });
const u = (k,v) => setForm(f=>({…f,[k]:v}));
return (
<div>
<div style={css.secLabel}>Configuración del proyecto — {p.name}</div>
<div style={{ display:“grid”, gridTemplateColumns:“1fr 1fr”, gap:“14px”, marginBottom:“14px” }}>
<div style={{ gridColumn:“1/-1” }}><label style={css.lbl}>Empresa target *</label><input value={form.company} onChange={e=>u(“company”,e.target.value)} placeholder=“Nombre de la empresa” style={css.inp} onFocus={css.fr} onBlur={css.bg} /></div>
<div style={{ gridColumn:“1/-1” }}><label style={css.lbl}>Descripción del negocio</label><textarea value={form.description} onChange={e=>u(“description”,e.target.value)} placeholder=“Breve descripción del negocio, modelo, presencia geográfica…” style={{…css.inp,minHeight:“80px”,resize:“vertical”}} onFocus={css.fr} onBlur={css.bg} /></div>
{[[“sector”,“Sector”,SECTORS],[“dealType”,“Tipo de deal”,DEAL_TYPES],[“geography”,“Geografía”,GEOGRAPHIES]].map(([key,label,opts])=>(
<div key={key}><label style={css.lbl}>{label}</label><select value={form[key]} onChange={e=>u(key,e.target.value)} style={{…css.inp,cursor:“pointer”}} onFocus={css.fr} onBlur={css.bg}><option value="">Seleccionar…</option>{opts.map(o=><option key={o} value={o}>{o}</option>)}</select></div>
))}
<div><label style={css.lbl}>Revenue estimado</label><input value={form.revenue} onChange={e=>u(“revenue”,e.target.value)} placeholder=“Ej: €45M” style={css.inp} onFocus={css.fr} onBlur={css.bg} /></div>
<div><label style={css.lbl}>EBITDA estimado</label><input value={form.ebitda} onChange={e=>u(“ebitda”,e.target.value)} placeholder=“Ej: €8M (18%)” style={css.inp} onFocus={css.fr} onBlur={css.bg} /></div>
<div style={{ gridColumn:“1/-1” }}><label style={css.lbl}>Notas del equipo</label><textarea value={form.notes} onChange={e=>u(“notes”,e.target.value)} placeholder=“Contexto interno, sensibilidades del cliente, foco del análisis…” style={{…css.inp,minHeight:“70px”,resize:“vertical”}} onFocus={css.fr} onBlur={css.bg} /></div>
</div>
<Btn onClick={()=>onSave(form)}>Guardar configuración</Btn>
</div>
);
};

// ─── KICKOFF MODULE ───────────────────────────────────────────────────────────
const KickoffModule = ({ project:p, onSave }) => {
const [result, setResult] = useState(p.kickoff || null);
const [loading, setLoading] = useState(false);
const [tab, setTab] = useState(“hypotheses”);

const run = async () => {
if (!p.company || !p.sector) return alert(“Completa Setup primero (empresa + sector).”);
setLoading(true);
try { const r = await claude(PROMPTS.kickoff(p).sys, PROMPTS.kickoff(p).usr); setResult(r); onSave(r); }
catch { alert(“Error generando kickoff.”); }
finally { setLoading(false); }
};

if (!result) return (
<div>
<div style={css.secLabel}>Project Kickoff Generator</div>
<div style={{…css.card({marginBottom:“18px”,background:”#111”})}}>
<div style={{ fontSize:“13px”, color:”#888”, lineHeight:“1.7” }}>Genera automáticamente: hipótesis del deal, work plan semana a semana, perfiles de expertos a llamar con preguntas, datos que necesitas, y red flags iniciales — todo basado en el setup del proyecto.</div>
</div>
{!p.company && <div style={{ color:”#f59e0b”, fontFamily:“monospace”, fontSize:“12px”, marginBottom:“14px” }}>⚠ Completa el Setup del proyecto primero (empresa + sector).</div>}
<Btn onClick={run} loading={loading} disabled={!p.company}>◈ Generar kickoff completo</Btn>
</div>
);

const TABS = [{id:“hypotheses”,label:“◌ Hipótesis”},{id:“experts”,label:“◉ Expertos”},{id:“workplan”,label:“◆ Work Plan”},{id:“flags”,label:“▲ Red Flags”}];
const WS_COLORS = { Market:”#3b82f6”, Competitive:”#f59e0b”, Customer:”#22c55e”, BP:”#a855f7”, Risk:”#E0294A” };

return (
<div>
<div style={{ display:“flex”, justifyContent:“space-between”, alignItems:“center”, marginBottom:“16px” }}>
<div style={css.secLabel}>Kickoff — {p.company}</div>
<div style={{ display:“flex”, gap:“10px”, alignItems:“center” }}>
<div style={{ background:`${SIGNAL_COLORS[result.deal_signal]}15`, border:`1px solid ${SIGNAL_COLORS[result.deal_signal]}40`, color:SIGNAL_COLORS[result.deal_signal], padding:“4px 12px”, borderRadius:“2px”, fontSize:“11px”, fontFamily:“monospace” }}>{result.deal_signal}</div>
<Btn variant=“ghost” onClick={run} loading={loading} style={{padding:“6px 14px”,fontSize:“10px”}}>↺ Regenerar</Btn>
</div>
</div>
<div style={{…css.card({marginBottom:“16px”,borderLeft:“3px solid #E0294A”,borderRadius:“0 4px 4px 0”,background:”#0e0e0e”})}}>
<div style={{ fontSize:“9px”, color:”#E0294A”, fontFamily:“monospace”, letterSpacing:“3px”, marginBottom:“6px” }}>INVESTMENT THESIS</div>
<div style={{ fontSize:“13px”, color:”#c8c4be”, lineHeight:“1.7” }}>{result.investment_thesis}</div>
{result.initial_market_size && <div style={{ marginTop:“8px”, fontSize:“11px”, color:”#555”, fontFamily:“monospace” }}>Est. market size: {result.initial_market_size}</div>}
</div>
<Tabs tabs={TABS} active={tab} onChange={setTab} />
<div style={{…css.card({borderTop:“none”,borderRadius:“0 0 4px 4px”})}}>
{tab === “hypotheses” && <div style={{ display:“grid”, gap:“10px” }}>
{result.hypotheses?.map(h => <div key={h.id} style={{ background:”#0e0e0e”, border:“1px solid #1e1e1e”, borderLeft:`3px solid ${WS_COLORS[h.workstream]||"#555"}`, borderRadius:“0 3px 3px 0”, padding:“12px 16px” }}>
<div style={{ display:“flex”, gap:“8px”, marginBottom:“5px” }}>
<span style={{ fontSize:“10px”, color:WS_COLORS[h.workstream]||”#555”, fontFamily:“monospace”, background:`${WS_COLORS[h.workstream]||"#555"}15`, padding:“2px 8px”, borderRadius:“2px” }}>{h.workstream}</span>
<span style={{ fontSize:“13px”, color:”#f0ede8” }}>{h.text}</span>
</div>
<div style={{ fontSize:“11px”, color:”#666”, paddingLeft:“4px” }}>↳ {h.why_critical}</div>
</div>)}
</div>}
{tab === “experts” && <div style={{ display:“grid”, gap:“12px” }}>
{result.expert_profiles?.map((ep,i) => <div key={i} style={{ background:”#0e0e0e”, border:“1px solid #1e1e1e”, borderRadius:“3px”, padding:“14px 16px” }}>
<div style={{ fontSize:“13px”, color:”#f0ede8”, marginBottom:“4px” }}>{ep.profile}</div>
<div style={{ fontSize:“12px”, color:”#666”, marginBottom:“10px” }}>{ep.why}</div>
<div style={{ fontSize:“10px”, color:”#E0294A”, fontFamily:“monospace”, marginBottom:“6px”, letterSpacing:“2px” }}>PREGUNTAS CLAVE</div>
{ep.questions?.map((q,qi) => <div key={qi} style={{ display:“flex”, gap:“8px”, padding:“4px 0”, fontSize:“12px”, color:”#888” }}><span style={{ color:”#E0294A”, flexShrink:0 }}>?</span>{q}</div>)}
</div>)}
</div>}
{tab === “workplan” && <div style={{ display:“grid”, gap:“10px” }}>
{result.week_plan?.map((wk,i) => <div key={i} style={{ display:“flex”, gap:“16px”, padding:“12px 0”, borderBottom:“1px solid #1a1a1a” }}>
<div style={{ fontSize:“11px”, color:”#E0294A”, fontFamily:“monospace”, width:“60px”, flexShrink:0 }}>{wk.week}</div>
<div>{wk.deliverables?.map((d,di) => <div key={di} style={{ display:“flex”, gap:“8px”, fontSize:“12px”, color:”#c8c4be”, padding:“2px 0” }}><span style={{ color:”#555” }}>—</span>{d}</div>)}</div>
</div>)}
</div>}
{tab === “flags” && <div style={{ display:“grid”, gap:“8px” }}>
{result.red_flags?.map((f,i) => <div key={i} style={{ display:“flex”, gap:“12px”, padding:“10px 0”, borderBottom:“1px solid #1a1a1a” }}>
<span style={{ color:”#E0294A”, flexShrink:0, fontSize:“14px” }}>▲</span>
<span style={{ fontSize:“13px”, color:”#c8c4be” }}>{f}</span>
</div>)}
</div>}
</div>
</div>
);
};

// ─── DATA MODULE ──────────────────────────────────────────────────────────────
const DataModule = ({ project:p, onSave }) => {
const [files, setFiles] = useState(p.data?.files || []);
const [manual, setManual] = useState(p.data?.manual || {});
const fileRef = useRef();
const MANUAL_FIELDS = [
{id:“tam”,label:“TAM (€M)”},{id:“sam”,label:“SAM (€M)”},{id:“market_cagr”,label:“Market CAGR (%)”},
{id:“target_share”,label:“Market share target (%)”},{id:“avg_ticket”,label:“Ticket medio (€)”},
{id:“churn”,label:“Churn rate (%)”},{id:“nps”,label:“NPS score”},{id:“employees”,label:“Empleados”},
];

const addFile = async (f) => {
if (!f) return;
try {
const data = await parseFile(f);
const newFile = { name:f.name, rows:data.rows.length, columns:data.columns, addedAt:Date.now() };
const updated = […files, newFile];
setFiles(updated);
onSave({ files:updated, manual });
} catch { alert(“Error leyendo archivo.”); }
};

const updateManual = (id, key, val) => {
const updated = { …manual, [id]:{ …(manual[id]||{}), [key]:val } };
setManual(updated);
onSave({ files, manual:updated });
};

return (
<div>
<div style={css.secLabel}>Data — Fuentes del proyecto</div>
{/* Files */}
<div style={{…css.card({marginBottom:“20px”})}}>
<div style={{ fontSize:“10px”, letterSpacing:“2px”, color:”#E0294A”, fontFamily:“monospace”, marginBottom:“12px” }}>ARCHIVOS SUBIDOS</div>
<div onDrop={e=>{e.preventDefault();addFile(e.dataTransfer.files[0]);}} onDragOver={e=>e.preventDefault()} onClick={()=>fileRef.current?.click()}
style={{ border:“2px dashed #2a2a2a”, borderRadius:“4px”, padding:“24px”, textAlign:“center”, cursor:“pointer”, marginBottom:“12px”, background:”#111” }}>
<input ref={fileRef} type=“file” accept=”.csv,.xlsx,.xls” style={{display:“none”}} onChange={e=>addFile(e.target.files[0])} />
<div style={{ fontSize:“20px”, color:”#333”, marginBottom:“8px” }}>⬆</div>
<div style={{ fontSize:“13px”, color:”#666” }}>SABI · Orbis · Statista · Euromonitor · IBISWorld — CSV o Excel</div>
</div>
{files.length === 0 ? <div style={{ fontSize:“12px”, color:”#444”, fontFamily:“monospace” }}>No hay archivos todavía</div> :
files.map((f,i) => <div key={i} style={{ display:“flex”, gap:“12px”, alignItems:“center”, padding:“8px 0”, borderBottom:“1px solid #1a1a1a” }}>
<span style={{ color:”#22c55e”, fontSize:“12px” }}>✓</span>
<span style={{ fontSize:“13px”, color:”#f0ede8”, flex:1 }}>{f.name}</span>
<span style={{ fontSize:“11px”, color:”#555”, fontFamily:“monospace” }}>{f.rows} filas · {f.columns?.length} cols</span>
</div>)}
</div>
{/* Manual fields */}
<div style={css.card({})}>
<div style={{ fontSize:“10px”, letterSpacing:“2px”, color:”#3b82f6”, fontFamily:“monospace”, marginBottom:“14px” }}>CAMPOS MANUALES — introduce datos de cualquier fuente</div>
<div style={{ display:“grid”, gridTemplateColumns:“1fr 1fr”, gap:“10px” }}>
{MANUAL_FIELDS.map(f => {
const v = manual[f.id] || {};
return <div key={f.id} style={{ background:”#0e0e0e”, border:`1px solid ${v.value?"#3b82f630":"#1e1e1e"}`, borderRadius:“3px”, padding:“12px 14px” }}>
<label style={{ …css.lbl, marginBottom:“8px”, color:v.value?”#3b82f6”:”#555” }}>{f.label}</label>
<input value={v.value||””} onChange={e=>updateManual(f.id,“value”,e.target.value)} placeholder=“Valor” style={{…css.inp,marginBottom:“6px”}} onFocus={css.fr} onBlur={css.bg} />
<input value={v.source||””} onChange={e=>updateManual(f.id,“source”,e.target.value)} placeholder=“Fuente (Ej: Euromonitor 2024)” style={{…css.inp,fontSize:“11px”,padding:“6px 10px”}} onFocus={css.fr} onBlur={css.bg} />
</div>;
})}
</div>
</div>
</div>
);
};

// ─── MARKET MODULE ────────────────────────────────────────────────────────────
const MarketModule = ({ project:p, onSave }) => {
const [result, setResult] = useState(p.market||null);
const [ctx, setCtx] = useState(””);
const [loading, setLoading] = useState(false);
const [tab, setTab] = useState(“summary”);

const run = async () => {
if (!p.sector) return alert(“Completa el Setup primero.”);
setLoading(true);
try { const r = await claude(PROMPTS.market(p,ctx).sys, PROMPTS.market(p,ctx).usr); setResult(r); onSave(r); }
catch { alert(“Error.”); } finally { setLoading(false); }
};

const manualData = p.data?.manual || {};

if (!result) return (
<div>
<div style={css.secLabel}>Market Sizing — Top-down + Bottom-up</div>
<textarea value={ctx} onChange={e=>setCtx(e.target.value)} placeholder=“Contexto adicional: segmento foco, hipótesis del equipo, datos ya conocidos…” style={{…css.inp,minHeight:“80px”,resize:“vertical”,marginBottom:“14px”}} onFocus={css.fr} onBlur={css.bg} />
{Object.keys(manualData).length > 0 && <div style={{…css.card({marginBottom:“14px”,background:”#111”})}}>
<div style={{ fontSize:“10px”, color:”#3b82f6”, fontFamily:“monospace”, marginBottom:“8px” }}>DATOS DE DATA MODULE DISPONIBLES</div>
{Object.entries(manualData).filter(([,v])=>v.value).map(([k,v])=><div key={k} style={{ fontSize:“12px”, color:”#888”, padding:“2px 0” }}>{k}: <span style={{ color:”#f0ede8” }}>{v.value}</span>{v.source&&<span style={{ color:”#555” }}> ({v.source})</span>}</div>)}
</div>}
<Btn onClick={run} loading={loading}>◎ Calcular market size</Btn>
</div>
);

const TABS = [{id:“summary”,label:“Summary”},{id:“topdown”,label:“▲ Top-Down”},{id:“bottomup”,label:“▼ Bottom-Up”},{id:“drivers”,label:“Drivers”}];
return (
<div>
<div style={{ display:“flex”, justifyContent:“space-between”, alignItems:“center”, marginBottom:“16px” }}>
<div style={css.secLabel}>{result.market_title}</div>
<Btn variant=“ghost” onClick={run} loading={loading} style={{padding:“6px 14px”,fontSize:“10px”}}>↺ Recalcular</Btn>
</div>
<div style={{ display:“grid”, gridTemplateColumns:“1fr 1fr 1fr auto”, gap:“10px”, marginBottom:“16px” }}>
{[[“TAM”,”#E0294A”,result.tam],[“SAM”,”#f59e0b”,result.sam],[“SOM”,”#22c55e”,result.som]].map(([label,color,d])=>(
<div key={label} style={css.card({border:`1px solid ${color}20`,padding:“14px 16px”})}>
<div style={{ fontSize:“9px”, letterSpacing:“3px”, color, fontFamily:“monospace”, marginBottom:“6px” }}>{label}</div>
<div style={{ fontSize:“22px”, color:”#f0ede8”, marginBottom:“4px” }}>€{d?.value?.toLocaleString()}M</div>
<div style={{ fontSize:“11px”, color:”#666” }}>{d?.definition}</div>
</div>
))}
<div style={css.card({padding:“14px 16px”})}>
<div style={{ fontSize:“9px”, color:”#555”, fontFamily:“monospace”, marginBottom:“8px” }}>CAGR</div>
<div style={{ fontSize:“16px”, color:”#f0ede8”, marginBottom:“6px” }}>{result.cagr_historical}</div>
<div style={{ fontSize:“9px”, color:”#555”, fontFamily:“monospace”, marginBottom:“4px” }}>FORECAST</div>
<div style={{ fontSize:“16px”, color:”#E0294A” }}>{result.cagr_forward}</div>
</div>
</div>
<Tabs tabs={TABS} active={tab} onChange={setTab} />
<div style={css.card({borderTop:“none”,borderRadius:“0 0 4px 4px”})}>
{tab === “summary” && <div>
<div style={{ background:”#0e0e0e”, borderLeft:“3px solid #E0294A”, padding:“16px 18px”, borderRadius:“0 3px 3px 0”, marginBottom:“14px”, display:“flex”, gap:“24px”, alignItems:“center” }}>
<div><div style={{ fontSize:“9px”, color:”#555”, fontFamily:“monospace”, marginBottom:“3px” }}>ESTIMACIÓN RECOMENDADA</div><div style={{ fontSize:“28px”, color:”#E0294A” }}>€{result.recommended_estimate?.toLocaleString()}M</div></div>
<div style={{ width:“1px”, height:“40px”, background:”#222” }} />
{[[“Top-Down”,result.top_down?.result],[“Bottom-Up”,result.bottom_up?.result],[“Gap”,result.gap_pct]].map(([l,v])=><div key={l}><div style={{ fontSize:“9px”, color:”#555”, fontFamily:“monospace”, marginBottom:“3px” }}>{l}</div><div style={{ fontSize:“16px”, color:”#c8c4be” }}>{typeof v===“number”?`€${v?.toLocaleString()}M`:v}</div></div>)}
</div>
</div>}
{[“topdown”,“bottomup”].includes(tab) && (() => {
const d = tab===“topdown” ? result.top_down : result.bottom_up;
const color = tab===“topdown” ? “#E0294A” : “#22c55e”;
return <div>{d?.steps?.map((s,i)=><div key={i} style={{ display:“grid”, gridTemplateColumns:“28px 1fr auto”, gap:“12px”, padding:“10px 0”, borderBottom:“1px solid #1a1a1a” }}>
<div style={{ width:“26px”, height:“26px”, borderRadius:“50%”, background:”#0e0e0e”, border:`1px solid ${color}`, display:“flex”, alignItems:“center”, justifyContent:“center”, fontSize:“11px”, color, fontFamily:“monospace” }}>{s.step}</div>
<div><div style={{ fontSize:“13px”, color:”#f0ede8”, marginBottom:“2px” }}>{s.label}</div><div style={{ fontSize:“11px”, color:”#666” }}>{s.assumption}</div></div>
<div style={{ fontSize:“15px”, color:”#c8c4be”, fontFamily:“monospace”, textAlign:“right” }}>{s.value}</div>
</div>)}</div>;
})()}
{tab === “drivers” && <div style={{ display:“grid”, gridTemplateColumns:“1fr 1fr”, gap:“20px” }}>
{[[“▲ Growth Drivers”,”#22c55e”,result.growth_drivers||[]],[“▼ Risks”,”#E0294A”,result.risks||[]]].map(([label,color,items])=>(
<div key={label}>
<div style={{ fontSize:“10px”, letterSpacing:“2px”, color, fontFamily:“monospace”, marginBottom:“12px” }}>{label}</div>
{items.map((item,i)=><div key={i} style={{ background:”#0e0e0e”, border:“1px solid #1e1e1e”, borderRadius:“3px”, padding:“10px 12px”, marginBottom:“8px” }}>
<div style={{ display:“flex”, justifyContent:“space-between”, marginBottom:“2px” }}>
<span style={{ fontSize:“12px”, color:”#f0ede8” }}>{item.driver||item.risk}</span>
<span style={{ fontSize:“9px”, color, fontFamily:“monospace” }}>{item.impact}</span>
</div>
</div>)}
</div>
))}
</div>}
</div>
</div>
);
};

// ─── BENCHMARK MODULE ─────────────────────────────────────────────────────────
const BenchmarkModule = ({ project:p, onSave }) => {
const [companies, setCompanies] = useState(p.benchmark?.companies || [{ name:p.company||””, isTarget:true, revenue:””, ebitda_margin:””, growth:”” }, { name:””, isTarget:false, revenue:””, ebitda_margin:””, growth:”” }, { name:””, isTarget:false, revenue:””, ebitda_margin:””, growth:”” }]);
const [result, setResult] = useState(p.benchmark?.result || null);
const [loading, setLoading] = useState(false);
const [tab, setTab] = useState(“insights”);

const updC = (i,k,v) => setCompanies(c=>c.map((co,idx)=>idx===i?{…co,[k]:v}:co));
const validCos = companies.filter(c=>c.name.trim());

const run = async () => {
setLoading(true);
try {
const { sys, usr } = PROMPTS.benchmark(validCos, p.sector);
const r = await claude(sys, usr);
setResult(r);
onSave({ companies, result:r });
} catch { alert(“Error.”); } finally { setLoading(false); }
};

const radarData = [“scale”,“growth”,“profitability”,“market_position”,“customer_value”].map(dim=>({
dimension:dim.replace(”_”,” “),
…Object.fromEntries((result?.radar_scores||[]).map(rs=>[rs.company,rs[dim]]))
}));

return (
<div>
<div style={css.secLabel}>Competitive Benchmarking</div>
{/* Company inputs */}
<div style={{ display:“grid”, gap:“8px”, marginBottom:“16px” }}>
{companies.map((co,ci)=><div key={ci} style={css.card({padding:“12px 14px”,border:`1px solid ${co.isTarget?"#E0294A30":"#1e1e1e"}`})}>
<div style={{ display:“flex”, gap:“10px”, alignItems:“center” }}>
<input value={co.name} onChange={e=>updC(ci,“name”,e.target.value)} placeholder={ci===0?“Empresa target”:` Competidor ${ci}`} style={{…css.inp,flex:1}} onFocus={css.fr} onBlur={css.bg} />
<button onClick={()=>updC(ci,“isTarget”,!co.isTarget)} style={{ background:co.isTarget?“rgba(224,41,74,0.15)”:”#111”, border:`1px solid ${co.isTarget?"#E0294A":"#2a2a2a"}`, color:co.isTarget?”#E0294A”:”#555”, padding:“8px 10px”, fontSize:“10px”, fontFamily:“monospace”, cursor:“pointer”, borderRadius:“2px”, whiteSpace:“nowrap” }}>{co.isTarget?“★ TARGET”:“TARGET?”}</button>
{companies.length>2&&<button onClick={()=>setCompanies(c=>c.filter((_,idx)=>idx!==ci))} style={{ background:“none”, border:“1px solid #2a2a2a”, color:”#555”, padding:“8px 10px”, fontFamily:“monospace”, cursor:“pointer”, borderRadius:“2px” }}>✕</button>}
{[[“revenue”,“Revenue (€M)”],[“ebitda_margin”,“EBITDA %”],[“growth”,“Growth %”]].map(([k,ph])=>(
<input key={k} value={co[k]||””} onChange={e=>updC(ci,k,e.target.value)} placeholder={ph} style={{…css.inp,width:“110px”}} onFocus={css.fr} onBlur={css.bg} />
))}
</div>
</div>)}
<button onClick={()=>setCompanies(c=>[…c,{name:””,isTarget:false,revenue:””,ebitda_margin:””,growth:””}])} style={{ background:“none”, border:“1px dashed #2a2a2a”, color:”#555”, padding:“8px”, fontFamily:“monospace”, fontSize:“12px”, cursor:“pointer”, borderRadius:“3px” }}>+ Añadir competidor</button>
</div>
<Btn onClick={run} loading={loading} disabled={validCos.length<2}>◎ Analizar benchmark</Btn>

```
  {result && <>
    <div style={{ height:"1px", background:"#1a1a1a", margin:"24px 0" }} />
    <div style={css.card({marginBottom:"14px",display:"flex",alignItems:"center",gap:"20px"})}>
      <div style={{ textAlign:"center" }}><div style={{ fontSize:"36px", color:result.competitive_intensity>=7?"#E0294A":"#f59e0b", fontFamily:"monospace", lineHeight:1 }}>{result.competitive_intensity}</div><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace" }}>INTENSIDAD</div></div>
      <div style={{ width:"1px", height:"40px", background:"#222" }} />
      <div style={{ fontSize:"13px", color:"#888", lineHeight:"1.7", flex:1 }}>{result.landscape_narrative}</div>
    </div>
    <Tabs tabs={[{id:"insights",label:"◈ Insights"},{id:"radar",label:"◎ Radar"},{id:"table",label:"⊞ KPIs"}]} active={tab} onChange={setTab} />
    <div style={css.card({borderTop:"none",borderRadius:"0 0 4px 4px"})}>
      {tab==="insights" && <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"20px" }}>
        <div>
          {result.insights?.map((ins,i)=><div key={i} style={{ display:"flex", gap:"10px", padding:"8px 0", borderBottom:"1px solid #1a1a1a" }}><span style={{ color:"#E0294A", flexShrink:0 }}>—</span><span style={{ fontSize:"13px", color:"#c8c4be", lineHeight:"1.6" }}>{ins}</span></div>)}
        </div>
        <div>
          {result.strategic_implications?.map((si,i)=><div key={i} style={{ background:"#0e0e0e", border:"1px solid #1e1e1e", borderRadius:"3px", padding:"12px 14px", marginBottom:"8px" }}>
            <div style={{ fontSize:"13px", color:"#f0ede8", marginBottom:"4px" }}>{si.implication}</div>
            <div style={{ fontSize:"12px", color:"#666" }}>{si.detail}</div>
          </div>)}
        </div>
      </div>}
      {tab==="radar" && <ResponsiveContainer width="100%" height={300}>
        <RadarChart data={radarData} margin={{top:10,right:40,bottom:10,left:40}}>
          <PolarGrid stroke="#1e1e1e" />
          <PolarAngleAxis dataKey="dimension" tick={{fill:"#888",fontSize:11,fontFamily:"monospace"}} />
          {result.radar_scores?.map((rs,i)=><Radar key={rs.company} name={rs.company} dataKey={rs.company} stroke={COLORS[i%COLORS.length]} fill={COLORS[i%COLORS.length]} fillOpacity={0.07} strokeWidth={2} />)}
          <Tooltip />
        </RadarChart>
      </ResponsiveContainer>}
      {tab==="table" && <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontFamily:"monospace" }}>
          <thead><tr>
            {["Empresa","Revenue","EBITDA %","Growth"].map(h=><th key={h} style={{ padding:"8px 12px", textAlign:h==="Empresa"?"left":"center", fontSize:"10px", color:"#555", borderBottom:"1px solid #222", background:"#0e0e0e" }}>{h}</th>)}
          </tr></thead>
          <tbody>{validCos.map((c,i)=><tr key={i} style={{ borderBottom:"1px solid #1a1a1a" }}>
            <td style={{ padding:"9px 12px", fontSize:"12px", color:c.isTarget?"#E0294A":"#f0ede8" }}>{c.isTarget?"★ ":""}{c.name}</td>
            {["revenue","ebitda_margin","growth"].map(k=><td key={k} style={{ padding:"9px 12px", textAlign:"center", fontSize:"12px", color:"#c8c4be" }}>{c[k]||"—"}</td>)}
          </tr>)}</tbody>
        </table>
      </div>}
    </div>
  </>}
</div>
```

);
};

// ─── BP MODULE ────────────────────────────────────────────────────────────────
const BPModule = ({ project:p, onSave }) => {
const [bpText, setBpText] = useState(p.bp?.bpText || “”);
const [result, setResult] = useState(p.bp?.result || null);
const [loading, setLoading] = useState(false);
const [tab, setTab] = useState(“overview”);
const [filter, setFilter] = useState(“All”);
const [copied, setCopied] = useState(false);

const VCFG = { Aggressive:{color:”#E0294A”,bg:“rgba(224,41,74,0.12)”,icon:“▲”}, Reasonable:{color:”#22c55e”,bg:“rgba(34,197,94,0.1)”,icon:“◉”}, Conservative:{color:”#3b82f6”,bg:“rgba(59,130,246,0.1)”,icon:“▼”}, Unsubstantiated:{color:”#f59e0b”,bg:“rgba(245,158,11,0.1)”,icon:”?”} };
const QC = s => s>=8?”#22c55e”:s>=6?”#f59e0b”:”#E0294A”;

const run = async () => {
if (!bpText.trim()) return;
setLoading(true);
try {
const { sys, usr } = PROMPTS.bp(bpText, p.company, p.sector);
const r = await claude(sys, usr);
setResult(r); onSave({ bpText, result:r });
} catch { alert(“Error.”); } finally { setLoading(false); }
};

const filtered = result?.assumptions?.filter(a=>filter===“All”||a.verdict===filter) || [];

return (
<div>
<div style={css.secLabel}>Business Plan Stress-Tester</div>
<textarea value={bpText} onChange={e=>setBpText(e.target.value)} placeholder={“Pega el business plan / P&L del management:\n\nYear 1: Revenue €45M (+22% YoY), EBITDA €8.5M…\nYear 2: Revenue €58M (+29% YoY)…\nManagement expects to expand from 45 to 80 clinics…”} style={{…css.inp,minHeight:“160px”,resize:“vertical”,marginBottom:“14px”}} onFocus={css.fr} onBlur={css.bg} />
<Btn onClick={run} loading={loading} disabled={bpText.trim().length<50}>▲ Stress-test</Btn>

```
  {result && <>
    <div style={{ height:"1px", background:"#1a1a1a", margin:"20px 0" }} />
    <div style={css.card({marginBottom:"14px",display:"grid",gridTemplateColumns:"auto 1px 1fr 1px auto",gap:"20px",alignItems:"center"})}>
      <div style={{ textAlign:"center" }}>
        <div style={{ fontSize:"40px", color:QC(result.credibility_score), fontFamily:"monospace", lineHeight:1 }}>{result.credibility_score}</div>
        <div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace" }}>/10</div>
        <div style={{ fontSize:"11px", color:QC(result.credibility_score), fontFamily:"monospace", marginTop:"3px" }}>{result.credibility_label}</div>
      </div>
      <div style={{ background:"#222", height:"50px" }} />
      <div style={{ fontSize:"13px", color:"#888", lineHeight:"1.7" }}>{result.credibility_summary}</div>
      <div style={{ background:"#222", height:"50px" }} />
      <div>{["Aggressive","Reasonable","Unsubstantiated"].map(v=>{
        const c=VCFG[v]; const cnt=result.assumptions?.filter(a=>a.verdict===v).length||0;
        return <div key={v} style={{ display:"flex", gap:"8px", alignItems:"center", marginBottom:"5px" }}><span style={{ color:c.color, fontSize:"10px" }}>{c.icon}</span><span style={{ fontSize:"11px", color:"#666", fontFamily:"monospace", width:"100px" }}>{v}</span><span style={{ fontSize:"15px", color:c.color, fontFamily:"monospace" }}>{cnt}</span></div>;
      })}</div>
    </div>
    <Tabs tabs={[{id:"overview",label:"Overview"},{id:"assumptions",label:"Assumptions"},{id:"scenarios",label:"Scenarios"},{id:"memo",label:"Memo"}]} active={tab} onChange={setTab} />
    <div style={css.card({borderTop:"none",borderRadius:"0 0 4px 4px"})}>
      {tab==="overview" && <div>
        {result.make_or_break?.map((m,i)=><div key={i} style={{ background:"#0e0e0e", borderLeft:"3px solid #E0294A", padding:"12px 16px", marginBottom:"8px", borderRadius:"0 3px 3px 0" }}>
          <div style={{ display:"flex", gap:"8px", marginBottom:"4px" }}><span style={{ color:"#E0294A", fontFamily:"monospace" }}>#{i+1}</span><span style={{ fontSize:"13px", color:"#f0ede8" }}>{m.assumption}</span></div>
          <div style={{ fontSize:"12px", color:"#777", paddingLeft:"20px" }}>{m.why}</div>
          <div style={{ fontSize:"11px", color:"#E0294A", fontFamily:"monospace", paddingLeft:"20px", marginTop:"3px" }}>EBITDA: {m.ebitda_impact}</div>
        </div>)}
      </div>}
      {tab==="assumptions" && <div>
        <div style={{ display:"flex", gap:"6px", marginBottom:"14px", flexWrap:"wrap" }}>
          {["All","Aggressive","Unsubstantiated","Reasonable","Conservative"].map(v=>{
            const cfg=VCFG[v];
            return <button key={v} onClick={()=>setFilter(v)} style={{ background:filter===v?(cfg?.bg||"rgba(224,41,74,0.15)"):"#111", border:`1px solid ${filter===v?(cfg?.color||"#E0294A"):"#2a2a2a"}`, color:filter===v?(cfg?.color||"#E0294A"):"#666", padding:"4px 12px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{v}</button>;
          })}
        </div>
        {filtered.map(a=>{
          const cfg=VCFG[a.verdict];
          return <div key={a.id} style={{ background:"#0e0e0e", border:"1px solid #1e1e1e", borderRadius:"3px", padding:"14px 16px", marginBottom:"8px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"8px" }}>
              <span style={{ fontSize:"13px", color:"#f0ede8" }}>{a.name}</span>
              <span style={{ background:cfg.bg, color:cfg.color, border:`1px solid ${cfg.color}40`, padding:"2px 8px", borderRadius:"2px", fontSize:"10px", fontFamily:"monospace" }}>{cfg.icon} {a.verdict}</span>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"8px", marginBottom:"8px" }}>
              {[["MGMT CLAIM",a.management_claim],["BENCHMARK",a.benchmark]].map(([l,v])=><div key={l} style={{ background:"#111", padding:"8px 10px", borderRadius:"2px" }}><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"3px" }}>{l}</div><div style={{ fontSize:"12px", color:"#c8c4be" }}>{v}</div></div>)}
            </div>
            <div style={{ display:"flex", gap:"8px", borderTop:"1px solid #1a1a1a", paddingTop:"8px" }}><span style={{ color:"#E0294A", flexShrink:0 }}>?</span><span style={{ fontSize:"12px", color:"#888", fontStyle:"italic" }}>{a.challenge_question}</span></div>
          </div>;
        })}
      </div>}
      {tab==="scenarios" && <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"14px" }}>
        {[["bear","Bear","#E0294A","▼"],["base","Base","#f59e0b","◉"],["bull","Bull","#22c55e","▲"]].map(([key,label,color,icon])=>{
          const sc=result.scenarios?.[key];
          return <div key={key} style={{ background:"#0e0e0e", border:`1px solid ${color}25`, borderRadius:"3px", padding:"16px" }}>
            <div style={{ display:"flex", gap:"6px", alignItems:"center", marginBottom:"12px" }}><span style={{ color, fontSize:"14px" }}>{icon}</span><span style={{ color, fontFamily:"monospace", fontSize:"11px" }}>{label.toUpperCase()}</span></div>
            <div style={{ marginBottom:"6px" }}><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"2px" }}>REVENUE</div><div style={{ fontSize:"18px", color }}>{sc?.revenue}</div></div>
            <div style={{ marginBottom:"10px" }}><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"2px" }}>EBITDA</div><div style={{ fontSize:"18px", color }}>{sc?.ebitda}</div></div>
            <div style={{ fontSize:"12px", color:"#777", borderTop:"1px solid #1a1a1a", paddingTop:"8px" }}>{sc?.rationale}</div>
          </div>;
        })}
      </div>}
      {tab==="memo" && <div>
        <div style={{ display:"flex", justifyContent:"flex-end", marginBottom:"12px" }}>
          <button onClick={()=>{navigator.clipboard.writeText(result.challenge_memo);setCopied(true);setTimeout(()=>setCopied(false),2000);}} style={{ background:copied?"#22c55e":"#1a1a1a", border:"1px solid #333", color:copied?"#fff":"#aaa", padding:"6px 14px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{copied?"✓ Copiado":"⎘ Copiar"}</button>
        </div>
        <div style={{ background:"#0e0e0e", border:"1px solid #1e1e1e", borderRadius:"3px", padding:"22px 26px", fontSize:"13px", color:"#c8c4be", lineHeight:"2", whiteSpace:"pre-line" }}>{result.challenge_memo}</div>
      </div>}
    </div>
  </>}
</div>
```

);
};

// ─── EXPERTS MODULE ───────────────────────────────────────────────────────────
const ExpertsModule = ({ project:p, onSave }) => {
const [calls, setCalls] = useState(p.experts || []);
const [transcript, setTranscript] = useState(””);
const [loading, setLoading] = useState(false);
const [activeCall, setActiveCall] = useState(null);
const [activeWs, setActiveWs] = useState(“market”);
const WS = [“market”,“competitive”,“customer”,“bp”,“risk”,“quotes”];

const addCall = async () => {
if (!transcript.trim()) return;
setLoading(true);
try {
const { sys, usr } = PROMPTS.expert(transcript);
const r = await claude(sys, usr);
const updated = […calls, { …r, id:Date.now(), transcript }];
setCalls(updated); onSave(updated);
setTranscript(””); setActiveCall(updated.length-1); setActiveWs(“market”);
} catch { alert(“Error.”); } finally { setLoading(false); }
};

return (
<div>
<div style={css.secLabel}>Expert Calls — {calls.length} llamada{calls.length!==1?“s”:””} sintetizada{calls.length!==1?“s”:””}</div>
{/* Add new call */}
<div style={css.card({marginBottom:“18px”})}>
<div style={{ fontSize:“10px”, letterSpacing:“2px”, color:”#E0294A”, fontFamily:“monospace”, marginBottom:“10px” }}>NUEVA LLAMADA</div>
<textarea value={transcript} onChange={e=>setTranscript(e.target.value)} placeholder={“Pega la transcripción de la llamada con el experto…\n\nEj: Interviewer: How do you see the competitive dynamics?\nExpert: The top 3 players now hold around 65% of market share…”} style={{…css.inp,minHeight:“140px”,resize:“vertical”,marginBottom:“12px”}} onFocus={css.fr} onBlur={css.bg} />
<Btn onClick={addCall} loading={loading} disabled={transcript.trim().length<50}>◉ Sintetizar llamada</Btn>
</div>

```
  {calls.length > 0 && <>
    {/* Call selector */}
    <div style={{ display:"flex", gap:"8px", flexWrap:"wrap", marginBottom:"16px" }}>
      {calls.map((c,i)=><button key={c.id} onClick={()=>{setActiveCall(i);setActiveWs("market");}} style={{ background:activeCall===i?"rgba(224,41,74,0.15)":"#111", border:`1px solid ${activeCall===i?"#E0294A":"#2a2a2a"}`, color:activeCall===i?"#E0294A":"#666", padding:"6px 14px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{c.expert_name||`Llamada ${i+1}`}</button>)}
    </div>

    {activeCall !== null && calls[activeCall] && (() => {
      const call = calls[activeCall];
      const sigColor = SIGNAL_COLORS[call.deal_signal]||"#888";
      return (
        <div>
          <div style={css.card({marginBottom:"14px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"18px"})}>
            <div><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"5px", letterSpacing:"2px" }}>EXPERTO</div><div style={{ fontSize:"14px", color:"#f0ede8", marginBottom:"3px" }}>{call.expert_name}</div><div style={{ fontSize:"12px", color:"#777" }}>{call.expert_profile}</div></div>
            <div><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"5px", letterSpacing:"2px" }}>CREDIBILIDAD</div><div style={{ display:"flex", gap:"3px", marginBottom:"4px" }}>{[1,2,3,4,5].map(n=><div key={n} style={{ width:"14px", height:"14px", borderRadius:"2px", background:n<=call.credibility_score?"#E0294A":"#2a2a2a" }} />)}</div><div style={{ fontSize:"11px", color:"#777" }}>{call.credibility_rationale}</div></div>
            <div><div style={{ fontSize:"9px", color:"#555", fontFamily:"monospace", marginBottom:"5px", letterSpacing:"2px" }}>DEAL SIGNAL</div><div style={{ display:"inline-flex", alignItems:"center", gap:"6px", background:`${sigColor}15`, border:`1px solid ${sigColor}40`, padding:"4px 10px", borderRadius:"2px", marginBottom:"6px" }}><div style={{ width:"6px", height:"6px", borderRadius:"50%", background:sigColor }} /><span style={{ color:sigColor, fontSize:"11px", fontFamily:"monospace" }}>{call.deal_signal}</span></div><div style={{ fontSize:"11px", color:"#777" }}>{call.deal_signal_rationale}</div></div>
          </div>
          <div style={{ display:"flex", gap:"4px", borderBottom:"1px solid #1e1e1e", overflowX:"auto" }}>
            {WS.map(ws=><button key={ws} onClick={()=>setActiveWs(ws)} style={{ background:activeWs===ws?"#141414":"none", border:"none", borderBottom:activeWs===ws?"2px solid #E0294A":"2px solid transparent", color:activeWs===ws?"#f0ede8":"#555", padding:"9px 14px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", whiteSpace:"nowrap" }}>{ws}</button>)}
          </div>
          <div style={css.card({borderTop:"none",borderRadius:"0 0 4px 4px"})}>
            {call.workstreams?.[activeWs] && (() => {
              const ws = call.workstreams[activeWs];
              return <div>
                {ws.insights?.map((ins,i)=><div key={i} style={{ display:"flex", gap:"10px", padding:"9px 0", borderBottom:"1px solid #1a1a1a" }}><span style={{ color:"#E0294A", flexShrink:0 }}>—</span><span style={{ fontSize:"13px", color:"#c8c4be", lineHeight:"1.6" }}>{ins}</span></div>)}
                <div style={{ background:"#0e0e0e", borderLeft:"3px solid #E0294A", padding:"10px 14px", marginTop:"14px", borderRadius:"0 2px 2px 0" }}>
                  <span style={{ fontSize:"9px", color:"#E0294A", fontFamily:"monospace", letterSpacing:"2px" }}>SO WHAT → </span>
                  <span style={{ fontSize:"12px", color:"#f0ede8" }}>{ws.so_what}</span>
                </div>
              </div>;
            })()}
          </div>
        </div>
      );
    })()}
  </>}
</div>
```

);
};

// ─── HYPOTHESES MODULE ────────────────────────────────────────────────────────
const HypothesesModule = ({ project:p, onSave }) => {
const [hyps, setHyps] = useState(() => {
const base = p.hypotheses || [];
const fromKickoff = p.kickoff?.hypotheses || [];
if (base.length > 0) return base;
return fromKickoff.map(h => ({ …h, status:“pending”, evidence:””, updatedAt:Date.now() }));
});
const [newText, setNewText] = useState(””);
const [newWs, setNewWs] = useState(“Market”);

const WS_COLORS = { Market:”#3b82f6”, Competitive:”#f59e0b”, Customer:”#22c55e”, BP:”#a855f7”, Risk:”#E0294A” };

const add = () => {
if (!newText.trim()) return;
const updated = […hyps, { id:Date.now(), workstream:newWs, text:newText, why_critical:””, status:“pending”, evidence:””, updatedAt:Date.now() }];
setHyps(updated); onSave(updated); setNewText(””);
};

const update = (id, key, val) => {
const updated = hyps.map(h => h.id===id ? {…h,[key]:val,updatedAt:Date.now()} : h);
setHyps(updated); onSave(updated);
};

const counts = Object.fromEntries(Object.keys(HYP_STATUS).map(s=>[s, hyps.filter(h=>h.status===s).length]));

return (
<div>
<div style={{ display:“flex”, justifyContent:“space-between”, alignItems:“center”, marginBottom:“16px” }}>
<div style={css.secLabel}>Hypothesis Log — {hyps.length} hipótesis</div>
<div style={{ display:“flex”, gap:“10px” }}>
{Object.entries(HYP_STATUS).map(([s,cfg])=><div key={s} style={{ display:“flex”, gap:“5px”, alignItems:“center” }}><span style={{ color:cfg.color, fontSize:“12px” }}>{cfg.icon}</span><span style={{ fontSize:“11px”, color:cfg.color, fontFamily:“monospace” }}>{counts[s]}</span></div>)}
</div>
</div>

```
  {/* Add hypothesis */}
  <div style={css.card({marginBottom:"18px"})}>
    <div style={{ display:"flex", gap:"10px", marginBottom:"10px" }}>
      <select value={newWs} onChange={e=>setNewWs(e.target.value)} style={{...css.inp,width:"160px",cursor:"pointer"}} onFocus={css.fr} onBlur={css.bg}>
        {Object.keys(WS_COLORS).map(ws=><option key={ws} value={ws}>{ws}</option>)}
      </select>
      <input value={newText} onChange={e=>setNewText(e.target.value)} onKeyDown={e=>e.key==="Enter"&&add()} placeholder="Nueva hipótesis..." style={{...css.inp,flex:1}} onFocus={css.fr} onBlur={css.bg} />
      <Btn onClick={add} disabled={!newText.trim()} style={{padding:"9px 18px",fontSize:"10px",letterSpacing:"2px"}}>+ Añadir</Btn>
    </div>
    {p.kickoff?.hypotheses?.length>0 && hyps.length===0 && <div style={{ fontSize:"11px", color:"#555", fontFamily:"monospace" }}>Las hipótesis del Kickoff se importan automáticamente.</div>}
  </div>

  {/* Hypothesis list */}
  <div style={{ display:"grid", gap:"8px" }}>
    {hyps.map(h=>{
      const stCfg = HYP_STATUS[h.status];
      const wsCfg = WS_COLORS[h.workstream]||"#555";
      return <div key={h.id} style={{ background:"#141414", border:`1px solid ${stCfg.color}30`, borderLeft:`3px solid ${stCfg.color}`, borderRadius:"0 4px 4px 0", padding:"14px 16px" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"10px", gap:"10px" }}>
          <div style={{ flex:1 }}>
            <div style={{ display:"flex", gap:"8px", alignItems:"center", marginBottom:"4px" }}>
              <span style={{ fontSize:"10px", color:wsCfg, fontFamily:"monospace", background:`${wsCfg}15`, padding:"2px 7px", borderRadius:"2px" }}>{h.workstream}</span>
              <span style={{ fontSize:"13px", color:"#f0ede8" }}>{h.text}</span>
            </div>
          </div>
          <select value={h.status} onChange={e=>update(h.id,"status",e.target.value)} style={{ background:`${stCfg.color}15`, border:`1px solid ${stCfg.color}40`, color:stCfg.color, fontFamily:"monospace", fontSize:"10px", padding:"4px 8px", borderRadius:"2px", cursor:"pointer", flexShrink:0, outline:"none" }}>
            {Object.entries(HYP_STATUS).map(([s,cfg])=><option key={s} value={s}>{cfg.icon} {cfg.label}</option>)}
          </select>
        </div>
        <input value={h.evidence||""} onChange={e=>update(h.id,"evidence",e.target.value)} placeholder="Evidencia / notas (fuente, dato, resultado de expert call...)" style={{...css.inp,fontSize:"12px",padding:"7px 12px"}} onFocus={css.fr} onBlur={css.bg} />
      </div>;
    })}
  </div>
</div>
```

);
};

// ─── STORYBOARD MODULE ────────────────────────────────────────────────────────
const StoryboardModule = ({ project:p, onSave }) => {
const [result, setResult] = useState(p.storyboard||null);
const [loading, setLoading] = useState(false);
const [activeSection, setActiveSection] = useState(0);
const [copied, setCopied] = useState(false);

const run = async () => {
if (!p.company) return alert(“Completa Setup primero.”);
setLoading(true);
const done = [“kickoff”,“market”,“benchmark”,“bp”,“experts”].filter(m=>!!p[m]);
try {
const { sys, usr } = PROMPTS.storyboard(p, done);
const r = await claude(sys, usr);
setResult(r); onSave(r);
} catch { alert(“Error.”); } finally { setLoading(false); }
};

const exportStoryboard = () => {
if (!result) return;
let out = `CDD STORYBOARD — ${p.company}\n${"=".repeat(50)}\n\n`;
out += `DEAL VERDICT: ${result.deal_verdict}\nSIGNAL: ${result.overall_signal}\n\n`;
out += `EXECUTIVE SUMMARY\n${"─".repeat(30)}\n${result.exec_summary?.investment_thesis}\n\n`;
out += `TOP FINDINGS:\n${result.exec_summary?.top_findings?.map((f,i)=>`${i+1}. ${f}`).join("\n")}\n\n`;
out += `KEY RISKS:\n${result.exec_summary?.key_risks?.map((r,i)=>`${i+1}. ${r}`).join("\n")}\n\n`;
out += `RECOMMENDATION: ${result.exec_summary?.recommendation}\n\n`;
result.sections?.forEach(s=>{
out += `\n${s.title.toUpperCase()}\n${"─".repeat(30)}\nRole: ${s.narrative_role}\n`;
s.slides?.forEach((sl,i)=>{out += `\nSlide ${i+1}: ${sl.title}\nMessage: ${sl.message}\nChart: ${sl.chart_type}\n`;});
});
navigator.clipboard.writeText(out);
setCopied(true); setTimeout(()=>setCopied(false),2000);
};

if (!result) return (
<div>
<div style={css.secLabel}>Storyboard Builder — narrativa del deck</div>
<div style={css.card({marginBottom:“16px”,background:”#111”})}>
<div style={{ fontSize:“13px”, color:”#888”, lineHeight:“1.7”, marginBottom:“12px” }}>Genera la estructura narrativa del deck CDD basándose en todos los análisis del proyecto. Propone qué slide va en qué posición, qué mensaje lleva cada sección, y qué tipo de gráfico usar.</div>
<div style={{ fontSize:“11px”, color:”#555”, fontFamily:“monospace” }}>
Módulos completados: {[“kickoff”,“market”,“benchmark”,“bp”,“experts”].filter(m=>!!p[m]).map(m=><span key={m} style={{ color:”#22c55e”, marginRight:“8px” }}>✓ {m}</span>)}
{[“kickoff”,“market”,“benchmark”,“bp”,“experts”].filter(m=>!p[m]).map(m=><span key={m} style={{ color:”#333”, marginRight:“8px” }}>○ {m}</span>)}
</div>
</div>
<Btn onClick={run} loading={loading}>◆ Generar storyboard</Btn>
</div>
);

const sigColor = SIGNAL_COLORS[result.overall_signal]||”#888”;
return (
<div>
<div style={{ display:“flex”, justifyContent:“space-between”, alignItems:“center”, marginBottom:“16px” }}>
<div style={css.secLabel}>Storyboard — {result.slide_count} slides</div>
<div style={{ display:“flex”, gap:“10px” }}>
<div style={{ background:`${sigColor}15`, border:`1px solid ${sigColor}40`, color:sigColor, padding:“4px 12px”, borderRadius:“2px”, fontSize:“11px”, fontFamily:“monospace” }}>{result.overall_signal}</div>
<button onClick={exportStoryboard} style={{ background:copied?”#22c55e”:”#1a1a1a”, border:“1px solid #333”, color:copied?”#fff”:”#aaa”, padding:“6px 14px”, fontSize:“11px”, fontFamily:“monospace”, cursor:“pointer”, borderRadius:“2px” }}>{copied?“✓ Copiado”:“⎘ Exportar”}</button>
<Btn variant=“ghost” onClick={run} loading={loading} style={{padding:“6px 14px”,fontSize:“10px”}}>↺ Regenerar</Btn>
</div>
</div>

```
  {/* Exec summary */}
  <div style={css.card({marginBottom:"16px",borderColor:"#E0294A30"})}>
    <div style={{ fontSize:"9px", letterSpacing:"3px", color:"#E0294A", fontFamily:"monospace", marginBottom:"8px" }}>EXECUTIVE SUMMARY</div>
    <div style={{ borderLeft:"3px solid #E0294A", paddingLeft:"14px", marginBottom:"14px" }}>
      <div style={{ fontSize:"13px", color:"#c8c4be", lineHeight:"1.7", marginBottom:"10px" }}>{result.exec_summary?.investment_thesis}</div>
      <div style={{ fontSize:"12px", color:"#888", fontStyle:"italic" }}>{result.exec_summary?.recommendation}</div>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"14px" }}>
      {[["TOP FINDINGS","#22c55e",result.exec_summary?.top_findings||[]],["KEY RISKS","#E0294A",result.exec_summary?.key_risks||[]]].map(([label,color,items])=>(
        <div key={label}><div style={{ fontSize:"9px", color, fontFamily:"monospace", letterSpacing:"2px", marginBottom:"6px" }}>{label}</div>
          {items.map((item,i)=><div key={i} style={{ display:"flex", gap:"8px", padding:"4px 0", fontSize:"12px", color:"#c8c4be" }}><span style={{ color, flexShrink:0 }}>·</span>{item}</div>)}
        </div>
      ))}
    </div>
  </div>

  {/* Sections nav + slides */}
  <div style={{ display:"grid", gridTemplateColumns:"200px 1fr", gap:"12px" }}>
    <div style={css.card({padding:0,overflow:"hidden"})}>
      {result.sections?.map((s,si)=>(
        <button key={s.id} onClick={()=>setActiveSection(si)} style={{ width:"100%", background:activeSection===si?"#1e1e1e":"none", border:"none", borderLeft:activeSection===si?"3px solid #E0294A":"3px solid transparent", borderBottom:"1px solid #111", color:activeSection===si?"#f0ede8":"#555", padding:"12px 14px", cursor:"pointer", textAlign:"left", fontSize:"11px" }}>
          <div style={{ color:activeSection===si?"#E0294A":"#333", fontSize:"9px", fontFamily:"monospace", marginBottom:"2px" }}>{s.slides?.length} slides</div>
          <div>{s.title}</div>
        </button>
      ))}
    </div>
    <div style={css.card({padding:"18px 20px"})}>
      {result.sections?.[activeSection] && (() => {
        const sec = result.sections[activeSection];
        return <div>
          <div style={{ fontSize:"9px", color:"#E0294A", fontFamily:"monospace", letterSpacing:"2px", marginBottom:"4px" }}>{sec.narrative_role}</div>
          <div style={{ fontSize:"16px", color:"#f0ede8", marginBottom:"16px" }}>{sec.title}</div>
          {sec.slides?.map((sl,i)=><div key={i} style={{ background:"#0e0e0e", border:"1px solid #1e1e1e", borderRadius:"3px", padding:"14px 16px", marginBottom:"8px" }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px" }}>
              <div style={{ fontSize:"13px", color:"#f0ede8" }}>{sl.title}</div>
              <span style={{ fontSize:"10px", color:"#555", fontFamily:"monospace", background:"#111", padding:"2px 8px", borderRadius:"2px" }}>{sl.chart_type}</span>
            </div>
            <div style={{ fontSize:"12px", color:"#888", lineHeight:"1.6", borderLeft:"2px solid #E0294A", paddingLeft:"10px" }}>{sl.message}</div>
            {sl.data_hint && <div style={{ fontSize:"11px", color:"#444", fontFamily:"monospace", marginTop:"6px" }}>→ {sl.data_hint}</div>}
          </div>)}
        </div>;
      })()}
    </div>
  </div>
</div>
```

);
};

// ─── EXPORT MODULE ────────────────────────────────────────────────────────────
const ExportModule = ({ project:p }) => {
const [copied, setCopied] = useState({});

const exportExcel = () => {
const wb = XLSX.utils.book_new();
// Setup sheet
const setup = [[“PROJECT ROOM EXPORT”,””,””,””],[“Project”,p.name,””,””],[“Company”,p.company||””,””,””],[“Sector”,p.sector||””,””,””],[“Deal Type”,p.dealType||””,””,””],[“Geography”,p.geography||””,””,””],[“Revenue”,p.revenue||””,””,””],[“EBITDA”,p.ebitda||””,””,””],[””,””,””,””],[“Description”,p.description||””,””,””],[“Notes”,p.notes||””,””,””]];
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(setup), “Setup”);

```
// Hypotheses
if (p.hypotheses?.length) {
  const hData = [["Workstream","Hypothesis","Status","Evidence"],
    ...p.hypotheses.map(h=>[h.workstream,h.text,HYP_STATUS[h.status]?.label||h.status,h.evidence||""])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hData), "Hypotheses");
}
// Market sizing
if (p.market) {
  const mData = [["MARKET SIZING","",""],["TAM",p.market.tam?.value,"€M"],["SAM",p.market.sam?.value,"€M"],["SOM",p.market.som?.value,"€M"],["CAGR Historical",p.market.cagr_historical,""],["CAGR Forward",p.market.cagr_forward,""],["Recommended estimate",p.market.recommended_estimate,"€M"]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mData), "Market Sizing");
}
// Expert calls summary
if (p.experts?.length) {
  const eData = [["Expert","Profile","Credibility","Deal Signal","Market","Competitive","Customer","BP","Risk"],
    ...p.experts.map(e=>[e.expert_name,e.expert_profile,e.credibility_score,e.deal_signal,
      e.workstreams?.market?.so_what||"",e.workstreams?.competitive?.so_what||"",
      e.workstreams?.customer?.so_what||"",e.workstreams?.bp?.so_what||"",
      e.workstreams?.risk?.so_what||""])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(eData), "Expert Calls");
}
// Data fields
if (p.data?.manual) {
  const dData = [["Field","Value","Source"],
    ...Object.entries(p.data.manual).filter(([,v])=>v.value).map(([k,v])=>[k,v.value,v.source||""])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dData), "Data Fields");
}
XLSX.writeFile(wb, `${p.name}_CDD_ProjectRoom_${new Date().toISOString().slice(0,10)}.xlsx`);
```

};

const copyText = (key, text) => { navigator.clipboard.writeText(text); setCopied(c=>({…c,[key]:true})); setTimeout(()=>setCopied(c=>({…c,[key]:false})),2000); };

const completedModules = [“kickoff”,“market”,“benchmark”,“bp”,“experts”,“storyboard”].filter(m=>!!p[m]);

return (
<div>
<div style={css.secLabel}>Export — {p.name}</div>
<div style={{ display:“grid”, gridTemplateColumns:“1fr 1fr”, gap:“10px”, marginBottom:“24px” }}>
{[“kickoff”,“market”,“benchmark”,“bp”,“experts”,“storyboard”].map(m=>(
<div key={m} style={css.card({padding:“12px 16px”,border:`1px solid ${p[m]?"#22c55e30":"#1e1e1e"}`})}>
<div style={{ display:“flex”, gap:“8px”, alignItems:“center” }}>
<span style={{ color:p[m]?”#22c55e”:”#333”, fontSize:“14px” }}>{p[m]?“✓”:“○”}</span>
<span style={{ fontSize:“13px”, color:p[m]?”#f0ede8”:”#555” }}>{m.charAt(0).toUpperCase()+m.slice(1)}</span>
</div>
</div>
))}
</div>

```
  <div style={{ display:"grid", gap:"12px" }}>
    {/* Excel export */}
    <div style={css.card({})}>
      <div style={{ fontSize:"10px", letterSpacing:"2px", color:"#22c55e", fontFamily:"monospace", marginBottom:"8px" }}>EXCEL — Todos los datos del proyecto</div>
      <div style={{ fontSize:"12px", color:"#666", marginBottom:"12px" }}>Exporta hipótesis, market sizing, expert calls, data fields y setup en un único Excel multi-hoja.</div>
      <Btn onClick={exportExcel} style={{padding:"10px 22px",fontSize:"10px"}}>⬇ Descargar Excel</Btn>
    </div>

    {/* Challenge memo */}
    {p.bp?.result?.challenge_memo && <div style={css.card({})}>
      <div style={{ fontSize:"10px", letterSpacing:"2px", color:"#E0294A", fontFamily:"monospace", marginBottom:"8px" }}>CHALLENGE MEMO — Listo para enviar al management</div>
      <div style={{ background:"#0e0e0e", border:"1px solid #1e1e1e", borderRadius:"3px", padding:"16px", fontSize:"12px", color:"#888", lineHeight:"1.8", whiteSpace:"pre-line", marginBottom:"10px", maxHeight:"120px", overflow:"hidden" }}>{p.bp.result.challenge_memo.slice(0,400)}...</div>
      <button onClick={()=>copyText("memo",p.bp.result.challenge_memo)} style={{ background:copied.memo?"#22c55e":"#1a1a1a", border:"1px solid #333", color:copied.memo?"#fff":"#aaa", padding:"7px 16px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{copied.memo?"✓ Copiado":"⎘ Copiar memo"}</button>
    </div>}

    {/* Storyboard */}
    {p.storyboard && <div style={css.card({})}>
      <div style={{ fontSize:"10px", letterSpacing:"2px", color:"#f59e0b", fontFamily:"monospace", marginBottom:"8px" }}>STORYBOARD — Estructura del deck ({p.storyboard.slide_count} slides)</div>
      <div style={{ fontSize:"12px", color:"#666", marginBottom:"12px" }}>Deal verdict: <span style={{ color:SIGNAL_COLORS[p.storyboard.overall_signal]||"#888" }}>{p.storyboard.deal_verdict}</span></div>
      <button onClick={()=>{
        let out=`CDD STORYBOARD — ${p.company}\n\nVERDICT: ${p.storyboard.deal_verdict}\n\n`;
        p.storyboard.sections?.forEach(s=>{out+=`\n${s.title}\n`;s.slides?.forEach((sl,i)=>{out+=`  Slide ${i+1}: ${sl.title}\n  → ${sl.message}\n`;});});
        copyText("story",out);
      }} style={{ background:copied.story?"#22c55e":"#1a1a1a", border:"1px solid #333", color:copied.story?"#fff":"#aaa", padding:"7px 16px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{copied.story?"✓ Copiado":"⎘ Copiar storyboard"}</button>
    </div>}

    {/* Project JSON */}
    <div style={css.card({})}>
      <div style={{ fontSize:"10px", letterSpacing:"2px", color:"#555", fontFamily:"monospace", marginBottom:"8px" }}>PROJECT SNAPSHOT — JSON completo</div>
      <div style={{ fontSize:"12px", color:"#666", marginBottom:"12px" }}>Copia todo el proyecto como JSON. Útil para backup o para pasar el contexto a otra herramienta.</div>
      <button onClick={()=>copyText("json",JSON.stringify(p,null,2))} style={{ background:copied.json?"#22c55e":"#1a1a1a", border:"1px solid #333", color:copied.json?"#fff":"#aaa", padding:"7px 16px", fontSize:"11px", fontFamily:"monospace", cursor:"pointer", borderRadius:"2px" }}>{copied.json?"✓ Copiado":"⎘ Copiar JSON"}</button>
    </div>
  </div>
</div>
```

);
};

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App() {
const [projects, setProjects] = useState({});
const [activeProjectId, setActiveProjectId] = useState(null);
const [loaded, setLoaded] = useState(false);

useEffect(() => {
loadProjects().then(p => { setProjects(p); setLoaded(true); });
}, []);

const save = async (updated) => {
setProjects(updated);
await saveProjects(updated);
};

const createProject = async (name) => {
const id = `proj_${Date.now()}`;
const newP = { id, name, createdAt:Date.now(), updatedAt:Date.now() };
const updated = { …projects, [id]:newP };
await save(updated);
setActiveProjectId(id);
};

const updateProject = async (updated) => {
const all = { …projects, [updated.id]:updated };
await save(all);
};

const deleteProject = async (id) => {
if (!window.confirm(”¿Eliminar este proyecto?”)) return;
const updated = { …projects };
delete updated[id];
await save(updated);
};

if (!loaded) return (
<div style={{ minHeight:“100vh”, background:”#0a0a0a”, display:“flex”, alignItems:“center”, justifyContent:“center”, fontFamily:“monospace”, color:”#555”, flexDirection:“column”, gap:“12px” }}>
<span style={{ animation:“spin 1s linear infinite”, display:“inline-block”, fontSize:“20px” }}>◌</span>
<span style={{ letterSpacing:“3px”, fontSize:“11px” }}>CARGANDO PROYECTOS…</span>
</div>
);

if (activeProjectId && projects[activeProjectId]) {
return <ProjectRoom project={projects[activeProjectId]} onUpdate={updateProject} onBack={()=>setActiveProjectId(null)} />;
}

return <HomeScreen projects={projects} onCreate={createProject} onOpen={setActiveProjectId} onDelete={deleteProject} />;
}
