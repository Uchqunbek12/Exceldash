import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, BarChart, Bar, Cell, Legend, CartesianGrid, PieChart, Pie } from "recharts";

/* ═══════════ SUPABASE SOZLAMALARI ═══════════ */
const SUPABASE_URL = "https://fqgkujhvgvorrdlcjbnf.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZxZ2t1amh2Z3ZvcnJkbGNqYm5mIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MzMwOTgsImV4cCI6MjA5MTIwOTA5OH0.Xa9B3zt6oro-P7ygeG45sQJsK8K5ezX0T1feZ3np4GA";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

/* ═══════════ ISH KUNLARI HISOBLAGICH ═══════════ */
function getWorkDayInfo() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 0-based
  const today = now.getDate();

  let totalWorkDays = 0;
  let elapsedWorkDays = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  for (let d = 1; d <= daysInMonth; d++) {
    const day = new Date(year, month, d).getDay(); // 0=Sun, 6=Sat
    if (day !== 0 && day !== 6) {
      totalWorkDays++;
      if (d <= today) elapsedWorkDays++;
    }
  }

  return { totalWorkDays, elapsedWorkDays };
}

/* ═══════════ PLAN (static, monthly) ═══════════ */
const PLAN_RAW = [
  // Солнечный
  ["Солнечный","Анвар (агент)",108256907,79985396,158008420,6,"Солнечный"],
  ["Солнечный","Жонибек (агент)",157043834,50626611,78623011,6,"Солнечный"],
  ["Солнечный","Фаррух (агент)",87525144,117383175,39492167,10,"Солнечный"],

  // Tashkent
  ["Чиланзар","Рихсивой (ТП)",132314496,85864554,88734505,12,"Tashkent"],
  ["Бектемир","Рихсивой (ТП)",273178190,46743195,99587914,3,"Tashkent"],
  ["Алмазар","Абдулбосид (агент)",35681781,5847616,7928154,9,"Tashkent"],
  ["Шайхантахур","Абдулбосид (агент)",100054410,38950164,40989828,12,"Tashkent"],
  ["Яшнабад","Абдурахмон (ТП)",34900662,18457806,30743612,9,"Tashkent"],
  ["Учтепа","Кристина (агент)",32139541,10978225,15424864,7,"Tashkent"],
  ["Мирабад","Кристина (агент)",20505509,20086704,29787110,10,"Tashkent"],
  ["Сергели","Умида (агент)",35130299,9665094,19055148,6,"Tashkent"],
  ["Яккасарай","Умида (агент)",32000517,10001264,10653905,5,"Tashkent"],
  ["М.Улугбек","Хусан (агент)",79019334,69668865,66064583,12,"Tashkent"],
  ["Юнусабад","Хусан (агент)",67739917,40275389,51175276,12,"Tashkent"],

  // Oblast
  ["Андижан","Бобурбек (агент)",91752555,27429497,47302459,8,"Oblast"],
  ["Коканд","Бобурбек (агент)",50748529,9835233,14964050,5,"Oblast"],
  ["Фергана","Бобурбек (агент)",144824685,29172156,44992484,10,"Oblast"],
  ["Каракалпакская республика","Даврон (ТП)",59815807,9759606,12916875,7,"Oblast"],
  ["Наманган","Даврон (ТП)",170962171,29512533,51300774,9,"Oblast"],
  ["Джиззах","Расул (ТП)",30202722,5310426,11148976,7,"Oblast"],
  ["Кашкадарья","Расул (ТП)",107408066,14851188,24671300,9,"Oblast"],
  ["Сырдарья","Расул (ТП)",14954587,2195754,6962606,2,"Oblast"],
  ["Самарканд","Хамид (SG)",467801624,59987845,124096366,10,"Oblast"],
  ["Хорезм","Хамид (SG)",149961930,25286322,43090950,10,"Oblast"],
  ["Бухара","Шахбоз (агент)",248938785,32482520,63184224,14,"Oblast"],
  ["Навои","Шахбоз (агент)",135751612,17554276,28007467,7,"Oblast"],
  ["Сурхандарья","Шахбоз (агент)",26458831,4552984,11469916,9,"Oblast"],
  ["Ташкентская область","Жалолиддин (ТП)",44222820,16190846,23543212,12,"Oblast"],
];
const REGIONS = [...new Set(PLAN_RAW.map(p => p[0]))];
const PLAN_AGENTS = [...new Set(PLAN_RAW.map(p => p[1]))];
const CATS = ["Все", "Tashkent", "Солнечный", "Oblast"];

/* ═══════════ INITIAL FACT DATA ═══════════ */
const INITIAL_DATA = {
  totalPlan: 0, totalFact: 0, totalPct: 0, uploadDate: "",
  orgSummary: [], regionSummary: [], agentSummary: [], daily: []
};

/* ═══════════ AGENT NORMALIZER ═══════════ */
const AGENT_MAP = {
  "Абдулбосид": "Абдулбосид (агент)", "Бобурбек": "Бобурбек (агент)",
  "Рихсивой": "Рихсивой (ТП)", "Хусан": "Хусан (агент)",
  "Шахбоз": "Шахбоз (агент)", "Умида": "Умида (агент)",
  "Анвар": "Анвар (агент)", "Фаррух": "Фаррух (агент)",
  "Жалолиддин": "Жалолиддин (ТП)", "Даврон": "Даврон (ТП)",
  "Хамид": "Хамид (SG)", "Жонибек": "Жонибек (агент)",
  "Ахмаджон": "Ахмаджон (менеджер)", "Расул": "Расул (ТП)",
  "Сатторов": "Абдурахмон (ТП)", "Абдурахмон": "Абдурахмон (ТП)",
  "Кристина":"Кристина (агент)",
};
function normAgent(raw) {
  const c = raw.replace(/^(DT|GT|SG)\s+/i, "").trim().split(/[\s(]/)[0];
  return AGENT_MAP[c] || null;
}

/* ═══════════ EXCEL PARSER ═══════════ */
function parseExcel(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
  let hi = 0;
  for (let i = 0; i < Math.min(rows.length, 20); i++) { if (rows[i]?.some(c => c && String(c).includes("Торговая точка"))) { hi = i; break; } }
  const dateRx = /^\d{2}\.\d{2}\.\d{4}$/, agentRx = /^(DT|GT|SG)\s+(.+)$/;
  const rF = {}, aF = {}, rT = {}, aT = {}, daily = [];
  let cR = null, cA = null;
  for (let i = hi + 1; i < rows.length; i++) {
    const row = rows[i], val = row?.[1]; if (val == null) continue;
    const s = String(val).trim(); if (!s || s === "ИТОГО:") continue;
    const dt = Number(row[2]) || 0, gt = Number(row[3]) || 0, sg = Number(row[4]) || 0, tot = Number(row[5]) || 0;
    if (dateRx.test(s)) { daily.push({ date: s, DT: dt, GT: gt, SG: sg, total: tot }); cR = null; cA = null; continue; }
    if (REGIONS.includes(s)) {
      cR = s; cA = null;
      if (!rF[s]) rF[s] = { DT: 0, GT: 0, SG: 0 };
      rF[s].DT += dt; rF[s].GT += gt; rF[s].SG += sg;
      if (!rT[s]) rT[s] = new Set(); continue;
    }
    const am = s.match(agentRx);
    if (am) {
      const p = am[1], m = normAgent(am[2].trim());
      cA = { prefix: p, mapped: m };
      if (m) {
        if (!aF[m]) aF[m] = { DT: 0, GT: 0, SG: 0 };
        if (p === "DT") aF[m].DT += dt; else if (p === "GT") aF[m].GT += gt; else if (p === "SG") aF[m].SG += sg;
        if (!aT[m]) aT[m] = new Set();
      }
      continue;
    }
    if (cR && rT[cR]) rT[cR].add(s);
    if (cA?.mapped && aT[cA.mapped]) aT[cA.mapped].add(s);
  }
  return { 
    rF, aF, 
    rT: Object.fromEntries(Object.entries(rT).map(([k, v]) => [k, v.size])), 
    aT: Object.fromEntries(Object.entries(aT).map(([k, v]) => [k, v.size])),
    daily 
  };
}

function buildData(f) {
  const rMap = {};
  PLAN_RAW.forEach(([r, , dt, gt, sg, akb, cat]) => {
    if (!rMap[r]) rMap[r] = { pdt: 0, pgt: 0, psg: 0, akb: 0, cat };
    rMap[r].pdt += dt; rMap[r].pgt += gt; rMap[r].psg += sg; rMap[r].akb += akb;
  });
  const regionSummary = REGIONS.map(r => {
    const p = rMap[r], plan = p.pdt + p.pgt + p.psg;
    const ff = f.rF[r] || { DT: 0, GT: 0, SG: 0 }, fact = ff.DT + ff.GT + ff.SG;
    const short = r === "Каракалпакская республика" ? "Каракалпакская р." : r === "Ташкентская область" ? "Таш. область" : r;
    return { region: short, category: p.cat, plan, fact, pct: plan ? Math.round(fact / plan * 1000) / 10 : 0, plan_akb: p.akb, fact_tochka: f.rT[r] || 0 };
  });
  const aMap = {};
  PLAN_RAW.forEach(([r, ag, dt, gt, sg, akb]) => {
    if (!aMap[ag]) aMap[ag] = { pdt: 0, pgt: 0, psg: 0, akb: 0, regs: new Set() };
    aMap[ag].pdt += dt; aMap[ag].pgt += gt; aMap[ag].psg += sg; aMap[ag].akb += akb;
    aMap[ag].regs.add(r === "Каракалпакская республика" ? "Каракалпакия" : r === "Ташкентская область" ? "Таш. обл." : r);
  });
  const agentSummary = PLAN_AGENTS.map(ag => {
    const p = aMap[ag], plan = p.pdt + p.pgt + p.psg;
    const ff = f.aF[ag] || { DT: 0, GT: 0, SG: 0 }, fact = ff.DT + ff.GT + ff.SG;
    return { agent: ag.split("(")[0].trim(), regions: [...p.regs].join(", "), plan, fact, pct: plan ? Math.round(fact / plan * 1000) / 10 : 0, plan_akb: p.akb, fact_tochka: f.aT[ag] || 0 };
  });
  let pDT = 0, pGT = 0, pSG = 0; PLAN_RAW.forEach(([, , dt, gt, sg]) => { pDT += dt; pGT += gt; pSG += sg; });
  let fDT = 0, fGT = 0, fSG = 0; Object.values(f.rF).forEach(v => { fDT += v.DT; fGT += v.GT; fSG += v.SG; });
  const orgSummary = [
    { org: "DELI TORG", plan: pDT, fact: fDT, pct: pDT ? Math.round(fDT / pDT * 1000) / 10 : 0, color: "#E53935" },
    { org: "GRAND TRADING", plan: pGT, fact: fGT, pct: pGT ? Math.round(fGT / pGT * 1000) / 10 : 0, color: "#1E88E5" },
    { org: "SIGNUM", plan: pSG, fact: fSG, pct: pSG ? Math.round(fSG / pSG * 1000) / 10 : 0, color: "#F9A825" },
  ];
  const tP = pDT + pGT + pSG, tF = fDT + fGT + fSG;
  const mo = { "01": "Янв", "02": "Фев", "03": "Мар", "04": "Апр", "05": "Май", "06": "Июн", "07": "Июл", "08": "Авг", "09": "Сен", "10": "Окт", "11": "Ноя", "12": "Дек" };
  const daily = f.daily.map(d => ({ ...d, dateLabel: d.date.substring(0, 2) + " " + (mo[d.date.substring(3, 5)] || "") }));
  return { totalPlan: tP, totalFact: tF, totalPct: tP ? Math.round(tF / tP * 1000) / 10 : 0, orgSummary, regionSummary, agentSummary, daily, uploadDate: new Date().toLocaleDateString("ru-RU") };
}

/* ═══════════ HELPERS ═══════════ */
const fmt = n => { if (n >= 1e9) return (n / 1e9).toFixed(1) + " млрд"; if (n >= 1e6) return (n / 1e6).toFixed(1) + " млн"; if (n >= 1e3) return (n / 1e3).toFixed(0) + " тыс"; return String(n); };
const pc = p => p >= 80 ? "#2E7D32" : p >= 40 ? "#E65100" : p >= 15 ? "#F57F17" : "#C62828";
const pb = p => p >= 80 ? "#E8F5E9" : p >= 40 ? "#FFF3E0" : p >= 15 ? "#FFFDE7" : "#FFEBEE";

/* ═══════════ FORECAST HELPER ═══════════ */
function calcForecast(fact, plan, elapsedWorkDays, totalWorkDays) {
  if (!elapsedWorkDays || !plan) return 0;
  const projected = (fact / elapsedWorkDays) * totalWorkDays;
  return Math.round(projected / plan * 1000) / 10;
}

/* ═══════════ SUPABASE STORAGE ═══════════ */
const sto = {
  async get() {
    try {
      const { data, error } = await supabase.from("dashboard").select("content").eq("id", 1).maybeSingle();
      if (error) throw error;
      return data ? JSON.parse(data.content) : null;
    } catch (e) { console.error("DB Get Error:", e); return null; }
  },
  async set(v) {
    try {
      const { error } = await supabase.from("dashboard").upsert({ id: 1, content: JSON.stringify(v) });
      if (error) throw error;
      return true;
    } catch (e) { console.error("DB Set Error:", e); return false; }
  },
};

/* ═══════════ ADMIN MODAL ═══════════ */
function AdminModal({ onClose, onPublish, uploadDate }) {
  const [step, setStep] = useState("pass");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState("");
  const [dragOver, setDragOver] = useState(false);

  const login = () => { if (pass === "1234") { setStep("upload"); setErr(""); } else setErr("Parol noto'g'ri!"); };
  
  const handleFileChange = (selectedFile) => {
    if (!selectedFile) return;
    setFile(selectedFile);
    setErr("");
  };

  const publish = async () => {
    if (!file) return;
    setLoading(true);
    setErr("");
    try {
      const parsed = parseExcel(await file.arrayBuffer());
      const data = buildData(parsed);
      const ok = await sto.set(data);
      if (ok) {
        setSuccess(`✅ Muvaffaqiyat! Serverga saqlandi.`);
        setTimeout(() => { onPublish(data); onClose(); }, 1500);
      } else throw new Error("Serverga saqlashda xatolik");
    } catch (e) { setErr("Xatolik: " + e.message); }
    setLoading(false);
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 1000, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", backdropFilter: "blur(4px)" }}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 20, padding: 36, width: 440, maxWidth: "92vw", boxShadow: "0 24px 64px rgba(0,0,0,0.25)", position: "relative" }}>
        <button onClick={onClose} style={{ position: "absolute", top: 14, right: 18, background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#90A4AE" }}>✕</button>

        {step === "pass" && (<>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}>🔐</div>
            <h2 style={{ fontFamily: "'Playfair Display',serif", fontSize: 22, color: "#263238", margin: "0 0 6px" }}>Admin Panel</h2>
            <p style={{ color: "#90A4AE", fontSize: 13, marginBottom: 22 }}>Parolni kiriting</p>
          </div>
          <input type="password" placeholder="••••" value={pass} onChange={e => setPass(e.target.value)}
            onKeyDown={e => e.key === "Enter" && login()}
            style={{ width: "100%", padding: "14px 18px", border: "2px solid #E0E0E0", borderRadius: 12, fontSize: 22, textAlign: "center", letterSpacing: 10, outline: "none", fontFamily: "inherit" }} />
          <button onClick={login} style={{ width: "100%", marginTop: 16, padding: 14, border: "none", borderRadius: 12, background: "linear-gradient(135deg,#1565C0,#0D47A1)", color: "#fff", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Kirish</button>
          {err && <div style={{ marginTop: 12, padding: 10, background: "#FFEBEE", borderRadius: 8, color: "#C62828", fontSize: 13, textAlign: "center" }}>❌ {err}</div>}
        </>)}

        {step === "upload" && (<>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}>📤</div>
            <h2 style={{ fontFamily: "'Playfair Display',serif", fontSize: 20, color: "#263238", margin: "0 0 6px" }}>Faylni yuklash</h2>
            <p style={{ color: "#90A4AE", fontSize: 13, marginBottom: 20 }}>1C dan olingan Excel ni tanlang</p>
          </div>
          <div
            onClick={() => document.getElementById("fu").click()}
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => { e.preventDefault(); setDragOver(false); handleFileChange(e.dataTransfer.files[0]); }}
            style={{ border: `3px dashed ${dragOver ? "#1565C0" : "#BBDEFB"}`, borderRadius: 14, padding: "32px 16px", cursor: "pointer", textAlign: "center", background: dragOver ? "#E3F2FD" : "#F8FBFF", transition: "all 0.2s" }}>
            <input id="fu" type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => handleFileChange(e.target.files[0])} />
            {file ? (<>
              <div style={{ fontSize: 32 }}>📊</div>
              <div style={{ fontWeight: 600, color: "#263238", marginTop: 6 }}>{file.name}</div>
              <div style={{ fontSize: 12, color: "#90A4AE" }}>{(file.size / 1024).toFixed(0)} KB</div>
            </>) : (<>
              <div style={{ fontSize: 32 }}>📁</div>
              <div style={{ color: "#1565C0", fontWeight: 600, marginTop: 6 }}>Faylni tanlang yoki tashlang</div>
              <div style={{ fontSize: 11, color: "#90A4AE", marginTop: 3 }}>.xlsx format</div>
            </>)}
          </div>
          <button onClick={publish} disabled={!file || loading}
            style={{ width: "100%", marginTop: 16, padding: 14, border: "none", borderRadius: 12, background: !file || loading ? "#B0BEC5" : "linear-gradient(135deg,#2E7D32,#1B5E20)", color: "#fff", fontSize: 15, fontWeight: 700, cursor: file && !loading ? "pointer" : "default", fontFamily: "inherit" }}>
            {loading ? "⏳ Yuklanmoqda..." : "🚀 Publish"}
          </button>
          {success && <div style={{ marginTop: 12, padding: 12, background: "#E8F5E9", border: "2px solid #66BB6A", borderRadius: 10, color: "#2E7D32", fontSize: 13, textAlign: "center" }}>{success}</div>}
          {err && <div style={{ marginTop: 12, padding: 10, background: "#FFEBEE", borderRadius: 8, color: "#C62828", fontSize: 13, textAlign: "center" }}>❌ {err}</div>}
          {uploadDate && <p style={{ marginTop: 14, textAlign: "center", fontSize: 12, color: "#B0BEC5" }}>Oxirgi yuklash: {uploadDate}</p>}
        </>)}
      </div>
    </div>
  );
}

/* ═══════════ MAIN APP ═══════════ */
export default function App() {
  const [data, setData] = useState(INITIAL_DATA);
  const [showAdmin, setShowAdmin] = useState(false);
  const [catFilter, setCatFilter] = useState("Все");
  const [tab, setTab] = useState("regions");

  // ISH KUNLARI
  const { totalWorkDays, elapsedWorkDays } = getWorkDayInfo();

  useEffect(() => { 
    (async () => { 
      const s = await sto.get(); 
      if (s) setData(s); 
    })(); 
  }, []);

  const filtered = useMemo(() => {
    if (!data.regionSummary) return [];
    return catFilter === "Все" ? data.regionSummary : data.regionSummary.filter(r => r.category === catFilter);
  }, [catFilter, data.regionSummary]);

  const fP = useMemo(() => filtered.reduce((s, r) => s + r.plan, 0), [filtered]);
  const fF = useMemo(() => filtered.reduce((s, r) => s + r.fact, 0), [filtered]);
  const fPct = useMemo(() => fP ? (fF / fP * 100).toFixed(1) : 0, [fF, fP]);

  const top5 = useMemo(() => data.regionSummary ? [...data.regionSummary].sort((a, b) => b.pct - a.pct).slice(0, 5) : [], [data.regionSummary]);
  const bot5 = useMemo(() => data.regionSummary ? [...data.regionSummary].sort((a, b) => a.pct - b.pct).slice(0, 5) : [], [data.regionSummary]);

  // UMUMIY PROGNOZ
  const totalForecastPct = useMemo(() => 
    calcForecast(data.totalFact || 0, data.totalPlan || 0, elapsedWorkDays, totalWorkDays),
    [data.totalFact, data.totalPlan, elapsedWorkDays, totalWorkDays]
  );

  return (
    <div style={{ fontFamily: "Arial, Calibri, sans-serif", background: "#FAFBFE", color: "#1a1a2e", minHeight: "100vh" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700;800;900&family=Merriweather+Sans:wght@300;400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0}
        .hd{background:linear-gradient(135deg,#1565C0,#0D47A1 40%,#01579B);padding:22px 28px;color:#fff;position:relative;overflow:hidden}
        .hd::before{content:'';position:absolute;top:-50%;right:-10%;width:400px;height:400px;background:radial-gradient(circle,rgba(255,255,255,.08),transparent 70%);border-radius:50%}
        .cd{background:#fff;border:1px solid #E8EAF0;border-radius:12px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,.04);transition:box-shadow .2s}
        .cd:hover{box-shadow:0 4px 20px rgba(0,0,0,.08)}
        .kp{background:#fff;border-radius:14px;padding:20px;border:1px solid #E8EAF0;box-shadow:0 2px 8px rgba(0,0,0,.03);position:relative;overflow:hidden}
        .st{width:5px;border-radius:3px;position:absolute;left:0;top:12px;bottom:12px}
        .bt{padding:9px 20px;border-radius:24px;border:2px solid #E0E0E0;background:#fff;color:#546E7A;cursor:pointer;font-family:Arial,Calibri,sans-serif;font-size:13px;font-weight:600;transition:all .25s}
        .bt.a{background:#1565C0;color:#fff;border-color:#1565C0;box-shadow:0 3px 12px rgba(21,101,192,.3)}
        .bt:hover:not(.a){border-color:#1565C0;color:#1565C0}
        table{width:100%;border-collapse:separate;border-spacing:0}
        thead{background:#F5F7FA}
        th{text-align:left;padding:12px;font-size:10.5px;text-transform:uppercase;letter-spacing:1.2px;color:#78909C;font-weight:700;border-bottom:2px solid #E0E0E0;font-family:Arial,Calibri,sans-serif}
        td{padding:11px 12px;font-size:13px;border-bottom:1px solid #F0F0F0;font-family:Arial,Calibri,sans-serif}
        tr:hover td{background:#F8FAFF}
        .br{height:7px;border-radius:4px;background:#ECEFF1;overflow:hidden;min-width:80px}
        .fl{height:100%;border-radius:4px;transition:width .5s ease-out}
        .bg{padding:3px 8px;border-radius:4px;font-size:10px;font-weight:700;text-transform:uppercase;font-family:Arial,Calibri,sans-serif}
        .pf{font-family:'Playfair Display',serif}
        .pct-val{font-family:Arial,Calibri,sans-serif;font-weight:800}
        @keyframes fu{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
        .an{animation:fu .45s ease forwards}
        .ab{position:absolute;top:14px;right:16px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;font-family:Arial,Calibri,sans-serif;backdrop-filter:blur(4px);transition:all .2s;z-index:2}
        .ab:hover{background:rgba(255,255,255,.3)}
        .wday-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 12px;font-size:12px;font-weight:700;font-family:Arial,Calibri,sans-serif;backdrop-filter:blur(4px)}
        .forecast-row{margin-top:6px;display:flex;align-items:center;gap:6px;font-size:11.5px;color:rgba(255,255,255,.85)}
      `}</style>

      {showAdmin && <AdminModal onClose={() => setShowAdmin(false)} onPublish={d => setData(d)} uploadDate={data.uploadDate} />}

      {/* HEADER */}
      <div className="hd">
        <button className="ab" onClick={() => setShowAdmin(true)}>⚙️ Admin</button>
        <div style={{ position: "relative", zIndex: 1 }}>
          <h1 className="pf" style={{ fontSize: 26, fontWeight: 800 }}> Май 2026 — Сотув Аналитикаси</h1>
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            {/* Ish kunlari badge */}
            <span className="wday-badge">
              📅 Oy: {totalWorkDays} ish kuni
            </span>
            <span className="wday-badge">
              ⏱️ Bugun: {elapsedWorkDays}-ish kuni
            </span>
            <span style={{ opacity: .7, fontSize: 13, fontFamily: "Arial,Calibri,sans-serif" }}>
              {data.daily?.length || 0} кун · 3 ташкилот · {data.regionSummary?.length || 0} регион · {data.agentSummary?.length || 0} агент
              {data.uploadDate && <span> · Янгиланган: {data.uploadDate}</span>}
            </span>
          </div>
        </div>
      </div>

      <div style={{ padding: "22px 24px" }}>
        {/* KPIs */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 24 }} className="an">
          {[
            { l: "Умумий план", v: fmt(data.totalPlan || 0), s: "Май 2026", c: "#1565C0", i: "🎯" },
            {
              l: "Факт сотув",
              v: fmt(data.totalFact || 0),
              extra: (
                <div style={{ marginTop: 4 }}>
                  <span className="pct-val" style={{ fontSize: 13, color: pc(data.totalPct || 0), background: pb(data.totalPct || 0), padding: "2px 7px", borderRadius: 5 }}>
                    {data.totalPct || 0}% бажарилди
                  </span>
                  {totalForecastPct > 0 && (
                    <span className="pct-val" style={{ marginLeft: 6, fontSize: 12, color: "#0D47A1", background: "#E3F2FD", padding: "2px 7px", borderRadius: 5 }}>
                      📈 {totalForecastPct}% прогноз
                    </span>
                  )}
                </div>
              ),
              c: "#2E7D32", i: "💰"
            },
            { l: "Қолган сумма", v: fmt((data.totalPlan || 0) - (data.totalFact || 0)), s: (100 - (data.totalPct || 0)).toFixed(1) + "% қолди", c: "#E65100", i: "📉" },
            { 
              l: "Торг. точкалар", 
              v: data.regionSummary?.reduce((s, r) => s + r.fact_tochka, 0) || 0, 
              s: (() => {
                const fact = data.regionSummary?.reduce((s, r) => s + r.fact_tochka, 0) || 0;
                const plan = data.regionSummary?.reduce((s, r) => s + r.plan_akb, 0) || 0;
                const pct = plan ? (fact / plan * 100).toFixed(1) : 0;
                return `АКБ: ${plan} (${pct}%)`;
              })(), 
              c: "#AD1457", 
              i: "🏪" 
            },
          ].map((k, i) => (
            <div key={i} className="kp"><div className="st" style={{ background: k.c }} />
              <div style={{ paddingLeft: 12 }}>
                <div style={{ fontSize: 10.5, color: "#90A4AE", textTransform: "uppercase", letterSpacing: 1.5, fontWeight: 700, fontFamily: "Arial,Calibri,sans-serif" }}>{k.l}</div>
                <div className="pf" style={{ fontSize: 26, fontWeight: 800, color: k.c, marginTop: 5 }}>{k.v}</div>
                {k.extra ? k.extra : <div style={{ fontSize: 12, color: "#78909C", marginTop: 3, fontFamily: "Arial,Calibri,sans-serif" }}>{k.s}</div>}
              </div>
              <div style={{ position: "absolute", top: 10, right: 14, fontSize: 34, opacity: .1 }}>{k.i}</div>
            </div>
          ))}
        </div>

        {/* ORG CARDS */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16, marginBottom: 24 }} className="an">
          {data.orgSummary?.map((o, i) => {
            const r = 40, circ = 2 * Math.PI * r, off = circ - (Math.min(o.pct, 100) / 100) * circ;
            const forecastPct = calcForecast(o.fact, o.plan, elapsedWorkDays, totalWorkDays);
            const fOff = circ - (Math.min(forecastPct, 100) / 100) * circ;
            return (
              <div key={i} className="cd" style={{ display: "flex", alignItems: "flex-start", gap: 18, borderTop: `4px solid ${o.color}` }}>
                <div style={{ position: "relative", width: 90, height: 90, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <svg width={90} height={90} style={{ transform: "rotate(-90deg)" }}>
                    <circle cx={45} cy={45} r={r} fill="none" stroke="#ECEFF1" strokeWidth={7} />
                    {forecastPct > o.pct && (
                      <circle cx={45} cy={45} r={r} fill="none" stroke={o.color} strokeWidth={7}
                        strokeDasharray={circ} strokeDashoffset={fOff}
                        strokeLinecap="round" opacity={0.25} />
                    )}
                    <circle cx={45} cy={45} r={r} fill="none" stroke={o.color} strokeWidth={7} strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round" />
                  </svg>
                  <span className="pct-val" style={{ position: "absolute", fontSize: 15, color: o.color }}>{o.pct}%</span>
                </div>
                <div style={{ flex: 1, paddingTop: 4 }}>
                  <div className="pf" style={{ fontSize: 16, fontWeight: 700, color: "#263238" }}>{o.org}</div>
                  <div style={{ fontSize: 12, color: "#90A4AE", marginTop: 5, fontFamily: "Arial,Calibri,sans-serif" }}>
                    План: <b style={{ color: "#455A64" }}>{fmt(o.plan)}</b>
                  </div>
                  <div style={{ fontSize: 12, color: "#90A4AE", fontFamily: "Arial,Calibri,sans-serif" }}>
                    Факт: <b style={{ color: o.color }}>{fmt(o.fact)}</b>
                  </div>
                  {forecastPct > 0 && (
                    <div style={{ marginTop: 7, padding: "4px 8px", background: "#E3F2FD", borderRadius: 6, display: "inline-flex", alignItems: "center", gap: 5 }}>
                      <span style={{ fontSize: 11, color: "#546E7A", fontFamily: "Arial,Calibri,sans-serif" }}>📈 Прогноз:</span>
                      <span className="pct-val" style={{ fontSize: 13, color: forecastPct >= 100 ? "#2E7D32" : "#0D47A1" }}>{forecastPct}%</span>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {/* LINE CHART */}
        <div className="cd an" style={{ marginBottom: 24, padding: 22 }}>
          <h3 className="pf" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Кунлик сотув динамикаси</h3>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={data.daily || []}>
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis dataKey="dateLabel" tick={{ fill: "#78909C", fontSize: 12, fontWeight: 600, fontFamily: "Arial,Calibri,sans-serif" }} />
              <YAxis tick={{ fill: "#78909C", fontSize: 10, fontFamily: "Arial,Calibri,sans-serif" }} tickFormatter={fmt} />
              <Tooltip contentStyle={{ background: "#FFF", border: "1px solid #E0E0E0", borderRadius: 10, fontSize: 12, boxShadow: "0 4px 16px rgba(0,0,0,.06)", fontFamily: "Arial,Calibri,sans-serif" }} formatter={v => [fmt(v), ""]} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600, fontFamily: "Arial,Calibri,sans-serif" }} />
              <Line type="monotone" dataKey="total" stroke="#1565C0" strokeWidth={3} name="Жами" dot={{ r: 5, fill: "#1565C0", stroke: "#fff", strokeWidth: 2 }} />
              <Line type="monotone" dataKey="DT" stroke="#E53935" strokeWidth={2} name="DELI TORG" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="GT" stroke="#1E88E5" strokeWidth={2} name="GRAND TRADING" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="SG" stroke="#F9A825" strokeWidth={2} name="SIGNUM" dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* TOP / BOTTOM */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16, marginBottom: 24 }} className="an">
          <div className="cd" style={{ borderTop: "4px solid #2E7D32" }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#2E7D32" }}>🏆 Топ-5 регионлар</h3>
            {top5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#E8F5E9" : "#FAFAFA" }}>
                <span style={{ fontSize: 20, width: 30, textAlign: "center" }}>{["🥇","🥈","🥉","4️⃣","5️⃣"][i]}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F", fontFamily: "Arial,Calibri,sans-serif" }}>{r.region}</span>
                <span className="pct-val" style={{ color: pc(r.pct), fontSize: 15 }}>{r.pct}%</span>
              </div>
            ))}
          </div>
          <div className="cd" style={{ borderTop: "4px solid #C62828" }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#C62828" }}>⚠️ Паст-5 регионлар</h3>
            {bot5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#FFEBEE" : "#FAFAFA" }}>
                <span style={{ width: 30, textAlign: "center", fontWeight: 800, color: "#C62828", fontSize: 15, fontFamily: "Arial,Calibri,sans-serif" }}>{i + 1}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F", fontFamily: "Arial,Calibri,sans-serif" }}>{r.region}</span>
                <span className="pct-val" style={{ color: "#C62828", fontSize: 15 }}>{r.pct}%</span>
              </div>
            ))}
          </div>
        </div>

        {/* TABS */}
        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <button className={`bt ${tab === "regions" ? "a" : ""}`} onClick={() => setTab("regions")}>Регионлар</button>
          <button className={`bt ${tab === "agents" ? "a" : ""}`} onClick={() => setTab("agents")}>Агентлар</button>
        </div>

        {tab === "regions" && (
          <div style={{ display: "flex", gap: 8, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
            {CATS.map(c => (
              <button key={c} className={`bt ${catFilter === c ? "a" : ""}`} onClick={() => setCatFilter(c)} style={{ fontSize: 12, padding: "6px 14px" }}>
                {c === "Все" ? "Барчаси" : c}
              </button>
            ))}
            {catFilter !== "Все" && <span style={{ marginLeft: "auto", fontSize: 12, color: "#78909C", fontFamily: "Arial,Calibri,sans-serif" }}>
              План: <b style={{ color: "#1565C0" }}>{fmt(fP)}</b> · Факт: <b style={{ color: "#2E7D32" }}>{fmt(fF)}</b> · <b className="pct-val" style={{ color: pc(+fPct) }}>{fPct}%</b>
            </span>}
          </div>
        )}

        {/* TABLES */}
        <div className="cd an" style={{ padding: 0, overflowX: "auto" }}>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>{tab === "regions" ? "Регион" : "Агент"}</th>
                <th>{tab === "regions" ? "Категория" : "Регионлар"}</th>
                <th style={{ textAlign: "right" }}>План</th>
                <th style={{ textAlign: "right" }}>Факт</th>
                <th style={{ textAlign: "center" }}>АКБ</th>
                <th style={{ textAlign: "center" }}>Точки</th>
                <th style={{ width: 110 }}>Прогресс</th>
                <th style={{ textAlign: "right" }}>%</th>
              </tr>
            </thead>
            <tbody>
              {(tab === "regions" ? filtered : (data.agentSummary || [])).sort((a, b) => b.pct - a.pct).map((item, i) => (
                <tr key={i}>
                  <td style={{ color: "#B0BEC5", fontWeight: 700 }}>{i + 1}</td>
                  <td style={{ fontWeight: 600, color: "#263238" }}>{item.region || item.agent}</td>
                  <td>
                    {tab === "regions" ? (
                      <span className="bg" style={{ background: item.category === "Tashkent" ? "#E3F2FD" : item.category === "Солнечный" ? "#FFF8E1" : "#E8F5E9", color: item.category === "Tashkent" ? "#1565C0" : item.category === "Солнечный" ? "#F57F17" : "#2E7D32" }}>{item.category}</span>
                    ) : (
                      <span style={{ fontSize: 11, color: "#90A4AE" }}>{item.regions}</span>
                    )}
                  </td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#546E7A" }}>{fmt(item.plan)}</td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 600, color: "#2E7D32" }}>{fmt(item.fact)}</td>
                  <td style={{ textAlign: "center", color: "#78909C" }}>{item.plan_akb}</td>
                  <td style={{ textAlign: "center", fontWeight: 700, color: "#E65100" }}>{item.fact_tochka}</td>
                  <td><div className="br"><div className="fl" style={{ width: Math.min(item.pct, 100) + "%", background: pc(item.pct) }} /></div></td>
                  <td style={{ textAlign: "right" }}><span className="pct-val" style={{ fontSize: 13, color: pc(item.pct), background: pb(item.pct), padding: "3px 7px", borderRadius: 6 }}>{item.pct}%</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* BAR CHART */}
        <div className="cd an" style={{ marginTop: 24, padding: 22 }}>
          <h3 className="pf" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Регионлар бўйича сотув (Топ-10)</h3>
          <ResponsiveContainer width="100%" height={350}>
            <BarChart data={data.regionSummary ? [...data.regionSummary].sort((a, b) => b.fact - a.fact).slice(0, 10) : []} layout="vertical">
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis type="number" tick={{ fill: "#78909C", fontSize: 10, fontFamily: "Arial,Calibri,sans-serif" }} tickFormatter={fmt} />
              <YAxis type="category" dataKey="region" tick={{ fill: "#455A64", fontSize: 11, fontWeight: 600, fontFamily: "Arial,Calibri,sans-serif" }} width={100} />
              <Tooltip formatter={v => [fmt(v), ""]} contentStyle={{ fontFamily: "Arial,Calibri,sans-serif" }} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600, fontFamily: "Arial,Calibri,sans-serif" }} />
              <Bar dataKey="plan" fill="#BBDEFB" name="План" radius={[0, 4, 4, 0]} />
              <Bar dataKey="fact" fill="#1565C0" name="Факт" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* PIE + SUMMARY */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16, marginTop: 24 }}>
          <div className="cd an" style={{ padding: 22 }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Факт бўйича улуш</h3>
            <ResponsiveContainer width="100%" height={230}>
              <PieChart>
                <Pie data={data.orgSummary || []} dataKey="fact" nameKey="org" cx="50%" cy="50%" outerRadius={85} innerRadius={35}
                  label={({ org, pct }) => `${org} ${pct}%`} labelLine={{ stroke: "#B0BEC5" }} strokeWidth={2} stroke="#fff">
                  {data.orgSummary?.map((o, i) => <Cell key={i} fill={o.color} />)}
                </Pie>
                <Tooltip formatter={v => [fmt(v), ""]} contentStyle={{ fontFamily: "Arial,Calibri,sans-serif" }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="cd an" style={{ padding: 22 }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Хулоса</h3>
            <div style={{ fontSize: 13.5, lineHeight: 2.2, color: "#455A64", fontFamily: "Arial,Calibri,sans-serif" }}>
              <div>🔴 Жами: <b className="pct-val" style={{ color: pc(data.totalPct || 0), fontSize: 15 }}>{data.totalPct || 0}%</b>
                {totalForecastPct > 0 && <span style={{ marginLeft: 8, fontSize: 12, color: "#0D47A1" }}>📈 Прогноз: <b className="pct-val">{totalForecastPct}%</b></span>}
              </div>
              {(() => { const b = data.regionSummary ? [...data.regionSummary].sort((a, b) => b.pct - a.pct)[0] : null; return b ? <div>🟢 Энг яхши: <b className="pct-val" style={{ color: "#2E7D32" }}>{b.region} — {b.pct}%</b></div> : null; })()}
              {data.orgSummary?.map((o, i) => {
                const fp = calcForecast(o.fact, o.plan, elapsedWorkDays, totalWorkDays);
                return (
                  <div key={i}>{["🔴","🔵","🟡"][i]} {o.org}: <b className="pct-val" style={{ color: o.color }}>{o.pct}%</b>
                    {fp > 0 && <span style={{ marginLeft: 6, fontSize: 12, color: "#546E7A" }}>📈 <b className="pct-val">{fp}%</b></span>}
                  </div>
                );
              })}
              {(() => { const b = data.agentSummary ? [...data.agentSummary].sort((a, b) => b.pct - a.pct)[0] : null; return b ? <div>👤 Топ: <b className="pct-val" style={{ color: "#AD1457" }}>{b.agent} — {b.pct}%</b></div> : null; })()}
              {(() => { const z = data.regionSummary?.filter(r => r.pct === 0) || []; return z.length ? <div>⚠️ 0%: <b className="pct-val" style={{ color: "#C62828" }}>{z.map(r => r.region).join(", ")}</b></div> : null; })()}
            </div>
          </div>
        </div>

        <div style={{ textAlign: "center", padding: "24px 0 8px", color: "#B0BEC5", fontSize: 11, letterSpacing: 1, fontFamily: "Arial,Calibri,sans-serif" }}>
          Май 2026 · DELI TORG · GRAND TRADING · SIGNUM
        </div>
      </div>
    </div>
  );
}
