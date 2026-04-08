mport React, { useState, useCallback } from 'react';
import * as XLSX from "xlsx";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, BarChart, Bar, Cell, Legend, CartesianGrid, PieChart, Pie } from "recharts";

/* ──────────── PLAN DATA (static, changes monthly) ──────────── */
// [region, agent, DT_sum, GT_sum, SG_sum, akb, category]
const PLAN_RAW = [
  ["Солнечный","Анвар (агент)",122213328,124912698,246760772,6,"Солнечный"],
  ["Солнечный","Жонибек (агент)",245254384,79063266,122785071,6,"Солнечный"],
  ["Солнечный","Фаррух (агент)",89836662,183316579,61674672,10,"Солнечный"],
  ["Чиланзар","Ахмаджон (менеджер)",165323494,170349513,160120389,12,"Tashkent"],
  ["Учтепа","Ахмаджон (менеджер)",45168038,27892531,33398080,7,"Tashkent"],
  ["Алмазар","Абдулбосид (агент)",44583451,11601279,14306262,9,"Tashkent"],
  ["Шайхантахур","Абдулбосид (агент)",166247997,77274511,73965670,12,"Tashkent"],
  ["Яшнабад","Абдурахмон (ТП)",43607463,36619048,55476493,7,"Tashkent"],
  ["Бектемир","Рихсивой (ТП)",175148974,130735363,179705238,3,"Tashkent"],
  ["Мирабад","Рихсивой (ТП)",25621096,39850672,53750495,9,"Tashkent"],
  ["Сергели","Умида (агент)",43894388,19174897,34384795,6,"Tashkent"],
  ["Яккасарай","Умида (агент)",34973283,13729367,13660756,6,"Tashkent"],
  ["М.Улугбек","Хусан (агент)",98732587,100218352,119212776,10,"Tashkent"],
  ["Юнусабад","Хусан (агент)",84639250,79903669,92345193,12,"Tashkent"],
  ["Андижан","Бобурбек (агент)",114642412,54418282,85356741,7,"Oblast"],
  ["Коканд","Бобурбек (агент)",63408957,19512441,27002456,4,"Oblast"],
  ["Фергана","Бобурбек (агент)",180954648,57875599,81188416,10,"Oblast"],
  ["Каракалпакская республика","Даврон (ТП)",74738283,19362402,23308351,6,"Oblast"],
  ["Наманган","Даврон (ТП)",213612751,58550884,92571654,8,"Oblast"],
  ["Джиззах","Расул (ТП)",37737509,10535528,20118199,7,"Oblast"],
  ["Кашкадарья","Расул (ТП)",134203563,29463761,44519075,9,"Oblast"],
  ["Сырдарья","Расул (ТП)",18685365,4356228,12563942,2,"Oblast"],
  ["Самарканд","Хамид (SG)",584505867,119011858,223930457,10,"Oblast"],
  ["Хорезм","Хамид (SG)",187373501,50166365,77757121,10,"Oblast"],
  ["Бухара","Шахбоз (агент)",311042486,64443140,114015201,14,"Oblast"],
  ["Навои","Шахбоз (агент)",169618080,34826506,50539151,7,"Oblast"],
  ["Сурхандарья","Шахбоз (агент)",33059615,9032815,20697330,9,"Oblast"],
  ["Ташкентская область","Жалолиддин (ТП)",55255255,32121551,42483454,12,"Oblast"],
];

const REGIONS = [...new Set(PLAN_RAW.map(p => p[0]))];
const PLAN_AGENTS = [...new Set(PLAN_RAW.map(p => p[1]))];
const CATS = ["Все", "Tashkent", "Солнечный", "Oblast"];

/* ──────────── AGENT NAME NORMALIZER ──────────── */
function normalizeAgent(rawName) {
  const clean = rawName.replace(/^(DT|GT|SG)\s+/i, "").trim();
  const first = clean.split(/[\s(]/)[0];
  const MAP = {
    "Абдулбосид": "Абдулбосид (агент)", "Бобурбек": "Бобурбек (агент)",
    "Рихсивой": "Рихсивой (ТП)", "Хусан": "Хусан (агент)",
    "Шахбоз": "Шахбоз (агент)", "Умида": "Умида (агент)",
    "Анвар": "Анвар (агент)", "Фаррух": "Фаррух (агент)",
    "Жалолиддин": "Жалолиддин (ТП)", "Даврон": "Даврон (ТП)",
    "Хамид": "Хамид (SG)", "Жонибек": "Жонибек (агент)",
    "Ахмаджон": "Ахмаджон (менеджер)", "Расул": "Расул (ТП)",
    "Сатторов": "Абдурахмон (ТП)", "Абдурахмон": "Абдурахмон (ТП)",
  };
  return MAP[first] || null;
}

/* ──────────── STORAGE ABSTRACTION ──────────── */
const store = {
  async get(key) {
    try {
      if (window.storage) { const r = await window.storage.get(key); return r ? JSON.parse(r.value) : null; }
    } catch { return null; }
    return null;
  },
  async set(key, val) {
    try {
      if (window.storage) { await window.storage.set(key, JSON.stringify(val)); return true; }
    } catch { return false; }
    return false;
  },
};

/* ──────────── EXCEL PARSER ──────────── */
function parseFactExcel(arrayBuf) {
  const wb = XLSX.read(arrayBuf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let headerIdx = 0;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    if (rows[i]?.some(c => c && String(c).includes("Торговая точка"))) { headerIdx = i; break; }
  }

  const dateRx = /^\d{2}\.\d{2}\.\d{4}$/;
  const agentRx = /^(DT|GT|SG)\s+(.+)$/;

  const regionFact = {};
  const agentFact = {};
  const regionTochka = {};
  const agentTochka = {};
  const daily = [];

  let curRegion = null, curAgent = null;

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const val = row?.[1];
    if (val == null) continue;
    const s = String(val).trim();
    if (s === "ИТОГО:" || !s) continue;

    const dt = Number(row[2]) || 0;
    const gt = Number(row[3]) || 0;
    const sg = Number(row[4]) || 0;
    const tot = Number(row[5]) || 0;

    if (dateRx.test(s)) {
      daily.push({ date: s, DT: dt, GT: gt, SG: sg, total: tot });
      curRegion = null; curAgent = null;
      continue;
    }

    if (REGIONS.includes(s)) {
      curRegion = s; curAgent = null;
      if (!regionFact[s]) regionFact[s] = { DT: 0, GT: 0, SG: 0 };
      regionFact[s].DT += dt;
      regionFact[s].GT += gt;
      regionFact[s].SG += sg;
      if (!regionTochka[s]) regionTochka[s] = new Set();
      continue;
    }

    const am = s.match(agentRx);
    if (am) {
      const prefix = am[1];
      const rawName = am[2].trim();
      const mapped = normalizeAgent(rawName);
      curAgent = { prefix, mapped };
      if (mapped) {
        if (!agentFact[mapped]) agentFact[mapped] = { DT: 0, GT: 0, SG: 0 };
        if (prefix === "DT") agentFact[mapped].DT += dt;
        else if (prefix === "GT") agentFact[mapped].GT += gt;
        else if (prefix === "SG") agentFact[mapped].SG += sg;
        if (!agentTochka[mapped]) agentTochka[mapped] = new Set();
      }
      continue;
    }

    // Trading point row
    if (curRegion && regionTochka[curRegion]) regionTochka[curRegion].add(s);
    if (curAgent?.mapped && agentTochka[curAgent.mapped]) agentTochka[curAgent.mapped].add(s);
  }

  // Convert sets to counts
  const rtc = {}; for (const k in regionTochka) rtc[k] = regionTochka[k].size;
  const atc = {}; for (const k in agentTochka) atc[k] = agentTochka[k].size;

  return { regionFact, agentFact, regionTochkaCount: rtc, agentTochkaCount: atc, daily };
}

/* ──────────── BUILD DASHBOARD DATA ──────────── */
function buildData(fact) {
  const { regionFact, agentFact, regionTochkaCount, agentTochkaCount, daily } = fact;

  // Region summary
  const regionMap = {};
  PLAN_RAW.forEach(([reg, ag, dt, gt, sg, akb, cat]) => {
    if (!regionMap[reg]) regionMap[reg] = { plan_dt: 0, plan_gt: 0, plan_sg: 0, akb: 0, category: cat };
    regionMap[reg].plan_dt += dt; regionMap[reg].plan_gt += gt; regionMap[reg].plan_sg += sg;
    regionMap[reg].akb += akb;
  });

  const regionSummary = REGIONS.map(r => {
    const p = regionMap[r];
    const plan = p.plan_dt + p.plan_gt + p.plan_sg;
    const f = regionFact[r] || { DT: 0, GT: 0, SG: 0 };
    const factTotal = f.DT + f.GT + f.SG;
    return {
      region: r === "Каракалпакская республика" ? "Каракалпакская р." : r === "Ташкентская область" ? "Таш. область" : r,
      regionFull: r, category: p.category,
      plan, fact: factTotal, pct: plan ? Math.round(factTotal / plan * 1000) / 10 : 0,
      plan_akb: p.akb, fact_tochka: regionTochkaCount[r] || 0,
    };
  });

  // Agent summary
  const agentMap = {};
  PLAN_RAW.forEach(([reg, ag, dt, gt, sg, akb]) => {
    if (!agentMap[ag]) agentMap[ag] = { plan_dt: 0, plan_gt: 0, plan_sg: 0, akb: 0, regions: new Set() };
    agentMap[ag].plan_dt += dt; agentMap[ag].plan_gt += gt; agentMap[ag].plan_sg += sg;
    agentMap[ag].akb += akb;
    agentMap[ag].regions.add(reg === "Каракалпакская республика" ? "Каракалпакия" : reg === "Ташкентская область" ? "Таш. обл." : reg);
  });

  const agentSummary = PLAN_AGENTS.map(ag => {
    const p = agentMap[ag];
    const plan = p.plan_dt + p.plan_gt + p.plan_sg;
    const f = agentFact[ag] || { DT: 0, GT: 0, SG: 0 };
    const factTotal = f.DT + f.GT + f.SG;
    const shortName = ag.split("(")[0].trim();
    return {
      agent: shortName, agentFull: ag, regions: [...p.regions].join(", "),
      plan, fact: factTotal, pct: plan ? Math.round(factTotal / plan * 1000) / 10 : 0,
      plan_akb: p.akb, fact_tochka: agentTochkaCount[ag] || 0,
    };
  });

  // Org summary
  let planDT = 0, planGT = 0, planSG = 0;
  PLAN_RAW.forEach(([, , dt, gt, sg]) => { planDT += dt; planGT += gt; planSG += sg; });
  let factDT = 0, factGT = 0, factSG = 0;
  Object.values(regionFact).forEach(f => { factDT += f.DT; factGT += f.GT; factSG += f.SG; });

  const orgSummary = [
    { org: "DELI TORG", plan: planDT, fact: factDT, pct: planDT ? Math.round(factDT / planDT * 1000) / 10 : 0, color: "#E53935" },
    { org: "GRAND TRADING", plan: planGT, fact: factGT, pct: planGT ? Math.round(factGT / planGT * 1000) / 10 : 0, color: "#1E88E5" },
    { org: "SIGNUM", plan: planSG, fact: factSG, pct: planSG ? Math.round(factSG / planSG * 1000) / 10 : 0, color: "#F9A825" },
  ];

  const totalPlan = planDT + planGT + planSG;
  const totalFact = factDT + factGT + factSG;

  // Format daily dates
  const months = { "01": "Янв", "02": "Фев", "03": "Мар", "04": "Апр", "05": "Май", "06": "Июн" };
  const dailyFormatted = daily.map(d => ({
    ...d,
    dateLabel: d.date.substring(0, 2) + " " + (months[d.date.substring(3, 5)] || d.date.substring(3, 5)),
  }));

  return {
    totalPlan, totalFact, totalPct: totalPlan ? Math.round(totalFact / totalPlan * 1000) / 10 : 0,
    orgSummary, regionSummary, agentSummary, daily: dailyFormatted,
    uploadDate: new Date().toLocaleDateString("ru-RU"),
  };
}

/* ──────────── HELPERS ──────────── */
const fmt = (n) => {
  if (n >= 1e9) return (n / 1e9).toFixed(1) + " млрд";
  if (n >= 1e6) return (n / 1e6).toFixed(1) + " млн";
  if (n >= 1e3) return (n / 1e3).toFixed(0) + " тыс";
  return String(n);
};
const pctColor = (p) => p >= 80 ? "#2E7D32" : p >= 40 ? "#E65100" : p >= 15 ? "#F57F17" : "#C62828";
const pctBg = (p) => p >= 80 ? "#E8F5E9" : p >= 40 ? "#FFF3E0" : p >= 15 ? "#FFFDE7" : "#FFEBEE";

/* ──────────── STYLES ──────────── */
const STYLES = `
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700;800;900&family=Merriweather+Sans:wght@300;400;500;600;700&display=swap');
* { box-sizing: border-box; margin: 0; }
body { font-family: 'Merriweather Sans', sans-serif; }
.header { background: linear-gradient(135deg, #1565C0 0%, #0D47A1 40%, #01579B 100%); padding: 24px 28px; color: white; position: relative; overflow: hidden; }
.header::before { content: ''; position: absolute; top: -50%; right: -10%; width: 400px; height: 400px; background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%); border-radius: 50%; }
.header h1 { font-family: 'Playfair Display', serif; font-size: 26px; font-weight: 800; position: relative; z-index: 1; }
.header p { opacity: 0.75; font-size: 13px; margin-top: 6px; position: relative; z-index: 1; }
.card { background: #fff; border: 1px solid #E8EAF0; border-radius: 12px; padding: 20px; box-shadow: 0 2px 12px rgba(0,0,0,0.04); transition: box-shadow 0.2s; }
.card:hover { box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
.kpi { background: #fff; border-radius: 14px; padding: 20px; border: 1px solid #E8EAF0; box-shadow: 0 2px 8px rgba(0,0,0,0.03); position: relative; overflow: hidden; }
.stripe { width: 5px; border-radius: 3px; position: absolute; left: 0; top: 12px; bottom: 12px; }
.btn { padding: 10px 22px; border-radius: 24px; border: 2px solid #E0E0E0; background: #fff; color: #546E7A; cursor: pointer; font-family: 'Merriweather Sans'; font-size: 13px; font-weight: 600; transition: all 0.25s; }
.btn.active { background: #1565C0; color: white; border-color: #1565C0; box-shadow: 0 3px 12px rgba(21,101,192,0.3); }
.btn:hover:not(.active) { border-color: #1565C0; color: #1565C0; }
table { width: 100%; border-collapse: separate; border-spacing: 0; }
thead { background: #F5F7FA; }
th { text-align: left; padding: 12px 12px; font-size: 10.5px; text-transform: uppercase; letter-spacing: 1.2px; color: #78909C; font-weight: 700; border-bottom: 2px solid #E0E0E0; }
td { padding: 11px 12px; font-size: 13px; border-bottom: 1px solid #F0F0F0; }
tr:hover td { background: #F8FAFF; }
.bar { height: 7px; border-radius: 4px; background: #ECEFF1; overflow: hidden; min-width: 80px; }
.fill { height: 100%; border-radius: 4px; transition: width 0.8s ease; }
.badge { font-size: 10px; padding: 3px 9px; border-radius: 12px; font-weight: 700; letter-spacing: 0.4px; }
.playfair { font-family: 'Playfair Display', serif; }
@keyframes fadeUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
.anim { animation: fadeUp 0.45s ease forwards; }
.admin-bg { min-height: 100vh; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #0D47A1, #1565C0, #1976D2); }
.admin-card { background: white; border-radius: 20px; padding: 40px; width: 420px; max-width: 90vw; box-shadow: 0 20px 60px rgba(0,0,0,0.2); text-align: center; }
.input { width: 100%; padding: 14px 18px; border: 2px solid #E0E0E0; border-radius: 12px; font-size: 16px; font-family: 'Merriweather Sans'; outline: none; transition: border 0.2s; }
.input:focus { border-color: #1565C0; }
.submit-btn { width: 100%; padding: 14px; border: none; border-radius: 12px; background: linear-gradient(135deg, #1565C0, #0D47A1); color: white; font-size: 16px; font-weight: 700; cursor: pointer; font-family: 'Merriweather Sans'; transition: transform 0.1s; }
.submit-btn:hover { transform: scale(1.02); }
.submit-btn:active { transform: scale(0.98); }
.upload-zone { border: 3px dashed #BBDEFB; border-radius: 16px; padding: 40px 20px; cursor: pointer; transition: all 0.3s; background: #F8FBFF; }
.upload-zone:hover { border-color: #1565C0; background: #EFF6FF; }
.upload-zone.dragover { border-color: #1565C0; background: #E3F2FD; }
.success-box { background: #E8F5E9; border: 2px solid #66BB6A; border-radius: 12px; padding: 20px; margin-top: 20px; }
.error-box { background: #FFEBEE; border: 2px solid #EF5350; border-radius: 12px; padding: 16px; margin-top: 16px; color: #C62828; font-size: 14px; }
.nav-bar { display: flex; justify-content: space-between; align-items: center; padding: 0 28px; background: #F5F7FA; border-bottom: 1px solid #E0E0E0; height: 48px; }
`;

/* ──────────── ADMIN PANEL ──────────── */
function AdminPanel({ onPublish, existingDate }) {
  const [pass, setPass] = useState("");
  const [authed, setAuthed] = useState(false);
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState(null);
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);

  const handleLogin = () => {
    if (pass === "1234") setAuthed(true);
    else setStatus({ type: "error", msg: "Parol noto'g'ri!" });
  };

  const handleFile = (f) => {
    if (f && (f.name.endsWith(".xlsx") || f.name.endsWith(".xls"))) {
      setFile(f);
      setStatus(null);
    } else {
      setStatus({ type: "error", msg: "Faqat .xlsx yoki .xls fayl yuklang!" });
    }
  };

  const handlePublish = async () => {
    if (!file) return;
    setLoading(true);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseFactExcel(buf);
      const data = buildData(parsed);
      await store.set("dashboard_data", data);
      setStatus({ type: "success", msg: `Ma'lumot muvaffaqiyatli yuklandi! ${data.daily.length} kun, ${data.regionSummary.length} region.` });
      onPublish(data);
    } catch (err) {
      setStatus({ type: "error", msg: "Xatolik: " + err.message });
    }
    setLoading(false);
  };

  if (!authed) {
    return (
      <div className="admin-bg">
        <div className="admin-card">
          <div style={{ fontSize: 48, marginBottom: 16 }}>🔐</div>
          <h2 className="playfair" style={{ fontSize: 24, color: "#263238", marginBottom: 8 }}>Admin Panel</h2>
          <p style={{ color: "#78909C", fontSize: 14, marginBottom: 24 }}>Parolni kiriting</p>
          <input className="input" type="password" placeholder="Parol" value={pass}
            onChange={e => setPass(e.target.value)}
            onKeyDown={e => e.key === "Enter" && handleLogin()}
            style={{ marginBottom: 16, textAlign: "center", fontSize: 20, letterSpacing: 8 }} />
          <button className="submit-btn" onClick={handleLogin}>Kirish</button>
          {status?.type === "error" && <div className="error-box">{status.msg}</div>}
          {existingDate && (
            <p style={{ marginTop: 20, fontSize: 13, color: "#90A4AE" }}>
              Oxirgi yuklash: <b style={{ color: "#1565C0" }}>{existingDate}</b>
            </p>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="admin-bg">
      <div className="admin-card">
        <div style={{ fontSize: 48, marginBottom: 16 }}>📤</div>
        <h2 className="playfair" style={{ fontSize: 22, color: "#263238", marginBottom: 8 }}>Faktni yuklash</h2>
        <p style={{ color: "#78909C", fontSize: 13, marginBottom: 24 }}>1C dan olingan Excel faylni yuklang</p>

        <div
          className={`upload-zone ${dragOver ? "dragover" : ""}`}
          onClick={() => document.getElementById("fileInput").click()}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
        >
          <input id="fileInput" type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={e => handleFile(e.target.files[0])} />
          {file ? (
            <div>
              <div style={{ fontSize: 36 }}>📊</div>
              <div style={{ fontWeight: 600, color: "#263238", marginTop: 8 }}>{file.name}</div>
              <div style={{ fontSize: 12, color: "#90A4AE" }}>{(file.size / 1024).toFixed(0)} KB</div>
            </div>
          ) : (
            <div>
              <div style={{ fontSize: 36 }}>📁</div>
              <div style={{ color: "#1565C0", fontWeight: 600, marginTop: 8 }}>Faylni tanlang yoki shu yerga tashlang</div>
              <div style={{ fontSize: 12, color: "#90A4AE", marginTop: 4 }}>.xlsx yoki .xls</div>
            </div>
          )}
        </div>

        <button className="submit-btn" onClick={handlePublish}
          disabled={!file || loading}
          style={{ marginTop: 20, opacity: !file || loading ? 0.6 : 1 }}>
          {loading ? "⏳ Yuklanmoqda..." : "🚀 Publish qilish"}
        </button>

        {status?.type === "success" && <div className="success-box">✅ {status.msg}</div>}
        {status?.type === "error" && <div className="error-box">❌ {status.msg}</div>}

        {existingDate && (
          <p style={{ marginTop: 16, fontSize: 12, color: "#90A4AE" }}>
            Oxirgi yuklash: <b>{existingDate}</b>
          </p>
        )}
      </div>
    </div>
  );
}

/* ──────────── DASHBOARD ──────────── */
function DashboardView({ data, onAdmin }) {
  const [catFilter, setCatFilter] = useState("Все");
  const [tab, setTab] = useState("regions");

  const filtered = useMemo(() => {
    if (catFilter === "Все") return data.regionSummary;
    return data.regionSummary.filter(r => r.category === catFilter);
  }, [catFilter, data]);

  const fPlan = filtered.reduce((s, r) => s + r.plan, 0);
  const fFact = filtered.reduce((s, r) => s + r.fact, 0);
  const fPct = fPlan ? (fFact / fPlan * 100).toFixed(1) : 0;

  const top5 = [...data.regionSummary].sort((a, b) => b.pct - a.pct).slice(0, 5);
  const bot5 = [...data.regionSummary].sort((a, b) => a.pct - b.pct).slice(0, 5);

  return (
    <div style={{ background: "#FAFBFE", minHeight: "100vh" }}>
      {/* Nav */}
      <div className="nav-bar">
        <span style={{ fontSize: 13, color: "#90A4AE" }}>Yangilangan: <b style={{ color: "#1565C0" }}>{data.uploadDate}</b></span>
        <button className="btn" onClick={onAdmin} style={{ padding: "6px 16px", fontSize: 12 }}>⚙️ Admin</button>
      </div>

      {/* Header */}
      <div className="header">
        <h1>Апрель 2026 — Сотув Аналитикаси</h1>
        <p>3 ташкилот · {data.regionSummary.length} регион · {data.agentSummary.length} агент · {data.daily.length} кун</p>
      </div>

      <div style={{ padding: "22px 24px" }}>
        {/* KPIs */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16, marginBottom: 24 }} className="anim">
          {[
            { l: "Умумий план", v: fmt(data.totalPlan), s: "Апрель 2026", c: "#1565C0", ic: "🎯" },
            { l: "Факт сотув", v: fmt(data.totalFact), s: data.totalPct + "% бажарилган", c: "#2E7D32", ic: "💰" },
            { l: "Қолган сумма", v: fmt(data.totalPlan - data.totalFact), s: (100 - data.totalPct).toFixed(1) + "% қолди", c: "#E65100", ic: "📉" },
            { l: "Торг. точкалар", v: data.regionSummary.reduce((s, r) => s + r.fact_tochka, 0), s: "АКБ план: " + data.regionSummary.reduce((s, r) => s + r.plan_akb, 0), c: "#AD1457", ic: "🏪" },
          ].map((k, i) => (
            <div key={i} className="kpi">
              <div className="stripe" style={{ background: k.c }} />
              <div style={{ paddingLeft: 12 }}>
                <div style={{ fontSize: 10.5, color: "#90A4AE", textTransform: "uppercase", letterSpacing: 1.5, fontWeight: 700 }}>{k.l}</div>
                <div className="playfair" style={{ fontSize: 26, fontWeight: 800, color: k.c, marginTop: 5 }}>{k.v}</div>
                <div style={{ fontSize: 12, color: "#78909C", marginTop: 3 }}>{k.s}</div>
              </div>
              <div style={{ position: "absolute", top: 10, right: 14, fontSize: 34, opacity: 0.1 }}>{k.ic}</div>
            </div>
          ))}
        </div>

        {/* Org cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16, marginBottom: 24 }} className="anim">
          {data.orgSummary.map((o, i) => {
            const r = 40, circ = 2 * Math.PI * r, off = circ - (Math.min(o.pct, 100) / 100) * circ;
            return (
              <div key={i} className="card" style={{ display: "flex", alignItems: "center", gap: 18, borderTop: `4px solid ${o.color}` }}>
                <div style={{ position: "relative", width: 90, height: 90, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <svg width={90} height={90} style={{ transform: "rotate(-90deg)" }}>
                    <circle cx={45} cy={45} r={r} fill="none" stroke="#ECEFF1" strokeWidth={7} />
                    <circle cx={45} cy={45} r={r} fill="none" stroke={o.color} strokeWidth={7} strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round" />
                  </svg>
                  <span className="playfair" style={{ position: "absolute", fontSize: 16, fontWeight: 800, color: o.color }}>{o.pct}%</span>
                </div>
                <div>
                  <div className="playfair" style={{ fontSize: 16, fontWeight: 700, color: "#263238" }}>{o.org}</div>
                  <div style={{ fontSize: 12, color: "#90A4AE", marginTop: 5 }}>План: <b style={{ color: "#455A64" }}>{fmt(o.plan)}</b></div>
                  <div style={{ fontSize: 12, color: "#90A4AE" }}>Факт: <b style={{ color: o.color }}>{fmt(o.fact)}</b></div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Line Chart */}
        <div className="card anim" style={{ marginBottom: 24, padding: 22 }}>
          <h3 className="playfair" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Кунлик сотув динамикаси</h3>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={data.daily}>
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis dataKey="dateLabel" tick={{ fill: "#78909C", fontSize: 12, fontWeight: 600 }} />
              <YAxis tick={{ fill: "#78909C", fontSize: 10 }} tickFormatter={fmt} />
              <Tooltip contentStyle={{ background: "#FFF", border: "1px solid #E0E0E0", borderRadius: 10, fontSize: 12, boxShadow: "0 4px 16px rgba(0,0,0,0.06)" }} formatter={(v) => [fmt(v), ""]} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600 }} />
              <Line type="monotone" dataKey="total" stroke="#1565C0" strokeWidth={3} name="Жами" dot={{ r: 5, fill: "#1565C0", stroke: "#fff", strokeWidth: 2 }} />
              <Line type="monotone" dataKey="DT" stroke="#E53935" strokeWidth={2} name="DELI TORG" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="GT" stroke="#1E88E5" strokeWidth={2} name="GRAND TRADING" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="SG" stroke="#F9A825" strokeWidth={2} name="SIGNUM" dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Top / Bottom */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }} className="anim">
          <div className="card" style={{ borderTop: "4px solid #2E7D32" }}>
            <h3 className="playfair" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#2E7D32" }}>Топ-5 регионлар</h3>
            {top5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#E8F5E9" : "#FAFAFA" }}>
                <span style={{ fontSize: 20, width: 30, textAlign: "center" }}>{["🥇","🥈","🥉","4️⃣","5️⃣"][i]}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F" }}>{r.region}</span>
                <span className="playfair" style={{ fontWeight: 800, color: pctColor(r.pct), fontSize: 15 }}>{r.pct}%</span>
              </div>
            ))}
          </div>
          <div className="card" style={{ borderTop: "4px solid #C62828" }}>
            <h3 className="playfair" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#C62828" }}>Паст-5 регионлар</h3>
            {bot5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#FFEBEE" : "#FAFAFA" }}>
                <span style={{ width: 30, textAlign: "center", fontWeight: 800, color: "#C62828", fontSize: 15 }}>{i + 1}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F" }}>{r.region}</span>
                <span className="playfair" style={{ fontWeight: 800, color: "#C62828", fontSize: 15 }}>{r.pct}%</span>
              </div>
            ))}
          </div>
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <button className={`btn ${tab === "regions" ? "active" : ""}`} onClick={() => setTab("regions")}>Регионлар</button>
          <button className={`btn ${tab === "agents" ? "active" : ""}`} onClick={() => setTab("agents")}>Агентлар</button>
        </div>

        {/* Category filter */}
        {tab === "regions" && (
          <div style={{ display: "flex", gap: 8, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
            {CATS.map(c => (
              <button key={c} className={`btn ${catFilter === c ? "active" : ""}`}
                onClick={() => setCatFilter(c)} style={{ fontSize: 12, padding: "6px 14px" }}>
                {c === "Все" ? "Барчаси" : c}
              </button>
            ))}
            {catFilter !== "Все" && (
              <span style={{ marginLeft: "auto", fontSize: 12, color: "#78909C" }}>
                План: <b style={{ color: "#1565C0" }}>{fmt(fPlan)}</b> · Факт: <b style={{ color: "#2E7D32" }}>{fmt(fFact)}</b> · <b style={{ color: pctColor(+fPct) }}>{fPct}%</b>
              </span>
            )}
          </div>
        )}

        {/* Region Table */}
        {tab === "regions" && (
          <div className="card anim" style={{ padding: 0, overflowX: "auto" }}>
            <table>
              <thead><tr>
                <th>#</th><th>Регион</th><th>Категория</th>
                <th style={{ textAlign: "right" }}>План</th><th style={{ textAlign: "right" }}>Факт</th>
                <th style={{ textAlign: "center" }}>АКБ</th><th style={{ textAlign: "center" }}>Точки</th>
                <th style={{ width: 110 }}>Прогресс</th><th style={{ textAlign: "right" }}>%</th>
              </tr></thead>
              <tbody>
                {[...filtered].sort((a, b) => b.pct - a.pct).map((r, i) => (
                  <tr key={i}>
                    <td style={{ color: "#B0BEC5", fontWeight: 700 }}>{i + 1}</td>
                    <td style={{ fontWeight: 600, color: "#263238" }}>{r.region}</td>
                    <td><span className="badge" style={{
                      background: r.category === "Tashkent" ? "#E3F2FD" : r.category === "Солнечный" ? "#FFF8E1" : "#E8F5E9",
                      color: r.category === "Tashkent" ? "#1565C0" : r.category === "Солнечный" ? "#F57F17" : "#2E7D32",
                    }}>{r.category}</span></td>
                    <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#546E7A" }}>{fmt(r.plan)}</td>
                    <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 600, color: "#2E7D32" }}>{fmt(r.fact)}</td>
                    <td style={{ textAlign: "center", color: "#78909C" }}>{r.plan_akb}</td>
                    <td style={{ textAlign: "center", fontWeight: 700, color: "#E65100" }}>{r.fact_tochka}</td>
                    <td><div className="bar"><div className="fill" style={{ width: Math.min(r.pct, 100) + "%", background: pctColor(r.pct) }} /></div></td>
                    <td style={{ textAlign: "right" }}><span className="playfair" style={{ fontWeight: 800, fontSize: 13, color: pctColor(r.pct), background: pctBg(r.pct), padding: "3px 7px", borderRadius: 6 }}>{r.pct}%</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Agent Table */}
        {tab === "agents" && (
          <div className="card anim" style={{ padding: 0, overflowX: "auto" }}>
            <table>
              <thead><tr>
                <th>#</th><th>Агент</th><th>Регионлар</th>
                <th style={{ textAlign: "right" }}>План</th><th style={{ textAlign: "right" }}>Факт</th>
                <th style={{ textAlign: "center" }}>АКБ</th><th style={{ textAlign: "center" }}>Точки</th>
                <th style={{ width: 110 }}>Прогресс</th><th style={{ textAlign: "right" }}>%</th>
              </tr></thead>
              <tbody>
                {[...data.agentSummary].sort((a, b) => b.pct - a.pct).map((a, i) => (
                  <tr key={i}>
                    <td style={{ color: "#B0BEC5", fontWeight: 700 }}>{i + 1}</td>
                    <td style={{ fontWeight: 700, color: "#263238" }}>{a.agent}</td>
                    <td style={{ fontSize: 11, color: "#90A4AE" }}>{a.regions}</td>
                    <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#546E7A" }}>{fmt(a.plan)}</td>
                    <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 600, color: "#2E7D32" }}>{fmt(a.fact)}</td>
                    <td style={{ textAlign: "center", color: "#78909C" }}>{a.plan_akb}</td>
                    <td style={{ textAlign: "center", fontWeight: 700, color: "#E65100" }}>{a.fact_tochka}</td>
                    <td><div className="bar"><div className="fill" style={{ width: Math.min(a.pct, 100) + "%", background: pctColor(a.pct) }} /></div></td>
                    <td style={{ textAlign: "right" }}><span className="playfair" style={{ fontWeight: 800, fontSize: 13, color: pctColor(a.pct), background: pctBg(a.pct), padding: "3px 7px", borderRadius: 6 }}>{a.pct}%</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Bar Chart */}
        <div className="card anim" style={{ marginTop: 24, padding: 22 }}>
          <h3 className="playfair" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Регионлар — План ва Факт</h3>
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={[...data.regionSummary].sort((a, b) => b.fact - a.fact).slice(0, 10)} layout="vertical">
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis type="number" tick={{ fill: "#78909C", fontSize: 10 }} tickFormatter={fmt} />
              <YAxis type="category" dataKey="region" tick={{ fill: "#455A64", fontSize: 11, fontWeight: 500 }} width={120} />
              <Tooltip contentStyle={{ background: "#FFF", border: "1px solid #E0E0E0", borderRadius: 10, fontSize: 12 }} formatter={(v) => [fmt(v), ""]} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600 }} />
              <Bar dataKey="plan" fill="#BBDEFB" name="План" radius={[0, 4, 4, 0]} />
              <Bar dataKey="fact" fill="#1565C0" name="Факт" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Pie + Summary */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 24 }}>
          <div className="card anim" style={{ padding: 22 }}>
            <h3 className="playfair" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Факт бўйича улуш</h3>
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={data.orgSummary} dataKey="fact" nameKey="org" cx="50%" cy="50%" outerRadius={85} innerRadius={42}
                  label={({ org, pct }) => `${org} ${pct}%`} labelLine={{ stroke: "#B0BEC5" }} strokeWidth={2} stroke="#fff">
                  {data.orgSummary.map((o, i) => <Cell key={i} fill={o.color} />)}
                </Pie>
                <Tooltip formatter={(v) => [fmt(v), ""]} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="card anim" style={{ padding: 22 }}>
            <h3 className="playfair" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Хулоса</h3>
            <div style={{ fontSize: 13.5, lineHeight: 2.2, color: "#455A64" }}>
              <div>🔴 Жами бажарилиш: <b style={{ color: pctColor(data.totalPct), fontSize: 15 }}>{data.totalPct}%</b></div>
              {(() => { const best = [...data.regionSummary].sort((a, b) => b.pct - a.pct)[0]; return best ? <div>🟢 Энг яхши: <b style={{ color: "#2E7D32" }}>{best.region} — {best.pct}%</b></div> : null; })()}
              {data.orgSummary.map((o, i) => (
                <div key={i} style={{ color: "#455A64" }}>{["🔴","🔵","🟡"][i]} {o.org}: <b style={{ color: o.color }}>{o.pct}%</b></div>
              ))}
              {(() => { const best = [...data.agentSummary].sort((a, b) => b.pct - a.pct)[0]; return best ? <div>👤 Топ агент: <b style={{ color: "#AD1457" }}>{best.agent} — {best.pct}%</b></div> : null; })()}
              {(() => { const zeros = data.regionSummary.filter(r => r.pct === 0); return zeros.length ? <div>⚠️ 0%: <b style={{ color: "#C62828" }}>{zeros.map(r => r.region).join(", ")}</b></div> : null; })()}
            </div>
          </div>
        </div>

        <div style={{ textAlign: "center", padding: "24px 0 8px", color: "#B0BEC5", fontSize: 11, letterSpacing: 1 }}>
          АПРЕЛЬ 2026 · DELI TORG · GRAND TRADING · SIGNUM
        </div>
      </div>
    </div>
  );
}

/* ──────────── MAIN APP ──────────── */
export default function App() {
  const [view, setView] = useState("loading");
  const [data, setData] = useState(null);

  useEffect(() => {
    (async () => {
      const saved = await store.get("dashboard_data");
      if (saved) { setData(saved); setView("dashboard"); }
      else setView("admin");
    })();
  }, []);

  if (view === "loading") {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#FAFBFE" }}>
        <style>{STYLES}</style>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 48 }}>⏳</div>
          <p style={{ color: "#78909C", marginTop: 12 }}>Yuklanmoqda...</p>
        </div>
      </div>
    );
  }

  return (
    <div>
      <style>{STYLES}</style>
      {view === "admin" ? (
        <AdminPanel
          existingDate={data?.uploadDate}
          onPublish={(d) => { setData(d); setView("dashboard"); }}
        />
      ) : (
        <DashboardView data={data} onAdmin={() => setView("admin")} />
      )}
    </div>
  );
}
