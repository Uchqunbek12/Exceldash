import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js"; // Supabase kutubxonasi
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, BarChart, Bar, Cell, Legend, CartesianGrid, PieChart, Pie } from "recharts";

/* 
  ═══════════ SUPABASE SOZLAMALARI ═══════════ 
  Supabase'dan olingan URL va KEY ni bu yerga qo'ying
*/
const SUPABASE_URL = "https://fqgkujhvgvorrdlcjbnf.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZxZ2t1amh2Z3ZvcnJkbGNqYm5mIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MzMwOTgsImV4cCI6MjA5MTIwOTA5OH0.Xa9B3zt6oro-P7ygeG45sQJsK8K5ezX0T1feZ3np4GA";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

/* ═══════════ PLAN (static, monthly) ═══════════ */
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

/* ═══════════ INITIAL FACT DATA (060426.xlsx pre-parsed) ═══════════ */
const INITIAL_DATA = {
  totalPlan: 7400000000, totalFact: 1091110262, totalPct: 14.7, uploadDate: "06.04.2026",
  orgSummary: [
    { org: "DELI TORG", plan: 3564082686, fact: 497115118, pct: 13.9, color: "#E53935" },
    { org: "GRAND TRADING", plan: 1658319104, fact: 499658332, pct: 30.1, color: "#1E88E5" },
    { org: "SIGNUM", plan: 2177598210, fact: 94336812, pct: 4.3, color: "#F9A825" },
  ],
  regionSummary: [
    { region: "Алмазар", category: "Tashkent", plan: 70490991, fact: 4173204, pct: 5.9, plan_akb: 9, fact_tochka: 4 },
    { region: "Андижан", category: "Oblast", plan: 254417434, fact: 4102560, pct: 1.6, plan_akb: 7, fact_tochka: 1 },
    { region: "Бектемир", category: "Tashkent", plan: 485589574, fact: 165098920, pct: 34.0, plan_akb: 3, fact_tochka: 5 },
    { region: "Бухара", category: "Oblast", plan: 489500828, fact: 134252911, pct: 27.4, plan_akb: 14, fact_tochka: 8 },
    { region: "Джиззах", category: "Oblast", plan: 68391236, fact: 4249000, pct: 6.2, plan_akb: 7, fact_tochka: 1 },
    { region: "Каракалпакская р.", category: "Oblast", plan: 117409037, fact: 8094500, pct: 6.9, plan_akb: 6, fact_tochka: 2 },
    { region: "Кашкадарья", category: "Oblast", plan: 208186398, fact: 2730550, pct: 1.3, plan_akb: 9, fact_tochka: 1 },
    { region: "Коканд", category: "Oblast", plan: 109923854, fact: 0, pct: 0.0, plan_akb: 4, fact_tochka: 0 },
    { region: "М.Улугбек", category: "Tashkent", plan: 318163715, fact: 52934041, pct: 16.6, plan_akb: 10, fact_tochka: 21 },
    { region: "Мирабад", category: "Tashkent", plan: 119222263, fact: 25238745, pct: 21.2, plan_akb: 9, fact_tochka: 3 },
    { region: "Навои", category: "Oblast", plan: 254983736, fact: 38803357, pct: 15.2, plan_akb: 7, fact_tochka: 6 },
    { region: "Наманган", category: "Oblast", plan: 364735289, fact: 19798000, pct: 5.4, plan_akb: 8, fact_tochka: 1 },
    { region: "Самарканд", category: "Oblast", plan: 927448182, fact: 27082256, pct: 2.9, plan_akb: 10, fact_tochka: 3 },
    { region: "Сергели", category: "Tashkent", plan: 97454080, fact: 38333287, pct: 39.3, plan_akb: 6, fact_tochka: 7 },
    { region: "Солнечный", category: "Солнечный", plan: 1275817432, fact: 82307796, pct: 6.5, plan_akb: 22, fact_tochka: 10 },
    { region: "Сурхандарья", category: "Oblast", plan: 62789760, fact: 5004000, pct: 8.0, plan_akb: 9, fact_tochka: 1 },
    { region: "Сырдарья", category: "Oblast", plan: 35605535, fact: 0, pct: 0.0, plan_akb: 2, fact_tochka: 0 },
    { region: "Таш. область", category: "Oblast", plan: 129860260, fact: 10238858, pct: 7.9, plan_akb: 12, fact_tochka: 3 },
    { region: "Учтепа", category: "Tashkent", plan: 106458650, fact: 124169114, pct: 116.6, plan_akb: 7, fact_tochka: 3 },
    { region: "Фергана", category: "Oblast", plan: 320018663, fact: 2584500, pct: 0.8, plan_akb: 10, fact_tochka: 1 },
    { region: "Хорезм", category: "Oblast", plan: 315296987, fact: 10733644, pct: 3.4, plan_akb: 10, fact_tochka: 3 },
    { region: "Чиланзар", category: "Tashkent", plan: 495793395, fact: 4856476, pct: 1.0, plan_akb: 12, fact_tochka: 3 },
    { region: "Шайхантахур", category: "Tashkent", plan: 317488178, fact: 252498717, pct: 79.5, plan_akb: 12, fact_tochka: 4 },
    { region: "Юнусабад", category: "Tashkent", plan: 256888113, fact: 65584265, pct: 25.5, plan_akb: 12, fact_tochka: 5 },
    { region: "Яккасарай", category: "Tashkent", plan: 62363406, fact: 130311, pct: 0.2, plan_akb: 6, fact_tochka: 1 },
    { region: "Яшнабад", category: "Tashkent", plan: 135703004, fact: 8111250, pct: 6.0, plan_akb: 7, fact_tochka: 7 },
  ],
  agentSummary: [
    { agent: "Анвар", regions: "Солнечный", plan: 493886798, fact: 18117636, pct: 3.7, plan_akb: 6, fact_tochka: 2 },
    { agent: "Жонибек", regions: "Солнечный", plan: 447102721, fact: 28315828, pct: 6.3, plan_akb: 6, fact_tochka: 1 },
    { agent: "Фаррух", regions: "Солнечный", plan: 334827913, fact: 16435420, pct: 4.9, plan_akb: 10, fact_tochka: 5 },
    { agent: "Ахмаджон", regions: "Учтепа, Чиланзар", plan: 602252045, fact: 129025590, pct: 21.4, plan_akb: 19, fact_tochka: 6 },
    { agent: "Абдулбосид", regions: "Алмазар, Шайхантахур", plan: 387979169, fact: 256671921, pct: 66.2, plan_akb: 21, fact_tochka: 8 },
    { agent: "Абдурахмон", regions: "Яшнабад", plan: 135703004, fact: 8111250, pct: 6.0, plan_akb: 7, fact_tochka: 7 },
    { agent: "Рихсивой", regions: "Бектемир, Мирабад", plan: 604811837, fact: 190337665, pct: 31.5, plan_akb: 12, fact_tochka: 6 },
    { agent: "Умида", regions: "Сергели, Яккасарай", plan: 159817486, fact: 38463598, pct: 24.1, plan_akb: 12, fact_tochka: 7 },
    { agent: "Хусан", regions: "М.Улугбек, Юнусабад", plan: 575051827, fact: 118518306, pct: 20.6, plan_akb: 22, fact_tochka: 17 },
    { agent: "Бобурбек", regions: "Андижан, Коканд, Фергана", plan: 684359951, fact: 6687060, pct: 1.0, plan_akb: 21, fact_tochka: 2 },
    { agent: "Даврон", regions: "Каракалпакия, Наманган", plan: 482144326, fact: 27892500, pct: 5.8, plan_akb: 14, fact_tochka: 2 },
    { agent: "Расул", regions: "Джиззах, Кашкадарья, Сырдарья", plan: 312183170, fact: 6979550, pct: 2.2, plan_akb: 18, fact_tochka: 2 },
    { agent: "Хамид", regions: "Самарканд, Хорезм", plan: 1242745169, fact: 37815900, pct: 3.0, plan_akb: 20, fact_tochka: 6 },
    { agent: "Шахбоз", regions: "Бухара, Навои, Сурхандарья", plan: 807274324, fact: 178060268, pct: 22.1, plan_akb: 30, fact_tochka: 7 },
    { agent: "Жалолиддин", regions: "Таш. область", plan: 129860260, fact: 9361038, pct: 7.2, plan_akb: 12, fact_tochka: 3 },
  ],
  daily: [
    { date: "01.04.2026", dateLabel: "01 Апр", DT: 146023772, GT: 173613622, SG: 19579412, total: 339216806 },
    { date: "02.04.2026", dateLabel: "02 Апр", DT: 114575287, GT: 218493658, SG: 38395352, total: 371464297 },
    { date: "03.04.2026", dateLabel: "03 Апр", DT: 235148464, GT: 30155847, SG: 27916700, total: 293221011 },
    { date: "04.04.2026", dateLabel: "04 Апр", DT: 1367595, GT: 34820837, SG: 6642068, total: 42830500 },
    { date: "06.04.2026", dateLabel: "06 Апр", DT: 0, GT: 42574368, SG: 1803280, total: 44377648 },
  ],
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
    rF, 
    aF, 
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

/* ═══════════ SUPABASE STORAGE ═══════════ */
const sto = {
  async get() {
    try {
      const { data, error } = await supabase.from("dashboard").select("content").eq("id", 1).single();
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
  const handleFile = f => { if (f && /\.xlsx?$/i.test(f.name)) { setFile(f); setErr(""); } else setErr("Faqat .xlsx fayl!"); };
  const publish = async () => {
    if (!file) return; setLoading(true); setErr("");
    try {
      const parsed = parseExcel(await file.arrayBuffer());
      const data = buildData(parsed);
      const ok = await sto.set(data);
      if (ok) {
        setSuccess(`✅ Muvaffaqiyat! Ma'lumotlar serverga saqlandi.`);
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
          <div className={`uz ${dragOver ? "dr" : ""}`}
            onClick={() => document.getElementById("fu").click()}
            onDragOver={e => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={e => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
            style={{ border: `3px dashed ${dragOver ? "#1565C0" : "#BBDEFB"}`, borderRadius: 14, padding: "32px 16px", cursor: "pointer", textAlign: "center", background: dragOver ? "#E3F2FD" : "#F8FBFF", transition: "all 0.2s" }}>
            <input id="fu" type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
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

  useEffect(() => { 
    (async () => { 
      const s = await sto.get(); 
      if (s) setData(s); 
    })(); 
  }, []);

  const filtered = useMemo(() => catFilter === "Все" ? data.regionSummary : data.regionSummary.filter(r => r.category === catFilter), [catFilter, data]);
  const fP = filtered.reduce((s, r) => s + r.plan, 0), fF = filtered.reduce((s, r) => s + r.fact, 0), fPct = fP ? (fF / fP * 100).toFixed(1) : 0;
  const top5 = useMemo(() => [...data.regionSummary].sort((a, b) => b.pct - a.pct).slice(0, 5), [data]);
  const bot5 = useMemo(() => [...data.regionSummary].sort((a, b) => a.pct - b.pct).slice(0, 5), [data]);

  return (
    <div style={{ fontFamily: "'Merriweather Sans',sans-serif", background: "#FAFBFE", color: "#1a1a2e", minHeight: "100vh" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700;800;900&family=Merriweather+Sans:wght@300;400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0}
        .hd{background:linear-gradient(135deg,#1565C0,#0D47A1 40%,#01579B);padding:22px 28px;color:#fff;position:relative;overflow:hidden}
        .hd::before{content:'';position:absolute;top:-50%;right:-10%;width:400px;height:400px;background:radial-gradient(circle,rgba(255,255,255,.08),transparent 70%);border-radius:50%}
        .cd{background:#fff;border:1px solid #E8EAF0;border-radius:12px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,.04);transition:box-shadow .2s}
        .cd:hover{box-shadow:0 4px 20px rgba(0,0,0,.08)}
        .kp{background:#fff;border-radius:14px;padding:20px;border:1px solid #E8EAF0;box-shadow:0 2px 8px rgba(0,0,0,.03);position:relative;overflow:hidden}
        .st{width:5px;border-radius:3px;position:absolute;left:0;top:12px;bottom:12px}
        .bt{padding:9px 20px;border-radius:24px;border:2px solid #E0E0E0;background:#fff;color:#546E7A;cursor:pointer;font-family:'Merriweather Sans';font-size:13px;font-weight:600;transition:all .25s}
        .bt.a{background:#1565C0;color:#fff;border-color:#1565C0;box-shadow:0 3px 12px rgba(21,101,192,.3)}
        .bt:hover:not(.a){border-color:#1565C0;color:#1565C0}
        table{width:100%;border-collapse:separate;border-spacing:0}
        thead{background:#F5F7FA}
        th{text-align:left;padding:12px;font-size:10.5px;text-transform:uppercase;letter-spacing:1.2px;color:#78909C;font-weight:700;border-bottom:2px solid #E0E0E0}
        td{padding:11px 12px;font-size:13px;border-bottom:1px solid #F0F0F0}
        tr:hover td{background:#F8FAFF}
        .br{height:7px;border-radius:4px;background:#ECEFF1;overflow:hidden;min-width:80px}
        .fl{height:100%;border-radius:4px;transition:width .5s ease-out}
        .bg{padding:3px 8px;border-radius:4px;font-size:10px;font-weight:700;text-transform:uppercase}
        .pf{font-family:'Playfair Display',serif}
        @keyframes fu{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
        .an{animation:fu .45s ease forwards}
        .ab{position:absolute;top:14px;right:16px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;font-family:'Merriweather Sans';backdrop-filter:blur(4px);transition:all .2s;z-index:2}
        .ab:hover{background:rgba(255,255,255,.3)}
      `}</style>

      {showAdmin && <AdminModal onClose={() => setShowAdmin(false)} onPublish={d => setData(d)} uploadDate={data.uploadDate} />}

      {/* HEADER */}
      <div className="hd">
        <button className="ab" onClick={() => setShowAdmin(true)}>⚙️ Admin</button>
        <div style={{ position: "relative", zIndex: 1 }}>
          <h1 className="pf" style={{ fontSize: 26, fontWeight: 800 }}>Апрель 2026 — Сотув Аналитикаси</h1>
          <p style={{ opacity: .75, fontSize: 13, marginTop: 6 }}>
            {data.daily.length} кун · 3 ташкилот · {data.regionSummary.length} регион · {data.agentSummary.length} агент
            {data.uploadDate && <span> · Янгиланган: {data.uploadDate}</span>}
          </p>
        </div>
      </div>

      <div style={{ padding: "22px 24px" }}>
        {/* KPIs */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16, marginBottom: 24 }} className="an">
          {[
            { l: "Умумий план", v: fmt(data.totalPlan), s: "Апрель 2026", c: "#1565C0", i: "🎯" },
            { l: "Факт сотув", v: fmt(data.totalFact), s: data.totalPct + "% бажарилган", c: "#2E7D32", i: "💰" },
            { l: "Қолган сумма", v: fmt(data.totalPlan - data.totalFact), s: (100 - data.totalPct).toFixed(1) + "% қолди", c: "#E65100", i: "📉" },
            { l: "Торг. точкалар", v: data.regionSummary.reduce((s, r) => s + r.fact_tochka, 0), s: "АКБ: " + data.regionSummary.reduce((s, r) => s + r.plan_akb, 0), c: "#AD1457", i: "🏪" },
          ].map((k, i) => (
            <div key={i} className="kp"><div className="st" style={{ background: k.c }} />
              <div style={{ paddingLeft: 12 }}>
                <div style={{ fontSize: 10.5, color: "#90A4AE", textTransform: "uppercase", letterSpacing: 1.5, fontWeight: 700 }}>{k.l}</div>
                <div className="pf" style={{ fontSize: 26, fontWeight: 800, color: k.c, marginTop: 5 }}>{k.v}</div>
                <div style={{ fontSize: 12, color: "#78909C", marginTop: 3 }}>{k.s}</div>
              </div>
              <div style={{ position: "absolute", top: 10, right: 14, fontSize: 34, opacity: .1 }}>{k.i}</div>
            </div>
          ))}
        </div>

        {/* ORG CARDS */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16, marginBottom: 24 }} className="an">
          {data.orgSummary.map((o, i) => {
            const r = 40, circ = 2 * Math.PI * r, off = circ - (Math.min(o.pct, 100) / 100) * circ;
            return (
              <div key={i} className="cd" style={{ display: "flex", alignItems: "center", gap: 18, borderTop: `4px solid ${o.color}` }}>
                <div style={{ position: "relative", width: 90, height: 90, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <svg width={90} height={90} style={{ transform: "rotate(-90deg)" }}>
                    <circle cx={45} cy={45} r={r} fill="none" stroke="#ECEFF1" strokeWidth={7} />
                    <circle cx={45} cy={45} r={r} fill="none" stroke={o.color} strokeWidth={7} strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round" />
                  </svg>
                  <span className="pf" style={{ position: "absolute", fontSize: 16, fontWeight: 800, color: o.color }}>{o.pct}%</span>
                </div>
                <div>
                  <div className="pf" style={{ fontSize: 16, fontWeight: 700, color: "#263238" }}>{o.org}</div>
                  <div style={{ fontSize: 12, color: "#90A4AE", marginTop: 5 }}>План: <b style={{ color: "#455A64" }}>{fmt(o.plan)}</b></div>
                  <div style={{ fontSize: 12, color: "#90A4AE" }}>Факт: <b style={{ color: o.color }}>{fmt(o.fact)}</b></div>
                </div>
              </div>
            );
          })}
        </div>

        {/* LINE CHART */}
        <div className="cd an" style={{ marginBottom: 24, padding: 22 }}>
          <h3 className="pf" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Кунлик сотув динамикаси</h3>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={data.daily}>
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis dataKey="dateLabel" tick={{ fill: "#78909C", fontSize: 12, fontWeight: 600 }} />
              <YAxis tick={{ fill: "#78909C", fontSize: 10 }} tickFormatter={fmt} />
              <Tooltip contentStyle={{ background: "#FFF", border: "1px solid #E0E0E0", borderRadius: 10, fontSize: 12, boxShadow: "0 4px 16px rgba(0,0,0,.06)" }} formatter={v => [fmt(v), ""]} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600 }} />
              <Line type="monotone" dataKey="total" stroke="#1565C0" strokeWidth={3} name="Жами" dot={{ r: 5, fill: "#1565C0", stroke: "#fff", strokeWidth: 2 }} />
              <Line type="monotone" dataKey="DT" stroke="#E53935" strokeWidth={2} name="DELI TORG" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="GT" stroke="#1E88E5" strokeWidth={2} name="GRAND TRADING" dot={{ r: 3 }} />
              <Line type="monotone" dataKey="SG" stroke="#F9A825" strokeWidth={2} name="SIGNUM" dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* TOP / BOTTOM */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }} className="an">
          <div className="cd" style={{ borderTop: "4px solid #2E7D32" }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#2E7D32" }}>🏆 Топ-5 регионлар</h3>
            {top5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#E8F5E9" : "#FAFAFA" }}>
                <span style={{ fontSize: 20, width: 30, textAlign: "center" }}>{["🥇","🥈","🥉","4️⃣","5️⃣"][i]}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F" }}>{r.region}</span>
                <span className="pf" style={{ fontWeight: 800, color: pc(r.pct), fontSize: 15 }}>{r.pct}%</span>
              </div>
            ))}
          </div>
          <div className="cd" style={{ borderTop: "4px solid #C62828" }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#C62828" }}>⚠️ Паст-5 регионлар</h3>
            {bot5.map((r, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, padding: "7px 10px", borderRadius: 8, background: i === 0 ? "#FFEBEE" : "#FAFAFA" }}>
                <span style={{ width: 30, textAlign: "center", fontWeight: 800, color: "#C62828", fontSize: 15 }}>{i + 1}</span>
                <span style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#37474F" }}>{r.region}</span>
                <span className="pf" style={{ fontWeight: 800, color: "#C62828", fontSize: 15 }}>{r.pct}%</span>
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
            {catFilter !== "Все" && <span style={{ marginLeft: "auto", fontSize: 12, color: "#78909C" }}>
              План: <b style={{ color: "#1565C0" }}>{fmt(fP)}</b> · Факт: <b style={{ color: "#2E7D32" }}>{fmt(fF)}</b> · <b style={{ color: pc(+fPct) }}>{fPct}%</b>
            </span>}
          </div>
        )}

        {/* REGION TABLE */}
        {tab === "regions" && (
          <div className="cd an" style={{ padding: 0, overflowX: "auto" }}>
            <table><thead><tr>
              <th>#</th><th>Регион</th><th>Категория</th><th style={{ textAlign: "right" }}>План</th><th style={{ textAlign: "right" }}>Факт</th>
              <th style={{ textAlign: "center" }}>АКБ</th><th style={{ textAlign: "center" }}>Точки</th><th style={{ width: 110 }}>Прогресс</th><th style={{ textAlign: "right" }}>%</th>
            </tr></thead><tbody>
              {[...filtered].sort((a, b) => b.pct - a.pct).map((r, i) => (
                <tr key={i}>
                  <td style={{ color: "#B0BEC5", fontWeight: 700 }}>{i + 1}</td>
                  <td style={{ fontWeight: 600, color: "#263238" }}>{r.region}</td>
                  <td><span className="bg" style={{ background: r.category === "Tashkent" ? "#E3F2FD" : r.category === "Солнечный" ? "#FFF8E1" : "#E8F5E9", color: r.category === "Tashkent" ? "#1565C0" : r.category === "Солнечный" ? "#F57F17" : "#2E7D32" }}>{r.category}</span></td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#546E7A" }}>{fmt(r.plan)}</td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 600, color: "#2E7D32" }}>{fmt(r.fact)}</td>
                  <td style={{ textAlign: "center", color: "#78909C" }}>{r.plan_akb}</td>
                  <td style={{ textAlign: "center", fontWeight: 700, color: "#E65100" }}>{r.fact_tochka}</td>
                  <td><div className="br"><div className="fl" style={{ width: Math.min(r.pct, 100) + "%", background: pc(r.pct) }} /></div></td>
                  <td style={{ textAlign: "right" }}><span className="pf" style={{ fontWeight: 800, fontSize: 13, color: pc(r.pct), background: pb(r.pct), padding: "3px 7px", borderRadius: 6 }}>{r.pct}%</span></td>
                </tr>
              ))}
            </tbody></table>
          </div>
        )}

        {/* AGENT TABLE */}
        {tab === "agents" && (
          <div className="cd an" style={{ padding: 0, overflowX: "auto" }}>
            <table><thead><tr>
              <th>#</th><th>Агент</th><th>Регионлар</th><th style={{ textAlign: "right" }}>План</th><th style={{ textAlign: "right" }}>Факт</th>
              <th style={{ textAlign: "center" }}>АКБ</th><th style={{ textAlign: "center" }}>Точки</th><th style={{ width: 110 }}>Прогресс</th><th style={{ textAlign: "right" }}>%</th>
            </tr></thead><tbody>
              {[...data.agentSummary].sort((a, b) => b.pct - a.pct).map((a, i) => (
                <tr key={i}>
                  <td style={{ color: "#B0BEC5", fontWeight: 700 }}>{i + 1}</td>
                  <td style={{ fontWeight: 700, color: "#263238" }}>{a.agent}</td>
                  <td style={{ fontSize: 11, color: "#90A4AE" }}>{a.regions}</td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#546E7A" }}>{fmt(a.plan)}</td>
                  <td style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 600, color: "#2E7D32" }}>{fmt(a.fact)}</td>
                  <td style={{ textAlign: "center", color: "#78909C" }}>{a.plan_akb}</td>
                  <td style={{ textAlign: "center", fontWeight: 700, color: "#E65100" }}>{a.fact_tochka}</td>
                  <td><div className="br"><div className="fl" style={{ width: Math.min(a.pct, 100) + "%", background: pc(a.pct) }} /></div></td>
                  <td style={{ textAlign: "right" }}><span className="pf" style={{ fontWeight: 800, fontSize: 13, color: pc(a.pct), background: pb(a.pct), padding: "3px 7px", borderRadius: 6 }}>{a.pct}%</span></td>
                </tr>
              ))}
            </tbody></table>
          </div>
        )}

        {/* BAR CHART */}
        <div className="cd an" style={{ marginTop: 24, padding: 22 }}>
          <h3 className="pf" style={{ fontSize: 18, fontWeight: 700, marginBottom: 18, color: "#263238" }}>Регионлар — План ва Факт</h3>
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={[...data.regionSummary].sort((a, b) => b.fact - a.fact).slice(0, 10)} layout="vertical">
              <CartesianGrid strokeDasharray="4 4" stroke="#E0E0E0" />
              <XAxis type="number" tick={{ fill: "#78909C", fontSize: 10 }} tickFormatter={fmt} />
              <YAxis type="category" dataKey="region" tick={{ fill: "#455A64", fontSize: 11, fontWeight: 600 }} width={100} />
              <Tooltip formatter={v => [fmt(v), ""]} />
              <Legend wrapperStyle={{ fontSize: 11, fontWeight: 600 }} />
              <Bar dataKey="plan" fill="#BBDEFB" name="План" radius={[0, 4, 4, 0]} />
              <Bar dataKey="fact" fill="#1565C0" name="Факт" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* PIE + SUMMARY */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 24 }}>
          <div className="cd an" style={{ padding: 22 }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Факт бўйича улуш</h3>
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={data.orgSummary} dataKey="fact" nameKey="org" cx="50%" cy="50%" outerRadius={85} innerRadius={42}
                  label={({ org, pct }) => `${org} ${pct}%`} labelLine={{ stroke: "#B0BEC5" }} strokeWidth={2} stroke="#fff">
                  {data.orgSummary.map((o, i) => <Cell key={i} fill={o.color} />)}
                </Pie>
                <Tooltip formatter={v => [fmt(v), ""]} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="cd an" style={{ padding: 22 }}>
            <h3 className="pf" style={{ fontSize: 16, fontWeight: 700, marginBottom: 14, color: "#263238" }}>Хулоса</h3>
            <div style={{ fontSize: 13.5, lineHeight: 2.2, color: "#455A64" }}>
              <div>🔴 Жами: <b style={{ color: pc(data.totalPct), fontSize: 15 }}>{data.totalPct}%</b></div>
              {(() => { const b = [...data.regionSummary].sort((a, b) => b.pct - a.pct)[0]; return b ? <div>🟢 Энг яхши: <b style={{ color: "#2E7D32" }}>{b.region} — {b.pct}%</b></div> : null; })()}
              {data.orgSummary.map((o, i) => <div key={i}>{["🔴","🔵","🟡"][i]} {o.org}: <b style={{ color: o.color }}>{o.pct}%</b></div>)}
              {(() => { const b = [...data.agentSummary].sort((a, b) => b.pct - a.pct)[0]; return b ? <div>👤 Топ: <b style={{ color: "#AD1457" }}>{b.agent} — {b.pct}%</b></div> : null; })()}
              {(() => { const z = data.regionSummary.filter(r => r.pct === 0); return z.length ? <div>⚠️ 0%: <b style={{ color: "#C62828" }}>{z.map(r => r.region).join(", ")}</b></div> : null; })()}
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
