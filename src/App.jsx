import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

const LOGO = "/logo.jpg";

// ═══ SUPABASE DB ═══
const db = {
  async getEmployees() {
    const { data } = await supabase.from("employees").select("*").eq("active", true).order("id");
    return (data || []).map(e => ({ id: String(e.id), name: e.name, position: e.position, salary: parseFloat(e.salary), empType: e.emp_type || "weekly" }));
  },
  async upsertEmployee(emp) {
    await supabase.from("employees").upsert({ id: emp.id, name: emp.name, position: emp.position, salary: emp.salary, active: true, emp_type: emp.empType || "weekly", updated_at: new Date().toISOString() });
  },
  async deleteEmployee(id) {
    await supabase.from("employees").update({ active: false }).eq("id", id);
  },
  async getClockEntries() {
    const { data } = await supabase.from("clock_entries").select("*").order("date", { ascending: false });
    return (data || []).map(e => ({ id: e.id, employeeId: String(e.employee_id), date: String(e.date).slice(0,10), checkIn: e.check_in, lunchOut: e.lunch_out, lunchIn: e.lunch_in, checkOut: e.check_out, punches: e.punches }));
  },
  async insertClockEntries(entries) {
    const rows = entries.map(e => ({ id: e.id, employee_id: e.employeeId, date: e.date, check_in: e.checkIn, lunch_out: e.lunchOut, lunch_in: e.lunchIn, check_out: e.checkOut, punches: e.punches }));
    await supabase.from("clock_entries").upsert(rows, { onConflict: "id" });
  },
  async updateClockEntry(entry) {
    const { error } = await supabase.from("clock_entries").update({ check_in: entry.checkIn, lunch_out: entry.lunchOut, lunch_in: entry.lunchIn, check_out: entry.checkOut }).eq("id", entry.id);
    if (error) {
      await supabase.from("clock_entries").delete().eq("id", entry.id);
      await supabase.from("clock_entries").insert({ id: entry.id, employee_id: entry.employeeId, date: entry.date, check_in: entry.checkIn, lunch_out: entry.lunchOut, lunch_in: entry.lunchIn, check_out: entry.checkOut, punches: 4 });
    }
  },
  async deleteClockEntry(id) { await supabase.from("clock_entries").delete().eq("id", id); },
  async deleteAllClockEntries() { await supabase.from("clock_entries").delete().neq("id", ""); },
  async getHolidays() {
    const { data } = await supabase.from("holidays").select("*").order("date");
    return (data || []).map(h => ({ id: h.id, date: String(h.date).slice(0,10), name: h.name }));
  },
  async addHoliday(h) { await supabase.from("holidays").insert({ date: h.date, name: h.name }); },
  async deleteHoliday(id) { await supabase.from("holidays").delete().eq("id", id); },
  async getPayrolls() {
    const { data } = await supabase.from("payrolls").select("*").order("id", { ascending: false });
    return data || [];
  },
  async getPayrollRows(pid) {
    const { data } = await supabase.from("payroll_rows").select("*").eq("payroll_id", pid);
    return (data || []).map(r => ({ employeeId: r.employee_id, name: r.name, position: r.position, salary: +r.salary, daily: +r.daily, hourly: +r.hourly, days: r.days, effectiveHrs: +r.effective_hrs, baseSalary: +r.base_salary, ot: { 0.25: +r.ot_25, 0.5: +r.ot_50, 0.75: +r.ot_75, 1.0: +r.ot_100 }, otPay: +r.ot_pay, ihssTotal: +(r.ihss_total||0), rap: +(r.rap||0), fuel: +r.fuel, vacation: +r.vacation, incapacity: +r.incapacity, advance: +r.advance, dec4: +r.dec4, dec3: +r.dec3, otherDed: +r.other_ded, totalEarned: +r.total_earned, totalDeductions: +r.total_deductions, netPay: +r.net_pay }));
  },
  async savePayroll(payroll) {
    const { data: ex } = await supabase.from("payrolls").select("id").eq("period", payroll.period).maybeSingle();
    let pid;
    if (ex) {
      pid = ex.id;
      await supabase.from("payroll_rows").delete().eq("payroll_id", pid);
      await supabase.from("payrolls").update({ week_num: payroll.weekNum, date_from: payroll.from, date_to: payroll.to, total_earned: payroll.rows.reduce((s,r) => s+r.totalEarned, 0), total_deductions: payroll.rows.reduce((s,r) => s+r.totalDeductions, 0), total_net: payroll.rows.reduce((s,r) => s+r.netPay, 0) }).eq("id", pid);
    } else {
      const { data } = await supabase.from("payrolls").insert({ period: payroll.period, week_num: payroll.weekNum, date_from: payroll.from, date_to: payroll.to, total_earned: payroll.rows.reduce((s,r) => s+r.totalEarned, 0), total_deductions: payroll.rows.reduce((s,r) => s+r.totalDeductions, 0), total_net: payroll.rows.reduce((s,r) => s+r.netPay, 0) }).select("id").single();
      pid = data.id;
    }
    const rows = payroll.rows.map(r => ({ payroll_id: pid, employee_id: r.employeeId, name: r.name, position: r.position, salary: r.salary, daily: r.daily, hourly: r.hourly, days: r.days, effective_hrs: r.effectiveHrs, base_salary: r.baseSalary, ot_25: r.ot[0.25], ot_50: r.ot[0.5], ot_75: r.ot[0.75], ot_100: r.ot[1.0], ot_pay: r.otPay, ihss_total: r.ihssTotal||0, rap: r.rap||0, fuel: r.fuel, vacation: r.vacation, incapacity: r.incapacity, advance: r.advance, dec4: r.dec4, dec3: r.dec3, other_ded: r.otherDed, total_earned: r.totalEarned, total_deductions: r.totalDeductions, net_pay: r.netPay }));
    await supabase.from("payroll_rows").insert(rows);
  },
  async deletePayroll(id) { await supabase.from("payroll_rows").delete().eq("payroll_id", id); await supabase.from("payrolls").delete().eq("id", id); },
};

// ═══ HELPERS ═══
function getScheduledHours(dow) { return dow >= 1 && dow <= 4 ? 9 : dow === 5 ? 8 : 0; }

// IHSS & RAP
const IHSS = { EM_TECHO: 11903.13, EM_TASA: 0.025, IVM_TECHO: 11903.13, IVM_TASA: 0.025 };
function calcIHSS_monthly(sal) { const b = Math.min(sal, IHSS.EM_TECHO); return { em: b*IHSS.EM_TASA, ivm: b*IHSS.IVM_TASA, total: b*IHSS.EM_TASA + b*IHSS.IVM_TASA }; }
function calcIHSS_biweekly(sal) { const m = calcIHSS_monthly(sal); return { em: m.em/2, ivm: m.ivm/2, total: m.total/2 }; }
function calcRAP_monthly(sal) { const exc = Math.max(0, sal - IHSS.IVM_TECHO); const f = exc * 0.015; return { feo3: f, fio3: f, employeeTotal: f*2, rl: sal*0.04, grandTotal: sal*0.04 + f*2 }; }

function isLastWeekOfMonth(from, to) {
  const s = new Date(from+"T12:00:00"), e = new Date(to+"T12:00:00");
  for (let d = new Date(s); d <= e; d.setDate(d.getDate()+1)) {
    const last = new Date(d.getFullYear(), d.getMonth()+1, 0).getDate();
    if (d.getDate() === last) return true;
  }
  return false;
}
function isSecondWeekOfMonth(from) { const d = new Date(from+"T12:00:00").getDate(); return d >= 8 && d <= 14; }

function formatL(n) { if (n==null||isNaN(n)) return "L. 0.00"; return "L. "+Number(n).toLocaleString("es-HN",{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fN(n) { return (!n||isNaN(n)||n===0) ? "" : Number(n).toLocaleString("es-HN",{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fmtTime(t) { return t ? new Date(t).toLocaleTimeString("es-HN",{hour:"2-digit",minute:"2-digit"}) : "—"; }
function fmtDate(d) { return new Date(d).toLocaleDateString("es-HN",{weekday:"short",day:"2-digit",month:"short"}); }

// ═══ CLOCK PARSER ═══
function parseClockCSV(csvText) {
  const text = csvText.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const sep = lines[0].includes(";") ? ";" : ",";
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(sep);
    if (cols.length < 2) continue;
    const time = cols[0].trim(), userId = cols[1].trim();
    if (!userId) continue;
    if (cols.length >= 10) { const ev = cols[9].trim(); if (ev==="Acceso denegado"||ev==="Dispositivo inicializado") continue; }
    let m = time.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) continue;
    const [,p1,p2,yyyy,hh,mi,ss] = m;
    let dd,mm;
    if (parseInt(p1)>12){dd=p1;mm=p2} else if(parseInt(p2)>12){mm=p1;dd=p2} else{dd=p1;mm=p2}
    dd=String(dd).padStart(2,"0"); mm=String(mm).padStart(2,"0");
    records.push({ userId, dateStr:`${yyyy}-${mm}-${dd}`, dt:new Date(+yyyy,+mm-1,+dd,+hh,+mi,+(ss||0)) });
  }
  return groupPunches(records);
}
function parseClockXLS(data) {
  try { 
    const wb = XLSX.read(data, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    
    // Try two approaches: raw values and formatted values
    const jsonRaw = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:""});
    const jsonFmt = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:""});
    
    console.log("XLS rows:", jsonRaw.length, "cols:", jsonRaw[0]?.length);
    if (jsonRaw.length > 1) console.log("Row1 raw:", jsonRaw[1]?.slice(0,3), "type:", typeof jsonRaw[1]?.[0]);
    if (jsonFmt.length > 1) console.log("Row1 fmt:", jsonFmt[1]?.slice(0,3), "type:", typeof jsonFmt[1]?.[0]);
    
    // Try raw first (serial numbers), then formatted (text dates)
    let records = parseXLSRows(jsonRaw);
    if (records.length === 0) records = parseXLSRows(jsonFmt);
    console.log("XLS total parsed:", records.length);
    return groupPunches(records);
  } catch(e) { console.error("parseClockXLS error:", e); return []; }
}
function parseXLSRows(json) {
  if (!json || json.length < 2) return [];
  const records = [];
  for (let i = 1; i < json.length; i++) {
    const row = json[i]; if (!row) continue;
    const rawTime = row[0], userId = String(row[1] ?? "").trim();
    if (!userId) continue;
    // Filter "Acceso denegado" if column 9 exists
    if (row.length >= 10) { const ev = String(row[9] ?? "").trim(); if (ev === "Acceso denegado" || ev === "Dispositivo inicializado") continue; }
    let dt = null, dateStr = null;
    if (typeof rawTime === "number" && rawTime > 10000) {
      // Excel serial number
      const epoch = new Date(1899, 11, 30); const d = new Date(epoch.getTime() + rawTime * 86400000);
      if (!isNaN(d.getTime())) { dt = d; dateStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
    } else if (rawTime instanceof Date) {
      dt = rawTime; dateStr = `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
    } else {
      const time = String(rawTime ?? "").trim();
      const m = time.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (!m) continue;
      const [,p1,p2,yyyy,hh,mi,ss] = m; let dd,mm;
      if (parseInt(p1)>12){dd=p1;mm=p2} else if(parseInt(p2)>12){mm=p1;dd=p2} else{dd=p1;mm=p2}
      dd=String(dd).padStart(2,"0"); mm=String(mm).padStart(2,"0");
      dt=new Date(+yyyy,+mm-1,+dd,+hh,+mi,+(ss||0)); dateStr=`${yyyy}-${mm}-${dd}`;
    }
    if (!dt || !dateStr || isNaN(dt.getTime())) continue;
    records.push({userId, dateStr, dt});
  }
  return records;
}
function groupPunches(records) {
  const groups={};records.forEach(r=>{const k=`${r.userId}_${r.dateStr}`;if(!groups[k])groups[k]={userId:r.userId,date:r.dateStr,punches:[]};groups[k].punches.push(r.dt)});
  const entries=[];
  Object.values(groups).forEach(g=>{
    g.punches.sort((a,b)=>a-b);
    // Deduplicate (within 2 min)
    const dd=[g.punches[0]];
    for(let i=1;i<g.punches.length;i++)if(g.punches[i]-g.punches[i-1]>120000)dd.push(g.punches[i]);
    
    // Classify by time-of-day ranges:
    // 6:00 - 10:00 → checkIn (entrada)
    // 10:00 - 12:30 → lunchOut (salida almuerzo)
    // 12:30 - 15:00 → lunchIn (regreso almuerzo)
    // 15:00 - 23:59 → checkOut (salida)
    let checkIn=null,lunchOut=null,lunchIn=null,checkOut=null;
    
    dd.forEach(p=>{
      const h=p.getHours()+p.getMinutes()/60;
      if(h>=6&&h<10&&!checkIn) checkIn=p;
      else if(h>=10&&h<12.5&&!lunchOut) lunchOut=p;
      else if(h>=12.5&&h<15&&!lunchIn) lunchIn=p;
      else if(h>=15&&!checkOut) checkOut=p;
      // Handle edge cases: if all slots for a range are filled, use next available
      else if(!checkIn) checkIn=p;
      else if(!lunchOut) lunchOut=p;
      else if(!lunchIn) lunchIn=p;
      else if(!checkOut) checkOut=p;
    });

    // For 4+ punches with positional fallback (if time-based didn't work well)
    if(dd.length>=4&&!checkIn){checkIn=dd[0];lunchOut=dd[1];lunchIn=dd[2];checkOut=dd[dd.length-1]}
    
    const e={
      id:`${g.userId}_${g.date}`,
      employeeId:g.userId,
      date:g.date,
      checkIn:checkIn?checkIn.toISOString():null,
      lunchOut:lunchOut?lunchOut.toISOString():null,
      lunchIn:lunchIn?lunchIn.toISOString():null,
      checkOut:checkOut?checkOut.toISOString():null,
      punches:dd.length
    };
    entries.push(e);
  });
  return entries.sort((a,b)=>a.date.localeCompare(b.date));
}
function calcDayHours(entry) {
  if(!entry.checkIn||!entry.checkOut)return 0;
  let cin=new Date(entry.checkIn);const cout=new Date(entry.checkOut);
  // Grace: 7:50-8:00am → cap at 8:00. Before 7:50 counts (early OT)
  const am8=new Date(cin);am8.setHours(8,0,0,0);const am750=new Date(cin);am750.setHours(7,50,0,0);
  if(cin>=am750&&cin<am8)cin=am8;
  let ms=cout-cin;
  if(entry.lunchOut&&entry.lunchIn)ms-=(new Date(entry.lunchIn)-new Date(entry.lunchOut));
  else if(entry.lunchOut)ms-=3600000;
  else if(ms>5*3600000)ms-=3600000;
  return Math.max(0,ms/3600000);
}

// ═══ PRINT ═══
function printPayroll(payroll) {
  const rows=payroll.rows;const t={sal:rows.reduce((s,r)=>s+(r.salary||0),0),base:rows.reduce((s,r)=>s+r.baseSalary,0),ot:rows.reduce((s,r)=>s+r.otPay,0),ihss:rows.reduce((s,r)=>s+(r.ihssTotal||0),0),rap:rows.reduce((s,r)=>s+(r.rap||0),0),earned:rows.reduce((s,r)=>s+r.totalEarned,0),ded:rows.reduce((s,r)=>s+r.totalDeductions,0),net:rows.reduce((s,r)=>s+r.netPay,0)};
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Planilla ${payroll.period}</title><style>@page{size:landscape;margin:10mm}*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:8pt}table{width:100%;border-collapse:collapse;margin-top:6px}th{background:#0a2351;color:#fff;padding:4px 3px;font-size:6.5pt;text-transform:uppercase;border:1px solid #0d2d6b;text-align:center}td{padding:3px;border:1px solid #c8d6e5;font-size:7.5pt}.r{text-align:right;font-family:'Courier New',monospace}.c{text-align:center}.name{font-weight:600;white-space:nowrap}.total-row{background:#e8eef6;font-weight:700}.total-row td{border-top:2px solid #0a2351}.net{color:#0a6847;font-weight:700}.ot-val{color:#b91c1c}.ihss{color:#6d28d9}.header{text-align:center;margin-bottom:10px}.header img{height:36px;margin-bottom:4px}.signatures{margin-top:30px;display:flex;justify-content:space-between}.sig-box{text-align:center;width:180px}.sig-line{border-top:1px solid #000;margin-top:40px;padding-top:3px;font-size:8pt}@media print{.no-print{display:none!important}}.no-print{position:fixed;top:10px;right:10px;z-index:999}.btn{padding:8px 20px;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;margin-right:6px}</style></head><body>
<div class="no-print"><button class="btn" style="background:#0a2351;color:#fff" onclick="window.print()">🖨️ Imprimir</button><button class="btn" style="background:#64748b;color:#fff" onclick="window.close()">✕ Cerrar</button></div>
<div class="header"><img src="${LOGO}" alt="Horeb"/><div style="font-size:10pt;color:#1a3a6b">Planilla de Empleados</div><div style="font-size:9pt;font-weight:bold;color:#0a2351;margin-top:3px">${payroll.period}</div></div>
<table><thead><tr><th rowspan="2">Cód</th><th rowspan="2">Nombre</th><th rowspan="2">Pos.</th><th rowspan="2">Sal.M.</th><th rowspan="2">Sal.D.</th><th rowspan="2">Días</th><th rowspan="2">Salario</th><th colspan="4">Horas Extras</th><th rowspan="2">Sal/Hr</th><th rowspan="2">Tot.OT</th><th rowspan="2">IHSS</th><th rowspan="2">RAP</th><th rowspan="2">Comb.</th><th rowspan="2">Adel.</th><th rowspan="2">Deveng.</th><th rowspan="2">Tot.Ded.</th><th rowspan="2">Neto</th></tr><tr><th>25%</th><th>50%</th><th>75%</th><th>100%</th></tr></thead><tbody>
${rows.map(r=>`<tr><td class="c">${r.employeeId}</td><td class="name">${r.name}</td><td>${r.position}</td><td class="r">${fN(r.salary)}</td><td class="r">${fN(r.daily)}</td><td class="c">${r.days}</td><td class="r">${fN(r.baseSalary)}</td><td class="r${r.ot[0.25]>0?' ot-val':''}">${r.ot[0.25]>0?r.ot[0.25].toFixed(1):''}</td><td class="r${r.ot[0.5]>0?' ot-val':''}">${r.ot[0.5]>0?r.ot[0.5].toFixed(1):''}</td><td class="r${r.ot[0.75]>0?' ot-val':''}">${r.ot[0.75]>0?r.ot[0.75].toFixed(1):''}</td><td class="r${r.ot[1.0]>0?' ot-val':''}">${r.ot[1.0]>0?r.ot[1.0].toFixed(1):''}</td><td class="r">${fN(r.hourly)}</td><td class="r ot-val">${fN(r.otPay)}</td><td class="r ihss">${fN(r.ihssTotal)}</td><td class="r ihss">${fN(r.rap)}</td><td class="r">${fN(r.fuel)}</td><td class="r">${fN(r.advance)}</td><td class="r" style="font-weight:600">${fN(r.totalEarned)}</td><td class="r">${fN(r.totalDeductions)}</td><td class="r net">${fN(r.netPay)}</td></tr>`).join("")}
<tr class="total-row"><td colspan="3" style="text-align:right">TOTALES</td><td class="r">${fN(t.sal)}</td><td></td><td></td><td class="r">${fN(t.base)}</td><td colspan="4"></td><td></td><td class="r ot-val">${fN(t.ot)}</td><td class="r ihss">${fN(t.ihss)}</td><td class="r ihss">${fN(t.rap)}</td><td colspan="2"></td><td class="r" style="font-weight:700">${fN(t.earned)}</td><td class="r">${fN(t.ded)}</td><td class="r net" style="font-size:9pt">${fN(t.net)}</td></tr></tbody></table>
<div class="signatures"><div class="sig-box"><div class="sig-line">Elaborado por</div></div><div class="sig-box"><div class="sig-line">Revisado por</div></div><div class="sig-box"><div class="sig-line">Autorizado por</div></div></div></body></html>`;
  const w=window.open("","_blank");if(w){w.document.write(html);w.document.close()}
}

// ═══ TABS ═══
const TABS=[
  {id:"dash",label:"Dashboard",icon:"📊"},
  {id:"emp",label:"Empleados",icon:"👥"},
  {id:"clock",label:"Reloj",icon:"🕐"},
  {id:"pay",label:"Planilla",icon:"📋"},
  {id:"conf",label:"Confidencial",icon:"🔒"},
  {id:"hist",label:"Historial",icon:"📁"},
];

// ═══ APP ═══
export default function App(){
  const[tab,setTab]=useState("dash");
  const[employees,setEmployees]=useState([]);
  const[clockEntries,setClockEntries]=useState([]);
  const[payrolls,setPayrolls]=useState([]);
  const[holidays,setHolidays]=useState([]);
  const[loading,setLoading]=useState(true);
  const refresh=useCallback(async()=>{const[e,c,p,h]=await Promise.all([db.getEmployees(),db.getClockEntries(),db.getPayrolls(),db.getHolidays()]);setEmployees(e);setClockEntries(c);setPayrolls(p);setHolidays(h)},[]);
  useEffect(()=>{refresh().then(()=>setLoading(false))},[refresh]);

  if(loading)return(<div style={S.loading}><style>{CSS}</style><img src={LOGO} alt="H" style={{height:50,marginBottom:20,borderRadius:8}}/><div style={S.spinner}/><p style={{color:"rgba(255,255,255,0.6)",marginTop:14,fontSize:13}}>Cargando...</p></div>);

  const T={dash:DashboardTab,emp:EmployeesTab,clock:ClockTab,pay:PayrollTab,conf:ConfidentialTab,hist:HistoryTab}[tab];
  return(<div style={S.app}><style>{CSS}</style>
    <header style={S.header}><div style={S.brand}><img src={LOGO} alt="H" style={S.logoImg}/><div style={S.brandDivider}/><div style={S.brandSub}>Sistema de Planilla</div></div>
    <nav style={S.nav}>{TABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{...S.navBtn,...(tab===t.id?S.navAct:{})}}><span>{t.icon}</span><span>{t.label}</span></button>)}</nav></header>
    <main style={S.main}><T employees={employees} refresh={refresh} clockEntries={clockEntries} payrolls={payrolls} holidays={holidays}/></main>
    <footer style={S.footer}><span>Impresos Horeb © {new Date().getFullYear()}</span><span style={{color:"#94a3b8"}}>v4.0</span></footer>
  </div>);
}

// ═══ DASHBOARD ═══
function DashboardTab({employees,payrolls,holidays,refresh}){
  const totalSal=employees.reduce((s,e)=>s+e.salary,0);
  const last=payrolls.length>0?payrolls[0]:null;
  const[lastRows,setLastRows]=useState(null);
  const[newH,setNewH]=useState({date:"",name:""});
  useEffect(()=>{if(last)db.getPayrollRows(last.id).then(setLastRows)},[last]);
  const addH=async()=>{if(!newH.date||!newH.name)return;await db.addHoliday(newH);await refresh();setNewH({date:"",name:""})};
  const delH=async(id)=>{await db.deleteHoliday(id);await refresh()};
  return(<div>
    <h2 style={S.title}>Panel de Control</h2>
    <div style={S.grid4}>
      <div style={{...S.card,borderLeft:"4px solid #3b82f6",padding:16}}><div style={{fontSize:12,color:"#64748b"}}>Empleados</div><div style={{fontSize:20,fontWeight:700}}>{employees.length}</div></div>
      <div style={{...S.card,borderLeft:"4px solid #8b5cf6",padding:16}}><div style={{fontSize:12,color:"#64748b"}}>Nómina Mensual</div><div style={{fontSize:20,fontWeight:700}}>{formatL(totalSal)}</div></div>
      <div style={{...S.card,borderLeft:"4px solid #c9a227",padding:16}}><div style={{fontSize:12,color:"#64748b"}}>Planillas</div><div style={{fontSize:20,fontWeight:700}}>{payrolls.length}</div></div>
      <div style={{...S.card,borderLeft:"4px solid #059669",padding:16}}><div style={{fontSize:12,color:"#64748b"}}>Última</div><div style={{fontSize:18,fontWeight:700}}>{last?.period||"—"}</div></div>
    </div>
    {last&&lastRows&&<div style={S.card}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><h3 style={S.cardTitle}>Última: {last.period}</h3><button style={S.btnGold} onClick={()=>printPayroll({period:last.period,rows:lastRows})}>🖨️ Imprimir</button></div>
      <div style={{display:"flex",gap:24,flexWrap:"wrap",marginTop:8}}><div><div style={{fontSize:11,color:"#64748b"}}>Neto</div><div style={{fontSize:24,fontWeight:700,color:"#059669"}}>{formatL(+last.total_net)}</div></div></div></div>}
    <div style={S.card}><h3 style={S.cardTitle}>🎌 Días Feriados</h3>
      <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"flex-end"}}>
        <div><label style={S.label}>Fecha</label><input style={{...S.input,width:160}} type="date" value={newH.date} onChange={e=>setNewH({...newH,date:e.target.value})}/></div>
        <div><label style={S.label}>Nombre</label><input style={{...S.input,width:200}} value={newH.name} onChange={e=>setNewH({...newH,name:e.target.value})} placeholder="Día del Trabajo"/></div>
        <button style={S.btnPrimary} onClick={addH}>+ Agregar</button>
      </div>
      {holidays.length===0?<p style={{color:"#94a3b8"}}>No hay feriados.</p>:<div style={{display:"flex",flexWrap:"wrap",gap:8}}>{holidays.map(h=><div key={h.id} style={{display:"flex",alignItems:"center",gap:6,padding:"6px 12px",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,fontSize:12}}><span>🎌</span><strong style={{color:"#92400e"}}>{h.name}</strong><span style={{color:"#a16207"}}>{h.date}</span><button style={{background:"none",border:"none",cursor:"pointer",color:"#dc2626"}} onClick={()=>delH(h.id)}>✕</button></div>)}</div>}
    </div>
  </div>);
}

// ═══ EMPLOYEES ═══
function EmployeesTab({employees,refresh}){
  const weeklyEmps=employees.filter(e=>{const t=e.empType||"weekly";return t==="weekly"||t==="weekly_nonclock"});
  const[modal,setModal]=useState(null);const[form,setForm]=useState({id:"",name:"",position:"",salary:""});const[saving,setSaving]=useState(false);
  const openAdd=()=>{setForm({id:"",name:"",position:"",salary:""});setModal("new")};
  const openEdit=e=>{setForm({...e,salary:String(e.salary)});setModal(e.id)};
  const doSave=async()=>{const emp={...form,salary:parseFloat(form.salary)||0};if(!emp.id||!emp.name)return;setSaving(true);await db.upsertEmployee(emp);await refresh();setSaving(false);setModal(null)};
  const doDelete=async id=>{if(confirm("¿Eliminar?")){await db.deleteEmployee(id);await refresh()}};
  return(<div>
    <div style={S.titleRow}><h2 style={S.title}>Empleados (Semanal)</h2><button style={S.btnPrimary} onClick={openAdd}>+ Agregar</button></div>
    {modal&&<div style={S.overlay}><div style={S.modal}><h3 style={{fontSize:18,fontWeight:700,marginBottom:16}}>{modal==="new"?"Nuevo":"Editar"}</h3>
      <div style={S.formGrid}><Field l="Código" v={form.id} o={v=>setForm({...form,id:v})} ph="008" dis={modal!=="new"}/><Field l="Nombre" v={form.name} o={v=>setForm({...form,name:v})}/><Field l="Posición" v={form.position} o={v=>setForm({...form,position:v})}/><Field l="Salario" v={form.salary} o={v=>setForm({...form,salary:v})} t="number"/></div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:20}}><button style={S.btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"...":"Guardar"}</button></div>
    </div></div>}
    <div style={S.card}><div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Cód","Nombre","Posición","Sal.Mensual","Sal.Diario","Sal./Hora",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=3&&i<=5?"right":"left"}}>{h}</th>)}</tr></thead><tbody>
      {weeklyEmps.map(e=><tr key={e.id} style={e.empType==="weekly_nonclock"?{background:"#f0f3f8"}:{}}><td style={S.td}><span style={S.badge}>{e.id}</span></td><td style={{...S.td,fontWeight:600,color:"#0a2351"}}>{e.name} {e.empType==="weekly_nonclock"&&<span style={{fontSize:9,color:"#1d4ed8",background:"#eff6ff",padding:"1px 5px",borderRadius:3,marginLeft:4}}>SIN RELOJ</span>}</td><td style={S.td}><span style={S.posBadge}>{e.position}</span></td><td style={S.tdM}>{formatL(e.salary)}</td><td style={S.tdM}>{formatL(e.salary/30)}</td><td style={S.tdM}>{formatL(e.salary/30/8)}</td><td style={{...S.td,textAlign:"center"}}><button style={S.tblBtn} onClick={()=>openEdit(e)}>Editar</button><button style={{...S.tblBtn,color:"#dc2626"}} onClick={()=>doDelete(e.id)}>Eliminar</button></td></tr>)}
    </tbody></table></div></div>
  </div>);
}

// ═══ CLOCK ═══
function ClockTab({employees,clockEntries,refresh}){
  const[mode,setMode]=useState("import");const[importResult,setImportResult]=useState(null);const[importing,setImporting]=useState(false);
  const[mf,setMf]=useState({empId:"",date:new Date().toISOString().slice(0,10),cin:"08:00",lout:"12:00",lin:"13:00",cout:"18:00"});
  const[filterEmp,setFilterEmp]=useState("");const[filterMonth,setFilterMonth]=useState(new Date().toISOString().slice(0,7));
  const[editEntry,setEditEntry]=useState(null);const[editForm,setEditForm]=useState({checkIn:"",lunchOut:"",lunchIn:"",checkOut:""});const[editSaving,setEditSaving]=useState(false);

  const finishImport=async(parsed)=>{if(parsed.length===0){setImportResult({error:"No se encontraron registros válidos."});setImporting(false);return}
    await db.insertClockEntries(parsed);await refresh();setImportResult({added:parsed.length});setImporting(false)};

  const handleFile=e=>{const file=e.target.files[0];if(!file)return;setImporting(true);setImportResult(null);
    const isXLS=file.name.match(/\.xls[xm]?$/i)&&!file.name.match(/\.csv$/i);
    console.log("File:",file.name,"Size:",file.size,"isXLS:",isXLS);
    const tryParse=text=>{let p=parseClockCSV(text);if(p.length>0)return p;return parseClockCSV(text.replace(/;/g,","))};
    if(isXLS){const r=new FileReader();r.onload=async evt=>{
      const arr=new Uint8Array(evt.target.result);
      let p=parseClockXLS(arr);
      if(p.length===0){
        // Fallback: convert XLS to CSV using SheetJS and parse as text
        try{
          const wb=XLSX.read(arr,{type:"array",raw:false});
          const ws=wb.Sheets[wb.SheetNames[0]];
          const csvText=XLSX.utils.sheet_to_csv(ws,{FS:";"});
          console.log("XLS→CSV fallback, first 200 chars:",csvText.substring(0,200));
          p=tryParse(csvText);
          console.log("CSV fallback parsed:",p.length);
        }catch(e){console.error("XLS→CSV fallback error:",e)}
      }
      if(p.length===0){
        // Final fallback: read as text (for HTML-based .xls files)
        const r2=new FileReader();r2.onload=async e2=>{let t=e2.target.result;if(t.includes("<table")){const doc=new DOMParser().parseFromString(t,"text/html");const lines=[];doc.querySelectorAll("tr").forEach(r=>{lines.push(Array.from(r.querySelectorAll("td,th")).map(c=>c.textContent.trim()).join(","))});t=lines.join("\n")}await finishImport(tryParse(t))};r2.readAsText(file,"ISO-8859-1");return;
      }
      await finishImport(p)};r.readAsArrayBuffer(file)}
    else{const r=new FileReader();r.onload=async evt=>{let p=tryParse(evt.target.result);if(p.length>0){await finishImport(p);return}const r2=new FileReader();r2.onload=async e2=>{await finishImport(tryParse(e2.target.result))};r2.readAsText(file,"ISO-8859-1")};r.readAsText(file,"UTF-8")}
    e.target.value=""};

  const addManual=async()=>{if(!mf.empId)return;const entry={id:`${mf.empId}_${mf.date}_m${Date.now()}`,employeeId:mf.empId,date:mf.date,checkIn:`${mf.date}T${mf.cin}:00`,lunchOut:`${mf.date}T${mf.lout}:00`,lunchIn:`${mf.date}T${mf.lin}:00`,checkOut:`${mf.date}T${mf.cout}:00`,punches:4};await db.insertClockEntries([entry]);await refresh()};
  const removeEntry=async id=>{await db.deleteClockEntry(id);await refresh()};
  const clearAll=async()=>{if(confirm("¿Borrar TODO?")){await db.deleteAllClockEntries();await refresh()}};

  const extractTime=iso=>{if(!iso)return"";try{const d=new Date(iso);return`${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`}catch{return""}};
  const openEdit=entry=>{setEditEntry(entry);setEditForm({checkIn:extractTime(entry.checkIn),lunchOut:extractTime(entry.lunchOut),lunchIn:extractTime(entry.lunchIn),checkOut:extractTime(entry.checkOut)})};
  const saveEdit=async()=>{if(!editEntry)return;setEditSaving(true);const d=editEntry.date;const toISO=t=>t?new Date(`${d}T${t}:00`).toISOString():null;
    await db.updateClockEntry({id:editEntry.id,employeeId:editEntry.employeeId,date:d,checkIn:toISO(editForm.checkIn),lunchOut:toISO(editForm.lunchOut),lunchIn:toISO(editForm.lunchIn),checkOut:toISO(editForm.checkOut)});
    await refresh();setEditSaving(false);setEditEntry(null)};

  const hasError=entry=>{const dow=new Date(entry.date+"T12:00:00").getDay();if(entry.checkIn&&!entry.checkOut&&dow===5)return"VIERNES SIN SALIDA";if(!entry.checkIn&&!entry.checkOut)return"Sin marcaciones";if(entry.checkIn&&!entry.checkOut)return"Falta salida";if(!entry.checkIn)return"Falta entrada";const hrs=calcDayHours(entry);if(hrs>16)return"Más de 16hrs";if(!entry.lunchOut&&!entry.lunchIn&&hrs>5)return"Sin almuerzo";return null};

  const filtered=clockEntries.filter(e=>(!filterEmp||e.employeeId===filterEmp)&&(!filterMonth||e.date.startsWith(filterMonth))).sort((a,b)=>b.date.localeCompare(a.date)||(a.employeeId||"").localeCompare(b.employeeId||""));
  const getN=id=>employees.find(e=>e.id===id)?.name||`ID ${id}`;
  const errCount=filtered.filter(e=>hasError(e)).length;

  return(<div>
    <h2 style={S.title}>Reloj Marcador</h2>
    {editEntry&&<div style={S.overlay}><div style={S.modal}><h3 style={{fontSize:18,fontWeight:700,marginBottom:16}}>Editar Marcación</h3>
      <div style={{background:"#f0f3f8",borderRadius:10,padding:12,marginBottom:16}}><div style={{fontWeight:700,color:"#0a2351"}}>{getN(editEntry.employeeId)}</div><div style={{fontSize:13,color:"#475569"}}>{fmtDate(editEntry.date+"T12:00:00")}</div>
        {hasError(editEntry)&&<div style={{marginTop:6,padding:"4px 8px",background:"#fef2f2",border:"1px solid #fecaca",borderRadius:6,fontSize:12,color:"#dc2626",fontWeight:600}}>⚠️ {hasError(editEntry)}</div>}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}><div><label style={S.label}>Entrada</label><input style={S.input} type="time" value={editForm.checkIn} onChange={e=>setEditForm({...editForm,checkIn:e.target.value})}/></div><div><label style={S.label}>Sale Almuerzo</label><input style={S.input} type="time" value={editForm.lunchOut} onChange={e=>setEditForm({...editForm,lunchOut:e.target.value})}/></div><div><label style={S.label}>Regresa</label><input style={S.input} type="time" value={editForm.lunchIn} onChange={e=>setEditForm({...editForm,lunchIn:e.target.value})}/></div><div><label style={S.label}>Salida</label><input style={S.input} type="time" value={editForm.checkOut} onChange={e=>setEditForm({...editForm,checkOut:e.target.value})}/></div></div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:20}}><button style={S.btnSec} onClick={()=>setEditEntry(null)}>Cancelar</button><button style={S.btnPrimary} onClick={saveEdit} disabled={editSaving}>{editSaving?"...":"Guardar"}</button></div>
    </div></div>}
    <div style={S.card}>
      <div style={{display:"flex",gap:8,marginBottom:16}}><button style={mode==="import"?S.btnPrimary:S.btnSec} onClick={()=>setMode("import")}>📥 Importar</button><button style={mode==="manual"?S.btnPrimary:S.btnSec} onClick={()=>setMode("manual")}>✍️ Manual</button></div>
      {mode==="import"?<div><p style={{fontSize:13,color:"#475569",marginBottom:12}}>Sube el archivo <strong>.xls</strong> o <strong>.csv</strong> del reloj.</p>
        <label style={S.fileLabel}>{importing?"⏳ Procesando...":"📄 Seleccionar archivo"}<input type="file" accept=".csv,.txt,.xls,.xlsx" onChange={handleFile} style={{display:"none"}} disabled={importing}/></label>
        {importResult&&<div style={{marginTop:12,padding:12,borderRadius:8,background:importResult.error?"#fef2f2":"#f0fdf4",border:`1px solid ${importResult.error?"#fecaca":"#bbf7d0"}`}}>{importResult.error?<p style={{color:"#dc2626",fontSize:13}}>⚠️ {importResult.error}</p>:<p style={{color:"#059669",fontSize:13}}>✅ {importResult.added} registros importados.</p>}</div>}
      </div>:<div style={S.formGrid}><Field l="Empleado" v={mf.empId} o={v=>setMf({...mf,empId:v})} t="select" opts={[{v:"",l:"—"}, ...employees.map(e=>({v:e.id,l:`${e.id}-${e.name}`}))]}/><Field l="Fecha" v={mf.date} o={v=>setMf({...mf,date:v})} t="date"/><Field l="Entrada" v={mf.cin} o={v=>setMf({...mf,cin:v})} t="time"/><Field l="Sal.Alm" v={mf.lout} o={v=>setMf({...mf,lout:v})} t="time"/><Field l="Reg.Alm" v={mf.lin} o={v=>setMf({...mf,lin:v})} t="time"/><Field l="Salida" v={mf.cout} o={v=>setMf({...mf,cout:v})} t="time"/><div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={addManual}>Registrar</button></div></div>}
    </div>
    <div style={S.card}><div style={S.titleRow}><h3 style={S.cardTitle}>Registros ({filtered.length}) {errCount>0&&<span style={{marginLeft:8,padding:"2px 8px",background:"#fef2f2",color:"#dc2626",borderRadius:20,fontSize:12,fontWeight:700}}>⚠️ {errCount}</span>}</h3>{clockEntries.length>0&&<button style={{...S.btnSec,color:"#dc2626",fontSize:12}} onClick={clearAll}>Borrar Todo</button>}</div>
      <div style={{display:"flex",gap:12,marginBottom:12,flexWrap:"wrap"}}><select style={{...S.input,width:200}} value={filterEmp} onChange={e=>setFilterEmp(e.target.value)}><option value="">Todos</option>{employees.map(e=><option key={e.id} value={e.id}>{e.id}-{e.name}</option>)}</select><input style={{...S.input,width:160}} type="month" value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}/></div>
      {filtered.length===0?<p style={{color:"#94a3b8"}}>No hay registros.</p>:<div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Empleado","Fecha","Entrada","Sal.Alm","Reg.Alm","Salida","Hrs","Extra","Estado",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=6?"right":"left",...(i>=8?{textAlign:"center"}:{})}}>{h}</th>)}</tr></thead><tbody>
        {filtered.map(e=>{const hrs=calcDayHours(e),dow=new Date(e.date+"T12:00:00").getDay(),extra=Math.max(0,hrs-getScheduledHours(dow)),isWE=dow===0||dow===6,error=hasError(e);
        return<tr key={e.id} style={{background:error?"#fef2f2":isWE?"#fffbeb":"transparent"}}>
          <td style={{...S.td,fontWeight:600,color:"#0a2351",whiteSpace:"nowrap"}}>{getN(e.employeeId)}</td>
          <td style={S.td}>{fmtDate(e.date+"T12:00:00")}{isWE&&<span style={{marginLeft:4,fontSize:10,color:"#c9a227",fontWeight:700}}>{dow===0?"DOM":"SÁB"}</span>}</td>
          <td style={{...S.td,color:e.checkIn?"#059669":"#dc2626",fontWeight:600}}>{fmtTime(e.checkIn)}</td>
          <td style={{...S.td,color:"#92400e"}}>{fmtTime(e.lunchOut)}</td><td style={{...S.td,color:"#92400e"}}>{fmtTime(e.lunchIn)}</td>
          <td style={{...S.td,color:e.checkOut?"#b91c1c":"#dc2626",fontWeight:600}}>{fmtTime(e.checkOut)}</td>
          <td style={{...S.td,textAlign:"right",fontFamily:"monospace",fontSize:12}}>{hrs.toFixed(1)}h</td>
          <td style={{...S.td,textAlign:"right",fontFamily:"monospace",fontSize:12,color:extra>0?"#b91c1c":"#cbd5e1",fontWeight:extra>0?700:400}}>{extra>0?`+${extra.toFixed(1)}h`:"—"}</td>
          <td style={{...S.td,textAlign:"center"}}>{error?<span style={{fontSize:11,color:"#dc2626",fontWeight:600}}>⚠️ {error}</span>:<span style={{fontSize:11,color:"#059669"}}>✓</span>}</td>
          <td style={{...S.td,textAlign:"center",whiteSpace:"nowrap"}}><button style={{...S.tblBtn,color:"#0a2351"}} onClick={()=>openEdit(e)}>✏️</button><button style={{...S.tblBtn,color:"#dc2626"}} onClick={()=>removeEntry(e.id)}>✕</button></td>
        </tr>})}
      </tbody></table></div>}
    </div>
  </div>);
}

// ═══ PAYROLL ═══
function PayrollTab({employees,clockEntries,refresh,holidays}){
  const[weekNum,setWeekNum]=useState("");const[from,setFrom]=useState("");const[to,setTo]=useState("");const[result,setResult]=useState(null);const[adj,setAdj]=useState({});const[saving,setSaving]=useState(false);
  const weeklyEmps=employees.filter(e=>{const t=e.empType||"weekly";return t==="weekly"||t==="weekly_nonclock"});

  const generate=()=>{
    if(!from||!to)return alert("Selecciona fechas.");
    const fs=from.slice(0,10),ts=to.slice(0,10);
    const pe=clockEntries.filter(c=>{const d=String(c.date).slice(0,10);return d>=fs&&d<=ts});
    const holidayDates=new Set(holidays.filter(h=>h.date>=fs&&h.date<=ts).map(h=>h.date));
    const holidayNames={};holidays.forEach(h=>{if(h.date>=fs&&h.date<=ts)holidayNames[h.date]=h.name});
    const byEmp={};pe.forEach(e=>{if(!byEmp[e.employeeId])byEmp[e.employeeId]=[];byEmp[e.employeeId].push(e)});

    // Working days (Mon-Fri excl first Friday, holidays INCLUDED as they count as worked)
    let workDays=0;
    const s0=new Date(fs+"T12:00:00"),e0=new Date(ts+"T12:00:00");
    for(let d=new Date(s0);d<=e0;d.setDate(d.getDate()+1)){const dow=d.getDay();const ds=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;if(dow>=1&&dow<=5&&!(ds===fs&&dow===5))workDays++}

    // Count holidays on weekdays (excl first Friday)
    let holOnWD=0;
    for(let d=new Date(s0);d<=e0;d.setDate(d.getDate()+1)){const dow=d.getDay();const ds=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;if(dow>=1&&dow<=5&&!(ds===fs&&dow===5)&&holidayDates.has(ds))holOnWD++}

    const applyIHSS=isLastWeekOfMonth(fs,ts);const applyRAP=isSecondWeekOfMonth(fs);
    const autoFills=[];

    const rows=weeklyEmps.map(emp=>{
      const isNC=emp.empType==="weekly_nonclock";
      if(isNC){
        const daily=emp.salary/30,hourly=daily/8;const a=adj[emp.id]||{};const faltas=+a.faltas||0;
        const daysPaid=Math.max(0,7-(faltas*2)),baseSalary=daily*daysPaid;
        const ihss=applyIHSS?calcIHSS_monthly(emp.salary):{em:0,ivm:0,total:0};
        const rapD=applyRAP?calcRAP_monthly(emp.salary):{employeeTotal:0};
        const fuel=+a.fuel||0,vacation=+a.vacation||0,incapacity=+a.incapacity||0,advance=+a.advance||0,dec4=+a.dec4||0,dec3=+a.dec3||0,otherDed=+a.otherDed||0;
        const totalEarned=baseSalary+fuel+vacation+incapacity+dec4+dec3;
        const totalDeductions=ihss.total+rapD.employeeTotal+advance+otherDed;
        return{employeeId:emp.id,name:emp.name,position:emp.position,salary:emp.salary,daily,hourly,isNonClock:true,daysWorked:workDays-faltas,absences:faltas,daysPaid,days:daysPaid,effectiveHrs:0,baseSalary,ot:{0.25:0,0.5:0,0.75:0,1.0:0},otPay:0,ihssTotal:ihss.total,rap:rapD.employeeTotal,fuel,vacation,incapacity,advance,dec4,dec3,otherDed,totalEarned,totalDeductions,netPay:totalEarned-totalDeductions};
      }
      // Clock employees
      let empEntries=(byEmp[emp.id]||[]).map(e=>{
        if(e.checkIn&&!e.checkOut){const dow=new Date(e.date+"T12:00:00").getDay();if(dow>=1&&dow<=5){const t=dow===5?"17:00:00":"18:00:00";autoFills.push({name:emp.name,date:e.date,time:dow===5?"5pm":"6pm"});return{...e,checkOut:`${e.date}T${t}`}}}return e});

      const clockedDays=empEntries.filter(e=>{if(!e.checkIn||!e.checkOut)return false;const dow=new Date(e.date+"T12:00:00").getDay();return dow>=1&&dow<=5&&!holidayDates.has(e.date)&&!(e.date===fs&&dow===5)}).length;
      const daysWorked=clockedDays+holOnWD;
      const absences=Math.max(0,workDays-daysWorked);const daysPaid=Math.max(0,7-(absences*2));
      const daily=emp.salary/30,hourly=daily/8,baseSalary=daily*daysPaid;

      // OT: extra = total effective - scheduled hours. Exit time determines RATE.
      // Skip OT for auto-filled entries (unknown real exit time)
      let ot={0.25:0,0.5:0,0.75:0,1.0:0},totalEff=0;
      empEntries.forEach(en=>{if(!en.checkIn||!en.checkOut)return;const dow=new Date(en.checkIn).getDay();const hrs=calcDayHours(en);totalEff+=hrs;
        if(en.autoFilled)return; // Don't calculate OT for auto-filled exits
        if(dow===0||holidayDates.has(en.date)){ot[1.0]+=hrs;return}
        if(dow===6){ot[0.25]+=hrs;return}
        const scheduled=getScheduledHours(dow);
        const extra=hrs-scheduled;
        if(extra<=0)return;
        // Exit time determines rate bracket
        const cH=new Date(en.checkOut).getHours()+new Date(en.checkOut).getMinutes()/60;
        if(cH<=19){ot[0.25]+=extra}
        else if(cH<=21){ot[0.25]+=Math.min(extra,1);if(extra>1)ot[0.5]+=extra-1}
        else{ot[0.25]+=Math.min(extra,1);const r=Math.max(0,extra-1);ot[0.5]+=Math.min(r,2);if(r>2)ot[0.75]+=r-2}});
      const otPay=Object.entries(ot).reduce((s,[r,h])=>s+h*hourly*(1+parseFloat(r)),0);

      const ihss=applyIHSS?calcIHSS_monthly(emp.salary):{total:0};
      const rapD=applyRAP?calcRAP_monthly(emp.salary):{employeeTotal:0};
      const a=adj[emp.id]||{};const fuel=+a.fuel||0,vacation=+a.vacation||0,incapacity=+a.incapacity||0,advance=+a.advance||0,dec4=+a.dec4||0,dec3=+a.dec3||0,otherDed=+a.otherDed||0;
      const totalEarned=baseSalary+otPay+fuel+vacation+incapacity+dec4+dec3;
      const totalDeductions=ihss.total+rapD.employeeTotal+advance+otherDed;
      return{employeeId:emp.id,name:emp.name,position:emp.position,salary:emp.salary,daily,hourly,daysWorked,absences,daysPaid,days:daysPaid,effectiveHrs:totalEff,baseSalary,ot,otPay,ihssTotal:ihss.total,rap:rapD.employeeTotal,fuel,vacation,incapacity,advance,dec4,dec3,otherDed,totalEarned,totalDeductions,netPay:totalEarned-totalDeductions};
    });
    setResult({period:`WK${weekNum||"?"} ${from} al ${to}`,rows,from,to,weekNum,workDays,holOnWD,applyIHSS,applyRAP,autoFills,holidayNames:Object.entries(holidayNames).map(([d,n])=>`${n} (${d})`)});
  };

  const doSave=async()=>{if(!result)return;setSaving(true);await db.savePayroll(result);await refresh();setSaving(false);alert("✅ Guardada.")};
  const updAdj=(eId,f,v)=>setAdj(p=>({...p,[eId]:{...(p[eId]||{}),[f]:v}}));

  return(<div>
    <h2 style={S.title}>Generar Planilla</h2>
    <div style={S.card}><h3 style={S.cardTitle}>📅 Período</h3><p style={{fontSize:12,color:"#64748b",marginBottom:10}}>Viernes 5pm → Viernes siguiente 5pm</p>
      <div style={S.formGrid}><Field l="Semana #" v={weekNum} o={setWeekNum} ph="3"/><Field l="Desde" v={from} o={setFrom} t="date"/><Field l="Hasta" v={to} o={setTo} t="date"/><div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={generate}>Generar</button></div></div></div>
    {result&&<>
      <div style={{...S.card,background:"#f0f3f8",border:"1px solid #d0daea"}}><div style={{display:"flex",gap:20,flexWrap:"wrap",fontSize:13}}>
        <div>Lab: <strong>{result.workDays}</strong></div>{result.holOnWD>0&&<div>Feriados: <strong style={{color:"#c9a227"}}>{result.holOnWD}</strong></div>}<div>Regla: <strong style={{color:"#059669"}}>{result.workDays}d=7d</strong></div>
        <div>IHSS: <strong style={{color:result.applyIHSS?"#7c3aed":"#94a3b8"}}>{result.applyIHSS?"✓ Aplica":"No"}</strong></div><div>RAP: <strong style={{color:result.applyRAP?"#0369a1":"#94a3b8"}}>{result.applyRAP?"✓ Aplica":"No"}</strong></div>
      </div>{result.holidayNames?.length>0&&<div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>{result.holidayNames.map((h,i)=><span key={i} style={{padding:"2px 8px",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:6,fontSize:12,color:"#92400e"}}>🎌 {h}</span>)}</div>}</div>
      {result.autoFills?.length>0&&<div style={{...S.card,background:"#fffbeb",border:"1px solid #fde68a"}}><h4 style={{fontSize:13,color:"#92400e",marginBottom:6}}>⚠️ Salida auto-completada</h4><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{result.autoFills.map((f,i)=><span key={i} style={{padding:"3px 10px",background:"#fff",border:"1px solid #fde68a",borderRadius:6,fontSize:12,color:"#92400e"}}>{f.name} — {f.date} → {f.time}</span>)}</div></div>}
      <div style={S.card}><div style={S.titleRow}><h3 style={S.cardTitle}>{result.period}</h3><div style={{display:"flex",gap:8}}><button style={S.btnGold} onClick={()=>printPayroll(result)}>🖨️</button><button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"...":"💾 Guardar"}</button></div></div>
        <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Cód","Nombre","Pos.","Sal.M.","Trab.","Faltas","Pago","Salario","Tot.OT","IHSS","RAP","Dev.","Ded.","Neto"].map((h,i)=><th key={i} style={{...S.th,fontSize:10,textAlign:i>=3?"right":"left"}}>{h}</th>)}</tr></thead><tbody>
          {result.rows.map(r=><tr key={r.employeeId} style={r.absences>0?{background:"#fffbeb"}:r.isNonClock?{background:"#f0f3f8"}:{}}>
            <td style={S.td}><span style={S.badge}>{r.employeeId}</span></td><td style={{...S.td,fontWeight:600,whiteSpace:"nowrap",fontSize:12,color:"#0a2351"}}>{r.name}{r.isNonClock&&<span style={{fontSize:9,color:"#1d4ed8",background:"#eff6ff",padding:"1px 4px",borderRadius:3,marginLeft:3}}>SR</span>}</td>
            <td style={{...S.td,fontSize:11}}>{r.position}</td><td style={S.tdM}>{formatL(r.salary)}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#059669"}}>{r.isNonClock?"—":`${r.daysWorked}d`}</td>
            <td style={{...S.tdM,fontWeight:700,color:r.absences>0?"#dc2626":"#059669"}}>{r.absences>0?`${r.absences}d`:"—"}</td>
            <td style={{...S.tdM,fontWeight:700,color:r.daysPaid<7?"#c9a227":"#0a2351"}}>{r.daysPaid}d</td>
            <td style={{...S.tdM,fontWeight:600}}>{formatL(r.baseSalary)}</td>
            <td style={{...S.tdM,color:r.otPay>0?"#b91c1c":"#d4d4d8"}}>{r.isNonClock?"N/A":formatL(r.otPay)}</td>
            <td style={{...S.tdM,color:"#7c3aed"}}>{r.ihssTotal>0?formatL(r.ihssTotal):"—"}</td>
            <td style={{...S.tdM,color:"#0369a1"}}>{r.rap>0?formatL(r.rap):"—"}</td>
            <td style={{...S.tdM,fontWeight:600}}>{formatL(r.totalEarned)}</td><td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#059669",fontSize:13}}>{formatL(r.netPay)}</td></tr>)}
        </tbody><tfoot><tr style={{background:"#e8eef6"}}><td colSpan={7} style={{...S.td,fontWeight:700}}>TOTALES</td><td style={{...S.tdM,fontWeight:700}}>{formatL(result.rows.reduce((s,r)=>s+r.baseSalary,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.otPay,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#7c3aed"}}>{formatL(result.rows.reduce((s,r)=>s+r.ihssTotal,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#0369a1"}}>{formatL(result.rows.reduce((s,r)=>s+r.rap,0))}</td><td style={{...S.tdM,fontWeight:700}}>{formatL(result.rows.reduce((s,r)=>s+r.totalEarned,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.totalDeductions,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#059669",fontSize:14}}>{formatL(result.rows.reduce((s,r)=>s+r.netPay,0))}</td></tr></tfoot></table></div></div>
      <div style={S.card}><h3 style={S.cardTitle}>⚙️ Ajustes</h3><div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Empleado","Tipo","Faltas","Comb.","Vac.","Incap.","Adel.","Dec4","Dec3","Otras"].map((h,i)=><th key={i} style={S.th}>{h}</th>)}</tr></thead><tbody>
        {weeklyEmps.map(emp=>{const a=adj[emp.id]||{};const isNC=emp.empType==="weekly_nonclock";return<tr key={emp.id} style={isNC?{background:"#f0f3f8"}:{}}><td style={{...S.td,fontWeight:600,fontSize:12,color:"#0a2351",whiteSpace:"nowrap"}}>{emp.name}</td><td style={{...S.td,fontSize:11}}>{isNC?<span style={{color:"#1d4ed8",fontSize:10}}>SR</span>:"Reloj"}</td><td style={S.td}>{isNC?<input style={{...S.input,width:50,padding:"3px",fontSize:12,textAlign:"right"}} type="number" value={a.faltas||""} onChange={e=>updAdj(emp.id,"faltas",e.target.value)} placeholder="0"/>:<span style={{color:"#94a3b8",fontSize:11}}>auto</span>}</td>
        {["fuel","vacation","incapacity","advance","dec4","dec3","otherDed"].map(f=><td key={f} style={S.td}><input style={{...S.input,width:70,padding:"3px",fontSize:12,textAlign:"right"}} type="number" value={a[f]||""} onChange={e=>updAdj(emp.id,f,e.target.value)} placeholder="0"/></td>)}</tr>})}
      </tbody></table></div></div>
    </>}
  </div>);
}

// ═══ CONFIDENTIAL ═══
function ConfidentialTab({employees,refresh}){
  const conf=employees.filter(e=>e.empType==="biweekly");
  const[modal,setModal]=useState(null);const[form,setForm]=useState({id:"",name:"",position:"",salary:""});const[saving,setSaving]=useState(false);
  const[from,setFrom]=useState("");const[to,setTo]=useState("");const[result,setResult]=useState(null);const[adj,setAdj]=useState({});const[paySaving,setPaySaving]=useState(false);
  const openAdd=()=>{setForm({id:"",name:"",position:"",salary:""});setModal("new")};
  const openEdit=e=>{setForm({...e,salary:String(e.salary)});setModal(e.id)};
  const doSave=async()=>{const emp={...form,salary:parseFloat(form.salary)||0,empType:"biweekly"};if(!emp.id||!emp.name)return;setSaving(true);await db.upsertEmployee(emp);await refresh();setSaving(false);setModal(null)};
  const doDelete=async id=>{if(confirm("¿Eliminar?")){await db.deleteEmployee(id);await refresh()}};
  const generate=()=>{if(!from||!to)return;const rows=conf.map(emp=>{const bw=emp.salary/2;const ihss=calcIHSS_biweekly(emp.salary);const a=adj[emp.id]||{};const adv=+a.advance||0,other=+a.otherDed||0;const te=bw;const td=ihss.total+adv+other;return{employeeId:emp.id,name:emp.name,position:emp.position,salary:emp.salary,baseSalary:bw,ihssEM:ihss.em,ihssIVM:ihss.ivm,ihssTotal:ihss.total,advance:adv,otherDed:other,totalEarned:te,totalDeductions:td,netPay:te-td,ot:{0.25:0,0.5:0,0.75:0,1.0:0},otPay:0,daily:emp.salary/30,hourly:emp.salary/30/8,days:15,effectiveHrs:0,fuel:0,vacation:0,incapacity:0,dec4:0,dec3:0,rap:0}});setResult({period:`Q ${from} al ${to}`,rows,from,to})};
  const doSaveP=async()=>{if(!result)return;setPaySaving(true);/* Save biweekly payroll to regular payrolls table */await db.savePayroll(result);await refresh();setPaySaving(false);alert("✅ Guardada.")};
  const updAdj=(id,f,v)=>setAdj(p=>({...p,[id]:{...(p[id]||{}),[f]:v}}));
  return(<div>
    <h2 style={S.title}>Planilla Confidencial (Quincenal)</h2>
    <div style={S.card}><div style={S.titleRow}><h3 style={S.cardTitle}>🔒 Empleados Confidenciales</h3><button style={S.btnPrimary} onClick={openAdd}>+ Agregar</button></div>
      {modal&&<div style={S.overlay}><div style={S.modal}><h3 style={{fontSize:18,fontWeight:700,marginBottom:16}}>{modal==="new"?"Nuevo":"Editar"}</h3>
        <div style={S.formGrid}><Field l="Código" v={form.id} o={v=>setForm({...form,id:v})} dis={modal!=="new"}/><Field l="Nombre" v={form.name} o={v=>setForm({...form,name:v})}/><Field l="Posición" v={form.position} o={v=>setForm({...form,position:v})}/><Field l="Salario Mensual" v={form.salary} o={v=>setForm({...form,salary:v})} t="number"/></div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:20}}><button style={S.btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"...":"Guardar"}</button></div></div></div>}
      {conf.length===0?<p style={{color:"#94a3b8"}}>No hay empleados confidenciales.</p>:<div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Cód","Nombre","Posición","Sal.Mensual","Quincenal","IHSS/Q",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=3?"right":"left"}}>{h}</th>)}</tr></thead><tbody>
        {conf.map(e=>{const ihss=calcIHSS_biweekly(e.salary);return<tr key={e.id}><td style={S.td}><span style={S.badge}>{e.id}</span></td><td style={{...S.td,fontWeight:600}}>{e.name}</td><td style={S.td}>{e.position}</td><td style={S.tdM}>{formatL(e.salary)}</td><td style={S.tdM}>{formatL(e.salary/2)}</td><td style={{...S.tdM,color:"#7c3aed"}}>{formatL(ihss.total)}</td><td style={{...S.td,textAlign:"center"}}><button style={S.tblBtn} onClick={()=>openEdit(e)}>Editar</button><button style={{...S.tblBtn,color:"#dc2626"}} onClick={()=>doDelete(e.id)}>Eliminar</button></td></tr>})}
      </tbody></table></div>}</div>
    <div style={S.card}><h3 style={S.cardTitle}>📋 Generar Quincenal</h3><div style={S.formGrid}><Field l="Desde" v={from} o={setFrom} t="date"/><Field l="Hasta" v={to} o={setTo} t="date"/><div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={generate}>Generar</button></div></div></div>
    {result&&<div style={S.card}><div style={S.titleRow}><h3 style={S.cardTitle}>{result.period}</h3><div style={{display:"flex",gap:8}}><button style={S.btnGold} onClick={()=>printPayroll(result)}>🖨️</button><button style={S.btnPrimary} onClick={doSaveP} disabled={paySaving}>{paySaving?"...":"💾 Guardar"}</button></div></div>
      <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Cód","Nombre","Sal.M.","Quincenal","IHSS EM","IHSS IVM","Tot.IHSS","Adel.","Otras","Tot.Ded.","Neto"].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=2?"right":"left"}}>{h}</th>)}</tr></thead><tbody>
        {result.rows.map(r=><tr key={r.employeeId}><td style={S.td}><span style={S.badge}>{r.employeeId}</span></td><td style={{...S.td,fontWeight:600}}>{r.name}</td><td style={S.tdM}>{formatL(r.salary)}</td><td style={{...S.tdM,fontWeight:600}}>{formatL(r.baseSalary)}</td><td style={{...S.tdM,color:"#7c3aed"}}>{formatL(r.ihssEM)}</td><td style={{...S.tdM,color:"#7c3aed"}}>{formatL(r.ihssIVM)}</td><td style={{...S.tdM,fontWeight:600,color:"#7c3aed"}}>{formatL(r.ihssTotal)}</td><td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.advance)}</td><td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.otherDed)}</td><td style={{...S.tdM,fontWeight:600,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td><td style={{...S.tdM,fontWeight:700,color:"#059669",fontSize:13}}>{formatL(r.netPay)}</td></tr>)}
      </tbody><tfoot><tr style={{background:"#e8eef6"}}><td colSpan={3} style={{...S.td,fontWeight:700}}>TOTALES</td><td style={{...S.tdM,fontWeight:700}}>{formatL(result.rows.reduce((s,r)=>s+r.baseSalary,0))}</td><td colSpan={2}></td><td style={{...S.tdM,fontWeight:700,color:"#7c3aed"}}>{formatL(result.rows.reduce((s,r)=>s+r.ihssTotal,0))}</td><td colSpan={2}></td><td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.totalDeductions,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#059669",fontSize:14}}>{formatL(result.rows.reduce((s,r)=>s+r.netPay,0))}</td></tr></tfoot></table></div>
      <div style={{marginTop:12}}><h4 style={{fontSize:13,fontWeight:700,color:"#475569",marginBottom:6}}>Ajustes</h4><div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Empleado","Adelanto","Otras Ded."].map((h,i)=><th key={i} style={S.th}>{h}</th>)}</tr></thead><tbody>{conf.map(emp=>{const a=adj[emp.id]||{};return<tr key={emp.id}><td style={{...S.td,fontWeight:600,fontSize:12}}>{emp.name}</td>{["advance","otherDed"].map(f=><td key={f} style={S.td}><input style={{...S.input,width:90,padding:"4px",fontSize:12,textAlign:"right"}} type="number" value={a[f]||""} onChange={e=>updAdj(emp.id,f,e.target.value)} placeholder="0"/></td>)}</tr>})}</tbody></table></div></div>
    </div>}
  </div>);
}

// ═══ HISTORY ═══
function HistoryTab({payrolls,refresh}){
  const[sel,setSel]=useState(null);const[rows,setRows]=useState(null);const[loading,setLoading]=useState(false);
  const select=async i=>{setSel(i);setLoading(true);const r=await db.getPayrollRows(payrolls[i].id);setRows(r);setLoading(false)};
  const del=async i=>{if(!confirm("¿Eliminar?"))return;await db.deletePayroll(payrolls[i].id);await refresh();setSel(null);setRows(null)};
  return(<div><h2 style={S.title}>Historial</h2>
    {payrolls.length===0?<div style={S.card}><p style={{color:"#94a3b8"}}>No hay planillas.</p></div>:<>
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>{payrolls.map((p,i)=><button key={i} onClick={()=>select(i)} style={sel===i?S.btnPrimary:S.btnSec}>{p.period}</button>)}</div>
      {loading&&<div style={S.card}><p>Cargando...</p></div>}
      {sel!==null&&rows&&!loading&&<div style={S.card}><div style={S.titleRow}><h3 style={S.cardTitle}>{payrolls[sel].period}</h3><div style={{display:"flex",gap:8}}><button style={S.btnGold} onClick={()=>printPayroll({period:payrolls[sel].period,rows})}>🖨️</button><button style={{...S.btnSec,color:"#dc2626"}} onClick={()=>del(sel)}>🗑️</button></div></div>
        <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>{["Cód","Nombre","Días","Salario","H.Extra","IHSS","RAP","Dev.","Ded.","Neto"].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=2?"right":"left"}}>{h}</th>)}</tr></thead><tbody>
          {rows.map(r=><tr key={r.employeeId}><td style={S.td}><span style={S.badge}>{r.employeeId}</span></td><td style={{...S.td,fontWeight:600}}>{r.name}</td><td style={S.tdM}>{r.days}</td><td style={S.tdM}>{formatL(r.baseSalary)}</td><td style={{...S.tdM,color:r.otPay>0?"#b91c1c":"#d4d4d8"}}>{formatL(r.otPay)}</td><td style={{...S.tdM,color:"#7c3aed"}}>{formatL(r.ihssTotal)}</td><td style={{...S.tdM,color:"#0369a1"}}>{formatL(r.rap)}</td><td style={S.tdM}>{formatL(r.totalEarned)}</td><td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td><td style={{...S.tdM,fontWeight:700,color:"#059669"}}>{formatL(r.netPay)}</td></tr>)}
        </tbody><tfoot><tr style={{background:"#e8eef6"}}><td colSpan={3} style={{...S.td,fontWeight:700}}>TOTALES</td><td style={{...S.tdM,fontWeight:700}}>{formatL(rows.reduce((s,r)=>s+r.baseSalary,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(rows.reduce((s,r)=>s+r.otPay,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#7c3aed"}}>{formatL(rows.reduce((s,r)=>s+r.ihssTotal,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#0369a1"}}>{formatL(rows.reduce((s,r)=>s+r.rap,0))}</td><td style={{...S.tdM,fontWeight:700}}>{formatL(rows.reduce((s,r)=>s+r.totalEarned,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(rows.reduce((s,r)=>s+r.totalDeductions,0))}</td><td style={{...S.tdM,fontWeight:700,color:"#059669",fontSize:14}}>{formatL(rows.reduce((s,r)=>s+r.netPay,0))}</td></tr></tfoot></table></div></div>}
    </>}
  </div>);
}

// ═══ SHARED ═══
function Field({l,v,o,ph,t="text",opts,dis}){return<div style={{display:"flex",flexDirection:"column",gap:4}}><label style={S.label}>{l}</label>{t==="select"?<select style={S.input} value={v} onChange={e=>o(e.target.value)}>{opts.map(x=><option key={x.v} value={x.v}>{x.l}</option>)}</select>:<input style={S.input} type={t} value={v} onChange={e=>o(e.target.value)} placeholder={ph} disabled={dis}/>}</div>}

// ═══ STYLES ═══
const CSS=`@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&family=JetBrains+Mono:wght@400&display=swap');@keyframes spin{to{transform:rotate(360deg)}}input:focus,select:focus{border-color:#1a5ab8!important;box-shadow:0 0 0 3px rgba(26,90,184,0.1)!important;outline:none}button{transition:all 0.15s}button:hover{filter:brightness(1.05)}tr:hover td{background:rgba(10,35,81,0.02)!important}*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-thumb{background:#b0bec5;border-radius:3px}`;

const S={
  app:{fontFamily:"'Source Sans 3',sans-serif",background:"#f1f5f9",minHeight:"100vh",color:"#1e293b"},
  loading:{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100vh",background:"linear-gradient(135deg,#0a2351,#1a4a8a)"},
  spinner:{width:36,height:36,border:"3px solid rgba(255,255,255,0.2)",borderTopColor:"#c9a227",borderRadius:"50%",animation:"spin 0.8s linear infinite"},
  header:{background:"linear-gradient(135deg,#0a2351,#0d2d6b)",padding:"0 20px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12,borderBottom:"3px solid #c9a227",position:"sticky",top:0,zIndex:100,minHeight:56},
  brand:{display:"flex",alignItems:"center",gap:10},logoImg:{height:32,borderRadius:4},brandDivider:{width:1,height:24,background:"rgba(255,255,255,0.15)"},brandSub:{color:"rgba(255,255,255,0.5)",fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase"},
  nav:{display:"flex",gap:2,flexWrap:"wrap"},navBtn:{display:"flex",alignItems:"center",gap:5,padding:"8px 12px",border:"none",borderRadius:6,background:"transparent",color:"rgba(255,255,255,0.5)",cursor:"pointer",fontSize:13,fontWeight:500,fontFamily:"inherit"},navAct:{background:"rgba(201,162,39,0.15)",color:"#c9a227",fontWeight:700},
  main:{maxWidth:1300,margin:"0 auto",padding:"20px 16px 40px"},footer:{textAlign:"center",padding:"16px",fontSize:12,color:"#64748b",borderTop:"1px solid #e2e8f0",display:"flex",justifyContent:"center",gap:20},
  title:{fontSize:22,fontWeight:700,marginBottom:16,color:"#0a2351"},titleRow:{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,marginBottom:12},
  grid4:{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:14,marginBottom:20},
  card:{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 3px rgba(0,0,0,0.05)",border:"1px solid #e2e8f0",marginBottom:16},cardTitle:{fontSize:15,fontWeight:700,marginBottom:10,color:"#334155"},
  table:{width:"100%",borderCollapse:"collapse",fontSize:13},th:{padding:"8px 10px",textAlign:"left",fontWeight:600,fontSize:11,color:"#64748b",borderBottom:"2px solid #d0daea",whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.04em",background:"#f8fafc"},
  td:{padding:"8px 10px",borderBottom:"1px solid #eef2f7",whiteSpace:"nowrap",fontSize:13},tdM:{padding:"8px 10px",borderBottom:"1px solid #eef2f7",whiteSpace:"nowrap",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12},
  badge:{background:"#e8eef6",color:"#0a2351",padding:"2px 8px",borderRadius:6,fontWeight:700,fontSize:11,fontFamily:"monospace"},posBadge:{background:"#f0fdf4",color:"#059669",padding:"2px 8px",borderRadius:6,fontSize:12},
  btnPrimary:{padding:"9px 16px",background:"linear-gradient(135deg,#0a2351,#163a72)",color:"#fff",border:"none",borderRadius:8,fontWeight:600,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap",boxShadow:"0 2px 6px rgba(10,35,81,0.2)"},
  btnSec:{padding:"9px 16px",background:"#fff",color:"#0a2351",border:"1px solid #d0daea",borderRadius:8,fontWeight:600,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap"},
  btnGold:{padding:"9px 16px",background:"linear-gradient(135deg,#c9a227,#dbb84a)",color:"#fff",border:"none",borderRadius:8,fontWeight:700,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap",boxShadow:"0 2px 6px rgba(201,162,39,0.3)"},
  tblBtn:{padding:"3px 10px",background:"#fff",color:"#0a2351",border:"1px solid #d0daea",borderRadius:6,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",marginRight:3},
  fileLabel:{display:"inline-flex",alignItems:"center",gap:6,padding:"10px 18px",background:"#eff6ff",color:"#0a2351",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",border:"2px dashed #93c5fd",fontFamily:"inherit"},
  formGrid:{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10},
  label:{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.04em"},
  input:{padding:"8px 10px",border:"1px solid #d0daea",borderRadius:8,fontSize:13,color:"#1e293b",background:"#fff",width:"100%",boxSizing:"border-box",fontFamily:"inherit"},
  overlay:{position:"fixed",inset:0,background:"rgba(10,35,81,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16,backdropFilter:"blur(4px)"},
  modal:{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:480,boxShadow:"0 20px 60px rgba(10,35,81,0.25)"},
};