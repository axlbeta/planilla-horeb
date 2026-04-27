import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

// ─── Supabase DB functions ───
const db = {
  async getEmployees() {
    const { data, error } = await supabase.from("employees").select("*").eq("active", true).order("id");
    if (error) { console.error("getEmployees:", error); return []; }
    return data.map(e => ({ id: String(e.id), name: e.name, position: e.position, salary: parseFloat(e.salary) }));
  },
  async upsertEmployee(emp) {
    const { error } = await supabase.from("employees").upsert({ id: emp.id, name: emp.name, position: emp.position, salary: emp.salary, active: true, updated_at: new Date().toISOString() });
    if (error) console.error("upsertEmployee:", error);
  },
  async deleteEmployee(id) {
    const { error } = await supabase.from("employees").update({ active: false }).eq("id", id);
    if (error) console.error("deleteEmployee:", error);
  },

  async getClockEntries() {
    const { data, error } = await supabase.from("clock_entries").select("*").order("date", { ascending: false });
    if (error) { console.error("getClockEntries:", error); return []; }
    return data.map(e => ({
      id: e.id, employeeId: String(e.employee_id), date: String(e.date).slice(0, 10),
      checkIn: e.check_in, lunchOut: e.lunch_out, lunchIn: e.lunch_in, checkOut: e.check_out,
      punches: e.punches, name: ""
    }));
  },
  async insertClockEntries(entries) {
    const rows = entries.map(e => ({
      id: e.id, employee_id: e.employeeId, date: e.date,
      check_in: e.checkIn, lunch_out: e.lunchOut, lunch_in: e.lunchIn, check_out: e.checkOut,
      punches: e.punches
    }));
    const { error } = await supabase.from("clock_entries").upsert(rows, { onConflict: "id" });
    if (error) console.error("insertClockEntries:", error);
  },
  async deleteClockEntry(id) {
    const { error } = await supabase.from("clock_entries").delete().eq("id", id);
    if (error) console.error("deleteClockEntry:", error);
  },
  async deleteAllClockEntries() {
    const { error } = await supabase.from("clock_entries").delete().neq("id", "");
    if (error) console.error("deleteAllClock:", error);
  },

  async getPayrolls() {
    const { data, error } = await supabase.from("payrolls").select("*").order("id", { ascending: false });
    if (error) { console.error("getPayrolls:", error); return []; }
    return data;
  },
  async getPayrollRows(payrollId) {
    const { data, error } = await supabase.from("payroll_rows").select("*").eq("payroll_id", payrollId);
    if (error) { console.error("getPayrollRows:", error); return []; }
    return data.map(r => ({
      employeeId: r.employee_id, name: r.name, position: r.position,
      salary: +r.salary, daily: +r.daily, hourly: +r.hourly, days: r.days,
      effectiveHrs: +r.effective_hrs, baseSalary: +r.base_salary,
      ot: { 0.25: +r.ot_25, 0.5: +r.ot_50, 0.75: +r.ot_75, 1.0: +r.ot_100 },
      otPay: +r.ot_pay, fuel: +r.fuel, vacation: +r.vacation, incapacity: +r.incapacity,
      advance: +r.advance, dec4: +r.dec4, dec3: +r.dec3, otherDed: +r.other_ded,
      totalEarned: +r.total_earned, totalDeductions: +r.total_deductions, netPay: +r.net_pay,
    }));
  },
  async savePayroll(payroll) {
    // Check if period exists
    const { data: existing } = await supabase.from("payrolls").select("id").eq("period", payroll.period).maybeSingle();
    let payrollId;
    if (existing) {
      payrollId = existing.id;
      await supabase.from("payroll_rows").delete().eq("payroll_id", payrollId);
      await supabase.from("payrolls").update({
        week_num: payroll.weekNum, date_from: payroll.from, date_to: payroll.to,
        total_earned: payroll.rows.reduce((s, r) => s + r.totalEarned, 0),
        total_deductions: payroll.rows.reduce((s, r) => s + r.totalDeductions, 0),
        total_net: payroll.rows.reduce((s, r) => s + r.netPay, 0),
      }).eq("id", payrollId);
    } else {
      const { data, error } = await supabase.from("payrolls").insert({
        period: payroll.period, week_num: payroll.weekNum,
        date_from: payroll.from, date_to: payroll.to,
        total_earned: payroll.rows.reduce((s, r) => s + r.totalEarned, 0),
        total_deductions: payroll.rows.reduce((s, r) => s + r.totalDeductions, 0),
        total_net: payroll.rows.reduce((s, r) => s + r.netPay, 0),
      }).select("id").single();
      if (error) { console.error("savePayroll:", error); return; }
      payrollId = data.id;
    }
    // Insert rows
    const rows = payroll.rows.map(r => ({
      payroll_id: payrollId, employee_id: r.employeeId, name: r.name, position: r.position,
      salary: r.salary, daily: r.daily, hourly: r.hourly, days: r.days,
      effective_hrs: r.effectiveHrs, base_salary: r.baseSalary,
      ot_25: r.ot[0.25], ot_50: r.ot[0.5], ot_75: r.ot[0.75], ot_100: r.ot[1.0],
      ot_pay: r.otPay, fuel: r.fuel, vacation: r.vacation, incapacity: r.incapacity,
      advance: r.advance, dec4: r.dec4, dec3: r.dec3, other_ded: r.otherDed,
      total_earned: r.totalEarned, total_deductions: r.totalDeductions, net_pay: r.netPay,
    }));
    const { error: rowErr } = await supabase.from("payroll_rows").insert(rows);
    if (rowErr) console.error("savePayrollRows:", rowErr);
  },
  async deletePayroll(id) {
    await supabase.from("payroll_rows").delete().eq("payroll_id", id);
    await supabase.from("payrolls").delete().eq("id", id);
  },
};

// ─── Schedule & calculations ───
function getScheduledHours(dow) { return dow >= 1 && dow <= 4 ? 9 : dow === 5 ? 8 : 0; }

function formatL(n) {
  if (n == null || isNaN(n)) return "L. 0.00";
  return "L. " + Number(n).toLocaleString("es-HN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fN(n) { return (!n || isNaN(n) || n === 0) ? "" : Number(n).toLocaleString("es-HN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtTime(t) { return t ? new Date(t).toLocaleTimeString("es-HN", { hour: "2-digit", minute: "2-digit" }) : "—"; }
function fmtDate(d) { return new Date(d).toLocaleDateString("es-HN", { weekday: "short", day: "2-digit", month: "short" }); }

function parseClockCSV(csvText) {
  const lines = csvText.trim().split("\n");
  if (lines.length < 2) return [];
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",");
    if (cols.length < 10) continue;
    const time = cols[0].trim(), userId = cols[1].trim(), evento = cols[9].trim();
    if (!userId || evento === "Acceso denegado" || evento === "Dispositivo inicializado") continue;
    const parts = time.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
    if (!parts) continue;
    const [, mm, dd, yyyy, hh, mi, ss] = parts;
    records.push({ userId, dateStr: `${yyyy}-${mm}-${dd}`, dt: new Date(+yyyy, +mm - 1, +dd, +hh, +mi, +ss) });
  }
  return groupPunches(records);
}

function parseClockXLS(data) {
  try {
    const wb = XLSX.read(data, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (json.length < 2) return [];
    const records = [];
    for (let i = 1; i < json.length; i++) {
      const row = json[i];
      if (!row || row.length < 10) continue;
      const time = String(row[0] || "").trim(), userId = String(row[1] || "").trim(), evento = String(row[9] || "").trim();
      if (!userId || evento === "Acceso denegado" || evento === "Dispositivo inicializado") continue;
      const parts = time.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
      if (!parts) continue;
      const [, mm, dd, yyyy, hh, mi, ss] = parts;
      records.push({ userId, dateStr: `${yyyy}-${mm}-${dd}`, dt: new Date(+yyyy, +mm - 1, +dd, +hh, +mi, +ss) });
    }
    return groupPunches(records);
  } catch (e) { console.error("XLS parse:", e); return []; }
}

function groupPunches(records) {
  const groups = {};
  records.forEach((r) => {
    const key = `${r.userId}_${r.dateStr}`;
    if (!groups[key]) groups[key] = { userId: r.userId, date: r.dateStr, punches: [] };
    groups[key].punches.push(r.dt);
  });
  const entries = [];
  Object.values(groups).forEach((g) => {
    g.punches.sort((a, b) => a - b);
    const dd = [g.punches[0]];
    for (let i = 1; i < g.punches.length; i++) if (g.punches[i] - g.punches[i - 1] > 120000) dd.push(g.punches[i]);
    let e = { id: `${g.userId}_${g.date}`, employeeId: g.userId, date: g.date, checkIn: null, lunchOut: null, lunchIn: null, checkOut: null, punches: dd.length };
    if (dd.length >= 4) { e.checkIn = dd[0].toISOString(); e.lunchOut = dd[1].toISOString(); e.lunchIn = dd[2].toISOString(); e.checkOut = dd[dd.length - 1].toISOString(); }
    else if (dd.length === 3) { e.checkIn = dd[0].toISOString(); e.lunchOut = dd[1].toISOString(); e.checkOut = dd[2].toISOString(); }
    else if (dd.length === 2) { e.checkIn = dd[0].toISOString(); e.checkOut = dd[1].toISOString(); }
    else if (dd.length === 1) { e.checkIn = dd[0].toISOString(); }
    entries.push(e);
  });
  return entries.sort((a, b) => a.date.localeCompare(b.date));
}

function calcDayHours(entry) {
  if (!entry.checkIn || !entry.checkOut) return 0;
  let ms = new Date(entry.checkOut) - new Date(entry.checkIn);
  if (entry.lunchOut && entry.lunchIn) ms -= (new Date(entry.lunchIn) - new Date(entry.lunchOut));
  else if (entry.lunchOut) ms -= 3600000;
  else if (ms > 5 * 3600000) ms -= 3600000;
  return Math.max(0, ms / 3600000);
}

function calcOvertime(entries) {
  const byEmp = {};
  entries.forEach((e) => { if (!byEmp[e.employeeId]) byEmp[e.employeeId] = []; byEmp[e.employeeId].push(e); });
  const results = {};
  Object.entries(byEmp).forEach(([empId, days]) => {
    let totalEffective = 0, regularDays = 0, ot = { 0.25: 0, 0.5: 0, 0.75: 0, 1.0: 0 };
    days.forEach((day) => {
      if (!day.checkIn || !day.checkOut) return;
      const dow = new Date(day.checkIn).getDay(), hrs = calcDayHours(day), scheduled = getScheduledHours(dow);
      if (dow === 0 || dow === 6) { ot[1.0] += hrs; if (hrs > 0) regularDays++; totalEffective += hrs; return; }
      regularDays++; totalEffective += hrs;
      if (hrs <= scheduled) return;
      const extra = hrs - scheduled, cH = new Date(day.checkOut).getHours() + new Date(day.checkOut).getMinutes() / 60;
      if (cH <= 19) ot[0.25] += extra;
      else if (cH <= 21) { ot[0.25] += Math.min(1, extra); if (extra > 1) ot[0.5] += extra - 1; }
      else { ot[0.25] += Math.min(1, extra); const r = extra - 1; if (r > 0) { ot[0.5] += Math.min(2, r); if (r > 2) ot[0.75] += r - 2; } }
    });
    results[empId] = { totalEffective, regularDays, ot };
  });
  return results;
}

function printPayroll(payroll) {
  const rows = payroll.rows;
  const totals = { salary: rows.reduce((s, r) => s + (r.salary || 0), 0), base: rows.reduce((s, r) => s + r.baseSalary, 0), ot: rows.reduce((s, r) => s + r.otPay, 0), earned: rows.reduce((s, r) => s + r.totalEarned, 0), ded: rows.reduce((s, r) => s + r.totalDeductions, 0), net: rows.reduce((s, r) => s + r.netPay, 0) };
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Planilla ${payroll.period}</title>
<style>@page{size:landscape;margin:12mm}*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:9pt;color:#000}.header{text-align:center;margin-bottom:12px}.header h1{font-size:14pt;margin-bottom:2px}.header h2{font-size:11pt;font-weight:normal;color:#444}.header .period{font-size:10pt;margin-top:4px;font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:8px}th{background:#1e293b;color:#fff;padding:5px 4px;font-size:7.5pt;text-transform:uppercase;border:1px solid #334155;text-align:center}td{padding:4px;border:1px solid #cbd5e1;font-size:8pt}.r{text-align:right;font-family:'Courier New',monospace;font-size:8pt}.name{font-weight:600;white-space:nowrap}.total-row{background:#f1f5f9;font-weight:700}.total-row td{border-top:2px solid #334155;padding:6px 4px}.net{color:#059669;font-weight:700;font-size:9pt}.ot-val{color:#dc2626}.signatures{margin-top:40px;display:flex;justify-content:space-between}.sig-box{text-align:center;width:200px}.sig-line{border-top:1px solid #000;margin-top:50px;padding-top:4px;font-size:9pt}@media print{button{display:none!important}}.no-print{position:fixed;top:10px;right:10px;z-index:999}.print-btn{padding:10px 24px;background:#3b82f6;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;margin-right:8px}.close-btn{padding:10px 24px;background:#64748b;color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer}</style></head><body>
<div class="no-print"><button class="print-btn" onclick="window.print()">🖨️ Imprimir</button><button class="close-btn" onclick="window.close()">✕ Cerrar</button></div>
<div class="header"><h1>Impresos Horeb</h1><h2>Planilla de Empleados</h2><div class="period">${payroll.period}</div></div>
<table><thead><tr><th rowspan="2">Cód.</th><th rowspan="2">Nombre del Empleado</th><th rowspan="2">Posición</th><th rowspan="2">Salario<br>Mensual</th><th rowspan="2">Salario<br>Diario</th><th rowspan="2">Días</th><th rowspan="2">Salario</th><th colspan="4">Horas Extras</th><th rowspan="2">Salario<br>Hr</th><th rowspan="2">Total<br>Extras</th><th rowspan="2">Combust.</th><th rowspan="2">Vacac.</th><th rowspan="2">Incap.</th><th rowspan="2">Adelanto<br>Ajustes</th><th rowspan="2">Dec.<br>Cuarto</th><th rowspan="2">Dec.<br>Tercero</th><th rowspan="2">Total<br>Devengado</th><th rowspan="2">Otras<br>Deduc.</th><th rowspan="2">Neto a<br>Pagar</th></tr><tr><th>0.25</th><th>0.50</th><th>0.75</th><th>1.00</th></tr></thead><tbody>
${rows.map(r => `<tr><td class="c">${r.employeeId}</td><td class="name">${r.name}</td><td>${r.position}</td><td class="r">${fN(r.salary)}</td><td class="r">${fN(r.daily)}</td><td class="c">${r.days}</td><td class="r">${fN(r.baseSalary)}</td><td class="r${r.ot[0.25] > 0 ? ' ot-val' : ''}">${r.ot[0.25] > 0 ? r.ot[0.25].toFixed(1) : ''}</td><td class="r${r.ot[0.5] > 0 ? ' ot-val' : ''}">${r.ot[0.5] > 0 ? r.ot[0.5].toFixed(1) : ''}</td><td class="r${r.ot[0.75] > 0 ? ' ot-val' : ''}">${r.ot[0.75] > 0 ? r.ot[0.75].toFixed(1) : ''}</td><td class="r${r.ot[1.0] > 0 ? ' ot-val' : ''}">${r.ot[1.0] > 0 ? r.ot[1.0].toFixed(1) : ''}</td><td class="r">${fN(r.hourly)}</td><td class="r ot-val">${fN(r.otPay)}</td><td class="r">${fN(r.fuel)}</td><td class="r">${fN(r.vacation)}</td><td class="r">${fN(r.incapacity)}</td><td class="r">${fN(r.advance)}</td><td class="r">${fN(r.dec4)}</td><td class="r">${fN(r.dec3)}</td><td class="r" style="font-weight:600">${fN(r.totalEarned)}</td><td class="r">${fN(r.otherDed)}</td><td class="r net">${fN(r.netPay)}</td></tr>`).join("")}
<tr class="total-row"><td colspan="3" style="text-align:right">TOTALES</td><td class="r">${fN(totals.salary)}</td><td></td><td></td><td class="r">${fN(totals.base)}</td><td colspan="4"></td><td></td><td class="r ot-val">${fN(totals.ot)}</td><td colspan="6"></td><td class="r" style="font-weight:700">${fN(totals.earned)}</td><td class="r">${fN(totals.ded)}</td><td class="r net" style="font-size:10pt">${fN(totals.net)}</td></tr>
</tbody></table>
<div class="signatures"><div class="sig-box"><div class="sig-line">Elaborado por</div></div><div class="sig-box"><div class="sig-line">Revisado por</div></div><div class="sig-box"><div class="sig-line">Autorizado por</div></div></div></body></html>`;
  const w = window.open("", "_blank");
  if (w) { w.document.write(html); w.document.close(); }
}

// ═══════════════════════════════════════
const TABS = [
  { id: "dash", label: "Dashboard", icon: "📊" },
  { id: "emp", label: "Empleados", icon: "👥" },
  { id: "clock", label: "Reloj", icon: "🕐" },
  { id: "pay", label: "Planilla", icon: "📋" },
  { id: "hist", label: "Historial", icon: "📁" },
];

export default function App() {
  const [tab, setTab] = useState("dash");
  const [employees, setEmployees] = useState([]);
  const [clockEntries, setClockEntries] = useState([]);
  const [payrolls, setPayrolls] = useState([]);
  const [loading, setLoading] = useState(true);

  const refresh = useCallback(async () => {
    const [emps, clock, pays] = await Promise.all([db.getEmployees(), db.getClockEntries(), db.getPayrolls()]);
    setEmployees(emps);
    setClockEntries(clock);
    setPayrolls(pays);
  }, []);

  useEffect(() => { refresh().then(() => setLoading(false)); }, [refresh]);

  if (loading) return (
    <div style={S.loading}><div style={S.spinner} /><p style={{ color: "#94a3b8", marginTop: 16, fontFamily: "'Segoe UI',sans-serif" }}>Conectando con Supabase...</p>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>
  );

  const TabComp = { dash: DashboardTab, emp: EmployeesTab, clock: ClockTab, pay: PayrollTab, hist: HistoryTab }[tab];

  return (
    <div style={S.app}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        input:focus,select:focus{border-color:#3b82f6!important;box-shadow:0 0 0 3px rgba(59,130,246,0.12)!important}
        button:hover{filter:brightness(1.05)}
        tr:hover{background:#f8fafc!important}
        *{box-sizing:border-box}
      `}</style>
      <header style={S.header}>
        <div style={S.brand}>
          <div style={S.logo}>H</div>
          <div><div style={S.brandName}>Impresos Horeb</div><div style={S.brandSub}>Sistema de Planilla</div></div>
        </div>
        <nav style={S.nav}>
          {TABS.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ ...S.navBtn, ...(tab === t.id ? S.navAct : {}) }}>
              <span style={{ fontSize: 16 }}>{t.icon}</span><span>{t.label}</span>
            </button>
          ))}
        </nav>
      </header>
      <main style={S.main}><div style={{ animation: "fadeIn 0.3s ease" }}>
        <TabComp employees={employees} refresh={refresh} clockEntries={clockEntries} payrolls={payrolls} />
      </div></main>
    </div>
  );
}

// ═══════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════
function DashboardTab({ employees, payrolls }) {
  const totalSalary = employees.reduce((s, e) => s + e.salary, 0);
  const last = payrolls.length > 0 ? payrolls[0] : null;
  const [lastRows, setLastRows] = useState(null);
  useEffect(() => { if (last) db.getPayrollRows(last.id).then(setLastRows); }, [last]);

  return (
    <div>
      <h2 style={S.title}>Panel de Control</h2>
      <div style={S.grid4}>
        <SC color="#3b82f6" label="Empleados" value={employees.length} />
        <SC color="#8b5cf6" label="Nómina Mensual" value={formatL(totalSalary)} />
        <SC color="#f59e0b" label="Planillas" value={payrolls.length} />
        <SC color="#059669" label="Última" value={last?.period || "—"} />
      </div>
      <div style={S.card}>
        <h3 style={S.cardTitle}>Horario</h3>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 8 }}>
          {["Lunes", "Martes", "Miércoles", "Jueves"].map((d) => <div key={d} style={S.schedBadge}><strong>{d}</strong><br />8am-6pm<br /><span style={{ fontSize: 11, color: "#64748b" }}>9hrs</span></div>)}
          <div style={{ ...S.schedBadge, borderColor: "#f59e0b" }}><strong>Viernes</strong><br />8am-5pm<br /><span style={{ fontSize: 11, color: "#64748b" }}>8hrs</span></div>
        </div>
        <p style={{ fontSize: 12, color: "#64748b", marginTop: 8 }}>Almuerzo: 12-1pm · Sábado compensado L-J · 44hrs/semana</p>
      </div>
      {last && lastRows && (
        <div style={S.card}>
          <div style={S.titleRow}>
            <h3 style={S.cardTitle}>Última: {last.period}</h3>
            <button style={{ ...S.btnPrimary, background: "#059669", fontSize: 12 }} onClick={() => printPayroll({ period: last.period, rows: lastRows })}>🖨️ Imprimir</button>
          </div>
          <div style={S.summRow}><span>Devengado:</span><strong>{formatL(+last.total_earned)}</strong></div>
          <div style={S.summRow}><span>Deducciones:</span><strong style={{ color: "#dc2626" }}>{formatL(+last.total_deductions)}</strong></div>
          <div style={{ ...S.summRow, borderTop: "2px solid #e2e8f0", paddingTop: 8, marginTop: 4 }}>
            <span style={{ fontWeight: 700 }}>Neto:</span>
            <strong style={{ color: "#059669", fontSize: 20 }}>{formatL(+last.total_net)}</strong>
          </div>
        </div>
      )}
    </div>
  );
}
function SC({ color, label, value }) {
  return <div style={{ ...S.card, borderLeft: `4px solid ${color}`, padding: 16 }}>
    <div style={{ fontSize: 12, color: "#64748b", fontWeight: 500, marginBottom: 4 }}>{label}</div>
    <div style={{ fontSize: 20, fontWeight: 700 }}>{value}</div>
  </div>;
}

// ═══════════════════════════════════════
// EMPLOYEES
// ═══════════════════════════════════════
function EmployeesTab({ employees, refresh }) {
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({ id: "", name: "", position: "", salary: "" });
  const [saving, setSaving] = useState(false);

  const openAdd = () => { setForm({ id: "", name: "", position: "", salary: "" }); setModal("new"); };
  const openEdit = (e) => { setForm({ ...e, salary: String(e.salary) }); setModal(e.id); };
  const doSave = async () => {
    const emp = { ...form, salary: parseFloat(form.salary) || 0 };
    if (!emp.id || !emp.name) return;
    setSaving(true);
    await db.upsertEmployee(emp);
    await refresh();
    setSaving(false);
    setModal(null);
  };
  const doDelete = async (id) => { if (confirm("¿Eliminar?")) { await db.deleteEmployee(id); await refresh(); } };

  return (
    <div>
      <div style={S.titleRow}><h2 style={S.title}>Empleados</h2><button style={S.btnPrimary} onClick={openAdd}>+ Agregar</button></div>
      {modal && (
        <div style={S.overlay}><div style={S.modal}>
          <h3 style={{ fontSize: 18, fontWeight: 700, marginBottom: 16 }}>{modal === "new" ? "Nuevo Empleado" : "Editar Empleado"}</h3>
          <div style={S.formGrid}>
            <Field label="Código" value={form.id} onChange={(v) => setForm({ ...form, id: v })} ph="008" disabled={modal !== "new"} />
            <Field label="Nombre" value={form.name} onChange={(v) => setForm({ ...form, name: v })} ph="Nombre completo" />
            <Field label="Posición" value={form.position} onChange={(v) => setForm({ ...form, position: v })} ph="Prensista" />
            <Field label="Salario Mensual" value={form.salary} onChange={(v) => setForm({ ...form, salary: v })} ph="0" type="number" />
          </div>
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 20 }}>
            <button style={S.btnSec} onClick={() => setModal(null)}>Cancelar</button>
            <button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving ? "Guardando..." : "Guardar"}</button>
          </div>
        </div></div>
      )}
      <div style={S.card}><div style={{ overflowX: "auto" }}>
        <table style={S.table}><thead><tr>
          {["Cód", "Nombre", "Posición", "Sal.Mensual", "Sal.Diario", "Sal./Hora", ""].map((h, i) => (
            <th key={i} style={{ ...S.th, textAlign: i >= 3 && i <= 5 ? "right" : "left" }}>{h}</th>
          ))}
        </tr></thead><tbody>
          {employees.map((e) => (
            <tr key={e.id}>
              <td style={S.td}><span style={S.badge}>{e.id}</span></td>
              <td style={{ ...S.td, fontWeight: 600 }}>{e.name}</td>
              <td style={S.td}><span style={S.posBadge}>{e.position}</span></td>
              <td style={S.tdM}>{formatL(e.salary)}</td>
              <td style={S.tdM}>{formatL(e.salary / 30)}</td>
              <td style={S.tdM}>{formatL(e.salary / 30 / 8)}</td>
              <td style={{ ...S.td, textAlign: "center" }}>
                <button style={S.iconBtn} onClick={() => openEdit(e)}>✏️</button>
                <button style={S.iconBtn} onClick={() => doDelete(e.id)}>🗑️</button>
              </td>
            </tr>
          ))}
        </tbody></table>
      </div></div>
    </div>
  );
}

// ═══════════════════════════════════════
// CLOCK
// ═══════════════════════════════════════
function ClockTab({ employees, clockEntries, refresh }) {
  const [mode, setMode] = useState("import");
  const [importResult, setImportResult] = useState(null);
  const [importing, setImporting] = useState(false);
  const [mf, setMf] = useState({ empId: "", date: new Date().toISOString().slice(0, 10), cin: "08:00", lout: "12:00", lin: "13:00", cout: "18:00" });
  const [filterEmp, setFilterEmp] = useState("");
  const [filterMonth, setFilterMonth] = useState(new Date().toISOString().slice(0, 7));

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImporting(true);
    const isXLS = file.name.match(/\.xls[xm]?$/i);
    if (isXLS) {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        let parsed = parseClockXLS(new Uint8Array(evt.target.result));
        if (parsed.length === 0) {
          const tr = new FileReader();
          tr.onload = async (e2) => { let t = e2.target.result; if (t.includes("<table") || t.includes("<html")) { const doc = new DOMParser().parseFromString(t, "text/html"); const lines = []; doc.querySelectorAll("tr").forEach(r => { lines.push(Array.from(r.querySelectorAll("td,th")).map(c => c.textContent.trim()).join(",")); }); t = lines.join("\n"); } await finishImport(parseClockCSV(t)); };
          tr.readAsText(file, "UTF-8"); return;
        }
        await finishImport(parsed);
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = async (evt) => await finishImport(parseClockCSV(evt.target.result));
      reader.readAsText(file, "UTF-8");
    }
    e.target.value = "";
  };

  const finishImport = async (parsed) => {
    if (parsed.length === 0) { setImportResult({ error: "No se encontraron registros válidos." }); setImporting(false); return; }
    const existingIds = new Set(clockEntries.map((c) => c.id));
    const newEntries = parsed.filter((p) => !existingIds.has(p.id));
    if (newEntries.length > 0) await db.insertClockEntries(newEntries);
    await refresh();
    setImportResult({ added: newEntries.length, total: parsed.length, skipped: parsed.length - newEntries.length });
    setImporting(false);
  };

  const addManual = async () => {
    if (!mf.empId) return;
    const entry = { id: `${mf.empId}_${mf.date}_m${Date.now()}`, employeeId: mf.empId, date: mf.date, checkIn: `${mf.date}T${mf.cin}:00`, lunchOut: `${mf.date}T${mf.lout}:00`, lunchIn: `${mf.date}T${mf.lin}:00`, checkOut: `${mf.date}T${mf.cout}:00`, punches: 4 };
    await db.insertClockEntries([entry]);
    await refresh();
  };

  const removeEntry = async (id) => { await db.deleteClockEntry(id); await refresh(); };
  const clearAll = async () => { if (confirm("¿Borrar TODAS las marcaciones?")) { await db.deleteAllClockEntries(); await refresh(); } };

  const filtered = clockEntries.filter((e) => (!filterEmp || e.employeeId === filterEmp) && (!filterMonth || e.date.startsWith(filterMonth))).sort((a, b) => b.date.localeCompare(a.date) || (a.employeeId || "").localeCompare(b.employeeId || ""));
  const getN = (id) => employees.find((e) => e.id === id)?.name || `ID ${id}`;

  return (
    <div>
      <h2 style={S.title}>Reloj Marcador</h2>
      <div style={S.card}>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <button style={mode === "import" ? S.btnPrimary : S.btnSec} onClick={() => setMode("import")}>📥 Importar Archivo</button>
          <button style={mode === "manual" ? S.btnPrimary : S.btnSec} onClick={() => setMode("manual")}>✍️ Manual</button>
        </div>
        {mode === "import" ? (
          <div>
            <p style={{ fontSize: 13, color: "#475569", marginBottom: 12 }}>Sube el archivo <strong>.xls</strong> o <strong>.csv</strong> del reloj marcador.</p>
            <label style={S.fileLabel}>{importing ? "⏳ Importando..." : "📄 Seleccionar archivo del reloj"}<input type="file" accept=".csv,.txt,.xls,.xlsx" onChange={handleFile} style={{ display: "none" }} disabled={importing} /></label>
            {importResult && (
              <div style={{ marginTop: 12, padding: 12, borderRadius: 8, background: importResult.error ? "#fef2f2" : "#f0fdf4", border: `1px solid ${importResult.error ? "#fecaca" : "#bbf7d0"}` }}>
                {importResult.error ? <p style={{ color: "#dc2626", fontSize: 13 }}>⚠️ {importResult.error}</p>
                  : <p style={{ color: "#059669", fontSize: 13 }}>✅ <strong>{importResult.added}</strong> registros guardados en Supabase.{importResult.skipped > 0 && ` (${importResult.skipped} duplicados omitidos)`}</p>}
              </div>
            )}
          </div>
        ) : (
          <div style={S.formGrid}>
            <Field label="Empleado" value={mf.empId} onChange={(v) => setMf({ ...mf, empId: v })} type="select" options={[{ v: "", l: "— Seleccionar —" }, ...employees.map((e) => ({ v: e.id, l: `${e.id} - ${e.name}` }))]} />
            <Field label="Fecha" value={mf.date} onChange={(v) => setMf({ ...mf, date: v })} type="date" />
            <Field label="Entrada" value={mf.cin} onChange={(v) => setMf({ ...mf, cin: v })} type="time" />
            <Field label="Sale Almuerzo" value={mf.lout} onChange={(v) => setMf({ ...mf, lout: v })} type="time" />
            <Field label="Regresa" value={mf.lin} onChange={(v) => setMf({ ...mf, lin: v })} type="time" />
            <Field label="Salida" value={mf.cout} onChange={(v) => setMf({ ...mf, cout: v })} type="time" />
            <div style={{ display: "flex", alignItems: "flex-end" }}><button style={S.btnPrimary} onClick={addManual}>Registrar</button></div>
          </div>
        )}
      </div>
      <div style={S.card}>
        <div style={S.titleRow}><h3 style={S.cardTitle}>Registros ({filtered.length})</h3>{clockEntries.length > 0 && <button style={{ ...S.btnSec, fontSize: 12, padding: "6px 10px", color: "#dc2626" }} onClick={clearAll}>🗑️ Borrar Todo</button>}</div>
        <div style={{ display: "flex", gap: 12, marginBottom: 12, flexWrap: "wrap" }}>
          <select style={{ ...S.input, width: 200 }} value={filterEmp} onChange={(e) => setFilterEmp(e.target.value)}><option value="">Todos</option>{employees.map((e) => <option key={e.id} value={e.id}>{e.id} - {e.name}</option>)}</select>
          <input style={{ ...S.input, width: 160 }} type="month" value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} />
        </div>
        {filtered.length === 0 ? <p style={{ color: "#94a3b8" }}>No hay registros.</p> : (
          <div style={{ overflowX: "auto" }}><table style={S.table}><thead><tr>
            {["Empleado", "Fecha", "Entrada", "Sal.Alm.", "Reg.Alm.", "Salida", "Hrs", "Extra", ""].map((h, i) => <th key={i} style={{ ...S.th, textAlign: i >= 6 ? "right" : "left", ...(i === 8 ? { textAlign: "center" } : {}) }}>{h}</th>)}
          </tr></thead><tbody>
            {filtered.map((e) => {
              const hrs = calcDayHours(e), dow = new Date(e.date + "T12:00:00").getDay(), extra = Math.max(0, hrs - getScheduledHours(dow)), isWE = dow === 0 || dow === 6;
              return (
                <tr key={e.id} style={isWE ? { background: "#fef3c7" } : {}}>
                  <td style={{ ...S.td, fontWeight: 600, whiteSpace: "nowrap" }}>{getN(e.employeeId)}</td>
                  <td style={S.td}>{fmtDate(e.date + "T12:00:00")}{isWE && <span style={{ marginLeft: 4, fontSize: 10, color: "#d97706", fontWeight: 700 }}>{dow === 0 ? "DOM" : "SÁB"}</span>}</td>
                  <td style={{ ...S.td, color: "#059669", fontWeight: 500 }}>{fmtTime(e.checkIn)}</td>
                  <td style={{ ...S.td, color: "#f59e0b" }}>{fmtTime(e.lunchOut)}</td>
                  <td style={{ ...S.td, color: "#f59e0b" }}>{fmtTime(e.lunchIn)}</td>
                  <td style={{ ...S.td, color: "#dc2626", fontWeight: 500 }}>{fmtTime(e.checkOut)}</td>
                  <td style={{ ...S.td, textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontSize: 13 }}>{hrs.toFixed(1)}h</td>
                  <td style={{ ...S.td, textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontSize: 13, color: extra > 0 ? "#dc2626" : "#cbd5e1", fontWeight: extra > 0 ? 600 : 400 }}>{extra > 0 ? `+${extra.toFixed(1)}h` : "—"}</td>
                  <td style={{ ...S.td, textAlign: "center" }}><button style={S.iconBtn} onClick={() => removeEntry(e.id)}>🗑️</button></td>
                </tr>
              );
            })}
          </tbody></table></div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════
// PAYROLL
// ═══════════════════════════════════════
function PayrollTab({ employees, clockEntries, refresh }) {
  const [weekNum, setWeekNum] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [result, setResult] = useState(null);
  const [adj, setAdj] = useState({});
  const [saving, setSaving] = useState(false);

  const generate = () => {
    if (!from || !to) return alert("Selecciona las fechas.");
    const fromStr = from.slice(0, 10);
    const toStr = to.slice(0, 10);
    const pe = clockEntries.filter((c) => {
      const d = String(c.date).slice(0, 10);
      return d >= fromStr && d <= toStr;
    });
    console.log("Clock entries in period:", pe.length, "from", fromStr, "to", toStr);
    console.log("All clock dates:", clockEntries.map(c => c.date).slice(0, 10));
    const otData = calcOvertime(pe);
    const rows = employees.map((emp) => {
      const d = otData[emp.id] || { totalEffective: 0, regularDays: 0, ot: { 0.25: 0, 0.5: 0, 0.75: 0, 1.0: 0 } };
      const daily = emp.salary / 30, hourly = daily / 8, baseSalary = daily * d.regularDays;
      const otPay = Object.entries(d.ot).reduce((sum, [rate, hrs]) => sum + hrs * hourly * (1 + parseFloat(rate)), 0);
      const a = adj[emp.id] || {};
      const fuel = +a.fuel || 0, vacation = +a.vacation || 0, incapacity = +a.incapacity || 0, advance = +a.advance || 0, dec4 = +a.dec4 || 0, dec3 = +a.dec3 || 0, otherDed = +a.otherDed || 0;
      const totalEarned = baseSalary + otPay + fuel + vacation + incapacity + dec4 + dec3;
      return { employeeId: emp.id, name: emp.name, position: emp.position, salary: emp.salary, daily, hourly, days: d.regularDays, effectiveHrs: d.totalEffective, baseSalary, ot: d.ot, otPay, fuel, vacation, incapacity, advance, dec4, dec3, otherDed, totalEarned, totalDeductions: advance + otherDed, netPay: totalEarned - advance - otherDed };
    });
    setResult({ period: `WK${weekNum || "?"} ${from} al ${to}`, rows, from, to, weekNum });
  };

  const doSave = async () => {
    if (!result) return;
    setSaving(true);
    await db.savePayroll(result);
    await refresh();
    setSaving(false);
    alert("✅ Planilla guardada en Supabase.");
  };

  const updAdj = (eId, f, v) => setAdj((p) => ({ ...p, [eId]: { ...(p[eId] || {}), [f]: v } }));

  return (
    <div>
      <h2 style={S.title}>Generar Planilla</h2>
      <div style={S.card}>
        <h3 style={S.cardTitle}>Período</h3>
        <div style={S.formGrid}>
          <Field label="Semana #" value={weekNum} onChange={setWeekNum} ph="3" />
          <Field label="Desde" value={from} onChange={setFrom} type="date" />
          <Field label="Hasta" value={to} onChange={setTo} type="date" />
          <div style={{ display: "flex", alignItems: "flex-end" }}><button style={S.btnPrimary} onClick={generate}>📋 Generar</button></div>
        </div>
      </div>
      {result && (<>
        <div style={S.card}>
          <div style={S.titleRow}><h3 style={S.cardTitle}>{result.period}</h3>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={{ ...S.btnPrimary, background: "#7c3aed" }} onClick={() => printPayroll(result)}>🖨️ Imprimir</button>
              <button style={{ ...S.btnPrimary, background: "#059669" }} onClick={doSave} disabled={saving}>{saving ? "Guardando..." : "💾 Guardar"}</button>
            </div>
          </div>
          <div style={{ overflowX: "auto" }}><table style={S.table}><thead><tr>
            {["Cód", "Nombre", "Pos.", "Sal.M.", "Días", "Hrs", "Salario", "OT25", "OT50", "OT75", "OT100", "Tot.OT", "Dev.", "Ded.", "Neto"].map((h, i) => <th key={i} style={{ ...S.th, fontSize: 10, textAlign: i >= 3 ? "right" : "left" }}>{h}</th>)}
          </tr></thead><tbody>
            {result.rows.map((r) => (
              <tr key={r.employeeId}>
                <td style={S.td}><span style={S.badge}>{r.employeeId}</span></td>
                <td style={{ ...S.td, fontWeight: 600, whiteSpace: "nowrap", fontSize: 12 }}>{r.name}</td>
                <td style={{ ...S.td, fontSize: 11 }}>{r.position}</td>
                <td style={S.tdM}>{formatL(r.salary)}</td>
                <td style={{ ...S.tdM, fontWeight: 600 }}>{r.days}</td>
                <td style={S.tdM}>{r.effectiveHrs.toFixed(1)}h</td>
                <td style={S.tdM}>{formatL(r.baseSalary)}</td>
                <td style={{ ...S.tdM, color: r.ot[0.25] > 0 ? "#d97706" : "#cbd5e1" }}>{r.ot[0.25] > 0 ? r.ot[0.25].toFixed(1) + "h" : "—"}</td>
                <td style={{ ...S.tdM, color: r.ot[0.5] > 0 ? "#ea580c" : "#cbd5e1" }}>{r.ot[0.5] > 0 ? r.ot[0.5].toFixed(1) + "h" : "—"}</td>
                <td style={{ ...S.tdM, color: r.ot[0.75] > 0 ? "#dc2626" : "#cbd5e1" }}>{r.ot[0.75] > 0 ? r.ot[0.75].toFixed(1) + "h" : "—"}</td>
                <td style={{ ...S.tdM, color: r.ot[1.0] > 0 ? "#7c3aed" : "#cbd5e1" }}>{r.ot[1.0] > 0 ? r.ot[1.0].toFixed(1) + "h" : "—"}</td>
                <td style={{ ...S.tdM, fontWeight: 600, color: r.otPay > 0 ? "#dc2626" : "#cbd5e1" }}>{formatL(r.otPay)}</td>
                <td style={S.tdM}>{formatL(r.totalEarned)}</td>
                <td style={{ ...S.tdM, color: "#dc2626" }}>{formatL(r.totalDeductions)}</td>
                <td style={{ ...S.tdM, fontWeight: 700, color: "#059669" }}>{formatL(r.netPay)}</td>
              </tr>
            ))}
          </tbody><tfoot><tr style={{ background: "#f1f5f9" }}>
            <td colSpan={6} style={{ ...S.td, fontWeight: 700 }}>TOTALES</td>
            <td style={{ ...S.tdM, fontWeight: 700 }}>{formatL(result.rows.reduce((s, r) => s + r.baseSalary, 0))}</td>
            <td colSpan={4}></td>
            <td style={{ ...S.tdM, fontWeight: 700, color: "#dc2626" }}>{formatL(result.rows.reduce((s, r) => s + r.otPay, 0))}</td>
            <td style={{ ...S.tdM, fontWeight: 700 }}>{formatL(result.rows.reduce((s, r) => s + r.totalEarned, 0))}</td>
            <td style={{ ...S.tdM, fontWeight: 700, color: "#dc2626" }}>{formatL(result.rows.reduce((s, r) => s + r.totalDeductions, 0))}</td>
            <td style={{ ...S.tdM, fontWeight: 700, color: "#059669", fontSize: 14 }}>{formatL(result.rows.reduce((s, r) => s + r.netPay, 0))}</td>
          </tr></tfoot></table></div>
        </div>
        <div style={S.card}>
          <h3 style={S.cardTitle}>Ajustes</h3>
          <p style={{ fontSize: 12, color: "#64748b", marginBottom: 8 }}>Ingresa montos y genera de nuevo para recalcular.</p>
          <div style={{ overflowX: "auto" }}><table style={S.table}><thead><tr>
            {["Empleado", "Combustible", "Vacaciones", "Incapacidad", "Adelanto", "Dec.4to", "Dec.3ro", "Otras Ded."].map((h, i) => <th key={i} style={S.th}>{h}</th>)}
          </tr></thead><tbody>
            {employees.map((emp) => { const a = adj[emp.id] || {}; return (
              <tr key={emp.id}><td style={{ ...S.td, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>{emp.name}</td>
              {["fuel", "vacation", "incapacity", "advance", "dec4", "dec3", "otherDed"].map((f) => <td key={f} style={S.td}><input style={{ ...S.input, width: 80, padding: "4px 6px", fontSize: 12 }} type="number" value={a[f] || ""} onChange={(e) => updAdj(emp.id, f, e.target.value)} placeholder="0" /></td>)}</tr>
            ); })}
          </tbody></table></div>
        </div>
      </>)}
    </div>
  );
}

// ═══════════════════════════════════════
// HISTORY
// ═══════════════════════════════════════
function HistoryTab({ payrolls, refresh }) {
  const [sel, setSel] = useState(null);
  const [rows, setRows] = useState(null);
  const [loadingRows, setLoadingRows] = useState(false);

  const selectPayroll = async (i) => {
    setSel(i);
    setLoadingRows(true);
    const r = await db.getPayrollRows(payrolls[i].id);
    setRows(r);
    setLoadingRows(false);
  };

  const del = async (i) => {
    if (!confirm("¿Eliminar esta planilla?")) return;
    await db.deletePayroll(payrolls[i].id);
    await refresh();
    setSel(null); setRows(null);
  };

  return (
    <div>
      <h2 style={S.title}>Historial de Planillas</h2>
      {payrolls.length === 0 ? <div style={S.card}><p style={{ color: "#94a3b8" }}>No hay planillas guardadas.</p></div> : (<>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
          {payrolls.map((p, i) => <button key={i} onClick={() => selectPayroll(i)} style={sel === i ? S.btnPrimary : S.btnSec}>{p.period}</button>)}
        </div>
        {sel !== null && rows && !loadingRows && (
          <div style={S.card}>
            <div style={S.titleRow}><h3 style={S.cardTitle}>{payrolls[sel].period}</h3>
              <div style={{ display: "flex", gap: 8 }}>
                <button style={{ ...S.btnPrimary, background: "#7c3aed" }} onClick={() => printPayroll({ period: payrolls[sel].period, rows })}>🖨️ Imprimir</button>
                <button style={{ ...S.btnSec, color: "#dc2626", fontSize: 12 }} onClick={() => del(sel)}>🗑️ Eliminar</button>
              </div>
            </div>
            <div style={{ overflowX: "auto" }}><table style={S.table}><thead><tr>
              {["Cód", "Nombre", "Posición", "Días", "Hrs", "Salario", "H.Extra", "Devengado", "Deducc.", "Neto"].map((h, i) => <th key={i} style={{ ...S.th, textAlign: i >= 3 ? "right" : "left" }}>{h}</th>)}
            </tr></thead><tbody>
              {rows.map((r) => (
                <tr key={r.employeeId}>
                  <td style={S.td}><span style={S.badge}>{r.employeeId}</span></td>
                  <td style={{ ...S.td, fontWeight: 600 }}>{r.name}</td>
                  <td style={S.td}>{r.position}</td>
                  <td style={S.tdM}>{r.days}</td>
                  <td style={S.tdM}>{r.effectiveHrs?.toFixed(1) || "—"}h</td>
                  <td style={S.tdM}>{formatL(r.baseSalary)}</td>
                  <td style={{ ...S.tdM, color: r.otPay > 0 ? "#dc2626" : "#cbd5e1" }}>{formatL(r.otPay)}</td>
                  <td style={S.tdM}>{formatL(r.totalEarned)}</td>
                  <td style={{ ...S.tdM, color: "#dc2626" }}>{formatL(r.totalDeductions)}</td>
                  <td style={{ ...S.tdM, fontWeight: 700, color: "#059669" }}>{formatL(r.netPay)}</td>
                </tr>
              ))}
            </tbody><tfoot><tr style={{ background: "#f1f5f9" }}>
              <td colSpan={5} style={{ ...S.td, fontWeight: 700 }}>TOTALES</td>
              <td style={{ ...S.tdM, fontWeight: 700 }}>{formatL(rows.reduce((s, r) => s + r.baseSalary, 0))}</td>
              <td style={{ ...S.tdM, fontWeight: 700, color: "#dc2626" }}>{formatL(rows.reduce((s, r) => s + r.otPay, 0))}</td>
              <td style={{ ...S.tdM, fontWeight: 700 }}>{formatL(rows.reduce((s, r) => s + r.totalEarned, 0))}</td>
              <td style={{ ...S.tdM, fontWeight: 700, color: "#dc2626" }}>{formatL(rows.reduce((s, r) => s + r.totalDeductions, 0))}</td>
              <td style={{ ...S.tdM, fontWeight: 700, color: "#059669", fontSize: 15 }}>{formatL(rows.reduce((s, r) => s + r.netPay, 0))}</td>
            </tr></tfoot></table></div>
          </div>
        )}
        {loadingRows && <div style={S.card}><p style={{ color: "#64748b" }}>Cargando detalle...</p></div>}
      </>)}
    </div>
  );
}

// ═══════════════════════════════════════
function Field({ label, value, onChange, ph, type = "text", options, disabled }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <label style={S.label}>{label}</label>
      {type === "select" ? <select style={S.input} value={value} onChange={(e) => onChange(e.target.value)}>{options.map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}</select>
        : <input style={S.input} type={type} value={value} onChange={(e) => onChange(e.target.value)} placeholder={ph} disabled={disabled} />}
    </div>
  );
}

const S = {
  app: { fontFamily: "'DM Sans','Segoe UI',sans-serif", background: "#f1f5f9", minHeight: "100vh", color: "#1e293b" },
  loading: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", background: "#0f172a" },
  spinner: { width: 40, height: 40, border: "4px solid #334155", borderTopColor: "#3b82f6", borderRadius: "50%", animation: "spin 0.8s linear infinite" },
  header: { background: "linear-gradient(135deg,#0f172a,#1e293b)", padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, borderBottom: "3px solid #3b82f6", position: "sticky", top: 0, zIndex: 100 },
  brand: { display: "flex", alignItems: "center", gap: 10 },
  logo: { width: 38, height: 38, borderRadius: 10, background: "linear-gradient(135deg,#3b82f6,#8b5cf6)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 18, fontFamily: "Georgia,serif" },
  brandName: { color: "#f1f5f9", fontWeight: 700, fontSize: 17 },
  brandSub: { color: "#64748b", fontSize: 11 },
  nav: { display: "flex", gap: 2, flexWrap: "wrap" },
  navBtn: { display: "flex", alignItems: "center", gap: 5, padding: "7px 12px", border: "none", borderRadius: 8, background: "transparent", color: "#94a3b8", cursor: "pointer", fontSize: 13, fontWeight: 500, fontFamily: "inherit" },
  navAct: { background: "rgba(59,130,246,0.15)", color: "#60a5fa", fontWeight: 600 },
  main: { maxWidth: 1300, margin: "0 auto", padding: "20px 16px" },
  title: { fontSize: 22, fontWeight: 700, marginBottom: 16, color: "#0f172a" },
  titleRow: { display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 },
  grid4: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 12, marginBottom: 20 },
  card: { background: "#fff", borderRadius: 12, padding: 18, boxShadow: "0 1px 3px rgba(0,0,0,0.05)", border: "1px solid #e2e8f0", marginBottom: 16 },
  cardTitle: { fontSize: 15, fontWeight: 700, marginBottom: 10, color: "#334155" },
  summRow: { display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: 15 },
  schedBadge: { padding: "8px 12px", borderRadius: 8, background: "#f0fdf4", border: "1px solid #bbf7d0", fontSize: 13, textAlign: "center" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { padding: "8px 10px", textAlign: "left", fontWeight: 600, fontSize: 11, color: "#64748b", borderBottom: "2px solid #e2e8f0", whiteSpace: "nowrap", textTransform: "uppercase", letterSpacing: "0.04em" },
  td: { padding: "8px 10px", borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap", fontSize: 13 },
  tdM: { padding: "8px 10px", borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontSize: 12 },
  badge: { background: "#eff6ff", color: "#2563eb", padding: "2px 8px", borderRadius: 6, fontWeight: 600, fontSize: 11, fontFamily: "monospace" },
  posBadge: { background: "#f0fdf4", color: "#15803d", padding: "2px 8px", borderRadius: 6, fontSize: 12 },
  btnPrimary: { padding: "9px 16px", background: "#3b82f6", color: "#fff", border: "none", borderRadius: 8, fontWeight: 600, cursor: "pointer", fontSize: 13, fontFamily: "inherit", whiteSpace: "nowrap" },
  btnSec: { padding: "9px 16px", background: "#f1f5f9", color: "#475569", border: "1px solid #e2e8f0", borderRadius: 8, fontWeight: 500, cursor: "pointer", fontSize: 13, fontFamily: "inherit", whiteSpace: "nowrap" },
  iconBtn: { background: "none", border: "none", cursor: "pointer", padding: 4, fontSize: 15, opacity: 0.7 },
  fileLabel: { display: "inline-flex", alignItems: "center", gap: 6, padding: "10px 18px", background: "#eff6ff", color: "#2563eb", borderRadius: 8, fontWeight: 600, fontSize: 13, cursor: "pointer", border: "2px dashed #93c5fd", fontFamily: "inherit" },
  formGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 10 },
  label: { fontSize: 11, fontWeight: 600, color: "#475569", textTransform: "uppercase", letterSpacing: "0.04em" },
  input: { padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 13, color: "#1e293b", background: "#fff", outline: "none", width: "100%", boxSizing: "border-box", fontFamily: "inherit" },
  overlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 999, padding: 16 },
  modal: { background: "#fff", borderRadius: 16, padding: 24, width: "100%", maxWidth: 480, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" },
};