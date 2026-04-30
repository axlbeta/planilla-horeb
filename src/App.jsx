import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

const LOGO = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAA8ADwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDwEksfagEqcZzQQVPFKATyeK/bj4gGNNHWjvTj0oAXApoGDQuQe9ONACE0daDg0ox7UCEFLTAcV63+y74U0Hxn431TRfENlHdWraPM6bmKmJw8YDqw6MMnmufFYmOGoyqy2RrSpOpNQXU8k5JpV64xX1vrXwf8Px+H/CyofC0FhYaNNc65rS2nnrOqCLEibSNxPzncTwM9TisNvhJp0XiG+ubmXwpD4Sh0pdYh1MafIWaFydqlDJxgLknuCOMnA8mHEGGkr69fzt977bnW8vqI+ZhleO1IQcFm6175e/s9Pp/htNTuvEtqNZW1S/k094gsBQnJiWQtkuAD2wenvWn8b/gp4Xhm8Tap4K1mK2vNItY9Qu9C8o7YYWUksjk8Z2swXkduOK3WeYOU1BSvfrZ26L8b77eZH1GsottbHzfwMUjcGl/DNBBJ6V65xgB3rqfhv441nwDrVxq+hpaPcT2rWri5iLrsYgnABHPyiuYHPSk6jioqUoVYOE1dMcZOL5luei6B8ZvHmjT6Q1lqFsttpWn/ANnwWbW4MDw8Z3rnLMdq/NkdOMDOXaz8aPG+qnX11GawnTXLFbCZPs5CwQLvwsQDfL99iSc5/CvNwCDQQetc39n4Xm5vZq/p53/M1+sVduZnbeJviZ4l8S+D7Xw1raaZexWqxrFeSWam8VU+6vm9eOmcZOTnqa0vEHxq8f674ObwvqGp27WksSw3E6W4W4uIx0V37j1wAT36mvOM5HvRn1qvqOG09xaO602YvrFXX3nqLSUAinAGuoxFA4xTR0pVOVpq0ALQelFIehoAQnBpTyPegcgUD74FAx6qAPejI9aD0puaCdz/2Q==";

// ─── Supabase DB ───
const db = {
  async getEmployees() {
    const { data, error } = await supabase.from("employees").select("*").eq("active", true).order("id");
    if (error) { console.error("getEmployees:", error); return []; }
    return data.map(e => ({ id: String(e.id), name: e.name, position: e.position, salary: parseFloat(e.salary), empType: e.emp_type || "weekly" }));
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
    return data.map(e => ({ id: e.id, employeeId: String(e.employee_id), date: String(e.date).slice(0, 10), checkIn: e.check_in, lunchOut: e.lunch_out, lunchIn: e.lunch_in, checkOut: e.check_out, punches: e.punches, name: "" }));
  },
  async insertClockEntries(entries) {
    const rows = entries.map(e => ({ id: e.id, employee_id: e.employeeId, date: e.date, check_in: e.checkIn, lunch_out: e.lunchOut, lunch_in: e.lunchIn, check_out: e.checkOut, punches: e.punches }));
    const { error } = await supabase.from("clock_entries").upsert(rows, { onConflict: "id" });
    if (error) console.error("insertClockEntries:", error);
  },
  async updateClockEntry(entry) {
    const { error } = await supabase.from("clock_entries").update({
      check_in: entry.checkIn, lunch_out: entry.lunchOut, lunch_in: entry.lunchIn, check_out: entry.checkOut
    }).eq("id", entry.id);
    if (error) {
      console.error("updateClockEntry:", error);
      // Fallback: delete and re-insert
      await supabase.from("clock_entries").delete().eq("id", entry.id);
      const { error: e2 } = await supabase.from("clock_entries").insert({
        id: entry.id, employee_id: entry.employeeId, date: entry.date,
        check_in: entry.checkIn, lunch_out: entry.lunchOut, lunch_in: entry.lunchIn, check_out: entry.checkOut,
        punches: 4
      });
      if (e2) console.error("updateClockEntry fallback:", e2);
    }
  },
  async deleteClockEntry(id) { await supabase.from("clock_entries").delete().eq("id", id); },
  async deleteAllClockEntries() { await supabase.from("clock_entries").delete().neq("id", ""); },

  // Holidays
  async getHolidays() {
    const { data, error } = await supabase.from("holidays").select("*").order("date");
    if (error) { console.error("getHolidays:", error); return []; }
    return data.map(h => ({ id: h.id, date: String(h.date).slice(0, 10), name: h.name }));
  },
  async addHoliday(holiday) {
    const { error } = await supabase.from("holidays").insert({ date: holiday.date, name: holiday.name });
    if (error) console.error("addHoliday:", error);
  },
  async deleteHoliday(id) {
    await supabase.from("holidays").delete().eq("id", id);
  },
  async getPayrolls() {
    const { data, error } = await supabase.from("payrolls").select("*").order("id", { ascending: false });
    if (error) { console.error("getPayrolls:", error); return []; }
    return data;
  },
  async getPayrollRows(payrollId) {
    const { data, error } = await supabase.from("payroll_rows").select("*").eq("payroll_id", payrollId);
    if (error) return [];
    return data.map(r => ({ employeeId: r.employee_id, name: r.name, position: r.position, salary: +r.salary, daily: +r.daily, hourly: +r.hourly, days: r.days, effectiveHrs: +r.effective_hrs, baseSalary: +r.base_salary, ot: { 0.25: +r.ot_25, 0.5: +r.ot_50, 0.75: +r.ot_75, 1.0: +r.ot_100 }, otPay: +r.ot_pay, fuel: +r.fuel, vacation: +r.vacation, incapacity: +r.incapacity, advance: +r.advance, dec4: +r.dec4, dec3: +r.dec3, otherDed: +r.other_ded, totalEarned: +r.total_earned, totalDeductions: +r.total_deductions, netPay: +r.net_pay }));
  },
  async savePayroll(payroll) {
    const { data: existing } = await supabase.from("payrolls").select("id").eq("period", payroll.period).maybeSingle();
    let payrollId;
    if (existing) {
      payrollId = existing.id;
      await supabase.from("payroll_rows").delete().eq("payroll_id", payrollId);
      await supabase.from("payrolls").update({ week_num: payroll.weekNum, date_from: payroll.from, date_to: payroll.to, total_earned: payroll.rows.reduce((s, r) => s + r.totalEarned, 0), total_deductions: payroll.rows.reduce((s, r) => s + r.totalDeductions, 0), total_net: payroll.rows.reduce((s, r) => s + r.netPay, 0) }).eq("id", payrollId);
    } else {
      const { data, error } = await supabase.from("payrolls").insert({ period: payroll.period, week_num: payroll.weekNum, date_from: payroll.from, date_to: payroll.to, total_earned: payroll.rows.reduce((s, r) => s + r.totalEarned, 0), total_deductions: payroll.rows.reduce((s, r) => s + r.totalDeductions, 0), total_net: payroll.rows.reduce((s, r) => s + r.netPay, 0) }).select("id").single();
      if (error) { console.error("savePayroll:", error); return; }
      payrollId = data.id;
    }
    const rows = payroll.rows.map(r => ({ payroll_id: payrollId, employee_id: r.employeeId, name: r.name, position: r.position, salary: r.salary, daily: r.daily, hourly: r.hourly, days: r.days, effective_hrs: r.effectiveHrs, base_salary: r.baseSalary, ot_25: r.ot[0.25], ot_50: r.ot[0.5], ot_75: r.ot[0.75], ot_100: r.ot[1.0], ot_pay: r.otPay, fuel: r.fuel, vacation: r.vacation, incapacity: r.incapacity, advance: r.advance, dec4: r.dec4, dec3: r.dec3, other_ded: r.otherDed, total_earned: r.totalEarned, total_deductions: r.totalDeductions, net_pay: r.netPay }));
    await supabase.from("payroll_rows").insert(rows);
  },
  async deletePayroll(id) { await supabase.from("payroll_rows").delete().eq("payroll_id", id); await supabase.from("payrolls").delete().eq("id", id); },
};

// ─── Helpers ───
function getScheduledHours(dow) { return dow >= 1 && dow <= 4 ? 9 : dow === 5 ? 8 : 0; }

// ─── IHSS & RAP Honduras 2025/2026 ───
const IHSS = {
  EM_TECHO: 11903.13,    // Techo Enfermedad y Maternidad
  EM_TASA: 0.025,         // 2.5% trabajador
  IVM_TECHO: 11903.13,   // Techo Invalidez, Vejez y Muerte
  IVM_TASA: 0.025,        // 2.5% trabajador
};
const RAP_TASA = 0.015; // 1.5% por fondo (FEO3 y FIO3)
// RAP employee: FEO3 = FIO3 = 1.5% × (salario - techo IHSS)
// RL (patrono) = 4% del salario (no se descuenta al empleado)
function calcRAP_monthly(salary) {
  const excedente = Math.max(0, salary - IHSS.IVM_TECHO);
  const feo3 = excedente * RAP_TASA;
  const fio3 = feo3;
  const rl = salary * 0.04;
  return { feo3, fio3, employeeTotal: feo3 + fio3, rl, grandTotal: rl + feo3 + fio3 };
}
function calcIHSS_monthly(salary) {
  const baseEM = Math.min(salary, IHSS.EM_TECHO);
  const baseIVM = Math.min(salary, IHSS.IVM_TECHO);
  const em = baseEM * IHSS.EM_TASA;
  const ivm = baseIVM * IHSS.IVM_TASA;
  return { em, ivm, total: em + ivm };
}
function calcIHSS_biweekly(salary) {
  const m = calcIHSS_monthly(salary);
  return { em: m.em / 2, ivm: m.ivm / 2, total: m.total / 2 };
}

// Determine which week of the month a date range falls in (1-5)
function getWeekOfMonth(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  const day = d.getDate();
  if (day <= 7) return 1;
  if (day <= 14) return 2;
  if (day <= 21) return 3;
  if (day <= 28) return 4;
  return 5;
}
// IHSS: applies when the period contains the last day of the month
// (28 Feb, 30 Apr/Jun/Sep/Nov, 31 Jan/Mar/May/Jul/Aug/Oct/Dec)
function isLastWeekOfMonth(fromDate, toDate) {
  const from = new Date(fromDate + "T12:00:00");
  const to = new Date(toDate + "T12:00:00");
  // Check each day in the period — if any day is the last day of its month, IHSS applies
  for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
    const lastDayOfMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    if (d.getDate() === lastDayOfMonth) return true;
  }
  return false;
}
// RAP: applies on the second week of the month (days 8-14)
function isSecondWeekOfMonth(fromDate) {
  const wk = getWeekOfMonth(fromDate);
  return wk === 2;
}
function formatL(n) { if (n == null || isNaN(n)) return "L. 0.00"; return "L. " + Number(n).toLocaleString("es-HN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fN(n) { return (!n || isNaN(n) || n === 0) ? "" : Number(n).toLocaleString("es-HN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtTime(t) { return t ? new Date(t).toLocaleTimeString("es-HN", { hour: "2-digit", minute: "2-digit" }) : "—"; }
function fmtDate(d) { return new Date(d).toLocaleDateString("es-HN", { weekday: "short", day: "2-digit", month: "short" }); }

function parseClockCSV(csvText) {
  // Normalize: replace semicolons with commas, handle CRLF
  const text = csvText.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  
  // Detect separator (semicolon or comma) from header
  const sep = lines[0].includes(";") ? ";" : ",";
  
  const records = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(sep);
    if (cols.length < 10) continue;
    const time = cols[0].trim(), userId = cols[1].trim(), evento = cols[9].trim();
    if (!userId || evento === "Acceso denegado" || evento === "Dispositivo inicializado") continue;
    
    // Try multiple date formats:
    // Format 1: MM/DD/YYYY HH:MM:SS (original)
    // Format 2: D/M/YYYY HH:MM (no leading zeros, no seconds)
    // Format 3: DD/MM/YYYY HH:MM:SS
    let dt = null, dateStr = null;
    
    // Try: DD/M/YYYY HH:MM or D/M/YYYY HH:MM (no seconds)
    let m = time.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      const [, p1, p2, yyyy, hh, mi, ss] = m;
      const sec = ss || "00";
      // Determine if DD/MM or MM/DD
      // If p1 > 12, it must be DD/MM
      // If p2 > 12, it must be MM/DD  
      // If both <= 12, check context (Honduras uses DD/MM)
      let dd, mm;
      if (parseInt(p1) > 12) { dd = p1; mm = p2; }
      else if (parseInt(p2) > 12) { mm = p1; dd = p2; }
      else { dd = p1; mm = p2; } // Default: DD/MM for Honduras
      
      dd = dd.padStart(2, "0");
      mm = mm.padStart(2, "0");
      dt = new Date(+yyyy, +mm - 1, +dd, +hh, +mi, +sec);
      dateStr = `${yyyy}-${mm}-${dd}`;
    }
    
    if (!dt || !dateStr) continue;
    records.push({ userId, dateStr, dt });
  }
  return groupPunches(records);
}
function parseClockXLS(data) {
  try { const wb = XLSX.read(data, { type: "array" }); const ws = wb.Sheets[wb.SheetNames[0]]; const json = XLSX.utils.sheet_to_json(ws, { header: 1 }); if (json.length < 2) return []; const records = [];
  for (let i = 1; i < json.length; i++) { const row = json[i]; if (!row || row.length < 10) continue; const time = String(row[0] || "").trim(), userId = String(row[1] || "").trim(), evento = String(row[9] || "").trim(); if (!userId || evento === "Acceso denegado" || evento === "Dispositivo inicializado") continue;
    let m = time.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) continue;
    const [, p1, p2, yyyy, hh, mi, ss] = m;
    let dd, mm;
    if (parseInt(p1) > 12) { dd = p1; mm = p2; } else if (parseInt(p2) > 12) { mm = p1; dd = p2; } else { dd = p1; mm = p2; }
    dd = String(dd).padStart(2, "0"); mm = String(mm).padStart(2, "0");
    records.push({ userId, dateStr: `${yyyy}-${mm}-${dd}`, dt: new Date(+yyyy, +mm - 1, +dd, +hh, +mi, +(ss||0)) }); }
  return groupPunches(records); } catch (e) { return []; }
}
function groupPunches(records) {
  const groups = {}; records.forEach((r) => { const key = `${r.userId}_${r.dateStr}`; if (!groups[key]) groups[key] = { userId: r.userId, date: r.dateStr, punches: [] }; groups[key].punches.push(r.dt); });
  const entries = []; Object.values(groups).forEach((g) => { g.punches.sort((a, b) => a - b); const dd = [g.punches[0]]; for (let i = 1; i < g.punches.length; i++) if (g.punches[i] - g.punches[i - 1] > 120000) dd.push(g.punches[i]);
  let e = { id: `${g.userId}_${g.date}`, employeeId: g.userId, date: g.date, checkIn: null, lunchOut: null, lunchIn: null, checkOut: null, punches: dd.length };
  if (dd.length >= 4) { e.checkIn = dd[0].toISOString(); e.lunchOut = dd[1].toISOString(); e.lunchIn = dd[2].toISOString(); e.checkOut = dd[dd.length - 1].toISOString(); } else if (dd.length === 3) { e.checkIn = dd[0].toISOString(); e.lunchOut = dd[1].toISOString(); e.checkOut = dd[2].toISOString(); } else if (dd.length === 2) { e.checkIn = dd[0].toISOString(); e.checkOut = dd[1].toISOString(); } else if (dd.length === 1) { e.checkIn = dd[0].toISOString(); }
  entries.push(e); }); return entries.sort((a, b) => a.date.localeCompare(b.date));
}
function calcDayHours(entry) {
  if (!entry.checkIn || !entry.checkOut) return 0;
  let cin = new Date(entry.checkIn);
  const cout = new Date(entry.checkOut);
  
  // Grace period at entry: only if arriving between 7:50am and 8:00am, count as 8:00am
  // If arriving earlier (7:00am, 7:30am etc), those hours DO count (could be overtime)
  const scheduled8am = new Date(cin);
  scheduled8am.setHours(8, 0, 0, 0);
  const grace10before = new Date(cin);
  grace10before.setHours(7, 50, 0, 0);
  if (cin >= grace10before && cin < scheduled8am) cin = scheduled8am;
  
  let ms = cout - cin;
  if (entry.lunchOut && entry.lunchIn) ms -= (new Date(entry.lunchIn) - new Date(entry.lunchOut));
  else if (entry.lunchOut) ms -= 3600000;
  else if (ms > 5 * 3600000) ms -= 3600000;
  return Math.max(0, ms / 3600000);
}
function calcOvertime(entries) {
  const GRACE = 10/60; // 10 minutes grace period in hours
  const byEmp = {}; entries.forEach((e) => { if (!byEmp[e.employeeId]) byEmp[e.employeeId] = []; byEmp[e.employeeId].push(e); }); const results = {};
  Object.entries(byEmp).forEach(([empId, days]) => { let totalEffective = 0, regularDays = 0, ot = { 0.25: 0, 0.5: 0, 0.75: 0, 1.0: 0 };
  days.forEach((day) => { if (!day.checkIn || !day.checkOut) return; const dow = new Date(day.checkIn).getDay(), hrs = calcDayHours(day), scheduled = getScheduledHours(dow);
    // Sunday = 100%
    if (dow === 0) { ot[1.0] += hrs; if (hrs > 0) regularDays++; totalEffective += hrs; return; }
    // Saturday = 25%
    if (dow === 6) { ot[0.25] += hrs; if (hrs > 0) regularDays++; totalEffective += hrs; return; }
    regularDays++; totalEffective += hrs; if (hrs <= scheduled + GRACE) return;
    const extra = hrs - scheduled - GRACE;
    const cH = new Date(day.checkOut).getHours() + new Date(day.checkOut).getMinutes() / 60;
    // 6pm-7pm (18-19) = 25%, 7:01pm-9pm (19-21) = 50%, 9pm-5am (21-5) = 75%
    if (cH <= 19) { ot[0.25] += extra; }
    else if (cH <= 21) {
      const hrs25 = Math.max(0, Math.min(extra, 1)); // up to 1hr at 25% (6-7pm)
      ot[0.25] += hrs25;
      ot[0.5] += extra - hrs25;
    } else {
      const hrs25 = Math.max(0, Math.min(extra, 1)); // 6-7pm
      const hrs50 = Math.max(0, Math.min(extra - hrs25, 2)); // 7-9pm
      const hrs75 = Math.max(0, extra - hrs25 - hrs50); // 9pm+
      ot[0.25] += hrs25; ot[0.5] += hrs50; ot[0.75] += hrs75;
    }
  });
  results[empId] = { totalEffective, regularDays, ot }; }); return results;
}
function printPayroll(payroll) {
  const rows = payroll.rows;
  const totals = { salary: rows.reduce((s,r)=>s+(r.salary||0),0), base: rows.reduce((s,r)=>s+r.baseSalary,0), ot: rows.reduce((s,r)=>s+r.otPay,0), ihss: rows.reduce((s,r)=>s+(r.ihssTotal||0),0), rap: rows.reduce((s,r)=>s+(r.rap||0),0), earned: rows.reduce((s,r)=>s+r.totalEarned,0), ded: rows.reduce((s,r)=>s+r.totalDeductions,0), net: rows.reduce((s,r)=>s+r.netPay,0) };
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Planilla ${payroll.period}</title>
<style>@page{size:landscape;margin:10mm}*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:8pt;color:#000}.header{text-align:center;margin-bottom:10px}.header img{height:36px;margin-bottom:4px}.header h2{font-size:10pt;font-weight:normal;color:#1a3a6b}.header .period{font-size:9pt;margin-top:3px;font-weight:bold;color:#0a2351}table{width:100%;border-collapse:collapse;margin-top:6px}th{background:#0a2351;color:#fff;padding:4px 3px;font-size:6.5pt;text-transform:uppercase;border:1px solid #0d2d6b;text-align:center}td{padding:3px;border:1px solid #c8d6e5;font-size:7.5pt}.r{text-align:right;font-family:'Courier New',monospace;font-size:7.5pt}.c{text-align:center}.name{font-weight:600;white-space:nowrap}.total-row{background:#e8eef6;font-weight:700}.total-row td{border-top:2px solid #0a2351;padding:5px 3px}.net{color:#0a6847;font-weight:700}.ot-val{color:#b91c1c}.ihss{color:#6d28d9}.signatures{margin-top:30px;display:flex;justify-content:space-between}.sig-box{text-align:center;width:180px}.sig-line{border-top:1px solid #000;margin-top:40px;padding-top:3px;font-size:8pt}@media print{.no-print{display:none!important}}.no-print{position:fixed;top:10px;right:10px;z-index:999}.print-btn{padding:8px 20px;background:#0a2351;color:#fff;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;margin-right:6px}.close-btn{padding:8px 20px;background:#64748b;color:#fff;border:none;border-radius:8px;font-size:12px;cursor:pointer}</style></head><body>
<div class="no-print"><button class="print-btn" onclick="window.print()">🖨️ Imprimir</button><button class="close-btn" onclick="window.close()">✕ Cerrar</button></div>
<div class="header"><img src="${LOGO}" alt="Horeb"/><h2>Planilla de Empleados</h2><div class="period">${payroll.period}</div></div>
<table><thead><tr>
<th rowspan="2">Cód.</th><th rowspan="2">Nombre</th><th rowspan="2">Posición</th><th rowspan="2">Sal.Mensual</th><th rowspan="2">Sal.Diario</th><th rowspan="2">Días</th><th rowspan="2">Salario</th>
<th colspan="4">Horas Extras</th><th rowspan="2">Sal/Hr</th><th rowspan="2">Total Extras</th>
<th rowspan="2">Comb.</th><th rowspan="2">Vac.</th><th rowspan="2">Incap.</th><th rowspan="2">Adelanto</th><th rowspan="2">Dec.4to</th><th rowspan="2">Dec.3ro</th>
<th rowspan="2">IHSS</th><th rowspan="2">RAP</th>
<th rowspan="2">Devengado</th><th rowspan="2">Tot.Ded.</th><th rowspan="2">Neto</th>
</tr><tr><th>25%</th><th>50%</th><th>75%</th><th>100%</th></tr></thead><tbody>
${rows.map(r=>`<tr>
<td class="c">${r.employeeId}</td><td class="name">${r.name}</td><td>${r.position}</td>
<td class="r">${fN(r.salary)}</td><td class="r">${fN(r.daily)}</td><td class="c">${r.days}</td><td class="r">${fN(r.baseSalary)}</td>
<td class="r${r.ot[0.25]>0?' ot-val':''}">${r.ot[0.25]>0?r.ot[0.25].toFixed(1):''}</td>
<td class="r${r.ot[0.5]>0?' ot-val':''}">${r.ot[0.5]>0?r.ot[0.5].toFixed(1):''}</td>
<td class="r${r.ot[0.75]>0?' ot-val':''}">${r.ot[0.75]>0?r.ot[0.75].toFixed(1):''}</td>
<td class="r${r.ot[1.0]>0?' ot-val':''}">${r.ot[1.0]>0?r.ot[1.0].toFixed(1):''}</td>
<td class="r">${fN(r.hourly)}</td><td class="r ot-val">${fN(r.otPay)}</td>
<td class="r">${fN(r.fuel)}</td><td class="r">${fN(r.vacation)}</td><td class="r">${fN(r.incapacity)}</td>
<td class="r">${fN(r.advance)}</td><td class="r">${fN(r.dec4)}</td><td class="r">${fN(r.dec3)}</td>
<td class="r ihss">${fN(r.ihssTotal)}</td><td class="r ihss">${fN(r.rap)}</td>
<td class="r" style="font-weight:600">${fN(r.totalEarned)}</td>
<td class="r">${fN(r.totalDeductions)}</td>
<td class="r net">${fN(r.netPay)}</td>
</tr>`).join("")}
<tr class="total-row">
<td colspan="3" style="text-align:right">TOTALES</td><td class="r">${fN(totals.salary)}</td><td></td><td></td><td class="r">${fN(totals.base)}</td>
<td colspan="4"></td><td></td><td class="r ot-val">${fN(totals.ot)}</td>
<td colspan="6"></td>
<td class="r ihss">${fN(totals.ihss)}</td><td class="r ihss">${fN(totals.rap)}</td>
<td class="r" style="font-weight:700">${fN(totals.earned)}</td><td class="r">${fN(totals.ded)}</td>
<td class="r net" style="font-size:9pt">${fN(totals.net)}</td>
</tr></tbody></table>
<div class="signatures"><div class="sig-box"><div class="sig-line">Elaborado por</div></div><div class="sig-box"><div class="sig-line">Revisado por</div></div><div class="sig-box"><div class="sig-line">Autorizado por</div></div></div></body></html>`;
  const w = window.open("", "_blank"); if (w) { w.document.write(html); w.document.close(); }
}

// ═══════════════════
const TABS = [
  { id: "dash", label: "Dashboard", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg> },
  { id: "emp", label: "Empleados", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg> },
  { id: "clock", label: "Reloj", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> },
  { id: "pay", label: "Planilla", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> },
  { id: "conf", label: "Confidencial", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg> },
  { id: "hist", label: "Historial", icon: <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="1.8"><path d="M12 8v4l3 3"/><circle cx="12" cy="12" r="10"/></svg> },
];

export default function App() {
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [tab, setTab] = useState("dash");
  const [employees, setEmployees] = useState([]);
  const [clockEntries, setClockEntries] = useState([]);
  const [payrolls, setPayrolls] = useState([]);
  const [holidays, setHolidays] = useState([]);
  const [loading, setLoading] = useState(false);

  // Check auth state on mount
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user || null);
      setAuthLoading(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user || null);
    });
    return () => subscription.unsubscribe();
  }, []);

  const refresh = useCallback(async () => { const [e, c, p, h] = await Promise.all([db.getEmployees(), db.getClockEntries(), db.getPayrolls(), db.getHolidays()]); setEmployees(e); setClockEntries(c); setPayrolls(p); setHolidays(h); }, []);

  useEffect(() => { if (user) { setLoading(true); refresh().then(() => setLoading(false)); } }, [user, refresh]);

  const handleLogout = async () => { await supabase.auth.signOut(); setUser(null); };

  // Global styles
  const globalStyles = `
    @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Source+Sans+3:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
    @keyframes spin{to{transform:rotate(360deg)}}
    @keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
    @keyframes slideUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
    input:focus,select:focus{border-color:#1a5ab8!important;box-shadow:0 0 0 3px rgba(26,90,184,0.1)!important;outline:none}
    button{transition:all 0.2s ease}
    button:hover{transform:translateY(-1px)}
    button:active{transform:translateY(0)}
    tr:hover td{background:rgba(10,35,81,0.02)!important}
    *{box-sizing:border-box;margin:0;padding:0}
    ::selection{background:#1a5ab8;color:#fff}
    ::-webkit-scrollbar{width:6px;height:6px}
    ::-webkit-scrollbar-track{background:#f0f3f8}
    ::-webkit-scrollbar-thumb{background:#b0bec5;border-radius:3px}
  `;

  // Auth loading
  if (authLoading) return (
    <div style={{ display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100vh",background:"linear-gradient(135deg,#0a2351 0%,#163a72 50%,#1a4a8a 100%)" }}>
      <style>{globalStyles}</style>
      <img src={LOGO} alt="Horeb" style={{ height: 50, marginBottom: 20, borderRadius: 8 }}/>
      <div style={{ width:36,height:36,border:"3px solid rgba(255,255,255,0.2)",borderTopColor:"#c9a227",borderRadius:"50%",animation:"spin 0.8s linear infinite" }}/>
    </div>
  );

  // Login screen
  if (!user) return <LoginScreen onLogin={setUser} globalStyles={globalStyles} />;

  // Data loading
  if (loading) return (
    <div style={{ display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100vh",background:"linear-gradient(135deg,#0a2351 0%,#163a72 50%,#1a4a8a 100%)" }}>
      <style>{globalStyles}</style>
      <img src={LOGO} alt="Horeb" style={{ height: 50, marginBottom: 20, borderRadius: 8 }}/>
      <div style={{ width:36,height:36,border:"3px solid rgba(255,255,255,0.2)",borderTopColor:"#c9a227",borderRadius:"50%",animation:"spin 0.8s linear infinite" }}/>
      <p style={{ color:"rgba(255,255,255,0.6)",marginTop:14,fontSize:13,fontFamily:"'Cormorant Garamond',Georgia,serif",letterSpacing:"0.1em",textTransform:"uppercase" }}>Cargando sistema</p>
    </div>
  );

  const TabComp = { dash: DashboardTab, emp: EmployeesTab, clock: ClockTab, pay: PayrollTab, conf: ConfidentialTab, hist: HistoryTab }[tab];

  return (
    <div style={S.app}>
      <style>{globalStyles}</style>
      <header style={S.header}>
        <div style={S.brand}>
          <img src={LOGO} alt="Horeb" style={S.logoImg}/>
          <div style={S.brandDivider}/>
          <div>
            <div style={S.brandSub}>Sistema de Planilla</div>
          </div>
        </div>
        <nav style={S.nav}>
          {TABS.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ ...S.navBtn, ...(tab === t.id ? S.navAct : {}) }}>
              {t.icon}<span>{t.label}</span>
            </button>
          ))}
          <div style={S.brandDivider}/>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:11,color:"rgba(255,255,255,0.4)",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.email}</span>
            <button onClick={handleLogout} style={{...S.navBtn,color:"#f87171",fontSize:12,padding:"6px 10px"}} title="Cerrar sesión">
              <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
              Salir
            </button>
          </div>
        </nav>
      </header>

      <main style={S.main}><div style={{ animation: "fadeIn 0.35s ease" }}>
        <TabComp employees={employees} refresh={refresh} clockEntries={clockEntries} payrolls={payrolls} holidays={holidays} />
      </div></main>

      <footer style={S.footer}>
        <span>Impresos Horeb © {new Date().getFullYear()}</span>
        <span style={{ color: "#94a3b8" }}>Sistema de Planilla v3.0</span>
      </footer>
    </div>
  );
}

// ═══ LOGIN SCREEN ═══
function LoginScreen({ onLogin, globalStyles }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [isRegister, setIsRegister] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!email || !password) return setError("Ingresa correo y contraseña.");
    setLoading(true);
    setError("");

    if (isRegister) {
      const { data, error: err } = await supabase.auth.signUp({ email, password });
      if (err) { setError(err.message); setLoading(false); return; }
      if (data.user) {
        setError("");
        setIsRegister(false);
        alert("✅ Cuenta creada. Ahora inicia sesión.");
      }
    } else {
      const { data, error: err } = await supabase.auth.signInWithPassword({ email, password });
      if (err) { setError(err.message === "Invalid login credentials" ? "Correo o contraseña incorrectos." : err.message); setLoading(false); return; }
      if (data.user) onLogin(data.user);
    }
    setLoading(false);
  };

  return (
    <div style={{
      display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",
      background:"linear-gradient(135deg,#0a2351 0%,#0d2d6b 40%,#1a4a8a 100%)",
      fontFamily:"'Source Sans 3','Segoe UI',sans-serif",padding:20,
    }}>
      <style>{globalStyles}</style>
      <div style={{
        background:"#fff",borderRadius:20,padding:0,width:"100%",maxWidth:420,
        boxShadow:"0 24px 80px rgba(0,0,0,0.4)",overflow:"hidden",animation:"slideUp 0.5s ease",
      }}>
        {/* Header */}
        <div style={{
          background:"linear-gradient(135deg,#0a2351,#163a72)",padding:"32px 40px 28px",textAlign:"center",
          borderBottom:"3px solid #c9a227",
        }}>
          <img src={LOGO} alt="Horeb" style={{height:45,marginBottom:12,borderRadius:6}}/>
          <div style={{color:"rgba(255,255,255,0.5)",fontSize:12,letterSpacing:"0.15em",textTransform:"uppercase",fontFamily:"'Cormorant Garamond',serif"}}>
            Sistema de Planilla
          </div>
        </div>

        {/* Form */}
        <div style={{padding:"32px 40px 36px"}}>
          <h2 style={{fontSize:20,fontWeight:700,color:"#0a2351",marginBottom:4,fontFamily:"'Cormorant Garamond',serif"}}>
            {isRegister ? "Crear Cuenta" : "Iniciar Sesión"}
          </h2>
          <p style={{fontSize:13,color:"#64748b",marginBottom:24}}>
            {isRegister ? "Registra un nuevo usuario autorizado." : "Ingresa tus credenciales para acceder."}
          </p>

          {error && (
            <div style={{padding:"10px 14px",background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,marginBottom:16,fontSize:13,color:"#dc2626",fontWeight:500}}>
              {error}
            </div>
          )}

          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.06em"}}>Correo Electrónico</label>
              <input
                type="email" value={email} onChange={e=>setEmail(e.target.value)}
                placeholder="usuario@empresa.com"
                style={{padding:"11px 14px",border:"1px solid #d0daea",borderRadius:10,fontSize:14,color:"#1e293b",background:"#f8fafc",width:"100%",fontFamily:"inherit"}}
              />
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.06em"}}>Contraseña</label>
              <input
                type="password" value={password} onChange={e=>setPassword(e.target.value)}
                placeholder="••••••••"
                onKeyDown={e=>e.key==="Enter"&&handleSubmit(e)}
                style={{padding:"11px 14px",border:"1px solid #d0daea",borderRadius:10,fontSize:14,color:"#1e293b",background:"#f8fafc",width:"100%",fontFamily:"inherit"}}
              />
            </div>
            <button
              onClick={handleSubmit} disabled={loading}
              style={{
                padding:"12px",background:"linear-gradient(135deg,#0a2351,#163a72)",color:"#fff",
                border:"none",borderRadius:10,fontWeight:700,fontSize:14,cursor:"pointer",
                fontFamily:"inherit",letterSpacing:"0.03em",marginTop:4,
                boxShadow:"0 4px 12px rgba(10,35,81,0.3)",
                opacity:loading?0.7:1,
              }}
            >
              {loading ? "Procesando..." : isRegister ? "Crear Cuenta" : "Ingresar"}
            </button>
          </div>

          <div style={{textAlign:"center",marginTop:20}}>
            <button
              onClick={()=>{setIsRegister(!isRegister);setError("")}}
              style={{background:"none",border:"none",color:"#1a5ab8",fontSize:13,cursor:"pointer",fontFamily:"inherit",fontWeight:500}}
            >
              {isRegister ? "← Ya tengo cuenta, iniciar sesión" : "Registrar nuevo usuario →"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══ DASHBOARD ═══
function DashboardTab({ employees, payrolls, holidays, refresh }) {
  const totalSalary = employees.reduce((s, e) => s + e.salary, 0);
  const last = payrolls.length > 0 ? payrolls[0] : null;
  const [lastRows, setLastRows] = useState(null);
  const [newHoliday, setNewHoliday] = useState({ date: "", name: "" });
  useEffect(() => { if (last) db.getPayrollRows(last.id).then(setLastRows); }, [last]);

  const addHoliday = async () => {
    if (!newHoliday.date || !newHoliday.name) return;
    await db.addHoliday(newHoliday);
    await refresh();
    setNewHoliday({ date: "", name: "" });
  };
  const removeHoliday = async (id) => { await db.deleteHoliday(id); await refresh(); };
  return (
    <div>
      <div style={S.pageHeader}><h2 style={S.title}>Panel de Control</h2><div style={S.goldLine}/></div>
      <div style={S.grid4}>
        <StatCard icon="👥" label="Empleados Activos" value={employees.length} accent="#1a5ab8"/>
        <StatCard icon="💰" label="Nómina Mensual" value={formatL(totalSalary)} accent="#0a6847"/>
        <StatCard icon="📋" label="Planillas Guardadas" value={payrolls.length} accent="#c9a227"/>
        <StatCard icon="📅" label="Última Planilla" value={last?.period || "—"} accent="#7c3aed"/>
      </div>
      <div style={S.card}>
        <h3 style={S.cardTitle}><span style={S.cardTitleIcon}>🕐</span> Horario Laboral</h3>
        <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10 }}>
          {["Lunes","Martes","Miércoles","Jueves"].map(d=><div key={d} style={S.schedCard}><div style={S.schedDay}>{d}</div><div style={S.schedTime}>8:00am — 6:00pm</div><div style={S.schedHrs}>9 hrs efectivas</div></div>)}
          <div style={{...S.schedCard,borderColor:"#c9a227",background:"linear-gradient(135deg,#fffbeb,#fef3c7)"}}><div style={{...S.schedDay,color:"#92400e"}}>Viernes</div><div style={S.schedTime}>8:00am — 5:00pm</div><div style={S.schedHrs}>8 hrs efectivas</div></div>
        </div>
        <div style={S.schedNote}>Almuerzo: 12:00 — 1:00pm · Sábado compensado de Lunes a Jueves · 44 horas semanales</div>
      </div>
      {last && lastRows && (
        <div style={{...S.card,background:"linear-gradient(135deg,#0a2351,#163a72)",color:"#fff",border:"none"}}>
          <div style={S.titleRow}>
            <h3 style={{...S.cardTitle,color:"rgba(255,255,255,0.7)"}}>Última Planilla: <span style={{color:"#c9a227"}}>{last.period}</span></h3>
            <button style={S.btnGold} onClick={()=>printPayroll({period:last.period,rows:lastRows})}>🖨️ Imprimir</button>
          </div>
          <div style={{display:"flex",gap:24,flexWrap:"wrap",marginTop:12}}>
            <div><div style={{fontSize:11,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:"0.08em"}}>Devengado</div><div style={{fontSize:22,fontWeight:700,color:"#fff",fontFamily:"'JetBrains Mono',monospace"}}>{formatL(+last.total_earned)}</div></div>
            <div><div style={{fontSize:11,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:"0.08em"}}>Deducciones</div><div style={{fontSize:22,fontWeight:700,color:"#f87171",fontFamily:"'JetBrains Mono',monospace"}}>{formatL(+last.total_deductions)}</div></div>
            <div><div style={{fontSize:11,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:"0.08em"}}>Neto a Pagar</div><div style={{fontSize:28,fontWeight:700,color:"#c9a227",fontFamily:"'JetBrains Mono',monospace"}}>{formatL(+last.total_net)}</div></div>
          </div>
        </div>
      )}

      {/* Holidays Management */}
      <div style={S.card}>
        <h3 style={S.cardTitle}><span style={S.cardTitleIcon}>🎌</span> Días Feriados</h3>
        <p style={{fontSize:12,color:"#64748b",marginBottom:12}}>Los feriados se pagan sin necesidad de marcación. No cuentan como falta.</p>
        <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div style={{display:"flex",flexDirection:"column",gap:4}}>
            <label style={S.label}>Fecha</label>
            <input style={{...S.input,width:160}} type="date" value={newHoliday.date} onChange={e=>setNewHoliday({...newHoliday,date:e.target.value})}/>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:4}}>
            <label style={S.label}>Nombre del Feriado</label>
            <input style={{...S.input,width:220}} value={newHoliday.name} onChange={e=>setNewHoliday({...newHoliday,name:e.target.value})} placeholder="Ej: Día de la Bandera"/>
          </div>
          <button style={S.btnPrimary} onClick={addHoliday}>+ Agregar</button>
        </div>
        {holidays.length===0?<p style={{color:"#94a3b8",fontStyle:"italic",fontSize:13}}>No hay feriados registrados.</p>:(
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
            {holidays.map(h=>(
              <div key={h.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 14px",background:"linear-gradient(135deg,#fffbeb,#fef3c7)",border:"1px solid #fde68a",borderRadius:8}}>
                <span style={{fontSize:18}}>🎌</span>
                <div>
                  <div style={{fontSize:13,fontWeight:700,color:"#92400e"}}>{h.name}</div>
                  <div style={{fontSize:11,color:"#a16207"}}>{fmtDate(h.date+"T12:00:00")}</div>
                </div>
                <button style={{background:"none",border:"none",cursor:"pointer",fontSize:14,color:"#dc2626",marginLeft:4}} onClick={()=>removeHoliday(h.id)}>✕</button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
function StatCard({icon,label,value,accent}){return<div style={{...S.card,borderTop:`3px solid ${accent}`,padding:"20px 18px"}}><div style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:24}}>{icon}</span><div><div style={{fontSize:11,color:"#64748b",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}</div><div style={{fontSize:20,fontWeight:700,color:"#0a2351",marginTop:2}}>{value}</div></div></div></div>}

// ═══ EMPLOYEES ═══
function EmployeesTab({ employees, refresh }) {
  const weeklyEmployees = employees.filter(e => (e.empType || "weekly") === "weekly");
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({ id:"",name:"",position:"",salary:"" });
  const [saving, setSaving] = useState(false);
  const openAdd=()=>{setForm({id:"",name:"",position:"",salary:""});setModal("new")};
  const openEdit=(e)=>{setForm({...e,salary:String(e.salary)});setModal(e.id)};
  const doSave=async()=>{const emp={...form,salary:parseFloat(form.salary)||0};if(!emp.id||!emp.name)return;setSaving(true);await db.upsertEmployee(emp);await refresh();setSaving(false);setModal(null)};
  const doDelete=async(id)=>{if(confirm("¿Eliminar este empleado?")) {await db.deleteEmployee(id);await refresh()}};
  return (
    <div>
      <div style={S.titleRow}><div style={S.pageHeader}><h2 style={S.title}>Empleados (Semanal)</h2><div style={S.goldLine}/></div><button style={S.btnPrimary} onClick={openAdd}>+ Nuevo Empleado</button></div>
      {modal&&(<div style={S.overlay}><div style={S.modal}>
        <h3 style={S.modalTitle}>{modal==="new"?"Nuevo Empleado":"Editar Empleado"}</h3>
        <div style={S.formGrid}>
          <Field label="Código" value={form.id} onChange={v=>setForm({...form,id:v})} ph="008" disabled={modal!=="new"}/>
          <Field label="Nombre Completo" value={form.name} onChange={v=>setForm({...form,name:v})} ph="Nombre del empleado"/>
          <Field label="Posición" value={form.position} onChange={v=>setForm({...form,position:v})} ph="Prensista"/>
          <Field label="Salario Mensual (L.)" value={form.salary} onChange={v=>setForm({...form,salary:v})} ph="0.00" type="number"/>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:24}}>
          <button style={S.btnSec} onClick={()=>setModal(null)}>Cancelar</button>
          <button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"Guardando...":"Guardar"}</button>
        </div>
      </div></div>)}
      <div style={S.card}><div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
        {["Cód","Nombre","Posición","Salario Mensual","Salario Diario","Salario/Hora",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=3&&i<=5?"right":"left"}}>{h}</th>)}
      </tr></thead><tbody>
        {weeklyEmployees.map(e=><tr key={e.id}><td style={S.td}><span style={S.badge}>{e.id}</span></td><td style={{...S.td,fontWeight:600,color:"#0a2351"}}>{e.name}</td><td style={S.td}><span style={S.posBadge}>{e.position}</span></td><td style={S.tdM}>{formatL(e.salary)}</td><td style={S.tdM}>{formatL(e.salary/30)}</td><td style={S.tdM}>{formatL(e.salary/30/8)}</td><td style={{...S.td,textAlign:"center"}}><button style={S.tblBtn} onClick={()=>openEdit(e)}>Editar</button><button style={{...S.tblBtn,color:"#dc2626",borderColor:"#fecaca"}} onClick={()=>doDelete(e.id)}>Eliminar</button></td></tr>)}
      </tbody></table></div></div>
    </div>
  );
}

// ═══ CLOCK ═══
function ClockTab({ employees, clockEntries, refresh }) {
  const [mode,setMode]=useState("import");const [importResult,setImportResult]=useState(null);const [importing,setImporting]=useState(false);
  const [mf,setMf]=useState({empId:"",date:new Date().toISOString().slice(0,10),cin:"08:00",lout:"12:00",lin:"13:00",cout:"18:00"});
  const [filterEmp,setFilterEmp]=useState("");const [filterMonth,setFilterMonth]=useState(new Date().toISOString().slice(0,7));
  const [editEntry,setEditEntry]=useState(null);
  const [editForm,setEditForm]=useState({checkIn:"",lunchOut:"",lunchIn:"",checkOut:""});
  const [editSaving,setEditSaving]=useState(false);

  const handleFile=(e)=>{const file=e.target.files[0];if(!file)return;setImporting(true);const isXLS=file.name.match(/\.xls[xm]?$/i);
    if(isXLS){const reader=new FileReader();reader.onload=async(evt)=>{let parsed=parseClockXLS(new Uint8Array(evt.target.result));if(parsed.length===0){const tr=new FileReader();tr.onload=async(e2)=>{let t=e2.target.result;if(t.includes("<table")||t.includes("<html")){const doc=new DOMParser().parseFromString(t,"text/html");const lines=[];doc.querySelectorAll("tr").forEach(r=>{lines.push(Array.from(r.querySelectorAll("td,th")).map(c=>c.textContent.trim()).join(","))});t=lines.join("\n")}await finishImport(parseClockCSV(t))};tr.readAsText(file,"UTF-8");return}await finishImport(parsed)};reader.readAsArrayBuffer(file)}
    else{const reader=new FileReader();reader.onload=async(evt)=>{let parsed=parseClockCSV(evt.target.result);if(parsed.length===0){const r2=new FileReader();r2.onload=async(e2)=>await finishImport(parseClockCSV(e2.target.result));r2.readAsText(file,"ISO-8859-1");return}await finishImport(parsed)};reader.readAsText(file,"UTF-8")}e.target.value=""};
  const finishImport=async(parsed)=>{if(parsed.length===0){setImportResult({error:"No se encontraron registros válidos."});setImporting(false);return}const existingIds=new Set(clockEntries.map(c=>c.id));const newE=parsed.filter(p=>!existingIds.has(p.id));if(newE.length>0)await db.insertClockEntries(newE);await refresh();setImportResult({added:newE.length,total:parsed.length,skipped:parsed.length-newE.length});setImporting(false)};
  const addManual=async()=>{if(!mf.empId)return;const entry={id:`${mf.empId}_${mf.date}_m${Date.now()}`,employeeId:mf.empId,date:mf.date,checkIn:`${mf.date}T${mf.cin}:00`,lunchOut:`${mf.date}T${mf.lout}:00`,lunchIn:`${mf.date}T${mf.lin}:00`,checkOut:`${mf.date}T${mf.cout}:00`,punches:4};await db.insertClockEntries([entry]);await refresh()};
  const removeEntry=async(id)=>{await db.deleteClockEntry(id);await refresh()};
  const clearAll=async()=>{if(confirm("¿Borrar TODAS las marcaciones?")) {await db.deleteAllClockEntries();await refresh()}};

  // ─── Edit functions ───
  const extractTime = (isoStr) => {
    if (!isoStr) return "";
    try { const d = new Date(isoStr); return `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`; }
    catch { return ""; }
  };
  const openEdit = (entry) => {
    setEditEntry(entry);
    setEditForm({
      checkIn: extractTime(entry.checkIn),
      lunchOut: extractTime(entry.lunchOut),
      lunchIn: extractTime(entry.lunchIn),
      checkOut: extractTime(entry.checkOut),
    });
  };
  const saveEdit = async () => {
    if (!editEntry) return;
    setEditSaving(true);
    const d = editEntry.date;
    const toISO = (time) => time ? new Date(`${d}T${time}:00`).toISOString() : null;
    const updated = {
      id: editEntry.id,
      employeeId: editEntry.employeeId,
      date: d,
      checkIn: toISO(editForm.checkIn),
      lunchOut: toISO(editForm.lunchOut),
      lunchIn: toISO(editForm.lunchIn),
      checkOut: toISO(editForm.checkOut),
    };
    await db.updateClockEntry(updated);
    await refresh();
    setEditSaving(false);
    setEditEntry(null);
  };

  // ─── Error detection ───
  const hasError = (entry) => {
    const dow = new Date(entry.date + "T12:00:00").getDay();
    if (entry.checkIn && !entry.checkOut && dow === 5) return "⚠️ VIERNES SIN SALIDA — Editar y poner 5:00pm";
    if (!entry.checkIn && !entry.checkOut) return "Sin marcaciones";
    if (entry.checkIn && !entry.checkOut) return "Falta hora de salida";
    if (!entry.checkIn && entry.checkOut) return "Falta hora de entrada";
    const hrs = calcDayHours(entry);
    if (hrs <= 0) return "0 horas registradas";
    if (hrs > 16) return "Más de 16 horas";
    if (!entry.lunchOut && !entry.lunchIn && hrs > 5) return "Sin marcación de almuerzo";
    if (entry.lunchOut && !entry.lunchIn) return "Falta regreso de almuerzo";
    if (!entry.lunchOut && entry.lunchIn) return "Falta salida de almuerzo";
    return null;
  };

  const filtered=clockEntries.filter(e=>(!filterEmp||e.employeeId===filterEmp)&&(!filterMonth||e.date.startsWith(filterMonth))).sort((a,b)=>b.date.localeCompare(a.date)||(a.employeeId||"").localeCompare(b.employeeId||""));
  const getN=id=>employees.find(e=>e.id===id)?.name||`ID ${id}`;
  const errorCount = filtered.filter(e => hasError(e)).length;

  return (
    <div>
      <div style={S.pageHeader}><h2 style={S.title}>Reloj Marcador</h2><div style={S.goldLine}/></div>

      {/* Edit Modal */}
      {editEntry && (
        <div style={S.overlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>Editar Marcación</h3>
          <div style={{background:"#f0f3f8",borderRadius:10,padding:14,marginBottom:18}}>
            <div style={{fontSize:14,fontWeight:700,color:"#0a2351"}}>{getN(editEntry.employeeId)}</div>
            <div style={{fontSize:13,color:"#475569",marginTop:2}}>{fmtDate(editEntry.date+"T12:00:00")} — {editEntry.date}</div>
            {hasError(editEntry) && <div style={{marginTop:8,padding:"6px 10px",background:"#fef2f2",border:"1px solid #fecaca",borderRadius:6,fontSize:12,color:"#dc2626",fontWeight:600}}>⚠️ {hasError(editEntry)}</div>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={S.label}>Hora Entrada</label>
              <input style={S.input} type="time" value={editForm.checkIn} onChange={e=>setEditForm({...editForm,checkIn:e.target.value})}/>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={S.label}>Sale Almuerzo</label>
              <input style={S.input} type="time" value={editForm.lunchOut} onChange={e=>setEditForm({...editForm,lunchOut:e.target.value})}/>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={S.label}>Regresa Almuerzo</label>
              <input style={S.input} type="time" value={editForm.lunchIn} onChange={e=>setEditForm({...editForm,lunchIn:e.target.value})}/>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:5}}>
              <label style={S.label}>Hora Salida</label>
              <input style={S.input} type="time" value={editForm.checkOut} onChange={e=>setEditForm({...editForm,checkOut:e.target.value})}/>
            </div>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:24}}>
            <button style={S.btnSec} onClick={()=>setEditEntry(null)}>Cancelar</button>
            <button style={S.btnPrimary} onClick={saveEdit} disabled={editSaving}>{editSaving?"Guardando...":"Guardar Cambios"}</button>
          </div>
        </div></div>
      )}

      <div style={S.card}>
        <div style={{display:"flex",gap:8,marginBottom:20}}>
          <button style={mode==="import"?S.btnPrimary:S.btnSec} onClick={()=>setMode("import")}>📥 Importar Archivo</button>
          <button style={mode==="manual"?S.btnPrimary:S.btnSec} onClick={()=>setMode("manual")}>✍️ Entrada Manual</button>
        </div>
        {mode==="import"?(<div>
          <p style={{fontSize:13,color:"#475569",marginBottom:14,lineHeight:1.6}}>Sube el archivo <strong>.xls</strong> o <strong>.csv</strong> exportado del reloj marcador. El sistema agrupa automáticamente las 4 marcaciones diarias por empleado.</p>
          <label style={S.fileLabel}>{importing?"⏳ Procesando archivo...":"📄 Seleccionar archivo del reloj (.xls / .csv)"}<input type="file" accept=".csv,.txt,.xls,.xlsx" onChange={handleFile} style={{display:"none"}} disabled={importing}/></label>
          {importResult&&(<div style={{marginTop:14,padding:14,borderRadius:10,background:importResult.error?"#fef2f2":"#f0fdf4",border:`1px solid ${importResult.error?"#fecaca":"#bbf7d0"}`}}>{importResult.error?<p style={{color:"#dc2626",fontSize:13}}>⚠️ {importResult.error}</p>:<p style={{color:"#0a6847",fontSize:13,fontWeight:500}}>✅ <strong>{importResult.added}</strong> registros guardados en base de datos.{importResult.skipped>0&&` (${importResult.skipped} duplicados omitidos)`}</p>}</div>)}
        </div>):(<div style={S.formGrid}>
          <Field label="Empleado" value={mf.empId} onChange={v=>setMf({...mf,empId:v})} type="select" options={[{v:"",l:"— Seleccionar —"},...employees.map(e=>({v:e.id,l:`${e.id} - ${e.name}`}))]}/>
          <Field label="Fecha" value={mf.date} onChange={v=>setMf({...mf,date:v})} type="date"/>
          <Field label="Entrada" value={mf.cin} onChange={v=>setMf({...mf,cin:v})} type="time"/>
          <Field label="Sale Almuerzo" value={mf.lout} onChange={v=>setMf({...mf,lout:v})} type="time"/>
          <Field label="Regresa" value={mf.lin} onChange={v=>setMf({...mf,lin:v})} type="time"/>
          <Field label="Salida" value={mf.cout} onChange={v=>setMf({...mf,cout:v})} type="time"/>
          <div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={addManual}>Registrar</button></div>
        </div>)}
      </div>

      <div style={S.card}>
        <div style={S.titleRow}>
          <h3 style={S.cardTitle}>
            <span style={S.cardTitleIcon}>📋</span> Registros ({filtered.length})
            {errorCount > 0 && <span style={{marginLeft:10,padding:"3px 10px",background:"#fef2f2",color:"#dc2626",borderRadius:20,fontSize:12,fontWeight:700,border:"1px solid #fecaca"}}>⚠️ {errorCount} con errores</span>}
          </h3>
          {clockEntries.length>0&&<button style={{...S.btnDanger}} onClick={clearAll}>Borrar Todo</button>}
        </div>
        <div style={{display:"flex",gap:12,marginBottom:14,flexWrap:"wrap"}}>
          <select style={{...S.input,width:220}} value={filterEmp} onChange={e=>setFilterEmp(e.target.value)}><option value="">Todos los empleados</option>{employees.map(e=><option key={e.id} value={e.id}>{e.id} - {e.name}</option>)}</select>
          <input style={{...S.input,width:170}} type="month" value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}/>
        </div>
        {filtered.length===0?<p style={{color:"#94a3b8",fontStyle:"italic"}}>No hay registros para este filtro.</p>:(
          <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
            {["Empleado","Fecha","Entrada","Sal.Almuerzo","Reg.Almuerzo","Salida","Hrs Efect.","Extra","Estado",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=6&&i<=7?"right":"left",...(i>=8?{textAlign:"center"}:{})}}>{h}</th>)}
          </tr></thead><tbody>
            {filtered.map(e=>{const hrs=calcDayHours(e),dow=new Date(e.date+"T12:00:00").getDay(),extra=Math.max(0,hrs-getScheduledHours(dow)),isWE=dow===0||dow===6;
            const error = hasError(e);
            const rowBg = error ? "#fef2f2" : isWE ? "#fffbeb" : {};
            return(<tr key={e.id} style={{background:rowBg}}>
              <td style={{...S.td,fontWeight:600,color:"#0a2351",whiteSpace:"nowrap"}}>{getN(e.employeeId)}</td>
              <td style={S.td}>{fmtDate(e.date+"T12:00:00")}{isWE&&<span style={{marginLeft:4,fontSize:10,color:"#c9a227",fontWeight:700,background:"#fffbeb",padding:"1px 5px",borderRadius:4}}>{dow===0?"DOM":"SÁB"}</span>}</td>
              <td style={{...S.td,color:e.checkIn?"#0a6847":"#dc2626",fontWeight:600}}>{e.checkIn?fmtTime(e.checkIn):<span style={{color:"#dc2626"}}>—</span>}</td>
              <td style={{...S.td,color:"#92400e"}}>{fmtTime(e.lunchOut)}</td>
              <td style={{...S.td,color:"#92400e"}}>{fmtTime(e.lunchIn)}</td>
              <td style={{...S.td,color:e.checkOut?"#b91c1c":"#dc2626",fontWeight:600}}>{e.checkOut?fmtTime(e.checkOut):<span style={{color:"#dc2626"}}>—</span>}</td>
              <td style={{...S.td,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{hrs.toFixed(1)}h</td>
              <td style={{...S.td,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:extra>0?"#b91c1c":"#cbd5e1",fontWeight:extra>0?700:400}}>{extra>0?`+${extra.toFixed(1)}h`:"—"}</td>
              <td style={{...S.td,textAlign:"center"}}>
                {error
                  ? <span style={{fontSize:11,color:"#dc2626",fontWeight:600,background:"#fef2f2",padding:"2px 8px",borderRadius:4,border:"1px solid #fecaca",whiteSpace:"nowrap"}}>⚠️ {error}</span>
                  : <span style={{fontSize:11,color:"#0a6847",fontWeight:500}}>✓ OK</span>
                }
              </td>
              <td style={{...S.td,textAlign:"center",whiteSpace:"nowrap"}}>
                <button style={{...S.tblBtn,color:"#0a2351",borderColor:"#b0c4de"}} onClick={()=>openEdit(e)}>✏️ Editar</button>
                <button style={{...S.tblBtn,color:"#dc2626",borderColor:"#fecaca"}} onClick={()=>removeEntry(e.id)}>✕</button>
              </td>
            </tr>)})}
          </tbody></table></div>)}
      </div>
    </div>
  );
}

// ═══ PAYROLL ═══
function PayrollTab({ employees, clockEntries, refresh, holidays }) {
  const [weekNum,setWeekNum]=useState("");const [from,setFrom]=useState("");const [to,setTo]=useState("");const [result,setResult]=useState(null);const [adj,setAdj]=useState({});const [saving,setSaving]=useState(false);

  const weeklyEmps = employees.filter(e => { const t = e.empType || "weekly"; return t === "weekly" || t === "weekly_nonclock"; });

  const generate=()=>{
    if(!from||!to)return alert("Selecciona las fechas.");
    const fs=from.slice(0,10), ts=to.slice(0,10);

    // Filter clock entries in date range
    const pe=clockEntries.filter(c=>{const d=String(c.date).slice(0,10);return d>=fs&&d<=ts});

    // Get holidays in this period
    const holidayDates = new Set(holidays.filter(h => h.date >= fs && h.date <= ts).map(h => h.date));
    const holidayNames = {};
    holidays.forEach(h => { if (h.date >= fs && h.date <= ts) holidayNames[h.date] = h.name; });
    // Count working days in period (Mon-Fri, excluding first Friday)
    // NOTE: Holidays ARE counted as working days (they are paid and count as "worked")
    // The employee only needs clock entries for non-holiday weekdays
    const workingDaysInPeriod = (() => {
      let count = 0;
      const start = new Date(fs + "T12:00:00");
      const end = new Date(ts + "T12:00:00");
      const firstDate = fs;
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const dow = d.getDay();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        const dateStr = `${yyyy}-${mm}-${dd}`;
        const isWeekday = dow >= 1 && dow <= 5;
        const isFirstFriday = dateStr === firstDate && dow === 5;
        // Count ALL weekdays except first Friday (holidays included as they count as worked)
        if (isWeekday && !isFirstFriday) count++;
      }
      return count;
    })();

    // Count how many holidays fall on weekdays in this period (excluding first Friday)
    const holidaysOnWeekdays = (() => {
      let count = 0;
      const start = new Date(fs + "T12:00:00");
      const end = new Date(ts + "T12:00:00");
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const dow = d.getDay();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        const dateStr = `${yyyy}-${mm}-${dd}`;
        const isFirstFriday = dateStr === fs && dow === 5;
        if (dow >= 1 && dow <= 5 && !isFirstFriday && holidayDates.has(dateStr)) count++;
      }
      return count;
    })();

    // Group entries by employee
    const byEmp = {};
    pe.forEach(e => {
      if (!byEmp[e.employeeId]) byEmp[e.employeeId] = [];
      byEmp[e.employeeId].push(e);
    });

    const fridayAutoFills = []; // Track auto-filled fridays for warnings

    const rows = weeklyEmps.map(emp => {
      const isNonClock = emp.empType === "weekly_nonclock";

      // Non-clock employees: always get full 7 days unless manually adjusted
      if (isNonClock) {
        const daily = emp.salary / 30;
        const hourly = daily / 8;
        const a = adj[emp.id] || {};
        const faltasManual = +a.faltas || 0; // Manual absence days
        const daysPaid = Math.max(0, 7 - (faltasManual * 2));
        const baseSalary = daily * daysPaid;
        const fuel=+a.fuel||0, vacation=+a.vacation||0, incapacity=+a.incapacity||0;
        const advance=+a.advance||0, dec4=+a.dec4||0, dec3=+a.dec3||0, otherDed=+a.otherDed||0;
        const applyIHSS = isLastWeekOfMonth(fs, ts);
        const ihss = applyIHSS ? calcIHSS_monthly(emp.salary) : { em: 0, ivm: 0, total: 0 };
        const applyRAP = isSecondWeekOfMonth(fs);
        const rapData = applyRAP ? calcRAP_monthly(emp.salary) : { feo3: 0, fio3: 0, employeeTotal: 0, rl: 0 };
        const rap = rapData.employeeTotal;
        const totalEarned = baseSalary + fuel + vacation + incapacity + dec4 + dec3;
        const totalDeductions = ihss.total + rap + advance + otherDed;
        return {
          employeeId: emp.id, name: emp.name, position: emp.position,
          salary: emp.salary, daily, hourly, isNonClock: true,
          daysWorked: workingDaysInPeriod - faltasManual, absences: faltasManual, daysPaid,
          holidaysInPeriod: holidayDates.size, days: daysPaid,
          effectiveHrs: 0, baseSalary,
          ot: { 0.25: 0, 0.5: 0, 0.75: 0, 1.0: 0 }, otPay: 0,
          ihssEM: ihss.em, ihssIVM: ihss.ivm, ihssTotal: ihss.total,
          rap, rapFeo3: rapData.feo3, rapFio3: rapData.fio3, rapRL: rapData.rl,
          fuel, vacation, incapacity, advance, dec4, dec3, otherDed,
          totalEarned, totalDeductions, netPay: totalEarned - totalDeductions,
        };
      }

      // Regular clock employees
      let empEntries = (byEmp[emp.id] || []).map(e => {
        // Auto-fill: Friday with checkIn but no checkOut → set checkOut to 5:00pm
        if (e.checkIn && !e.checkOut) {
          const dow = new Date(e.date + "T12:00:00").getDay();
          if (dow === 5) { // Friday
            fridayAutoFills.push({ name: emp.name, date: e.date });
            return { ...e, checkOut: `${e.date}T17:00:00`, autoFilled: true };
          }
        }
        return e;
      });

      // Count days with valid clock entries (Mon-Fri, excluding holidays and first Friday)
      // Holidays are NOT required to have clock entries - they count as worked automatically
      const clockedDays = empEntries.filter(e => {
        if (!e.checkIn || !e.checkOut) return false;
        const dow = new Date(e.date + "T12:00:00").getDay();
        const isHoliday = holidayDates.has(e.date);
        const isFirstFriday = e.date === fs && dow === 5;
        return dow >= 1 && dow <= 5 && !isHoliday && !isFirstFriday;
      }).length;

      // Count holidays on weekdays directly here (bulletproof)
      let empHolidayCount = 0;
      {
        const start = new Date(fs + "T12:00:00");
        const end = new Date(ts + "T12:00:00");
        for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
          const dow = d.getDay();
          const yyyy = d.getFullYear();
          const mm = String(d.getMonth() + 1).padStart(2, "0");
          const dd = String(d.getDate()).padStart(2, "0");
          const dateStr = `${yyyy}-${mm}-${dd}`;
          const isFirstFriday = dateStr === fs && dow === 5;
          if (dow >= 1 && dow <= 5 && !isFirstFriday && holidayDates.has(dateStr)) empHolidayCount++;
        }
      }

      // Total days worked = clocked days + holidays (holidays auto-count as worked)
      const daysWorked = clockedDays + empHolidayCount;

      // How many holidays fell in this period
      const holidaysInPeriod = holidayDates.size;

      // Pay logic
      const absences = Math.max(0, workingDaysInPeriod - daysWorked);
      const daysPaid = Math.max(0, 7 - (absences * 2));

      console.log(`[${emp.name}] clocked=${clockedDays} holidays=${empHolidayCount} worked=${daysWorked} required=${workingDaysInPeriod} absences=${absences} paid=${daysPaid}`);

      const daily = emp.salary / 30;
      const hourly = daily / 8;
      const baseSalary = daily * daysPaid;

      // Overtime
      let ot = { 0.25: 0, 0.5: 0, 0.75: 0, 1.0: 0 };
      let totalEffective = 0;

      empEntries.forEach(entry => {
        if (!entry.checkIn || !entry.checkOut) return;
        const dow = new Date(entry.checkIn).getDay();
        const hrs = calcDayHours(entry);
        totalEffective += hrs;

        // Sunday or holiday → 100%
        if (dow === 0 || holidayDates.has(entry.date)) {
          ot[1.0] += hrs;
          return;
        }
        // Saturday → 25%
        if (dow === 6) {
          ot[0.25] += hrs;
          return;
        }

        const scheduled = getScheduledHours(dow);
        const GRACE = 10/60; // 10 min grace period
        if (hrs <= scheduled + GRACE) return;
        const extra = hrs - scheduled - GRACE;
        const cH = new Date(entry.checkOut).getHours() + new Date(entry.checkOut).getMinutes() / 60;
        // 6pm-7pm (18-19) = 25%, 7:01pm-9pm (19-21) = 50%, 9pm-5am (21-5) = 75%
        if (cH <= 19) { ot[0.25] += extra; }
        else if (cH <= 21) {
          const hrs25 = Math.max(0, Math.min(extra, 1));
          ot[0.25] += hrs25;
          ot[0.5] += extra - hrs25;
        } else {
          const hrs25 = Math.max(0, Math.min(extra, 1));
          const hrs50 = Math.max(0, Math.min(extra - hrs25, 2));
          const hrs75 = Math.max(0, extra - hrs25 - hrs50);
          ot[0.25] += hrs25; ot[0.5] += hrs50; ot[0.75] += hrs75;
        }
      });

      const otPay = Object.entries(ot).reduce((sum, [rate, hrs]) => sum + hrs * hourly * (1 + parseFloat(rate)), 0);

      const a = adj[emp.id] || {};
      const fuel=+a.fuel||0, vacation=+a.vacation||0, incapacity=+a.incapacity||0;
      const advance=+a.advance||0, dec4=+a.dec4||0, dec3=+a.dec3||0, otherDed=+a.otherDed||0;

      // IHSS: only last week of the month (full monthly amount)
      const applyIHSS = isLastWeekOfMonth(fs, ts);
      const ihss = applyIHSS ? calcIHSS_monthly(emp.salary) : { em: 0, ivm: 0, total: 0 };

      // RAP 1.5%: only second week of the month (full monthly amount)
      const applyRAP = isSecondWeekOfMonth(fs);
      const rapData = applyRAP ? calcRAP_monthly(emp.salary) : { feo3: 0, fio3: 0, employeeTotal: 0, rl: 0 };
      const rap = rapData.employeeTotal; // Only employee portion is deducted

      const totalEarned = baseSalary + otPay + fuel + vacation + incapacity + dec4 + dec3;
      const totalDeductions = ihss.total + rap + advance + otherDed;

      return {
        employeeId: emp.id, name: emp.name, position: emp.position,
        salary: emp.salary, daily, hourly,
        daysWorked, absences, daysPaid, holidaysInPeriod,
        days: daysPaid,
        effectiveHrs: totalEffective, baseSalary,
        ot, otPay,
        ihssEM: ihss.em, ihssIVM: ihss.ivm, ihssTotal: ihss.total,
        rap, rapFeo3: rapData.feo3, rapFio3: rapData.fio3, rapRL: rapData.rl,
        fuel, vacation, incapacity, advance, dec4, dec3, otherDed,
        totalEarned, totalDeductions, netPay: totalEarned - totalDeductions,
      };
    });

    const applyIHSS = isLastWeekOfMonth(fs, ts);
    const applyRAP = isSecondWeekOfMonth(fs);
    setResult({ period: `WK${weekNum||"?"} ${from} al ${to}`, rows, from, to, weekNum, workingDaysInPeriod, holidaysInPeriod: holidayDates.size, holidayNames: Object.entries(holidayNames).map(([d,n])=>`${n} (${d})`), applyIHSS, applyRAP, fridayAutoFills });
  };

  const doSave=async()=>{if(!result)return;setSaving(true);await db.savePayroll(result);await refresh();setSaving(false);alert("✅ Planilla guardada.")};
  const updAdj=(eId,f,v)=>setAdj(p=>({...p,[eId]:{...(p[eId]||{}),[f]:v}}));
  return (
    <div>
      <div style={S.pageHeader}><h2 style={S.title}>Generar Planilla</h2><div style={S.goldLine}/></div>
      <div style={S.card}>
        <h3 style={S.cardTitle}><span style={S.cardTitleIcon}>📅</span> Período</h3>
        <p style={{fontSize:12,color:"#64748b",marginBottom:12}}>Período de lectura: Viernes 5:00pm → Viernes siguiente 5:00pm. Selecciona las fechas de la semana laboral (Lunes a Viernes).</p>
        <div style={S.formGrid}>
          <Field label="Semana #" value={weekNum} onChange={setWeekNum} ph="3"/>
          <Field label="Desde" value={from} onChange={setFrom} type="date"/>
          <Field label="Hasta" value={to} onChange={setTo} type="date"/>
          <div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={generate}>Generar Planilla</button></div>
        </div>
      </div>

      {result&&(<>
        {/* Pay rules summary */}
        <div style={{...S.card,background:"linear-gradient(135deg,#f0f3f8,#e8eef6)",border:"1px solid #d0daea"}}>
          <div style={{display:"flex",gap:24,flexWrap:"wrap",fontSize:13,marginBottom:result.holidayNames?.length>0?10:0}}>
            <div><span style={{color:"#64748b"}}>Días laborales:</span> <strong style={{color:"#0a2351"}}>{result.workingDaysInPeriod}</strong></div>
            {result.holidaysInPeriod>0&&<div><span style={{color:"#64748b"}}>Feriados:</span> <strong style={{color:"#c9a227"}}>{result.holidaysInPeriod} (pagados)</strong></div>}
            <div><span style={{color:"#64748b"}}>Regla:</span> <strong style={{color:"#0a6847"}}>{result.workingDaysInPeriod}d trabajados = 7d pagados</strong></div>
            <div><span style={{color:"#64748b"}}>Falta:</span> <strong style={{color:"#b91c1c"}}>-2 días/ausencia</strong></div>
            <div><span style={{color:"#64748b"}}>IHSS:</span> <strong style={{color:result.applyIHSS?"#7c3aed":"#94a3b8"}}>{result.applyIHSS?"✓ Aplica (última semana)":"No aplica esta semana"}</strong></div>
            <div><span style={{color:"#64748b"}}>RAP 1.5%:</span> <strong style={{color:result.applyRAP?"#0369a1":"#94a3b8"}}>{result.applyRAP?"✓ Aplica (2da semana)":"No aplica esta semana"}</strong></div>
          </div>
          {result.holidayNames?.length>0&&(
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {result.holidayNames.map((h,i)=><span key={i} style={{padding:"3px 10px",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:6,fontSize:12,color:"#92400e",fontWeight:600}}>🎌 {h}</span>)}
            </div>
          )}
        </div>

        {/* Friday auto-fill warnings */}
        {result.fridayAutoFills?.length > 0 && (
          <div style={{...S.card, background:"#fffbeb", border:"1px solid #fde68a"}}>
            <h4 style={{fontSize:13,fontWeight:700,color:"#92400e",marginBottom:8}}>⚠️ Viernes con salida auto-completada a 5:00pm</h4>
            <p style={{fontSize:12,color:"#a16207",marginBottom:8}}>Los siguientes empleados tenían marca de entrada el viernes pero no de salida. Se usó 5:00pm como hora de salida para el cálculo. Verifica que sea correcto.</p>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {result.fridayAutoFills.map((f,i) => (
                <span key={i} style={{padding:"4px 12px",background:"#fff",border:"1px solid #fde68a",borderRadius:6,fontSize:12,color:"#92400e",fontWeight:600}}>{f.name} — {fmtDate(f.date+"T12:00:00")}</span>
              ))}
            </div>
          </div>
        )}

        <div style={S.card}>
          <div style={S.titleRow}><h3 style={S.cardTitle}><span style={S.cardTitleIcon}>📋</span> {result.period}</h3>
            <div style={{display:"flex",gap:8}}>
              <button style={S.btnGold} onClick={()=>printPayroll(result)}>🖨️ Imprimir</button>
              <button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"Guardando...":"💾 Guardar"}</button>
            </div>
          </div>
          <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
            {["Cód","Nombre","Pos.","Sal.M.","Trab.","Faltas","Pago","Salario","Tot.OT","IHSS","RAP","Devengado","Deducc.","Neto"].map((h,i)=><th key={i} style={{...S.th,fontSize:10,textAlign:i>=3?"right":"left"}}>{h}</th>)}
          </tr></thead><tbody>
            {result.rows.map(r=><tr key={r.employeeId} style={r.absences>0?{background:"#fffbeb"}:r.isNonClock?{background:"#f0f3f8"}:{}}>
              <td style={S.td}><span style={S.badge}>{r.employeeId}</span></td>
              <td style={{...S.td,fontWeight:600,whiteSpace:"nowrap",fontSize:12,color:"#0a2351"}}>{r.name} {r.isNonClock&&<span style={{fontSize:9,color:"#1d4ed8",background:"#eff6ff",padding:"1px 5px",borderRadius:3,marginLeft:4}}>SIN RELOJ</span>}</td>
              <td style={{...S.td,fontSize:11}}>{r.position}</td>
              <td style={S.tdM}>{formatL(r.salary)}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a6847"}}>{r.isNonClock?"—":`${r.daysWorked}d`}</td>
              <td style={{...S.tdM,fontWeight:700,color:r.absences>0?"#dc2626":"#0a6847"}}>{r.absences>0?`${r.absences}d`:"—"}</td>
              <td style={{...S.tdM,fontWeight:700,color:r.daysPaid<7?"#c9a227":"#0a2351"}}>{r.daysPaid}d</td>
              <td style={{...S.tdM,fontWeight:600}}>{formatL(r.baseSalary)}</td>
              <td style={{...S.tdM,fontWeight:700,color:r.otPay>0?"#b91c1c":"#d4d4d8"}}>{r.isNonClock?"N/A":formatL(r.otPay)}</td>
              <td style={{...S.tdM,color:"#7c3aed",fontSize:11}}>{r.ihssTotal>0?formatL(r.ihssTotal):"—"}</td>
              <td style={{...S.tdM,color:"#0369a1",fontSize:11}} title={r.rap>0?`FEO3: ${formatL(r.rapFeo3)} + FIO3: ${formatL(r.rapFio3)}\nRL patrono: ${formatL(r.rapRL)}`:""}>
                {r.rap>0?formatL(r.rap):"—"}
              </td>
              <td style={{...S.tdM,fontWeight:600}}>{formatL(r.totalEarned)}</td>
              <td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:13}}>{formatL(r.netPay)}</td>
            </tr>)}
          </tbody><tfoot><tr style={{background:"#e8eef6"}}>
            <td colSpan={7} style={{...S.td,fontWeight:700,color:"#0a2351",textTransform:"uppercase",fontSize:11,letterSpacing:"0.05em"}}>Totales</td>
            <td style={{...S.tdM,fontWeight:700,color:"#0a2351"}}>{formatL(result.rows.reduce((s,r)=>s+r.baseSalary,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.otPay,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#7c3aed"}}>{formatL(result.rows.reduce((s,r)=>s+r.ihssTotal,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#0369a1"}}>{formatL(result.rows.reduce((s,r)=>s+r.rap,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#0a2351"}}>{formatL(result.rows.reduce((s,r)=>s+r.totalEarned,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.totalDeductions,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:15}}>{formatL(result.rows.reduce((s,r)=>s+r.netPay,0))}</td>
          </tr></tfoot></table></div>
        </div>
        <div style={S.card}>
          <h3 style={S.cardTitle}><span style={S.cardTitleIcon}>⚙️</span> Ajustes por Empleado</h3>
          <p style={{fontSize:12,color:"#64748b",marginBottom:10}}>Ingresa los montos y genera de nuevo la planilla para que se recalcule.</p>
          <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
            {["Empleado","Tipo","Faltas","Combustible","Vacaciones","Incapacidad","Adelanto","Dec. 4to","Dec. 3ro","Otras Ded."].map((h,i)=><th key={i} style={S.th}>{h}</th>)}
          </tr></thead><tbody>
            {weeklyEmps.map(emp=>{const a=adj[emp.id]||{};const isNC=emp.empType==="weekly_nonclock";return(<tr key={emp.id} style={isNC?{background:"#f0f3f8"}:{}}>
              <td style={{...S.td,fontWeight:600,fontSize:12,whiteSpace:"nowrap",color:"#0a2351"}}>{emp.name}</td>
              <td style={{...S.td,fontSize:11}}>{isNC?<span style={{padding:"2px 8px",background:"#eff6ff",color:"#1d4ed8",borderRadius:4,fontSize:10,fontWeight:600}}>Sin Reloj</span>:<span style={{fontSize:10,color:"#64748b"}}>Reloj</span>}</td>
              <td style={S.td}>{isNC?<input style={{...S.input,width:60,padding:"5px 8px",fontSize:12,textAlign:"right"}} type="number" value={a.faltas||""} onChange={e=>updAdj(emp.id,"faltas",e.target.value)} placeholder="0"/>:<span style={{color:"#94a3b8",fontSize:11}}>auto</span>}</td>
            {["fuel","vacation","incapacity","advance","dec4","dec3","otherDed"].map(f=><td key={f} style={S.td}><input style={{...S.input,width:85,padding:"5px 8px",fontSize:12,textAlign:"right"}} type="number" value={a[f]||""} onChange={e=>updAdj(emp.id,f,e.target.value)} placeholder="0.00"/></td>)}</tr>)})}
          </tbody></table></div>
        </div>
      </>)}
    </div>
  );
}

// ═══ HISTORY ═══
function HistoryTab({ payrolls, refresh }) {
  const [sel,setSel]=useState(null);const [rows,setRows]=useState(null);const [loadingRows,setLoadingRows]=useState(false);
  const selectPayroll=async i=>{setSel(i);setLoadingRows(true);setRows(await db.getPayrollRows(payrolls[i].id));setLoadingRows(false)};
  const del=async i=>{if(!confirm("¿Eliminar esta planilla del historial?"))return;await db.deletePayroll(payrolls[i].id);await refresh();setSel(null);setRows(null)};
  return (
    <div>
      <div style={S.pageHeader}><h2 style={S.title}>Historial de Planillas</h2><div style={S.goldLine}/></div>
      {payrolls.length===0?<div style={S.card}><p style={{color:"#94a3b8",fontStyle:"italic"}}>No hay planillas guardadas aún.</p></div>:(<>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:20}}>
          {payrolls.map((p,i)=><button key={i} onClick={()=>selectPayroll(i)} style={sel===i?{...S.btnPrimary,boxShadow:"0 4px 12px rgba(10,35,81,0.3)"}:S.btnSec}>{p.period}</button>)}
        </div>
        {loadingRows&&<div style={S.card}><p style={{color:"#64748b"}}>Cargando detalle de planilla...</p></div>}
        {sel!==null&&rows&&!loadingRows&&(
          <div style={S.card}>
            <div style={S.titleRow}><h3 style={S.cardTitle}><span style={S.cardTitleIcon}>📋</span> {payrolls[sel].period}</h3>
              <div style={{display:"flex",gap:8}}>
                <button style={S.btnGold} onClick={()=>printPayroll({period:payrolls[sel].period,rows})}>🖨️ Imprimir Planilla</button>
                <button style={S.btnDanger} onClick={()=>del(sel)}>Eliminar</button>
              </div>
            </div>
            <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
              {["Cód","Nombre","Posición","Días","Horas","Salario","H. Extra","Devengado","Deducciones","Neto a Pagar"].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=3?"right":"left"}}>{h}</th>)}
            </tr></thead><tbody>
              {rows.map(r=><tr key={r.employeeId}><td style={S.td}><span style={S.badge}>{r.employeeId}</span></td><td style={{...S.td,fontWeight:600,color:"#0a2351"}}>{r.name}</td><td style={S.td}>{r.position}</td><td style={{...S.tdM,fontWeight:700}}>{r.days}</td><td style={S.tdM}>{r.effectiveHrs?.toFixed(1)||"—"}h</td><td style={S.tdM}>{formatL(r.baseSalary)}</td><td style={{...S.tdM,color:r.otPay>0?"#b91c1c":"#d4d4d8",fontWeight:r.otPay>0?700:400}}>{formatL(r.otPay)}</td><td style={{...S.tdM,fontWeight:600}}>{formatL(r.totalEarned)}</td><td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td><td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:13}}>{formatL(r.netPay)}</td></tr>)}
            </tbody><tfoot><tr style={{background:"#e8eef6"}}>
              <td colSpan={5} style={{...S.td,fontWeight:700,color:"#0a2351",textTransform:"uppercase",fontSize:11}}>Totales</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a2351"}}>{formatL(rows.reduce((s,r)=>s+r.baseSalary,0))}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(rows.reduce((s,r)=>s+r.otPay,0))}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a2351"}}>{formatL(rows.reduce((s,r)=>s+r.totalEarned,0))}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(rows.reduce((s,r)=>s+r.totalDeductions,0))}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:15}}>{formatL(rows.reduce((s,r)=>s+r.netPay,0))}</td>
            </tr></tfoot></table></div>
          </div>)}
      </>)}
    </div>
  );
}

// ═══ CONFIDENTIAL (Biweekly Payroll) ═══
function ConfidentialTab({ employees, refresh }) {
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({ id:"",name:"",position:"",salary:"" });
  const [saving, setSaving] = useState(false);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [result, setResult] = useState(null);
  const [adj, setAdj] = useState({});
  const [paySaving, setPaySaving] = useState(false);

  // Filter confidential employees (emp_type = 'biweekly')
  const confEmployees = employees.filter(e => e.empType === "biweekly");

  const openAdd=()=>{setForm({id:"",name:"",position:"",salary:""});setModal("new")};
  const openEdit=(e)=>{setForm({...e,salary:String(e.salary)});setModal(e.id)};
  const doSave=async()=>{
    const emp={...form,salary:parseFloat(form.salary)||0};
    if(!emp.id||!emp.name)return;
    setSaving(true);
    await supabase.from("employees").upsert({id:emp.id,name:emp.name,position:emp.position,salary:emp.salary,active:true,emp_type:"biweekly",updated_at:new Date().toISOString()});
    await refresh();
    setSaving(false);
    setModal(null);
  };
  const doDelete=async(id)=>{if(confirm("¿Eliminar?")) {await supabase.from("employees").update({active:false}).eq("id",id);await refresh()}};

  const generate = () => {
    if (!from || !to) return alert("Selecciona las fechas de la quincena.");
    const rows = confEmployees.map(emp => {
      const biweeklySalary = emp.salary / 2; // Quincenal = mensual / 2
      const ihss = calcIHSS_biweekly(emp.salary);
      const a = adj[emp.id] || {};
      const advance = +a.advance || 0, otherDed = +a.otherDed || 0;
      const totalEarned = biweeklySalary;
      const totalDeductions = ihss.total + advance + otherDed;
      return {
        employeeId: emp.id, name: emp.name, position: emp.position,
        salary: emp.salary, baseSalary: biweeklySalary,
        ihssEM: ihss.em, ihssIVM: ihss.ivm, ihssTotal: ihss.total,
        advance, otherDed,
        totalEarned, totalDeductions, netPay: totalEarned - totalDeductions,
      };
    });
    setResult({ period: `Q ${from} al ${to}`, rows, from, to });
  };

  const doSavePayroll = async () => {
    if (!result) return;
    setPaySaving(true);
    const { data: existing } = await supabase.from("payrolls_biweekly").select("id").eq("period", result.period).maybeSingle();
    let pid;
    if (existing) {
      pid = existing.id;
      await supabase.from("payroll_rows_biweekly").delete().eq("payroll_id", pid);
      await supabase.from("payrolls_biweekly").update({ date_from: result.from, date_to: result.to, total_earned: result.rows.reduce((s,r)=>s+r.totalEarned,0), total_deductions: result.rows.reduce((s,r)=>s+r.totalDeductions,0), total_net: result.rows.reduce((s,r)=>s+r.netPay,0) }).eq("id", pid);
    } else {
      const { data } = await supabase.from("payrolls_biweekly").insert({ period: result.period, date_from: result.from, date_to: result.to, total_earned: result.rows.reduce((s,r)=>s+r.totalEarned,0), total_deductions: result.rows.reduce((s,r)=>s+r.totalDeductions,0), total_net: result.rows.reduce((s,r)=>s+r.netPay,0) }).select("id").single();
      pid = data.id;
    }
    await supabase.from("payroll_rows_biweekly").insert(result.rows.map(r=>({ payroll_id:pid, employee_id:r.employeeId, name:r.name, position:r.position, salary:r.salary, base_salary:r.baseSalary, ihss_em:r.ihssEM, ihss_ivm:r.ihssIVM, ihss_total:r.ihssTotal, other_ded:r.otherDed, total_earned:r.totalEarned, total_deductions:r.totalDeductions, net_pay:r.netPay })));
    setPaySaving(false);
    alert("✅ Planilla quincenal guardada.");
  };

  const updAdj=(eId,f,v)=>setAdj(p=>({...p,[eId]:{...(p[eId]||{}),[f]:v}}));

  const printConf = () => {
    if (!result) return;
    const rows = result.rows;
    const totals = { base: rows.reduce((s,r)=>s+r.baseSalary,0), ihss: rows.reduce((s,r)=>s+r.ihssTotal,0), earned: rows.reduce((s,r)=>s+r.totalEarned,0), ded: rows.reduce((s,r)=>s+r.totalDeductions,0), net: rows.reduce((s,r)=>s+r.netPay,0) };
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Planilla Quincenal ${result.period}</title><style>@page{size:landscape;margin:12mm}*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:9pt}.header{text-align:center;margin-bottom:12px}.header img{height:40px;margin-bottom:6px}.header h2{font-size:11pt;color:#1a3a6b}.header .period{font-size:10pt;margin-top:4px;font-weight:bold;color:#0a2351}table{width:100%;border-collapse:collapse;margin-top:8px}th{background:#0a2351;color:#fff;padding:5px 4px;font-size:8pt;text-transform:uppercase;border:1px solid #0d2d6b;text-align:center}td{padding:4px;border:1px solid #c8d6e5;font-size:8pt}.r{text-align:right;font-family:'Courier New',monospace}.name{font-weight:600;white-space:nowrap}.total-row{background:#e8eef6;font-weight:700}.total-row td{border-top:2px solid #0a2351}.net{color:#0a6847;font-weight:700}.signatures{margin-top:40px;display:flex;justify-content:space-between}.sig-box{text-align:center;width:200px}.sig-line{border-top:1px solid #000;margin-top:50px;padding-top:4px;font-size:9pt}@media print{.no-print{display:none!important}}.no-print{position:fixed;top:10px;right:10px;z-index:999}.print-btn{padding:10px 24px;background:#0a2351;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;margin-right:8px}.close-btn{padding:10px 24px;background:#64748b;color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer}</style></head><body><div class="no-print"><button class="print-btn" onclick="window.print()">🖨️ Imprimir</button><button class="close-btn" onclick="window.close()">✕ Cerrar</button></div><div class="header"><img src="${LOGO}" alt="Horeb"/><h2>Planilla Quincenal — Empleados Confidenciales</h2><div class="period">${result.period}</div></div><table><thead><tr><th>Cód.</th><th>Nombre</th><th>Posición</th><th>Sal.Mensual</th><th>Sal.Quincenal</th><th>IHSS EM</th><th>IHSS IVM</th><th>Total IHSS</th><th>Adelanto</th><th>Otras Ded.</th><th>Total Ded.</th><th>Neto</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${r.employeeId}</td><td class="name">${r.name}</td><td>${r.position}</td><td class="r">${fN(r.salary)}</td><td class="r">${fN(r.baseSalary)}</td><td class="r">${fN(r.ihssEM)}</td><td class="r">${fN(r.ihssIVM)}</td><td class="r">${fN(r.ihssTotal)}</td><td class="r">${fN(r.advance)}</td><td class="r">${fN(r.otherDed)}</td><td class="r">${fN(r.totalDeductions)}</td><td class="r net">${fN(r.netPay)}</td></tr>`).join("")}<tr class="total-row"><td colspan="4" style="text-align:right">TOTALES</td><td class="r">${fN(totals.base)}</td><td colspan="2"></td><td class="r">${fN(totals.ihss)}</td><td colspan="2"></td><td class="r">${fN(totals.ded)}</td><td class="r net">${fN(totals.net)}</td></tr></tbody></table><div class="signatures"><div class="sig-box"><div class="sig-line">Elaborado por</div></div><div class="sig-box"><div class="sig-line">Revisado por</div></div><div class="sig-box"><div class="sig-line">Autorizado por</div></div></div></body></html>`;
    const w = window.open("","_blank"); if(w){w.document.write(html);w.document.close()}
  };

  return (
    <div>
      <div style={S.pageHeader}><h2 style={S.title}>Planilla Confidencial (Quincenal)</h2><div style={S.goldLine}/></div>

      {/* Confidential Employees */}
      <div style={S.card}>
        <div style={S.titleRow}><h3 style={S.cardTitle}><span style={S.cardTitleIcon}>🔒</span> Empleados Confidenciales</h3><button style={S.btnPrimary} onClick={openAdd}>+ Agregar</button></div>
        <p style={{fontSize:12,color:"#64748b",marginBottom:12}}>Empleados de confianza que no marcan reloj. Pago quincenal (salario mensual ÷ 2).</p>

        {modal&&(<div style={S.overlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>{modal==="new"?"Nuevo Empleado Confidencial":"Editar"}</h3>
          <div style={S.formGrid}>
            <Field label="Código" value={form.id} onChange={v=>setForm({...form,id:v})} ph="C01" disabled={modal!=="new"}/>
            <Field label="Nombre" value={form.name} onChange={v=>setForm({...form,name:v})} ph="Nombre completo"/>
            <Field label="Posición" value={form.position} onChange={v=>setForm({...form,position:v})} ph="Gerente"/>
            <Field label="Salario Mensual" value={form.salary} onChange={v=>setForm({...form,salary:v})} ph="0.00" type="number"/>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:24}}>
            <button style={S.btnSec} onClick={()=>setModal(null)}>Cancelar</button>
            <button style={S.btnPrimary} onClick={doSave} disabled={saving}>{saving?"Guardando...":"Guardar"}</button>
          </div>
        </div></div>)}

        {confEmployees.length===0?<p style={{color:"#94a3b8",fontStyle:"italic"}}>No hay empleados confidenciales registrados.</p>:(
          <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
            {["Cód","Nombre","Posición","Sal. Mensual","Sal. Quincenal","IHSS/Quincena",""].map((h,i)=><th key={i} style={{...S.th,textAlign:i>=3?"right":"left"}}>{h}</th>)}
          </tr></thead><tbody>
            {confEmployees.map(e=>{const ihss=calcIHSS_biweekly(e.salary);return(
              <tr key={e.id}><td style={S.td}><span style={S.badge}>{e.id}</span></td><td style={{...S.td,fontWeight:600,color:"#0a2351"}}>{e.name}</td><td style={S.td}><span style={S.posBadge}>{e.position}</span></td><td style={S.tdM}>{formatL(e.salary)}</td><td style={S.tdM}>{formatL(e.salary/2)}</td><td style={{...S.tdM,color:"#7c3aed"}}>{formatL(ihss.total)}</td><td style={{...S.td,textAlign:"center"}}><button style={S.tblBtn} onClick={()=>openEdit(e)}>Editar</button><button style={{...S.tblBtn,color:"#dc2626",borderColor:"#fecaca"}} onClick={()=>doDelete(e.id)}>Eliminar</button></td></tr>
            )})}
          </tbody></table></div>
        )}
      </div>

      {/* Generate Biweekly Payroll */}
      <div style={S.card}>
        <h3 style={S.cardTitle}><span style={S.cardTitleIcon}>📋</span> Generar Planilla Quincenal</h3>
        <div style={S.formGrid}>
          <Field label="Desde" value={from} onChange={setFrom} type="date"/>
          <Field label="Hasta" value={to} onChange={setTo} type="date"/>
          <div style={{display:"flex",alignItems:"flex-end"}}><button style={S.btnPrimary} onClick={generate}>Generar</button></div>
        </div>
      </div>

      {result&&(
        <div style={S.card}>
          <div style={S.titleRow}><h3 style={S.cardTitle}><span style={S.cardTitleIcon}>📋</span> {result.period}</h3>
            <div style={{display:"flex",gap:8}}>
              <button style={S.btnGold} onClick={printConf}>🖨️ Imprimir</button>
              <button style={S.btnPrimary} onClick={doSavePayroll} disabled={paySaving}>{paySaving?"Guardando...":"💾 Guardar"}</button>
            </div>
          </div>
          <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
            {["Cód","Nombre","Posición","Sal.Mensual","Quincenal","IHSS EM","IHSS IVM","Tot.IHSS","Adelanto","Otras Ded.","Tot.Ded.","Neto"].map((h,i)=><th key={i} style={{...S.th,fontSize:10,textAlign:i>=3?"right":"left"}}>{h}</th>)}
          </tr></thead><tbody>
            {result.rows.map(r=><tr key={r.employeeId}>
              <td style={S.td}><span style={S.badge}>{r.employeeId}</span></td>
              <td style={{...S.td,fontWeight:600,color:"#0a2351",whiteSpace:"nowrap"}}>{r.name}</td>
              <td style={{...S.td,fontSize:11}}>{r.position}</td>
              <td style={S.tdM}>{formatL(r.salary)}</td>
              <td style={{...S.tdM,fontWeight:600}}>{formatL(r.baseSalary)}</td>
              <td style={{...S.tdM,color:"#7c3aed"}}>{formatL(r.ihssEM)}</td>
              <td style={{...S.tdM,color:"#7c3aed"}}>{formatL(r.ihssIVM)}</td>
              <td style={{...S.tdM,fontWeight:600,color:"#7c3aed"}}>{formatL(r.ihssTotal)}</td>
              <td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.advance)}</td>
              <td style={{...S.tdM,color:"#b91c1c"}}>{formatL(r.otherDed)}</td>
              <td style={{...S.tdM,fontWeight:600,color:"#b91c1c"}}>{formatL(r.totalDeductions)}</td>
              <td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:13}}>{formatL(r.netPay)}</td>
            </tr>)}
          </tbody><tfoot><tr style={{background:"#e8eef6"}}>
            <td colSpan={4} style={{...S.td,fontWeight:700,color:"#0a2351",textTransform:"uppercase",fontSize:11}}>Totales</td>
            <td style={{...S.tdM,fontWeight:700}}>{formatL(result.rows.reduce((s,r)=>s+r.baseSalary,0))}</td>
            <td colSpan={2}></td>
            <td style={{...S.tdM,fontWeight:700,color:"#7c3aed"}}>{formatL(result.rows.reduce((s,r)=>s+r.ihssTotal,0))}</td>
            <td colSpan={2}></td>
            <td style={{...S.tdM,fontWeight:700,color:"#b91c1c"}}>{formatL(result.rows.reduce((s,r)=>s+r.totalDeductions,0))}</td>
            <td style={{...S.tdM,fontWeight:700,color:"#0a6847",fontSize:15}}>{formatL(result.rows.reduce((s,r)=>s+r.netPay,0))}</td>
          </tr></tfoot></table></div>

          {/* Adjustments */}
          <div style={{marginTop:16}}>
            <h4 style={{fontSize:13,fontWeight:700,color:"#475569",marginBottom:8}}>Ajustes</h4>
            <div style={{overflowX:"auto"}}><table style={S.table}><thead><tr>
              {["Empleado","Adelanto","Otras Ded."].map((h,i)=><th key={i} style={S.th}>{h}</th>)}
            </tr></thead><tbody>
              {confEmployees.map(emp=>{const a=adj[emp.id]||{};return(<tr key={emp.id}>
                <td style={{...S.td,fontWeight:600,fontSize:12,color:"#0a2351"}}>{emp.name}</td>
                {["advance","otherDed"].map(f=><td key={f} style={S.td}><input style={{...S.input,width:100,padding:"5px 8px",fontSize:12,textAlign:"right"}} type="number" value={a[f]||""} onChange={e=>updAdj(emp.id,f,e.target.value)} placeholder="0.00"/></td>)}
              </tr>)})}
            </tbody></table></div>
            <p style={{fontSize:11,color:"#64748b",marginTop:6}}>Ingresa ajustes y genera de nuevo para recalcular.</p>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══ SHARED ═══
function Field({label,value,onChange,ph,type="text",options,disabled}){return(
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    <label style={S.label}>{label}</label>
    {type==="select"?<select style={S.input} value={value} onChange={e=>onChange(e.target.value)}>{options.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}</select>
    :<input style={S.input} type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={ph} disabled={disabled}/>}
  </div>
)}

// ═══ STYLES ═══
const S = {
  app: { fontFamily:"'Source Sans 3','Segoe UI',sans-serif",background:"#f0f3f8",minHeight:"100vh",color:"#1e293b" },
  // Header
  header: { background:"linear-gradient(135deg,#0a2351 0%,#0d2d6b 100%)",padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12,borderBottom:"3px solid #c9a227",position:"sticky",top:0,zIndex:100,minHeight:60 },
  brand: { display:"flex",alignItems:"center",gap:12 },
  logoImg: { height:36,borderRadius:4,objectFit:"contain" },
  brandDivider: { width:1,height:28,background:"rgba(255,255,255,0.15)" },
  brandSub: { color:"rgba(255,255,255,0.6)",fontSize:12,fontWeight:500,letterSpacing:"0.08em",textTransform:"uppercase",fontFamily:"'Cormorant Garamond',Georgia,serif" },
  nav: { display:"flex",gap:3,flexWrap:"wrap" },
  navBtn: { display:"flex",alignItems:"center",gap:6,padding:"10px 14px",border:"none",borderRadius:6,background:"transparent",color:"rgba(255,255,255,0.55)",cursor:"pointer",fontSize:13,fontWeight:500,fontFamily:"inherit",letterSpacing:"0.02em" },
  navAct: { background:"rgba(201,162,39,0.15)",color:"#c9a227",fontWeight:700 },
  // Main
  main: { maxWidth:1320,margin:"0 auto",padding:"28px 20px 40px" },
  footer: { textAlign:"center",padding:"20px",fontSize:12,color:"#64748b",borderTop:"1px solid #e2e8f0",display:"flex",justifyContent:"center",gap:20 },
  // Typography
  pageHeader: { marginBottom:24 },
  title: { fontSize:26,fontWeight:700,color:"#0a2351",letterSpacing:"-0.02em",fontFamily:"'Cormorant Garamond','Source Sans 3',serif",marginBottom:6 },
  goldLine: { width:48,height:3,background:"linear-gradient(90deg,#c9a227,#e8d48b)",borderRadius:2 },
  // Cards
  grid4: { display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:16,marginBottom:24 },
  card: { background:"#fff",borderRadius:14,padding:22,boxShadow:"0 1px 4px rgba(10,35,81,0.06),0 0 0 1px rgba(10,35,81,0.04)",marginBottom:20 },
  cardTitle: { fontSize:15,fontWeight:700,marginBottom:14,color:"#0a2351",display:"flex",alignItems:"center",gap:8 },
  cardTitleIcon: { fontSize:18 },
  titleRow: { display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,marginBottom:14 },
  // Schedule
  schedCard: { padding:"12px 14px",borderRadius:10,background:"linear-gradient(135deg,#f0f3f8,#e8eef6)",border:"1px solid #d0daea",textAlign:"center" },
  schedDay: { fontSize:13,fontWeight:700,color:"#0a2351",marginBottom:4 },
  schedTime: { fontSize:12,color:"#475569",fontWeight:500 },
  schedHrs: { fontSize:11,color:"#64748b",marginTop:2 },
  schedNote: { fontSize:12,color:"#64748b",marginTop:12,fontStyle:"italic" },
  summRow: { display:"flex",justifyContent:"space-between",padding:"6px 0",fontSize:15 },
  // Table
  table: { width:"100%",borderCollapse:"separate",borderSpacing:0,fontSize:13 },
  th: { padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:10,color:"#64748b",borderBottom:"2px solid #d0daea",whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.06em",background:"#f8fafc" },
  td: { padding:"10px 12px",borderBottom:"1px solid #eef2f7",whiteSpace:"nowrap",fontSize:13 },
  tdM: { padding:"10px 12px",borderBottom:"1px solid #eef2f7",whiteSpace:"nowrap",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12 },
  // Badges
  badge: { background:"#e8eef6",color:"#0a2351",padding:"3px 10px",borderRadius:6,fontWeight:700,fontSize:11,fontFamily:"'JetBrains Mono',monospace" },
  posBadge: { background:"#f0fdf4",color:"#0a6847",padding:"3px 10px",borderRadius:6,fontSize:12,fontWeight:500,border:"1px solid #bbf7d0" },
  // Buttons
  btnPrimary: { padding:"10px 20px",background:"linear-gradient(135deg,#0a2351,#163a72)",color:"#fff",border:"none",borderRadius:8,fontWeight:600,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap",letterSpacing:"0.02em",boxShadow:"0 2px 8px rgba(10,35,81,0.2)" },
  btnSec: { padding:"10px 20px",background:"#fff",color:"#0a2351",border:"1px solid #d0daea",borderRadius:8,fontWeight:600,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap" },
  btnGold: { padding:"10px 20px",background:"linear-gradient(135deg,#c9a227,#dbb84a)",color:"#fff",border:"none",borderRadius:8,fontWeight:700,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap",boxShadow:"0 2px 8px rgba(201,162,39,0.3)",letterSpacing:"0.02em" },
  btnDanger: { padding:"10px 20px",background:"#fff",color:"#dc2626",border:"1px solid #fecaca",borderRadius:8,fontWeight:600,cursor:"pointer",fontSize:13,fontFamily:"inherit",whiteSpace:"nowrap" },
  tblBtn: { padding:"4px 12px",background:"#fff",color:"#0a2351",border:"1px solid #d0daea",borderRadius:6,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",marginRight:4 },
  iconBtn: { background:"none",border:"none",cursor:"pointer",padding:4,fontSize:15,opacity:0.7 },
  fileLabel: { display:"inline-flex",alignItems:"center",gap:8,padding:"12px 22px",background:"linear-gradient(135deg,#e8eef6,#f0f3f8)",color:"#0a2351",borderRadius:10,fontWeight:600,fontSize:13,cursor:"pointer",border:"2px dashed #b0c4de",fontFamily:"inherit",letterSpacing:"0.02em" },
  // Forms
  formGrid: { display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12 },
  label: { fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.06em" },
  input: { padding:"9px 12px",border:"1px solid #d0daea",borderRadius:8,fontSize:13,color:"#1e293b",background:"#fff",width:"100%",boxSizing:"border-box",fontFamily:"inherit" },
  // Modal
  overlay: { position:"fixed",inset:0,background:"rgba(10,35,81,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:999,padding:16,backdropFilter:"blur(4px)" },
  modal: { background:"#fff",borderRadius:18,padding:28,width:"100%",maxWidth:500,boxShadow:"0 24px 64px rgba(10,35,81,0.25)" },
  modalTitle: { fontSize:20,fontWeight:700,marginBottom:20,color:"#0a2351",fontFamily:"'Cormorant Garamond',serif",borderBottom:"2px solid #c9a227",paddingBottom:10,display:"inline-block" },
};