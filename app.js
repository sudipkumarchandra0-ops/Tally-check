/* Customer Tally Check
 * Everything runs in the browser. No file content leaves this page.
 *
 * Part 1 (TallyCore) is pure logic with no page code, so it can also be
 * loaded in Node for testing. Part 2 wires it to the page.
 */
(function (root) {
"use strict";

/* ===================== Part 1: core logic ===================== */

// Header names are compared after normH(): lower case, no "(Rs)", "₹" or punctuation.
const ALIASES = {
  id: ["customer id", "cust id", "customerid", "custid", "customer no", "customer number", "customer code", "cust no", "account id"],
  name: ["customer name", "cust name", "name"],
  open_date: ["account open date", "open date", "account opening date", "acct open date", "date of opening"],
  opening: ["opening balance", "opening bal"],
  deposits: ["total deposits", "total deposit", "deposits", "total credits", "total credit", "credits"],
  withdrawals: ["total withdrawals", "total withdrawal", "withdrawals", "total debits", "total debit", "debits"],
  balance: ["current balance", "closing balance", "balance", "current bal", "closing bal"],
  txns: ["no of transactions", "number of transactions", "no of txns", "transactions", "transaction count", "txn count"],
  last_date: ["last transaction date", "last txn date", "last transaction"],
  date: ["date", "transaction date", "txn date"],
  type: ["transaction type", "type", "txn type"],
  amount: ["amount", "transaction amount", "txn amount"]
};
const FIELDS = {
  name: { label: "Customer Name", kind: "text" },
  open_date: { label: "Account Open Date", kind: "date" },
  opening: { label: "Opening Balance", kind: "money" },
  deposits: { label: "Total Deposits", kind: "money" },
  withdrawals: { label: "Total Withdrawals", kind: "money" },
  balance: { label: "Current Balance", kind: "money" },
  txns: { label: "No. of Transactions", kind: "int" },
  last_date: { label: "Last Transaction Date", kind: "date" }
};
const STATUS_LABEL = { ok: "OK", bad: "MISMATCH", missing: "NOT IN MASTER", warn: "NEEDS A LOOK" };
const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

/* ---------- small helpers ---------- */
const normH = h => String(h == null ? "" : h).toLowerCase()
  .replace(/₹/g, " ").replace(/\(\s*rs\.?\s*\)/g, " ").replace(/\brs\b\.?/g, " ")
  .replace(/[^a-z0-9]+/g, " ").trim();
const normId = v => v == null ? "" : String(v).trim().toUpperCase();
const normText = v => String(v == null ? "" : v).trim().replace(/\s+/g, " ").toLowerCase();
const pad = n => String(n).padStart(2, "0");
const fmtN = n => Number(n).toLocaleString("en-IN");

// Numbers may arrive as text with commas, "₹" or "Rs". Returns null for blank, NaN for junk.
function num(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return v;
  let s = String(v).trim().replace(/^rs\.?/i, "").replace(/[₹,\s]/g, "");
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (s === "" || s === "-") return null;
  const n = Number(s.replace(/^−/, "-"));
  return isNaN(n) ? NaN : (neg ? -n : n);
}
const r2 = n => Math.round(n * 100) / 100;

// Every date becomes DD-MM-YYYY. Excel stores real dates as serial numbers.
function normDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return pad(v.getDate()) + "-" + pad(v.getMonth() + 1) + "-" + v.getFullYear();
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return pad(d.getUTCDate()) + "-" + pad(d.getUTCMonth() + 1) + "-" + d.getUTCFullYear();
  }
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/))) return pad(m[1]) + "-" + pad(m[2]) + "-" + m[3];
  if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) return pad(m[3]) + "-" + pad(m[2]) + "-" + m[1];
  if ((m = s.match(/^(\d{1,2})[-\s\/]([a-z]{3})[a-z]*[-\s\/,]+(\d{4})$/i)) && MONTHS[m[2].toLowerCase()])
    return pad(m[1]) + "-" + pad(MONTHS[m[2].toLowerCase()]) + "-" + m[3];
  if (/^\d+(\.\d+)?$/.test(s) && Number(s) > 20000 && Number(s) < 80000) return normDate(Number(s));
  return s;
}

function money(n, sym) {
  if (n == null || isNaN(n)) return "—";
  const s = Math.abs(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return (n < 0 ? "−" : "") + (sym == null ? "₹" : sym) + s;
}
const rs = n => money(n, "Rs ");
const signed = d => (d > 0 ? "+" : d < 0 ? "−" : "") + money(Math.abs(d), "");

function show(kind, v, sym) {
  if (v == null || v === "") return "—";
  if (kind === "money") { const n = num(v); return isNaN(n) ? String(v) : money(n, sym); }
  if (kind === "int") { const n = num(v); return isNaN(n) ? String(v) : fmtN(n); }
  if (kind === "date") return normDate(v);
  return String(v);
}
function parseKind(kind, v) {
  if (kind === "money") { const n = num(v); return n == null ? null : (isNaN(n) ? NaN : r2(n)); }
  if (kind === "int") { const n = num(v); return n == null ? null : n; }
  if (kind === "date") return normDate(v);
  return v == null || String(v).trim() === "" ? null : String(v).trim();
}

/* ---------- reading ---------- */
// First sheet of a workbook as an array of rows. Dates stay as serial numbers.
function workbookRows(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = root.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  return rows.map(r => { let e = r.length; while (e > 0 && (r[e - 1] == null || r[e - 1] === "")) e--; return r.slice(0, e); });
}
function csvRows(text) {
  return root.Papa.parse(text, { skipEmptyLines: "greedy" }).data;
}

function findHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 25); i++) {
    if ((rows[i] || []).some(c => ALIASES.id.includes(normH(c)))) return i;
  }
  return -1;
}
// map: field key -> first column index. perCol: column index -> field key (first occurrence only).
function mapCols(header) {
  const map = {}, perCol = [];
  header.forEach((h, i) => {
    const n = normH(h);
    for (const k in ALIASES) {
      if (ALIASES[k].includes(n)) { if (map[k] == null) { map[k] = i; perCol[i] = k; } break; }
    }
  });
  return { map, perCol };
}
const noHeader = () => new Error("Couldn’t find a “Customer ID” column in the first 25 rows of this file.");
const noRows = () => new Error("The file has a header row but no customer rows under it.");

const PARSERS = {
  txn(rows) {
    const h = findHeader(rows); if (h < 0) throw noHeader();
    const { map } = mapCols(rows[h]);
    if (map.type == null || map.amount == null || map.balance == null) {
      if (map.deposits != null || map.withdrawals != null || map.txns != null)
        throw new Error("This looks like the customer summary, not the transaction master. Put it in box 2, “Customer summary”.");
      throw new Error("This file needs Transaction Type, Amount and Current Balance columns.");
    }
    const agg = new Map(); let n = 0, prevId = null;
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || []; const id = normId(r[map.id]); if (!id) continue; n++;
      const rowNo = i + 1;
      let a = agg.get(id);
      if (!a) {
        a = { name: map.name != null ? (parseKind("text", r[map.name]) || "") : undefined, opening: null,
          dep: 0, wdl: 0, balance: null, txns: 0, last_date: map.date != null ? null : undefined,
          breaks: [], other: [], prev: null, split: false };
        agg.set(id, a);
      } else if (prevId !== id) a.split = true;
      prevId = id;
      const type = normText(r[map.type]); const amt = num(r[map.amount]); const bal = num(r[map.balance]);
      if (map.date != null) { const d = normDate(r[map.date]); if (d) a.last_date = d; }
      let expected = null;
      if (type.includes("opening")) {
        a.opening = amt != null && !isNaN(amt) ? r2(amt) : (bal != null && !isNaN(bal) ? r2(bal) : null);
        expected = a.opening;
      } else {
        a.txns++;
        if (type.startsWith("deposit") || type.startsWith("credit")) {
          a.dep += Math.round((amt || 0) * 100);
          if (a.prev != null && amt != null) expected = a.prev + amt;
        } else if (type.startsWith("withdraw") || type.startsWith("debit")) {
          a.wdl += Math.round((amt || 0) * 100);
          if (a.prev != null && amt != null) expected = a.prev - amt;
        } else a.other.push(rowNo);
      }
      if (expected != null && bal != null && Math.abs(expected - bal) > 0.0101) a.breaks.push(rowNo);
      a.prev = bal; a.balance = bal == null || isNaN(bal) ? null : r2(bal);
    }
    if (!agg.size) throw noRows();
    agg.forEach(a => { a.deposits = a.dep / 100; a.withdrawals = a.wdl / 100; delete a.dep; delete a.wdl; delete a.prev; });
    return { map: agg, rowCount: n, info: `${fmtN(n)} transactions for ${fmtN(agg.size)} customers` };
  },

  sum(rows) {
    const h = findHeader(rows); if (h < 0) throw noHeader();
    const { map } = mapCols(rows[h]);
    if (map.type != null && map.amount != null)
      throw new Error("This looks like the transaction master, not the customer summary. Put it in box 1, “Transaction master”.");
    const keys = Object.keys(FIELDS).filter(k => map[k] != null && k !== "name");
    if (!keys.length) throw new Error("This file needs columns such as Total Deposits, Total Withdrawals or Current Balance.");
    if (map.name != null) keys.unshift("name");
    const out = new Map(); const dupes = new Set();
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || []; const id = normId(r[map.id]); if (!id) continue;
      if (out.has(id)) dupes.add(id);
      const rec = {}; keys.forEach(k => rec[k] = parseKind(FIELDS[k].kind, r[map[k]])); out.set(id, rec);
    }
    if (!out.size) throw noRows();
    return { map: out, keys, dupes, info: `${fmtN(out.size)} customers, ${keys.length} columns to compare` };
  },

  input(rows) {
    const h = findHeader(rows); if (h < 0) throw noHeader();
    const header = rows[h]; const { map, perCol } = mapCols(header);
    if (map.type != null && map.amount != null)
      throw new Error("This looks like the transaction master. Put it in box 1, “Transaction master”, and put your staff’s sheet here.");
    const keys = Object.keys(FIELDS).filter(k => map[k] != null);
    if (!keys.length) throw new Error("Add at least one column to check next to Customer ID, such as Customer Name or Current Balance.");
    const data = [];
    for (let i = h + 1; i < rows.length; i++) { const r = rows[i] || []; if (!normId(r[map.id])) continue; data.push(r); }
    if (!data.length) throw noRows();
    return { header, perCol, idCol: map.id, keys, rows: data,
      info: `${fmtN(data.length)} customer${data.length === 1 ? "" : "s"} to check · ${keys.map(k => FIELDS[k].label).join(", ")}` };
  }
};

/* ---------- comparing ---------- */
function same(kind, a, b, tol) {
  if (a == null && b == null) return true;
  if (a == null || b == null) return false;
  if (kind === "money") return !isNaN(a) && !isNaN(b) && Math.round(Math.abs(a - b) * 100) <= Math.round(tol * 100);
  if (kind === "int") return a === b;
  if (kind === "date") return a === b;
  return normText(a) === normText(b);
}

// state: { txn, sum, input } as returned by PARSERS (txn/sum may be null). tol in rupees.
function runTally(state, tol) {
  tol = Math.max(0, Number(tol) || 0);
  const sum = state.sum ? state.sum.map : null, txn = state.txn ? state.txn.map : null;

  // 1. Masters against each other, and running balances
  const masterIssues = new Map();
  const addIssue = (id, msg) => { if (!masterIssues.has(id)) masterIssues.set(id, []); masterIssues.get(id).push(msg); };
  const rowList = l => (l.length > 5 ? l.slice(0, 5).join(", ") + " and " + (l.length - 5) + " more" : l.join(", "));
  if (txn) txn.forEach((a, id) => {
    if (a.breaks.length) addIssue(id, `Running balance doesn’t add up in the transaction master (row${a.breaks.length > 1 ? "s" : ""} ${rowList(a.breaks)})`);
    if (a.other.length) addIssue(id, `Unknown transaction type in the transaction master (row${a.other.length > 1 ? "s" : ""} ${rowList(a.other)})`);
    if (a.opening == null) addIssue(id, "No opening balance row in the transaction master");
    if (a.split) addIssue(id, "Transaction rows for this customer aren’t grouped together");
  });
  if (sum && state.sum.dupes) state.sum.dupes.forEach(id => addIssue(id, "Appears more than once in the customer summary"));
  if (sum && txn) {
    const CMP = ["name", "opening", "deposits", "withdrawals", "balance", "txns", "last_date"];
    sum.forEach((s, id) => {
      const t = txn.get(id);
      if (!t) { addIssue(id, "In the customer summary but has no rows in the transaction master"); return; }
      CMP.forEach(k => {
        if (s[k] === undefined || t[k] === undefined) return;
        if (k === "last_date" && (s[k] == null || !/^\d\d-\d\d-\d{4}$/.test(s[k]))) return;
        const kind = FIELDS[k].kind;
        if (!same(kind, s[k], t[k], 0)) {
          let msg = `${FIELDS[k].label}: summary shows ${show(kind, s[k], "Rs ")}, transactions give ${show(kind, t[k], "Rs ")}`;
          if (kind === "money" && s[k] != null && t[k] != null && !isNaN(s[k])) msg += ` (difference ${signed(r2(s[k] - t[k]))})`;
          addIssue(id, msg);
        }
      });
    });
    txn.forEach((t, id) => { if (!sum.has(id)) addIssue(id, "In the transaction master but not in the customer summary"); });
  }

  // 2. Input against the master (summary first, transaction totals as fallback)
  const inp = state.input, seen = new Map();
  inp.rows.forEach(r => { const id = normId(r[inp.idCol]); seen.set(id, (seen.get(id) || 0) + 1); });
  const unchecked = new Set();
  const rows = inp.rows.map((r, idx) => {
    const id = normId(r[inp.idCol]);
    const s = sum && sum.get(id), t = txn && txn.get(id);
    const res = { idx, id, raw: r, cells: {}, status: "ok", notes: [] };
    if (!s && !t) { res.status = "missing"; res.notes.push("Customer ID not found in the master files"); }
    else {
      inp.keys.forEach(k => {
        const ci = inp.perCol.indexOf(k), kind = FIELDS[k].kind;
        const iv = parseKind(kind, r[ci]);
        let mv = s ? s[k] : undefined;
        if (mv === undefined && t) mv = t[k];
        if (mv === undefined) { res.cells[k] = { state: "na" }; unchecked.add(k); return; }
        const ok = same(kind, iv, mv, tol);
        const cell = { state: ok ? "ok" : "bad", iv, mv };
        if (!ok) {
          if (kind === "money" && iv != null && !isNaN(iv) && mv != null) cell.diff = r2(iv - mv);
          const entered = iv == null ? "blank" : (Number.isNaN(iv) ? `“${r[ci]}” (not a number)` : show(kind, iv, "Rs "));
          res.notes.push(`${FIELDS[k].label}: entered ${entered}, master ${show(kind, mv, "Rs ")}` +
            (cell.diff != null ? ` (difference ${signed(cell.diff)})` : ""));
        }
        res.cells[k] = cell;
      });
      if (Object.values(res.cells).some(c => c.state === "bad")) res.status = "bad";
    }
    if (seen.get(id) > 1) { res.notes.push("This Customer ID appears more than once in the input"); if (res.status === "ok") res.status = "warn"; }
    if (masterIssues.has(id)) { res.notes.push("The master files disagree for this customer: " + masterIssues.get(id).join("; ")); if (res.status === "ok") res.status = "warn"; }
    return res;
  });
  const counts = { all: 0, ok: 0, bad: 0, missing: 0, warn: 0 };
  rows.forEach(r => { counts.all++; counts[r.status]++; });
  return { rows, counts, masterIssues, tol, unchecked: [...unchecked], at: new Date(),
    masterUsed: sum ? "summary" : "transactions", bothMasters: !!(sum && txn) };
}

/* ---------- exports ---------- */
function stamp(d) { d = d || new Date(); return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "-" + pad(d.getHours()) + pad(d.getMinutes()); }

function buildWorkbook(state, result) {
  const X = root.XLSX, inp = state.input, wb = X.utils.book_new();
  const nameCol = inp.perCol.indexOf("name");
  const H = { font: { bold: true, color: { rgb: "16212B" } }, fill: { patternType: "solid", fgColor: { rgb: "FFF2A8" } }, alignment: { vertical: "center", wrapText: true } };
  const RED = { fill: { patternType: "solid", fgColor: { rgb: "FFC7CE" } }, font: { color: { rgb: "9C0006" }, bold: true } };
  const GREEN = { fill: { patternType: "solid", fgColor: { rgb: "C6EFCE" } }, font: { color: { rgb: "006100" }, bold: true } };
  const AMBER = { fill: { patternType: "solid", fgColor: { rgb: "FFEB9C" } }, font: { color: { rgb: "7A5300" }, bold: true } };
  const MONEY = "#,##0.00";
  const styleHeader = (ws, n) => { for (let c = 0; c < n; c++) { const a = X.utils.encode_cell({ r: 0, c }); if (ws[a]) ws[a].s = H; } };

  // Sheet 1: Tally Result
  const header = inp.header.map(h => h == null ? "" : h).concat(["Tally Status", "Remarks"]);
  const nCols = header.length, sCol = nCols - 2;
  const aoa = [header];
  result.rows.forEach(r => {
    const row = inp.header.map((_, i) => {
      const v = r.raw[i]; const k = inp.perCol[i];
      // write money as real numbers so the #,##0.00 format applies
      if (k && FIELDS[k] && FIELDS[k].kind === "money") { const n = num(v); if (n != null && !isNaN(n)) return n; }
      return v == null ? "" : v;
    });
    row.push(STATUS_LABEL[r.status], r.notes.join("; ") || "Matches the master files");
    aoa.push(row);
  });
  const ws = X.utils.aoa_to_sheet(aoa);
  styleHeader(ws, nCols);
  result.rows.forEach((r, ri) => {
    const R = ri + 1;
    inp.perCol.forEach((k, ci) => {
      if (!k || !FIELDS[k]) return;
      const a = X.utils.encode_cell({ r: R, c: ci }); if (!ws[a]) ws[a] = { t: "s", v: "" };
      const kind = FIELDS[k].kind;
      if (kind === "money" && ws[a].t === "n") ws[a].z = MONEY;
      if (kind === "date" && ws[a].t === "n") ws[a].z = "dd-mm-yyyy";
      const cell = r.cells[k];
      if (cell && cell.state === "bad") {
        ws[a].s = kind === "money" && ws[a].t === "n" ? Object.assign({ numFmt: MONEY }, RED) : RED;
        ws[a].c = [{ a: "Tally check", t: `Master value: ${kind === "money" ? rs(cell.mv) : show(kind, cell.mv)}` }];
        ws[a].c.hidden = true;
      } else if (kind === "money" && ws[a].t === "n") ws[a].s = { numFmt: MONEY };
    });
    const ida = X.utils.encode_cell({ r: R, c: inp.idCol });
    if ((r.status === "bad" || r.status === "missing") && ws[ida]) ws[ida].s = RED;
    ws[X.utils.encode_cell({ r: R, c: sCol })].s = r.status === "ok" ? GREEN : (r.status === "warn" ? AMBER : RED);
    ws[X.utils.encode_cell({ r: R, c: sCol + 1 })].s = { alignment: { wrapText: true, vertical: "top" } };
  });
  ws["!cols"] = header.map((h, i) => ({ wch: i === sCol + 1 ? 70 : Math.max(12, String(h).length + 4) }));
  X.utils.book_append_sheet(wb, ws, "Tally Result");

  // Sheet 2: Mismatches
  const mm = [["Customer ID", "Customer Name", "Field", "Your Input", "Master Value", "Difference (Rs)"]];
  result.rows.forEach(r => {
    const nm = nameCol >= 0 ? (r.raw[nameCol] == null ? "" : r.raw[nameCol]) : "";
    if (r.status === "missing") mm.push([r.id, nm, "Customer ID", "Not found in master", "", ""]);
    Object.keys(r.cells).forEach(k => {
      const c = r.cells[k]; if (c.state !== "bad") return;
      const kind = FIELDS[k].kind;
      const iv = c.iv == null ? "" : (kind === "money" && !Number.isNaN(c.iv) ? c.iv : (Number.isNaN(c.iv) ? String(r.raw[inp.perCol.indexOf(k)]) : show(kind, c.iv)));
      mm.push([r.id, nm, FIELDS[k].label, iv, kind === "money" ? c.mv : show(kind, c.mv), c.diff != null ? c.diff : ""]);
    });
  });
  const ws2 = X.utils.aoa_to_sheet(mm);
  styleHeader(ws2, 6);
  for (let R = 1; R < mm.length; R++) for (const c of [3, 4, 5]) {
    const a = X.utils.encode_cell({ r: R, c }); if (ws2[a] && ws2[a].t === "n") { ws2[a].z = MONEY; ws2[a].s = { numFmt: MONEY }; }
  }
  ws2["!cols"] = [{ wch: 13 }, { wch: 26 }, { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 16 }];
  X.utils.book_append_sheet(wb, ws2, "Mismatches");

  // Sheet 3: Master Check (only when the masters disagree)
  if (result.masterIssues.size) {
    const mi = [["Customer ID", "Problem in the master files"]];
    result.masterIssues.forEach((l, id) => l.forEach(x => mi.push([id, x])));
    const ws3 = X.utils.aoa_to_sheet(mi); styleHeader(ws3, 2);
    ws3["!cols"] = [{ wch: 13 }, { wch: 100 }];
    X.utils.book_append_sheet(wb, ws3, "Master Check");
  }
  return wb;
}

function buildCsv(state, result) {
  const q = v => { const s = String(v == null ? "" : v); return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const nameCol = state.input.perCol.indexOf("name");
  const lines = [["Customer ID", "Customer Name", "Status", "Details"].join(",")];
  result.rows.filter(r => r.status !== "ok").forEach(r =>
    lines.push([r.id, nameCol >= 0 ? r.raw[nameCol] : "", STATUS_LABEL[r.status], r.notes.join("; ")].map(q).join(",")));
  return "﻿" + lines.join("\r\n") + "\r\n";
}

const TallyCore = { ALIASES, FIELDS, STATUS_LABEL, normH, num, normDate, money, signed, show, parseKind,
  workbookRows, csvRows, findHeader, mapCols, PARSERS, runTally, buildWorkbook, buildCsv, stamp };
if (typeof module !== "undefined" && module.exports) module.exports = TallyCore;
root.TallyCore = TallyCore;

/* ===================== Part 2: the page ===================== */
if (typeof document === "undefined") return;

const S = { txn: null, sum: null, input: null };
let RESULT = null, FILTER = "all";
const $ = s => document.querySelector(s);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const tick = () => new Promise(r => setTimeout(r, 30));
function toast(msg) { const t = $("#toast"); t.textContent = msg; t.classList.add("show"); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove("show"), 3500); }

// Library check: show a clear message instead of failing silently
const missingLibs = [];
if (typeof root.Papa === "undefined") missingLibs.push("CSV reader");
if (typeof root.XLSX === "undefined") missingLibs.push("Excel reader");
if (missingLibs.length) {
  const b = $("#libError");
  b.textContent = `The ${missingLibs.join(" and ")} couldn’t load, so ${missingLibs.length > 1 ? "these files" : "some files"} can’t be read yet. ` +
    "Check your internet connection and reload the page. If it keeps happening, your network may be blocking cdnjs.cloudflare.com or cdn.jsdelivr.net; ask IT to allow them.";
  b.hidden = false;
}

async function readRows(file) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  if (ext === "csv") {
    if (typeof root.Papa === "undefined") throw new Error("The CSV reader didn’t load. Check your connection and reload the page.");
    return csvRows(await file.text());
  }
  if (ext === "xlsx" || ext === "xls") {
    if (typeof root.XLSX === "undefined") throw new Error("The Excel reader didn’t load. Check your connection and reload the page.");
    let wb;
    try { wb = root.XLSX.read(await file.arrayBuffer(), { type: "array" }); }
    catch (e) { throw new Error("This Excel file couldn’t be opened. It may be damaged or password-protected."); }
    return workbookRows(wb);
  }
  throw new Error("Use a .csv, .xlsx or .xls file.");
}

/* ---------- upload boxes ---------- */
document.querySelectorAll(".slot").forEach(el => {
  const slot = el.dataset.slot, input = el.querySelector("input[type=file]");
  input.addEventListener("change", () => { if (input.files[0]) loadSlot(el, input.files[0]); input.value = ""; });
  el.addEventListener("dragover", e => { e.preventDefault(); el.classList.add("drag"); });
  el.addEventListener("dragleave", () => el.classList.remove("drag"));
  el.addEventListener("drop", e => { e.preventDefault(); el.classList.remove("drag"); const f = e.dataTransfer.files[0]; if (f) loadSlot(el, f); });
  el.querySelector(".clear").addEventListener("click", () => {
    S[slot] = null; el.className = "slot"; el.querySelector(".file-name").textContent = "";
    setStatus(el, "", ""); el.querySelector(".clear").hidden = true; updateRun();
  });
});
function setStatus(el, cls, msg) { const s = el.querySelector(".status"); s.className = "status " + cls; s.textContent = msg; }
async function loadSlot(el, file) {
  const slot = el.dataset.slot;
  el.className = "slot"; el.querySelector(".file-name").textContent = file.name;
  setStatus(el, "busy", "Reading the file…"); await tick();
  try {
    const parsed = PARSERS[slot](await readRows(file)); parsed.fileName = file.name; S[slot] = parsed;
    el.classList.add("ok"); setStatus(el, "ok", parsed.info);
  } catch (err) {
    S[slot] = null; el.classList.add("err"); setStatus(el, "err", err.message || "This file couldn’t be read.");
  }
  el.querySelector(".clear").hidden = false; updateRun();
}
function updateRun() {
  const ready = !!S.input && (!!S.sum || !!S.txn);
  $("#run").disabled = !ready;
  $("#runmsg").textContent = ready
    ? (S.sum && S.txn ? "Both master files loaded. They’ll also be checked against each other." : "One master file loaded. Add the other to cross-check the masters too.")
    : (!S.input ? "Add the input file and at least one master file." : "Add at least one master file.");
}

/* ---------- run ---------- */
$("#run").addEventListener("click", async () => {
  const b = $("#run"); b.disabled = true; b.textContent = "Checking…"; await tick();
  try {
    RESULT = runTally(S, $("#tol").value);
    renderResults(); $("#results").hidden = false;
    $("#results").scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "start" });
  } catch (e) { toast("Something went wrong while checking: " + (e.message || e)); }
  b.textContent = "Run tally"; b.disabled = false;
});

/* ---------- results ---------- */
function renderResults() {
  const c = RESULT.counts;
  for (const k in c) $("#c-" + k).textContent = fmtN(c[k]);
  const v = $("#verdict"), wrong = c.bad + c.missing;
  if (wrong) { v.className = "verdict bad"; v.textContent = `${fmtN(wrong)} of ${fmtN(c.all)} customer${c.all === 1 ? "" : "s"} ${wrong === 1 ? "doesn’t" : "don’t"} match.`; }
  else if (c.warn) { v.className = "verdict warn"; v.textContent = `No wrong figures in ${fmtN(c.all)} customers, but ${fmtN(c.warn)} need${c.warn === 1 ? "s" : ""} a look.`; }
  else { v.className = "verdict good"; v.textContent = `All ${fmtN(c.all)} customer${c.all === 1 ? "" : "s"} match the master files.`; }
  const used = [S.txn && S.txn.fileName, S.sum && S.sum.fileName].filter(Boolean).join(" and ");
  let meta = `Checked ${S.input.fileName} against ${used} at ${RESULT.at.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })}. Allowed difference ${money(RESULT.tol)}.`;
  if (RESULT.unchecked.length) meta += ` Not checked (not in the master files): ${RESULT.unchecked.map(k => FIELDS[k].label).join(", ")}.`;
  $("#meta").textContent = meta;
  const mi = RESULT.masterIssues, box = $("#masterNotice");
  if (mi.size) {
    const items = [...mi.entries()].slice(0, 300).map(([id, l]) => `<li><b>${esc(id)}</b>: ${l.map(esc).join("; ")}</li>`).join("");
    box.innerHTML = `<details class="notice"><summary>The master files have problems for ${fmtN(mi.size)} customer${mi.size > 1 ? "s" : ""}. Fix the masters before trusting those results.</summary><ul>${items}</ul>${mi.size > 300 ? "<p>Showing the first 300. The marked Excel file lists them all.</p>" : ""}</details>`;
  } else if (RESULT.bothMasters) {
    box.innerHTML = `<p class="meta" style="margin-top:6px">The transaction master and the customer summary agree with each other for every customer, and every running balance adds up.</p>`;
  } else box.innerHTML = "";
  renderTable();
}
document.querySelectorAll(".chip").forEach(ch => ch.addEventListener("click", () => {
  FILTER = ch.dataset.f; document.querySelectorAll(".chip").forEach(x => x.setAttribute("aria-pressed", String(x === ch))); renderTable();
}));
$("#q").addEventListener("input", () => renderTable());

function renderTable() {
  const inp = S.input, q = normText($("#q").value), nameCol = inp.perCol.indexOf("name");
  let list = RESULT.rows.filter(r => FILTER === "all" || r.status === FILTER);
  if (q) list = list.filter(r => r.id.toLowerCase().includes(q) || (nameCol >= 0 && normText(r.raw[nameCol]).includes(q)));
  const cols = inp.header.map((h, i) => ({ i, label: h == null ? "" : String(h), key: inp.perCol[i] })).filter(c => c.label !== "" || c.i === inp.idCol);
  const isNum = k => k && FIELDS[k] && (FIELDS[k].kind === "money" || FIELDS[k].kind === "int");
  $("#tbl thead").innerHTML = "<tr>" + cols.map(c => `<th scope="col" class="${isNum(c.key) ? "num" : ""}">${esc(c.label)}</th>`).join("") + '<th scope="col">Result</th></tr>';
  const LIMIT = 500;
  const html = list.slice(0, LIMIT).map(r => {
    const tds = cols.map(c => {
      if (c.i === inp.idCol) return `<td class="idcell">${esc(r.id)}</td>`;
      const k = c.key, raw = r.raw[c.i];
      if (!k || !FIELDS[k]) return `<td>${esc(raw == null ? "" : raw)}</td>`;
      const kind = FIELDS[k].kind, cell = r.cells[k], cls = isNum(k) ? "num" : "";
      if (cell && cell.state === "bad") {
        const d = cell.diff != null ? ` <span>(${signed(cell.diff)})</span>` : "";
        return `<td class="${cls} bad"><span class="struck">${esc(show(kind, raw))}</span><span class="fix">${esc(show(kind, cell.mv))}${d}</span></td>`;
      }
      return `<td class="${cls}">${esc(show(kind, raw))}</td>`;
    }).join("");
    const tag = { ok: '<span class="tag ok">Matched</span>', bad: '<span class="tag bad">Mismatch</span>', missing: '<span class="tag missing">Not in master</span>', warn: '<span class="tag warn">Needs a look</span>' }[r.status];
    const notes = r.status === "bad" ? "" : r.notes.map(n => `<span class="note">${esc(n)}</span>`).join("");
    return `<tr class="${r.status === "bad" || r.status === "missing" ? "row-bad" : ""}">${tds}<td>${tag}${notes}</td></tr>`;
  }).join("");
  $("#tbl tbody").innerHTML = html || `<tr><td colspan="${cols.length + 1}" class="empty">No customers match this filter.</td></tr>`;
  const more = $("#more");
  if (list.length > LIMIT) { more.hidden = false; more.textContent = `Showing the first ${LIMIT} of ${fmtN(list.length)}. Search for a Customer ID or name, or download the Excel file to see every row.`; }
  else more.hidden = true;
}

/* ---------- downloads ---------- */
function save(filename, blob) {
  try {
    const url = URL.createObjectURL(blob); const a = document.createElement("a");
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000); toast("Downloaded " + filename);
  } catch (e) { toast("The file couldn’t be downloaded" + (e && e.message ? ": " + e.message : ".")); }
}
$("#dlx").addEventListener("click", () => {
  if (!RESULT) return;
  if (typeof root.XLSX === "undefined") { toast("The Excel writer didn’t load. Reload the page and try again."); return; }
  try {
    const buf = root.XLSX.write(buildWorkbook(S, RESULT), { bookType: "xlsx", type: "array" });
    save(`Tally_Result_${stamp()}.xlsx`, new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  } catch (e) { toast("The Excel file couldn’t be built: " + (e.message || e)); }
});
$("#dlc").addEventListener("click", () => {
  if (!RESULT) return;
  if (RESULT.counts.all === RESULT.counts.ok) { toast("Nothing to list: every customer matched."); return; }
  save(`Tally_Mismatches_${stamp()}.csv`, new Blob([buildCsv(S, RESULT)], { type: "text/csv;charset=utf-8" }));
});
updateRun();
})(typeof window !== "undefined" ? window : globalThis);
