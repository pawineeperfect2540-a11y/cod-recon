/* ============================================================
   ระบบเช็คยอด COD & บัตรเครดิต — Perfect Group Intertrading
   ============================================================ */

// ---------- ค่าเริ่มต้น: จับคู่ขนส่ง/ประเภท กับบัญชีธนาคาร ----------
const DEFAULT_MAPPING = {
  best: { label: "Best Express", account: "5903035196", expectedDays: 1, tag: "tag-best" },
  jnt:  { label: "J&T Express",  account: "5903035201", expectedDays: 1, tag: "tag-jnt" },
  kerry:{ label: "Kerry (KEX)",  account: "7622505234", expectedDays: 1, tag: "tag-kerry" },
  cc:   { label: "บัตรเครดิต (PaySolution)", account: "7622505234", expectedDays: 2, tag: "tag-cc" },
};

function loadMapping() {
  try {
    const saved = JSON.parse(localStorage.getItem("cod_mapping"));
    if (saved) return Object.assign({}, DEFAULT_MAPPING, saved);
  } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_MAPPING));
}
let MAPPING = loadMapping();

// วันหยุดธนาคารทางการปี 2569 (พ.ศ.) ตามประกาศธนาคารแห่งประเทศไทย (ข้อมูล ณ 29 มิ.ย. 2569)
// ใช้เป็นค่าเริ่มต้นเท่านั้น — แก้ไข/ลบ/เพิ่มเองได้ในแท็บ "ตั้งค่า"
// ควรตรวจสอบกับประกาศทางการอีกครั้งเป็นระยะ เพราะ ธปท. อาจประกาศวันหยุดพิเศษเพิ่มเติมภายหลัง
const DEFAULT_HOLIDAYS_2026 = [
  { date: "2026-01-01", name: "วันขึ้นปีใหม่" },
  { date: "2026-01-02", name: "วันหยุดทำการเพิ่มเป็นกรณีพิเศษ" },
  { date: "2026-03-03", name: "วันมาฆบูชา" },
  { date: "2026-04-06", name: "วันจักรี" },
  { date: "2026-04-13", name: "วันสงกรานต์" },
  { date: "2026-04-14", name: "วันสงกรานต์" },
  { date: "2026-04-15", name: "วันสงกรานต์" },
  { date: "2026-05-01", name: "วันแรงงานแห่งชาติ" },
  { date: "2026-05-04", name: "วันฉัตรมงคล" },
  { date: "2026-06-01", name: "ชดเชยวันวิสาขบูชา" },
  { date: "2026-06-03", name: "วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าสุทิดาฯ" },
  { date: "2026-07-28", name: "วันเฉลิมพระชนมพรรษา ร.10" },
  { date: "2026-07-29", name: "วันอาสาฬหบูชา" },
  { date: "2026-08-12", name: "วันแม่แห่งชาติ" },
  { date: "2026-10-13", name: "วันนวมินทรมหาราช" },
  { date: "2026-10-16", name: "วันหยุดทำการเพิ่มเติมเป็นกรณีพิเศษ" },
  { date: "2026-10-23", name: "วันปิยมหาราช" },
  { date: "2026-12-07", name: "ชดเชยวันชาติ/วันพ่อแห่งชาติ" },
  { date: "2026-12-10", name: "วันรัฐธรรมนูญ" },
  { date: "2026-12-31", name: "วันสิ้นปี" },
];

function loadHolidays() {
  try {
    const saved = JSON.parse(localStorage.getItem("cod_holidays"));
    if (Array.isArray(saved) && saved.length && typeof saved[0] === "object") return saved;
    if (Array.isArray(saved) && saved.length && typeof saved[0] === "string") {
      // ข้อมูลเก่าเป็น array ของสตริงวันที่ล้วน — แปลงให้มีชื่อกำกับ
      return saved.map(iso => ({ date: iso, name: "วันหยุด" }));
    }
  } catch (e) {}
  return DEFAULT_HOLIDAYS_2026.map(h => ({ ...h }));
}
let HOLIDAYS = loadHolidays(); // array of {date:"YYYY-MM-DD", name}
function isoDate(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
function holidayName(d) {
  const iso = isoDate(d);
  const h = HOLIDAYS.find(x => x.date === iso);
  return h ? h.name : null;
}
function isHoliday(d) {
  return holidayName(d) !== null;
}
function isBusinessDay(d) {
  const dow = d.getDay(); // 0 = Sun, 6 = Sat
  if (dow === 0 || dow === 6) return false;
  if (isHoliday(d)) return false;
  return true;
}
function nextBusinessDay(d) {
  let cur = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  while (!isBusinessDay(cur)) cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
  return cur;
}
function addDays(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

// ---------- state ----------
const uploaded = { courierFiles: { best: [], jnt: [], kerry: [] }, statementFiles: [], ccFile: null };
let lastResults = []; // [{courierKey,label,expectedDate,actualDate,lateDays,amount,count,status}]
let lastDepositsByAccount = {}; // account -> [{date,amount,description,used}] — full bank ledger for the two-side summary

// ================= Utility: date parsing =================
function excelSerialToDate(n) {
  // Excel date serial (1900 system)
  const utc_days = Math.floor(n - 25569);
  const utc_value = utc_days * 86400;
  return new Date(utc_value * 1000);
}
function parseDateCell(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") return excelSerialToDate(v);
  const s = String(v).trim();
  if (!s) return null;
  // YYYY/MM/DD or YYYY-MM-DD
  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  // DD-MM-YYYY or DD/MM/YYYY
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(s);
  if (!isNaN(d)) return d;
  return null;
}
function fmtDate(d) {
  if (!d) return "-";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}
function dayDiff(a, b) {
  const A = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const B = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((B - A) / 86400000);
}
function normAcct(s) {
  return String(s || "").replace(/[^0-9]/g, "");
}
function toNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}
function fmtMoney(n) {
  return Number(n || 0).toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ================= Sheet reading helper =================
function sheetToRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}
function findHeaderRow(rows, mustHaveCols) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(c => (c == null ? "" : String(c).trim()));
    if (mustHaveCols.every(col => row.includes(col))) return i;
  }
  return -1;
}
function rowsToObjects(rows, headerIdx) {
  const headers = rows[headerIdx].map(c => (c == null ? "" : String(c).trim()));
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c === null || c === "")) continue;
    const obj = {};
    headers.forEach((h, idx) => { if (h) obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}

// ================= Parsers: courier COD files =================
function parseKerryWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = sheetToRows(ws);
  // find remit date from header block
  let remitDate = null;
  for (const row of rows) {
    for (const cell of row) {
      if (typeof cell === "string" && cell.includes("Remit Date")) {
        const m = cell.match(/Remit Date:\s*([\d\/\-]+)/);
        if (m) remitDate = parseDateCell(m[1]);
      }
    }
  }
  const headerIdx = findHeaderRow(rows, ["Consignment No.", "Net COD Remit"]);
  if (headerIdx === -1) throw new Error("ไม่พบตารางข้อมูลในไฟล์ Kerry (โครงสร้างไฟล์อาจเปลี่ยนไป)");
  const objs = rowsToObjects(rows, headerIdx);
  const items = [];
  for (const o of objs) {
    if (!o["Consignment No."] || String(o["Consignment No."]).toLowerCase() === "total") continue;
    const amount = toNum(o["Net COD Remit"]) || 0;
    if (!amount) continue;
    items.push({
      waybill: o["Consignment No."],
      recipient: o["Recipient Name"] || "",
      signDate: parseDateCell(o["COD Collected Date"] || o["POD Date"]),
      remitDate: remitDate,
      amount: amount,
    });
  }
  return items;
}

function parseBestWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = sheetToRows(ws);
  const headerIdx = findHeaderRow(rows, ["Waybill Number", "COD Amount"]);
  if (headerIdx === -1) throw new Error("ไม่พบตารางข้อมูลในไฟล์ Best (โครงสร้างไฟล์อาจเปลี่ยนไป)");
  const objs = rowsToObjects(rows, headerIdx);
  const items = [];
  for (const o of objs) {
    if (!o["Waybill Number"] || String(o["Waybill Number"]).toLowerCase() === "total") continue;
    const amount = toNum(o["COD Amount"]) || 0;
    if (!amount) continue;
    items.push({
      waybill: o["Waybill Number"],
      recipient: o["Recipient Name"] || "",
      signDate: parseDateCell(o["Sign DT"]),
      remitDate: parseDateCell(o["Remittance DT"]),
      amount: amount,
    });
  }
  return items;
}

// J&T format ไม่มีตัวอย่างจริงในตอนนี้ — ลองใช้โครงสร้างแบบ Best เป็นค่าเริ่มต้น
// (คอลัมน์ Waybill/Sign DT/Remittance DT/Amount) ถ้าไฟล์จริงต่างจากนี้ ระบบจะแจ้งเตือนให้ปรับ parser
function parseJntWorkbook(wb) {
  try {
    return parseBestWorkbook(wb);
  } catch (e) {
    throw new Error("ยังไม่รู้จักโครงสร้างไฟล์ J&T — กรุณาแจ้งตัวอย่างไฟล์จริงเพื่อปรับระบบ (ตอนนี้ยังไม่มีข้อมูลตัวอย่าง)");
  }
}

const COURIER_PARSERS = { kerry: parseKerryWorkbook, best: parseBestWorkbook, jnt: parseJntWorkbook };

// ================= Parser: bank statement =================
function parseStatementWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = sheetToRows(ws);
  const headerIdx = findHeaderRow(rows, ["Account Number", "Deposit"]);
  if (headerIdx === -1) throw new Error("ไม่พบตารางข้อมูลใน statement (โครงสร้างไฟล์อาจเปลี่ยนไป)");
  const objs = rowsToObjects(rows, headerIdx);
  const byAccount = {};
  for (const o of objs) {
    const acct = normAcct(o["Account Number"]);
    if (!acct) continue;
    const deposit = toNum(o["Deposit"]) || 0;
    if (!deposit) continue;
    if (!byAccount[acct]) byAccount[acct] = [];
    byAccount[acct].push({
      date: parseDateCell(o["Date"]),
      amount: deposit,
      description: o["Description"] || "",
      used: false,
    });
  }
  return byAccount;
}

// ================= Parser: PaySolution credit card =================
function parseCCWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = sheetToRows(ws);
  const headerIdx = findHeaderRow(rows, ["ReferenceNo", "Net"]);
  if (headerIdx === -1) throw new Error("ไม่พบตารางข้อมูลในไฟล์บัตรเครดิต (โครงสร้างไฟล์อาจเปลี่ยนไป)");
  const objs = rowsToObjects(rows, headerIdx);
  const items = [];
  for (const o of objs) {
    if (!o["ReferenceNo"]) continue;
    if (String(o["StatusName"] || "").toLowerCase().includes("cancel")) continue;
    const net = toNum(o["Net"]) || 0;
    if (!net) continue;
    items.push({
      waybill: o["ReferenceNo"],
      recipient: o["ProductDetail"] || "",
      signDate: parseDateCell(o["paymentDate"] || o["OrderDateTime"]),
      remitDate: parseDateCell(o["paymentDate"] || o["OrderDateTime"]),
      amount: net,
    });
  }
  return items;
}

// ================= Matching engine =================
function dedupeByWaybill(items) {
  const seen = new Set();
  const out = [];
  for (const it of items) {
    const key = String(it.waybill).trim();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(it);
  }
  return out;
}

function groupIntoBatches(items) {
  const map = {};
  for (const it of items) {
    if (!it.remitDate) continue;
    const key = it.remitDate.toDateString();
    if (!map[key]) map[key] = { date: it.remitDate, amount: 0, count: 0, items: [] };
    map[key].amount += it.amount;
    map[key].count += 1;
    map[key].items.push(it);
  }
  return Object.values(map).sort((a, b) => a.date - b.date);
}

function matchBatchesAgainstDeposits(batches, deposits, expectedDays, label, tolerance = 1) {
  const results = [];
  for (const batch of batches) {
    // ปรับวันที่คาดว่าเงินเข้าให้เลี่ยงเสาร์-อาทิตย์-วันหยุดนักขัตฤกษ์ (ถ้าตกวันหยุด ให้เลื่อนไปวันทำการถัดไป)
    const rawExpected = addDays(batch.date, expectedDays);
    const adjustedExpected = nextBusinessDay(rawExpected);

    let found = null;
    let candidates = deposits.filter(d => !d.used && Math.abs(d.amount - batch.amount) <= tolerance);
    // window: from batch date -1 to +30 days
    candidates = candidates.filter(d => {
      const diff = dayDiff(batch.date, d.date);
      return diff >= -1 && diff <= 30;
    });
    candidates.sort((a, b) => a.date - b.date);
    if (candidates.length) found = candidates[0];

    if (found) {
      found.used = true;
      found.matchedLabel = label;
      const isLate = dayDiff(adjustedExpected, found.date) > 0;
      const lateDays = isLate ? dayDiff(adjustedExpected, found.date) : 0;
      results.push({
        expectedDate: adjustedExpected,
        rawExpectedDate: batch.date,
        actualDate: found.date,
        lateDays: lateDays,
        amount: batch.amount,
        count: batch.count,
        status: isLate ? "late" : "ontime",
        items: batch.items,
      });
    } else {
      results.push({
        expectedDate: adjustedExpected,
        rawExpectedDate: batch.date,
        actualDate: null,
        lateDays: null,
        amount: batch.amount,
        count: batch.count,
        status: "missing",
        items: batch.items,
      });
    }
  }
  return results;
}

// ================= Main run =================
async function runReconciliation() {
  const logEl = document.getElementById("runLog");
  const statusEl = document.getElementById("runStatus");
  logEl.style.display = "block";
  logEl.textContent = "";
  const log = (msg) => { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; };

  try {
    statusEl.textContent = "กำลังอ่านไฟล์...";

    // parse statements -> deposits by account
    let depositsByAccount = {};
    for (const f of uploaded.statementFiles) {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const parsed = parseStatementWorkbook(wb);
      for (const acct in parsed) {
        if (!depositsByAccount[acct]) depositsByAccount[acct] = [];
        depositsByAccount[acct] = depositsByAccount[acct].concat(parsed[acct]);
      }
      log(`✓ อ่าน statement: ${f.name}`);
    }

    lastResults = [];
    lastDepositsByAccount = depositsByAccount;
    const order = ["best", "jnt", "kerry", "cc"];
    for (const key of order) {
      const cfg = MAPPING[key];
      const acct = normAcct(cfg.account);
      const deposits = depositsByAccount[acct] || [];

      let items = [];
      if (key === "cc") {
        if (uploaded.ccFile) {
          const wb = XLSX.read(await uploaded.ccFile.arrayBuffer(), { type: "array" });
          items = parseCCWorkbook(wb);
          log(`✓ อ่านไฟล์บัตรเครดิต: ${uploaded.ccFile.name} (${items.length} รายการ)`);
        }
      } else {
        const files = uploaded.courierFiles[key] || [];
        for (const f of files) {
          const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
          const parsed = COURIER_PARSERS[key](wb);
          items = items.concat(parsed);
          log(`✓ อ่านไฟล์ ${cfg.label}: ${f.name} (${parsed.length} รายการ)`);
        }
      }
      if (!items.length) continue;

      const beforeDedup = items.length;
      items = dedupeByWaybill(items);
      if (items.length < beforeDedup) {
        log(`⚠ ${cfg.label}: พบเลขพัสดุ/เลขอ้างอิงซ้ำ ${beforeDedup - items.length} รายการ (จากไฟล์ที่ช่วงวันที่ทับกัน) — ตัดออกให้แล้ว นับครั้งเดียวต่อ 1 พัสดุ`);
      }

      const batches = key === "kerry" ? groupIntoBatches(items) : groupIntoBatchesByDay(items);
      const matched = matchBatchesAgainstDeposits(batches, deposits, cfg.expectedDays, cfg.label);
      matched.forEach(m => lastResults.push(Object.assign({ courierKey: key, label: cfg.label, tag: cfg.tag }, m)));

      if (!depositsByAccount[acct]) log(`⚠ ไม่พบ statement ของบัญชี ${cfg.account} (${cfg.label}) — ยังไม่ได้อัปโหลด`);
    }

    statusEl.textContent = `เสร็จสิ้น — พบ ${lastResults.length} รอบโอน`;
    log(`\nเสร็จสิ้นการประมวลผล: ${lastResults.length} รอบโอนทั้งหมด`);
    renderResults();
    renderTwoSideSummary();
    let firebaseFailed = false;
    if (window.FIREBASE_ENABLED && db) {
      log("กำลังบันทึกประวัติลง Firebase...");
      const saveResult = await saveResultsToFirebase();
      if (saveResult.failed > 0) {
        firebaseFailed = true;
        log(`✗ บันทึกไม่สำเร็จ ${saveResult.failed} รายการ — ${saveResult.errors.join("; ")}`);
        log("  (สาเหตุที่พบบ่อย: Firestore Security Rules ยังไม่อนุญาตให้เขียนข้อมูล — เช็คได้ในแท็บ \"ตั้งค่า\")");
      } else {
        log(`✓ บันทึกประวัติลง Firebase สำเร็จ ${saveResult.ok} รายการ (ดูได้ในแท็บ "ประวัติย้อนหลัง")`);
      }
    } else {
      log("⚠ ยังไม่ได้เชื่อมต่อ Firebase — ผลลัพธ์รอบนี้จะไม่ถูกบันทึกไว้ดูย้อนหลัง");
    }
    if (!firebaseFailed) switchView("results");
  } catch (err) {
    statusEl.textContent = "เกิดข้อผิดพลาด";
    log("✗ " + err.message);
    console.error(err);
  }
}

function groupIntoBatchesByDay(items) {
  // same as groupIntoBatches but kept separate name for clarity (per-row remitDate already set)
  return groupIntoBatches(items);
}

// ================= Render results =================
// ================= Two-side executive summary =================
function accountGroups() {
  const groups = {};
  Object.keys(MAPPING).forEach(key => {
    const cfg = MAPPING[key];
    const acct = normAcct(cfg.account);
    if (!groups[acct]) groups[acct] = { labels: [], keys: [] };
    groups[acct].labels.push(cfg.label);
    groups[acct].keys.push(key);
  });
  return groups;
}

const THAI_WEEKDAY = ["วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์"];
function nonBusinessLabel(d) {
  const hn = holidayName(d);
  if (hn) return hn;
  const dow = d.getDay();
  if (dow === 0 || dow === 6) return THAI_WEEKDAY[dow];
  return null;
}

function renderTwoSideSummary(fromStr, toStr) {
  const el = document.getElementById("twosideContent");
  if (!Object.keys(lastDepositsByAccount).length && !lastResults.length) {
    el.innerHTML = `<div class="empty">ยังไม่มีข้อมูล — ไปที่แท็บ "อัปโหลดข้อมูล" แล้วกดประมวลผลก่อนค่ะ</div>`;
    return;
  }
  const fromD = fromStr ? new Date(fromStr) : null;
  const toD = toStr ? new Date(toStr) : null;
  const inRange = (d) => (!fromD || d >= fromD) && (!toD || d <= toD);

  const groups = accountGroups();
  let html = "";
  for (const acct in groups) {
    const { labels } = groups[acct];
    const deposits = (lastDepositsByAccount[acct] || []).filter(d => inRange(d.date)).slice().sort((a, b) => a.date - b.date);
    const missingBatches = lastResults.filter(r => normAcct(MAPPING[r.courierKey].account) === acct && r.status === "missing" && inRange(r.rawExpectedDate));

    html += `<div class="section-title">บัญชี ${acct} (${labels.join(" / ")})</div>`;
    html += `<div class="card" style="padding:0;overflow-x:auto;margin-bottom:14px;"><table>
      <thead><tr><th>วันที่เงินเข้าบัญชี</th><th>ยอดจากบัญชี</th><th>ตรง?</th><th>หมายเหตุ / ที่มา</th></tr></thead><tbody>`;

    if (fromD && toD) {
      // แสดงครบทุกวันในช่วงที่เลือก — วันหยุด/เสาร์-อาทิตย์ที่ไม่มีเงินเข้าจะขึ้นชื่อวันหยุดแทนที่จะหายไปจากตาราง
      const byDay = {};
      deposits.forEach(d => {
        const key = isoDate(d.date);
        if (!byDay[key]) byDay[key] = [];
        byDay[key].push(d);
      });
      for (let cur = new Date(fromD); cur <= toD; cur = addDays(cur, 1)) {
        const key = isoDate(cur);
        const dayDeposits = byDay[key];
        if (dayDeposits && dayDeposits.length) {
          dayDeposits.forEach(d => {
            const ok = d.used;
            html += `<tr>
              <td>${fmtDate(d.date)}</td>
              <td>${fmtMoney(d.amount)}</td>
              <td>${ok ? '<span class="badge ontime">✓</span>' : '<span class="badge missing">✕</span>'}</td>
              <td>${ok ? `ตรงกับ ${d.matchedLabel || "-"}` : `ยังไม่ได้อัปโหลดไฟล์ COD ของ ${labels.join("/")} ที่ตรงกับยอดนี้ (หรือไม่ใช่ยอด COD) — อ้างอิงจาก statement: ${d.description || "-"}`}</td>
            </tr>`;
          });
        } else {
          const nb = nonBusinessLabel(cur);
          if (nb) {
            html += `<tr>
              <td>${fmtDate(cur)}</td>
              <td style="color:var(--text-faint);">-</td>
              <td><span class="badge" style="background:rgba(255,255,255,.06);color:var(--text-faint);">-</span></td>
              <td style="color:var(--text-faint);">${nb}</td>
            </tr>`;
          } else {
            html += `<tr>
              <td>${fmtDate(cur)}</td>
              <td style="color:var(--text-faint);">-</td>
              <td><span class="badge missing">✕</span></td>
              <td>ไม่มีเงินเข้าวันนี้ (เป็นวันทำการปกติ)</td>
            </tr>`;
          }
        }
      }
    } else {
      if (!deposits.length) {
        html += `<tr><td colspan="4" style="color:var(--text-faint);">ไม่มีข้อมูลในช่วงวันที่นี้ (หรือยังไม่ได้อัปโหลด statement ของบัญชีนี้)</td></tr>`;
      }
      for (const d of deposits) {
        const ok = d.used;
        html += `<tr>
          <td>${fmtDate(d.date)}</td>
          <td>${fmtMoney(d.amount)}</td>
          <td>${ok ? '<span class="badge ontime">✓</span>' : '<span class="badge missing">✕</span>'}</td>
          <td>${ok ? `ตรงกับ ${d.matchedLabel || "-"}` : `ยังไม่ได้อัปโหลดไฟล์ COD ของ ${labels.join("/")} ที่ตรงกับยอดนี้ (หรือไม่ใช่ยอด COD) — อ้างอิงจาก statement: ${d.description || "-"}`}</td>
        </tr>`;
      }
    }
    html += `</tbody></table></div>`;

    if (missingBatches.length) {
      html += `<div class="config-note" style="margin-bottom:22px;">`;
      missingBatches.forEach(m => {
        html += `⚠ ${m.label} แจ้งว่าจะโอน ${fmtMoney(m.amount)} บาท (${m.count} รายการ) วันที่ ${fmtDate(m.rawExpectedDate)} — ยังไม่พบเงินเข้าบัญชีนี้<br>`;
      });
      html += `</div>`;
    } else {
      html += `<div style="color:var(--green);font-size:13px;margin-bottom:22px;">✓ ยอดที่ขนส่ง/บัตรเครดิตแจ้งไว้ทั้งหมดในบัญชีนี้ เข้าครบแล้ว</div>`;
    }
  }
  el.innerHTML = html;
}

function exportTwoSideExcel(fromStr, toStr) {
  const fromD = fromStr ? new Date(fromStr) : null;
  const toD = toStr ? new Date(toStr) : null;
  const inRange = (d) => (!fromD || d >= fromD) && (!toD || d <= toD);

  const groups = accountGroups();
  const wb = XLSX.utils.book_new();
  for (const acct in groups) {
    const { labels } = groups[acct];
    const deposits = (lastDepositsByAccount[acct] || []).filter(d => inRange(d.date)).slice().sort((a, b) => a.date - b.date);
    const data = [["วันที่", "ยอดจากบัญชี", "ตรง?", "หมายเหตุ / ที่มา"]];
    if (fromD && toD) {
      const byDay = {};
      deposits.forEach(d => {
        const key = isoDate(d.date);
        if (!byDay[key]) byDay[key] = [];
        byDay[key].push(d);
      });
      for (let cur = new Date(fromD); cur <= toD; cur = addDays(cur, 1)) {
        const dayDeposits = byDay[isoDate(cur)];
        if (dayDeposits && dayDeposits.length) {
          dayDeposits.forEach(d => {
            data.push([fmtDate(d.date), d.amount, d.used ? "✓" : "✕", d.used ? `ตรงกับ ${d.matchedLabel || "-"}` : `ยังไม่ได้อัปโหลดไฟล์ COD ของ ${labels.join("/")} ที่ตรงกับยอดนี้ (หรือไม่ใช่ยอด COD) — อ้างอิง: ${d.description || "-"}`]);
          });
        } else {
          const nb = nonBusinessLabel(cur);
          data.push([fmtDate(cur), "-", nb ? "-" : "✕", nb || "ไม่มีเงินเข้าวันนี้ (เป็นวันทำการปกติ)"]);
        }
      }
    } else {
      deposits.forEach(d => {
        data.push([fmtDate(d.date), d.amount, d.used ? "✓" : "✕", d.used ? `ตรงกับ ${d.matchedLabel || "-"}` : `ยังไม่ได้อัปโหลดไฟล์ COD ของ ${labels.join("/")} ที่ตรงกับยอดนี้ (หรือไม่ใช่ยอด COD) — อ้างอิง: ${d.description || "-"}`]);
      });
    }
    const missingBatches = lastResults.filter(r => normAcct(MAPPING[r.courierKey].account) === acct && r.status === "missing" && inRange(r.rawExpectedDate));
    if (missingBatches.length) {
      data.push([]);
      data.push(["ยอดที่ขนส่งแจ้งแต่ยังไม่พบในบัญชีนี้"]);
      missingBatches.forEach(m => data.push([fmtDate(m.rawExpectedDate), m.amount, "✕", `${m.label} (${m.count} รายการ)`]));
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const sheetName = (acct + " " + labels.join("_")).slice(0, 31).replace(/[\\/?*[\]:]/g, "");
    XLSX.utils.book_append_sheet(wb, ws, sheetName || acct);
  }
  XLSX.writeFile(wb, `สรุปผู้บริหาร-2ฝั่ง-${new Date().toISOString().slice(0,10)}.xlsx`);
}

function renderResults() {
  const body = document.getElementById("resultsBody");
  const empty = document.getElementById("resultsEmpty");
  const statsEl = document.getElementById("resultStats");
  const filter = document.querySelector(".pill.active")?.dataset.filter || "all";

  const rows = lastResults.filter(r => filter === "all" || r.status === filter)
    .sort((a, b) => b.expectedDate - a.expectedDate);

  body.innerHTML = "";
  empty.style.display = rows.length ? "none" : "block";
  if (!rows.length) {
    empty.textContent = lastResults.length
      ? `ไม่มีรายการที่ตรงกับตัวกรองนี้ (มีทั้งหมด ${lastResults.length} รอบโอนในสถานะอื่น — ลองกด "ทั้งหมด")`
      : `ยังไม่มีผลการตรวจสอบ — ไปที่แท็บ "อัปโหลดข้อมูล" เพื่อเริ่มค่ะ`;
  }

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.className = "result-row";
    tr.style.cursor = "pointer";
    const statusLabel = { ontime: "ตรงเวลา", late: "โอนช้า", missing: "ยังไม่พบ" }[r.status];
    tr.innerHTML = `
      <td><span class="tag-courier ${r.tag}">${r.label}</span> <span style="color:var(--text-faint);font-size:11px;">▸</span></td>
      <td>${fmtDate(r.rawExpectedDate)}</td>
      <td>${fmtDate(r.expectedDate)}</td>
      <td>${r.actualDate ? fmtDate(r.actualDate) : "-"}</td>
      <td>${r.lateDays != null ? (r.lateDays > 0 ? r.lateDays + " วัน" : "-") : "-"}</td>
      <td>${fmtMoney(r.amount)}</td>
      <td>${r.count}</td>
      <td><span class="badge ${r.status}">${statusLabel}</span></td>
    `;
    body.appendChild(tr);

    const detailTr = document.createElement("tr");
    detailTr.className = "detail-row";
    detailTr.style.display = "none";
    const items = r.items || [];
    const itemRows = items.map(it => `
      <tr>
        <td style="padding-left:26px;">${it.waybill}</td>
        <td>${it.recipient || "-"}</td>
        <td>${it.signDate ? fmtDate(it.signDate) : "-"}</td>
        <td>${fmtMoney(it.amount)}</td>
      </tr>`).join("");
    detailTr.innerHTML = `
      <td colspan="8" style="background:var(--bg-panel2);padding:10px 16px;">
        <table style="width:100%;">
          <thead><tr><th style="padding-left:26px;">เลขพัสดุ</th><th>ผู้รับ</th><th>วันที่เซ็นรับ</th><th>ยอด</th></tr></thead>
          <tbody>${itemRows}</tbody>
        </table>
      </td>
    `;
    body.appendChild(detailTr);

    tr.addEventListener("click", () => {
      const showing = detailTr.style.display !== "none";
      detailTr.style.display = showing ? "none" : "table-row";
      tr.querySelector("span[style*='faint']").textContent = showing ? "▸" : "▾";
    });
  });

  const total = lastResults.length;
  const ontime = lastResults.filter(r => r.status === "ontime").length;
  const late = lastResults.filter(r => r.status === "late").length;
  const missing = lastResults.filter(r => r.status === "missing").length;
  const missingAmt = lastResults.filter(r => r.status === "missing").reduce((s, r) => s + r.amount, 0);

  statsEl.innerHTML = `
    <div class="card"><h3>รอบโอนทั้งหมด</h3><div class="hint">จากไฟล์ที่อัปโหลด</div><div class="stat teal">${total}</div></div>
    <div class="card"><h3>ตรงเวลา</h3><div class="hint">เข้าไม่เกินกำหนด</div><div class="stat green">${ontime}</div></div>
    <div class="card"><h3>โอนช้า</h3><div class="hint">เข้าช้ากว่ากำหนด</div><div class="stat amber">${late}</div></div>
    <div class="card"><h3>ยังไม่พบ</h3><div class="hint">รวม ${fmtMoney(missingAmt)} บาท — ต้องตามขนส่ง</div><div class="stat red">${missing}</div></div>
  `;
}

// ================= Export to Excel =================
function exportResultsExcel(rows, filename) {
  const data = [["ขนส่ง/ประเภท", "วันที่ขนส่งแจ้งจะโอน", "วันที่ครบกำหนด (เว้นวันหยุด)", "วันที่เข้าจริง", "จำนวนวันช้า", "ยอดเงิน", "จำนวนรายการ", "สถานะ"]];
  const statusLabel = { ontime: "ตรงเวลา", late: "โอนช้า", missing: "ยังไม่พบ" };
  rows.forEach(r => {
    data.push([
      r.label, fmtDate(r.rawExpectedDate), fmtDate(r.expectedDate), r.actualDate ? fmtDate(r.actualDate) : "-",
      r.lateDays != null && r.lateDays > 0 ? r.lateDays : "-",
      r.amount, r.count, statusLabel[r.status],
    ]);
  });
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "สรุป COD");

  // sheet 2: itemized waybill detail (only for batches with items and status late/missing take priority but include all)
  const detailData = [["ขนส่ง/ประเภท", "สถานะรอบโอน", "เลขพัสดุ", "ผู้รับ", "วันที่เซ็นรับ", "ยอด"]];
  rows.forEach(r => {
    (r.items || []).forEach(it => {
      detailData.push([r.label, statusLabel[r.status], it.waybill, it.recipient || "-", it.signDate ? fmtDate(it.signDate) : "-", it.amount]);
    });
  });
  if (detailData.length > 1) {
    const ws2 = XLSX.utils.aoa_to_sheet(detailData);
    XLSX.utils.book_append_sheet(wb, ws2, "รายละเอียดพัสดุ");
  }

  XLSX.writeFile(wb, filename);
}

// ================= Firebase =================
let db = null;
function initFirebase() {
  if (!window.FIREBASE_ENABLED || !window.FIREBASE_CONFIG || !window.FIREBASE_CONFIG.apiKey) {
    document.getElementById("fbWarning").style.display = "block";
    document.getElementById("fbStatus").textContent = "";
    return;
  }
  try {
    firebase.initializeApp(window.FIREBASE_CONFIG);
    db = firebase.firestore();
    document.getElementById("fbStatus").textContent = "🟢 เชื่อมต่อ Firebase แล้ว";
  } catch (e) {
    console.error(e);
    document.getElementById("fbWarning").style.display = "block";
  }
}

function monthKeyLocal(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
}

async function saveResultsToFirebase() {
  if (!db) return { ok: 0, failed: 0, errors: [] };
  const batch = {};
  for (const r of lastResults) {
    const period = monthKeyLocal(r.rawExpectedDate || r.expectedDate);
    const id = `${r.courierKey}_${period}`;
    if (!batch[id]) batch[id] = { courierKey: r.courierKey, label: r.label, period, items: [], savedAt: new Date().toISOString() };
    batch[id].items.push(r);
  }
  let ok = 0, failed = 0, errors = [];
  for (const id in batch) {
    try {
      await db.collection("cod_periods").doc(id).set(batch[id]);
      ok++;
    } catch (e) {
      failed++;
      errors.push(`${id}: ${e.message}`);
      console.error("save failed", id, e);
    }
  }
  return { ok, failed, errors };
}

async function loadHistoryForMonth(monthStr) {
  const body = document.getElementById("historyBody");
  const empty = document.getElementById("historyEmpty");
  body.innerHTML = "";
  if (!db) {
    empty.textContent = "ยังไม่ได้เชื่อมต่อ Firebase — ดูวิธีตั้งค่าในแท็บ \"ตั้งค่า\"";
    empty.style.display = "block";
    return;
  }
  let snap;
  try {
    snap = await db.collection("cod_periods").where("period", "==", monthStr).get();
  } catch (e) {
    empty.textContent = `โหลดข้อมูลไม่สำเร็จ: ${e.message} (ส่วนใหญ่เกิดจาก Firestore Security Rules ยังไม่อนุญาตให้อ่านข้อมูล collection "cod_periods")`;
    empty.style.display = "block";
    console.error(e);
    return;
  }
  if (snap.empty) {
    empty.textContent = "ไม่พบข้อมูลย้อนหลังของเดือนนี้";
    empty.style.display = "block";
    return;
  }
  empty.style.display = "none";
  window.__historyRows = [];
  snap.forEach(doc => {
    const d = doc.data();
    const ontime = d.items.filter(i => i.status === "ontime").length;
    const late = d.items.filter(i => i.status === "late").length;
    const missing = d.items.filter(i => i.status === "missing").length;
    const totalAmt = d.items.reduce((s, i) => s + i.amount, 0);
    window.__historyRows.push(...d.items.map(i => Object.assign({ label: d.label, courierKey: d.courierKey, tag: MAPPING[d.courierKey]?.tag }, i, { expectedDate: new Date(i.expectedDate), actualDate: i.actualDate ? new Date(i.actualDate) : null })));
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${new Date(d.savedAt).toLocaleString("th-TH")}</td>
      <td><span class="tag-courier ${MAPPING[d.courierKey]?.tag || ""}">${d.label}</span></td>
      <td>${fmtMoney(totalAmt)}</td>
      <td>${ontime}</td>
      <td>${late}</td>
      <td>${missing}</td>
      <td></td>
    `;
    body.appendChild(tr);
  });
}

// ================= UI wiring =================
function switchView(name) {
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  document.getElementById("view-" + name).classList.add("active");
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.view === name));
}

function buildCourierUploadGrid() {
  const grid = document.getElementById("courierUploadGrid");
  grid.innerHTML = "";
  ["best", "jnt", "kerry"].forEach(key => {
    const cfg = MAPPING[key];
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <h3><span class="tag-courier ${cfg.tag}">${cfg.label}</span></h3>
      <div class="hint">บัญชี SCB ${cfg.account}</div>
      <div class="dropzone" id="dz-${key}">📦 ลากไฟล์ COD มาวาง หรือคลิกเพื่อเลือก (.xlsx ได้หลายไฟล์)
        <input type="file" id="file-${key}" accept=".xlsx,.xls" multiple style="display:none;">
      </div>
      <div class="log" id="log-${key}" style="display:none;"></div>
    `;
    grid.appendChild(card);
    wireDropzone(`dz-${key}`, `file-${key}`, (files) => {
      uploaded.courierFiles[key] = mergeFiles(uploaded.courierFiles[key] || [], files);
      refreshCourierLog(key);
    });
  });
}

function refreshCourierLog(key) {
  const dz = document.getElementById(`dz-${key}`);
  const logEl = document.getElementById(`log-${key}`);
  const list = uploaded.courierFiles[key] || [];
  if (list.length) dz.classList.add("filled"); else dz.classList.remove("filled");
  renderFileLog(logEl, list, (idx) => {
    uploaded.courierFiles[key].splice(idx, 1);
    refreshCourierLog(key);
  });
}

function refreshStatementLog() {
  const dz = document.getElementById("dz-statement");
  const logEl = document.getElementById("statementFileList");
  if (uploaded.statementFiles.length) dz.classList.add("filled"); else dz.classList.remove("filled");
  renderFileLog(logEl, uploaded.statementFiles, (idx) => {
    uploaded.statementFiles.splice(idx, 1);
    refreshStatementLog();
  });
}

function mergeFiles(existing, newFileList) {
  const merged = [...existing];
  Array.from(newFileList).forEach(f => {
    if (!merged.some(e => e.name === f.name && e.size === f.size)) merged.push(f);
  });
  return merged;
}

function renderFileLog(logEl, files, onRemove) {
  logEl.style.display = "block";
  logEl.innerHTML = files.map((f, i) => `
    <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;padding:2px 0;">
      <span>• ${f.name}</span>
      <span data-idx="${i}" style="cursor:pointer;color:var(--red);font-weight:700;flex-shrink:0;">✕</span>
    </div>`).join("");
  logEl.querySelectorAll("[data-idx]").forEach(btn => {
    btn.addEventListener("click", () => onRemove(Number(btn.dataset.idx)));
  });
}

function wireDropzone(dzId, inputId, onFiles) {
  const dz = document.getElementById(dzId);
  const input = document.getElementById(inputId);
  dz.addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    if (input.files.length) onFiles(input.files);
    input.value = "";
  });
  dz.addEventListener("dragover", e => { e.preventDefault(); dz.style.borderColor = "var(--teal)"; });
  dz.addEventListener("dragleave", () => { dz.style.borderColor = ""; });
  dz.addEventListener("drop", e => {
    e.preventDefault();
    dz.style.borderColor = "";
    if (e.dataTransfer.files.length) onFiles(e.dataTransfer.files);
  });
}

function buildMappingTable() {
  const tbody = document.querySelector("#mappingTable tbody");
  tbody.innerHTML = "";
  Object.keys(MAPPING).forEach(key => {
    const cfg = MAPPING[key];
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="tag-courier ${cfg.tag}">${cfg.label}</span></td>
      <td><input type="text" data-key="${key}" data-field="account" value="${cfg.account}" style="width:160px;"></td>
      <td><input type="number" data-key="${key}" data-field="expectedDays" value="${cfg.expectedDays}" style="width:100px;"> วัน</td>
    `;
    tbody.appendChild(tr);
  });
}

function buildHolidayList() {
  const el = document.getElementById("holidayList");
  const sorted = [...HOLIDAYS].sort((a, b) => a.date.localeCompare(b.date));
  if (!sorted.length) {
    el.innerHTML = `<div style="color:var(--text-faint);font-size:13px;">ยังไม่ได้เพิ่มวันหยุดใดๆ</div>`;
    return;
  }
  el.innerHTML = sorted.map(h => {
    const d = new Date(h.date);
    return `<span class="pill" style="display:inline-flex;align-items:center;gap:8px;margin:0 8px 8px 0;">
      ${fmtDate(d)} — ${h.name}
      <span data-remove="${h.date}" style="cursor:pointer;color:var(--red);font-weight:700;">✕</span>
    </span>`;
  }).join("");
  el.querySelectorAll("[data-remove]").forEach(btn => {
    btn.addEventListener("click", () => {
      HOLIDAYS = HOLIDAYS.filter(h => h.date !== btn.dataset.remove);
      localStorage.setItem("cod_holidays", JSON.stringify(HOLIDAYS));
      buildHolidayList();
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initFirebase();
  buildCourierUploadGrid();
  buildMappingTable();
  buildHolidayList();

  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => switchView(btn.dataset.view));
  });

  wireDropzone("dz-statement", "file-statement", (files) => {
    uploaded.statementFiles = mergeFiles(uploaded.statementFiles, files);
    refreshStatementLog();
  });

  wireDropzone("dz-cc", "file-cc", (files) => {
    uploaded.ccFile = files[0];
    const dz = document.getElementById("dz-cc");
    dz.classList.add("filled");
    dz.textContent = "✓ " + files[0].name;
  });

  document.getElementById("btnRun").addEventListener("click", runReconciliation);

  document.querySelectorAll(".pill").forEach(p => {
    p.addEventListener("click", () => {
      document.querySelectorAll(".pill").forEach(x => x.classList.remove("active"));
      p.classList.add("active");
      renderResults();
    });
  });

  document.getElementById("btnExportExcel").addEventListener("click", () => {
    const filter = document.querySelector(".pill.active")?.dataset.filter || "all";
    const rows = lastResults.filter(r => filter === "all" || r.status === filter);
    exportResultsExcel(rows, `สรุปยอด-COD-${new Date().toISOString().slice(0,10)}.xlsx`);
  });

  document.getElementById("btnExportTwoSide").addEventListener("click", () => {
    exportTwoSideExcel(document.getElementById("twosideFrom").value, document.getElementById("twosideTo").value);
  });

  document.querySelector('[data-view="twoside"]').addEventListener("click", () => {
    renderTwoSideSummary(document.getElementById("twosideFrom").value, document.getElementById("twosideTo").value);
  });

  document.getElementById("btnTwoSideFilter").addEventListener("click", () => {
    renderTwoSideSummary(document.getElementById("twosideFrom").value, document.getElementById("twosideTo").value);
  });

  document.getElementById("btnTwoSideLast7").addEventListener("click", () => {
    const to = new Date();
    const from = new Date();
    from.setDate(from.getDate() - 6);
    const toStr = to.toISOString().slice(0, 10);
    const fromStr = from.toISOString().slice(0, 10);
    document.getElementById("twosideFrom").value = fromStr;
    document.getElementById("twosideTo").value = toStr;
    renderTwoSideSummary(fromStr, toStr);
  });

  document.getElementById("btnTwoSideClear").addEventListener("click", () => {
    document.getElementById("twosideFrom").value = "";
    document.getElementById("twosideTo").value = "";
    renderTwoSideSummary();
  });

  document.getElementById("btnLoadHistory").addEventListener("click", () => {
    const val = document.getElementById("historyMonth").value; // YYYY-MM
    if (!val) return;
    loadHistoryForMonth(val);
  });

  document.getElementById("btnExportHistory").addEventListener("click", () => {
    const rows = window.__historyRows || [];
    if (!rows.length) { alert("กรุณาโหลดประวัติของเดือนก่อน"); return; }
    exportResultsExcel(rows, `สรุปผู้บริหาร-COD-${document.getElementById("historyMonth").value}.xlsx`);
  });

  document.getElementById("btnAddHoliday").addEventListener("click", () => {
    const val = document.getElementById("newHolidayDate").value; // YYYY-MM-DD
    const nameInput = document.getElementById("newHolidayName");
    const name = (nameInput.value || "วันหยุดพิเศษ").trim();
    if (!val) return;
    if (!HOLIDAYS.some(h => h.date === val)) {
      HOLIDAYS.push({ date: val, name });
      localStorage.setItem("cod_holidays", JSON.stringify(HOLIDAYS));
      buildHolidayList();
    }
    document.getElementById("newHolidayDate").value = "";
    nameInput.value = "";
  });

  document.getElementById("btnSaveMapping").addEventListener("click", () => {
    document.querySelectorAll("#mappingTable input").forEach(inp => {
      const key = inp.dataset.key, field = inp.dataset.field;
      MAPPING[key][field] = field === "expectedDays" ? Number(inp.value) : inp.value;
    });
    localStorage.setItem("cod_mapping", JSON.stringify(MAPPING));
    document.getElementById("mappingSaved").textContent = "✓ บันทึกแล้ว";
    buildCourierUploadGrid();
    setTimeout(() => document.getElementById("mappingSaved").textContent = "", 2000);
  });
});
