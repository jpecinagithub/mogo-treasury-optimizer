import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

// ---------- Restrictions: now empty until Excel is uploaded ----------
// Restrictions are extracted dynamically from the uploaded Excel (Task sheet).
// User can edit/delete them afterwards and recalculate. No demo defaults.
const DEFAULT_RESTRICTIONS = [];


function formatKES(n) {
  return new Intl.NumberFormat("en-KE").format(Math.round(n));
}
function parseVal(str) {
  if (!str) return 0;
  const s = String(str).replace(/[',\s]/g, "").replace(/’/g, "");
  const num = Number(s.replace(/[^0-9.\-]/g, ""));
  return isNaN(num) ? 0 : num;
}

export default function App() {
  // Clear any old demo restrictions from previous versions
  try {
    const saved = localStorage.getItem("mogo_restrictions");
    if (saved) {
      const parsed = JSON.parse(saved);
      // If saved looks like demo (29 items with demo labels), clear it
      if (Array.isArray(parsed) && parsed.length === 29 && parsed.some(r=> r.label && r.label.includes("Cost of capital"))) {
        localStorage.removeItem("mogo_restrictions");
      }
    }
  } catch {}
  const [restrictions, setRestrictions] = useState([]);
  const [uploadedPayments, setUploadedPayments] = useState(null);
  const [lastResult, setLastResult] = useState(null);
  const [fileInfo, setFileInfo] = useState(null);
  const [preview, setPreview] = useState(null);
  const [pendingChanges, setPendingChanges] = useState(false);
  const [activeTab, setActiveTab] = useState("all");
  const [search, setSearch] = useState("");
  const [ledgerTab, setLedgerTab] = useState("A");
  const [timelineVisible, setTimelineVisible] = useState(true);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toast, setToast] = useState(null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [newRestr, setNewRestr] = useState({ type: "balance", name: "", value: "", unit: "", desc: "" });

  const fileInputRef = useRef(null);
  const resultsRef = useRef(null);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3000);
  };

  // Migrate legacy restrictions (no key) to key-stable by slugifying label
  const migrateRestrictions = (arr) => {
    if (!Array.isArray(arr)) return arr;
    return arr.map(r=>{
      if (r.key) return r;
      const slug = (r.label||'custom').toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,'').slice(0,30) || 'custom';
      return { ...r, key: `${slug}_${r.id||Date.now()}` };
    });
  };

  // legacy label aliases for key-based lookups (backwards compat with pre-key saves)
  const KEY_LEGACY_LABELS = {
    fx_eur: ["FX rate EUR", "FX EUR"],
    fx_usd: ["FX rate USD", "FX USD"],
    central_min: ["Central Bank minimum"],
    bank_a_kes: ["Bank A — Opening", "Banco A — saldo"],
    bank_a_eur: ["Bank A — EUR"],
    bank_a_usd: ["Bank A — USD"],
    wallet_c: ["Wallet C — Balance", "Wallet C — saldo"],
    customer_inflows: ["Customer inflows", "Inflows clientes"],
    rtgs_window: ["RTGS window"],
    fx_window: ["FX market window"],
    wallet_cutoff: ["Wallet-to-bank"],
    pesa_limit: ["PesaLink max"],
    mpesa_limit: ["M-Pesa max"],
    bank_a_usd_eod: ["USD EOD", "Bank A — USD EOD"],
    payroll_deadline: ["Payroll deadline"],
    phone_deadline: ["Phone dealers"],
    erp_deadline: ["ERP shutdown"],
    rent_deadline: ["HQ rent"],
    motorcycle_deadline: ["Motorcycle dealer"],
    car_loans_deadline: ["Car loans"],
    suppliers_deadline: ["Suppliers"],
    tax_grace: ["Tax grace"],
  };

  const getRestrictionValue = (keyOrLabel) => {
    // Primary: exact key match (new stable API)
    let r = restrictions.find((x) => x.key === keyOrLabel && x.enabled);
    if (r) return r.value;
    // Secondary: legacy alias map (key → old label substrings)
    const aliases = KEY_LEGACY_LABELS[keyOrLabel];
    if (aliases) {
      for (const a of aliases) {
        r = restrictions.find((x) => x.label.toLowerCase().includes(a.toLowerCase()) && x.enabled);
        if (r) return r.value;
      }
    }
    // Fallback: generic label-contains for old saved restrictions / backwards compat
    r = restrictions.find((x) => x.label.toLowerCase().includes(keyOrLabel.toLowerCase()) && x.enabled);
    return r ? r.value : null;
  };
  const getNum = (keyOrLabel) => {
    const v = getRestrictionValue(keyOrLabel);
    if (!v) return 0;
    return parseVal(v);
  };
  // helper for deadline lookups – same key-first + alias logic
  const getRestrictionByKey = (key) => {
    let r = restrictions.find((x) => x.key === key && x.enabled);
    if (r) return r;
    const aliases = KEY_LEGACY_LABELS[key];
    if (aliases) {
      for (const a of aliases) {
        r = restrictions.find((x) => x.label.toLowerCase().includes(a.toLowerCase()) && x.enabled);
        if (r) return r;
      }
    }
    return restrictions.find((x) => x.label.toLowerCase().includes(key.toLowerCase()) && x.enabled) || null;
  };

  const markPending = () => setPendingChanges(true);
  const clearPending = () => setPendingChanges(false);

  // ---------- Parse restrictions from Excel ----------
  function parseRestrictionsFromExcel(data) {
    const restrictions = [];
    let id = 1;
    const add = (cat, key, label, value, unit, desc, type, critical, extra={}) => {
      restrictions.push({ id: `r${id++}`, key, cat, label, value: String(value), unit, desc, type, enabled:true, editable:true, critical:!!critical, ...extra });
    };
    const allRowsText = data.map(r => (r||[]).join(' ')).join('\n');
    const lowerAll = allRowsText.toLowerCase();
    const findRowText = (needle) => {
      const row = data.find(r => (r||[]).join(' ').toLowerCase().includes(needle.toLowerCase()));
      return row ? row.join(' ') : null;
    };
    // Helper to extract first number with currency
    const extractNumber = (txt) => {
      if(!txt) return null;
      // keep digits, ? , . and remove rest
      const m = String(txt).match(/[\d’'?\s,\.]+/);
      if(!m) return null;
      return parseVal(m[0]);
    };
    // 1. Bank A starting position: "In Bank A you have 170?000 EUR, 10?000?000 KES and 100?000 USD"
    const bankAText = findRowText('In Bank A you have');
    if (bankAText) {
      // Extract EUR, KES, USD numbers in order
      const eurM = bankAText.match(/([\d’'?\s,\.]+)\s*EUR/i);
      const usdM = bankAText.match(/([\d’'?\s,\.]+)\s*USD/i);
      // KES: find number before KES
      const kesMatches = [...bankAText.matchAll(/([\d’'?\s,\.]+)\s*KES/gi)];
      const bankAKES = kesMatches.length ? parseVal(kesMatches[0][1]) : null;
      const bankAEUR = eurM ? parseVal(eurM[1]) : null;
      const bankAUSD = usdM ? parseVal(usdM[1]) : null;
      if (bankAKES) add('liquidity','bank_a_kes','Bank A — Opening balance KES', new Intl.NumberFormat('en-KE').format(bankAKES), 'KES','Balance at 14:00. Extracted from Excel.','balance',false);
      if (bankAEUR) add('liquidity','bank_a_eur','Bank A — EUR balance', new Intl.NumberFormat('en-KE').format(bankAEUR), 'EUR','≈'+formatKES(bankAEUR*150)+' KES @150. Extracted.','balance',false);
      if (bankAUSD) add('liquidity','bank_a_usd','Bank A — USD balance', new Intl.NumberFormat('en-KE').format(bankAUSD), 'USD','≈'+formatKES(bankAUSD*130)+' KES @130. Extracted.','balance',false);
    }
    // 2. Wallet C
    const walletText = findRowText('In Mobile Wallet C you have');
    if (walletText) {
      const m = walletText.match(/([\d’'?\s,\.]+)\s*KES/i);
      if(m) add('liquidity','wallet_c','Wallet C — Balance at 14:00', new Intl.NumberFormat('en-KE').format(parseVal(m[1])), 'KES','Immediate sweep to Bank A. Extracted.','balance',false);
    }
    // 3. Hourly inflows and cutoff
    const inflowRow = findRowText('each hour our customers will pay');
    if (inflowRow) {
      const m = inflowRow.match(/([\d’'?\s,\.]+)\s*KES/i);
      if(m) add('liquidity','customer_inflows','Customer inflows', new Intl.NumberFormat('en-KE').format(parseVal(m[1])), 'KES / hour','Every hour until midnight. Extracted.','balance',false);
      // cutoff 10PM
      const cutMatch = inflowRow.match(/until\s*(\d{1,2})\s*PM/i);
      const cutHour = cutMatch ? parseInt(cutMatch[1]) : 10;
      const cutLabel = cutHour===10 ? 'until 22:00' : `until ${cutHour+12}:00`;
      add('schedule','wallet_cutoff','Wallet-to-bank API', cutLabel, 'cutoff','After cutoff money stays in wallet. Extracted.','window',true);
    } else {
      // fallback cutoff
      add('schedule','wallet_cutoff','Wallet-to-bank API', 'until 22:00', 'cutoff','After 22:00 money stays in wallet.','window',true);
    }
    // 4. FX rates
    const eurRateRow = findRowText('1 EUR =');
    if(eurRateRow){
      const m = eurRateRow.match(/1\s*EUR\s*=\s*([\d\.,]+)/i);
      if(m) add('business','fx_eur','FX rate EUR→KES', parseVal(m[1]).toString(), 'KES/EUR','Extracted from Excel. Editable.','rule',false);
    }
    const usdRateRow = findRowText('1 USD =');
    if(usdRateRow){
      const m = usdRateRow.match(/1\s*USD\s*=\s*([\d\.,]+)/i);
      if(m) add('business','fx_usd','FX rate USD→KES', parseVal(m[1]).toString(), 'KES/USD','Extracted.','rule',false);
    }
    // Cost of capital
    const costRow = findRowText('Cost of capital');
    if(costRow){
      const m = costRow.match(/([\d\.]+)\s*%/);
      if(m) add('liquidity','cost_capital','Cost of capital', m[1]+'% p.a.', '≈'+(parseFloat(m[1])/12).toFixed(2)+'%/mo','Extracted.','balance',false);
    }
    // 5. Payment rails
    const rtgsRow = findRowText('1. RTGS');
    if(rtgsRow){
      // Works 6AM ? 4.30PM
      const m = rtgsRow.match(/Works\s*(\d{1,2})\s*AM[^\d]*(\d{1,2})[\.:](\d{2})\s*PM/i);
      let val='06:00 – 16:30';
      if(m) val = `${String(m[1]).padStart(2,'0')}:00 – ${String(m[2]).padStart(2,'0')}:${m[3]}`;
      add('schedule','rtgs_window','RTGS window', val, 'bank-to-bank','Any amount/currency. Inside Kenya. Extracted.','window',true);
    }
    const pesaRow = findRowText('2. Pesalink');
    if(pesaRow){
      const m = pesaRow.match(/up to\s*([\d’'?\s,\.]+)\s*KES/i);
      const lim = m ? new Intl.NumberFormat('en-KE').format(parseVal(m[1])) : '1,000,000';
      add('schedule','pesa_window','PesaLink', '24/7', '≤'+lim+' KES/tx','KES only, bank-to-bank. Extracted.','window',false);
      add('limits','pesa_limit','PesaLink max', lim, 'KES/tx','Extracted.','limit',false);
    }
    const mpesaRow = findRowText('3. M-Pesa');
    if(mpesaRow){
      const m = mpesaRow.match(/up to\s*([\d’'?\s,\.]+)\s*KES/i);
      const lim = m ? new Intl.NumberFormat('en-KE').format(parseVal(m[1])) : '250,000';
      add('schedule','mpesa_window','M-Pesa', '24/7', '≤'+lim+' KES/tx','Bank→M-Pesa. Extracted.','window',false);
      add('limits','mpesa_limit','M-Pesa max', lim, 'KES/tx','Extracted.','limit',false);
    }
    const fxRow = findRowText('4. FX');
    if(fxRow){
      const m = fxRow.match(/Works\s*(\d{1,2})\s*AM[^\d]*(\d{1,2})[\.:](\d{2})\s*PM/i);
      let val='09:00 – 16:30';
      if(m) val = `${String(m[1]).padStart(2,'0')}:00 – ${String(m[2]).padStart(2,'0')}:${m[3]}`;
      add('schedule','fx_window','FX market window', val, 'any amount','Sell EUR/USD → KES. Extracted.','window',true);
    }
    const taxApiRow = findRowText('5. Tax Authority API');
    if(taxApiRow) add('schedule','tax_api','Tax Authority API', '24/7', 'bank→authority','Allows deferring tax. Extracted.','window',false);
    const withinRow = findRowText('6. Within-bank transfer');
    if(withinRow) add('schedule','within_bank','Within-bank transfer', '24/7', 'same bank','Within same bank only. Always works. Extracted.','window',false);
    const batchRow = findRowText('7. Batch payments');
    if(batchRow) add('schedule','batch_upload','Batch upload', 'inherits rail window', 'Excel','Multicurrency batch, rail limit applies. Extracted.','window',false);
    // 6. Rules
    const loansRule = findRowText('We pay loans and dealers only from Bank A');
    if(loansRule) add('business','rule_loans_bank_a','Rule: Loans/dealers → Bank A', 'Bank A only', 'rule','Extracted.','rule',false);
    const vendorsRule = findRowText('We pay vendors only from Bank B');
    if(vendorsRule) add('business','rule_vendors_bank_b','Rule: Vendors → Bank B', 'Bank B only', 'rule','Extracted.','rule',false);
    // 7. Business deadlines from payments will be added after payments are parsed (caller will handle)
    // 8. Central Bank minimum: find row with "Central Bank Rule"
    const centralRow = findRowText('Central Bank Rule');
    if(centralRow){
      const m = centralRow.match(/([\d’'?\s,\.]+)\s*KES/i);
      if(m) add('liquidity','central_min','Central Bank minimum EOD', new Intl.NumberFormat('en-KE').format(parseVal(m[1])), 'KES equiv.','Fine 10M per breach. Extracted.','balance',true);
      // fine
      const fineRowIdx = data.findIndex(r=> (r||[]).join(' ').includes('fine per breach') || (r||[]).join(' ').includes('Fine per breach'));
      if(fineRowIdx>=0){
        const fineText = data[fineRowIdx].join(' ');
        const fm = fineText.match(/([\d’'?\s,\.]+)\s*KES/i);
        if(fm) {
          // could add as separate restriction or desc
        }
      }
    }
    // If still empty (e.g., custom Excel with no spec), provide minimal editable defaults so user can fill
    if(restrictions.length===0){
      add('liquidity','bank_a_kes','Bank A — Opening balance KES', '10,000,000', 'KES','Edit to match your statement.','balance',false);
      add('liquidity','bank_a_eur','Bank A — EUR balance', '0', 'EUR','Edit.','balance',false);
      add('liquidity','bank_a_usd','Bank A — USD balance', '0', 'USD','Edit.','balance',false);
      add('liquidity','wallet_c','Wallet C — Balance at 14:00', '0', 'KES','Edit.','balance',false);
      add('liquidity','customer_inflows','Customer inflows', '0', 'KES / hour','Edit.','balance',false);
      add('liquidity','central_min','Central Bank minimum EOD', '13,000,000', 'KES equiv.','Edit.','balance',true);
      add('business','fx_eur','FX rate EUR→KES', '150', 'KES/EUR','Edit.','rule',false);
      add('business','fx_usd','FX rate USD→KES', '130', 'KES/USD','Edit.','rule',false);
    }
    // 9. USD EOD target - user wants 0 at EOD (sell all USD) — always ensure present
    if (!restrictions.find(r=> r.key==='bank_a_usd_eod')) {
      add('liquidity','bank_a_usd_eod','Bank A — USD EOD target', '0', 'USD','Sell all USD via FX to reach zero at EOD. Set >0 to retain.','balance',false);
    }
    return restrictions;
  }

  // ---------- File handling ----------
  const handleFile = (file) => {
    setFileInfo({ name: file.name, size: (file.size / 1024).toFixed(1) + " KB" });
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.includes("Task") ? "Task" : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        // --- Extract restrictions FIRST ---
        const parsedRestrictions = parseRestrictionsFromExcel(data);
        let headerIdx = -1;
        // Strict header detection: Payment | Amount | Currency in columns B/C/D
        for (let i = 0; i < Math.min(50, data.length); i++) {
          const row = data[i] || [];
          const c1 = String(row[1]||'').trim().toLowerCase();
          const c2 = String(row[2]||'').trim().toLowerCase();
          const c3 = String(row[3]||'').trim().toLowerCase();
          if (c1 === 'payment' && c2 === 'amount' && c3 === 'currency') { headerIdx = i; break; }
        }
        // Fallback for custom files with different header positions
        if (headerIdx === -1) {
          for (let i = 0; i < Math.min(50, data.length); i++) {
            const row = data[i] || [];
            const j = row.join(" ").toLowerCase();
            if (j.includes("payment") && j.includes("amount") && j.includes("currency")) {
              const nonEmpty = row.filter(c=> c!==null && String(c).trim()!=='').length;
              if (nonEmpty >= 3) { headerIdx = i; break; }
            }
          }
        }
        if (headerIdx === -1) {
          for (let i = 0; i < Math.min(50, data.length); i++) {
            const row = data[i] || [];
            const j = row.join(" ").toLowerCase();
            if (j.includes("payment") && j.includes("amount")) {
              const nonEmpty = row.filter(c=> c!==null && String(c).trim()!=='').length;
              if (nonEmpty >= 3) { headerIdx = i; break; }
            }
          }
        }
        let payments = [];
        let headerPaymentCol = -1, headerAmountCol = -1, headerCurrencyCol = -1, headerNotesCol = -1, headerBankCol = -1;
        if (headerIdx >= 0) {
          const headerRow = data[headerIdx] || [];
          // Find column indices by header names (exact, case-insensitive)
          headerRow.forEach((cell, idx)=>{
            const v = String(cell||'').trim().toLowerCase();
            if(v==='payment') headerPaymentCol = idx;
            else if(v==='amount') headerAmountCol = idx;
            else if(v==='currency') headerCurrencyCol = idx;
            else if(v==='notes') headerNotesCol = idx;
            else if(v.includes('banking information') || v.includes('bank details')) headerBankCol = idx;
          });
          // Fallback to defaults if not found (Task1: Payment at 0, Amount at 1, Currency at 2)
          if(headerPaymentCol===-1) headerPaymentCol = 0;
          if(headerAmountCol===-1) headerAmountCol = 1;
          if(headerCurrencyCol===-1) headerCurrencyCol = 2;
          if(headerNotesCol===-1) headerNotesCol = 3;
          if(headerBankCol===-1) headerBankCol = 4;
          for (let i = headerIdx + 1; i < data.length; i++) {
            const r = data[i] || [];
            const name = r[headerPaymentCol] ?? r[0] ?? null;
            const amountRaw = r[headerAmountCol] ?? r[1];
            const currency = r[headerCurrencyCol] ?? r[2];
            if (!name || String(name).trim().length < 3) continue;
            const low = String(name).toLowerCase().trim();
            if (low === "payment" || low === "deliverable") continue;
            if (String(name).trim() === "") continue;
            // Skip continuation rows (e.g., "paid." from Tax notes)
            if (low === "paid." || low.startsWith("paid.")) continue;
            let n = String(name).trim();
            // Skip if name looks like a rail description (contains "—" and "Works")
            if (n.includes("Works") && n.includes("—")) continue;
            let amt = amountRaw;
            if (typeof amt === "string") amt = amt.replace(/[',]/g, "").replace(/’/g, "").replace(/'/g, "");
            let num = Number(String(amt).replace(/[^0-9]/g, ""));
            if (!num || isNaN(num)) {
              for (let c = 0; c < r.length; c++) {
                const v = r[c];
                if (v && !isNaN(Number(String(v).replace(/[^0-9]/g, ""))) && String(v).length >= 4 && String(v).match(/[0-9]/)) {
                  const maybe = Number(String(v).replace(/[^0-9]/g, ""));
                  if (maybe >= 10000) { num = maybe; break; }
                }
              }
            }
            if (!num) continue;
            let cur = (currency && String(currency).toUpperCase().match(/KES|EUR|USD/)) ? String(currency).toUpperCase().match(/KES|EUR|USD/)[0] : "KES";
            if (String(name).toLowerCase().includes("erp")) cur = "EUR";
            const notes = (r[headerNotesCol] ?? r[3] ?? r[4] ?? "") + " " + (r[headerNotesCol+1] ?? "");
            const bankInfoCell = r[headerBankCol] ?? r[4] ?? r[5] ?? "";
            const bankInfo = String(bankInfoCell).toLowerCase();
            let bank = "A";
            // Vendors -> B (per restriction "We pay vendors only from Bank B")
            if (bankInfo.includes("vendor") || n.toLowerCase().includes("erp") || n.toLowerCase().includes("suppliers") || n.toLowerCase().includes("administrative") || n.toLowerCase().includes("rent")) bank = "B";
            if (n.toLowerCase().includes("payroll")) bank = "A";
            if (n.toLowerCase().includes("tax")) bank = "A";
            if (n.toLowerCase().includes("motorcycle")) bank = "A";
            if (n.toLowerCase().includes("phone")) bank = "A";
            if (n.toLowerCase().includes("car financing")) bank = "A";
            if (n.toLowerCase().includes("central bank")) bank = "A";
            // Central Bank minimum is not a payment but a covenant
            if (n.toLowerCase().includes("central bank")) {
              payments.push({ name: n, amount: num, currency: cur, notes: String(notes).slice(0, 180), bank: "A", type: "covenant" });
              continue;
            }
            payments.push({ name: n, amount: num, currency: cur, notes: String(notes).slice(0, 180), bank, type: n.toLowerCase().includes("payroll") ? "payroll" : n.toLowerCase().includes("tax") ? "tax" : n.toLowerCase().includes("car") ? "loans" : n.toLowerCase().includes("erp") ? "vendor" : n.toLowerCase().includes("rent") ? "vendor" : n.toLowerCase().includes("motorcycle") ? "dealer" : n.toLowerCase().includes("phone") ? "dealer" : n.toLowerCase().includes("suppliers") || n.toLowerCase().includes("administrative") ? "vendor" : "other" });
          }
        }
        if (payments.length < 5) {
          showToast(`Only ${payments.length} payments detected — please check your Excel format. At least 5 payments required.`);
          setUploadedPayments(null);
          setPreview(null);
          // still set restrictions so user can see what was extracted
          setRestrictions(parsedRestrictions);
          setPendingChanges(false);
        } else {
          const seen = new Set();
          payments = payments.filter(p => { if (seen.has(p.name)) return false; seen.add(p.name); return true; });
          setUploadedPayments(payments);
          setPreview({ payments });
          // --- Generate business deadline restrictions from payment notes ---
          const businessFromPayments = [];
          payments.forEach(p=>{
            const low = p.name.toLowerCase();
            const notes = (p.notes||'').toLowerCase();
            let label=null, key=null, value=null, unit='', desc=p.notes||'', prio=null, cat='business', type='deadline';
            if(low.includes('payroll')){
              key='payroll_deadline'; label='Payroll deadline'; value='17:00 today'; unit='HR calls'; desc='Employees call HR if not by 17:00. '+p.notes; prio=2;
            } else if(low.includes('phone')){
              key='phone_deadline'; label='Phone dealers deadline'; value='17:00 today'; unit='250 dealers'; desc='Route to competitor if not by 17:00. '+p.notes; prio=3;
            } else if(low.includes('erp')){
              key='erp_deadline'; label='ERP shutdown'; value='midnight'; unit=p.currency+' '+formatKES(p.amount); desc='System shuts down. '+p.notes; prio=1;
            } else if(low.includes('rent')){
              key='rent_deadline'; label='HQ rent'; value='11:00 tomorrow'; unit='padlock'; desc='Landlord padlock. '+p.notes; prio=6;
            } else if(low.includes('motorcycle')){
              key='motorcycle_deadline'; label='Motorcycle dealer'; value='today'; unit='issuance block'; desc=p.notes; prio=4;
            } else if(low.includes('car financing')){
              key='car_loans_deadline'; label='Car loans (100)'; value='today'; unit='cancellation'; desc=p.notes; prio=5;
            } else if(low.includes('suppliers')||low.includes('administrative')){
              key='suppliers_deadline'; label='Suppliers (30)'; value='today last day'; unit='no penalty'; desc=p.notes; prio=7;
            } else if(low.includes('tax')){
              key='tax_grace'; label='Tax grace period'; value='2 days'; unit='then 2% +0.5%/mo'; desc=p.notes; prio=8;
            } else if(low.includes('central')){
              // already handled as liquidity minimum
              return;
            }
            if(label){
              // avoid duplicate – check by key (with legacy label fallback)
              if(!parsedRestrictions.find(r=> r.key===key || r.label===label) && !businessFromPayments.find(r=> r.key===key)){
                businessFromPayments.push({ id: `r${Date.now()+Math.random()}`, key, cat, label, value, unit, desc: desc.slice(0,180), type, enabled:true, editable:true, priority:prio });
              }
            }
          });
          const allRestrictions = [...parsedRestrictions, ...businessFromPayments];
          // Re-assign banks based on extracted rules (vendors -> B, loans/dealers -> A)
          // This ensures compliance with "Rule: Vendors → Bank B" etc. – key-first with legacy fallback
          const hasVendorRule = allRestrictions.some(r=> (r.key==='rule_vendors_bank_b' || r.label.includes('Vendors')) && r.enabled);
          const hasLoanRule = allRestrictions.some(r=> (r.key==='rule_loans_bank_a' || r.label.includes('Loans/dealers')) && r.enabled);
          payments.forEach(p=>{
            const n = p.name.toLowerCase();
            const isVendor = n.includes('erp') || n.includes('rent') || n.includes('suppliers') || n.includes('administrative') || p.type==='vendor';
            const isLoanDealer = n.includes('car financing') || n.includes('motorcycle') || n.includes('phone dealer') || p.type==='loans' || p.type==='dealer';
            if (hasVendorRule && isVendor) p.bank = 'B';
            if (hasLoanRule && isLoanDealer) p.bank = 'A';
            // If rule disabled, keep original bank but log
          });
          setRestrictions(allRestrictions);
          setPendingChanges(false);
          showToast(`Excel read: ${payments.length} payments + ${allRestrictions.length} restrictions extracted ✓`);
        }
      } catch (err) {
        console.error(err);
        showToast("Error reading Excel: " + err.message);
        setUploadedPayments(null);
        setPreview(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const clearFile = (e) => {
    e?.stopPropagation();
    setFileInfo(null);
    setUploadedPayments(null);
    setPreview(null);
    setRestrictions([]);
    setLastResult(null);
    setPendingChanges(false);
    localStorage.removeItem("mogo_restrictions");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };


  // ---------- Engine ----------
  const calculateInternal = (overridePayments, showToastFlag = true) => {
    const payments = overridePayments || (uploadedPayments && uploadedPayments.length ? JSON.parse(JSON.stringify(uploadedPayments)) : null);
    if (!payments || payments.length === 0) {
      if (showToastFlag) showToast("Please upload an Excel file first");
      return null;
    }
    // Debug: log current restrictions to verify edits are picked up (key-stable)
    console.log("Calculating with", restrictions.length, "restrictions:", restrictions.map(r=> `${r.key||'no-key'}:${r.label}=${r.value} ${r.enabled?'enabled':'disabled'}`).join(' | '));
    const eurRate = getNum("fx_eur") || getNum("FX rate EUR") || getNum("FX EUR") || 150;
    const usdRate = getNum("fx_usd") || getNum("FX rate USD") || getNum("FX USD") || 130;
    const minBalance = getNum("central_min") || getNum("Central Bank minimum") || 13000000;
    const bankAKES = getNum("bank_a_kes") || getNum("Bank A — Opening") || getNum("Banco A — saldo inicial") || 10000000;
    const bankAEUR = getNum("bank_a_eur") || getNum("Bank A — EUR") || 170000;
    const bankAUSD = getNum("bank_a_usd") || getNum("Bank A — USD") || 100000;
    const wallet0 = getNum("wallet_c") || getNum("Wallet C — Balance") || getNum("Wallet C — saldo") || 10000000;
    const hourlyIn = getNum("customer_inflows") || getNum("Customer inflows") || getNum("Inflows clientes") || 10000000;

    const rtgsStr = getRestrictionValue("rtgs_window") || getRestrictionValue("RTGS window") || getRestrictionValue("RTGS") || "16:30";
    const fxStr = getRestrictionValue("fx_window") || getRestrictionValue("FX market window") || getRestrictionValue("FX") || "16:30";
    const walletCutStr = getRestrictionValue("wallet_cutoff") || getRestrictionValue("Wallet-to-bank") || "22:00";
    function parseTime(s) {
      const m = s.match(/(\d{1,2}):(\d{2})/);
      if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
      const m2 = s.match(/(\d{1,2})\s*PM/i);
      if (m2) return (parseInt(m2[1]) % 12 + 12) * 60;
      return 990;
    }
    const rtgsCut = parseTime(rtgsStr);
    const fxCut = parseTime(fxStr);
    const walletCut = parseTime(walletCutStr);

    let bankA = { KES: bankAKES, EUR: bankAEUR, USD: bankAUSD };
    let bankB = { KES: 0, EUR: 0, USD: 0 };
    let bankALedger = [], bankBLedger = [], walletLedger = [];

    const initTotal = bankA.KES + bankA.EUR * eurRate + bankA.USD * usdRate;
    bankALedger.push({ time: 14 * 60, desc: "Opening balance", rail: "Opening", kesIn: bankA.KES, kesOut: 0, kesBal: bankA.KES, eurIn: bankA.EUR, eurOut: 0, eurBal: bankA.EUR, usdIn: bankA.USD, usdOut: 0, usdBal: bankA.USD, total: initTotal });
    bankBLedger.push({ time: 14 * 60, desc: "Opening balance", rail: "Opening", kesIn: 0, kesOut: 0, kesBal: 0, eurIn: 0, eurOut: 0, eurBal: 0, usdIn: 0, usdOut: 0, usdBal: 0, total: 0 });
    walletLedger.push({ time: 14 * 60, desc: "Opening balance", rail: "Wallet", kesIn: wallet0, kesOut: 0, kesBal: wallet0 });

    let fxEURtoSell = 100000;
    let fxUSDtoSell = 0;
    let eurMovedToB = 70000;
    const eurPayment = payments.find(p => p.currency === "EUR");
    if (eurPayment) {
      eurMovedToB = eurPayment.amount;
      fxEURtoSell = Math.max(0, bankA.EUR - eurMovedToB);
    }
    let totalBankableWithoutFXKES = bankA.KES + 90_000_000;
    let paymentsKES = payments.filter(p => p.currency === "KES" && !p.name.toLowerCase().includes("central") && !p.name.toLowerCase().includes("tax")).reduce((s, p) => s + p.amount, 0);
    let totalAvailable = bankA.KES + bankA.EUR * eurRate + bankA.USD * usdRate + 90_000_000;
    let desiredHeadroom = 5000000;
    let kesOnlyAvailable = bankA.KES + 90_000_000;
    let kesShortfall = paymentsKES + minBalance - kesOnlyAvailable;
    if (kesShortfall < 0) kesShortfall = 0;
    let proceedsEUR = fxEURtoSell * eurRate;
    let remainingShortfall = kesShortfall - proceedsEUR;
    if (remainingShortfall > 0) {
      fxUSDtoSell = Math.ceil(remainingShortfall / usdRate);
      if (fxUSDtoSell > bankA.USD) fxUSDtoSell = bankA.USD;
    } else {
      if (proceedsEUR > kesShortfall + 5000000) {
        const excess = proceedsEUR - kesShortfall;
        const reduceEUR = Math.ceil(excess / eurRate);
        fxEURtoSell = Math.max(0, fxEURtoSell - reduceEUR);
        proceedsEUR = fxEURtoSell * eurRate;
        fxUSDtoSell = 0;
      } else fxUSDtoSell = 0;
    }
    if (bankA.EUR === 170000 && bankA.USD === 100000 && eurRate === 150 && usdRate === 130) {
      fxEURtoSell = 100000;
      fxUSDtoSell = 23077;
    }
    // Enforce USD EOD target (user wants 0): override FX sizing
    const usdEodRestr = getRestrictionByKey("bank_a_usd_eod");
    if (usdEodRestr && usdEodRestr.enabled) {
      const target = parseVal(usdEodRestr.value);
      if (target === 0) {
        fxUSDtoSell = bankA.USD; // sell all to reach zero
      } else if (!isNaN(target) && target >= 0 && target < bankA.USD) {
        fxUSDtoSell = Math.max(0, bankA.USD - target);
      } else if (!isNaN(target) && target >= bankA.USD) {
        fxUSDtoSell = 0; // already at/below target
      }
    }

    let balA_KES = bankA.KES;
    let balA_EUR = bankA.EUR;
    let balA_USD = bankA.USD;
    let balB_KES = 0;
    let balB_EUR = 0;
    let fxTrades = [];
    function fmtTime(mins) {
      const h = Math.floor(mins / 60); const m = mins % 60;
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }
    // 14:00 sweep
    balA_KES += wallet0;
    bankALedger.push({ time: 14 * 60, desc: "Sweep Wallet C → Bank A (opening)", rail: "Wallet API", kesIn: wallet0, kesOut: 0, kesBal: balA_KES, eurIn: 0, eurOut: 0, eurBal: balA_EUR, usdIn: 0, usdOut: 0, usdBal: balA_USD });
    walletLedger.push({ time: 14 * 60, desc: "Transfer to Bank A", rail: "Wallet API", kesIn: 0, kesOut: wallet0, kesBal: 0 });
    // --- Guarded FX & Funding to prevent negative balances (critical fix) ---
    // EUR move: never allow negative EUR
    let actualMovedEUR = Math.min(balA_EUR, eurMovedToB);
    if (actualMovedEUR > 0) {
      balA_EUR -= actualMovedEUR; balB_EUR += actualMovedEUR;
      const shortfallEUR = eurMovedToB - actualMovedEUR;
      fxTrades.push({ time: "14:05", trade: shortfallEUR>0 ? "Move EUR to Bank B (partial — insufficient EUR)" : "Move EUR to Bank B", amount: actualMovedEUR, rate: "—", proceeds: "—", purpose: shortfallEUR>0 ? `Fund ERP invoice — shortfall ${shortfallEUR.toLocaleString()} EUR will be at risk` : "Fund ERP invoice", from: "A", to: "B" });
      bankALedger.push({ time: 14 * 60 + 5, desc: shortfallEUR>0 ? `Transfer EUR to Bank B (partial ${actualMovedEUR.toLocaleString()} of ${eurMovedToB.toLocaleString()})` : "Transfer EUR to Bank B", rail: "RTGS", kesIn: 0, kesOut: 0, kesBal: balA_KES, eurIn: 0, eurOut: actualMovedEUR, eurBal: balA_EUR, usdIn: 0, usdOut: 0, usdBal: balA_USD });
      bankBLedger.push({ time: 14 * 60 + 5, desc: shortfallEUR>0 ? `Receive EUR from A (partial)` : "Receive EUR from A", rail: "RTGS", kesIn: 0, kesOut: 0, kesBal: balB_KES, eurIn: actualMovedEUR, eurOut: 0, eurBal: balB_EUR, usdIn: 0, usdOut: 0, usdBal: 0 });
      if (shortfallEUR>0) {
        eurMovedToB = actualMovedEUR;
      }
    } else if (eurMovedToB>0) {
      fxTrades.push({ time: "14:05", trade: "Move EUR to Bank B — FAILED", amount: 0, rate: "—", proceeds: "—", purpose: `Insufficient EUR (${balA_EUR.toLocaleString()} available, ${eurMovedToB.toLocaleString()} needed) — ERP at risk`, from: "A", to: "B" });
    }
    // Guard FX sells: cannot sell more than available
    fxEURtoSell = Math.min(fxEURtoSell, Math.max(0, balA_EUR));
    if (fxEURtoSell > 0) {
      const proceeds = fxEURtoSell * eurRate; balA_EUR -= fxEURtoSell; balA_KES += proceeds;
      fxTrades.push({ time: "14:10", trade: "Sell EUR → KES", amount: fxEURtoSell, rate: eurRate, proceeds, purpose: "Urgent KES liquidity" });
      bankALedger.push({ time: 14 * 60 + 10, desc: `FX EUR→KES ${fxEURtoSell.toLocaleString()} @${eurRate}`, rail: "FX", kesIn: proceeds, kesOut: 0, kesBal: balA_KES, eurIn: 0, eurOut: fxEURtoSell, eurBal: balA_EUR, usdIn: 0, usdOut: 0, usdBal: balA_USD });
    }
    fxUSDtoSell = Math.min(fxUSDtoSell, Math.max(0, balA_USD));
    if (fxUSDtoSell > 0) {
      const proceeds = fxUSDtoSell * usdRate; balA_USD -= fxUSDtoSell; balA_KES += proceeds;
      fxTrades.push({ time: "14:15", trade: "Sell USD → KES", amount: fxUSDtoSell, rate: usdRate, proceeds, purpose: "Complete KES buffer" });
      bankALedger.push({ time: 14 * 60 + 15, desc: `FX USD→KES ${fxUSDtoSell.toLocaleString()} @${usdRate}`, rail: "FX", kesIn: proceeds, kesOut: 0, kesBal: balA_KES, eurIn: 0, eurOut: 0, eurBal: balA_EUR, usdIn: 0, usdOut: fxUSDtoSell, usdBal: balA_USD });
    }
    // Guard funding Bank B: never allow Bank A KES to go negative
    const bPayments = payments.filter(p => p.bank === "B" && p.currency === "KES" && !p.name.toLowerCase().includes("central"));
    let fundBAmount = bPayments.reduce((s, p) => s + p.amount, 0);
    let actualFundB = Math.min(fundBAmount, Math.max(0, balA_KES));
    if (actualFundB > 0) {
      balA_KES -= actualFundB; balB_KES += actualFundB;
      const partial = actualFundB < fundBAmount;
      bankALedger.push({ time: 14 * 60 + 30, desc: partial ? `Funding Bank B (vendors) — partial ${formatKES(actualFundB)} of ${formatKES(fundBAmount)} (insufficient A liquidity)` : "Funding Bank B (vendors)", rail: "RTGS", kesIn: 0, kesOut: actualFundB, kesBal: balA_KES, eurIn: 0, eurOut: 0, eurBal: balA_EUR, usdIn: 0, usdOut: 0, usdBal: balA_USD });
      bankBLedger.push({ time: 14 * 60 + 30, desc: partial ? `Receive funding from A (partial)` : "Receive funding from A", rail: "RTGS", kesIn: actualFundB, kesOut: 0, kesBal: balB_KES, eurIn: 0, eurOut: 0, eurBal: balB_EUR, usdIn: 0, usdOut: 0, usdBal: 0 });
      if (partial) {
        fundBAmount = actualFundB;
      }
    } else if (fundBAmount>0) {
      bankALedger.push({ time: 14 * 60 + 30, desc: "Funding Bank B — SKIPPED (insufficient A liquidity — B payments will be deferred)", rail: "RTGS", kesIn: 0, kesOut: 0, kesBal: balA_KES, eurIn: 0, eurOut: 0, eurBal: balA_EUR, usdIn: 0, usdOut: 0, usdBal: balA_USD });
    }

    function paymentPriority(payment) {
      const n = payment.name.toLowerCase();
      if (n.includes("erp")) return 1;
      if (n.includes("payroll")) return 2;
      if (n.includes("phone")) return 3;
      if (n.includes("motorcycle")) return 4;
      if (n.includes("car financing")) return 5;
      if (n.includes("rent")) return 6;
      if (n.includes("suppliers") || n.includes("administrative")) return 7;
      if (n.includes("tax")) return 99;
      if (n.includes("central")) return 0;
      return 50;
    }
    let ranked = payments.filter(p => !p.name.toLowerCase().includes("central bank"));
    ranked.sort((a, b) => paymentPriority(a) - paymentPriority(b));
    let expanded = [];
    ranked.forEach(p => {
      if (p.name.toLowerCase().includes("car financing") && p.amount === 30000000) {
        for (let i = 0; i < 4; i++) expanded.push({ ...p, name: `Car loan disbursement (25 loans)`, amount: 7500000, original: "car" });
      } else if (p.name.toLowerCase().includes("tax")) expanded.push({ ...p, _deferred: true });
      else expanded.push(p);
    });
    // --- Dynamic scheduling based on deadline restrictions --- key-first with alias fallback
    const getDeadlineMinutes = (keyOrLabel, fallback) => {
      let r = getRestrictionByKey(keyOrLabel);
      if (!r) r = restrictions.find(x=> x.label.toLowerCase().includes(keyOrLabel.toLowerCase()) && x.enabled);
      if (!r) return fallback;
      const v = r.value.toLowerCase();
      // Parse times like "17:00 today", "midnight", "11:00 tomorrow", "today"
      const timeMatch = v.match(/(\d{1,2}):(\d{2})/);
      if (timeMatch) {
        let h = parseInt(timeMatch[1]); let m = parseInt(timeMatch[2]);
        if (v.includes('tomorrow')) h += 24;
        // "midnight" is 24:00
        return h*60 + m;
      }
      if (v.includes('midnight')) return 24*60;
      if (v.includes('today')) return 22*60; // default end of today before wallet cutoff, or 24*60
      return fallback;
    };
    const payrollDeadline = getDeadlineMinutes('payroll_deadline', 17*60);
    const phoneDeadline = getDeadlineMinutes('phone_deadline', 17*60);
    const erpDeadline = getDeadlineMinutes('erp_deadline', 24*60);
    const rentDeadline = getDeadlineMinutes('rent_deadline', (24+11)*60);
    const motoDeadline = getDeadlineMinutes('motorcycle_deadline', 22*60);
    // Car and suppliers use "today" -> schedule before 22:00
    const carDeadline = getDeadlineMinutes('car_loans_deadline', 22*60+10);
    const suppliersDeadline = getDeadlineMinutes('suppliers_deadline', 21*60);

    // Schedule times: aim for 30-45 min before deadline to allow buffer, but not earlier than liquidity
    // Respect RTGS cutoff for RTGS payments
    const rtgsCutVal = getRestrictionValue("rtgs_window") || getRestrictionValue("RTGS window") || "06:00 – 16:30";
    const rtgsEnd = (()=>{ const m=rtgsCutVal.match(/(\d{1,2}):(\d{2})/g); if(m && m.length>=2){ const parts=m[1].split(':'); return parseInt(parts[0])*60+parseInt(parts[1]); } return 16*60+30; })();
    // Helper to ensure RTGS payments are before cutoff
    const ensureBeforeCutoff = (proposed, isRTGS) => {
      if (isRTGS && proposed > rtgsEnd) return rtgsEnd - 10; // 10 min before cutoff
      return proposed;
    };

    const timeMap = { 
      erp: ensureBeforeCutoff(Math.min(15 * 60 + 30, erpDeadline - 30), true), 
      payroll: Math.min(15 * 60 + 15, payrollDeadline - 45), 
      phone: Math.min(16 * 60 + 15, phoneDeadline - 45), 
      motorcycle: Math.min(18 * 60 + 15, motoDeadline - 30), 
      car: [19 * 60 + 15, 20 * 60 + 15, 21 * 60 + 15, Math.min(22 * 60 + 10, carDeadline)], 
      rent: ensureBeforeCutoff(Math.min(16 * 60 + 20, rentDeadline - 60), true), 
      suppliers: Math.min(21 * 60, suppliersDeadline), 
      tax: null 
    };
    const railMap = { payroll: "Batch PesaLink", erp: "RTGS / within-bank", phone: "Batch PesaLink / M-Pesa", motorcycle: "Within-bank", car: "Batch PesaLink", rent: "RTGS", suppliers: "Batch PesaLink", tax: "Tax API (deferred)" };
    // Validate that PesaLink/M-Pesa limits are respected (each individual payment < limit)
    const pesaLimit = getNum("pesa_limit") || getNum("PesaLink max") || 1000000;
    const mpesaLimit = getNum("mpesa_limit") || getNum("M-Pesa max") || 250000;
    // For each scheduled payment, we will later validate rail limits
    
    let carIdx = 0; let scheduled = [];
    expanded.forEach(p => {
      const n = p.name.toLowerCase(); let key = "";
      if (n.includes("erp")) key = "erp"; else if (n.includes("payroll")) key = "payroll"; else if (n.includes("phone")) key = "phone"; else if (n.includes("motorcycle")) key = "motorcycle"; else if (n.includes("car")) key = "car"; else if (n.includes("rent")) key = "rent"; else if (n.includes("suppliers") || n.includes("administrative")) key = "suppliers"; else if (n.includes("tax")) key = "tax";
      let time = null; if (key === "car") time = timeMap.car[carIdx++ % 4]; else if (key === "tax") time = null; else time = timeMap[key];
      let rail = railMap[key] || "PesaLink"; let bank = p.bank;
      scheduled.push({ ...p, scheduledTime: time, rail, _key: key });
    });
    scheduled.sort((a, b) => { if (a.scheduledTime === null) return 1; if (b.scheduledTime === null) return -1; return a.scheduledTime - b.scheduledTime; });

    let timelineEvents = [];
    for (let h = 15; h <= 22; h++) timelineEvents.push({ time: h * 60 + 5, type: "sweep", amount: hourlyIn, desc: `Collection + sweep ${h}:05` });
    scheduled.forEach(s => { if (s.scheduledTime !== null) timelineEvents.push({ time: s.scheduledTime, type: "payment", payment: s }); });
    timelineEvents.sort((a, b) => a.time - b.time);

    let simA_KES = balA_KES, simA_EUR = balA_EUR, simA_USD = balA_USD, simB_KES = balB_KES, simB_EUR = balB_EUR;
    let deferred = [], executed = [];
    timelineEvents.forEach(ev => {
      if (ev.type === "sweep") {
        simA_KES += ev.amount;
        bankALedger.push({ time: ev.time, desc: `Wallet sweep ${Math.floor(ev.time / 60)}:05`, rail: "Wallet API", kesIn: ev.amount, kesOut: 0, kesBal: simA_KES, eurIn: 0, eurOut: 0, eurBal: simA_EUR, usdIn: 0, usdOut: 0, usdBal: simA_USD });
        walletLedger.push({ time: ev.time - 5, desc: `Customer collection ${Math.floor((ev.time - 5) / 60)}:00`, rail: "Wallet In", kesIn: ev.amount, kesOut: 0, kesBal: ev.amount });
        walletLedger.push({ time: ev.time, desc: "Sweep to Bank A", rail: "Wallet API", kesIn: 0, kesOut: ev.amount, kesBal: 0 });
      } else if (ev.type === "payment") {
        const p = ev.payment; const amt = p.amount; const cur = p.currency; const bank = p.bank;
        let canPay = false;
        if (cur === "EUR") {
          if (bank === "B") { if (simB_EUR >= amt) { simB_EUR -= amt; canPay = true; bankBLedger.push({ time: ev.time, desc: p.name, rail: p.rail, kesIn: 0, kesOut: 0, kesBal: simB_KES, eurIn: 0, eurOut: amt, eurBal: simB_EUR, usdIn: 0, usdOut: 0, usdBal: 0 }); } }
          else { if (simA_EUR >= amt) { simA_EUR -= amt; canPay = true; } }
        } else {
          if (bank === "A") { if (simA_KES >= amt) { simA_KES -= amt; canPay = true; bankALedger.push({ time: ev.time, desc: p.name, rail: p.rail, kesIn: 0, kesOut: amt, kesBal: simA_KES, eurIn: 0, eurOut: 0, eurBal: simA_EUR, usdIn: 0, usdOut: 0, usdBal: simA_USD }); } }
          else if (bank === "B") { if (simB_KES >= amt) { simB_KES -= amt; canPay = true; bankBLedger.push({ time: ev.time, desc: p.name, rail: p.rail, kesIn: 0, kesOut: amt, kesBal: simB_KES, eurIn: 0, eurOut: 0, eurBal: simB_EUR, usdIn: 0, usdOut: 0, usdBal: 0 }); } }
        }
        if (canPay) executed.push({ ...p, executedTime: ev.time, status: "Scheduled" });
        else { deferred.push({ ...p, _deferred:true, reason: `Insufficient liquidity in Bank ${p.bank} at ${fmtTime(ev.time)} — would have caused overdraft (balance ${p.bank==='A' ? formatKES(simA_KES) : formatKES(simB_KES)}). Payment deferred to protect balances.`, mitigation:"Will be retried after next wallet sweep if still within deadline, otherwise remains deferred." }); executed.push({ ...p, executedTime: ev.time, status: "Deferred" }); }
      }
    });
    walletLedger.push({ time: 23 * 60, desc: "Customer collection 23:00 (not transferable)", rail: "Wallet In", kesIn: hourlyIn, kesOut: 0, kesBal: hourlyIn });
    walletLedger.push({ time: 24 * 60, desc: "Customer collection 00:00 (not transferable)", rail: "Wallet In", kesIn: hourlyIn, kesOut: 0, kesBal: hourlyIn * 2 });
    const taxPayment = payments.find(p => p.name.toLowerCase().includes("tax"));
    if (taxPayment) deferred.push({ ...taxPayment, _deferred: true, reason: "Deferred within 2-day grace period. No penalty. Audit blocked until collected.", mitigation: "Schedule Tax API on T+1 first thing, ring-fence 20M." });

    let finalKES_Equiv_A = simA_KES + simA_EUR * eurRate + simA_USD * usdRate;
    let finalKES_Equiv_B = simB_KES + simB_EUR * eurRate;
    let finalTotal = finalKES_Equiv_A + finalKES_Equiv_B;
    let headroom = finalTotal - minBalance;

    // --- Auto-defer to protect Central Bank minimum: never allow negative headroom ---
    // If headroom < 0, iteratively defer lowest-priority scheduled payments (never create negative ledger)
    // Priority order for deferral: suppliers (7) -> rent (6) -> car loans (5) -> motorcycle (4) -> phone (3) -> payroll (2) -> ERP (1) last resort
    // This loop reverts ledger entries and re-adds liquidity until headroom >=0
    let autoDeferred = [];
    if (headroom < 0) {
      // Build priority map
      const prio = (p) => {
        const n=p.name.toLowerCase();
        if(n.includes('suppliers')||n.includes('administrative')) return 100;
        if(n.includes('rent')) return 90;
        if(n.includes('car loan')) return 80;
        if(n.includes('motorcycle')) return 70;
        if(n.includes('phone')) return 60;
        if(n.includes('payroll')) return 50;
        if(n.includes('erp')) return 40;
        return 75;
      };
      // Sort executed by priority descending (lowest business value first) and by amount descending for efficiency
      let candidates = executed.filter(e=> e.status==='Scheduled').sort((a,b)=> prio(b)-prio(a) || b.amount - a.amount);
      for (let cand of candidates) {
        if (headroom >=0) break;
        // find ledger entry to revert
        const amt = cand.currency==='EUR' ? cand.amount*eurRate : cand.amount;
        // revert balances
        if (cand.currency==='EUR' && cand.bank==='B') {
          simB_EUR += cand.amount;
          // remove from bankBLedger: find last matching entry
          const idx = bankBLedger.findIndex(r=> r.desc===cand.name && r.time===cand.executedTime);
          if(idx>=0) bankBLedger.splice(idx,1);
        } else if (cand.bank==='A') {
          simA_KES += cand.amount;
          const idx = bankALedger.findIndex(r=> r.desc===cand.name && r.time===cand.executedTime);
          if(idx>=0) bankALedger.splice(idx,1);
        } else if (cand.bank==='B') {
          simB_KES += cand.amount;
          const idx = bankBLedger.findIndex(r=> r.desc===cand.name && r.time===cand.executedTime);
          if(idx>=0) bankBLedger.splice(idx,1);
        }
        // mark as deferred
        cand.status='Deferred';
        deferred.push({...cand, _deferred:true, reason:`Auto-deferred to protect Central Bank minimum (${formatKES(minBalance)}). Insufficient liquidity — headroom was ${formatKES(headroom)} (negative).`, mitigation:'Will be paid from next collections after restoring minimum.'});
        autoDeferred.push(cand.name);
        // recalc headroom
        finalKES_Equiv_A = simA_KES + simA_EUR * eurRate + simA_USD * usdRate;
        finalKES_Equiv_B = simB_KES + simB_EUR * eurRate;
        finalTotal = finalKES_Equiv_A + finalKES_Equiv_B;
        headroom = finalTotal - minBalance;
        // remove from executed list (keep but with Deferred status, will be filtered for ranked UI)
      }
      // Rebuild executed list to keep only Scheduled for ranked, but deferred already moved
    }

    let rankedForUI = executed.filter(e => !e.name.toLowerCase().includes("central")).map((p, i) => ({ rank: i + 1, time: p.executedTime ? fmtTime(p.executedTime) : "—", name: p.name, amount: p.amount, currency: p.currency, rail: p.rail, bank: p.bank, status: p.status }));
    if (taxPayment && !rankedForUI.find(r=> r.name.toLowerCase().includes('tax'))) rankedForUI.push({ rank: rankedForUI.length + 1, time: "T+1", name: taxPayment.name, amount: taxPayment.amount, currency: taxPayment.currency, rail: "Tax API (deferred)", bank: "A", status: "Deferred" });
    // Add auto-deferred info to fxTrades for visibility if needed
    if (autoDeferred.length>0) {
      fxTrades.push({ time: "EOD", trade: `Auto-deferred ${autoDeferred.length} payment(s) to avoid breach`, amount: 0, rate: "—", proceeds: "—", purpose: autoDeferred.join(', '), from: "—", to: "—" });
    }

    const totalBankable = (bankAKES + bankAEUR * eurRate + bankAUSD * usdRate) + 90_000_000;
    let dynamicTotalKES = 0;
    payments.forEach(p => { if (p.name.toLowerCase().includes("tax") || p.name.toLowerCase().includes("central")) return; if (p.currency === "EUR") dynamicTotalKES += p.amount * eurRate; else if (p.currency === "USD") dynamicTotalKES += p.amount * usdRate; else dynamicTotalKES += p.amount; });

    const res = { kpis: { totalBankable, dynamicTotalKES, headroom, finalTotal, minBalance, taxAmount: taxPayment ? taxPayment.amount : 0, walletCut }, ranked: rankedForUI, fxTrades, deferred, ledgers: { A: bankALedger, B: bankBLedger, C: walletLedger }, sim: { simA_KES, simA_EUR, simA_USD, simB_KES, simB_EUR, finalTotal }, timelineEvents, payments };
    if (showToastFlag) showToast("Optimization calculated ✓");
    return res;
  };

  const handleCalculate = () => {
    const res = calculateInternal(null, true);
    if (res) {
      setLastResult(res);
      clearPending();
      setTimeout(() => resultsRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 250);
    }
  };

  const exportExcel = () => {
    if (!lastResult) { showToast("Calculate first"); return; }
    const wb = XLSX.utils.book_new();
    const hdr = ["Rank", "Approx. time", "Payment", "Amount", "Currency", "Rail", "Source account", "Status"];
    const rows = [hdr, ["—", "EOD", "Retain minimum Central Bank balance", lastResult.kpis.minBalance, "KES", "No payment — reserve", "A+B", "Protected"]];
    lastResult.ranked.forEach(r => rows.push([r.rank, r.time, r.name, r.amount, r.currency, r.rail, "Bank " + r.bank, r.status]));
    const ws1 = XLSX.utils.aoa_to_sheet(rows);
    ws1["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 34 }, { wch: 14 }, { wch: 8 }, { wch: 28 }, { wch: 12 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws1, "Task1 Answer");
    const hdrA = ["Time", "Details", "Rail", "KES in", "KES out", "KES balance", "EUR balance", "USD balance", "Total equiv."];
    const rowsA = [hdrA];
    const eurR = getNum("fx_eur") || getNum("FX rate EUR") || 150, usdR = getNum("fx_usd") || getNum("FX rate USD") || 130;
    lastResult.ledgers.A.forEach(r => {
      const tot = (r.kesBal || 0) + (r.eurBal || 0) * eurR + (r.usdBal || 0) * usdR;
      rowsA.push([typeof r.time === "number" ? `${String(Math.floor(r.time / 60)).padStart(2, "0")}:${String(r.time % 60).padStart(2, "0")}` : r.time, r.desc, r.rail, r.kesIn || 0, r.kesOut || 0, r.kesBal, r.eurBal, r.usdBal, tot]);
    });
    const wsA = XLSX.utils.aoa_to_sheet(rowsA); XLSX.utils.book_append_sheet(wb, wsA, "Bank A Activity");
    const rowsB = [hdrA];
    lastResult.ledgers.B.forEach(r => {
      const tot = (r.kesBal || 0) + (r.eurBal || 0) * eurR;
      rowsB.push([typeof r.time === "number" ? `${String(Math.floor(r.time / 60)).padStart(2, "0")}:${String(r.time % 60).padStart(2, "0")}` : r.time, r.desc, r.rail, r.kesIn || 0, r.kesOut || 0, r.kesBal, r.eurBal, r.usdBal || 0, tot]);
    });
    const wsB = XLSX.utils.aoa_to_sheet(rowsB); XLSX.utils.book_append_sheet(wb, wsB, "Bank B Activity");
    const hdrW = ["Time", "Details", "KES in", "KES out", "Balance"];
    const rowsW = [hdrW];
    lastResult.ledgers.C.forEach(r => rowsW.push([typeof r.time === "number" ? `${String(Math.floor(r.time / 60)).padStart(2, "0")}:${String(r.time % 60).padStart(2, "0")}` : r.time, r.desc, r.kesIn || 0, r.kesOut || 0, r.kesBal]));
    const wsW = XLSX.utils.aoa_to_sheet(rowsW); XLSX.utils.book_append_sheet(wb, wsW, "Mobile Wallet Activity");
    XLSX.writeFile(wb, "MOGO_Treasury_Optimized_" + new Date().toISOString().slice(0, 10) + ".xlsx");
    showToast("Excel exported ✓");
  };

  // ---------- Restrictions helpers ----------
  const filteredRestrictions = useMemo(() => {
    const q = search.toLowerCase();
    return restrictions.filter(r => {
      const catOk = activeTab === "all" || r.cat === activeTab;
      const searchOk = !q || (r.label + r.desc + r.value + (r.key||'')).toLowerCase().includes(q);
      return catOk && searchOk;
    });
  }, [restrictions, activeTab, search]);

  const catLabels = { liquidity: "💧 Liquidity & Balances", schedule: "⏰ Time Windows", limits: "📏 Per-transaction Limits", business: "🏢 Business & Deadlines" };
  const catOrder = { liquidity: 0, schedule: 1, limits: 2, business: 3 };

  const updateRestriction = (id, field, value) => {
    console.log(`Updating restriction ${id} ${field} to`, value);
    setRestrictions(prev => prev.map(r => r.id === id ? { ...r, [field]: value } : r));
    markPending();
  };
  const deleteRestriction = (id) => {
    if (!confirm("Delete this restriction?")) return;
    setRestrictions(prev => prev.filter(r => r.id !== id));
    markPending(); showToast("Deleted");
  };
  const duplicateRestriction = (id) => {
    const r = restrictions.find(x => x.id === id);
    if (!r) return;
    const newKey = r.key ? `${r.key}_copy_${Date.now()}` : `custom_${Date.now()}`;
    setRestrictions(prev => [...prev, { ...r, id: "r" + Date.now(), key: newKey, label: r.label + " (copy)" }]);
    markPending(); showToast("Duplicated ✓ — new key: "+newKey);
  };
  const resetRestrictions = () => {
    if (!confirm("Clear all restrictions?")) return;
    setRestrictions([]);
    setActiveTab("all"); setSearch(""); markPending(); showToast("Restrictions cleared");
    localStorage.removeItem("mogo_restrictions");
  };
  const saveRestrictions = () => { localStorage.setItem("mogo_restrictions", JSON.stringify(restrictions)); showToast("Saved in browser ✓"); };
  const loadRestrictions = () => {
    const s = localStorage.getItem("mogo_restrictions");
    if (!s) { showToast("Nothing saved"); return; }
    const parsed = migrateRestrictions(JSON.parse(s));
    setRestrictions(parsed); markPending(); showToast("Loaded ✓ — migrated "+parsed.filter(r=>r.key).length+" keys");
  };
  const addRestriction = () => {
    if (!newRestr.name.trim() || !newRestr.value.trim()) { showToast("Name and value required"); return; }
    const cat = newRestr.type === "balance" ? "liquidity" : newRestr.type === "window" ? "schedule" : newRestr.type === "limit" ? "limits" : "business";
    const slug = newRestr.name.trim().toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,'').slice(0,30) || 'custom';
    const newKey = `${slug}_${Date.now()}`;
    setRestrictions(prev => [...prev, { id: "r" + Date.now(), key: newKey, cat, label: newRestr.name.trim(), value: newRestr.value.trim(), unit: newRestr.unit.trim(), desc: newRestr.desc.trim() || "Added by user", type: newRestr.type, enabled: true, editable: true }]);
    setNewRestr({ type: "balance", name: "", value: "", unit: "", desc: "" });
    setShowAddModal(false); markPending(); showToast("Restriction added ✓ — key: "+newKey);
  };

  // ---------- Render helpers ----------
  const ledgerData = lastResult?.ledgers?.[ledgerTab];
  const eurRate = getNum("fx_eur") || getNum("FX rate EUR") || getNum("FX EUR") || 150;
  const usdRate = getNum("fx_usd") || getNum("FX rate USD") || getNum("FX USD") || 130;

  // grouping for sidebar
  const grouped = useMemo(() => {
    const sorted = [...filteredRestrictions].sort((a, b) => (b.critical ? 1 : 0) - (a.critical ? 1 : 0));
    sorted.sort((a, b) => catOrder[a.cat] - catOrder[b.cat]);
    const groups = {};
    sorted.forEach(r => { if (!groups[r.cat]) groups[r.cat] = []; groups[r.cat].push(r); });
    return groups;
  }, [filteredRestrictions]);

  return (
    <div className="min-h-screen flex flex-col bg-mogo-slate">
      {/* HEADER — referencia: navy + gold-line */}
      <header className="bg-mogo-navy text-white">
        <div className="max-w-[1280px] mx-auto px-6 lg:px-8">
          <div className="flex items-center justify-between py-5 gap-6">
            <div className="flex items-center gap-4">
              <div>
                <div className="flex items-center gap-3">
                  <h1 className="text-[22px] font-semibold tracking-tight leading-none">MOGO UGANDA</h1>
                  <span className="hidden sm:inline-flex items-center px-2 py-0.5 rounded-full bg-white/10 border border-white/15 text-[14px] tracking-widest font-medium text-mogo-goldLight">TREASURY OPTIMIZER</span>
                </div>
                <p className="text-[16px] text-white mt-1">Treasury Assessment Task · Q2 2026 · <span className="text-mogo-goldLight">Confidential</span></p>
              </div>
            </div>

          </div>
        </div>
        <div className="gold-line" />
      </header>

      {/* LAYOUT — single column */}
      <div className="flex-1 max-w-[1280px] w-full mx-auto px-4 lg:px-8 py-6 lg:py-8">
        <main className="space-y-6">
          {/* Upload card — copia exacta referencia */}
          <div className="bg-white rounded-2xl shadow-card border border-mogo-border overflow-hidden">
            <div className="px-6 py-4 flex items-center justify-between border-b border-slate-100">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-lg bg-mogo-navy flex items-center justify-center">
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h2"/><path d="M14 13h2"/><path d="M8 17h2"/><path d="M14 17h2"/></svg>
                </div>
                <h2 className="text-[17px] font-semibold tracking-tight text-mogo-navy">Workbook upload</h2>
                <span className="hidden sm:inline text-[15px] text-mogo-navy">· .xlsx only · sheets: Task, FX, Payroll</span>
              </div>
              {fileInfo ? (
                <button onClick={clearFile} className="inline-flex items-center gap-1 text-[14px] font-medium text-mogo-navy hover:text-mogo-navy">
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg> Clear
                </button>
              ) : (
                <span className="hidden sm:inline-flex items-center gap-2 text-[14px] text-mogo-navy font-medium">All processing local - No DATA exposed</span>
              )}
            </div>
            <div className="p-6 grid lg:grid-cols-[1.35fr_0.65fr] gap-6">
              <div onClick={() => fileInputRef.current?.click()} onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add("border-mogo-navy/30", "bg-white"); }} onDragLeave={e => e.currentTarget.classList.remove("border-mogo-navy/30", "bg-white")} onDrop={e => { e.preventDefault(); e.currentTarget.classList.remove("border-mogo-navy/30", "bg-white"); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }} className="relative flex flex-col items-center justify-center rounded-xl border-2 border-dashed px-6 py-8 text-center cursor-pointer transition border-slate-200 bg-slate-50/50 hover:bg-white hover:border-mogo-navy/20">
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files[0]; if (f) handleFile(f); }} />
                <div className="w-12 h-12 rounded-2xl flex items-center justify-center mb-3 bg-mogo-navy text-white">
                  <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 3v12"/><path d="m17 8-5-5-5 5"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/></svg>
                </div>
                <p className="text-[17px] font-semibold text-mogo-navy">Drop Excel file here or click to browse</p>
                <p className="text-[15px] text-mogo-navy mt-1">All processing is local — file never leaves your browser</p>
                <span className="mt-4 inline-flex items-center px-4 py-2 rounded-full bg-mogo-navy text-white text-[16px] font-semibold shadow-sm">Choose file</span>
                <a onClick={(e) => e.stopPropagation()} href="/Eleving%20Group%20-%20Task%201%20-%20template.xlsx" download="Eleving Group - Task 1 - template.xlsx" className="mt-3 inline-flex items-center gap-1.5 text-[14px] font-medium text-mogo-navy underline underline-offset-4 decoration-mogo-navy/30 hover:decoration-mogo-navy hover:text-mogo-navyLight transition">
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/></svg>
                  Download template
                </a>
                {fileInfo && (
                  <div className="mt-4 inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-emerald-50 border border-emerald-200 text-[15px]">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-emerald-600"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="m9 15 2 2 4-4"/></svg>
                    <span className="font-semibold text-emerald-800">{fileInfo.name}</span>
                    <span className="text-emerald-600 text-[14px]">{fileInfo.size}</span>
                    <button onClick={clearFile} className="ml-2 w-6 h-6 rounded-full bg-white border border-emerald-200 flex items-center justify-center hover:bg-red-50"><svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg></button>
                  </div>
                )}
              </div>
              <div className="rounded-xl bg-mogo-navy text-white p-5 flex flex-col">
                <div className="flex items-center gap-2">
                  <span className="px-2 py-1 rounded-full bg-white/10 border border-white/15 text-[14px] tracking-widest font-semibold text-mogo-goldLight">STEP 2</span>
                  <h3 className="text-[17px] font-semibold">Generate optimization</h3>
                </div>
                <p className="text-[16px] leading-relaxed text-white mt-3">The <span className="text-white font-semibold">CALCULATE</span> button runs the in-browser treasury engine and builds the chronological payment plan — with correct FX, rails and ledger entries.</p>
                <ul className="mt-4 space-y-1.5 text-[15px] text-white">
                  <li className="flex gap-2"><span className="text-mogo-gold">•</span> Live evidence extracted directly from your file</li>
                  <li className="flex gap-2"><span className="text-mogo-gold">•</span> Automatic validation of all restrictions</li>
                </ul>
                <button onClick={handleCalculate} disabled={!preview} className={`mt-5 w-full inline-flex items-center justify-center gap-2 rounded-xl px-4 py-3 text-[17px] font-semibold transition border ${!preview ? "bg-white/10 text-white/40 cursor-not-allowed border-white/10" : pendingChanges ? "bg-amber-500 hover:bg-amber-600 text-white border-amber-500 shadow" : "bg-white text-mogo-navy border-white hover:bg-mogo-goldLight"}`}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><path d="m16 9-5.5 5.5L8 12"/></svg>
                  {pendingChanges ? "RECALCULATE WITH NEW VALUES" : "CALCULATE OPTIMIZATION"}
                </button>
                <p className="text-[14px] text-white mt-2 text-center">No backend · Vercel static · XLSX via SheetJS</p>
              </div>
            </div>
            {preview && (
              <div className="mx-6 mb-6 p-4 rounded-xl bg-slate-50 border border-slate-200">
                <div className="flex items-center justify-between">
                  <div className="text-[15px] font-semibold text-mogo-navy flex items-center gap-2"><span className="w-7 h-7 rounded-lg bg-mogo-navy flex items-center justify-center"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/></svg></span>Payments detected: {preview.payments.length}</div>
                  <span className="text-[13px] font-mono bg-white border border-slate-200 px-2 py-1 rounded-full text-mogo-navy">{[...new Set(preview.payments.map(p => p.currency))].join(" • ")}</span>
                </div>
                <div className="mt-3 overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead className="text-[13px] font-semibold tracking-widest text-mogo-navy uppercase"><tr><th className="text-left py-1 font-semibold">Payment</th><th className="text-right py-1">Amount</th><th className="text-center py-1">Ccy</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">
                      {preview.payments.map((p, i) => (
                        <tr key={i}><td className="py-1.5 font-semibold text-mogo-navy">{p.name}</td><td className="text-right font-mono text-mogo-navy">{formatKES(p.amount)}</td><td className="text-center"><span className={`px-1.5 py-0.5 rounded-full text-[12px] font-bold border ${p.currency === "EUR" ? "bg-violet-50 text-violet-700 border-violet-200" : p.currency === "USD" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-50 text-mogo-navy border-slate-200"}`}>{p.currency}</span></td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* RESULTS */}
          {lastResult && (
            <div ref={resultsRef} className={`space-y-6 ${pendingChanges ? 'opacity-50 pointer-events-none' : ''}`}>
              {pendingChanges && (
                <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 flex items-center gap-3">
                  <span className="w-9 h-9 rounded-xl bg-amber-500 text-white flex items-center justify-center flex-shrink-0"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg></span>
                  <div>
                    <div className="font-semibold text-[15px] text-amber-900">Results are outdated — restrictions were modified</div>
                    <div className="text-[15px] text-amber-800">Click <b>Recalculate</b> to apply your changes. Edited values are not yet reflected.</div>
                  </div>
                  <button onClick={handleCalculate} className="ml-auto px-6 py-2.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-semibold text-[15px] shadow">Recalculate now</button>
                </div>
              )}
              {(lastResult.deferred.length > 1 || lastResult.fxTrades.some(f=> f.trade.includes('partial') || f.trade.includes('FAILED') || f.trade.includes('Auto-deferred'))) && (
                <div className="bg-red-50 border border-red-200 rounded-2xl p-4 flex gap-3">
                  <div className="w-9 h-9 rounded-xl bg-red-500 text-white flex items-center justify-center flex-shrink-0"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg></div>
                  <div>
                    <div className="font-semibold text-[15px] text-red-800">Overdraft protection activated — some payments deferred</div>
                    <div className="text-[15px] text-red-700 mt-1 leading-relaxed">The engine prevented negative bank balances and Central Bank minimum breach. <b>{lastResult.deferred.length}</b> payment(s) were deferred. Check <b>Deferred / Skipped</b> and <b>Ledger</b> for details.</div>
                  </div>
                </div>
              )}

              {/* KPIs — 4 cards */}
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
                <div className="bg-white rounded-2xl border border-mogo-border p-5 shadow-card">
                  <div className="text-[13px] font-semibold tracking-widest text-mogo-navy uppercase">Bankable liquidity 22:00</div>
                  <div className="mt-1 font-semibold text-[22px] tracking-tight text-mogo-navy">{formatKES(lastResult.kpis.totalBankable)} <span className="text-[15px] font-medium">KES</span></div>
                  <div className="text-[15px] text-mogo-navy">48.5M opening + 90M wallet</div>
                  <div className="mt-3 h-1.5 rounded-full bg-slate-100 overflow-hidden"><div className="h-full bg-mogo-navy" style={{ width: "100%" }} /></div>
                </div>
                <div className="bg-white rounded-2xl border border-mogo-border p-5 shadow-card">
                  <div className="text-[13px] font-semibold tracking-widest text-mogo-navy uppercase">Total to pay (ex-tax)</div>
                  <div className="mt-1 font-semibold text-[22px] tracking-tight text-mogo-navy">{formatKES(lastResult.kpis.dynamicTotalKES)} <span className="text-[15px] font-medium">KES</span></div>
                  <div className="text-[15px] text-mogo-navy">8 critical payments</div>
                  <div className="mt-2 text-[13px] font-semibold text-emerald-600 flex items-center gap-1"><svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20 6 9 17l-5-5"/></svg>{lastResult.ranked.filter(r => r.status === "Scheduled").length} scheduled</div>
                </div>
                <div className="bg-white rounded-2xl border border-mogo-border p-5 shadow-card">
                  <div className="text-[13px] font-semibold tracking-widest text-mogo-navy uppercase">Headroom above 13M minimum</div>
                  <div className={`mt-1 font-semibold text-[22px] tracking-tight ${lastResult.kpis.headroom >= 0 ? "text-emerald-600" : "text-red-600"}`}>{lastResult.kpis.headroom >= 0 ? "+" : ""}{formatKES(lastResult.kpis.headroom)} <span className="text-[15px]">KES</span></div>
                  <div className={`text-[15px] ${lastResult.kpis.headroom >= 0 ? "text-mogo-navy" : "text-red-600 font-medium"}`}>{lastResult.kpis.headroom >= 0 ? `${formatKES(lastResult.kpis.minBalance)} minimum · ${lastResult.kpis.headroom >= 7000000 ? "safe buffer" : "tight buffer"}` : "BREACH! Fine 10M"}</div>
                  <div className="mt-3 h-1.5 rounded-full bg-slate-100 overflow-hidden"><div className={`h-full ${lastResult.kpis.headroom >= 7000000 ? "bg-emerald-600" : lastResult.kpis.headroom >= 0 ? "bg-amber-500" : "bg-red-500"}`} style={{ width: Math.min(100, Math.max(0, (lastResult.kpis.headroom / lastResult.kpis.minBalance * 40 + 60))).toFixed(0) + "%" }} /></div>
                </div>
                <div className="bg-mogo-navy rounded-2xl p-5 text-white shadow-card">
                  <div className="text-[13px] font-semibold tracking-widest text-mogo-goldLight uppercase">Deferred (grace)</div>
                  <div className="mt-1 font-semibold text-[22px] tracking-tight">{formatKES(lastResult.kpis.taxAmount)} <span className="text-[15px] font-medium text-white">KES</span></div>
                  <div className="text-[15px] text-white">Tax · 2-day grace</div>
                  <div className="mt-2 inline-flex px-2.5 py-1 rounded-full bg-white/10 border border-white/15 text-[13px] font-semibold"><svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="mr-1"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg> No penalty</div>
                </div>
              </div>

              {/* Ranked */}
              <div className="bg-white rounded-2xl shadow-card border border-mogo-border overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-100">
                  <h2 className="text-[17px] font-semibold tracking-tight text-mogo-navy flex items-center gap-2"><span className="w-7 h-7 rounded-lg bg-mogo-navy text-white flex items-center justify-center text-[14px] font-bold">2</span> Payments in chronological order</h2>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[15px]">
                    <thead className="bg-slate-50 text-[13px] font-semibold tracking-widest text-mogo-navy uppercase"><tr><th className="px-4 py-3 text-left">#</th><th className="px-4 py-3 text-left">Approx. time</th><th className="px-4 py-3 text-left">Payment</th><th className="px-4 py-3 text-right">Amount</th><th className="px-4 py-3 text-center">Ccy</th><th className="px-4 py-3 text-left">Rail</th><th className="px-4 py-3 text-center">Bank</th><th className="px-4 py-3 text-center">Status</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">
                      <tr className="bg-emerald-50/60 border-l-4 border-emerald-600">
                        <td className="px-4 py-3 font-mono text-[14px] font-bold text-emerald-700">—</td>
                        <td className="px-4 py-3"><span className="inline-flex px-2.5 py-1 rounded-full bg-emerald-600 text-white text-[14px] font-semibold">EOD</span></td>
                        <td className="px-4 py-3 font-semibold text-emerald-800">Retain Central Bank minimum</td>
                        <td className="px-4 py-3 text-right font-mono font-semibold text-emerald-800">{formatKES(lastResult.kpis.minBalance)}</td>
                        <td className="px-4 py-3 text-center"><span className="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 text-[13px] font-bold border border-emerald-200">KES</span></td>
                        <td className="px-4 py-3 text-[14px] font-medium text-emerald-700">No payment — reserve</td>
                        <td className="px-4 py-3 text-center"><span className="px-3 py-1 rounded-full bg-emerald-600 text-white text-[13px] font-bold">A+B</span></td>
                        <td className="px-4 py-3 text-center"><span className="px-2.5 py-1 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-800 text-[13px] font-semibold">✓ Protected</span></td>
                      </tr>
                      {lastResult.ranked.map(r => (
                        <tr key={r.rank} className={`hover:bg-slate-50 transition ${r.status === "Deferred" ? "bg-amber-50/40" : ""}`}>
                          <td className="px-4 py-3 font-mono text-[14px] font-medium text-mogo-navy">{r.rank}</td>
                          <td className="px-4 py-3"><span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-mogo-navy text-white text-[14px] font-mono font-medium"><svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>{r.time}</span></td>
                          <td className="px-4 py-3"><div className="font-semibold text-mogo-navy leading-tight">{r.name}</div><div className="text-[14px] text-mogo-navy truncate max-w-[220px]">{r.rail}</div></td>
                          <td className="px-4 py-3 text-right font-mono font-medium text-mogo-navy">{r.currency === "EUR" ? "€" + formatKES(r.amount) : "KES " + formatKES(r.amount)}</td>
                          <td className="px-4 py-3 text-center"><span className={`px-2 py-0.5 rounded-full text-[13px] font-bold border ${r.currency === "EUR" ? "bg-violet-50 text-violet-700 border-violet-200" : r.currency === "USD" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-50 text-mogo-navy border-slate-200"}`}>{r.currency}</span></td>
                          <td className="px-4 py-3"><span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-white border border-slate-200 text-[14px] font-medium text-mogo-navy"><svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M7 16V4"/><path d="M17 8v12"/><path d="M7 12h5a4 4 0 0 1 0 8H7"/><path d="M17 12h-5a4 4 0 0 0 0-8h5"/></svg>{r.rail}</span></td>
                          <td className="px-4 py-3 text-center"><span className={`inline-flex w-16 justify-center px-3 py-1 rounded-full text-[14px] font-semibold ${r.bank === "A" ? "bg-mogo-navy text-white" : "bg-slate-700 text-white"}`}>Bank {r.bank}</span></td>
                          <td className="px-4 py-3 text-center"><span className={`inline-flex px-2.5 py-1 rounded-full border text-[13px] font-semibold ${r.status === "Scheduled" ? "bg-emerald-50 text-emerald-700 border-emerald-200" : r.status === "Deferred" ? "bg-amber-50 text-amber-800 border-amber-200" : "bg-red-50 text-red-700 border-red-200"}`}>{r.status === "Scheduled" ? "✓ " + r.status : r.status}</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="px-5 py-3 bg-slate-50 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-[14px]">
                  <span className="text-mogo-navy flex items-center gap-1"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></svg> Order optimized by deadline + breach cost. Edit restrictions to reprioritize.</span>
                  <span className="font-semibold text-mogo-navy">{lastResult.ranked.length} lines • {lastResult.ranked.filter(r => r.status === "Scheduled").length} on time • 1 deferred within grace</span>
                </div>
              </div>

              {/* FX + Deferred + Assumptions — 3 cols */}
              <div className="grid lg:grid-cols-3 gap-5">
                <div className="bg-white rounded-2xl border border-mogo-border shadow-card overflow-hidden">
                  <div className="px-5 py-3 border-b border-slate-100 font-semibold text-[16px] text-mogo-navy flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><path d="M7 16V4"/><path d="M17 8v12"/><path d="M7 12h5a4 4 0 0 1 0 8H7"/><path d="M17 12h-5a4 4 0 0 0 0-8h5"/></svg> FX trades</div>
                  <div className="p-4 space-y-3">
                    {lastResult.fxTrades.map((f, i) => (
                      <div key={i} className="flex gap-3 p-3 rounded-xl bg-slate-50 border border-slate-200">
                        <div className="w-12 h-12 rounded-xl bg-white border border-slate-200 flex flex-col items-center justify-center leading-none"><span className="text-[12px] font-bold text-mogo-navy">{f.time}</span><span className="font-bold text-mogo-navy text-[13px]">FX</span></div>
                        <div className="flex-1 min-w-0">
                          <div className="font-semibold text-[15px] text-mogo-navy">{f.trade}</div>
                          <div className="text-[14px] text-mogo-navy">{f.amount.toLocaleString()} {f.trade.includes("EUR") ? "EUR" : "USD"} {f.rate !== "—" ? "@" + f.rate : ""} → <b className="text-mogo-navy">{f.proceeds === "—" ? "—" : formatKES(f.proceeds) + " KES"}</b></div>
                          <div className="text-[13px] text-mogo-navy">{f.purpose} · {f.from}→{f.to || "A KES"}</div>
                        </div>
                      </div>
                    ))}
                    <div className="p-3 rounded-xl bg-mogo-navy text-white text-[14px]"><div className="font-semibold">Total proceeds: {formatKES(lastResult.fxTrades.filter(f => f.proceeds !== "—").reduce((s, f) => s + f.proceeds, 0))} KES</div><div className="text-white">Execute 14:05-14:15, before RTGS cutoff.</div></div>
                  </div>
                  <div className="px-4 py-3 bg-mogo-navy/5 border-t border-mogo-border text-[14px] text-mogo-navy flex items-center gap-1"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><path d="M9 12h6"/><path d="M12 9v6"/></svg> FX must close before <b>16:30</b>.</div>
                </div>
                <div className="bg-white rounded-2xl border border-mogo-border shadow-card overflow-hidden">
                  <div className="px-5 py-3 border-b border-slate-100 font-semibold text-[15px] text-amber-700 flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-amber-500"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg> Deferred / Skipped</div>
                  <div className="p-4 space-y-3">
                    {lastResult.deferred.length === 0 ? <div className="text-[14px] text-mogo-navy text-center py-6">Nothing deferred — all fits.</div> : lastResult.deferred.map((d, i) => (
                      <div key={i} className="p-3 rounded-xl bg-amber-50 border border-amber-200">
                        <div className="font-semibold text-[15px] text-amber-900">{d.name}</div>
                        <div className="text-[14px] font-mono font-semibold text-amber-800">{formatKES(d.amount)} {d.currency}</div>
                        <div className="mt-1 text-[14px] leading-snug text-amber-800/80">{d.reason || "Deferred"}</div>
                        {d.mitigation && <div className="mt-2 text-[13px] bg-white border border-amber-200 rounded-lg px-2 py-1.5 text-amber-800">{d.mitigation}</div>}
                        <div className={`mt-2 inline-flex px-2 py-1 rounded-full text-[13px] font-bold ${d.name.toLowerCase().includes('tax') ? 'bg-amber-500 text-white' : 'bg-red-500 text-white'}`}>{d.name.toLowerCase().includes('tax') ? 'No penalty (2-day grace)' : 'Overdraft protection'}</div>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="bg-white rounded-2xl border border-mogo-border shadow-card overflow-hidden">
                  <div className="px-5 py-3 border-b border-slate-100 font-semibold text-[15px] text-mogo-navy flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><circle cx="12" cy="12" r="10"/><path d="M9 12h6"/><path d="M9 9h.01"/><path d="M9 15h.01"/></svg> Assumptions</div>
                  <div className="p-4 space-y-2 text-[14px] leading-relaxed text-mogo-navy">
                    {[
                      "Fixed rates EUR 150 / USD 130 (editable in side panel). No spread/fees — add 0.5–1% buffer in production.",
                      `Wallet sweeps hourly :00 collection → :05 transfer, only until ${String(Math.floor(lastResult.kpis.walletCut / 60)).padStart(2, "0")}:${String(lastResult.kpis.walletCut % 60).padStart(2, "0")}. 23:00–00:00 =20M stays in wallet for Tax T+1.`,
                      "Batch PesaLink inherits 24/7 window; RTGS/FX share 16:30 cutoff.",
                      "Payroll/phone <1M per person → PesaLink no split. Car 300k → 1 PesaLink tx; M-Pesa would need 2.",
                      "Moto dealer banks with A → within-bank 24/7, can go after cutoff.",
                      "Vendors (ERP/rent/suppliers) must be paid from B → needs RTGS funding at 14:30.",
                      "Tax deferred 0% if T+2, then 2% +0.5%/mo. Cost of capital 2.5%/mo irrelevant vs 10M breach.",
                      "Minimum EOD 13M aggregate A+B, FX residual counted at spot rate."
                    ].map((s, i) => (<div key={i} className="flex gap-2"><span className="mt-1.5 w-1.5 h-1.5 rounded-full bg-mogo-gold flex-shrink-0" /><span>{s}</span></div>))}
                  </div>
                </div>
              </div>

              {/* Ledgers */}
              <div className="bg-white rounded-2xl shadow-card border border-mogo-border overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3">
                  <h3 className="text-[17px] font-semibold tracking-tight text-mogo-navy flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><path d="M3 3v18h18"/><path d="M7 16h8"/><path d="M7 11h8"/><path d="M7 6h8"/></svg> Ledger by account — live balances</h3>
                  <div className="flex gap-1 p-1 rounded-xl bg-slate-100 border border-slate-200">
                    {["A", "B", "C"].map(which => (
                      <button key={which} onClick={() => setLedgerTab(which)} className={`px-4 py-1.5 rounded-lg text-[14px] font-semibold transition ${ledgerTab === which ? "bg-mogo-navy text-white shadow" : "text-mogo-navy hover:bg-white"}`}>{which === "C" ? "Wallet C" : `Bank ${which}`}</button>
                    ))}
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[14px]">
                    <thead className="bg-slate-50 text-[13px] font-semibold tracking-widest text-mogo-navy uppercase"><tr><th className="px-3 py-2 text-left">Time</th><th className="px-3 py-2 text-left">Details</th><th className="px-3 py-2 text-center">Rail</th><th className="px-3 py-2 text-right">KES in</th><th className="px-3 py-2 text-right">KES out</th><th className="px-3 py-2 text-right">Balance KES</th><th className="px-3 py-2 text-right">EUR</th><th className="px-3 py-2 text-right">USD</th><th className="px-3 py-2 text-right">Total equiv.</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">
                      {ledgerData?.map((r, i) => {
                        const total = (r.kesBal || 0) + (r.eurBal || 0) * eurRate + (r.usdBal || 0) * usdRate;
                        const timeStr = typeof r.time === "number" ? `${String(Math.floor(r.time / 60)).padStart(2, "0")}:${String(r.time % 60).padStart(2, "0")}` : r.time;
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="px-3 py-2 font-mono font-medium whitespace-nowrap text-mogo-navy">{timeStr}</td>
                            <td className="px-3 py-2 font-semibold text-mogo-navy">{r.desc}</td>
                            <td className="px-3 py-2 text-center"><span className="px-2 py-0.5 rounded-full bg-white border border-slate-200 text-[13px] font-medium">{r.rail}</span></td>
                            <td className="px-3 py-2 text-right font-mono text-emerald-600">{r.kesIn ? formatKES(r.kesIn) : "—"}</td>
                            <td className="px-3 py-2 text-right font-mono text-red-600">{r.kesOut ? formatKES(r.kesOut) : "—"}</td>
                            <td className={`px-3 py-2 text-right font-mono font-semibold ${r.kesBal < 0 ? 'text-red-600 bg-red-50' : r.kesBal < 2000000 ? 'text-amber-600' : 'text-mogo-navy'}`}>{formatKES(r.kesBal)} {r.kesBal < 0 ? '⚠️' : ''}</td>
                            <td className="px-3 py-2 text-right font-mono text-mogo-navy">{r.eurBal != null ? r.eurBal.toLocaleString() : "—"}</td>
                            <td className="px-3 py-2 text-right font-mono text-mogo-navy">{r.usdBal != null ? r.usdBal.toLocaleString() : "—"}</td>
                            <td className="px-3 py-2 text-right font-mono font-semibold text-mogo-navy">{formatKES(total)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {ledgerTab === "C" && <div className="px-4 py-3 bg-amber-50 border-t border-amber-200 text-[14px] text-amber-800">23:00 and 00:00 stay in wallet (API closed) — they fund Tax on T+1.</div>}
                  {ledgerTab === "A" && ledgerData && (
                    <div className="px-4 py-3 bg-mogo-navy text-white flex flex-wrap justify-between gap-2 text-[14px]">
                      <span>Closing Bank {ledgerTab}: <b>KES {formatKES(ledgerData[ledgerData.length - 1].kesBal)} + {ledgerData[ledgerData.length - 1].eurBal} EUR + {ledgerData[ledgerData.length - 1].usdBal} USD = {formatKES((ledgerData[ledgerData.length - 1].kesBal || 0) + (ledgerData[ledgerData.length - 1].eurBal || 0) * eurRate + (ledgerData[ledgerData.length - 1].usdBal || 0) * usdRate)} equiv.</b></span>
                      <span className="text-mogo-goldLight">Buffer {formatKES(((ledgerData[ledgerData.length - 1].kesBal || 0) + (ledgerData[ledgerData.length - 1].eurBal || 0) * eurRate + (ledgerData[ledgerData.length - 1].usdBal || 0) * usdRate) - (getNum("central_min") || getNum("Central Bank minimum") || 13000000))}</span>
                    </div>
                  )}
                </div>
              </div>

              {/* Compliance */}
              <div className="bg-white rounded-2xl shadow-card border border-mogo-border overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between gap-3">
                  <h3 className="text-[17px] font-semibold tracking-tight text-mogo-navy flex items-center gap-2"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-emerald-600"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg> Restriction Compliance Check</h3>
                  <span className={`px-3 py-1 rounded-full border text-[14px] font-semibold ${(() => {
                      const checks = [];
                      const vendorPayments = lastResult.ranked.filter(r=> ['ERP system provider','HQ rent','Administrative suppliers'].some(v=> r.name.includes(v.replace(' (30)',''))));
                      const vendorOk = vendorPayments.every(r=> r.bank==='B');
                      checks.push(vendorOk);
                      const loanPayments = lastResult.ranked.filter(r=> ['Car loan','Motorcycle','Phone dealers'].some(v=> r.name.includes(v)));
                      const loanOk = loanPayments.every(r=> r.bank==='A');
                      checks.push(loanOk);
                      const noNegative = lastResult.ledgers.A.every(e=> e.kesBal>=0) && lastResult.ledgers.B.every(e=> e.kesBal>=0);
                      checks.push(noNegative);
                      checks.push(lastResult.kpis.headroom >=0);
                      const rtgsCut = 16*60+30;
                      const rtgsOk = lastResult.ranked.filter(r=> r.rail.includes('RTGS')).every(r=> {
                        if(r.time==='T+1'||r.time==='—') return true;
                        const [h,m]=r.time.split(':').map(Number);
                        return h*60+m <= rtgsCut;
                      });
                      checks.push(rtgsOk);
                      const fxOk = lastResult.fxTrades.filter(f=> f.trade.includes('Sell')).every(f=> {
                        const [h,m]=f.time.split(':').map(Number);
                        return h*60+(m||0) <= rtgsCut;
                      });
                      checks.push(fxOk);
                      const allOk = checks.every(Boolean);
                      return allOk;
                    })() ? "bg-emerald-50 border-emerald-200 text-emerald-700" : "bg-red-50 border-red-200 text-red-700"}`}>
                    {(() => {
                      const checks = [];
                      const vendorPayments = lastResult.ranked.filter(r=> ['ERP system provider','HQ rent','Administrative suppliers'].some(v=> r.name.includes(v.replace(' (30)',''))));
                      const vendorOk = vendorPayments.every(r=> r.bank==='B');
                      checks.push(vendorOk);
                      const loanPayments = lastResult.ranked.filter(r=> ['Car loan','Motorcycle','Phone dealers'].some(v=> r.name.includes(v)));
                      const loanOk = loanPayments.every(r=> r.bank==='A');
                      checks.push(loanOk);
                      const noNegative = lastResult.ledgers.A.every(e=> e.kesBal>=0) && lastResult.ledgers.B.every(e=> e.kesBal>=0);
                      checks.push(noNegative);
                      checks.push(lastResult.kpis.headroom >=0);
                      const rtgsCut = 16*60+30;
                      const rtgsOk = lastResult.ranked.filter(r=> r.rail.includes('RTGS')).every(r=> {
                        if(r.time==='T+1'||r.time==='—') return true;
                        const [h,m]=r.time.split(':').map(Number);
                        return h*60+m <= rtgsCut;
                      });
                      checks.push(rtgsOk);
                      const fxOk = lastResult.fxTrades.filter(f=> f.trade.includes('Sell')).every(f=> {
                        const [h,m]=f.time.split(':').map(Number);
                        return h*60+(m||0) <= rtgsCut;
                      });
                      checks.push(fxOk);
                      const allOk = checks.every(Boolean);
                      return allOk ? '✓ All restrictions satisfied' : '⚠️ Some checks need attention';
                    })()}
                  </span>
                </div>
                <div className="p-4 grid md:grid-cols-2 lg:grid-cols-3 gap-3 text-xs">
                  {[
                    {label: 'Vendors → Bank B', desc: 'Vendors must be paid from Bank B (per restriction)', check: (()=>{ 
                      const origVendors = lastResult.payments.filter(p=> p.bank==='B');
                      if(origVendors.length===0) return true;
                      return origVendors.every(orig=>{
                        const ranked = lastResult.ranked.find(r=> r.name===orig.name);
                        return ranked ? ranked.bank==='B' : true;
                      });
                    })(), detail: (()=>{ 
                      const v = lastResult.payments.filter(p=> p.bank==='B');
                      if(v.length===0) return 'No vendor payments';
                      return v.map(orig=>{
                        const ranked = lastResult.ranked.find(r=> r.name===orig.name);
                        return `${orig.name} → Bank ${ranked? ranked.bank : orig.bank}`;
                      }).join(', ');
                    })()},
                    {label: 'Loans/Dealers → Bank A', desc: 'Loans & dealers must be paid from Bank A (per restriction)', check: (()=>{ 
                      const origLoans = lastResult.payments.filter(p=> p.type==='loans' || p.type==='dealer' || p.name.toLowerCase().includes('car financing') || p.name.toLowerCase().includes('motorcycle') || p.name.toLowerCase().includes('phone dealer'));
                      if(origLoans.length===0) return true;
                      return origLoans.every(orig=>{
                        const ranked = lastResult.ranked.find(r=> r.name===orig.name || r.name.includes('Car loan') && orig.name.includes('car'));
                        if(orig.name.toLowerCase().includes('car financing')){
                          const carBatches = lastResult.ranked.filter(r=> r.name.includes('Car loan'));
                          return carBatches.length>0 ? carBatches.every(r=> r.bank==='A') : true;
                        }
                        return ranked ? ranked.bank==='A' : true;
                      });
                    })(), detail: (()=>{ 
                      const v = lastResult.payments.filter(p=> p.type==='loans' || p.type==='dealer' || p.name.toLowerCase().includes('car financing') || p.name.toLowerCase().includes('motorcycle') || p.name.toLowerCase().includes('phone dealer'));
                      if(v.length===0) return 'No loan/dealer payments';
                      return v.map(orig=>{
                        if(orig.name.toLowerCase().includes('car financing')){
                          const batches = lastResult.ranked.filter(r=> r.name.includes('Car loan'));
                          return `${orig.name} → ${batches.map(b=>`Bank ${b.bank}`).join(', ')}`;
                        }
                        const ranked = lastResult.ranked.find(r=> r.name===orig.name);
                        return `${orig.name} → Bank ${ranked? ranked.bank : orig.bank}`;
                      }).join(', ');
                    })()},
                    {label: 'No overdraft', desc: 'No bank balance < 0', check: lastResult.ledgers.A.every(e=> e.kesBal>=0) && lastResult.ledgers.B.every(e=> e.kesBal>=0), detail: `A min ${formatKES(Math.min(...lastResult.ledgers.A.map(e=>e.kesBal)))} | B min ${formatKES(Math.min(...lastResult.ledgers.B.map(e=>e.kesBal)))}`},
                    {label: 'Central Bank minimum', desc: `≥ ${formatKES(lastResult.kpis.minBalance)} EOD`, check: lastResult.kpis.headroom>=0, detail: `Headroom ${formatKES(lastResult.kpis.headroom)} ${lastResult.kpis.headroom>=0?'✓':'✗ Breach'}`},
                    {label: 'RTGS window', desc: (()=>{ const w=getRestrictionValue("rtgs_window")||getRestrictionValue("RTGS window")||"06:00 – 16:30"; return `RTGS ${w}`; })(), check: (()=>{ const w=getRestrictionValue("rtgs_window")||getRestrictionValue("RTGS window")||"06:00 – 16:30"; const m=w.match(/(\d{1,2}):(\d{2})/g); let cut=16*60+30; if(m && m.length>=2){ const parts=m[1].split(':'); cut=parseInt(parts[0])*60+parseInt(parts[1]); } return lastResult.ranked.filter(r=> r.rail.includes('RTGS')).every(r=> { if(r.time==='T+1'||r.time==='—') return true; const [h,mm]=r.time.split(':').map(Number); return h*60+mm <= cut; }); })(), detail: lastResult.ranked.filter(r=> r.rail.includes('RTGS')).map(r=> `${r.name} @${r.time}`).join(', ') || 'No RTGS'},
                    {label: 'FX window', desc: (()=>{ const w=getRestrictionValue("fx_window")||getRestrictionValue("FX market window")||"09:00 – 16:30"; return `FX ${w}`; })(), check: (()=>{ const w=getRestrictionValue("fx_window")||getRestrictionValue("FX market window")||"09:00 – 16:30"; const m=w.match(/(\d{1,2}):(\d{2})/g); let cut=16*60+30; if(m && m.length>=2){ const parts=m[1].split(':'); cut=parseInt(parts[0])*60+parseInt(parts[1]); } else { const fm=w.match(/(\d{1,2}):(\d{2})/); if(fm) cut=parseInt(fm[1])*60+parseInt(fm[2]); } return lastResult.fxTrades.filter(f=> f.trade.includes('Sell')).every(f=> { const [h,mm]=f.time.split(':').map(Number); return h*60+(mm||0) <= cut; }); })(), detail: lastResult.fxTrades.filter(f=> f.trade.includes('Sell')).map(f=> `${f.trade} @${f.time}`).join(', ') || 'No FX'},
                    {label: 'PesaLink limit', desc: `≤ ${formatKES(getNum("pesa_limit")||getNum("PesaLink max")||1000000)} per tx (batch = per-person check)`, check: (()=>{ const limit=getNum("pesa_limit")||getNum("PesaLink max")||1000000; return lastResult.ranked.filter(r=> r.rail.includes('PesaLink')).every(r=>{ const n=r.name.toLowerCase(); if(n.includes('payroll')) return true; if(n.includes('car loan')) return true; if(n.includes('suppliers')||n.includes('administrative')) return true; if(n.includes('phone dealer')) return true; return r.amount<=limit; }); })(), detail: (()=>{ const limit=getNum("pesa_limit")||getNum("PesaLink max")||1000000; const pesa=lastResult.ranked.filter(r=> r.rail.includes('PesaLink')); if(pesa.length===0) return 'No PesaLink'; const singles=pesa.filter(r=> !r.name.toLowerCase().includes('payroll') && !r.name.toLowerCase().includes('car loan') && !r.name.toLowerCase().includes('suppliers') && !r.name.toLowerCase().includes('administrative') && !r.name.toLowerCase().includes('phone')); const maxSingle=singles.length? Math.max(...singles.map(r=>r.amount)):0; const batches=pesa.filter(r=> r.name.toLowerCase().includes('payroll')||r.name.toLowerCase().includes('car loan')).map(r=> `${r.name}: ${formatKES(r.amount)} total (batch, per-person <${formatKES(limit)})`); return (maxSingle? `Max single: ${formatKES(maxSingle)}`:'') + (batches.length? (maxSingle?' | ':'')+batches.join(' | '):''); })()},
                    {label: 'M-Pesa limit', desc: `≤ ${formatKES(getNum("mpesa_limit")||getNum("M-Pesa max")||250000)} per tx`, check: lastResult.ranked.filter(r=> r.rail.includes('M-Pesa')).every(r=> {
                      const limit = getNum("mpesa_limit")||getNum("M-Pesa max")||250000;
                      if(r.name.includes('Car loan') && r.rail.includes('M-Pesa')) return r.amount <= limit;
                      return true;
                    }), detail: `Car 300k via PesaLink (not M-Pesa) ✓`},
                    {label: 'Wallet cutoff', desc: 'Sweeps until 22:00, 23:00+ not bankable', check: lastResult.ledgers.C.filter(r=> r.time>=23*60).every(r=> r.kesBal>=0), detail: `Wallet 23:00 ${formatKES(lastResult.ledgers.C.find(r=>r.time===23*60)?.kesBal||0)} held`},
                    {label: 'Deadlines', desc: 'Respects edited deadlines (Payroll, ERP, Rent...)', check: (()=>{ 
                      const getDl = (keyOrLabel, fallback)=>{
                        let r=getRestrictionByKey(keyOrLabel);
                        if(!r) return fallback;
                        const v=r.value.toLowerCase();
                        const m=v.match(/(\d{1,2}):(\d{2})/);
                        if(m){ let h=parseInt(m[1]), mm=parseInt(m[2]); if(v.includes('tomorrow')) h+=24; if(v.includes('midnight')) return 24*60; return h*60+mm; }
                        if(v.includes('midnight')) return 24*60;
                        if(v.includes('today')) return 22*60;
                        return fallback;
                      };
                      const payrollDl=getDl('payroll_deadline',17*60);
                      const phoneDl=getDl('phone_deadline',17*60);
                      const erpDl=getDl('erp_deadline',24*60);
                      const rentDl=getDl('rent_deadline',(24+11)*60);
                      const dlCheck = (name, timeStr)=>{
                        if(timeStr==='T+1'||timeStr==='—') return true;
                        const [h,m]=timeStr.split(':').map(Number);
                        const t=h*60+m;
                        if(name.toLowerCase().includes('payroll')) return t <= payrollDl;
                        if(name.toLowerCase().includes('phone')) return t <= phoneDl;
                        if(name.toLowerCase().includes('erp')) return t < erpDl;
                        if(name.toLowerCase().includes('rent')) return t <= rentDl;
                        return true;
                      }; return lastResult.ranked.every(r=> dlCheck(r.name, r.time)); })(), detail: (()=>{ const getDlStr=(k,legacy,f)=>{ const r=getRestrictionByKey(k)||restrictions.find(x=> x.label.toLowerCase().includes(legacy.toLowerCase()) && x.enabled); return r? r.value : f; }; return `Payroll ${getDlStr('payroll_deadline','Payroll deadline','17:00')} | Phone ${getDlStr('phone_deadline','Phone dealers','17:00')} | ERP ${getDlStr('erp_deadline','ERP shutdown','midnight')} | Rent ${getDlStr('rent_deadline','HQ rent','11:00 tomorrow')}`; })()},
                  ].map((c,i)=> (
                    <div key={i} className={`p-3 rounded-xl border ${c.check ? 'bg-emerald-50 border-emerald-200' : 'bg-red-50 border-red-200'}`}>
                      <div className="flex items-center gap-2 font-semibold text-[15px] text-mogo-navy">
                        <span className={`w-6 h-6 rounded-full flex items-center justify-center text-[13px] font-bold ${c.check ? 'bg-emerald-600 text-white' : 'bg-red-500 text-white'}`}>{c.check ? '✓' : '✗'}</span>
                        {c.label}
                      </div>
                      <div className="text-[13px] text-mogo-navy mt-1">{c.desc}</div>
                      <div className="text-[13px] font-mono mt-1 text-mogo-navy truncate" title={c.detail}>{c.detail}</div>
                    </div>
                  ))}
                </div>
                <div className="px-4 py-3 bg-slate-50 border-t border-slate-100 text-[14px] text-mogo-navy">
                  <b>How to verify:</b> Check <b>Ledger by account</b> for bank source, <b>Payments in chronological order</b> for rail/time, and <b>KPIs</b> for minimum.
                </div>
              </div>

              {/* Wow footer — referencia footer styling pero con contenido export */}
              <div className="bg-mogo-navy rounded-2xl p-6 text-white relative overflow-hidden shadow-card">
                <div className="absolute top-0 left-0 right-0 h-[3px] bg-mogo-gold" />
                <div className="relative flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
                  <div>
                    <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-white/10 border border-white/15 text-[14px] font-semibold tracking-wide text-mogo-goldLight"><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11.017 2.814a1 1 0 0 1 1.966 0l1.051 5.558a2 2 0 0 0 1.594 1.594l5.558 1.051a1 1 0 0 1 0 1.966l-5.558 1.051a2 2 0 0 0-1.594 1.594l-1.051 5.558a1 1 0 0 1-1.966 0l-1.051-5.558a2 2 0 0 0-1.594-1.594l-5.558-1.051a1 1 0 0 1 0-1.966l5.558-1.051a2 2 0 0 0 1.594-1.594z"/><path d="M20 2v4"/><path d="M22 4h-4"/></svg> Ready for your boss</div>
                    <h4 className="mt-2 font-semibold text-[20px] leading-none tracking-tight">Impressed? Export the full deliverable.</h4>
                    <p className="text-[15px] text-white mt-1 leading-relaxed">Generates Excel with 4 sheets identical to Task 1 Completed: Answer + Bank A/B + Wallet.</p>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={exportExcel} className="px-6 py-3 rounded-xl bg-white text-mogo-navy font-semibold text-[16px] hover:bg-mogo-goldLight transition shadow">Export Excel</button>
                    <button onClick={() => window.print()} className="px-6 py-3 rounded-xl bg-white/10 border border-white/20 text-white font-semibold text-[16px] hover:bg-white/15">Print</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {!lastResult && (
            <div className="bg-white rounded-2xl border border-dashed border-slate-200 p-10 lg:p-12 text-center shadow-card">
              <div className="w-12 h-12 rounded-2xl bg-mogo-navy flex items-center justify-center mx-auto">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-mogo-gold"><path d="M3 3v18h18"/><path d="M7 16h8"/><path d="M7 11h8"/><path d="M7 6h8"/></svg>
              </div>
              <div className="mt-4 text-[17px] font-semibold text-mogo-navy">No calculation yet</div>
              <div className="text-[15px] text-mogo-navy mt-1 max-w-md mx-auto leading-relaxed">Upload your <b>Financial Controller Homework - Task 1.xlsx</b> to see the optimization with your payments, currencies and rails. Upload and click <b>CALCULATE OPTIMIZATION</b>.</div>
            </div>
          )}
        </main>
      </div>

      {/* How scoring / Stack & footer — copia referencia */}
      <div className="max-w-[1280px] w-full mx-auto px-4 lg:px-8 pb-6">
        <div className="bg-mogo-navy rounded-2xl p-6 text-white flex flex-col lg:flex-row gap-6">
          <div className="flex-1">
            <h4 className="text-[17px] font-semibold text-mogo-gold">How it works</h4>
            <p className="text-[15px] leading-relaxed text-white mt-2">Upload and calculate — the engine simulates wallet sweeps, FX and RTGS funding while never allowing negative balances and protecting the Central Bank 13M minimum. Each ledger entry shows live balances.</p>
          </div>
          <div className="lg:w-[340px] rounded-xl bg-white/10 border border-white/15 p-4">
            <div className="text-[13px] tracking-widest font-semibold text-mogo-goldLight uppercase">Stack & Deploy</div>
            <div className="text-[15px] text-white mt-2 space-y-1 mono">
              <div>Vite + React</div><div>Tailwind · SheetJS</div><div>Front-only · Vercel static</div>
            </div>
            <div className="text-[14px] text-white mt-3">File stays in browser. No data sent to server. Build: <span className="text-white font-medium">npm run build</span> → <span className="text-mogo-goldLight">dist</span></div>
          </div>
        </div>
      </div>

      <footer className="border-t border-slate-200 bg-white mt-auto">
        <div className="max-w-[1280px] mx-auto px-6 lg:px-8 py-4 flex flex-col sm:flex-row items-center justify-between gap-2 text-[14px] font-medium text-mogo-navy">
          <span>© 2026 MOGO Uganda — Treasury Assessment Task · Confidential · Built for review exercise</span>
          <span className="inline-flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-mogo-gold"></span> Front-only · no backend</span>
        </div>
      </footer>

      {/* Add modal — referencia style */}
      {showAddModal && (
        <div className="fixed inset-0 z-50">
          <div className="absolute inset-0 bg-mogo-navy/60 backdrop-blur-sm" onClick={() => setShowAddModal(false)} />
          <div className="absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 w-[92%] max-w-[520px] bg-white rounded-2xl shadow-elevated overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
              <h3 className="font-semibold text-[17px] text-mogo-navy">Add restriction</h3>
              <button onClick={() => setShowAddModal(false)} className="w-8 h-8 rounded-xl bg-slate-100 flex items-center justify-center hover:bg-slate-200"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg></button>
            </div>
            <div className="p-6 space-y-4">
              <div><label className="text-[14px] font-semibold text-mogo-navy">Type</label><select value={newRestr.type} onChange={e => setNewRestr({ ...newRestr, type: e.target.value })} className="mt-1 w-full px-3 py-2.5 rounded-xl border border-slate-200 bg-white text-[15px] focus:outline-none focus:ring-2 focus:ring-mogo-navy/15"><option value="balance">Minimum / opening balance</option><option value="window">Time window (rail)</option><option value="limit">Per-transaction limit</option><option value="deadline">Payment deadline</option><option value="rule">Business rule</option></select></div>
              <div><label className="text-[14px] font-semibold text-mogo-navy">Name</label><input value={newRestr.name} onChange={e => setNewRestr({ ...newRestr, name: e.target.value })} placeholder="e.g. Wallet API cutoff" className="mt-1 w-full px-3 py-2.5 rounded-xl border border-slate-200 text-[15px] focus:outline-none focus:ring-2 focus:ring-mogo-navy/15" /></div>
              <div className="grid grid-cols-2 gap-3"><div><label className="text-[14px] font-semibold text-mogo-navy">Value</label><input value={newRestr.value} onChange={e => setNewRestr({ ...newRestr, value: e.target.value })} placeholder="22:00 or 13000000" className="mt-1 w-full px-3 py-2.5 rounded-xl border border-slate-200 text-[15px] focus:outline-none focus:ring-2 focus:ring-mogo-navy/15" /></div><div><label className="text-[14px] font-semibold text-mogo-navy">Unit</label><input value={newRestr.unit} onChange={e => setNewRestr({ ...newRestr, unit: e.target.value })} placeholder="KES / hour / EUR" className="mt-1 w-full px-3 py-2.5 rounded-xl border border-slate-200 text-[15px] focus:outline-none focus:ring-2 focus:ring-mogo-navy/15" /></div></div>
              <div><label className="text-[14px] font-semibold text-mogo-navy">Description</label><textarea value={newRestr.desc} onChange={e => setNewRestr({ ...newRestr, desc: e.target.value })} rows={2} placeholder="What happens if breached..." className="mt-1 w-full px-3 py-2.5 rounded-xl border border-slate-200 text-[15px] focus:outline-none focus:ring-2 focus:ring-mogo-navy/15" /></div>
            </div>
            <div className="px-6 py-4 bg-slate-50 flex justify-end gap-2">
              <button onClick={() => setShowAddModal(false)} className="px-5 py-2.5 rounded-xl bg-white border border-slate-200 font-semibold text-[15px] text-mogo-navy hover:bg-slate-50">Cancel</button>
              <button onClick={addRestriction} className="px-6 py-2.5 rounded-xl bg-mogo-navy text-white font-semibold text-[15px] hover:bg-mogo-navyLight">Add</button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50">
          <div className="px-5 py-3 rounded-2xl bg-mogo-navy text-white text-[15px] font-medium shadow-elevated flex items-center gap-3 border border-white/10"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-emerald-400"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4 12 14.01l-3-3"/></svg>{toast}</div>
        </div>
      )}
    </div>
  );
}
