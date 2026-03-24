const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
// Configure multer to accept any field names for files
const upload = multer({ 
  dest: path.join(__dirname, 'uploads_temp/'), 
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    // Accept all files
    cb(null, true);
  }
});

app.use(express.json());
// Serve static files with no caching to ensure updates are always loaded
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  }
}));

const DATA_FILE = path.join(__dirname, 'wtd_data.json');

// =====================
// DATA STORAGE - Multi-week
// Structure: { weeks: { "2026-03-10": { week, period, days: {} }, ... } }
// =====================
function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf8');
      if (raw && raw.trim()) {
        const parsed = JSON.parse(raw);
        // Migrate old single-week format to multi-week
        if (parsed.days && !parsed.weeks) {
          const weekKey = parsed.week || 'unknown';
          return { weeks: { [weekKey]: { week: weekKey, period: parsed.period || '', days: parsed.days } } };
        }
        if (parsed.weeks) return parsed;
      }
    }
  } catch(e) {
    console.error('Error loading data:', e.message);
  }
  return { weeks: {} };
}

function saveData(data) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
  } catch(e) {
    console.error('Error saving data:', e.message);
  }
}

// =====================
// TUESDAY-MONDAY WEEK + PERIOD TRACKING
// =====================
function getWeekKey(dateStr) {
  const d = new Date(dateStr + 'T12:00:00Z');
  const day = d.getUTCDay();
  const daysFromTue = (day + 5) % 7;
  const tue = new Date(d);
  tue.setUTCDate(d.getUTCDate() - daysFromTue);
  return tue.toISOString().split('T')[0];
}

function getPeriodWeek(dateStr) {
  const ANCHOR = new Date('2024-01-02T12:00:00Z');
  const weekKey = getWeekKey(dateStr);
  const weekStart = new Date(weekKey + 'T12:00:00Z');
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const weeksSinceAnchor = Math.round((weekStart - ANCHOR) / msPerWeek);
  const totalWeeks = Math.max(0, weeksSinceAnchor);
  const period = Math.floor(totalWeeks / 4) + 1;
  const week = (totalWeeks % 4) + 1;
  return `P${period}W${week}`;
}

function getWeekDateRange(weekKey) {
  if (!weekKey || weekKey === 'unknown') return '';
  try {
    const tue = new Date(weekKey + 'T12:00:00Z');
    const mon = new Date(tue);
    mon.setUTCDate(tue.getUTCDate() + 6);
    const fmt = d => `${d.getUTCMonth()+1}/${d.getUTCDate()}`;
    return `${fmt(tue)}-${fmt(mon)}`;
  } catch(e) { return ''; }
}

// =====================
// EXCEL PARSER
// Col 0=StoreID, Col 3=OnTime%, Col 8=InStore, Col 9=#Del,
// Col 11=Make, Col 13=%<4, Col 15=OvenCut, Col 17=Production,
// Col 19=%<15, Col 24=Rack
// =====================
function parseSOSExcel(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let reportDate = null;
  let reportType = 'daily';

  try {
    const dateCell = raw[1] && raw[1][23];
    if (dateCell instanceof Date) {
      const y = dateCell.getFullYear();
      const m = String(dateCell.getMonth() + 1).padStart(2, '0');
      const d = String(dateCell.getDate()).padStart(2, '0');
      reportDate = `${y}-${m}-${d}`;
    } else if (typeof dateCell === 'string') {
      const m2 = dateCell.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m2) reportDate = `${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`;
      const m1 = dateCell.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (m1) reportDate = m1[0];
    }
  } catch(e) {}

  if (!reportDate) {
    outer: for (let i = 0; i < Math.min(5, raw.length); i++) {
      for (let j = 0; j < Math.min(30, (raw[i] || []).length); j++) {
        const cell = raw[i][j];
        if (cell instanceof Date && cell.getFullYear() > 2020) {
          const y = cell.getFullYear();
          const m = String(cell.getMonth() + 1).padStart(2, '0');
          const d = String(cell.getDate()).padStart(2, '0');
          reportDate = `${y}-${m}-${d}`;
          break outer;
        } else if (typeof cell === 'string') {
          const m2 = cell.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (m2 && parseInt(m2[3]) > 2020) { reportDate = `${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`; break outer; }
          const m1 = cell.match(/(\d{4})-(\d{2})-(\d{2})/);
          if (m1 && parseInt(m1[1]) > 2020) { reportDate = m1[0]; break outer; }
        }
      }
    }
  }

  for (let i = 0; i < Math.min(5, raw.length); i++) {
    for (let j = 0; j < (raw[i] || []).length; j++) {
      if (typeof raw[i][j] === 'string' && raw[i][j].toLowerCase().includes('week')) {
        reportType = 'weekly';
      }
    }
  }

  const stores = [];
  for (let i = 0; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row[0] || typeof row[0] !== 'string') continue;
    if (!row[0].match(/^S0\d{5}$/) && !row[0].match(/^S\d{6}$/)) continue;
    // Check row has valid data (use deliveries or make as indicator)
    const deliveries = parseInt(row[9]) || 0;
    const make = row[11] || null;
    if (!deliveries && !make) continue;
    stores.push({
      store_id:       row[0].trim(),
      on_time:        row[3]  || null,
      // in_store comes ONLY from Above Store PDF, not from SOS Excel
      deliveries:     deliveries,
      make:           make,
      pct_lt4:        row[13] || null,
      oven_cut:       row[15] || null,
      production:     row[17] || null,
      pct_lt15:       row[19] || null,
      rack:           row[24] || null
      // ist_lt19_pct and ist_gt25_count come from PDF only
    });
  }

  console.log(`Excel parsed: ${stores.length} stores, date=${reportDate}`);
  if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} Make=${stores[0].make} pct_lt4=${stores[0].pct_lt4}`);
  return { stores, reportDate, reportType, source: 'sos_excel' };
}

// =====================
// DELIVERY PERFORMANCE REPORT PARSER
// Col 0 = "Store Name (SID)" or area coach name or totals row
// Col 2 = Total Deliveries
// Col 7 = Avg Production Time (decimal minutes)
// Col 8 = Make Time < 4min % (decimal 0-1)
// Col 9 = Production Time < 15min %
// =====================
function parseDeliveryExcel(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let reportDate = null;
  let reportType = 'daily';

  // Extract date from Row1: "Date Range:3/17/2026 - 3/17/2026"
  try {
    const dateRow = raw[1] && raw[1][0];
    if (typeof dateRow === 'string') {
      const m = dateRow.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) reportDate = `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
    }
  } catch(e) {}

  // Check if weekly range
  try {
    const d1 = raw[1] && raw[1][0];
    const d2 = raw[3] && raw[3][0];
    if (d1 && d2 && typeof d1 === 'string' && typeof d2 === 'string') {
      const m1 = d1.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      const m2 = d2.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m1 && m2) {
        const start = new Date(`${m1[3]}-${m1[1].padStart(2,'0')}-${m1[2].padStart(2,'0')}`);
        const end = new Date(`${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`);
        if ((end - start) > 86400000) reportType = 'weekly';
        reportDate = `${m2[3]}-${m2[1].padStart(2,'0')}-${m2[2].padStart(2,'0')}`;
      }
    }
  } catch(e) {}

  const stores = [];
  for (let i = 5; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row[0] || typeof row[0] !== 'string') continue;
    const storeMatch = row[0].match(/\(S(\d{6})\)\s*$/);
    if (!storeMatch) continue;
    const store_id = 'S' + storeMatch[1];

    const avgProdMins = typeof row[7] === 'number' ? row[7] : parseFloat(row[7]);
    const makePct4 = typeof row[8] === 'number' ? row[8] : parseFloat(row[8]);
    const prodPct15 = typeof row[9] === 'number' ? row[9] : parseFloat(row[9]);
    const totalDel = typeof row[2] === 'number' ? row[2] : parseInt(row[2]);

    let productionStr = null;
    if (!isNaN(avgProdMins) && avgProdMins > 0) {
      const mins = Math.floor(avgProdMins);
      const secs = Math.round((avgProdMins - mins) * 60);
      productionStr = `${mins}:${String(secs).padStart(2,'0')}`;
    }

    let pct_lt4_str = null;
    if (!isNaN(makePct4)) pct_lt4_str = (makePct4 * 100).toFixed(1) + '%';

    let pct_lt15_str = null;
    if (!isNaN(prodPct15)) pct_lt15_str = (prodPct15 * 100).toFixed(1) + '%';

    stores.push({
      store_id,
      production: productionStr,
      pct_lt15: pct_lt15_str,
      pct_lt4: pct_lt4_str,
      deliveries: isNaN(totalDel) ? 0 : totalDel,
      _source: 'delivery'
    });
  }

  console.log(`Delivery Excel parsed: ${stores.length} stores, date=${reportDate}`);
  if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} Prod=${stores[0].production} pct_lt15=${stores[0].pct_lt15}`);
  return { stores, reportDate, reportType, source: 'delivery_excel' };
}

// =====================
// ABOVE STORE PDF PARSER (local, no API needed)
// Extracts IST bucket counts: <10, 10-14, 15-18, 19-25, >25
// =====================
const { execSync } = require('child_process');

// =====================
// IST TRACKER BY TERRITORY EXCEL PARSER
// Parses the IST Tracker by Territory format
// Col 0 = Level (STORE, AREA, REGION, TOTAL)
// Col 4 = Store # (store ID as number)
// Col 5 = Store Name
// Col 7 = IST <10 # (count)
// Col 9 = IST 10-14 # (count)
// Col 11 = IST 15-18 # (count)
// Col 13 = IST 19-25 # (count)
// Col 15 = IST >25 # (count)
// Col 16 = Avg IST (mins)
// Col 17 = IST <19 %
// Col 18 = IST >25 %
// =====================
function parseISTTrackerExcel(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let reportDate = null;
  let reportType = 'daily';

  // Try to find date in the file - IST Tracker reports usually don't have explicit dates
  // We'll use today's date or look for it in filename
  try {
    const filename = filePath.split('/').pop();
    const dateMatch = filename.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (dateMatch) {
      reportDate = `${dateMatch[3]}-${dateMatch[1].padStart(2,'0')}-${dateMatch[2].padStart(2,'0')}`;
    }
  } catch(e) {}

  // If no date in filename, use today
  if (!reportDate) {
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    reportDate = `${y}-${m}-${d}`;
  }

  const stores = [];
  for (let i = 0; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row[0] || typeof row[0] !== 'string') continue;
    
    // Only process STORE rows (not AREA, REGION, or TOTAL)
    if (row[0].trim() !== 'STORE') continue;
    
    // Column 4 contains the store number
    const storeNum = row[4];
    if (!storeNum || (typeof storeNum !== 'number' && typeof storeNum !== 'string')) continue;
    
    // Format store ID as S followed by 6 digits
    const store_id = 'S' + String(storeNum).replace(/^S/, '').padStart(6, '0');
    
    // Extract IST bucket counts
    const istLt10 = parseInt(row[7]) || 0;
    const ist10to14 = parseInt(row[9]) || 0;
    const ist15to18 = parseInt(row[11]) || 0;
    const ist19to25 = parseInt(row[13]) || 0;
    const istGt25 = parseInt(row[15]) || 0;
    
    // Calculate total IST orders
    const totalOrders = istLt10 + ist10to14 + ist15to18 + ist19to25 + istGt25;
    
    // Calculate weighted average IST time
    const avgIstmMins = typeof row[16] === 'number' ? row[16] : parseFloat(row[16]) || 0;
    
    // Format avg IST as MM:SS
    let inStore = null;
    if (!isNaN(avgIstmMins) && avgIstmMins > 0) {
      const mins = Math.floor(avgIstmMins);
      const secs = Math.round((avgIstmMins - mins) * 60);
      inStore = `${mins}:${String(secs).padStart(2,'0')}`;
    }
    
    // Calculate IST <19 % and IST >25 %
    const istLt19Pct = typeof row[17] === 'number' ? row[17] : parseFloat(row[17]) || 0;
    const istGt25Pct = typeof row[18] === 'number' ? row[18] : parseFloat(row[18]) || 0;
    
    // Format percentages as strings with %
    const ist_lt19_pct_str = istLt19Pct > 0 ? (istLt19Pct * 100).toFixed(1) + '%' : null;
    const ist_gt25_pct_str = istGt25Pct > 0 ? (istGt25Pct * 100).toFixed(1) + '%' : null;
    
    stores.push({
      store_id: store_id,
      in_store: inStore,
      deliveries: totalOrders,
      ist_lt19_pct: ist_lt19_pct_str,
      ist_gt25_count: istGt25,
      ist_lt10_count: istLt10,
      ist_10_14_count: ist10to14,
      ist_15_18_count: ist15to18,
      ist_19_25_count: ist19to25,
      _source: 'ist_tracker'
    });
  }

  console.log(`IST Tracker Excel parsed: ${stores.length} stores, date=${reportDate}`);
  if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} InStore=${stores[0].in_store} Lt19%=${stores[0].ist_lt19_pct}`);
  return { stores, reportDate, reportType, source: 'ist_tracker' };
}

function parseAboveStorePDFLocal(filePath) {
  try {
    const text = execSync(`pdftotext "${filePath}" -`, { maxBuffer: 20 * 1024 * 1024 }).toString();
    const storeBlocks = text.split(/~+\s*\n\s*\nStore:/);
    const stores = [];

    for (const block of storeBlocks.slice(1)) {
      const storeMatch = block.trim().match(/^(S?\d+)\s+\(([^)]+)\)/);
      if (!storeMatch) continue;
      const rawId = storeMatch[1];
      const store_id = 'S' + rawId.replace(/^S/, '').padStart(6, '0');

      const dateMatch = block.match(/For Bus\.Date\s+\S+-(\d{2}\/\d{2}\/\d{2})/);
      let reportDate = null;
      if (dateMatch) {
        const parts = dateMatch[1].split('/');
        const yr = parseInt(parts[2]) < 100 ? 2000 + parseInt(parts[2]) : parseInt(parts[2]);
        reportDate = `${yr}-${parts[0].padStart(2,'0')}-${parts[1].padStart(2,'0')}`;
      }

      // Extract IST counts from the section after the 5 colons
      const colonBlock = block.match(/:\s*\n:\s*\n:\s*\n:\s*\n:\s*\n\s*\n([\s\S]+?)(?:Orders per Dispatch|Averages:|Cash controls)/);
      if (!colonBlock) continue;

      const afterColons = colonBlock[1];
      const lines = afterColons.trim().split('\n');
      const counts = [];
      for (const line of lines) {
        const l = line.trim();
        if (!l) continue;
        // "6 100.00%" or just "6" - extract leading integer
        const m1 = l.match(/^(-?\d+)(?:\s+[\d.]+%)?$/);
        if (m1) { counts.push(parseInt(m1[1])); continue; }
        // Pure percentage - skip
        if (l.match(/^[\d.]+%$/)) continue;
        // "6 100.00%" inline
        const m2 = l.match(/^(-?\d+)\s+[\d.]+%/);
        if (m2) { counts.push(parseInt(m2[1])); }
      }

      if (counts.length < 5) continue;
      const [ist_lt10, ist_1014, ist_1518, ist_1925, ist_gt25] = counts;
      const total = ist_lt10 + ist_1014 + ist_1518 + ist_1925 + ist_gt25;
      // Compute lt19 pct from actual counts
      const ist_lt19_pct = total > 0 ? parseFloat(((ist_lt10 + ist_1014 + ist_1518) / total * 100).toFixed(1)) : 0;

      // Extract ist_avg from "Averages:" section
      let ist_avg = null;
      const avgMatch = block.match(/Averages:\s*\n\s*([\d.]+)/);
      if (avgMatch) {
        ist_avg = parseFloat(avgMatch[1]);
      }
      // Fallback: estimate avg from bucket midpoints if not found
      if (!ist_avg && total > 0) {
        const weightedSum = (ist_lt10 * 8) + (ist_1014 * 12) + (ist_1518 * 16.5) + (ist_1925 * 22) + (ist_gt25 * 27);
        ist_avg = Math.round((weightedSum / total) * 10) / 10;
      }

      stores.push({
        store_id,
        reportDate,
        ist_lt10, ist_1014, ist_1518, ist_1925, ist_gt25,
        ist_lt19_pct,
        ist_gt25_count: ist_gt25,
        total_orders: total,
        ist_avg
      });
    }

    // Use most common date as report date
    const dates = stores.map(s => s.reportDate).filter(Boolean);
    const dateCounts = {};
    dates.forEach(d => { dateCounts[d] = (dateCounts[d] || 0) + 1; });
    const reportDate = Object.entries(dateCounts).sort((a,b) => b[1]-a[1])[0]?.[0] || null;

    console.log(`Above Store PDF parsed: ${stores.length} stores, date=${reportDate}`);
    if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} 19-25:${stores[0].ist_1925} >25:${stores[0].ist_gt25}`);
    return { stores, reportDate, reportType: 'daily', source: 'pdf' };
  } catch(e) {
    console.error('PDF parse error:', e.message);
    return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: e.message };
  }
}

// =====================
// PDF PARSER via Claude API (fallback)
// =====================
async function parseAboveStorePDF(filePath) {
  try {
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: 'PDF parsing requires ANTHROPIC_API_KEY' };
    const fileBuffer = fs.readFileSync(filePath);
    const base64Data = fileBuffer.toString('base64');
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-3-5-sonnet-20241022', max_tokens: 4000,
        messages: [{ role: 'user', content: [
          { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64Data } },
          { type: 'text', text: `Extract Speed of Service data from this Pizza Hut Above Store report. Return ONLY a JSON array with no markdown, no backticks, no explanation. Each object must have these exact keys:\n{\n  "store_id": "S039xxx",\n  "bus_date": "MM/DD/YYYY",\n  "total_orders": number,\n  "ist_avg": number,\n  "ist_lt10_count": number, "ist_lt10_pct": number,\n  "ist_1014_count": number, "ist_1014_pct": number,\n  "ist_1518_count": number, "ist_1518_pct": number,\n  "ist_1925_count": number, "ist_1925_pct": number,\n  "ist_gt25_count": number, "ist_gt25_pct": number,\n  "ist_lt19_count": number, "ist_lt19_pct": number\n}\nReturn [] if no store data found.` }
        ]}]
      })
    });
    if (!response.ok) return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: `Claude API error ${response.status}` };
    const apiResult = await response.json();
    const rawText = apiResult.content?.map(c => c.text || '').join('').trim();
    let cleaned = rawText.replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();
    const jsonMatch = cleaned.match(/\[[\s\S]*\]/);
    if (jsonMatch) cleaned = jsonMatch[0];
    let stores;
    try { stores = JSON.parse(cleaned); } catch(e) { return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: 'Failed to parse Claude response' }; }
    if (!Array.isArray(stores)) return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: 'Non-array response' };
    let reportDate = null;
    for (const s of stores) {
      if (s.bus_date) {
        const parts = s.bus_date.split('/');
        if (parts.length === 3) { reportDate = `${parts[2].length===2?'20'+parts[2]:parts[2]}-${parts[0].padStart(2,'0')}-${parts[1].padStart(2,'0')}`; break; }
      }
    }
    stores.forEach(s => {
      const bp = (s.ist_lt10_pct||0)+(s.ist_1014_pct||0)+(s.ist_1518_pct||0);
      if (!s.ist_lt19_pct || s.ist_lt19_pct===0) s.ist_lt19_pct = parseFloat(bp.toFixed(1));
    });
    return { stores, reportDate, reportType: 'daily', source: 'pdf' };
  } catch(e) {
    return { stores: [], reportDate: null, reportType: 'daily', source: 'pdf', error: e.message };
  }
}

// =====================
// UPLOAD
// =====================
app.post('/api/upload', upload.any(), async (req, res) => {
  const tempFiles = req.files ? req.files.map(f => f.path) : [];
  try {
    if (!req.files || !req.files.length) return res.status(400).json({ error: 'No files received' });
    const uploaderName = req.body.uploaderName || 'Unknown';
    let allData = loadData();
    const results = [], errors = [];

    for (const file of req.files) {
      const isExcel = file.originalname.match(/\.xlsx?$/i);
      const isPdf = file.originalname.match(/\.pdf$/i);
      let parsed = null;
      try {
        if (isExcel) {
          // Detect file type by peeking at content
          const wb2 = XLSX.readFile(file.path, { cellDates: false });
          const ws2 = wb2.Sheets[wb2.SheetNames[0]];
          const raw2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null });
          const firstCell = raw2[0] && raw2[0][0];
          const firstRow = raw2[0] && raw2[0].join(' ');
          
          if (typeof firstCell === 'string' && firstCell.includes('Delivery Performance')) {
            parsed = parseDeliveryExcel(file.path);
          } else if (typeof firstRow === 'string' && (firstRow.includes('IST <10 #') || firstRow.includes('IST Tracker'))) {
            // IST Tracker by Territory format
            parsed = parseISTTrackerExcel(file.path);
          } else {
            parsed = parseSOSExcel(file.path);
          }
        }
        else if (isPdf) {
          // Try local pdftotext parser first (faster, no API needed)
          parsed = parseAboveStorePDFLocal(file.path);
          // Fall back to Claude API if local parse gets very few stores
          if (!parsed.stores || parsed.stores.length < 10) {
            parsed = await parseAboveStorePDF(file.path);
          }
        }
        else { errors.push(`${file.originalname}: unsupported file type`); continue; }
      } catch(e) { errors.push(`${file.originalname}: ${e.message}`); continue; }

      if (!parsed) { errors.push(`${file.originalname}: no data returned`); continue; }
      if (parsed.error) { errors.push(`${file.originalname}: ${parsed.error}`); if (!parsed.stores?.length) continue; }
      if (!parsed.stores || !parsed.stores.length) { errors.push(`${file.originalname}: no store data found`); continue; }
      if (!parsed.stores || !Array.isArray(parsed.stores)) { 
        errors.push(`${file.originalname}: invalid store data format`); 
        console.log('Invalid stores:', parsed);
        continue; 
      }

      const { stores, reportDate, reportType } = parsed;
      const finalDate = reportDate || new Date().toISOString().split('T')[0];
      
      if (!finalDate) {
        errors.push(`${file.originalname}: Could not determine report date`);
        continue;
      }
      
      const weekKey = getWeekKey(finalDate);
      const periodWeek = getPeriodWeek(finalDate);

      // NEVER wipe existing weeks - just add/update
      if (!allData.weeks[weekKey]) {
        allData.weeks[weekKey] = { week: weekKey, period: periodWeek, days: {} };
      }
      if (!allData.weeks[weekKey].days[finalDate]) {
        allData.weeks[weekKey].days[finalDate] = { date: finalDate, type: reportType, uploader: uploaderName, stores: [] };
      }

      const existing = {};
      if (allData.weeks[weekKey].days[finalDate].stores) {
        allData.weeks[weekKey].days[finalDate].stores.forEach(s => { existing[s.store_id] = s; });
      }
      stores.forEach(s => {
        const align = ALIGNMENT[s.store_id];
        if (parsed.source === 'delivery_excel') {
          // Merge only delivery-specific fields; don't overwrite SOS fields
          if (existing[s.store_id]) {
            existing[s.store_id].production = s.production;
            existing[s.store_id].pct_lt15 = s.pct_lt15;
            if (!existing[s.store_id].deliveries || existing[s.store_id].deliveries === 0) {
              existing[s.store_id].deliveries = s.deliveries;
            }
          } else if (align) {
            // SOS not yet uploaded for this store - create stub with delivery data
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name,
              area: align.area_coach,
              area_coach: align.area_coach,
              region_coach: align.region_coach,
              production: s.production,
              pct_lt15: s.pct_lt15,
              deliveries: s.deliveries,
              in_store: null, make: null, pct_lt4: s.pct_lt4, on_time: null, rack: null,
              ist_lt19_pct: null, ist_gt25_count: 0
            };
          }
        } else if (parsed.source === 'pdf') {
          // PDF: merge IST bucket counts + ist_avg (in_store) into existing store data
          if (existing[s.store_id]) {
            existing[s.store_id].ist_lt10 = s.ist_lt10;
            existing[s.store_id].ist_1014 = s.ist_1014;
            existing[s.store_id].ist_1518 = s.ist_1518;
            existing[s.store_id].ist_1925 = s.ist_1925;
            existing[s.store_id].ist_gt25 = s.ist_gt25;
            existing[s.store_id].ist_gt25_count = s.ist_gt25;
            existing[s.store_id].ist_lt19_pct = s.ist_lt19_pct;
            // Use PDF in-store average as the primary in_store value
            if (s.ist_avg) existing[s.store_id].in_store = s.ist_avg;
          } else if (align) {
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name, area: align.area_coach, area_coach: align.area_coach,
              region_coach: align.region_coach,
              ist_lt10: s.ist_lt10, ist_1014: s.ist_1014, ist_1518: s.ist_1518,
              ist_1925: s.ist_1925, ist_gt25: s.ist_gt25,
              ist_gt25_count: s.ist_gt25, ist_lt19_pct: s.ist_lt19_pct,
              in_store: s.ist_avg || null, make: null, production: null, pct_lt4: null,
              pct_lt15: null, on_time: null, rack: null, deliveries: 0
            };
          }
        } else if (parsed.source === 'sos_excel') {
          // SOS Excel: merge only SOS-specific fields; don't overwrite PDF IST fields
          if (existing[s.store_id]) {
            existing[s.store_id].on_time = s.on_time;
            existing[s.store_id].deliveries = s.deliveries;
            existing[s.store_id].make = s.make;
            existing[s.store_id].pct_lt4 = s.pct_lt4;
            existing[s.store_id].oven_cut = s.oven_cut;
            existing[s.store_id].production = s.production;
            existing[s.store_id].pct_lt15 = s.pct_lt15;
            existing[s.store_id].rack = s.rack;
            // DO NOT overwrite in_store, ist_lt19_pct, ist_gt25_count - those come from PDF
          } else if (align) {
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name,
              area: align.area_coach,
              area_coach: align.area_coach,
              region_coach: align.region_coach,
              on_time: s.on_time,
              deliveries: s.deliveries,
              make: s.make,
              pct_lt4: s.pct_lt4,
              oven_cut: s.oven_cut,
              production: s.production,
              pct_lt15: s.pct_lt15,
              rack: s.rack,
              // IST fields are null until PDF is uploaded
              in_store: null,
              ist_lt10: null, ist_1014: null, ist_1518: null,
              ist_1925: null, ist_gt25: null,
              ist_gt25_count: null, ist_lt19_pct: null
            };
          }
        } else {
          if (!align) return;
          existing[s.store_id] = { ...existing[s.store_id], ...s, ...align };
        }
      });
      allData.weeks[weekKey].days[finalDate].stores = Object.values(existing);

      results.push({ file: file.originalname, date: finalDate, week: weekKey, period: periodWeek, type: reportType, storeCount: allData.weeks[weekKey].days[finalDate].stores.length, source: parsed.source });
    }

    tempFiles.forEach(f => { try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch(e) {} });
    saveData(allData);
    return res.json({ success: true, results, errors, data: computeAllWeeks(allData) });
  } catch(e) {
    console.error('Upload error:', e);
    console.error('Error stack:', e.stack);
    tempFiles.forEach(f => { try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch(e2) {} });
    return res.status(500).json({ error: e.message || 'Server error', stack: e.stack });
  }
});

app.get('/api/data', (req, res) => {
  try {
    const allData = loadData();
    res.json({ data: computeAllWeeks(allData) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/reset', (req, res) => {
  saveData({ weeks: {} });
  res.json({ success: true });
});

app.post('/api/ai', async (req, res) => {
  try {
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(400).json({ error: 'AI features require ANTHROPIC_API_KEY' });
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-3-5-sonnet-20241022', max_tokens: 1000, system: 'You are Velocity, a Pizza Hut Speed of Service intelligence agent. Be direct, specific, action-oriented. No corporate fluff. Plain text only.', messages: req.body.messages })
    });
    res.json(await response.json());
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// =====================
// HELPERS
// =====================
function parseTimeToMins(s) {
  if (!s || s === 'NaN' || s === '0:00' || s === '—') return 0;
  const neg = s.startsWith('-');
  const clean = s.replace('-', '').split(':');
  return (neg ? -1 : 1) * ((parseInt(clean[0]) || 0) + (parseInt(clean[1]) || 0) / 60);
}
function parsePct(s) { if (!s) return 0; return parseFloat(String(s).replace('%', '')) || 0; }
function fmtTime(m) {
  if (!m || m <= 0) return '0:00';
  const i = Math.floor(Math.abs(m));
  const s = Math.round((Math.abs(m) - i) * 60);
  return `${m < 0 ? '-' : ''}${i}:${String(s).padStart(2, '0')}`;
}

function computeWeek(wtd) {
  const days = Object.values(wtd.days || {}).sort((a, b) => a.date.localeCompare(b.date));
  if (!days.length) return { days: [], stores: [], week: wtd.week, period: wtd.period || '', dateRange: getWeekDateRange(wtd.week) };

  const allIds = new Set();
  days.forEach(d => {
    if (d.stores && Array.isArray(d.stores)) {
      d.stores.forEach(s => allIds.add(s.store_id));
    }
  });

  const storeWTD = [];
  for (const sid of allIds) {
    const dd = days.map(d => d.stores.find(s => s.store_id === sid)).filter(Boolean);
    if (!dd.length) continue;
    const align = ALIGNMENT[sid] || {};
    if (!align.name) continue;

    const avgIST = dd.reduce((a, s) => a + (s.ist_avg || s.in_store || 0), 0) / dd.length;
    const avgMk = dd.reduce((a, s) => a + parseTimeToMins(s.make), 0) / dd.length;
    const avgPr = dd.reduce((a, s) => a + parseTimeToMins(s.production), 0) / dd.length;
    const avgPct4 = dd.reduce((a, s) => a + parsePct(s.pct_lt4), 0) / dd.length;
    const avgPct15 = dd.reduce((a, s) => a + parsePct(s.pct_lt15), 0) / dd.length;
    const avgOT = dd.reduce((a, s) => a + parsePct(s.on_time), 0) / dd.length;

    const avgLt19 = dd.reduce((a, s) => {
      let val = 0;
      if (typeof s.ist_lt19_pct === 'number' && s.ist_lt19_pct > 0) val = s.ist_lt19_pct;
      else if (typeof s.ist_lt19_pct === 'string') val = parseFloat(s.ist_lt19_pct) || 0;
      if (val === 0) val = (s.ist_lt10_pct||0)+(s.ist_1014_pct||0)+(s.ist_1518_pct||0);
      if (val === 0 && (s.ist_avg || s.in_store)) {
        const ist = s.ist_avg || s.in_store;
        val = ist <= 19 ? 85 : ist <= 22 ? 60 : ist <= 25 ? 40 : 25;
      }
      return a + val;
    }, 0) / dd.length;

    const avgGt25Count = dd.reduce((a, s) => a + (s.ist_gt25_count || s.ist_gt25 || 0), 0) / dd.length;
    const totalDel = dd.reduce((a, s) => a + (s.total_orders || s.deliveries || 0), 0);
    const last = dd[dd.length - 1];

    // Sum IST bucket counts across all days (from PDF uploads)
    const totalLt10 = dd.reduce((a, s) => a + (s.ist_lt10 || 0), 0);
    const total1014 = dd.reduce((a, s) => a + (s.ist_1014 || 0), 0);
    const total1518 = dd.reduce((a, s) => a + (s.ist_1518 || 0), 0);
    const total1925 = dd.reduce((a, s) => a + (s.ist_1925 || 0), 0);
    const totalGt25 = dd.reduce((a, s) => a + (s.ist_gt25 || s.ist_gt25_count || 0), 0);

    const daily = {};
    days.forEach(d => {
      if (!d.stores || !Array.isArray(d.stores)) return;
      const s = d.stores.find(s => s.store_id === sid);
      if (s) daily[d.date] = {
        in_store: s.ist_avg||s.in_store, make: s.make, production: s.production,
        on_time: s.on_time, deliveries: s.total_orders||s.deliveries,
        pct_lt4: s.pct_lt4, pct_lt15: s.pct_lt15, ist_lt19_pct: s.ist_lt19_pct,
        ist_gt25_count: s.ist_gt25_count||s.ist_gt25||0,
        ist_lt10: s.ist_lt10||0, ist_1014: s.ist_1014||0, ist_1518: s.ist_1518||0,
        ist_1925: s.ist_1925||0, ist_gt25: s.ist_gt25||0
      };
    });

    let lastLt19 = last.ist_lt19_pct;
    if (!lastLt19 || lastLt19 === 0) lastLt19 = (last.ist_lt10_pct||0)+(last.ist_1014_pct||0)+(last.ist_1518_pct||0);
    if (!lastLt19 || lastLt19 === 0) { const ist = last.ist_avg||last.in_store; lastLt19 = ist<=19?85:ist<=22?60:ist<=25?40:25; }

    storeWTD.push({
      store_id: sid, name: align.name, area: align.area_coach, area_coach: align.area_coach, region_coach: align.region_coach,
      days_reported: dd.length,
      wtd_in_store: Math.round(avgIST*10)/10, wtd_make: fmtTime(avgMk), wtd_pct_lt4: avgPct4.toFixed(1)+'%',
      wtd_production: fmtTime(avgPr), wtd_pct_lt15: avgPct15.toFixed(1)+'%', wtd_on_time: avgOT.toFixed(1)+'%',
      wtd_deliveries: totalDel, wtd_lt19_pct: avgLt19.toFixed(1), wtd_gt25_avg: avgGt25Count.toFixed(1),
      wtd_ist_lt10: totalLt10, wtd_ist_1014: total1014, wtd_ist_1518: total1518,
      wtd_ist_1925: total1925, wtd_ist_gt25: totalGt25,
      in_store: last.ist_avg||last.in_store, make: last.make, pct_lt4: last.pct_lt4,
      production: last.production, pct_lt15: last.pct_lt15, on_time: last.on_time,
      deliveries: last.total_orders||last.deliveries,
      ist_lt19_pct: typeof lastLt19==='number' ? lastLt19 : parseFloat(lastLt19)||0,
      ist_gt25_count: last.ist_gt25_count||last.ist_gt25||0,
      ist_lt10: last.ist_lt10||0, ist_1014: last.ist_1014||0, ist_1518: last.ist_1518||0,
      ist_1925: last.ist_1925||0, ist_gt25: last.ist_gt25||0,
      daily
    });
  }

  return {
    week: wtd.week, period: wtd.period||'', dateRange: getWeekDateRange(wtd.week),
    days: days.map(d => ({ date: d.date, type: d.type, storeCount: d.stores.length, uploader: d.uploader })),
    stores: storeWTD
  };
}

function computeAllWeeks(allData) {
  const weeks = allData.weeks || {};
  const weekKeys = Object.keys(weeks).sort((a, b) => b.localeCompare(a)); // newest first
  return weekKeys.map(k => computeWeek(weeks[k]));
}

const ALIGNMENT = {"S038876":{"name":"Senoia","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039377":{"name":"Griffin","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039378":{"name":"Union City","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039379":{"name":"Jefferson St","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039384":{"name":"Newnan","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039454":{"name":"Zebulon","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039465":{"name":"Senoia Rd","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039383":{"name":"Stockbridge","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039388":{"name":"Jonesboro Rd","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039393":{"name":"Lovejoy","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039429":{"name":"Ola","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039461":{"name":"County Line","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039513":{"name":"Jodeco","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039521":{"name":"Kellytown","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039522":{"name":"Ellenwood","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039375":{"name":"Bells Ferry Rd","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039376":{"name":"CrossRds","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039382":{"name":"Glade Rd","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039387":{"name":"Kennesaw","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039392":{"name":"Towne Lake","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039462":{"name":"Acworth/Emerson","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039380":{"name":"Windy Hill","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039386":{"name":"Powder Springs","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039389":{"name":"Lithia Springs","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039410":{"name":"Mableton","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039451":{"name":"Bolton","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039525":{"name":"Smyrna","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039527":{"name":"Austell Rd","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039412":{"name":"Miracle Strip","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039413":{"name":"Navarre","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039414":{"name":"Gulf Breeze","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039415":{"name":"Miramar Beach","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039416":{"name":"Niceville","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039430":{"name":"Racetrack","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039529":{"name":"Crestview","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039381":{"name":"Fairburn Rd","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039385":{"name":"Ridge Rd","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039390":{"name":"East Paulding","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039391":{"name":"Hwy 5","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039526":{"name":"Dallas","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039417":{"name":"Collinsville","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039419":{"name":"Martinsville","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039421":{"name":"College Rd","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039424":{"name":"Gate City Blvd","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039427":{"name":"Pyramid Village","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039436":{"name":"Battleground","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039457":{"name":"E. Greensboro","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039418":{"name":"Riverside Dr","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039422":{"name":"South Church","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039423":{"name":"Graham","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039432":{"name":"Mebane","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039433":{"name":"Elton Way","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039455":{"name":"Spring Garden","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039456":{"name":"Whitsett","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039420":{"name":"Harrisonburg","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039425":{"name":"Elkton","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039426":{"name":"Woodstock","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039428":{"name":"Stuarts Draft","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039431":{"name":"Staunton","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039435":{"name":"Shoppers World","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039450":{"name":"Orange","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039453":{"name":"JMU/Market","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039466":{"name":"Waynesboro","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039400":{"name":"E Palmetto","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039401":{"name":"Darlington","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039402":{"name":"2nd Loop","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039403":{"name":"Marion","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039394":{"name":"Elberton","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039395":{"name":"Abbeville","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039396":{"name":"Hartwell","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039398":{"name":"Royston","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039399":{"name":"Lavonia","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039404":{"name":"Greenwood Bypass","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039405":{"name":"Simpsonville","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039407":{"name":"Newberry","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039408":{"name":"Seneca","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S040090":{"name":"Main","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040091":{"name":"Silver City","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040093":{"name":"Missouri","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040102":{"name":"Deming","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S039180":{"name":"Zaragosa","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039182":{"name":"Vista","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039185":{"name":"Gateway","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039318":{"name":"Socorro","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039323":{"name":"Tierre Este","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S041651":{"name":"Eastlake","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S040082":{"name":"Taylor Ranch","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040084":{"name":"7th/Lomas","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040101":{"name":"Washington/Zuni","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040107":{"name":"Coors/Barcelona","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040108":{"name":"Wyoming/Harper","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040111":{"name":"303 Coors","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S038729":{"name":"Kenworthy","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039174":{"name":"University","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039175":{"name":"Airway","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039178":{"name":"CrossRds EP","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039192":{"name":"Resler","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039324":{"name":"Outlet Mall","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039448":{"name":"Dyer","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S040088":{"name":"Los Lunas","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040096":{"name":"Belen","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040099":{"name":"Candelaria","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040100":{"name":"T or C","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040110":{"name":"Bull Chicks","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S039589":{"name":"Rio Rancho","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040094":{"name":"Villa Linda Mall","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040104":{"name":"Southern","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040105":{"name":"Las Vegas","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040106":{"name":"Espanola","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040109":{"name":"Unser & McMahon","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S039173":{"name":"Yarbrough","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039176":{"name":"Lovington","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039177":{"name":"Hobbs","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039179":{"name":"George Dieter","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039188":{"name":"Carlsbad","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039518":{"name":"Hobbs North","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039530":{"name":"Montana","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S040083":{"name":"20th St","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040085":{"name":"North Gallup","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040086":{"name":"Main Street","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040087":{"name":"East Gallup","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040092":{"name":"Aztec","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040112":{"name":"Durango","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"}};

// Enable CORS for all routes
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`Velocity running on port ${PORT}`));
