const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { FISCAL_CALENDAR_2026, getPeriodForDate, getDatesInPeriod, getWeeksInPeriod } = require('./fiscalCalendar');

const app = express();
// Configure multer to accept any field names for files
const upload = multer({ 
  dest: path.join(__dirname, 'uploads_temp/'), 
  limits: { fileSize: 50 * 1024 * 1024, fieldSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    // Accept all files
    cb(null, true);
  }
});

app.use(express.json());

// Enable CORS for all routes - MUST be before any route definitions
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Serve static files with no caching to ensure updates are always loaded
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  }
}));

// =====================
// PERSISTENT STORAGE SETUP
// Uses Render Disk at /var/data if available, falls back to local directory.
// On first run with a fresh disk, seeds data from the bundled wtd_data.json.
// =====================
const DISK_PATH = process.env.RENDER_DISK_PATH || '/var/data';
const DISK_DATA_FILE = path.join(DISK_PATH, 'wtd_data.json');
const BUNDLED_DATA_FILE = path.join(__dirname, 'wtd_data.json');

function getDataFilePath() {
  try {
    if (!fs.existsSync(DISK_PATH)) fs.mkdirSync(DISK_PATH, { recursive: true });
    // Seed if disk file missing, empty, or has no weeks (e.g. from a failed deploy)
    let needsSeed = !fs.existsSync(DISK_DATA_FILE);
    if (!needsSeed) {
      try {
        const raw = fs.readFileSync(DISK_DATA_FILE, 'utf8');
        const parsed = JSON.parse(raw);
        if (!parsed.weeks || Object.keys(parsed.weeks).length === 0) needsSeed = true;
      } catch(e) { needsSeed = true; }
    }
    if (needsSeed && fs.existsSync(BUNDLED_DATA_FILE)) {
      console.log('Seeding persistent disk from bundled wtd_data.json...');
      fs.copyFileSync(BUNDLED_DATA_FILE, DISK_DATA_FILE);
      console.log('Disk seeded successfully.');
    }
    return DISK_DATA_FILE;
  } catch(e) {
    console.warn('Disk path not available, using local file:', e.message);
    return BUNDLED_DATA_FILE;
  }
}

const DATA_FILE = getDataFilePath();

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
        if (parsed.weeks) {
          // Ensure all stores have ist_avg calculated from bucket data
          ensureISTFromBuckets(parsed);
          return parsed;
        }
      }
    }
  } catch(e) {
    console.error('Error loading data:', e.message);
  }
  return { weeks: {} };
}

// Calculate IST average from bucket distribution for stores missing ist_avg
function ensureISTFromBuckets(data) {
  if (!data.weeks) return;
  
  for (const weekKey of Object.keys(data.weeks)) {
    const week = data.weeks[weekKey];
    
    // Fix week-level stores
    if (week.stores) {
      for (const store of week.stores) {
        if (store.ist_avg === null || store.ist_avg === undefined) {
          const ist = calculateISTFromBuckets(store);
          if (ist !== null) {
            store.ist_avg = ist;
            if (!store.wtd_in_store) store.wtd_in_store = ist;
          }
        }
      }
    }
    
    // Fix daily stores
    if (week.days) {
      for (const dayKey of Object.keys(week.days)) {
        const day = week.days[dayKey];
        if (day.stores) {
          for (const store of day.stores) {
            if (store.ist_avg === null || store.ist_avg === undefined) {
              const ist = calculateISTFromBuckets(store);
              if (ist !== null) {
                store.ist_avg = ist;
              }
            }
          }
        }
      }
    }
  }
}

// Calculate IST average from bucket distribution
function calculateISTFromBuckets(store) {
  const lt10 = store.ist_lt10 || 0;
  const t1014 = store.ist_1014 || 0;
  const t1518 = store.ist_1518 || 0;
  const t1925 = store.ist_1925 || 0;
  const gt25 = store.ist_gt25 || 0;
  const total = lt10 + t1014 + t1518 + t1925 + gt25;
  
  if (total === 0) return null;
  
  // Approximate midpoints for each bucket
  const avgIST = (lt10 * 8 + t1014 * 12 + t1518 * 16.5 + t1925 * 22 + gt25 * 30) / total;
  return Math.round(avgIST * 10) / 10;
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

const FISCAL_CALENDAR = {
  "2025-12-30": "P1W1",
  "2026-01-06": "P1W2",
  "2026-01-13": "P1W3",
  "2026-01-20": "P1W4",
  "2026-01-27": "P2W1",
  "2026-02-03": "P2W2",
  "2026-02-10": "P2W3",
  "2026-02-17": "P2W4",
  "2026-02-24": "P3W1",
  "2026-03-03": "P3W2",
  "2026-03-10": "P3W3",
  "2026-03-17": "P3W4",
  "2026-03-24": "P4W1",
  "2026-03-31": "P4W2",
  "2026-04-07": "P4W3",
  "2026-04-14": "P4W4",
  "2026-04-21": "P5W1",
  "2026-04-28": "P5W2",
  "2026-05-05": "P5W3",
  "2026-05-12": "P5W4",
  "2026-05-19": "P6W1",
  "2026-05-26": "P6W2",
  "2026-06-02": "P6W3",
  "2026-06-09": "P6W4",
  "2026-06-16": "P7W1",
  "2026-06-23": "P7W2",
  "2026-06-30": "P7W3",
  "2026-07-07": "P7W4",
  "2026-07-14": "P8W1",
  "2026-07-21": "P8W2",
  "2026-07-28": "P8W3",
  "2026-08-04": "P8W4",
  "2026-08-11": "P9W1",
  "2026-08-18": "P9W2",
  "2026-08-25": "P9W3",
  "2026-09-01": "P9W4",
  "2026-09-08": "P10W1",
  "2026-09-15": "P10W2",
  "2026-09-22": "P10W3",
  "2026-09-29": "P10W4",
  "2026-10-06": "P11W1",
  "2026-10-13": "P11W2",
  "2026-10-20": "P11W3",
  "2026-10-27": "P11W4",
  "2026-11-03": "P12W1",
  "2026-11-10": "P12W2",
  "2026-11-17": "P12W3",
  "2026-11-24": "P12W4",
  "2026-12-01": "P13W1",
  "2026-12-08": "P13W2",
  "2026-12-15": "P13W3",
  "2026-12-22": "P13W4",
  "2026-12-29": "P1W1",
  "2027-01-05": "P1W2",
  "2027-01-12": "P1W3",
  "2027-01-19": "P1W4",
  "2027-01-26": "P2W1",
  "2027-02-02": "P2W2",
  "2027-02-09": "P2W3",
  "2027-02-16": "P2W4",
  "2027-02-23": "P3W1",
  "2027-03-02": "P3W2",
  "2027-03-09": "P3W3",
  "2027-03-16": "P3W4",
  "2027-03-23": "P4W1",
  "2027-03-30": "P4W2",
  "2027-04-06": "P4W3",
  "2027-04-13": "P4W4",
  "2027-04-20": "P5W1",
  "2027-04-27": "P5W2",
  "2027-05-04": "P5W3",
  "2027-05-11": "P5W4",
  "2027-05-18": "P6W1",
  "2027-05-25": "P6W2",
  "2027-06-01": "P6W3",
  "2027-06-08": "P6W4",
  "2027-06-15": "P7W1",
  "2027-06-22": "P7W2",
  "2027-06-29": "P7W3",
  "2027-07-06": "P7W4",
  "2027-07-13": "P8W1",
  "2027-07-20": "P8W2",
  "2027-07-27": "P8W3",
  "2027-08-03": "P8W4",
  "2027-08-10": "P9W1",
  "2027-08-17": "P9W2",
  "2027-08-24": "P9W3",
  "2027-08-31": "P9W4",
  "2027-09-07": "P10W1",
  "2027-09-14": "P10W2",
  "2027-09-21": "P10W3",
  "2027-09-28": "P10W4",
  "2027-10-05": "P11W1",
  "2027-10-12": "P11W2",
  "2027-10-19": "P11W3",
  "2027-10-26": "P11W4",
  "2027-11-02": "P12W1",
  "2027-11-09": "P12W2",
  "2027-11-16": "P12W3",
  "2027-11-23": "P12W4",
  "2027-11-30": "P13W1",
  "2027-12-07": "P13W2",
  "2027-12-14": "P13W3",
  "2027-12-21": "P13W4",
  "2027-12-28": "P1W1",
  "2028-01-04": "P1W2",
  "2028-01-11": "P1W3",
  "2028-01-18": "P1W4",
  "2028-01-25": "P2W1",
  "2028-02-01": "P2W2",
  "2028-02-08": "P2W3",
  "2028-02-15": "P2W4",
  "2028-02-22": "P3W1",
  "2028-02-29": "P3W2",
  "2028-03-07": "P3W3",
  "2028-03-14": "P3W4",
  "2028-03-21": "P4W1",
  "2028-03-28": "P4W2",
  "2028-04-04": "P4W3",
  "2028-04-11": "P4W4",
  "2028-04-18": "P5W1",
  "2028-04-25": "P5W2",
  "2028-05-02": "P5W3",
  "2028-05-09": "P5W4",
  "2028-05-16": "P6W1",
  "2028-05-23": "P6W2",
  "2028-05-30": "P6W3",
  "2028-06-06": "P6W4",
  "2028-06-13": "P7W1",
  "2028-06-20": "P7W2",
  "2028-06-27": "P7W3",
  "2028-07-04": "P7W4",
  "2028-07-11": "P8W1",
  "2028-07-18": "P8W2",
  "2028-07-25": "P8W3",
  "2028-08-01": "P8W4",
  "2028-08-08": "P9W1",
  "2028-08-15": "P9W2",
  "2028-08-22": "P9W3",
  "2028-08-29": "P9W4",
  "2028-09-05": "P10W1",
  "2028-09-12": "P10W2",
  "2028-09-19": "P10W3",
  "2028-09-26": "P10W4",
  "2028-10-03": "P11W1",
  "2028-10-10": "P11W2",
  "2028-10-17": "P11W3",
  "2028-10-24": "P11W4",
  "2028-10-31": "P12W1",
  "2028-11-07": "P12W2",
  "2028-11-14": "P12W3",
  "2028-11-21": "P12W4",
  "2028-11-28": "P13W1",
  "2028-12-05": "P13W2",
  "2028-12-12": "P13W3",
  "2028-12-19": "P13W4",
  "2028-12-26": "P1W1",
  "2029-01-02": "P1W2",
  "2029-01-09": "P1W3",
  "2029-01-16": "P1W4",
  "2029-01-23": "P2W1",
  "2029-01-30": "P2W2",
  "2029-02-06": "P2W3",
  "2029-02-13": "P2W4",
  "2029-02-20": "P3W1",
  "2029-02-27": "P3W2",
  "2029-03-06": "P3W3",
  "2029-03-13": "P3W4",
  "2029-03-20": "P4W1",
  "2029-03-27": "P4W2",
  "2029-04-03": "P4W3",
  "2029-04-10": "P4W4",
  "2029-04-17": "P5W1",
  "2029-04-24": "P5W2",
  "2029-05-01": "P5W3",
  "2029-05-08": "P5W4",
  "2029-05-15": "P6W1",
  "2029-05-22": "P6W2",
  "2029-05-29": "P6W3",
  "2029-06-05": "P6W4",
  "2029-06-12": "P7W1",
  "2029-06-19": "P7W2",
  "2029-06-26": "P7W3",
  "2029-07-03": "P7W4",
  "2029-07-10": "P8W1",
  "2029-07-17": "P8W2",
  "2029-07-24": "P8W3",
  "2029-07-31": "P8W4",
  "2029-08-07": "P9W1",
  "2029-08-14": "P9W2",
  "2029-08-21": "P9W3",
  "2029-08-28": "P9W4",
  "2029-09-04": "P10W1",
  "2029-09-11": "P10W2",
  "2029-09-18": "P10W3",
  "2029-09-25": "P10W4",
  "2029-10-02": "P11W1",
  "2029-10-09": "P11W2",
  "2029-10-16": "P11W3",
  "2029-10-23": "P11W4",
  "2029-10-30": "P12W1",
  "2029-11-06": "P12W2",
  "2029-11-13": "P12W3",
  "2029-11-20": "P12W4",
  "2029-11-27": "P13W1",
  "2029-12-04": "P13W2",
  "2029-12-11": "P13W3",
  "2029-12-18": "P13W4",
  "2029-12-25": "P1W1",
  "2030-01-01": "P1W2",
  "2030-01-08": "P1W3",
  "2030-01-15": "P1W4",
  "2030-01-22": "P2W1",
  "2030-01-29": "P2W2",
  "2030-02-05": "P2W3",
  "2030-02-12": "P2W4",
  "2030-02-19": "P3W1",
  "2030-02-26": "P3W2",
  "2030-03-05": "P3W3",
  "2030-03-12": "P3W4",
  "2030-03-19": "P4W1",
  "2030-03-26": "P4W2",
  "2030-04-02": "P4W3",
  "2030-04-09": "P4W4",
  "2030-04-16": "P5W1",
  "2030-04-23": "P5W2",
  "2030-04-30": "P5W3",
  "2030-05-07": "P5W4",
  "2030-05-14": "P6W1",
  "2030-05-21": "P6W2",
  "2030-05-28": "P6W3",
  "2030-06-04": "P6W4",
  "2030-06-11": "P7W1",
  "2030-06-18": "P7W2",
  "2030-06-25": "P7W3",
  "2030-07-02": "P7W4",
  "2030-07-09": "P8W1",
  "2030-07-16": "P8W2",
  "2030-07-23": "P8W3",
  "2030-07-30": "P8W4",
  "2030-08-06": "P9W1",
  "2030-08-13": "P9W2",
  "2030-08-20": "P9W3",
  "2030-08-27": "P9W4",
  "2030-09-03": "P10W1",
  "2030-09-10": "P10W2",
  "2030-09-17": "P10W3",
  "2030-09-24": "P10W4",
  "2030-10-01": "P11W1",
  "2030-10-08": "P11W2",
  "2030-10-15": "P11W3",
  "2030-10-22": "P11W4",
  "2030-10-29": "P12W1",
  "2030-11-05": "P12W2",
  "2030-11-12": "P12W3",
  "2030-11-19": "P12W4",
  "2030-11-26": "P13W1",
  "2030-12-03": "P13W2",
  "2030-12-10": "P13W3",
  "2030-12-17": "P13W4",
};

function getPeriodWeek(dateStr) {
  // Use official Ayvaz/PH fiscal calendar lookup (from U.S. Period Calendars 2026-2030)
  const weekKey = getWeekKey(dateStr);
  if (FISCAL_CALENDAR[weekKey]) return FISCAL_CALENDAR[weekKey];
  // Fallback: find nearest known week
  const keys = Object.keys(FISCAL_CALENDAR).sort();
  const target = new Date(weekKey + 'T12:00:00Z').getTime();
  let closest = null, minDiff = Infinity;
  for (const k of keys) {
    const diff = Math.abs(new Date(k + 'T12:00:00Z').getTime() - target);
    if (diff < minDiff) { minDiff = diff; closest = k; }
  }
  if (closest) return FISCAL_CALENDAR[closest];
  return 'P?W?';
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
    
    // ONLY extract Make time and %<4 from SOS Excel
    // All other fields (IST, In-Store, deliveries) come from Above Store PDF
    const make = row[11] || null;
    const pctLt4 = row[13] || null;
    
    // Skip if no make time data
    if (!make) continue;
    
    stores.push({
      store_id: row[0].trim(),
      make: make,
      pct_lt4: pctLt4
      // NOTE: In-Store Time comes ONLY from Above Store PDF, NOT from SOS Excel
    });
  }

  console.log(`SOS Excel parsed: ${stores.length} stores, date=${reportDate}`);
  if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} Make=${stores[0].make} %<4=${stores[0].pct_lt4}`);
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
      // Use ist_avg, NOT in_store — in_store is reserved for Above Store PDF only
      ist_avg: inStore,
      deliveries: totalOrders,
      ist_lt19_pct: ist_lt19_pct_str,
      ist_gt25_count: istGt25,
      // Use standard bucket field names matching the rest of the codebase
      ist_lt10: istLt10,
      ist_1014: ist10to14,
      ist_1518: ist15to18,
      ist_1925: ist19to25,
      ist_gt25: istGt25
    });
  }

  console.log(`IST Tracker Excel parsed: ${stores.length} stores, date=${reportDate}`);
  if (stores.length > 0) console.log(`  Sample: ${stores[0].store_id} IST_avg=${stores[0].ist_avg} Lt19%=${stores[0].ist_lt19_pct}`);
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
        let yr = parseInt(parts[2]) < 100 ? 2000 + parseInt(parts[2]) : parseInt(parts[2]);
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

      // Extract "In-Store Time" directly from "Averages:" section - no math, just read the value
      // The PDF format has labels and values separated, with "In-Store Time" label followed by its value "XX mins"
      let ist_avg = null;
      // Match "In-Store Time" (label), then find the corresponding value line that ends with "mins"
      // The value appears after a blank line and is in the format "XX mins"
      const inStoreValueMatch = block.match(/In-Store Time[\s\S]*?\n\s*(\d+)\s+mins/i);
      if (inStoreValueMatch) {
        ist_avg = parseFloat(inStoreValueMatch[1]);
      } else {
        // Fallback: try same-line format
        const sameLineMatch = block.match(/In-Store Time\s*:\s*(\d+(?:\.\d+)?)/i);
        if (sameLineMatch) {
          ist_avg = parseFloat(sameLineMatch[1]);
        }
      }
      // NO fallback calculation - if not found, leave as null (PDF should always have it)

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

    // Client sends lightweight week keys so we know what weeks the client has
    // (used for logging/diagnostics only - server manages its own data)
    if (req.body.existingWeekKeys) {
      try {
        const clientWeekKeys = JSON.parse(req.body.existingWeekKeys);
        console.log(`Client has ${clientWeekKeys.length} existing weeks: ${clientWeekKeys.map(w=>w.week).join(', ')}`);
      } catch(e) {
        console.warn('Could not parse existingWeekKeys:', e.message);
      }
    }
    const results = [], errors = [];

    for (const file of req.files) {
      const isExcel = file.originalname.match(/\.xlsx?$/i);
      const isPdf = file.originalname.match(/\.pdf$/i);

      // Identify file types
      const isAboveStore = file.originalname.match(/above.?store/i);
      const isSOSExcel = file.originalname.match(/speed.?of.?service|PH_Speed|PH_Delivery/i);
      const isISTTracker = file.originalname.match(/Velocity_IST|IST_Tracker/i);

      if (isPdf && !isAboveStore) {
        errors.push(`${file.originalname}: Unrecognized PDF. Only upload the Above Store Report PDF.`);
        continue;
      }
      if (isExcel && !isSOSExcel && !isISTTracker) {
        errors.push(`${file.originalname}: Unrecognized Excel. Upload SOS Excel or IST Tracker.`);
        continue;
      }
      let parsed = null;
      try {
        if (isExcel && isISTTracker) {
          parsed = parseISTTrackerExcel(file.path);
        } else if (isExcel) {
          parsed = parseSOSExcel(file.path);
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
              area: align.area,
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
            // Update area and coach info if available from alignment
            if (align) {
              existing[s.store_id].area = align.area;
              existing[s.store_id].area_coach = align.area_coach;
              existing[s.store_id].region_coach = align.region_coach;
            }
          } else if (align) {
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name, area: align.area, area_coach: align.area_coach,
              region_coach: align.region_coach,
              ist_lt10: s.ist_lt10, ist_1014: s.ist_1014, ist_1518: s.ist_1518,
              ist_1925: s.ist_1925, ist_gt25: s.ist_gt25,
              ist_gt25_count: s.ist_gt25, ist_lt19_pct: s.ist_lt19_pct,
              in_store: s.ist_avg || null, make: null, production: null, pct_lt4: null,
              pct_lt15: null, on_time: null, rack: null, deliveries: 0
            };
          }
        } else if (parsed.source === 'sos_excel') {
          // SOS Excel: ONLY merge make time and %<4 - nothing else
          // In-Store Time comes ONLY from Above Store PDF
          if (existing[s.store_id]) {
            // Only update make and pct_lt4 fields
            existing[s.store_id].make = s.make;
            existing[s.store_id].pct_lt4 = s.pct_lt4;
            // DO NOT overwrite any other fields (IST, deliveries, on_time, etc.)
          } else if (align) {
            // Create new store record with only make and pct_lt4
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name,
              area: align.area,
              area_coach: align.area_coach,
              region_coach: align.region_coach,
              make: s.make,
              pct_lt4: s.pct_lt4,
              // All other fields null until provided by other sources
              in_store: null, ist_avg: null,
              ist_lt10: null, ist_1014: null, ist_1518: null,
              ist_1925: null, ist_gt25: null,
              ist_gt25_count: null, ist_lt19_pct: null,
              deliveries: null, on_time: null, production: null, pct_lt15: null, rack: null
            };
          }
        } else if (parsed.source === 'ist_tracker') {
          // IST Tracker Excel: ONLY update IST bucket counts + ist_lt19_pct
          // NEVER set in_store or ist_avg - those come from Above Store PDF only
          if (existing[s.store_id]) {
            // Only update bucket counts if we have them
            if (s.ist_lt10 != null) existing[s.store_id].ist_lt10 = s.ist_lt10;
            if (s.ist_1014 != null) existing[s.store_id].ist_1014 = s.ist_1014;
            if (s.ist_1518 != null) existing[s.store_id].ist_1518 = s.ist_1518;
            if (s.ist_1925 != null) existing[s.store_id].ist_1925 = s.ist_1925;
            if (s.ist_gt25 != null) existing[s.store_id].ist_gt25 = s.ist_gt25;
            if (s.ist_gt25_count != null) existing[s.store_id].ist_gt25_count = s.ist_gt25_count;
            if (s.ist_lt19_pct != null) existing[s.store_id].ist_lt19_pct = s.ist_lt19_pct;
            if (s.deliveries) existing[s.store_id].deliveries = s.deliveries;
            // Update area and coach info if available from alignment
            if (align) {
              existing[s.store_id].area = align.area;
              existing[s.store_id].area_coach = align.area_coach;
              existing[s.store_id].region_coach = align.region_coach;
            }
            // DO NOT set in_store or ist_avg - PDF is the source of truth for those
          } else if (align) {
            // No prior record: create stub. ist_avg only used as display fallback until PDF uploaded
            existing[s.store_id] = {
              store_id: s.store_id,
              name: align.name, area: align.area, area_coach: align.area_coach,
              region_coach: align.region_coach,
              ist_lt10: s.ist_lt10||0, ist_1014: s.ist_1014||0, ist_1518: s.ist_1518||0,
              ist_1925: s.ist_1925||0, ist_gt25: s.ist_gt25||0,
              ist_gt25_count: s.ist_gt25_count||0, ist_lt19_pct: s.ist_lt19_pct||null,
              deliveries: s.deliveries||0,
              // ist_avg as temporary fallback only — will be overwritten when PDF is uploaded
              ist_avg: s.ist_avg || null,
              in_store: null, make: null, production: null, pct_lt4: null,
              pct_lt15: null, on_time: null, rack: null
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

// Get available periods from the data (for PTD dropdown)
app.get('/api/periods', (req, res) => {
  try {
    const allData = loadData();
    const weeks = allData.weeks || {};
    
    // Find all unique periods in the data
    const periodMap = {};
    Object.values(weeks).forEach(w => {
      if (w.period) {
        // Extract period (e.g., 'P4W2' -> 'P4')
        const periodKey = w.period.replace(/W\d+$/, '');
        if (!periodMap[periodKey]) {
          periodMap[periodKey] = {
            period: periodKey,
            name: FISCAL_CALENDAR_2026[periodKey]?.name || periodKey,
            weeks: [],
            weeksWithData: []
          };
        }
        periodMap[periodKey].weeks.push(w.period);
        periodMap[periodKey].weeksWithData.push({
          periodWeek: w.period,
          week: w.week,
          days: Object.keys(w.days || {}).length
        });
      }
    });
    
    // Only return periods with multiple weeks of data
    const periods = Object.values(periodMap).filter(p => p.weeksWithData.length > 0);
    
    res.json({ periods });
  } catch(e) { 
    res.status(500).json({ error: e.message }); 
  }
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

  // Check if we have pre-calculated WTD stores data from Excel import
  if (wtd.stores && wtd.stores.length > 0 && wtd.stores[0].wtd_in_store) {
    // Use the pre-calculated WTD data directly
    const storeWTD = wtd.stores.map(s => {
      const align = ALIGNMENT[s.store_id] || {};
      return {
        ...s,
        name: s.name || align.name || 'Unknown',
        area: s.area || align.area || '',
        area_coach: s.area_coach || align.area_coach || '',
        region_coach: s.region_coach || align.region_coach || ''
      };
    });
    
    return {
      week: wtd.week,
      period: wtd.period || '',
      dateRange: getWeekDateRange(wtd.week),
      days: days.map(d => ({ date: d.date, type: d.type, storeCount: (d.stores||[]).length, uploader: d.uploader || 'Matt Hester' })),
      stores: storeWTD,
      weekStart: days.length ? (d => (d.getUTCMonth()+1)+'/'+(d.getUTCDate()))(new Date(days[0].date+'T12:00:00Z')) : '',
      weekEnd: days.length ? (d => (d.getUTCMonth()+1)+'/'+(d.getUTCDate()))(new Date(days[days.length-1].date+'T12:00:00Z')) : ''
    };
  }

  // Fall back to calculation from daily data
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
      store_id: sid, name: align.name, area: align.area, area_coach: align.area_coach, region_coach: align.region_coach,
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
    days: days.map(d => ({ date: d.date, type: d.type, storeCount: d.stores.length, uploader: d.uploader || 'Matt Hester' })),
    stores: storeWTD,
    weekStart: days.length ? (d => (d.getUTCMonth()+1)+'/'+(d.getUTCDate()))(new Date(days[0].date+'T12:00:00Z')) : '',
    weekEnd: days.length ? (d => (d.getUTCMonth()+1)+'/'+(d.getUTCDate()))(new Date(days[days.length-1].date+'T12:00:00Z')) : ''
  };
}

function computeAllWeeks(allData) {
  const weeks = allData.weeks || {};
  const weekKeys = Object.keys(weeks).sort((a, b) => b.localeCompare(a)); // newest first
  return weekKeys.map(k => computeWeek(weeks[k]));
}

const ALIGNMENT = {"S038876":{"name":"Senoia","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039377":{"name":"Griffin","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039378":{"name":"Union City","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039379":{"name":"Jefferson St","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039384":{"name":"Newnan","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039454":{"name":"Zebulon","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039465":{"name":"Senoia Rd","area":"Area 2011","area_coach":"Darian Spikes","region_coach":"Harold Lacoste"},"S039383":{"name":"Stockbridge","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039388":{"name":"Jonesboro Rd","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039393":{"name":"Lovejoy","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039429":{"name":"Ola","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039461":{"name":"County Line","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039513":{"name":"Jodeco","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039521":{"name":"Kellytown","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039522":{"name":"Ellenwood","area":"Area 2016","area_coach":"Ebony Simmons","region_coach":"Harold Lacoste"},"S039375":{"name":"Bells Ferry Rd","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039376":{"name":"CrossRds","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039382":{"name":"Glade Rd","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039387":{"name":"Kennesaw","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039392":{"name":"Towne Lake","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039462":{"name":"Acworth/Emerson","area":"Area 2022","area_coach":"Ja'Don McNeil","region_coach":"Harold Lacoste"},"S039380":{"name":"Windy Hill","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039386":{"name":"Powder Springs","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039389":{"name":"Lithia Springs","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039410":{"name":"Mableton","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039451":{"name":"Bolton","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039525":{"name":"Smyrna","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039527":{"name":"Austell Rd","area":"Area 2000","area_coach":"Jorge Garcia","region_coach":"Harold Lacoste"},"S039412":{"name":"Miracle Strip","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039413":{"name":"Navarre","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039414":{"name":"Gulf Breeze","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039415":{"name":"Miramar Beach","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039416":{"name":"Niceville","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039430":{"name":"Racetrack","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039529":{"name":"Crestview","area":"Area 2015","area_coach":"Marc Gannon","region_coach":"Harold Lacoste"},"S039381":{"name":"Fairburn Rd","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039385":{"name":"Ridge Rd","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039390":{"name":"East Paulding","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039391":{"name":"Hwy 5","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039526":{"name":"Dallas","area":"Area 2034","area_coach":"Michelle Meehan","region_coach":"Harold Lacoste"},"S039417":{"name":"Collinsville","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039419":{"name":"Martinsville","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039421":{"name":"College Rd","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039424":{"name":"Gate City Blvd","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039427":{"name":"Pyramid Village","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039436":{"name":"Battleground","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039457":{"name":"E. Greensboro","area":"Area 2041","area_coach":"ARNWINE-OPEN","region_coach":"Preston Arnwine"},"S039418":{"name":"Riverside Dr","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039422":{"name":"South Church","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039423":{"name":"Graham","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039432":{"name":"Mebane","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039433":{"name":"Elton Way","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039455":{"name":"Spring Garden","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039456":{"name":"Whitsett","area":"Area 2017","area_coach":"Emmanuel Boateng","region_coach":"Preston Arnwine"},"S039420":{"name":"Harrisonburg","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039425":{"name":"Elkton","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039426":{"name":"Woodstock","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039428":{"name":"Stuarts Draft","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039431":{"name":"Staunton","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039435":{"name":"Shoppers World","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039450":{"name":"Orange","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039453":{"name":"JMU/Market","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039466":{"name":"Waynesboro","area":"Area 2004","area_coach":"Erin Pizzo","region_coach":"Preston Arnwine"},"S039400":{"name":"E Palmetto","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039401":{"name":"Darlington","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039402":{"name":"2nd Loop","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039403":{"name":"Marion","area":"Area 2009","area_coach":"Royal Mitchell","region_coach":"Preston Arnwine"},"S039394":{"name":"Elberton","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039395":{"name":"Abbeville","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039396":{"name":"Hartwell","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039398":{"name":"Royston","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039399":{"name":"Lavonia","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039404":{"name":"Greenwood Bypass","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039405":{"name":"Simpsonville","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039407":{"name":"Newberry","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S039408":{"name":"Seneca","area":"Area 2048","area_coach":"Russell Kowalczyk","region_coach":"Preston Arnwine"},"S040090":{"name":"Main","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040091":{"name":"Silver City","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040093":{"name":"Missouri","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S040102":{"name":"Deming","area":"Area 2002","area_coach":"Brenda Marta","region_coach":"Terrance Spillane"},"S039180":{"name":"Zaragosa","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039182":{"name":"Vista","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039185":{"name":"Gateway","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039318":{"name":"Socorro","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S039323":{"name":"Tierre Este","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S041651":{"name":"Eastlake","area":"Area 2010","area_coach":"Constance Miranda","region_coach":"Terrance Spillane"},"S040082":{"name":"Taylor Ranch","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040084":{"name":"7th/Lomas","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040101":{"name":"Washington/Zuni","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040107":{"name":"Coors/Barcelona","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040108":{"name":"Wyoming/Harper","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S040111":{"name":"303 Coors","area":"Area 2033","area_coach":"Eric Harstine","region_coach":"Terrance Spillane"},"S038729":{"name":"Kenworthy","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039174":{"name":"University","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039175":{"name":"Airway","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039178":{"name":"CrossRds EP","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039192":{"name":"Resler","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039324":{"name":"Outlet Mall","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S039448":{"name":"Dyer","area":"Area 2024","area_coach":"Javier Martinez","region_coach":"Terrance Spillane"},"S040088":{"name":"Los Lunas","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040096":{"name":"Belen","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040099":{"name":"Candelaria","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040100":{"name":"T or C","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S040110":{"name":"Bull Chicks","area":"Area 2055","area_coach":"Kevin Dunn","region_coach":"Terrance Spillane"},"S039589":{"name":"Rio Rancho","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040094":{"name":"Villa Linda Mall","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040104":{"name":"Southern","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040105":{"name":"Las Vegas","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040106":{"name":"Espanola","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S040109":{"name":"Unser & McMahon","area":"Area 2039","area_coach":"Max Losey","region_coach":"Terrance Spillane"},"S039173":{"name":"Yarbrough","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039176":{"name":"Lovington","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039177":{"name":"Hobbs","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039179":{"name":"George Dieter","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039188":{"name":"Carlsbad","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039518":{"name":"Hobbs North","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S039530":{"name":"Montana","area":"Area 2043","area_coach":"Oscar Gutierrez","region_coach":"Terrance Spillane"},"S040083":{"name":"20th St","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040085":{"name":"North Gallup","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040086":{"name":"Main Street","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040087":{"name":"East Gallup","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040092":{"name":"Aztec","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"},"S040112":{"name":"Durango","area":"Area 2008","area_coach":"Tami Elliott-Baker","region_coach":"Terrance Spillane"}};

// =====================
// AUTOMATION API ENDPOINTS
// For automated report uploads from ODS
// =====================

// Simple auth token for automation (in production, use proper auth)
const AUTOMATION_TOKEN = process.env.AUTOMATION_TOKEN || 'velocity-auto-2024';

// Middleware to verify automation requests
function verifyAutomationAuth(req, res, next) {
  const token = req.headers['x-automation-token'] || req.query.token;
  if (token !== AUTOMATION_TOKEN) {
    return res.status(401).json({ error: 'Unauthorized - Invalid automation token' });
  }
  next();
}

// Upload Speed of Service report (XLSX) from automation
app.post('/api/automation/upload-sos', verifyAutomationAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;

    // Use the existing SOS Excel parser
    const parsed = parseSOSExcel(filePath);
    
    if (!parsed.stores || parsed.stores.length === 0) {
      // Clean up temp file
      try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch(e) {}
      return res.status(400).json({ error: 'No store data found in SOS Excel' });
    }

    // Use report date from file, or default to yesterday
    let dateStr = parsed.reportDate;
    if (!dateStr) {
      const yesterday = new Date();
      yesterday.setDate(yesterday.getDate() - 1);
      dateStr = yesterday.toISOString().split('T')[0];
    }
    
    const weekKey = getWeekKey(dateStr);
    const periodWeek = FISCAL_CALENDAR[weekKey] || '';

    // Load existing data
    const allData = loadData();
    if (!allData.weeks[weekKey]) {
      allData.weeks[weekKey] = { week: weekKey, period: periodWeek, days: {} };
    }
    if (!allData.weeks[weekKey].days[dateStr]) {
      allData.weeks[weekKey].days[dateStr] = { date: dateStr, type: 'automation', stores: [], uploader: 'automation' };
    }

    const existing = {};
    (allData.weeks[weekKey].days[dateStr].stores || []).forEach(s => { existing[s.store_id] = s; });

    let storeCount = 0;
    parsed.stores.forEach(s => {
      const sid = s.store_id;
      const align = ALIGNMENT[sid];
      if (!align && !existing[sid]) return;

      if (existing[sid]) {
        // Only update make and pct_lt4 from SOS Excel
        existing[sid].make = s.make;
        existing[sid].pct_lt4 = s.pct_lt4;
      } else if (align) {
        existing[sid] = {
          store_id: sid,
          name: align.name,
          area: align.area,
          area_coach: align.area_coach,
          region_coach: align.region_coach,
          make: s.make,
          pct_lt4: s.pct_lt4,
          in_store: null,
          ist_lt10: null, ist_1014: null, ist_1518: null,
          ist_1925: null, ist_gt25: null,
          ist_gt25_count: null, ist_lt19_pct: null,
          deliveries: null, on_time: null, production: null, pct_lt15: null, rack: null
        };
      }
      storeCount++;
    });

    allData.weeks[weekKey].days[dateStr].stores = Object.values(existing);
    saveData(allData);

    // Clean up temp file
    try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch(e) {}

    res.json({ 
      success: true, 
      message: 'Speed of Service report uploaded successfully',
      date: dateStr,
      week: weekKey,
      period: periodWeek,
      storeCount 
    });
  } catch (e) {
    console.error('Automation SOS upload error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Upload Daily Dispatch report (PDF) from automation
app.post('/api/automation/upload-dispatch', verifyAutomationAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;

    // Use the existing PDF parser (local, no API needed)
    const parsed = parseAboveStorePDFLocal(filePath);
    
    if (!parsed.stores || parsed.stores.length === 0) {
      return res.status(400).json({ error: 'No store data found in PDF' });
    }

    // Get yesterday's date for the report
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = yesterday.toISOString().split('T')[0];
    const weekKey = getWeekKey(dateStr);
    const periodWeek = FISCAL_CALENDAR[weekKey] || '';

    // Load existing data
    const allData = loadData();
    if (!allData.weeks[weekKey]) {
      allData.weeks[weekKey] = { week: weekKey, period: periodWeek, days: {} };
    }
    if (!allData.weeks[weekKey].days[dateStr]) {
      allData.weeks[weekKey].days[dateStr] = { date: dateStr, type: 'automation', stores: [], uploader: 'automation' };
    }

    const existing = {};
    (allData.weeks[weekKey].days[dateStr].stores || []).forEach(s => { existing[s.store_id] = s; });

    let storeCount = 0;
    parsed.stores.forEach(s => {
      const sid = s.store_id;
      const align = ALIGNMENT[sid];
      if (!align && !existing[sid]) return;

      if (existing[sid]) {
        // Merge IST data from PDF
        existing[sid].ist_lt10 = s.ist_lt10 || existing[sid].ist_lt10;
        existing[sid].ist_1014 = s.ist_1014 || existing[sid].ist_1014;
        existing[sid].ist_1518 = s.ist_1518 || existing[sid].ist_1518;
        existing[sid].ist_1925 = s.ist_1925 || existing[sid].ist_1925;
        existing[sid].ist_gt25 = s.ist_gt25 || existing[sid].ist_gt25;
        existing[sid].ist_gt25_count = s.ist_gt25 || existing[sid].ist_gt25_count;
        existing[sid].ist_lt19_pct = s.ist_lt19_pct || existing[sid].ist_lt19_pct;
        if (s.ist_avg) existing[sid].in_store = s.ist_avg;
      } else if (align) {
        existing[sid] = {
          store_id: sid,
          name: align.name,
          area: align.area,
          area_coach: align.area_coach,
          region_coach: align.region_coach,
          ist_lt10: s.ist_lt10,
          ist_1014: s.ist_1014,
          ist_1518: s.ist_1518,
          ist_1925: s.ist_1925,
          ist_gt25: s.ist_gt25,
          ist_gt25_count: s.ist_gt25,
          ist_lt19_pct: s.ist_lt19_pct,
          in_store: s.ist_avg || null,
          make: null, pct_lt4: null, production: null, pct_lt15: null,
          on_time: null, rack: null, deliveries: null
        };
      }
      storeCount++;
    });

    allData.weeks[weekKey].days[dateStr].stores = Object.values(existing);
    saveData(allData);

    // Clean up temp file
    try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch(e) {}

    res.json({ 
      success: true, 
      message: 'Daily Dispatch report uploaded successfully',
      date: dateStr,
      week: weekKey,
      period: periodWeek,
      storeCount 
    });
  } catch (e) {
    console.error('Automation Dispatch upload error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Send daily report emails
app.post('/api/automation/send-emails', verifyAutomationAuth, async (req, res) => {
  try {
    const nodemailer = require('nodemailer');
    
    // Gmail SMTP configuration
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com',
        pass: process.env.VELOCITY_EMAIL_PASS || 'dewnkjrxbgmodwwd'
      }
    });

    // Get yesterday's data
    const allData = loadData();
    const weekKeys = Object.keys(allData.weeks || {}).sort((a, b) => b.localeCompare(a));
    
    if (weekKeys.length === 0) {
      return res.status(400).json({ error: 'No data available' });
    }

    const latestWeek = allData.weeks[weekKeys[0]];
    const dayKeys = Object.keys(latestWeek.days || {}).sort((a, b) => b.localeCompare(a));
    
    if (dayKeys.length === 0) {
      return res.status(400).json({ error: 'No daily data available' });
    }

    // Use computed week data which has WTD values
    const weekComputed = computeWeek(latestWeek);
    const stores = weekComputed.stores || [];

    // Calculate summary metrics from WTD values
    const validStores = stores.filter(s => s.wtd_in_store || s.ist_avg || s.in_store || s.make);
    const avgInStore = validStores.reduce((a, s) => a + (s.wtd_in_store || s.ist_avg || 0), 0) / validStores.length;
    const avgPctLt4 = validStores.reduce((a, s) => {
      const pct = parseFloat(String(s.wtd_pct_lt4 || '0').replace('%', '')) || 0;
      return a + pct;
    }, 0) / validStores.length;
    const avgOnTime = validStores.reduce((a, s) => {
      const pct = parseFloat(String(s.wtd_on_time || '0').replace('%', '')) || 0;
      return a + pct;
    }, 0) / validStores.length;
    const totalDeliveries = validStores.reduce((a, s) => a + (s.wtd_deliveries || 0), 0);

    // Sort by WTD in-store time for top/bottom performers
    // Use ist_avg as fallback if wtd_in_store is not available
    const sortedByIST = [...validStores].sort((a, b) => {
      const aIST = a.wtd_in_store || a.ist_avg || 999;
      const bIST = b.wtd_in_store || b.ist_avg || 999;
      return aIST - bIST;
    });
    const topPerformers = sortedByIST.slice(0, 5);
    const bottomPerformers = sortedByIST.slice(-5).reverse();

    const dashboardUrl = 'https://00p2f.app.super.myninja.ai';

    // Email HTML
    const generateHTML = (isAreaCoach = false, areaFilter = null) => {
      let filteredStores = validStores;
      let filteredTop = topPerformers;
      let filteredBottom = bottomPerformers;
      
      if (areaFilter) {
        filteredStores = validStores.filter(s => s.area === areaFilter);
        const areaSorted = [...filteredStores].sort((a, b) => (a.wtd_in_store || a.ist_avg || a.in_store || 999) - (b.wtd_in_store || b.ist_avg || b.in_store || 999));
        filteredTop = areaSorted.slice(0, 5);
        filteredBottom = areaSorted.slice(-5).reverse();
      }

      const getISTColor = (ist) => {
        if (!ist) return '#666';
        if (ist <= 19) return '#28a745';
        if (ist <= 22) return '#ffc107';
        if (ist <= 25) return '#fd7e14';
        return '#dc3545';
      };

      return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
    h1 { color: #e31837; border-bottom: 3px solid #e31837; padding-bottom: 10px; }
    h2 { color: #333; margin-top: 30px; }
    .summary-box { background: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .summary-box table { width: 100%; border-collapse: collapse; }
    .summary-box td { padding: 8px; font-size: 16px; }
    .summary-box td:last-child { font-weight: bold; text-align: right; }
    .dashboard-btn { 
      display: inline-block; background: #e31837; color: white; padding: 12px 24px; 
      text-decoration: none; border-radius: 5px; margin: 20px 0; font-weight: bold;
    }
    table.stores { width: 100%; border-collapse: collapse; margin: 10px 0; }
    table.stores th { background: #333; color: white; padding: 10px; text-align: left; }
    table.stores td { padding: 8px; border-bottom: 1px solid #ddd; }
    .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 12px; }
  </style>
</head>
<body>
  <h1>🍕 Velocity Daily Report</h1>
  <p><strong>Report Date:</strong> ${dayKeys[0]}</p>
  ${areaFilter ? `<p><strong>Area:</strong> ${areaFilter}</p>` : ''}
  
  <div class="summary-box">
    <h2 style="margin-top:0">📊 Yesterday's Summary</h2>
    <table>
      <tr><td>Stores Reporting:</td><td>${filteredStores.length}</td></tr>
      <tr><td>Avg In-Store Time:</td><td style="color: ${getISTColor(avgInStore)}">${avgInStore.toFixed(1)} mins</td></tr>
      <tr><td>Avg % <4 Min:</td><td>${avgPctLt4.toFixed(1)}%</td></tr>
      <tr><td>Avg On-Time %:</td><td>${avgOnTime.toFixed(1)}%</td></tr>
      <tr><td>Total Deliveries:</td><td>${totalDeliveries.toLocaleString()}</td></tr>
    </table>
  </div>

  <h2>🏆 Top 5 Performers</h2>
  <table class="stores">
    <tr><th>Store</th><th>In-Store</th><th>Make</th><th>%<4</th></tr>
    ${filteredTop.map(s => `
      <tr>
        <td><strong>${s.name}</strong><br><small>${s.store_id}</small></td>
        <td style="color: ${getISTColor(s.wtd_in_store || s.ist_avg || s.in_store)}">${s.wtd_in_store || s.ist_avg || s.in_store || '—'} mins</td>
        <td>${s.wtd_make || s.make || '—'}</td>
        <td>${s.wtd_pct_lt4 || s.pct_lt4 || '—'}</td>
      </tr>
    `).join('')}
  </table>

  <h2>⚠️ Bottom 5 Performers</h2>
  <table class="stores">
    <tr><th>Store</th><th>In-Store</th><th>Make</th><th>%<4</th></tr>
    ${filteredBottom.map(s => `
      <tr>
        <td><strong>${s.name}</strong><br><small>${s.store_id}</small></td>
        <td style="color: ${getISTColor(s.wtd_in_store || s.ist_avg || s.in_store)}">${s.wtd_in_store || s.ist_avg || s.in_store || '—'} mins</td>
        <td>${s.wtd_make || s.make || '—'}</td>
        <td>${s.wtd_pct_lt4 || s.pct_lt4 || '—'}</td>
      </tr>
    `).join('')}
  </table>

  <div class="footer">
    <p>This is an automated email from Velocity - Pizza Hut Speed of Service Dashboard</p>
    <p>Generated: ${new Date().toLocaleString()}</p>
  </div>
</body>
</html>`;
    };

    // Email distribution lists
    const areaCoaches = [
      { name: 'Jorge Garcia', area: '2000', email: 'jgarcia@ayvazpizza.com' },
      { name: 'Darian Spikes', area: '2011', email: 'dspikes@ayvazpizza.com' },
      { name: 'Marc Gannon', area: '2015', email: 'mgannon@ayvazpizza.com' },
      { name: 'Ebony Simmons', area: '2016', email: 'esimmons@ayvazpizza.com' },
      { name: 'Jadon McNeil', area: '2022', email: 'jmcneil@ayvazpizza.com' },
      { name: 'Michelle Meehan', area: '2034', email: 'mmeehan@ayvazpizza.com' }
    ];
    const peers = [
      { name: 'Preston Arnwine', email: 'parnwine@ayvazpizza.com' },
      { name: 'Terrance Spillane', email: 'tspillane@ayvazpizza.com' }
    ];
    const vp = { name: 'Matt Hester', email: 'mhester@ayvazpizza.com' };
    // Region coach (Harold Lacoste) gets full summary
    const regionCoach = { name: 'Harold Lacoste', email: 'hlacoste@ayvazpizza.com' };

    const results = { sent: [], failed: [] };

    // Generate Excel attachment for all emails
    const excelBuffer = generateExcelExport(weekComputed, allData);
    const excelAttachment = {
      filename: `Velocity_Report_${dayKeys[0]}.xlsx`,
      content: excelBuffer,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };

    // TESTING MODE: Only send to Harold Lacoste
    // Uncomment the sections below when ready to go live
    
    /* DISABLED FOR TESTING
    // Send to Area Coaches (each gets their area's data)
    for (const coach of areaCoaches) {
      try {
        const info = await transporter.sendMail({
          from: `"Velocity Reports" <${process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com'}>`,
          to: coach.email,
          cc: vp.email,
          subject: `Velocity Daily Report - ${coach.area} - ${dayKeys[0]}`,
          html: generateHTML(true, coach.area),
          attachments: [excelAttachment]
        });
        results.sent.push({ to: coach.email, area: coach.area, messageId: info.messageId });
      } catch (e) {
        results.failed.push({ to: coach.email, error: e.message });
      }
    }

    // Send summary to peers
    for (const peer of peers) {
      try {
        const info = await transporter.sendMail({
          from: `"Velocity Reports" <${process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com'}>`,
          to: peer.email,
          subject: `Velocity Daily Report - Summary - ${dayKeys[0]}`,
          html: generateHTML(false),
          attachments: [excelAttachment]
        });
        results.sent.push({ to: peer.email, messageId: info.messageId });
      } catch (e) {
        results.failed.push({ to: peer.email, error: e.message });
      }
    }
    */

    // Send summary to Region Coach (Harold Lacoste) - ONLY THIS IS ACTIVE FOR TESTING
    try {
      const info = await transporter.sendMail({
        from: `"Velocity Reports" <${process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com'}>`,
        to: regionCoach.email,
        subject: `Velocity Daily Report - Region Summary - ${dayKeys[0]}`,
        html: generateHTML(false),
        attachments: [excelAttachment]
      });
      results.sent.push({ to: regionCoach.email, messageId: info.messageId });
    } catch (e) {
      results.failed.push({ to: regionCoach.email, error: e.message });
    }

    res.json({ 
      success: true, 
      date: dayKeys[0],
      totalSent: results.sent.length,
      totalFailed: results.failed.length,
      results 
    });
  } catch (e) {
    console.error('Send emails error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Generate Excel export for email attachment
// Generate Excel export for email attachment
function generateExcelExport(weekData, allData = null) {
  const wb = XLSX.utils.book_new();
  const stores = weekData.stores || [];
  const period = weekData.period || '';
  const weekStart = weekData.weekStart || '';
  const weekEnd = weekData.weekEnd || '';

  // weekData.days is a summary array; get raw day data from allData for per-day store access
  const rawWeekDays = (allData && allData.weeks && weekData.week && allData.weeks[weekData.week])
    ? (allData.weeks[weekData.week].days || {}) : {};
  const dayData = (typeof rawWeekDays === 'object' && !Array.isArray(rawWeekDays)) ? rawWeekDays : {};
  const dayKeys = Object.keys(dayData).sort();

  const round1 = v => v != null && !isNaN(v) ? Math.round(v * 10) / 10 : null;

  function calcTotals(storeList) {
    const totalOrders = storeList.reduce((s, x) => s + (x.total_orders || x.wtd_deliveries || 0), 0);
    const weightedIST = storeList.reduce((s, x) => {
      const orders = x.total_orders || x.wtd_deliveries || 0;
      const ist = x.wtd_in_store || x.ist_avg || 0;
      return s + ist * orders;
    }, 0);
    const avgIST = totalOrders > 0 ? round1(weightedIST / totalOrders) : null;
    const lt10  = storeList.reduce((s, x) => s + (x.ist_lt10  || 0), 0);
    const i1014 = storeList.reduce((s, x) => s + (x.ist_1014  || 0), 0);
    const i1518 = storeList.reduce((s, x) => s + (x.ist_1518  || 0), 0);
    const i1925 = storeList.reduce((s, x) => s + (x.ist_1925  || 0), 0);
    const gt25  = storeList.reduce((s, x) => s + (x.ist_gt25  || 0), 0);
    const lt19  = lt10 + i1014 + i1518;
    const lt19Pct = totalOrders > 0 ? lt19 / totalOrders : null;
    return { totalOrders, avgIST, lt10, i1014, i1518, i1925, gt25, lt19Pct };
  }

  function pct(n, d) { return d > 0 ? n / d : null; }

  const HEADERS = ['Level','Region','Area Coach','Store #','Store Name','Avg IST (mins)',
    'Total Orders','IST <10 #','IST <10 %','IST 10-14 #','IST 10-14 %',
    'IST 15-18 #','IST 15-18 %','IST 19-25 #','IST 19-25 %',
    'IST >25 #','IST >25 %','IST <19 %'];

  function groupStores(storeList) {
    const byRegion = {};
    storeList.forEach(s => {
      const r = s.region_coach || 'Unknown';
      if (!byRegion[r]) byRegion[r] = { stores: [], byArea: {} };
      byRegion[r].stores.push(s);
      const a = s.area || 'Unknown';
      if (!byRegion[r].byArea[a]) byRegion[r].byArea[a] = { stores: [], coach: s.area_coach || '' };
      byRegion[r].byArea[a].stores.push(s);
    });
    return byRegion;
  }

  function buildRows(title, storeList, showAvgIST) {
    if (showAvgIST === undefined) showAvgIST = true;
    const rows = [[title], HEADERS];
    const t = calcTotals(storeList);
    rows.push(['TOTAL','ALL REGIONS','','',storeList.length+' Stores',
      showAvgIST ? t.avgIST : null, t.totalOrders,
      t.lt10, pct(t.lt10,t.totalOrders), t.i1014, pct(t.i1014,t.totalOrders),
      t.i1518, pct(t.i1518,t.totalOrders), t.i1925, pct(t.i1925,t.totalOrders),
      t.gt25, pct(t.gt25,t.totalOrders), t.lt19Pct]);
    const byRegion = groupStores(storeList);
    for (const region of Object.keys(byRegion)) {
      const rdata = byRegion[region];
      const rt = calcTotals(rdata.stores);
      rows.push(['REGION', region,'','',rdata.stores.length+' Stores',
        showAvgIST ? rt.avgIST : null, rt.totalOrders,
        rt.lt10, pct(rt.lt10,rt.totalOrders), rt.i1014, pct(rt.i1014,rt.totalOrders),
        rt.i1518, pct(rt.i1518,rt.totalOrders), rt.i1925, pct(rt.i1925,rt.totalOrders),
        rt.gt25, pct(rt.gt25,rt.totalOrders), rt.lt19Pct]);
      for (const areaKey of Object.keys(rdata.byArea)) {
        const adata = rdata.byArea[areaKey];
        const at = calcTotals(adata.stores);
        rows.push(['AREA', region, adata.coach, areaKey, adata.stores.length+' Stores',
          showAvgIST ? at.avgIST : null, at.totalOrders,
          at.lt10, pct(at.lt10,at.totalOrders), at.i1014, pct(at.i1014,at.totalOrders),
          at.i1518, pct(at.i1518,at.totalOrders), at.i1925, pct(at.i1925,at.totalOrders),
          at.gt25, pct(at.gt25,at.totalOrders), at.lt19Pct]);
        adata.stores.forEach(s => {
          const orders = s.total_orders || s.wtd_deliveries || 0;
          const ist = showAvgIST ? round1(s.wtd_in_store || s.ist_avg || null) : null;
          rows.push(['STORE', s.region_coach||'', s.area_coach||'', s.store_id||'', s.name||'',
            ist, orders,
            s.ist_lt10||0, s.ist_lt10_pct||pct(s.ist_lt10,orders),
            s.ist_1014||0, s.ist_1014_pct||pct(s.ist_1014,orders),
            s.ist_1518||0, s.ist_1518_pct||pct(s.ist_1518,orders),
            s.ist_1925||0, s.ist_1925_pct||pct(s.ist_1925,orders),
            s.ist_gt25||0, s.ist_gt25_pct||pct(s.ist_gt25,orders),
            s.ist_lt19_pct||null]);
        });
      }
    }
    return rows;
  }

  const DAY_SHORT = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const MON_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  function dayLabel(dateStr) {
    const [y,m,d] = dateStr.split('-').map(Number);
    const dt = new Date(y, m-1, d);
    return DAY_SHORT[dt.getDay()]+', '+MON_SHORT[m-1]+' '+d;
  }
  function dayLabelShort(dateStr) {
    const [y,m,d] = dateStr.split('-').map(Number);
    const dt = new Date(y, m-1, d);
    return DAY_SHORT[dt.getDay()]+', '+m+'/'+d;
  }

  function fmtDelta(prev, curr, isStore) {
    if (prev == null || curr == null || isNaN(prev) || isNaN(curr)) return '\u2013';
    const diff = round1(curr - prev);
    if (diff === null) return '\u2013';
    if (isStore) {
      const d = Math.round(diff);
      if (d > 0) return '\u25b2 +'+d;
      if (d < 0) return '\u25bc '+d;
      return '\u2013';
    } else {
      if (diff > 0) return '\u25b2 +'+diff;
      if (diff < 0) return '\u25bc '+diff;
      return '\u2013';
    }
  }

  function getGroupedIST(storeList) {
    if (!storeList || !storeList.length) return null;
    return calcTotals(storeList.map(s => ({
      ...s,
      wtd_in_store: s.ist_avg || s.in_store || s.wtd_in_store
    }))).avgIST;
  }

  // ─── WTD IST ───
  const wtdWs = XLSX.utils.aoa_to_sheet(buildRows('WTD IST \u2014 '+period+' '+weekStart+'-'+weekEnd, stores, true));
  XLSX.utils.book_append_sheet(wb, wtdWs, 'WTD IST');

  // ─── PTD IST ─── aggregate all weeks in same period
  let ptdStoreMap = {};
  const periodBase = period ? period.replace(/W\d+$/, '') : '';
  if (allData && allData.weeks && periodBase) {
    const ptdRawWeeks = Object.values(allData.weeks).filter(w => w.period && w.period.replace(/W\d+$/, '') === periodBase);
    ptdRawWeeks.forEach(rawWk => {
      const wkC = computeWeek(rawWk);
      (wkC.stores || []).forEach(s => {
        const sid = s.store_id;
        if (!ptdStoreMap[sid]) {
          ptdStoreMap[sid] = { store_id: sid, name: s.name, area: s.area, area_coach: s.area_coach,
            region_coach: s.region_coach, ist_lt10: 0, ist_1014: 0, ist_1518: 0,
            ist_1925: 0, ist_gt25: 0, total_orders: 0, wtd_deliveries: 0,
            wtd_in_store: null, ist_avg: null };
        }
        const orders = s.total_orders || s.wtd_deliveries || 0;
        ptdStoreMap[sid].ist_lt10  = (ptdStoreMap[sid].ist_lt10  || 0) + (s.ist_lt10  || 0);
        ptdStoreMap[sid].ist_1014  = (ptdStoreMap[sid].ist_1014  || 0) + (s.ist_1014  || 0);
        ptdStoreMap[sid].ist_1518  = (ptdStoreMap[sid].ist_1518  || 0) + (s.ist_1518  || 0);
        ptdStoreMap[sid].ist_1925  = (ptdStoreMap[sid].ist_1925  || 0) + (s.ist_1925  || 0);
        ptdStoreMap[sid].ist_gt25  = (ptdStoreMap[sid].ist_gt25  || 0) + (s.ist_gt25  || 0);
        ptdStoreMap[sid].total_orders  = (ptdStoreMap[sid].total_orders  || 0) + orders;
        ptdStoreMap[sid].wtd_deliveries = ptdStoreMap[sid].total_orders;
      });
    });
  }
  const ptdStoreList = Object.values(ptdStoreMap).length ? Object.values(ptdStoreMap) : stores;
  const ptdTitle = 'PTD IST \u2014 '+(periodBase||period)+' (Period To Date)';
  const ptdWs = XLSX.utils.aoa_to_sheet(buildRows(ptdTitle, ptdStoreList, false));
  XLSX.utils.book_append_sheet(wb, ptdWs, 'PTD IST');

  // ─── Daily sheets ───
  dayKeys.forEach(dayKey => {
    const dayStores = dayData[dayKey] ? dayData[dayKey].stores || [] : [];
    if (!dayStores.length) return;
    const label = dayLabel(dayKey);
    const ws = XLSX.utils.aoa_to_sheet(buildRows(label+' \u2014 '+period, dayStores, true));
    XLSX.utils.book_append_sheet(wb, ws, label);
  });

  // ─── WTD Trend (day-over-day with delta columns) ───
  const trendDayHeaders = dayKeys.map(dayLabelShort);
  const trendDeltaHeaders = dayKeys.slice(1).map((dk, i) => '\u0394 '+dayLabelShort(dayKeys[i])+'\u2192'+dayLabelShort(dk));
  const trendHeaders = ['Level','Region','Area Coach','Store #','Store Name',...trendDayHeaders,...trendDeltaHeaders];
  const trendRows = [['In-Store Time Trend \u2014 '+period+' '+weekStart+'-'+weekEnd], trendHeaders];

  function addTrendRows(storeList) {
    const byRegion = groupStores(storeList);
    const totalISTs = dayKeys.map(dk => getGroupedIST(dayData[dk] ? dayData[dk].stores : []));
    const totalDeltas = dayKeys.slice(1).map((dk,i) => fmtDelta(totalISTs[i], totalISTs[i+1], false));
    trendRows.push(['TOTAL','ALL REGIONS','','',storeList.length+' Stores',...totalISTs,...totalDeltas]);
    for (const region of Object.keys(byRegion)) {
      const rdata = byRegion[region];
      const rISTs = dayKeys.map(dk => getGroupedIST((dayData[dk] ? dayData[dk].stores||[] : []).filter(s => s.region_coach === region)));
      const rDeltas = dayKeys.slice(1).map((dk,i) => fmtDelta(rISTs[i], rISTs[i+1], false));
      trendRows.push(['REGION', region,'','',rdata.stores.length+' Stores',...rISTs,...rDeltas]);
      for (const areaKey of Object.keys(rdata.byArea)) {
        const adata = rdata.byArea[areaKey];
        const aISTs = dayKeys.map(dk => getGroupedIST((dayData[dk] ? dayData[dk].stores||[] : []).filter(s => s.area === areaKey)));
        const aDeltas = dayKeys.slice(1).map((dk,i) => fmtDelta(aISTs[i], aISTs[i+1], false));
        trendRows.push(['AREA', region, adata.coach, areaKey, adata.stores.length+' Stores',...aISTs,...aDeltas]);
        adata.stores.forEach(s => {
          const sISTs = dayKeys.map(dk => {
            const match = (dayData[dk] ? dayData[dk].stores||[] : []).find(x => x.store_id === s.store_id);
            return match ? round1(match.ist_avg || match.in_store || match.wtd_in_store) : null;
          });
          const sDeltas = dayKeys.slice(1).map((dk,i) => fmtDelta(sISTs[i], sISTs[i+1], true));
          trendRows.push(['STORE', s.region_coach||'', s.area_coach||'', s.store_id||'', s.name||'',...sISTs,...sDeltas]);
        });
      }
    }
  }
  addTrendRows(stores);
  const trendWs = XLSX.utils.aoa_to_sheet(trendRows);
  XLSX.utils.book_append_sheet(wb, trendWs, 'Trend');

  // ─── PTD Trend (week-over-week within period) ───
  if (allData && allData.weeks && periodBase) {
    const ptdRawWeeks = Object.values(allData.weeks)
      .filter(w => w.period && w.period.replace(/W\d+$/, '') === periodBase)
      .sort((a, b) => (a.week||'').localeCompare(b.week||''));
    if (ptdRawWeeks.length > 1) {
      const ptdWkComputed = ptdRawWeeks.map(wk => computeWeek(wk));
      const wkLabels = ptdRawWeeks.map(w => w.period || w.week);
      const ptdTrendDeltaHdrs = wkLabels.slice(1).map((wl,i) => '\u0394 '+wkLabels[i]+'\u2192'+wl);
      const ptdTrendHdrs = ['Level','Region','Area Coach','Store #','Store Name',...wkLabels,...ptdTrendDeltaHdrs];
      const ptdTRows = [['PTD In-Store Time Trend \u2014 '+periodBase], ptdTrendHdrs];
      const ptdAllStores = {};
      ptdWkComputed.forEach(wk => (wk.stores||[]).forEach(s => { if (!ptdAllStores[s.store_id]) ptdAllStores[s.store_id] = s; }));
      const ptdSList = Object.values(ptdAllStores);
      const ptdByRegion = groupStores(ptdSList);
      const totWkISTs = ptdWkComputed.map(wk => calcTotals(wk.stores||[]).avgIST);
      const totWkDeltas = wkLabels.slice(1).map((wl,i) => fmtDelta(totWkISTs[i], totWkISTs[i+1], false));
      ptdTRows.push(['TOTAL','ALL REGIONS','','',ptdSList.length+' Stores',...totWkISTs,...totWkDeltas]);
      for (const region of Object.keys(ptdByRegion)) {
        const rdata = ptdByRegion[region];
        const rWkISTs = ptdWkComputed.map(wk => calcTotals((wk.stores||[]).filter(s => s.region_coach===region)).avgIST);
        const rWkDeltas = wkLabels.slice(1).map((wl,i) => fmtDelta(rWkISTs[i], rWkISTs[i+1], false));
        ptdTRows.push(['REGION', region,'','',rdata.stores.length+' Stores',...rWkISTs,...rWkDeltas]);
        for (const areaKey of Object.keys(rdata.byArea)) {
          const adata = rdata.byArea[areaKey];
          const aWkISTs = ptdWkComputed.map(wk => calcTotals((wk.stores||[]).filter(s => s.area===areaKey)).avgIST);
          const aWkDeltas = wkLabels.slice(1).map((wl,i) => fmtDelta(aWkISTs[i], aWkISTs[i+1], false));
          ptdTRows.push(['AREA', region, adata.coach, areaKey, adata.stores.length+' Stores',...aWkISTs,...aWkDeltas]);
          adata.stores.forEach(s => {
            const sWkISTs = ptdWkComputed.map(wk => {
              const match = (wk.stores||[]).find(x => x.store_id===s.store_id);
              return match ? round1(match.wtd_in_store||match.ist_avg) : null;
            });
            const sWkDeltas = wkLabels.slice(1).map((wl,i) => fmtDelta(sWkISTs[i], sWkISTs[i+1], true));
            ptdTRows.push(['STORE', s.region_coach||'', s.area_coach||'', s.store_id||'', s.name||'',...sWkISTs,...sWkDeltas]);
          });
        }
      }
      const ptdTrendWs = XLSX.utils.aoa_to_sheet(ptdTRows);
      XLSX.utils.book_append_sheet(wb, ptdTrendWs, 'PTD Trend');
    }
  }

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}


// Send test email
app.post('/api/automation/test-email', verifyAutomationAuth, async (req, res) => {
  try {
    const nodemailer = require('nodemailer');
    const { to } = req.body;
    
    if (!to) {
      return res.status(400).json({ error: 'Email address required' });
    }

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com',
        pass: process.env.VELOCITY_EMAIL_PASS || 'dewnkjrxbgmodwwd'
      }
    });

    const info = await transporter.sendMail({
      from: `"Velocity Reports" <${process.env.VELOCITY_EMAIL_USER || 'velocityai.reports@gmail.com'}>`,
      to: to,
      subject: 'Velocity Email Test - Success!',
      html: `
        <h1>✅ Velocity Email is Working!</h1>
        <p>If you received this email, the Velocity automation email system is configured correctly.</p>
        <p>Daily reports will be sent automatically at 7:00 AM.</p>
        <p>Time sent: ${new Date().toLocaleString()}</p>
      `
    });

    res.json({ success: true, messageId: info.messageId, response: info.response });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Get automation status
app.get('/api/automation/status', verifyAutomationAuth, (req, res) => {
  try {
    const allData = loadData();
    const weeks = Object.keys(allData.weeks || {}).sort((a, b) => b.localeCompare(a));
    
    // Get last upload info
    let lastUpload = null;
    if (weeks.length > 0) {
      const latestWeek = allData.weeks[weeks[0]];
      const days = Object.keys(latestWeek.days || {}).sort((a, b) => b.localeCompare(a));
      if (days.length > 0) {
        lastUpload = {
          date: days[0],
          week: weeks[0],
          storeCount: latestWeek.days[days[0]].stores?.length || 0
        };
      }
    }

    res.json({
      status: 'active',
      lastUpload,
      totalWeeks: weeks.length
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// =====================
// ODS Auto-Pull: fetch Daily Dispatch Performance from oneVIEW and ingest
// POST /api/automation/pull-ods
// =====================
app.post('/api/automation/pull-ods', verifyAutomationAuth, async (req, res) => {
  try {
    const https = require('https');
    const querystring = require('querystring');

    const ODS_ORG  = process.env.ODS_ORG      || 'dgi';
    const ODS_USER = process.env.ODS_USER      || 'hlacoste';
    const ODS_PASS = process.env.ODS_PASSWORD;
    if (!ODS_PASS) return res.status(500).json({ error: 'ODS_PASSWORD env variable not set on Render' });

    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = yesterday.toISOString().split('T')[0];
    console.log(`[ODS Pull] Fetching Daily Dispatch Performance for ${dateStr}`);

    function httpsReq(options, postData) {
      return new Promise((resolve, reject) => {
        const req = https.request(options, (r) => {
          const chunks = [];
          r.on('data', d => chunks.push(d));
          r.on('end', () => resolve({ body: Buffer.concat(chunks), headers: r.headers, statusCode: r.statusCode }));
        });
        req.on('error', reject);
        if (postData) req.write(postData);
        req.end();
      });
    }

    function mergeCookies(existing, newSetCookies) {
      const map = {};
      (existing || '').split('; ').forEach(c => { const [k,v] = c.split('='); if (k&&v) map[k.trim()]=v.trim(); });
      [].concat(newSetCookies||[]).forEach(c => { const p = c.split(';')[0]; const [k,v] = p.split('='); if (k&&v) map[k.trim()]=v.trim(); });
      return Object.entries(map).map(([k,v]) => `${k}=${v}`).join('; ');
    }

    const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36';

    // Step 1: GET login page for initial cookie
    const r1 = await httpsReq({ hostname:'bi.onedatasource.com', path:'/asp/login.html', method:'GET', headers:{'User-Agent':UA} });
    let cookie = mergeCookies('', r1.headers['set-cookie']);

    // Step 2: POST login
    const loginBody = querystring.stringify({ orgCode:ODS_ORG, userId:ODS_USER, password:ODS_PASS, _eventId:'login', locale:'en_US', timezone:'America/New_York' });
    const r2 = await httpsReq({
      hostname:'bi.onedatasource.com', path:'/asp/login.html', method:'POST',
      headers:{ 'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(loginBody),'Cookie':cookie,'User-Agent':UA }
    }, loginBody);
    cookie = mergeCookies(cookie, r2.headers['set-cookie']);
    if (r2.statusCode >= 400) return res.status(500).json({ error:`ODS login failed: HTTP ${r2.statusCode}` });

    // Step 3: GET report parameters page to get CSRF + flowExecutionKey
    const r3 = await httpsReq({
      hostname:'bi.onedatasource.com',
      path:'/asp/flow.html?_flowId=aboveStoreInStoreReportsFlow&_eventId=selectParameters&selectedReportId=457',
      method:'GET', headers:{'Cookie':cookie,'User-Agent':UA}
    });
    cookie = mergeCookies(cookie, r3.headers['set-cookie']);
    const html3 = r3.body.toString();
    const csrfMatch = html3.match(/name="OWASP_CSRFTOKEN"\s+value="([^"]+)"/);
    const csrfToken = csrfMatch ? csrfMatch[1] : '';
    const flowKeyMatch = html3.match(/_flowExecutionKey=([^"&\s]+)/);
    const flowKey = flowKeyMatch ? flowKeyMatch[1] : 'e1s1';
    console.log(`[ODS Pull] flowKey=${flowKey} csrf=${csrfToken?'ok':'missing'}`);

    // Step 4: POST form to get PDF
    const reportBody = querystring.stringify({
      _eventId:'retrieveReports', orgTypes:'territory', orgTypeValues:'26',
      storesInOrgType:'all', selectedDate:dateStr, exportFormat:'pdf', OWASP_CSRFTOKEN:csrfToken
    });
    const r4 = await httpsReq({
      hostname:'bi.onedatasource.com',
      path:`/asp/flow.html?_flowId=aboveStoreInStoreReportsFlow&_flowExecutionKey=${flowKey}&_eventId=selectParameters&selectedReportId=457`,
      method:'POST',
      headers:{ 'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(reportBody),'Cookie':cookie,'User-Agent':UA,'Referer':'https://bi.onedatasource.com/' }
    }, reportBody);
    cookie = mergeCookies(cookie, r4.headers['set-cookie']);
    const ct = r4.headers['content-type'] || '';
    console.log(`[ODS Pull] Report response: HTTP ${r4.statusCode}, type=${ct}, size=${r4.body.length}`);

    if (!ct.includes('pdf') && r4.body.length < 5000) {
      return res.status(500).json({ error:'Did not receive PDF', statusCode:r4.statusCode, contentType:ct, preview:r4.body.toString().substring(0,300) });
    }

    // Step 5: Save PDF and parse
    const tmpPdf = path.join(os.tmpdir(), `dispatch_${dateStr}.pdf`);
    fs.writeFileSync(tmpPdf, r4.body);
    const parsed = parseAboveStorePDFLocal(tmpPdf);
    try { fs.unlinkSync(tmpPdf); } catch(e) {}

    if (!parsed.stores || !parsed.stores.length) {
      return res.status(200).json({ success:false, message:'PDF downloaded but no stores parsed', date:dateStr, pdfBytes:r4.body.length });
    }

    // Step 6: Merge into wtd_data
    const weekKey = getWeekKey(dateStr);
    const allData2 = loadData();
    if (!allData2.weeks[weekKey]) allData2.weeks[weekKey] = { week:weekKey, period:FISCAL_CALENDAR[weekKey]||'', days:{} };
    if (!allData2.weeks[weekKey].days[dateStr]) allData2.weeks[weekKey].days[dateStr] = { date:dateStr, type:'ods_auto', stores:[], uploader:'ODS Auto' };
    const existing = {};
    (allData2.weeks[weekKey].days[dateStr].stores||[]).forEach(s => { existing[s.store_id]=s; });
    parsed.stores.forEach(s => {
      const align = ALIGNMENT[s.store_id];
      if (!align && !existing[s.store_id]) return;
      if (existing[s.store_id]) Object.assign(existing[s.store_id], s);
      else existing[s.store_id] = { ...s, ...(align||{}) };
    });
    allData2.weeks[weekKey].days[dateStr].stores = Object.values(existing);
    saveData(allData2);
    console.log(`[ODS Pull] Done: ${parsed.stores.length} stores ingested for ${dateStr}`);
    res.json({ success:true, date:dateStr, week:weekKey, storeCount:parsed.stores.length });

  } catch(e) {
    console.error('[ODS Pull] Error:', e);
    res.status(500).json({ error:e.message });
  }
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`Velocity running on port ${PORT}`));
