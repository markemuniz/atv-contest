const express = require('express');
const multer  = require('multer');
const fs      = require('fs');
const path    = require('path');
const XLSX    = require('xlsx');
const Papa    = require('papaparse');
const jwt     = require('jsonwebtoken');
const cookieParser = require('cookie-parser');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Config ────────────────────────────────────────────────────────────────
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'travel2026';
const JWT_SECRET     = process.env.JWT_SECRET     || 'atv-secret-key-change-in-prod';
const DATA_FILE      = path.join(__dirname, 'data', 'leaderboard.json');

// ── Middleware ─────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use('/public', express.static(path.join(__dirname, 'public')));

// Multer for file uploads (memory storage)
const upload = multer({ storage: multer.memoryStorage() });

// Ensure data dir exists
if (!fs.existsSync(path.join(__dirname, 'data')))
  fs.mkdirSync(path.join(__dirname, 'data'));
if (!fs.existsSync(DATA_FILE))
  fs.writeFileSync(DATA_FILE, JSON.stringify({ days: {} }));

// ── Auth middleware ────────────────────────────────────────────────────────
function requireAdmin(req, res, next) {
  const token = req.cookies?.admin_token || req.headers?.authorization?.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try {
    jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Invalid or expired token' });
  }
}

// ── Processing logic (same as HTML version) ───────────────────────────────
const EXCLUDE_STARTS   = ['Jude','Kaleo','Bruce','Jeff','Kevin','Ian','Cynthia','Nick'];
const EXCLUDE_CONTAINS = ['DTBK'];

function isExcluded(name) {
  if (!name || name.trim().toLowerCase() === 'blank') return true;
  if (EXCLUDE_CONTAINS.some(s => name.includes(s))) return true;
  if (EXCLUDE_STARTS.some(s => name.startsWith(s))) return true;
  return false;
}

function norm(s) {
  return (s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseXLSXDate(raw) {
  if (!raw) return { dateStr: null };
  const s = String(raw).trim();
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})/);
  if (mdy) {
    let [, mo, dd, yy] = mdy;
    if (yy.length === 2) yy = '20' + yy;
    return { dateStr: `${yy}-${mo.padStart(2,'0')}-${dd.padStart(2,'0')}` };
  }
  const iso = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return { dateStr: iso[1] };
  return { dateStr: null };
}

function parseSalesReport(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws['!ref']);
  let headerRow = -1;
  for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell && cell.v === 'Order ID') { headerRow = r; break; }
  }
  if (headerRow < 0) throw new Error('Could not find header row in sales report');
  const data = XLSX.utils.sheet_to_json(ws, { range: headerRow, raw: false });

  const orderSales = new Map();
  const orderMeta  = new Map();
  const dateCounts = {};

  for (const row of data) {
    const id = String(row['Order ID'] || '').trim();
    if (!id) continue;
    const sales = parseFloat(row['Net Sales']) || 0;
    const cust  = norm(String(row['Customer Name'] || ''));
    const { dateStr } = parseXLSXDate(row['Order Time']);
    if (dateStr) dateCounts[dateStr] = (dateCounts[dateStr] || 0) + 1;
    orderSales.set(id, (orderSales.get(id) || 0) + sales);
    if (!orderMeta.has(id)) orderMeta.set(id, { cust_norm: cust });
  }

  const reportDate = Object.entries(dateCounts).sort((a, b) => b[1] - a[1])[0]?.[0];
  return { reportDate, orderSales, orderMeta };
}

function parseAttrCSV(buffer, reportDate) {
  const text = buffer.toString('utf8');
  const { data: rows } = Papa.parse(text, { header: true, skipEmptyLines: true });
  const out = [];
  for (const row of rows) {
    const helped = (row['helped_by'] || '').trim().replace(/\.\s*$/, '');
    if (isExcluded(helped)) continue;
    const dtUTC  = new Date(row['order_time']);
    const dtET   = new Date(dtUTC.getTime() - 4 * 60 * 60 * 1000);
    const rowDate = dtET.toISOString().slice(0, 10);
    if (rowDate !== reportDate) continue;
    const cust_norm = norm((row['customer_first'] || '') + ' ' + (row['customer_last'] || ''));
    out.push({ helped_by: helped, cust_norm });
  }
  return out;
}

function matchAndCalc(attrRows, orderSales, orderMeta) {
  const custLookup = new Map();
  for (const [id, { cust_norm }] of orderMeta) {
    if (!custLookup.has(cust_norm)) custLookup.set(cust_norm, []);
    custLookup.get(cust_norm).push(id);
  }
  const used = new Set();
  const result = {};
  for (const attr of attrRows) {
    const candidates = custLookup.get(attr.cust_norm) || [];
    const orderId = candidates.find(id => !used.has(id));
    if (!orderId) continue;
    used.add(orderId);
    const bud = attr.helped_by;
    if (!result[bud]) result[bud] = { orders: 0, totalSales: 0 };
    result[bud].orders++;
    result[bud].totalSales += orderSales.get(orderId) || 0;
  }
  for (const b of Object.keys(result))
    result[b].avgTicket = result[b].totalSales / result[b].orders;
  return result;
}

// ── Routes ─────────────────────────────────────────────────────────────────

// Admin login
app.post('/api/login', (req, res) => {
  const { password } = req.body;
  if (password !== ADMIN_PASSWORD)
    return res.status(401).json({ error: 'Wrong password' });
  const token = jwt.sign({ admin: true }, JWT_SECRET, { expiresIn: '24h' });
  res.cookie('admin_token', token, { httpOnly: true, maxAge: 86400000 });
  res.json({ ok: true });
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('admin_token');
  res.json({ ok: true });
});

// Upload reports
app.post('/api/upload',
  requireAdmin,
  upload.fields([{ name: 'sales', maxCount: 10 }, { name: 'attr', maxCount: 10 }]),
  (req, res) => {
    try {
      const salesFiles = req.files?.['sales'] || [];
      const attrFiles  = req.files?.['attr']  || [];
      if (!salesFiles.length || !attrFiles.length)
        return res.status(400).json({ error: 'Need at least one sales report and one attributed CSV' });

      const db = JSON.parse(fs.readFileSync(DATA_FILE));
      const newDates = [];
      const skipped  = [];

      for (const salesFile of salesFiles) {
        const { reportDate, orderSales, orderMeta } = parseSalesReport(salesFile.buffer);
        if (!reportDate) { skipped.push(salesFile.originalname + ' (no date)'); continue; }
        if (db.days[reportDate]) { skipped.push(reportDate + ' (already loaded)'); continue; }

        // Try each attr CSV for this date
        const dayBuds = {};
        for (const attrFile of attrFiles) {
          const attrRows = parseAttrCSV(attrFile.buffer, reportDate);
          if (!attrRows.length) continue;
          const dayData = matchAndCalc(attrRows, orderSales, orderMeta);
          for (const [bud, stats] of Object.entries(dayData)) {
            if (!dayBuds[bud]) dayBuds[bud] = { orders: 0, totalSales: 0 };
            dayBuds[bud].orders     += stats.orders;
            dayBuds[bud].totalSales += stats.totalSales;
          }
        }

        for (const b of Object.keys(dayBuds))
          dayBuds[b].avgTicket = dayBuds[b].totalSales / dayBuds[b].orders;

        if (Object.keys(dayBuds).length) {
          db.days[reportDate] = dayBuds;
          newDates.push(reportDate);
        } else {
          skipped.push(reportDate + ' (no matches found)');
        }
      }

      fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2));
      res.json({ ok: true, added: newDates, skipped });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  }
);

// Delete a day
app.delete('/api/day/:date', requireAdmin, (req, res) => {
  const db = JSON.parse(fs.readFileSync(DATA_FILE));
  if (!db.days[req.params.date])
    return res.status(404).json({ error: 'Date not found' });
  delete db.days[req.params.date];
  fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2));
  res.json({ ok: true });
});

// Get leaderboard data (public)
app.get('/api/leaderboard', (req, res) => {
  const db = JSON.parse(fs.readFileSync(DATA_FILE));
  res.json(db);
});

// ── Pages ──────────────────────────────────────────────────────────────────
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'leaderboard.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

app.listen(PORT, () => console.log(`ATV Contest running on http://localhost:${PORT}`));
