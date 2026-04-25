const express = require('express');
const Database = require('better-sqlite3');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = 3000;
const ADMIN_PASSWORD = 'prayer2025';

// --- Database setup ---
const db = new Database(path.join(__dirname, 'prayers.db'));
db.pragma('journal_mode = WAL');

db.exec(`
  CREATE TABLE IF NOT EXISTS prayers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT NOT NULL,
    name TEXT NOT NULL DEFAULT 'Anonymous',
    request TEXT NOT NULL,
    duration TEXT NOT NULL DEFAULT 'one-time',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    expires_at TEXT
  )
`);

// Add columns if upgrading from older schema
try { db.exec(`ALTER TABLE prayers ADD COLUMN duration TEXT NOT NULL DEFAULT 'one-time'`); } catch(e) {}
try { db.exec(`ALTER TABLE prayers ADD COLUMN expires_at TEXT`); } catch(e) {}

// --- Duration config ---
const durations = {
  'one-time':   { label: 'One-time',    days: 1 },
  '3-months':   { label: '3 Months',    days: 90 },
  '6-months':   { label: '6 Months',    days: 180 },
  '9-months':   { label: '9 Months',    days: 270 },
  '12-months':  { label: '12 Months',   days: 365 },
  '2-years':    { label: '2 Years',     days: 730 },
  '3-years':    { label: '3 Years',     days: 1095 },
  'perpetual':  { label: 'Perpetual',   days: null }
};

function calculateExpiry(duration) {
  const d = durations[duration];
  if (!d || d.days === null) return null; // perpetual = no expiry
  const now = new Date();
  now.setDate(now.getDate() + d.days);
  return now.toISOString().slice(0, 19).replace('T', ' ');
}

// --- Auto-delete expired prayers ---
function purgeExpired() {
  const result = db.prepare(
    `DELETE FROM prayers WHERE expires_at IS NOT NULL AND expires_at <= datetime('now','localtime')`
  ).run();
  if (result.changes > 0) {
    console.log(`Purged ${result.changes} expired prayer(s)`);
  }
}

// Purge on startup
purgeExpired();
// Purge every hour
setInterval(purgeExpired, 60 * 60 * 1000);

// --- Middleware ---
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// --- Categories ---
const categories = {
  ill:          'For the Ill',
  dead:         'For the Deceased',
  job:          'Job Search',
  pregnancy:    'Pregnancy & Children',
  immigration:  'Immigration',
  reversion:    'Reversion/Conversion',
  spouse:       'Future Spouse'
};

// --- API Routes ---

// Get all prayers (ADMIN ONLY) — excludes expired
app.post('/api/admin/prayers', (req, res) => {
  const { password, category, search } = req.body;
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  let sql = 'SELECT * FROM prayers WHERE (expires_at IS NULL OR expires_at > datetime(\'now\',\'localtime\'))';
  const params = [];

  if (category) {
    sql += ' AND category = ?';
    params.push(category);
  }
  if (search) {
    sql += ' AND (name LIKE ? OR request LIKE ?)';
    params.push(`%${search}%`, `%${search}%`);
  }

  sql += ' ORDER BY created_at DESC';

  const prayers = db.prepare(sql).all(...params);
  res.json(prayers);
});

// Create a prayer
app.post('/api/prayers', (req, res) => {
  const { category, name, request, duration } = req.body;

  if (!category || !categories[category]) {
    return res.status(400).json({ error: 'Invalid category' });
  }
  if (!request || !request.trim()) {
    return res.status(400).json({ error: 'Prayer request is required' });
  }
  const dur = durations[duration] ? duration : 'one-time';
  const expiresAt = calculateExpiry(dur);

  const stmt = db.prepare(
    'INSERT INTO prayers (category, name, request, duration, expires_at) VALUES (?, ?, ?, ?, ?)'
  );
  const result = stmt.run(
    category,
    (name && name.trim()) || 'Anonymous',
    request.trim(),
    dur,
    expiresAt
  );

  const prayer = db.prepare('SELECT * FROM prayers WHERE id = ?').get(result.lastInsertRowid);
  res.status(201).json(prayer);
});

// Delete a prayer (ADMIN ONLY)
app.post('/api/admin/prayers/delete', (req, res) => {
  const { password, id } = req.body;
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const result = db.prepare('DELETE FROM prayers WHERE id = ?').run(id);
  if (result.changes === 0) {
    return res.status(404).json({ error: 'Prayer not found' });
  }
  res.json({ success: true });
});

// --- Health check (for Render) ---
app.get('/healthz00m', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// --- Admin: verify password ---
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  if (password === ADMIN_PASSWORD) {
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Invalid password' });
  }
});

// --- Admin: export to Excel ---
app.post('/api/admin/export', (req, res) => {
  const { password } = req.body;
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const wb = XLSX.utils.book_new();

  for (const [key, label] of Object.entries(categories)) {
    const rows = db.prepare(
      `SELECT name AS "Requestor", request AS "Intercession For", duration AS "Duration", created_at AS "Submitted", expires_at AS "Expires"
       FROM prayers WHERE category = ? AND (expires_at IS NULL OR expires_at > datetime('now','localtime'))
       ORDER BY created_at DESC`
    ).all(key);

    // Map duration keys to labels
    const mapped = rows.map(r => ({
      ...r,
      Duration: durations[r.Duration] ? durations[r.Duration].label : r.Duration,
      Expires: r.Expires || 'Never'
    }));

    const ws = XLSX.utils.json_to_sheet(
      mapped.length > 0 ? mapped : [{ Requestor: '', 'Intercession For': '', Duration: '', Submitted: '', Expires: '' }]
    );

    ws['!cols'] = [{ wch: 20 }, { wch: 60 }, { wch: 14 }, { wch: 20 }, { wch: 20 }];

    const sheetName = label.length > 31 ? label.substring(0, 31) : label;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename="prayer-requests.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// --- Admin: get stats ---
app.post('/api/admin/stats', (req, res) => {
  const { password } = req.body;
  if (password !== ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const active = `(expires_at IS NULL OR expires_at > datetime('now','localtime'))`;
  const total = db.prepare(`SELECT COUNT(*) as count FROM prayers WHERE ${active}`).get().count;
  const byCategory = db.prepare(
    `SELECT category, COUNT(*) as count FROM prayers WHERE ${active} GROUP BY category`
  ).all();

  res.json({ total, byCategory });
});

// --- Start server ---
app.listen(PORT, () => {
  console.log(`Prayer App running at http://localhost:${PORT}`);
  console.log(`Admin password: ${ADMIN_PASSWORD}`);
});
