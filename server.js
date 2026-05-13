const express = require('express');
const https   = require('https');
const http    = require('http');
const { URL } = require('url');
const path    = require('path');
const { Pool } = require('pg');

const app = express();
app.use(express.json({ limit: '20mb' }));

// ── PostgreSQL ────────────────────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false,
});

pool.query(`
  CREATE TABLE IF NOT EXISTS app_state (
    id          INTEGER PRIMARY KEY DEFAULT 1,
    data        JSONB NOT NULL,
    updated_at  TIMESTAMPTZ DEFAULT NOW()
  )
`).catch(e => console.error('DB init error:', e.message));

// ── CORS ──────────────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// ── Serve o dashboard ─────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'Dashboard_v2.html'));
});

// ── API: ler estado ───────────────────────────────────────────────────────────
app.get('/api/state', async (req, res) => {
  try {
    const result = await pool.query('SELECT data, updated_at FROM app_state WHERE id = 1');
    if (result.rows.length === 0) return res.json({ data: null });
    res.json({ data: result.rows[0].data, updated_at: result.rows[0].updated_at });
  } catch (e) {
    console.error('GET /api/state:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── API: guardar estado ───────────────────────────────────────────────────────
app.post('/api/state', async (req, res) => {
  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ error: 'data required' });
    await pool.query(`
      INSERT INTO app_state (id, data, updated_at) VALUES (1, $1, NOW())
      ON CONFLICT (id) DO UPDATE SET data = $1, updated_at = NOW()
    `, [data]);
    res.json({ ok: true });
  } catch (e) {
    console.error('POST /api/state:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── Proxy → API Closum ────────────────────────────────────────────────────────
app.get('/proxy', (req, res) => {
  const target = req.query.url;
  if (!target) return res.status(400).json({ error: 'url param required' });

  let parsed;
  try { parsed = new URL(target); }
  catch (e) { return res.status(400).json({ error: 'invalid url' }); }

  if (!parsed.hostname.endsWith('closum.com')) {
    return res.status(403).json({ error: 'forbidden domain' });
  }

  const lib = parsed.protocol === 'https:' ? https : http;
  const options = {
    hostname : parsed.hostname,
    port     : parsed.port || (parsed.protocol === 'https:' ? 443 : 80),
    path     : parsed.pathname + (parsed.search || ''),
    method   : 'GET',
    headers  : { 'User-Agent': 'TheCode-Dashboard/2.0', 'Accept': 'application/json' },
  };

  const proxyReq = lib.request(options, (proxyRes) => {
    res.set('Access-Control-Allow-Origin', '*');
    res.set('Content-Type', proxyRes.headers['content-type'] || 'application/json');
    res.status(proxyRes.statusCode);
    proxyRes.pipe(res);
  });

  proxyReq.on('error', (err) => {
    console.error('Proxy error:', err.message);
    res.status(502).json({ error: 'proxy error', detail: err.message });
  });

  proxyReq.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`The Code Dashboard → http://localhost:${PORT}`));
