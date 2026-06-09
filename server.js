const express = require('express');
const https   = require('https');
const http    = require('http');
const { URL } = require('url');
const path    = require('path');
const { Pool } = require('pg');
const XLSX    = require('xlsx');

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

// ── API: guardar estado (shallow JSONB merge — atómico) ──────────────────────
app.post('/api/state', async (req, res) => {
  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ error: 'data required' });

    const result = await pool.query(`
      INSERT INTO app_state (id, data, updated_at) VALUES (1, $1::jsonb, NOW())
      ON CONFLICT (id) DO UPDATE
        SET data = COALESCE(app_state.data, '{}'::jsonb) || $1::jsonb,
            updated_at = NOW()
      RETURNING updated_at
    `, [JSON.stringify(data)]);

    res.json({ ok: true, updated_at: result.rows[0].updated_at });
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

// ════════════════════════════════════════════════════════════════════════════
//                           ONEDRIVE EXCEL SYNC
// ════════════════════════════════════════════════════════════════════════════

const EXCEL_SYNC_CUTOFF = '2026-06-09'; // Cutoff fixo: ignora linhas com data anterior

const CLUBES = [
  { unidade: 'Saldanha',             comercial: 'Íris',     envVar: 'ONEDRIVE_LINK_SALDANHA' },
  { unidade: 'Nova',                 comercial: 'Diogo',    envVar: 'ONEDRIVE_LINK_NOVA' },
  { unidade: 'Cidade Universitária', comercial: 'Henrique', envVar: 'ONEDRIVE_LINK_CU' },
];

const MONTHS_PT_SHORT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

const FIELD_MAP = {
  'Efetuadas': 'chamadasEf',
  'Atendidas': 'chamadasAt',
  'Vis.Marc.': 'marc',
  'Vis.Real.': 'visitas',
  'Vendas':    'conv',
  'Marc.Dia':  'marcDia',
  'LEADS':     'leadsRec',
  'WI':        'walkin',
  'POS':       'pos',
  'REFS':      'contAng',
  'OUT':       'outbound',
};

// ── Encode share URL → base64url (formato exigido pelo Shares API) ───────────
function encodeShareUrl(shareUrl) {
  return Buffer.from(shareUrl).toString('base64')
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
}

// ── Extrai o link canónico (1drv.ms) do parâmetro "redeem" do URL longo ──────
function getCanonicalShareUrl(url) {
  try {
    const u = new URL(url);
    const redeem = u.searchParams.get('redeem');
    if (redeem) {
      const b64 = redeem.replace(/-/g,'+').replace(/_/g,'/');
      const pad = b64.length % 4 === 0 ? '' : '='.repeat(4 - (b64.length % 4));
      const decoded = Buffer.from(b64 + pad, 'base64').toString('utf-8');
      if (decoded.startsWith('http')) return decoded;
    }
  } catch (_) {}
  return url;
}

// ── Download .xlsx — tenta vários endpoints até um funcionar ─────────────────
async function downloadShared(shareUrl) {
  const canonical = getCanonicalShareUrl(shareUrl);
  const b64canonical = encodeShareUrl(canonical);
  const b64original  = encodeShareUrl(shareUrl);

  const attempts = [
    { name: 'graph/shares (canonical)',    url: `https://graph.microsoft.com/v1.0/shares/u!${b64canonical}/driveItem/content` },
    { name: 'api.onedrive/shares (canonical)', url: `https://api.onedrive.com/v1.0/shares/u!${b64canonical}/root/content` },
    { name: 'api.onedrive/shares (original)',  url: `https://api.onedrive.com/v1.0/shares/u!${b64original}/root/content` },
    { name: 'canonical + download=1', url: canonical + (canonical.includes('?')?'&':'?') + 'download=1' },
    { name: 'original + download=1',  url: shareUrl + (shareUrl.includes('?')?'&':'?') + 'download=1' },
  ];

  const headers = { 'User-Agent': 'Mozilla/5.0 (compatible; TheCode-Dashboard/2.0)' };
  const errors = [];

  for (const a of attempts) {
    try {
      const res = await fetch(a.url, { redirect: 'follow', headers });
      if (!res.ok) { errors.push(`${a.name}: HTTP ${res.status}`); continue; }
      const buf = Buffer.from(await res.arrayBuffer());
      // Sanity: ficheiros .xlsx começam sempre por PK (assinatura ZIP, 50 4B)
      if (buf.length > 200 && buf[0] === 0x50 && buf[1] === 0x4B) {
        console.log(`[Excel sync] ✓ download via ${a.name} (${buf.length} bytes)`);
        return buf;
      }
      // Pode ter recebido HTML (página de login/redirect)
      const snippet = buf.slice(0, 80).toString('utf-8').replace(/\s+/g,' ').slice(0, 60);
      errors.push(`${a.name}: not xlsx (${buf.length}B, "${snippet}")`);
    } catch (e) {
      errors.push(`${a.name}: ${e.message}`);
    }
  }

  throw new Error('all methods failed → ' + errors.join(' | '));
}

// ── Normalize date cell value → "YYYY-MM-DD" ─────────────────────────────────
function normalizeDate(v, year) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0') + '-' + String(v.getDate()).padStart(2,'0');
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return d.y + '-' + String(d.m).padStart(2,'0') + '-' + String(d.d).padStart(2,'0');
  }
  if (typeof v === 'string') {
    const s = v.trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2,'0') + '-' + m[3].padStart(2,'0');
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return m[3] + '-' + m[2].padStart(2,'0') + '-' + m[1].padStart(2,'0');
    m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (m) return year + '-' + m[2].padStart(2,'0') + '-' + m[1].padStart(2,'0');
  }
  return null;
}

// ── Parse one monthly sheet → array of inputs rows ───────────────────────────
function parseSheet(sheet, year) {
  // 2D array (header:1 keeps positional indexing)
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

  // Find header row: row containing both "DATA" and "Efetuadas"
  let headerRow = -1;
  for (let i = 0; i < Math.min(20, grid.length); i++) {
    const row = (grid[i] || []).map(c => (c == null ? '' : String(c).trim()));
    if (row.includes('DATA') && row.includes('Efetuadas')) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return [];

  const headers = (grid[headerRow] || []).map(c => (c == null ? '' : String(c).trim()));
  const colMap = {};
  headers.forEach((h, i) => { if (h) colMap[h] = i; });

  const dataCol = colMap['DATA'];
  if (dataCol === undefined) return [];

  // First pass: build inputs rows from Chamadas&Visitas + Funil de Leads sections
  const inputs = [];
  for (let i = headerRow + 1; i < grid.length; i++) {
    const row = grid[i] || [];
    const dateStr = normalizeDate(row[dataCol], year);
    if (!dateStr) continue;

    const obj = { data: dateStr };
    for (const [excelH, schemaKey] of Object.entries(FIELD_MAP)) {
      const col = colMap[excelH];
      if (col === undefined) continue;
      const val = row[col];
      if (val === null || val === undefined || val === '' || val === '-' || val === '—') continue;
      const num = Number(val);
      if (!isNaN(num)) obj[schemaKey] = num;
    }
    inputs.push(obj);
  }

  // Second pass: Inscrições → count Source="WI" per Data Ins. → walkinSale
  const dataInsCol = colMap['Data Ins.'];
  const sourceCol  = colMap['Source'];
  const wiByDate = {};
  if (dataInsCol !== undefined && sourceCol !== undefined) {
    for (let i = headerRow + 1; i < grid.length; i++) {
      const row = grid[i] || [];
      const ds = normalizeDate(row[dataInsCol], year);
      const src = String(row[sourceCol] || '').trim().toUpperCase();
      if (!ds || src !== 'WI') continue;
      wiByDate[ds] = (wiByDate[ds] || 0) + 1;
    }
  }

  // Merge walkinSale onto inputs
  inputs.forEach(r => { r.walkinSale = wiByDate[r.data] || 0; });

  // Always set outbound default 0 if absent
  inputs.forEach(r => { if (r.outbound === undefined) r.outbound = 0; });

  return inputs;
}

// ── Sync one club's file ──────────────────────────────────────────────────────
async function syncClub(club) {
  const shareUrl = process.env[club.envVar];
  if (!shareUrl) throw new Error(`${club.envVar} not set`);

  const buf = await downloadShared(shareUrl);
  const wb  = XLSX.read(buf, { type: 'buffer', cellDates: true });

  const now = new Date();
  const year = now.getFullYear();
  const monthName = MONTHS_PT_SHORT[now.getMonth()];

  // Find sheet matching current month (case-insensitive, accent-tolerant)
  const norm = s => String(s||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase();
  const target = norm(monthName);
  const sheetName = wb.SheetNames.find(n => norm(n) === target);
  if (!sheetName) throw new Error(`Sheet "${monthName}" not found in workbook`);

  const sheet = wb.Sheets[sheetName];
  const rows  = parseSheet(sheet, year);

  // Filter by cutoff and attach unidade + comercial
  const filtered = rows
    .filter(r => r.data >= EXCEL_SYNC_CUTOFF)
    .map(r => ({ ...r, unidade: club.unidade, comercial: club.comercial }));

  return { unidade: club.unidade, count: filtered.length, rows: filtered, monthName };
}

// ── Sync all clubs and persist to DB ─────────────────────────────────────────
let _syncInProgress = false;
async function runExcelSync(triggeredBy = 'cron') {
  if (_syncInProgress) return { skipped: true, reason: 'already in progress' };
  _syncInProgress = true;
  const startedAt = new Date().toISOString();
  const summary = { startedAt, triggeredBy, perClub: [], total: 0, errors: [] };

  try {
    for (const club of CLUBES) {
      try {
        const r = await syncClub(club);
        summary.perClub.push({ unidade: r.unidade, count: r.count, month: r.monthName });
        summary.total += r.count;
        summary.__rows = (summary.__rows || []).concat(r.rows);
      } catch (e) {
        console.error(`[Excel sync] ${club.unidade} failed:`, e.message);
        summary.errors.push({ unidade: club.unidade, error: e.message });
      }
    }

    // Read current state (apenas para preservar inputs legacy anteriores ao cutoff)
    const cur = await pool.query('SELECT data FROM app_state WHERE id = 1');
    const dbState = (cur.rows[0] && cur.rows[0].data) ? cur.rows[0].data : {};
    const oldInputs = Array.isArray(dbState.inputs) ? dbState.inputs : [];
    const keepOld = oldInputs.filter(r => {
      const d = String(r.data || '').slice(0, 10);
      return d && d < EXCEL_SYNC_CUTOFF;
    });
    const newInputs = [...keepOld, ...(summary.__rows || [])];
    const excelSyncStatus = {
      lastRun: startedAt,
      finishedAt: new Date().toISOString(),
      perClub: summary.perClub,
      errors: summary.errors,
      total: summary.total,
      triggeredBy,
    };

    // Garantir que a linha existe (caso seja a primeira escrita) e fazer UPDATE
    // ATÓMICO apenas dos campos geridos pelo sync (inputs + excelSync), via
    // jsonb_set. Outros campos (kpisM, metas, tarefas, etc.) NÃO são tocados
    // — evita race conditions com edições do utilizador durante a sync.
    await pool.query(`
      INSERT INTO app_state (id, data) VALUES (1, '{}'::jsonb)
      ON CONFLICT (id) DO NOTHING
    `);
    await pool.query(`
      UPDATE app_state
      SET data = jsonb_set(
            jsonb_set(COALESCE(data, '{}'::jsonb), '{inputs}',    $1::jsonb, true),
            '{excelSync}', $2::jsonb, true
          ),
          updated_at = NOW()
      WHERE id = 1
    `, [JSON.stringify(newInputs), JSON.stringify(excelSyncStatus)]);

    console.log(`[Excel sync] OK · ${summary.total} rows · ${summary.errors.length} errors`);
    delete summary.__rows;
    return summary;
  } catch (e) {
    console.error('[Excel sync] FATAL:', e.message);
    summary.errors.push({ unidade: null, error: e.message });
    return summary;
  } finally {
    _syncInProgress = false;
  }
}

// ── Endpoints ────────────────────────────────────────────────────────────────
app.get('/api/excel-sync', async (req, res) => {
  try {
    const cur = await pool.query('SELECT data FROM app_state WHERE id = 1');
    const sync = (cur.rows[0] && cur.rows[0].data && cur.rows[0].data.excelSync) || null;
    res.json({ status: sync, cutoff: EXCEL_SYNC_CUTOFF, inProgress: _syncInProgress });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/excel-sync', async (req, res) => {
  try {
    const result = await runExcelSync('manual');
    res.json({ ok: true, summary: result });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Schedule: every 30 minutes + 30s after startup ───────────────────────────
const SYNC_INTERVAL_MS = 30 * 60 * 1000;
function startExcelSyncScheduler() {
  const hasAnyLink = CLUBES.some(c => process.env[c.envVar]);
  if (!hasAnyLink) {
    console.log('[Excel sync] No ONEDRIVE_LINK_* env vars set — scheduler not started.');
    return;
  }
  console.log('[Excel sync] Scheduler ON — first run in 30s, then every 30 min.');
  setTimeout(() => runExcelSync('startup'), 30 * 1000);
  setInterval(() => runExcelSync('cron'), SYNC_INTERVAL_MS);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`The Code Dashboard → http://localhost:${PORT}`);
  // Sync OneDrive automático desactivado — import é manual via UI.
  // startExcelSyncScheduler();
});
