const express = require('express');
const https   = require('https');
const http    = require('http');
const { URL } = require('url');
const path    = require('path');

const app = express();

// ── CORS (necessário para o dashboard chamar /proxy) ─────────────────────────
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

// ── Proxy → API Closum ────────────────────────────────────────────────────────
// Chamada:  GET /proxy?url=https://api.closum.com/v2/lead/?api-key=...
// (mesmo padrão que o proxy.ps1 usava em localhost:3001)
app.get('/proxy', (req, res) => {
  const target = req.query.url;
  if (!target) return res.status(400).json({ error: 'url param required' });

  let parsed;
  try { parsed = new URL(target); }
  catch (e) { return res.status(400).json({ error: 'invalid url' }); }

  // Whitelist: só api.closum.com
  if (!parsed.hostname.endsWith('closum.com')) {
    return res.status(403).json({ error: 'forbidden domain' });
  }

  const lib = parsed.protocol === 'https:' ? https : http;
  const options = {
    hostname : parsed.hostname,
    port     : parsed.port || (parsed.protocol === 'https:' ? 443 : 80),
    path     : parsed.pathname + (parsed.search || ''),
    method   : 'GET',
    headers  : {
      'User-Agent' : 'TheCode-Dashboard/2.0',
      'Accept'     : 'application/json',
    },
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
