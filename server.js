const express = require('express');
const path = require('path');
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
// CONFIGURACION — edita estos valores antes de desplegar
// ============================================================
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || 'TU_API_KEY_AQUI';
const PORT = process.env.PORT || 3000;

// Config QAD (también se puede actualizar desde el panel admin)
let qadConfig = {
  url:     process.env.QAD_URL     || '',
  port:    process.env.QAD_PORT    || '',
  token:   process.env.QAD_TOKEN   || '',
  company: process.env.QAD_COMPANY || 'VIVATEX',
  user:    process.env.QAD_USER    || '',
  pass:    process.env.QAD_PASS    || '',
};
// ============================================================

// ---- Ruta principal chat ----
app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body;
  if (!messages || !system) return res.status(400).json({ error: 'Faltan datos' });

  try {
    // 1. Intentar obtener datos de QAD si la pregunta lo requiere
    const lastMsg = messages[messages.length - 1]?.content || '';
    let qadData = '';
    if (qadConfig.url) {
      qadData = await fetchQADData(lastMsg);
    }

    // 2. Agregar datos QAD al contexto si existen
    const fullSystem = qadData
      ? `${system}\n\nDatos actuales de QAD:\n${qadData}`
      : system;

    // 3. Llamar a Claude API
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5',
        max_tokens: 2048,
        system: fullSystem,
        messages: messages,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error('Anthropic API error:', errText);
      return res.status(500).json({ error: 'Error en API de IA' });
    }

    const data = await response.json();
    const reply = data.content?.[0]?.text || 'Sin respuesta';
    res.json({ reply });

  } catch (err) {
    console.error('Error en /api/chat:', err);
    res.status(500).json({ error: 'Error interno del servidor' });
  }
});

// ---- Función para obtener datos de QAD ----
async function fetchQADData(userQuery) {
  // Detectar qué módulo de QAD consultar según la pregunta
  const q = userQuery.toLowerCase();
  let endpoint = '';

  if (q.includes('venta') || q.includes('orden') || q.includes('pedido'))
    endpoint = '/api/sales/summary';
  else if (q.includes('inventario') || q.includes('existencia') || q.includes('stock'))
    endpoint = '/api/inventory/status';
  else if (q.includes('compra') || q.includes('proveedor'))
    endpoint = '/api/purchase/orders';
  else if (q.includes('cliente'))
    endpoint = '/api/customer/list';
  else if (q.includes('financiero') || q.includes('finanza') || q.includes('balance'))
    endpoint = '/api/finance/summary';
  else
    return ''; // No requiere datos de QAD

  try {
    const baseUrl = `${qadConfig.url}:${qadConfig.port || 80}${endpoint}`;
    const response = await fetch(baseUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${qadConfig.token}`,
        'X-QAD-Company': qadConfig.company,
        'Content-Type': 'application/json',
      },
      signal: AbortSignal.timeout(8000),
    });

    if (!response.ok) return '';
    const data = await response.json();
    return JSON.stringify(data, null, 2);
  } catch {
    return ''; // QAD no disponible, VIVA responde sin datos en tiempo real
  }
}

// ---- Rutas de configuración QAD (solo desde panel admin) ----
app.post('/api/admin/qad-config', (req, res) => {
  const { url, port, token, company, user, pass } = req.body;
  qadConfig = { url, port, token, company, user, pass };
  res.json({ ok: true });
});

app.get('/api/admin/qad-status', async (req, res) => {
  if (!qadConfig.url) return res.json({ connected: false, reason: 'Sin configurar' });
  try {
    const r = await fetch(`${qadConfig.url}:${qadConfig.port || 80}/ping`, {
      headers: { 'Authorization': `Bearer ${qadConfig.token}` },
      signal: AbortSignal.timeout(4000),
    });
    res.json({ connected: r.ok, status: r.status });
  } catch (e) {
    res.json({ connected: false, reason: e.message });
  }
});

// ---- PWA Manifest ----
app.get('/manifest.json', (req, res) => {
  res.json({
    name: 'VIVA — Vivatex IA',
    short_name: 'VIVA',
    description: 'Asistente inteligente de Grupo Vivatex',
    start_url: '/',
    display: 'standalone',
    background_color: '#1a1f16',
    theme_color: '#6BBF3E',
    orientation: 'portrait-primary',
    icons: [
      { src: '/icons/icon-192.png', sizes: '192x192', type: 'image/png' },
      { src: '/icons/icon-512.png', sizes: '512x512', type: 'image/png' },
    ],
  });
});

// ---- Service Worker ----
app.get('/sw.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript');
  res.send(`
const CACHE = 'viva-v1';
const ASSETS = ['/', '/index.html'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS))));
self.addEventListener('fetch', e => {
  if (e.request.url.includes('/api/')) return;
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
  `);
});

// ---- Catch-all para SPA ----
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`\n🌿 VIVA — Servidor corriendo en http://localhost:${PORT}`);
  console.log(`   Panel admin: http://localhost:${PORT}  (usuario SISTEMAS1900)\n`);
});
