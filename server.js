const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || 'TU_API_KEY_AQUI';
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'qad_data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
// ============================================================

// Almacenamiento de datos QAD en memoria
let qadDataCache = {};
let qadLastUpdate = null;

// ---- Multer para subida de archivos ----
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, allowed.includes(ext));
  }
});

// ---- Función para parsear Excel/CSV ----
function parseFile(buffer, filename) {
  try {
    const ext = path.extname(filename).toLowerCase();
    let workbook;
    if (ext === '.csv') {
      workbook = xlsx.read(buffer, { type: 'buffer', raw: false });
    } else {
      workbook = xlsx.read(buffer, { type: 'buffer' });
    }
    const result = {};
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      result[sheetName] = xlsx.utils.sheet_to_json(sheet, { defval: '' });
    });
    return result;
  } catch (e) {
    console.error('Error parseando archivo:', e);
    return null;
  }
}

// ---- Ruta para recibir archivos de QAD (Opción A) ----
app.post('/api/qad/upload', upload.array('files', 20), (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: 'No se recibieron archivos' });
  }

  let totalSheets = 0;
  req.files.forEach(file => {
    const parsed = parseFile(file.buffer, file.originalname);
    if (parsed) {
      const baseName = path.basename(file.originalname, path.extname(file.originalname));
      Object.keys(parsed).forEach(sheet => {
        const key = `${baseName}_${sheet}`.replace(/\s+/g, '_').toLowerCase();
        qadDataCache[key] = {
          data: parsed[sheet],
          filename: file.originalname,
          sheet,
          updatedAt: new Date().toISOString()
        };
        totalSheets++;
      });
    }
  });

  qadLastUpdate = new Date().toISOString();
  console.log(`📊 QAD: ${req.files.length} archivos recibidos, ${totalSheets} hojas procesadas`);
  res.json({ ok: true, files: req.files.length, sheets: totalSheets, updatedAt: qadLastUpdate });
});

// ---- Estado de datos QAD ----
app.get('/api/qad/status', (req, res) => {
  const keys = Object.keys(qadDataCache);
  res.json({
    hasData: keys.length > 0,
    sheets: keys,
    lastUpdate: qadLastUpdate,
    totalRecords: keys.reduce((sum, k) => sum + (qadDataCache[k]?.data?.length || 0), 0)
  });
});

// ---- Limpiar datos QAD ----
app.delete('/api/qad/clear', (req, res) => {
  qadDataCache = {};
  qadLastUpdate = null;
  res.json({ ok: true });
});

// ---- Ruta principal chat ----
app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body;
  if (!messages || !system) return res.status(400).json({ error: 'Faltan datos' });

  try {
    // Preparar contexto de datos QAD
    let qadContext = '';
    const keys = Object.keys(qadDataCache);
    if (keys.length > 0) {
      const lastMsg = messages[messages.length - 1]?.content?.toLowerCase() || '';
      
      // Buscar datos relevantes según la pregunta
      const relevantData = [];
      keys.forEach(key => {
        const cache = qadDataCache[key];
        const isRelevant =
          lastMsg.includes('venta') && (key.includes('venta') || key.includes('sale')) ||
          lastMsg.includes('inventario') && (key.includes('inventario') || key.includes('stock') || key.includes('inventory')) ||
          lastMsg.includes('cliente') && (key.includes('cliente') || key.includes('customer')) ||
          lastMsg.includes('compra') && (key.includes('compra') || key.includes('purchase')) ||
          lastMsg.includes('proveedor') && (key.includes('proveedor') || key.includes('supplier')) ||
          lastMsg.includes('financiero') && (key.includes('financiero') || key.includes('finance') || key.includes('balance')) ||
          lastMsg.includes('orden') && (key.includes('orden') || key.includes('order')) ||
          lastMsg.includes('pedido') && (key.includes('pedido') || key.includes('order')) ||
          lastMsg.includes('stock') && (key.includes('stock') || key.includes('inventario'));

        if (isRelevant || lastMsg.includes('todo') || lastMsg.includes('resumen') || lastMsg.includes('reporte')) {
          const sample = cache.data.slice(0, 50);
          relevantData.push(`\n### ${cache.filename} - Hoja: ${cache.sheet} (${cache.data.length} registros, actualizado: ${cache.updatedAt})\n${JSON.stringify(sample, null, 2)}`);
        }
      });

      if (relevantData.length > 0) {
        qadContext = `\n\nDATOS ACTUALES DE QAD (última actualización: ${qadLastUpdate}):\n${relevantData.join('\n')}`;
      } else {
        // Si no hay datos relevantes específicos, dar resumen general
        const summary = keys.map(k => `- ${qadDataCache[k].filename} / ${qadDataCache[k].sheet}: ${qadDataCache[k].data.length} registros`).join('\n');
        qadContext = `\n\nDatos disponibles en QAD (última actualización: ${qadLastUpdate}):\n${summary}\n\nPide al usuario que sea más específico sobre qué datos necesita.`;
      }
    }

    const fullSystem = system + qadContext;

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
        messages,
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
  console.log(`   Panel admin: http://localhost:${PORT}  (usuario SISTEMAS1900)`);
  console.log(`   Endpoint QAD upload: POST /api/qad/upload\n`);
});
