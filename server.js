const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const xlsx = require('xlsx');
const { Pool } = require('pg');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
const PORT = process.env.PORT || 3000;
const DATABASE_URL = process.env.DATABASE_URL || '';
// ============================================================

// Base de datos PostgreSQL
let pool = null;

async function initDB() {
  if (!DATABASE_URL) {
    console.log('⚠️  Sin DATABASE_URL — usando memoria (datos no persisten)');
    return;
  }
  try {
    pool = new Pool({ connectionString: DATABASE_URL, ssl: { rejectUnauthorized: false } });
    await pool.query(`
      CREATE TABLE IF NOT EXISTS qad_data (
        id SERIAL PRIMARY KEY,
        sheet_key VARCHAR(255) UNIQUE NOT NULL,
        filename VARCHAR(500),
        sheet_name VARCHAR(255),
        data JSONB,
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `);
    console.log('✅ Base de datos PostgreSQL conectada');
  } catch (e) {
    console.error('❌ Error conectando DB:', e.message);
    pool = null;
  }
}

// Guardar datos QAD en DB
async function saveQADToDB(key, filename, sheetName, data) {
  if (!pool) return;
  try {
    await pool.query(`
      INSERT INTO qad_data (sheet_key, filename, sheet_name, data, updated_at)
      VALUES ($1, $2, $3, $4, NOW())
      ON CONFLICT (sheet_key) DO UPDATE
      SET filename=$2, sheet_name=$3, data=$4, updated_at=NOW()
    `, [key, filename, sheetName, JSON.stringify(data)]);
  } catch (e) {
    console.error('Error guardando en DB:', e.message);
  }
}

// Cargar datos QAD desde DB
async function loadQADFromDB() {
  if (!pool) return {};
  try {
    const result = await pool.query('SELECT * FROM qad_data ORDER BY updated_at DESC');
    const cache = {};
    result.rows.forEach(row => {
      cache[row.sheet_key] = {
        data: row.data,
        filename: row.filename,
        sheet: row.sheet_name,
        updatedAt: row.updated_at
      };
    });
    return cache;
  } catch (e) {
    console.error('Error cargando DB:', e.message);
    return {};
  }
}

// Limpiar datos QAD de DB
async function clearQADFromDB() {
  if (!pool) return;
  try {
    await pool.query('DELETE FROM qad_data');
  } catch (e) {
    console.error('Error limpiando DB:', e.message);
  }
}

// Cache en memoria (respaldo)
let qadDataCache = {};
let qadLastUpdate = null;

// Multer para subida de archivos
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, allowed.includes(ext));
  }
});

// Parsear Excel/CSV
function parseFile(buffer, filename) {
  try {
    const ext = path.extname(filename).toLowerCase();
    const workbook = xlsx.read(buffer, { type: 'buffer' });
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

// ---- Subir archivos QAD ----
app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: 'No se recibieron archivos' });
  }
  let totalSheets = 0;
  for (const file of req.files) {
    const parsed = parseFile(file.buffer, file.originalname);
    if (parsed) {
      const baseName = path.basename(file.originalname, path.extname(file.originalname));
      for (const sheet of Object.keys(parsed)) {
        const key = `${baseName}_${sheet}`.replace(/\s+/g, '_').toLowerCase();
        const entry = {
          data: parsed[sheet],
          filename: file.originalname,
          sheet,
          updatedAt: new Date().toISOString()
        };
        qadDataCache[key] = entry;
        await saveQADToDB(key, file.originalname, sheet, parsed[sheet]);
        totalSheets++;
      }
    }
  }
  qadLastUpdate = new Date().toISOString();
  console.log(`📊 QAD: ${req.files.length} archivos, ${totalSheets} hojas`);
  res.json({ ok: true, files: req.files.length, sheets: totalSheets, updatedAt: qadLastUpdate });
});

// ---- Estado QAD ----
app.get('/api/qad/status', async (req, res) => {
  const cache = pool ? await loadQADFromDB() : qadDataCache;
  const keys = Object.keys(cache);
  res.json({
    hasData: keys.length > 0,
    sheets: keys,
    lastUpdate: qadLastUpdate,
    totalRecords: keys.reduce((sum, k) => sum + (cache[k]?.data?.length || 0), 0),
    persistent: !!pool
  });
});

// ---- Limpiar QAD ----
app.delete('/api/qad/clear', async (req, res) => {
  qadDataCache = {};
  qadLastUpdate = null;
  await clearQADFromDB();
  res.json({ ok: true });
});

// ---- Chat principal ----
app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body;
  if (!messages || !system) return res.status(400).json({ error: 'Faltan datos' });

  try {
    // Cargar datos QAD (desde DB si disponible, sino memoria)
    const cache = pool ? await loadQADFromDB() : qadDataCache;
    const keys = Object.keys(cache);
    let qadContext = '';

    if (keys.length > 0) {
      const lastMsg = messages[messages.length - 1]?.content?.toLowerCase() || '';
      const relevantData = [];

      keys.forEach(key => {
        const c = cache[key];
        const isRelevant =
          (lastMsg.includes('venta') && (key.includes('venta') || key.includes('sale'))) ||
          (lastMsg.includes('inventario') && (key.includes('inventario') || key.includes('stock'))) ||
          (lastMsg.includes('cliente') && (key.includes('cliente') || key.includes('customer'))) ||
          (lastMsg.includes('compra') && (key.includes('compra') || key.includes('purchase'))) ||
          (lastMsg.includes('proveedor') && (key.includes('proveedor') || key.includes('supplier'))) ||
          (lastMsg.includes('financiero') || lastMsg.includes('finanza') || lastMsg.includes('balance')) ||
          (lastMsg.includes('orden') && (key.includes('orden') || key.includes('order'))) ||
          lastMsg.includes('todo') || lastMsg.includes('resumen') || lastMsg.includes('reporte') ||
          lastMsg.includes('excel') || lastMsg.includes('tabla');

        if (isRelevant) {
          const sample = Array.isArray(c.data) ? c.data.slice(0, 100) : c.data;
          relevantData.push(`\n### ${c.filename} — Hoja: ${c.sheet} (${Array.isArray(c.data) ? c.data.length : '?'} registros)\n${JSON.stringify(sample, null, 2)}`);
        }
      });

      if (relevantData.length > 0) {
        qadContext = `\n\nDATOS QAD ACTUALES (${new Date(qadLastUpdate || Date.now()).toLocaleString('es-MX')}):\n${relevantData.join('\n')}`;
      } else {
        qadContext = `\n\nDatos disponibles en QAD:\n${keys.map(k => `- ${cache[k].filename} / ${cache[k].sheet}: ${Array.isArray(cache[k].data) ? cache[k].data.length : '?'} registros`).join('\n')}`;
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
    start_url: '/',
    display: 'standalone',
    background_color: '#1a1f16',
    theme_color: '#6BBF3E',
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
const CACHE = 'viva-v2';
const ASSETS = ['/', '/index.html'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS))));
self.addEventListener('fetch', e => {
  if (e.request.url.includes('/api/')) return;
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
  `);
});

// ---- Catch-all ----
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ---- Arrancar ----
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`\n🌿 VIVA — Servidor corriendo en http://localhost:${PORT}`);
    console.log(`   Panel admin: http://localhost:${PORT}  (usuario SISTEMAS1900)`);
    console.log(`   Base de datos: ${pool ? '✅ PostgreSQL persistente' : '⚠️  Memoria (no persiste)'}\n`);
  });
});
