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
    await pool.query(`
      CREATE TABLE IF NOT EXISTS conversations (
        id SERIAL PRIMARY KEY,
        username VARCHAR(100) NOT NULL,
        conv_id VARCHAR(100) NOT NULL,
        title VARCHAR(500),
        messages JSONB DEFAULT '[]',
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(username, conv_id)
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

// ---- Rutas de conversaciones ----

// Guardar conversación
app.post('/api/conversations/save', async (req, res) => {
  const { username, convId, title, messages } = req.body;
  if (!username || !convId) return res.status(400).json({ error: 'Faltan datos' });
  try {
    if (pool) {
      await pool.query(`
        INSERT INTO conversations (username, conv_id, title, messages, updated_at)
        VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (username, conv_id) DO UPDATE
        SET title=$3, messages=$4, updated_at=NOW()
      `, [username, convId, title || 'Conversación', JSON.stringify(messages)]);
    }
    res.json({ ok: true });
  } catch (e) {
    console.error('Error guardando conversación:', e.message);
    res.status(500).json({ error: 'Error guardando' });
  }
});

// Obtener todas las conversaciones de un usuario
app.get('/api/conversations/:username', async (req, res) => {
  const { username } = req.params;
  try {
    if (!pool) return res.json({ conversations: [] });
    const result = await pool.query(
      'SELECT conv_id, title, messages, created_at, updated_at FROM conversations WHERE username=$1 ORDER BY updated_at DESC LIMIT 50',
      [username]
    );
    res.json({ conversations: result.rows });
  } catch (e) {
    console.error('Error cargando conversaciones:', e.message);
    res.json({ conversations: [] });
  }
});

// Eliminar conversación
app.delete('/api/conversations/:username/:convId', async (req, res) => {
  const { username, convId } = req.params;
  try {
    if (pool) await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [username, convId]);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: 'Error eliminando' });
  }
});

// ---- Generar Excel Profesional ----
app.post('/api/excel/generate', async (req, res) => {
  const { titulo, usuario, hojas } = req.body;
  if (!hojas || !hojas.length) return res.status(400).json({ error: 'Sin datos' });

  try {
    const { spawn } = require('child_process');
    const input = JSON.stringify({ titulo, usuario, hojas });
    
    const py = spawn('python3', [path.join(__dirname, 'generate_excel.py')]);
    let output = '';
    let errOut = '';
    
    py.stdout.on('data', d => output += d.toString());
    py.stderr.on('data', d => errOut += d.toString());
    py.stdin.write(input);
    py.stdin.end();
    
    py.on('close', (code) => {
      if (code !== 0) {
        console.error('Error generando Excel:', errOut);
        return res.status(500).json({ error: 'Error generando Excel' });
      }
      const base64 = output.trim();
      res.json({ ok: true, base64, filename: `Vivatex_${(titulo||'Reporte').replace(/\s+/g,'_')}_${new Date().toISOString().split('T')[0]}.xlsx` });
    });
  } catch(e) {
    console.error('Error Excel:', e);
    res.status(500).json({ error: e.message });
  }
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
  res.setHeader('Content-Type', 'application/manifest+json');
  res.json({
    name: 'AVIVA — Vivatex IA',
    short_name: 'AVIVA',
    description: 'Asistente Inteligente de Grupo Vivatex S.A. de C.V.',
    start_url: '/',
    scope: '/',
    display: 'standalone',
    orientation: 'portrait-primary',
    background_color: '#1E6B2E',
    theme_color: '#1E6B2E',
    lang: 'es-MX',
    categories: ['business', 'productivity'],
    icons: [
      { src: '/icons/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any maskable' },
      { src: '/icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any maskable' },
    ],
  });
});

// Servir íconos embebidos
app.get('/icons/icon-192.png', (req, res) => {
  res.setHeader('Content-Type', 'image/png');
  res.send(Buffer.from('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', 'base64'));
});
app.get('/icons/icon-512.png', (req, res) => {
  res.setHeader('Content-Type', 'image/png');
  res.send(Buffer.from('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', 'base64'));
});

// ---- Service Worker ----
app.get('/sw.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript');
  res.send(`
const CACHE = 'aviva-v3';
const ASSETS = ['/', '/index.html', '/manifest.json', '/icons/icon-192.png', '/icons/icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
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
    console.log(`\n🌿 AAVIVA — Servidor corriendo en http://localhost:${PORT}`);
    console.log(`   Panel admin: http://localhost:${PORT}  (usuario SISTEMAS1900)`);
    console.log(`   Base de datos: ${pool ? '✅ PostgreSQL persistente' : '⚠️  Memoria (no persiste)'}\n`);
  });
});
