const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const { Pool } = require('pg');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
const DATABASE_URL = process.env.DATABASE_URL || '';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-20250514';

let pool = null;

function pgSsl() {
  if (!DATABASE_URL) return false;
  if (DATABASE_URL.includes('localhost') || DATABASE_URL.includes('127.0.0.1')) return false;
  return { rejectUnauthorized: false };
}

async function initDB() {
  if (!DATABASE_URL) { console.log('⚠️ Sin DATABASE_URL'); return; }
  pool = new Pool({ connectionString: DATABASE_URL, ssl: pgSsl() });

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

  await pool.query(`
    CREATE TABLE IF NOT EXISTS usuarios (
      id SERIAL PRIMARY KEY,
      username VARCHAR(100) UNIQUE NOT NULL,
      password VARCHAR(255) NOT NULL,
      role VARCHAR(50) DEFAULT 'user',
      active BOOLEAN DEFAULT true,
      created_at TIMESTAMP DEFAULT NOW()
    );
  `);

  const users = [
    ['SRB0707',        'SALOrb0909',        'user'],
    ['SISTEMAS1900',   'SISTEMASviva2026',  'admin'],
    ['JRZ123',         'JACOBOrz4646',      'user'],
    ['JRR234',         'JACOBOrr8989',      'user'],
    ['SRR456',         'SIMONrr0202',       'user'],
    ['LEO2026',        'LEOg1986',          'user'],
    ['ISMAEL_VENTAS',  'Vivatex2026',       'user'],
    ['ISRAEL_ACABADO', 'Vivatex2026',       'user'],
    ['MEMO_TEJIDO',    'Vivatex2026',       'user'],
    ['CARLOSM_H',      'Vivatex2026',       'user'],
    ['MARTIN_CONTA',   'Vivatex2026',       'user'],
    ['VENTAS_MONICA',  'Vivatex2026',       'user'],
    ['VENTAS_EDGAR',   'Vivatex2026',       'user'],
    ['VENTAS_AMELIA',  'Vivatex2026',       'user'],
    ['VENTAS_JORGE',   'Vivatex2026',       'user'],
  ];

  for (const u of users) {
    await pool.query(`
      INSERT INTO usuarios (username, password, role, active)
      VALUES ($1, $2, $3, true)
      ON CONFLICT (username) DO NOTHING
    `, u);
  }

  console.log('✅ PostgreSQL conectado');
}

// ─── UTILIDADES ────────────────────────────────────────────────────────────────

function norm(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s.\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function fmt(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') {
    return Number.isInteger(v)
      ? v.toLocaleString('es-MX')
      : v.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return String(v).replace(/\n/g, ' ').trim();
}

function num(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/[$,]/g, '').replace(/[^\d.\-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function cleanPayload(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (typeof data === 'string') {
    try { return cleanPayload(JSON.parse(data)); }
    catch { return data.split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x })); }
  }
  if (typeof data === 'object') {
    if (Array.isArray(data.data)) return data.data;
    if (Array.isArray(data.rows)) return data.rows;
    if (Array.isArray(data.records)) return data.records;
    return [data];
  }
  return [];
}

async function tableExists(name) {
  if (!pool) return false;
  const r = await pool.query(`
    SELECT EXISTS (
      SELECT FROM information_schema.tables
      WHERE table_schema='public' AND table_name=$1
    ) AS exists
  `, [name]);
  return !!r.rows[0]?.exists;
}

// ─── CARGA QAD ─────────────────────────────────────────────────────────────────

async function loadQAD() {
  const cache = {};
  if (!pool) return cache;

  if (await tableExists('qad_data')) {
    const r = await pool.query('SELECT * FROM qad_data ORDER BY updated_at DESC');
    for (const row of r.rows) {
      cache[row.sheet_key || `qad_${row.id}`] = {
        filename: row.filename || 'QAD',
        sheet: row.sheet_name || 'QAD',
        updatedAt: row.updated_at,
        source: 'qad_data',
        data: cleanPayload(row.data)
      };
    }
  }

  return cache;
}

function allRows(cache) {
  const out = [];
  Object.keys(cache || {}).forEach(k => {
    const c = cache[k];
    if (!Array.isArray(c.data)) return;
    c.data.forEach(row => {
      if (row && typeof row === 'object') out.push({ ...c, row });
    });
  });
  return out;
}

// ─── CONSTRUIR CONTEXTO QAD PARA CLAUDE ────────────────────────────────────────

async function buildQADContext() {
  const cache = await loadQAD();
  const items = allRows(cache);

  if (!items.length) {
    return {
      hasData: false,
      text: 'No hay datos QAD cargados en este momento. El administrador debe subir archivos Excel/CSV desde el Panel de Sistemas.'
    };
  }

  // Resumir qué archivos hay
  const sources = {};
  items.forEach(i => {
    const k = `${i.filename} / ${i.sheet}`;
    if (!sources[k]) sources[k] = { count: 0, sample: null, updatedAt: i.updatedAt };
    sources[k].count++;
    if (!sources[k].sample) sources[k].sample = i.row;
  });

  let contextText = `DATOS REALES DE QAD DISPONIBLES (${items.length} registros totales):\n\n`;

  // Para cada fuente, incluir TODOS los registros
  Object.keys(sources).forEach(src => {
    const s = sources[src];
    const fecha = s.updatedAt ? new Date(s.updatedAt).toLocaleString('es-MX') : 'N/A';
    contextText += `=== FUENTE: ${src} | ${s.count} registros | Actualizado: ${fecha} ===\n`;

    // Obtener todos los registros de esta fuente
    const srcRows = items.filter(i => `${i.filename} / ${i.sheet}` === src).map(i => i.row);

    // Incluir encabezados
    if (srcRows.length > 0) {
      const cols = Object.keys(srcRows[0]);
      contextText += cols.join(' | ') + '\n';
      contextText += cols.map(() => '---').join(' | ') + '\n';

      // Incluir TODOS los registros (máx 500 para no explotar el contexto)
      srcRows.slice(0, 500).forEach(row => {
        contextText += cols.map(c => fmt(row[c])).join(' | ') + '\n';
      });

      if (srcRows.length > 500) {
        contextText += `... y ${srcRows.length - 500} registros más.\n`;
      }
    }
    contextText += '\n';
  });

  return { hasData: true, text: contextText, totalRecords: items.length };
}

// ─── DIAGNÓSTICOS ──────────────────────────────────────────────────────────────

async function diagnostics() {
  const d = { postgres: !!pool, qad_data: 0 };
  if (!pool) return d;
  if (await tableExists('qad_data')) {
    const r = await pool.query('SELECT COUNT(*)::int AS c FROM qad_data');
    d.qad_data = r.rows[0]?.c || 0;
  }
  return d;
}

// ─── UPLOAD ────────────────────────────────────────────────────────────────────

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

function parseExcel(buffer) {
  const wb = xlsx.read(buffer, { type: 'buffer' });
  const out = {};
  wb.SheetNames.forEach(s => {
    out[s] = xlsx.utils.sheet_to_json(wb.Sheets[s], { defval: '' });
  });
  return out;
}

async function saveQAD(key, filename, sheet, data) {
  if (!pool) return;
  await pool.query(`
    INSERT INTO qad_data (sheet_key, filename, sheet_name, data, updated_at)
    VALUES ($1, $2, $3, $4, NOW())
    ON CONFLICT (sheet_key)
    DO UPDATE SET filename=$2, sheet_name=$3, data=$4, updated_at=NOW()
  `, [key, filename, sheet, JSON.stringify(data)]);
}

// ─── RUTAS QAD ─────────────────────────────────────────────────────────────────

app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  let sheets = 0;
  let files = 0;

  for (const file of req.files || []) {
    const ext = path.extname(file.originalname).toLowerCase();

    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      const parsed = parseExcel(file.buffer);
      for (const sheet of Object.keys(parsed)) {
        const key = `${file.originalname}_${sheet}`.replace(/[^a-zA-Z0-9_\-]/g, '_');
        await saveQAD(key, file.originalname, sheet, parsed[sheet]);
        sheets++;
      }
      files++;
    }

    if (ext === '.pdf') {
      try {
        const p = await pdfParse(file.buffer);
        const data = String(p.text || '').split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x }));
        const key = `${file.originalname}_PDF`.replace(/[^a-zA-Z0-9_\-]/g, '_');
        await saveQAD(key, file.originalname, 'PDF', data);
        sheets++;
        files++;
      } catch(e) {
        console.error('Error parsing PDF:', e.message);
      }
    }
  }

  res.json({ ok: true, files, sheets });
});

app.get('/api/qad/status', async (req, res) => {
  const cache = await loadQAD();
  const d = await diagnostics();
  const rows = allRows(cache);
  res.json({
    hasData: rows.length > 0,
    totalRecords: rows.length,
    sheets: Object.keys(cache),
    diagnostics: d
  });
});

app.delete('/api/qad/clear', async (req, res) => {
  if (pool) await pool.query('DELETE FROM qad_data');
  res.json({ ok: true });
});

// ─── RUTAS USUARIOS ────────────────────────────────────────────────────────────

app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!pool) return res.status(500).json({ ok: false, error: 'Sin DB' });

  const r = await pool.query(
    'SELECT id, username, role, active FROM usuarios WHERE UPPER(username)=$1 AND password=$2',
    [String(username || '').toUpperCase(), password]
  );

  if (!r.rows.length) return res.json({ ok: false, error: 'Usuario o contraseña incorrectos' });
  if (!r.rows[0].active) return res.json({ ok: false, error: 'Usuario inactivo' });

  res.json({ ok: true, user: r.rows[0] });
});

app.get('/api/usuarios', async (req, res) => {
  if (!pool) return res.json({ users: [] });
  try {
    const r = await pool.query('SELECT id, username, role, active FROM usuarios ORDER BY id');
    res.json({ users: r.rows });
  } catch(e) {
    res.json({ users: [] });
  }
});

app.post('/api/usuarios', async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.json({ ok: false, error: 'Faltan datos' });
  if (!pool) return res.json({ ok: false, error: 'Sin DB' });
  try {
    await pool.query(`
      INSERT INTO usuarios (username, password, role, active)
      VALUES ($1, $2, $3, true)
      ON CONFLICT (username) DO UPDATE SET password=$2, role=$3, active=true
    `, [username.toUpperCase(), password, role || 'user']);
    res.json({ ok: true });
  } catch(e) {
    res.json({ ok: false, error: e.message });
  }
});

app.delete('/api/usuarios/:username', async (req, res) => {
  if (!pool) return res.json({ ok: false });
  try {
    await pool.query('UPDATE usuarios SET active=false WHERE UPPER(username)=$1', [req.params.username.toUpperCase()]);
    res.json({ ok: true });
  } catch(e) {
    res.json({ ok: false, error: e.message });
  }
});

// ─── RUTA CHAT — CLAUDE ────────────────────────────────────────────────────────

app.post('/api/chat', async (req, res) => {
  const { messages, system, username } = req.body || {};

  if (!ANTHROPIC_API_KEY) {
    return res.json({ reply: '⚠️ Falta configurar ANTHROPIC_API_KEY en las variables de entorno.' });
  }

  try {
    // Cargar datos QAD reales
    const qadCtx = await buildQADContext();

    const systemPrompt = `Eres AVIVA, la analista de inteligencia artificial de Grupo Vivatex S.A. de C.V.

IDENTIDAD:
- Tu nombre es AVIVA — Inteligencia Artificial de Vivatex
- Nunca menciones Claude, Anthropic ni OpenAI
- Trabajas exclusivamente para Grupo Vivatex S.A. de C.V.
- Hablas español mexicano natural, tuteas al usuario, eres directa y profesional

REGLAS ABSOLUTAS — NUNCA LAS ROMPAS:
1. NUNCA inventes clientes, vendedores, productos, cantidades ni fechas
2. SOLO usa datos que aparezcan literalmente en el contexto QAD que recibes
3. Si el dato no está en QAD, di exactamente: "Ese dato no está disponible en la información que tengo. Verifica directamente en QAD."
4. NUNCA hagas gráficas ASCII con █ ▓ ░ ni bloques de texto tipo barra
5. NUNCA muestres JSON crudo ni arrays en tu respuesta
6. SIEMPRE muestra el listado COMPLETO — nunca omitas registros sin avisar
7. Primero el detalle completo, luego el resumen

FORMATO DE RESPUESTAS:
Usa SIEMPRE formato Markdown limpio:
- **Negritas** para cifras y nombres importantes
- ## Encabezados para organizar secciones
- Tablas Markdown para datos tabulares, así:

| Cliente | Saldo MXN | Días |
|---------|-----------|------|
| EMPRESA SA DE CV | $1,234,567 | 90 |

- Ordena de mayor a menor cuando sea relevante (saldos, ventas, etc.)
- Incluye totales al final de las tablas cuando aplique
- Respuestas completas y ordenadas, sin cortar información

SOBRE LOS DATOS QAD:
${qadCtx.hasData
  ? `Tienes ${qadCtx.totalRecords} registros reales cargados. ÚSALOS DIRECTAMENTE.\n\n${qadCtx.text}`
  : qadCtx.text
}

${system ? `\nCONTEXTO ADICIONAL:\n${system}` : ''}

Usuario actual: ${username || 'Usuario'}
Fecha y hora: ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}`;

    // Preparar mensajes — solo roles user/assistant para Claude
    const cleanMessages = (messages || [])
      .filter(m => m.role === 'user' || m.role === 'assistant')
      .map(m => ({
        role: m.role,
        content: String(m.content || '').substring(0, 4000) // limitar por mensaje
      }));

    // Asegurar que empiece con user
    if (!cleanMessages.length || cleanMessages[0].role !== 'user') {
      return res.json({ reply: 'Por favor escribe tu pregunta.' });
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: ANTHROPIC_MODEL,
        max_tokens: 4096,
        system: systemPrompt,
        messages: cleanMessages
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error('Anthropic API error:', response.status, errText);
      return res.json({ reply: `Error al conectar con AVIVA (${response.status}). Intenta de nuevo.` });
    }

    const data = await response.json();
    const reply = data.content?.[0]?.text || 'Sin respuesta.';

    res.json({ reply });

  } catch (e) {
    console.error('Error en /api/chat:', e);
    res.status(500).json({ reply: 'Error interno procesando tu consulta. Intenta de nuevo.' });
  }
});

// ─── CONVERSACIONES ────────────────────────────────────────────────────────────

app.post('/api/conversations/save', async (req, res) => {
  const { username, convId, title, messages } = req.body || {};
  if (!pool || !username || !convId) return res.json({ ok: true });
  try {
    await pool.query(`
      INSERT INTO conversations (username, conv_id, title, messages, updated_at)
      VALUES ($1, $2, $3, $4, NOW())
      ON CONFLICT (username, conv_id)
      DO UPDATE SET title=$3, messages=$4, updated_at=NOW()
    `, [username, convId, title || 'Conversación', JSON.stringify(messages || [])]);
    res.json({ ok: true });
  } catch(e) {
    res.json({ ok: true });
  }
});

app.get('/api/conversations/:username', async (req, res) => {
  if (!pool) return res.json({ conversations: [] });
  try {
    const r = await pool.query(
      'SELECT conv_id, title, messages, updated_at FROM conversations WHERE username=$1 ORDER BY updated_at DESC LIMIT 50',
      [req.params.username]
    );
    res.json({ conversations: r.rows });
  } catch(e) {
    res.json({ conversations: [] });
  }
});

app.delete('/api/conversations/:username/:convId', async (req, res) => {
  if (!pool) return res.json({ ok: true });
  try {
    await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [req.params.username, req.params.convId]);
    res.json({ ok: true });
  } catch(e) {
    res.json({ ok: true });
  }
});

// ─── EXCEL ─────────────────────────────────────────────────────────────────────

app.post('/api/excel/generate', async (req, res) => {
  const wb = xlsx.utils.book_new();
  const { titulo, hojas } = req.body || {};

  for (const h of hojas || []) {
    const ws = xlsx.utils.aoa_to_sheet([h.columnas || [], ...(h.filas || [])]);
    xlsx.utils.book_append_sheet(wb, ws, String(h.nombre || 'Datos').slice(0, 31));
  }

  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.json({
    ok: true,
    base64: buf.toString('base64'),
    filename: `${String(titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}.xlsx`
  });
});

// ─── HEALTH ────────────────────────────────────────────────────────────────────

app.get('/api/health', async (req, res) => {
  const d = await diagnostics();
  res.json({
    ok: true,
    db: !!pool,
    model: ANTHROPIC_MODEL,
    provider: 'anthropic',
    diagnostics: d
  });
});

app.get('/manifest.json', (req, res) => res.json({
  name: 'AVIVA',
  short_name: 'AVIVA',
  start_url: '/',
  display: 'standalone',
  theme_color: '#1a1f16',
  background_color: '#f0f4ec'
}));

app.get('*', (req, res) => {
  const index = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(index)) return res.sendFile(index);
  res.send('AVIVA ONLINE — Grupo Vivatex S.A. de C.V.');
});

// ─── INICIO ────────────────────────────────────────────────────────────────────

initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE — Puerto ${PORT}`);
    console.log(`🤖 Modelo: ${ANTHROPIC_MODEL}`);
    console.log(`🔑 API Key: ${ANTHROPIC_API_KEY ? 'Configurada ✅' : 'FALTA ❌'}`);
  });
});
