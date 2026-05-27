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
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';

let pool = null;
let qadDataCache = {};

function pgSsl() {
  if (!DATABASE_URL) return false;
  if (DATABASE_URL.includes('localhost') || DATABASE_URL.includes('127.0.0.1')) return false;
  return { rejectUnauthorized: false };
}

async function initDB() {
  if (!DATABASE_URL) {
    console.log('⚠️ Sin DATABASE_URL');
    return;
  }

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
    ['SRB0707', 'SALOrb0909', 'user'],
    ['SISTEMAS1900', 'SISTEMASviva2026', 'admin'],
    ['JRZ123', 'JACOBOrz4646', 'user'],
    ['JRR234', 'JACOBOrr8989', 'user'],
    ['SRR456', 'SIMONrr0202', 'user'],
    ['LEO2026', 'LEOg1986', 'user'],
    ['ISMAEL_VENTAS', 'Vivatex2026', 'user'],
    ['ISRAEL_ACABADO', 'Vivatex2026', 'user'],
    ['MEMO_TEJIDO', 'Vivatex2026', 'user'],
    ['CARLOSM_H', 'Vivatex2026', 'user'],
    ['MARTIN_CONTA', 'Vivatex2026', 'user'],
    ['VENTAS_MONICA', 'Vivatex2026', 'user'],
    ['VENTAS_EDGAR', 'Vivatex2026', 'user'],
    ['VENTAS_AMELIA', 'Vivatex2026', 'user'],
    ['VENTAS_JORGE', 'Vivatex2026', 'user']
  ];

  for (const u of users) {
    await pool.query(`
      INSERT INTO usuarios (username,password,role,active)
      VALUES ($1,$2,$3,true)
      ON CONFLICT (username) DO NOTHING
    `, u);
  }

  console.log('✅ PostgreSQL conectado');
}

function norm(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s.-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function esc(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
  const n = Number(String(v || '').replace(/[$,]/g, '').replace(/[^\d.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function cleanPayload(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (typeof data === 'string') {
    try { return cleanPayload(JSON.parse(data)); }
    catch {
      return data.split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x }));
    }
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

async function loadQAD() {
  const cache = {};

  if (!pool) return cache;

  if (await tableExists('qad_data')) {
    const r = await pool.query('SELECT * FROM qad_data ORDER BY updated_at DESC');
    for (const row of r.rows) {
      cache[row.sheet_key || `qad_data_${row.id}`] = {
        filename: row.filename || 'QAD',
        sheet: row.sheet_name || 'QAD',
        updatedAt: row.updated_at,
        source: 'qad_data',
        data: cleanPayload(row.data)
      };
    }
  }

  if (await tableExists('qad_files')) {
    const r = await pool.query('SELECT * FROM qad_files ORDER BY created_at DESC');
    for (const row of r.rows) {
      cache[`qad_files_${row.id}`] = {
        filename: row.filename || 'QAD FILE',
        sheet: 'QAD_FILE',
        updatedAt: row.created_at || row.updated_at,
        source: 'qad_files',
        data: cleanPayload(row.data || row.content || row.text || row.file_content || '')
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
      if (row && typeof row === 'object') {
        out.push({ ...c, row });
      }
    });
  });
  return out;
}

function usefulColumns(rows) {
  if (!rows.length) return [];

  const all = new Set();
  rows.forEach(r => Object.keys(r || {}).forEach(k => all.add(k)));

  const preferred = [
    'cliente', 'nombre', 'razon', 'razón', 'codigo', 'código',
    'vendedor', 'saldo', 'vencido', 'corriente', 'credito', 'crédito',
    'limite', 'límite', 'factura', 'pedido', 'fecha', 'importe',
    'total', 'cantidad', 'producto', 'descripcion', 'descripción',
    'dias', 'días', 'antiguedad', 'antigüedad'
  ];

  const cols = [...all];

  const withData = cols.filter(col => {
    const filled = rows.slice(0, 80).filter(r => {
      const v = r[col];
      return v !== null && v !== undefined && String(v).trim() !== '';
    }).length;
    return filled > 0;
  });

  const priority = withData.filter(c => preferred.some(p => norm(c).includes(norm(p))));
  const rest = withData.filter(c => !priority.includes(c));

  return [...priority, ...rest].slice(0, 14);
}

function tableHTML(rows, max = 40) {
  if (!rows.length) return '<p><b>No hay datos disponibles.</b></p>';

  const limited = rows.slice(0, max);
  const cols = usefulColumns(limited);

  if (!cols.length) return '<p><b>No hay columnas útiles para mostrar.</b></p>';

  return `
<div style="overflow-x:auto;width:100%;margin:12px 0;">
<table style="border-collapse:collapse;width:100%;font-size:13px;">
<thead>
<tr>
${cols.map(c => `<th style="border:1px solid #aac49f;background:#eaf4e4;padding:8px;text-align:left;">${esc(c)}</th>`).join('')}
</tr>
</thead>
<tbody>
${limited.map(r => `
<tr>
${cols.map(c => `<td style="border:1px solid #d6e3cf;padding:8px;vertical-align:top;">${esc(fmt(r[c]))}</td>`).join('')}
</tr>
`).join('')}
</tbody>
</table>
</div>
${rows.length > max ? `<p><em>Mostrando ${max} de ${rows.length} registros encontrados.</em></p>` : ''}
`;
}

function searchTerms(msg) {
  const stop = new Set([
    'dame','reporte','tabla','cliente','clientes','vendedor','vendedores',
    'informacion','información','datos','qad','de','del','la','el','los',
    'las','un','una','en','que','tienes','tengo','por','para','con',
    'saldo','saldos','cartera','pagos','pago','atrasados','atrasado'
  ]);

  return norm(msg).split(' ').filter(w => w.length > 2 && !stop.has(w));
}

function scoreRow(row, terms) {
  const text = norm(JSON.stringify(row));
  let score = 0;
  terms.forEach(t => {
    if (text.includes(t)) score += 100;
    text.split(' ').forEach(w => {
      if (w.startsWith(t) || t.startsWith(w)) score += 10;
    });
  });
  return score;
}

function bestMoneyCol(rows) {
  const cols = usefulColumns(rows);
  const preferred = ['saldo', 'vencido', 'total', 'importe', 'cartera', 'monto'];
  let best = null, bestScore = -1;

  cols.forEach(c => {
    let s = 0;
    const nc = norm(c);
    preferred.forEach(p => { if (nc.includes(p)) s += 100; });
    rows.slice(0, 80).forEach(r => { if (Math.abs(num(r[c])) > 0) s += 1; });
    if (s > bestScore) { bestScore = s; best = c; }
  });

  return best;
}

async function diagnostics() {
  const d = { postgres: !!pool, qad_data: 0, qad_files: 0 };

  if (!pool) return d;

  if (await tableExists('qad_data')) {
    const r = await pool.query('SELECT COUNT(*)::int AS c FROM qad_data');
    d.qad_data = r.rows[0]?.c || 0;
  }

  if (await tableExists('qad_files')) {
    const r = await pool.query('SELECT COUNT(*)::int AS c FROM qad_files');
    d.qad_files = r.rows[0]?.c || 0;
  }

  return d;
}

function filesReply(cache) {
  const rows = Object.values(cache).map(c => ({
    Archivo: c.filename,
    Hoja: c.sheet,
    Registros: Array.isArray(c.data) ? c.data.length : 0,
    Fuente: c.source,
    Actualizado: c.updatedAt ? new Date(c.updatedAt).toLocaleString('es-MX') : ''
  }));

  return `
<h2>Registros QAD disponibles</h2>
${tableHTML(rows, 200)}
`;
}

async function qadReply(message) {
  const cache = await loadQAD();
  const items = allRows(cache);
  const msg = norm(message);

  if (!items.length) {
    const d = await diagnostics();
    return `
<h2>No encontré registros QAD cargados</h2>
<p>Diagnóstico:</p>
${tableHTML([{
      PostgreSQL: d.postgres ? 'Conectado' : 'No conectado',
      qad_data: d.qad_data,
      qad_files: d.qad_files
    }], 1)}
`;
  }

  if (
    msg.includes('que tienes') ||
    msg.includes('registros') ||
    msg.includes('archivos') ||
    msg.includes('qad cargados')
  ) {
    return filesReply(cache);
  }

  const isBusiness =
    msg.includes('cliente') || msg.includes('cartera') || msg.includes('saldo') ||
    msg.includes('vendedor') || msg.includes('venta') || msg.includes('pedido') ||
    msg.includes('inventario') || msg.includes('reporte') || msg.includes('tabla') ||
    msg.includes('pago') || msg.includes('atrasado');

  if (!isBusiness) return null;

  const terms = searchTerms(message);
  let matches = [];

  if (terms.length) {
    matches = items
      .map(i => ({ ...i, score: scoreRow(i.row, terms) }))
      .filter(i => i.score > 0)
      .sort((a, b) => b.score - a.score);
  }

  if (!matches.length) {
    matches = items.filter(i => {
      const t = norm(`${i.filename} ${i.sheet} ${JSON.stringify(i.row)}`);
      return t.includes('cliente') || t.includes('saldo') || t.includes('cartera') || t.includes('vendedor');
    }).map(i => ({ ...i, score: 1 }));
  }

  if (!matches.length) {
    return `
<h2>No encontré coincidencias exactas</h2>
<p>Pero sí hay registros QAD cargados.</p>
${filesReply(cache)}
`;
  }

  let rows = matches.map(m => m.row);

  if (
    msg.includes('top') ||
    msg.includes('principal') ||
    msg.includes('mayor') ||
    msg.includes('atrasado') ||
    msg.includes('vencido')
  ) {
    const col = bestMoneyCol(rows);
    if (col) rows = [...rows].sort((a, b) => num(b[col]) - num(a[col]));
    rows = rows.slice(0, 10);
  } else {
    rows = rows.slice(0, 60);
  }

  const sources = {};
  matches.forEach(m => {
    const k = `${m.filename} / ${m.sheet}`;
    sources[k] = (sources[k] || 0) + 1;
  });

  const sourceRows = Object.keys(sources).map(k => ({
    Fuente: k,
    Coincidencias: sources[k]
  }));

  return `
<h2>Resumen ejecutivo</h2>
<p>Encontré <b>${matches.length}</b> registros reales en QAD relacionados con tu consulta.</p>
<p>La información siguiente sale directamente de PostgreSQL/QAD. No se inventan clientes, vendedores, saldos ni fechas.</p>

<h2>Fuentes revisadas</h2>
${tableHTML(sourceRows, 20)}

<h2>Tabla principal</h2>
${tableHTML(rows, 60)}

<h2>Observaciones</h2>
<ul>
<li>Solo se muestran columnas que sí tienen información útil.</li>
<li>Los campos vacíos se ocultan para que la tabla no se llene de “NO DISPONIBLE”.</li>
<li>Si quieres todo el detalle completo, pide: <b>exporta esto a Excel</b>.</li>
</ul>
`;
}

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
    VALUES ($1,$2,$3,$4,NOW())
    ON CONFLICT (sheet_key)
    DO UPDATE SET filename=$2, sheet_name=$3, data=$4, updated_at=NOW()
  `, [key, filename, sheet, JSON.stringify(data)]);
}

app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  let sheets = 0;

  for (const file of req.files || []) {
    const ext = path.extname(file.originalname).toLowerCase();

    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      const parsed = parseExcel(file.buffer);
      for (const sheet of Object.keys(parsed)) {
        const key = `${file.originalname}_${sheet}_${Date.now()}`.replace(/[^a-zA-Z0-9_-]/g, '_');
        await saveQAD(key, file.originalname, sheet, parsed[sheet]);
        sheets++;
      }
    }

    if (ext === '.pdf') {
      const p = await pdfParse(file.buffer);
      const data = String(p.text || '').split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x }));
      const key = `${file.originalname}_PDF_${Date.now()}`.replace(/[^a-zA-Z0-9_-]/g, '_');
      await saveQAD(key, file.originalname, 'PDF', data);
      sheets++;
    }
  }

  res.json({ ok: true, sheets });
});

app.get('/api/qad/status', async (req, res) => {
  const cache = await loadQAD();
  const d = await diagnostics();
  res.json({
    hasData: allRows(cache).length > 0,
    totalRecords: allRows(cache).length,
    sheets: Object.keys(cache),
    diagnostics: d
  });
});

app.delete('/api/qad/clear', async (req, res) => {
  if (pool) await pool.query('DELETE FROM qad_data');
  res.json({ ok: true });
});

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

app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body || {};
  const lastMsg = messages?.[messages.length - 1]?.content || '';

  try {
    const direct = await qadReply(lastMsg);
    if (direct) return res.json({ reply: direct });

    if (!OPENAI_API_KEY) return res.json({ reply: 'AVIVA está conectada, pero falta OPENAI_API_KEY.' });

    const r = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${OPENAI_API_KEY}` },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        temperature: 0,
        max_tokens: 1500,
        messages: [
          { role: 'system', content: `${system || ''}\nNo inventes datos. Responde breve y profesional.` },
          ...messages.map(m => ({
            role: m.role === 'assistant' ? 'assistant' : 'user',
            content: String(m.content || '')
          }))
        ]
      })
    });

    const data = await r.json();
    res.json({ reply: data.choices?.[0]?.message?.content || 'Sin respuesta.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ reply: 'Error interno procesando QAD.' });
  }
});

app.post('/api/conversations/save', async (req, res) => res.json({ ok: true }));
app.get('/api/conversations/:username', async (req, res) => res.json({ conversations: [] }));
app.delete('/api/conversations/:username/:convId', async (req, res) => res.json({ ok: true }));

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

app.get('/api/health', async (req, res) => res.json({ ok: true, db: !!pool, model: OPENAI_MODEL, diagnostics: await diagnostics() }));

app.get('/manifest.json', (req, res) => res.json({ name: 'AVIVA', short_name: 'AVIVA', start_url: '/', display: 'standalone' }));

app.get('*', (req, res) => {
  const index = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(index)) return res.sendFile(index);
  res.send('AVIVA ONLINE');
});

initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE PORT ${PORT}`);
    console.log(`🧠 Model: ${OPENAI_MODEL}`);
  });
});
