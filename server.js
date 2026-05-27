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
const ANTHROPIC_MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-5';

let pool = null;

function pgSsl() {
  if (!DATABASE_URL) return false;
  if (DATABASE_URL.includes('localhost') || DATABASE_URL.includes('127.0.0.1')) return false;
  return { rejectUnauthorized: false };
}

async function initDB() {
  if (!DATABASE_URL) { console.log('⚠️ Sin DATABASE_URL'); return; }
  pool = new Pool({ connectionString: DATABASE_URL, ssl: pgSsl() });

  await pool.query(`CREATE TABLE IF NOT EXISTS qad_data (
    id SERIAL PRIMARY KEY, sheet_key VARCHAR(255) UNIQUE NOT NULL,
    filename VARCHAR(500), sheet_name VARCHAR(255), data JSONB,
    updated_at TIMESTAMP DEFAULT NOW()
  )`);

  await pool.query(`CREATE TABLE IF NOT EXISTS conversations (
    id SERIAL PRIMARY KEY, username VARCHAR(100) NOT NULL,
    conv_id VARCHAR(100) NOT NULL, title VARCHAR(500),
    messages JSONB DEFAULT '[]', summary TEXT DEFAULT '',
    created_at TIMESTAMP DEFAULT NOW(), updated_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(username, conv_id)
  )`);

  await pool.query(`CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL, role VARCHAR(50) DEFAULT 'user',
    active BOOLEAN DEFAULT true, created_at TIMESTAMP DEFAULT NOW()
  )`);

  await pool.query(`CREATE TABLE IF NOT EXISTS user_memory (
    id SERIAL PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL,
    memory TEXT DEFAULT '', updated_at TIMESTAMP DEFAULT NOW()
  )`);

  const users = [
    ['SRB0707','SALOrb0909','user'],['SISTEMAS1900','SISTEMASviva2026','admin'],
    ['JRZ123','JACOBOrz4646','user'],['JRR234','JACOBOrr8989','user'],
    ['SRR456','SIMONrr0202','user'],['LEO2026','LEOg1986','user'],
    ['ISMAEL_VENTAS','Vivatex2026','user'],['ISRAEL_ACABADO','Vivatex2026','user'],
    ['MEMO_TEJIDO','Vivatex2026','user'],['CARLOSM_H','Vivatex2026','user'],
    ['MARTIN_CONTA','Vivatex2026','user'],['VENTAS_MONICA','Vivatex2026','user'],
    ['VENTAS_EDGAR','Vivatex2026','user'],['VENTAS_AMELIA','Vivatex2026','user'],
    ['VENTAS_JORGE','Vivatex2026','user'],
  ];
  for (const u of users) {
    await pool.query(`INSERT INTO usuarios (username,password,role,active) VALUES ($1,$2,$3,true) ON CONFLICT (username) DO NOTHING`, u);
  }
  console.log('✅ PostgreSQL conectado');
}

// ─── UTILIDADES ───────────────────────────────────────────────────
function fmt(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return v.toLocaleString('es-MX', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  return String(v).replace(/\n/g, ' ').trim();
}

function norm(v) {
  return String(v || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ').trim();
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
    return [data];
  }
  return [];
}

async function tableExists(name) {
  if (!pool) return false;
  const r = await pool.query(`SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_schema='public' AND table_name=$1) AS exists`, [name]);
  return !!r.rows[0]?.exists;
}

// ─── MEMORIA DE USUARIO ───────────────────────────────────────────
async function getUserMemory(username) {
  if (!pool) return '';
  try {
    const r = await pool.query('SELECT memory FROM user_memory WHERE username=$1', [username]);
    return r.rows[0]?.memory || '';
  } catch(e) { return ''; }
}

async function saveUserMemory(username, memory) {
  if (!pool) return;
  try {
    await pool.query(`INSERT INTO user_memory (username, memory, updated_at) VALUES ($1,$2,NOW())
      ON CONFLICT (username) DO UPDATE SET memory=$2, updated_at=NOW()`, [username, memory]);
  } catch(e) {}
}

// ─── QAD ─────────────────────────────────────────────────────────
async function getQADSources() {
  if (!pool) return [];
  try {
    const r = await pool.query(`SELECT sheet_key, filename, sheet_name, updated_at, jsonb_array_length(data) as count FROM qad_data ORDER BY updated_at DESC`);
    return r.rows;
  } catch(e) { return []; }
}

// Detectar qué hojas son relevantes según la pregunta
function detectRelevantSheets(userMessage, sources) {
  const msg = norm(userMessage);

  // Mapeo de palabras clave → patrones de nombre de archivo/hoja
  const mappings = [
    { keys: ['cartera','saldo','cliente','cxc','cobrar','deuda','vencido','antigüedad','antiguedad','adeudo','pago','atrasado','moroso','cobranza'], patterns: ['cxc','saldo','cliente','cartera','cobrar','antiguedad','antigüedad'] },
    { keys: ['venta','factura','remision','remisión','ingreso','pedido'], patterns: ['venta','factura','remision','ingreso','pedido'] },
    { keys: ['inventario','stock','existencia','almacen','almacén','producto','tela','material'], patterns: ['inventario','stock','existencia','almacen','producto','tela'] },
    { keys: ['proveedor','cxp','pagar','compra','abastecimiento'], patterns: ['proveedor','cxp','pagar','compra'] },
    { keys: ['especificacion','especificación','ficha','tecnica','técnica','tela','cardigan','jersey'], patterns: ['especificacion','ficha','tecnica','tela'] },
    { keys: ['produccion','producción','manufactura','fabricacion','tejido','acabado'], patterns: ['produccion','manufactura','fabricacion','tejido'] },
  ];

  // Encontrar qué categoría aplica
  let relevantPatterns = [];
  for (const map of mappings) {
    if (map.keys.some(k => msg.includes(k))) {
      relevantPatterns = map.patterns;
      break;
    }
  }

  // Si no hay categoría específica, devolver todas las fuentes
  if (!relevantPatterns.length) return sources;

  // Filtrar fuentes que coincidan con los patrones
  const relevant = sources.filter(src => {
    const srcNorm = norm(`${src.filename} ${src.sheet_name}`);
    return relevantPatterns.some(p => srcNorm.includes(p));
  });

  // Si encontró fuentes relevantes, usar esas; si no, usar todas
  return relevant.length > 0 ? relevant : sources;
}

async function searchQAD(userMessage, maxRows = 500) {
  if (!pool) return { rows: [], total: 0, sources: [] };

  const allSources = await getQADSources();
  if (!allSources.length) return { rows: [], total: 0, sources: [] };

  const totalRows = allSources.reduce((s, r) => s + (parseInt(r.count) || 0), 0);

  // Detectar hojas relevantes para esta pregunta
  const relevantSources = detectRelevantSheets(userMessage, allSources);

  console.log(`Buscando en ${relevantSources.length} de ${allSources.length} hojas para: "${userMessage.substring(0,50)}"`);
  console.log(`Hojas seleccionadas: ${relevantSources.map(s => s.filename).join(', ')}`);

  const stopWords = new Set(['dame','reporte','tabla','de','del','la','el','los','las','un','una','en','que','tienes','por','para','con','me','mi','mis','tu','sus','hay','son','mas','top','cinco','diez','todos','todas','cuales','cual','como','cuando','donde','quiero','ver','mostrar','genera','dime','muestra','necesito']);
  const msg = norm(userMessage);
  const keywords = msg.split(' ').filter(w => w.length > 2 && !stopWords.has(w));

  let allMatches = [];

  for (const src of relevantSources) {
    try {
      // Siempre traer TODOS los registros de hojas relevantes (son pocas y específicas)
      const r = await pool.query(
        `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1`,
        [src.sheet_key]
      );
      r.rows.forEach(row => allMatches.push({ ...row.row, _src: `${src.filename} / ${src.sheet_name}` }));
    } catch(e) {
      console.error('Error buscando en', src.sheet_key, e.message);
    }
  }

  // Si no encontró nada en hojas relevantes, buscar en todas con keywords
  if (!allMatches.length && keywords.length > 0) {
    for (const src of allSources.slice(0, 10)) {
      try {
        const conditions = keywords.map((_, i) => `lower(elem::text) LIKE $${i + 2}`).join(' OR ');
        const q = `SELECT elem as row FROM qad_data, jsonb_array_elements(data) AS elem WHERE sheet_key = $1 AND (${conditions}) LIMIT 100`;
        const r = await pool.query(q, [src.sheet_key, ...keywords.map(k => `%${k}%`)]);
        r.rows.forEach(row => allMatches.push({ ...row.row, _src: `${src.filename} / ${src.sheet_name}` }));
      } catch(e) {}
    }
  }

  return { rows: allMatches.slice(0, maxRows), total: totalRows, sources: relevantSources };
}

async function buildQADContext(userMessage = '') {
  try {
    if (!pool) return { hasData: false, text: 'Sin conexión a base de datos.' };
    if (!(await tableExists('qad_data'))) return { hasData: false, text: 'No hay datos QAD cargados.' };

    const countR = await pool.query('SELECT COUNT(*) as c FROM qad_data');
    if (parseInt(countR.rows[0].c) === 0) return { hasData: false, text: 'No hay archivos QAD cargados. El administrador debe subirlos desde el Panel de Sistemas → Archivos QAD.' };

    const { rows, total, sources } = await searchQAD(userMessage, 500);
    if (!rows.length) return { hasData: false, text: 'No encontré datos relacionados con tu consulta.' };

    // Resumen de fuentes
    let ctx = `DATOS REALES DE QAD — ${total.toLocaleString('es-MX')} registros totales\n`;
    ctx += `Archivos cargados:\n`;
    sources.forEach(s => {
      ctx += `  • ${s.filename} / ${s.sheet_name} — ${s.count} registros — ${new Date(s.updated_at).toLocaleString('es-MX')}\n`;
    });
    ctx += `\nDATOS PARA ESTA CONSULTA (${rows.length} registros):\n\n`;

    // Agrupar por fuente
    const bySrc = {};
    rows.forEach(row => {
      const src = row._src || 'QAD';
      if (!bySrc[src]) bySrc[src] = [];
      const clean = { ...row };
      delete clean._src;
      bySrc[src].push(clean);
    });

    Object.keys(bySrc).forEach(src => {
      const srcRows = bySrc[src];
      if (!srcRows.length) return;
      const cols = Object.keys(srcRows[0]);
      ctx += `=== ${src} ===\n`;
      ctx += cols.join(' | ') + '\n';
      ctx += cols.map(() => '---').join(' | ') + '\n';
      srcRows.forEach(row => { ctx += cols.map(c => fmt(row[c])).join(' | ') + '\n'; });
      ctx += '\n';
    });

    if (total > rows.length) ctx += `\n⚠️ Mostrando ${rows.length} de ${total} registros. Para ver más especifica cliente, vendedor o período.`;

    return { hasData: true, text: ctx, total, shown: rows.length };
  } catch(e) {
    console.error('buildQADContext error:', e.message);
    return { hasData: false, text: 'Error cargando QAD: ' + e.message };
  }
}

// ─── EXCEL PROFESIONAL ────────────────────────────────────────────
function generateProfessionalExcel(data) {
  const wb = xlsx.utils.book_new();
  const { titulo = 'Reporte', subtitulo = '', periodo = '', usuario = '', hojas = [] } = data;

  // Colores corporativos Vivatex
  const VERDE = '1E6B2E';
  const VERDE_MED = '4A9A20';
  const VERDE_L = 'D4EDDA';
  const BLANCO = 'FFFFFF';
  const GRIS = 'F5F5F5';
  const AZUL = '1A3A5C';
  const NEGRO = '1C1C1C';

  const borderStyle = {
    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
  };

  hojas.forEach(hoja => {
    const { nombre = 'Datos', columnas = [], filas = [], totales = null } = hoja;
    const ws = {};
    const nCols = columnas.length;
    let maxRow = 1;

    const setCell = (r, c, value, opts = {}) => {
      const addr = xlsx.utils.encode_cell({ r: r - 1, c: c - 1 });
      ws[addr] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
      if (opts.font || opts.fill || opts.alignment || opts.border || opts.numFmt) {
        ws[addr].s = {
          font: opts.font || {},
          fill: opts.fill || {},
          alignment: opts.alignment || { vertical: 'center', wrapText: false },
          border: opts.border || {},
          numFmt: opts.numFmt || ''
        };
      }
      if (r > maxRow) maxRow = r;
    };

    // Fila 1: Header empresa
    for (let c = 1; c <= nCols; c++) {
      setCell(1, c, c === 1 ? `GRUPO VIVATEX S.A. DE C.V.  ·  ${titulo.toUpperCase()}` : '', {
        font: { name: 'Calibri', sz: 13, bold: true, color: { argb: 'FF' + BLANCO } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + VERDE } },
        alignment: { horizontal: 'left', vertical: 'center' }
      });
    }

    // Fila 2: Subtítulo/meta
    const metaText = [subtitulo, periodo ? `Período: ${periodo}` : '', usuario ? `Usuario: ${usuario}` : ''].filter(Boolean).join('   |   ');
    for (let c = 1; c <= nCols; c++) {
      setCell(2, c, c === 1 ? metaText : '', {
        font: { name: 'Calibri', sz: 9, italic: true, color: { argb: 'FF555555' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + VERDE_L } },
        alignment: { horizontal: 'left', vertical: 'center' }
      });
    }

    // Fila 3: Fecha generación
    const fechaText = `Generado: ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}   |   Confidencial · Uso exclusivo interno`;
    for (let c = 1; c <= nCols; c++) {
      setCell(3, c, c === 1 ? fechaText : '', {
        font: { name: 'Calibri', sz: 8, italic: true, color: { argb: 'FF888888' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFAFA' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      });
    }

    // Fila 4: Espacio
    for (let c = 1; c <= nCols; c++) {
      setCell(4, c, '', { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } });
    }

    // Fila 5: Headers columnas
    columnas.forEach((col, i) => {
      setCell(5, i + 1, col, {
        font: { name: 'Calibri', sz: 11, bold: true, color: { argb: 'FF' + BLANCO } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + AZUL } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
        border: borderStyle
      });
    });

    // Filas de datos
    filas.forEach((fila, ri) => {
      const bgColor = ri % 2 === 0 ? 'FFFFFFFF' : 'FFF5FAF3';
      fila.forEach((val, ci) => {
        const isNum = typeof val === 'number';
        setCell(6 + ri, ci + 1, val, {
          font: { name: 'Calibri', sz: 10, color: { argb: 'FF' + NEGRO } },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } },
          alignment: { horizontal: isNum ? 'right' : 'left', vertical: 'center' },
          border: borderStyle,
          numFmt: isNum ? '#,##0.00' : ''
        });
      });
    });

    // Fila totales
    if (totales && totales.length) {
      const tr = 6 + filas.length + 1;
      totales.forEach((val, ci) => {
        setCell(tr, ci + 1, val, {
          font: { name: 'Calibri', sz: 11, bold: true, color: { argb: 'FF' + BLANCO } },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + VERDE_MED } },
          alignment: { horizontal: typeof val === 'number' ? 'right' : 'center', vertical: 'center' },
          border: borderStyle,
          numFmt: typeof val === 'number' ? '#,##0.00' : ''
        });
      });
    }

    // Anchos de columna automáticos
    const colWidths = columnas.map((col, i) => {
      let maxW = col.length + 4;
      filas.forEach(fila => {
        const len = String(fila[i] ?? '').length;
        if (len > maxW) maxW = len;
      });
      return { wch: Math.min(Math.max(maxW + 2, 12), 45) };
    });

    // Alturas de fila
    ws['!rows'] = [
      { hpt: 36 }, // header empresa
      { hpt: 18 }, // meta
      { hpt: 14 }, // fecha
      { hpt: 8 },  // espacio
      { hpt: 28 }, // headers columnas
      ...filas.map(() => ({ hpt: 20 })),
      ...(totales ? [{ hpt: 24 }] : [])
    ];

    ws['!cols'] = colWidths;
    ws['!ref'] = xlsx.utils.encode_range({ r: 0, c: 0 }, { r: maxRow - 1, c: nCols - 1 });

    // Merges para header y meta
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: nCols - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: nCols - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: nCols - 1 } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: nCols - 1 } },
    ];

    xlsx.utils.book_append_sheet(wb, ws, String(nombre).slice(0, 31));
  });

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx', bookSST: false, cellStyles: true });
}

// ─── UPLOAD ───────────────────────────────────────────────────────
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

function parseExcel(buffer) {
  const wb = xlsx.read(buffer, { type: 'buffer' });
  const out = {};
  wb.SheetNames.forEach(s => { out[s] = xlsx.utils.sheet_to_json(wb.Sheets[s], { defval: '' }); });
  return out;
}

async function saveQAD(key, filename, sheet, data) {
  if (!pool) return;
  await pool.query(`INSERT INTO qad_data (sheet_key,filename,sheet_name,data,updated_at) VALUES ($1,$2,$3,$4,NOW()) ON CONFLICT (sheet_key) DO UPDATE SET filename=$2,sheet_name=$3,data=$4,updated_at=NOW()`, [key, filename, sheet, JSON.stringify(data)]);
}

app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  let sheets = 0, files = 0;
  for (const file of req.files || []) {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      try {
        const parsed = parseExcel(file.buffer);
        for (const sheet of Object.keys(parsed)) {
          const key = `${file.originalname}_${sheet}`.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 200);
          await saveQAD(key, file.originalname, sheet, parsed[sheet]);
          sheets++;
        }
        files++;
      } catch(e) { console.error('Error Excel:', file.originalname, e.message); }
    }
    if (ext === '.pdf') {
      try {
        const p = await pdfParse(file.buffer);
        const data = String(p.text || '').split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x }));
        const key = `${file.originalname}_PDF`.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 200);
        await saveQAD(key, file.originalname, 'PDF', data);
        sheets++; files++;
      } catch(e) { console.error('Error PDF:', file.originalname, e.message); }
    }
  }
  res.json({ ok: true, files, sheets });
});

app.get('/api/qad/status', async (req, res) => {
  if (!pool) return res.json({ hasData: false, totalRecords: 0, sheets: [] });
  try {
    const r = await pool.query('SELECT sheet_key, filename, sheet_name, jsonb_array_length(data) as count, updated_at FROM qad_data ORDER BY updated_at DESC');
    const total = r.rows.reduce((s, row) => s + (parseInt(row.count) || 0), 0);
    res.json({ hasData: total > 0, totalRecords: total, sheets: r.rows });
  } catch(e) { res.json({ hasData: false, totalRecords: 0, sheets: [] }); }
});

app.delete('/api/qad/clear', async (req, res) => {
  if (pool) await pool.query('DELETE FROM qad_data');
  res.json({ ok: true });
});

// ─── USUARIOS ─────────────────────────────────────────────────────
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!pool) return res.status(500).json({ ok: false, error: 'Sin DB' });
  try {
    const r = await pool.query('SELECT id,username,role,active FROM usuarios WHERE UPPER(username)=$1 AND password=$2', [String(username || '').toUpperCase(), password]);
    if (!r.rows.length) return res.json({ ok: false, error: 'Usuario o contraseña incorrectos' });
    if (!r.rows[0].active) return res.json({ ok: false, error: 'Usuario inactivo' });
    res.json({ ok: true, user: r.rows[0] });
  } catch(e) { res.status(500).json({ ok: false, error: 'Error de servidor' }); }
});

app.get('/api/usuarios', async (req, res) => {
  if (!pool) return res.json({ users: [] });
  try { const r = await pool.query('SELECT id,username,role,active FROM usuarios ORDER BY id'); res.json({ users: r.rows }); }
  catch(e) { res.json({ users: [] }); }
});

app.post('/api/usuarios', async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.json({ ok: false, error: 'Faltan datos' });
  if (!pool) return res.json({ ok: false, error: 'Sin DB' });
  try {
    await pool.query(`INSERT INTO usuarios (username,password,role,active) VALUES ($1,$2,$3,true) ON CONFLICT (username) DO UPDATE SET password=$2,role=$3,active=true`, [username.toUpperCase(), password, role || 'user']);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: false, error: e.message }); }
});

app.delete('/api/usuarios/:username', async (req, res) => {
  if (!pool) return res.json({ ok: false });
  try { await pool.query('UPDATE usuarios SET active=false WHERE UPPER(username)=$1', [req.params.username.toUpperCase()]); res.json({ ok: true }); }
  catch(e) { res.json({ ok: false }); }
});

// ─── MEMORIA ──────────────────────────────────────────────────────
app.get('/api/memory/:username', async (req, res) => {
  const mem = await getUserMemory(req.params.username);
  res.json({ memory: mem });
});

app.post('/api/memory/:username', async (req, res) => {
  const { memory } = req.body || {};
  await saveUserMemory(req.params.username, memory || '');
  res.json({ ok: true });
});

// ─── CHAT ─────────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages, username } = req.body || {};
  if (!ANTHROPIC_API_KEY) return res.json({ reply: '⚠️ Falta ANTHROPIC_API_KEY en Railway.' });

  try {
    const lastUserMsg = (messages || []).filter(m => m.role === 'user').slice(-1)[0]?.content || '';
    const [qadCtx, userMemory] = await Promise.all([
      buildQADContext(lastUserMsg),
      getUserMemory(username || '')
    ]);

    const systemPrompt = `Eres AVIVA, la analista de inteligencia artificial de Grupo Vivatex S.A. de C.V.

IDENTIDAD:
- Tu nombre es AVIVA. NUNCA menciones Claude, Anthropic ni OpenAI.
- Trabajas exclusivamente para Grupo Vivatex S.A. de C.V.
- Hablas español mexicano natural, tuteas al usuario, eres directa y profesional.
- Tienes memoria de conversaciones anteriores con cada usuario.

REGLAS DE ORO — NUNCA LAS ROMPAS:
1. NUNCA inventes clientes, vendedores, productos, cantidades ni fechas
2. SOLO usa datos que aparezcan LITERALMENTE en el contexto QAD que recibes
3. Si el dato no está en QAD: "Ese dato no está en la información que tengo. Verifica en QAD."
4. NUNCA hagas gráficas ASCII con █ ▓ ░ ni tablas de texto plano
5. NUNCA muestres JSON crudo, arrays ni objetos técnicos
6. Muestra SIEMPRE el listado COMPLETO — nunca omitas registros sin avisar
7. Copia nombres EXACTAMENTE como están en los datos — sin cambiar ni una letra
8. Copia números EXACTAMENTE — sin redondear ni aproximar

FORMATO OBLIGATORIO DE RESPUESTAS:

Para CUALQUIER dato tabular usa SIEMPRE tabla Markdown:
| Columna 1 | Columna 2 | Columna 3 |
|-----------|-----------|-----------|
| Dato 1    | Dato 2    | Dato 3    |
| **TOTAL** | **$X,XXX**| |

Reglas de formato:
- ## Encabezado principal (una sola vez al inicio)
- **Negritas** para valores importantes y totales
- Ordena de MAYOR a MENOR por saldo/importe cuando aplique
- Incluye siempre fila de TOTALES al final de tablas numéricas
- Después de la tabla agrega 2-3 líneas de análisis ejecutivo
- Respuestas completas, claras y fáciles de leer

CUANDO GENERES UN EXCEL agrega al FINAL de tu respuesta (después del análisis):
%%EXCEL%%{"titulo":"Nombre","subtitulo":"Desc","periodo":"Mayo 2026","hojas":[{"nombre":"Hoja1","columnas":["Col1","Col2"],"filas":[["val1",1234]],"totales":["TOTAL",1234]}]}%%EXCEL%%

${userMemory ? `\nMEMORIA DE ESTE USUARIO:\n${userMemory}\n` : ''}

DATOS QAD DISPONIBLES:
${qadCtx.hasData ? qadCtx.text : qadCtx.text}

Usuario: ${username || 'Usuario'}
Fecha: ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}`;

    let cleanMessages = (messages || [])
      .filter(m => (m.role === 'user' || m.role === 'assistant') && String(m.content || '').trim())
      .slice(-12)
      .map(m => ({ role: m.role, content: String(m.content || '').substring(0, 3000) }));

    while (cleanMessages.length && cleanMessages[0].role !== 'user') cleanMessages.shift();
    if (!cleanMessages.length) {
      const fallback = String(lastUserMsg || '').trim();
      cleanMessages = [{ role: 'user', content: fallback || '¡Hola!' }];
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 4096, system: systemPrompt, messages: cleanMessages })
    });

    const data = await response.json();
    if (!response.ok) {
      console.error('Anthropic error:', JSON.stringify(data));
      return res.json({ reply: `Error: ${data?.error?.message || response.status}` });
    }

    let reply = data.content?.[0]?.text || 'Sin respuesta.';

    // Extraer datos Excel si los hay
    let excelData = null;
    const excelMatch = reply.match(/%%EXCEL%%([\s\S]*?)%%EXCEL%%/);
    if (excelMatch) {
      try {
        excelData = JSON.parse(excelMatch[1].trim());
        reply = reply.replace(/%%EXCEL%%[\s\S]*?%%EXCEL%%/, '').trim();
      } catch(e) {}
    }

    // Actualizar memoria si hay info relevante
    const memoryTriggers = ['recuerda','mi nombre es','llámame','soy el','soy la','trabajo en','mi área'];
    if (memoryTriggers.some(t => lastUserMsg.toLowerCase().includes(t))) {
      const currentMem = await getUserMemory(username || '');
      const newMem = currentMem + `\n[${new Date().toLocaleDateString('es-MX')}] ${lastUserMsg.substring(0, 200)}`;
      await saveUserMemory(username || '', newMem.substring(0, 2000));
    }

    res.json({ reply, excelData });

  } catch(e) {
    console.error('Error /api/chat:', e.message);
    res.status(500).json({ reply: 'Error interno. Intenta de nuevo.' });
  }
});

// ─── CONVERSACIONES ───────────────────────────────────────────────
app.post('/api/conversations/save', async (req, res) => {
  const { username, convId, title, messages } = req.body || {};
  if (!pool || !username || !convId) return res.json({ ok: true });
  try {
    await pool.query(`INSERT INTO conversations (username,conv_id,title,messages,updated_at) VALUES ($1,$2,$3,$4,NOW()) ON CONFLICT (username,conv_id) DO UPDATE SET title=$3,messages=$4,updated_at=NOW()`,
      [username, convId, title || 'Conversación', JSON.stringify(messages || [])]);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: true }); }
});

app.get('/api/conversations/:username', async (req, res) => {
  if (!pool) return res.json({ conversations: [] });
  try {
    const r = await pool.query('SELECT conv_id,title,messages,updated_at FROM conversations WHERE username=$1 ORDER BY updated_at DESC LIMIT 50', [req.params.username]);
    res.json({ conversations: r.rows });
  } catch(e) { res.json({ conversations: [] }); }
});

app.delete('/api/conversations/:username/:convId', async (req, res) => {
  if (!pool) return res.json({ ok: true });
  try { await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [req.params.username, req.params.convId]); res.json({ ok: true }); }
  catch(e) { res.json({ ok: true }); }
});

// ─── EXCEL PROFESIONAL ────────────────────────────────────────────
app.post('/api/excel/generate', async (req, res) => {
  try {
    const buf = generateProfessionalExcel(req.body || {});
    res.json({
      ok: true,
      base64: buf.toString('base64'),
      filename: `Vivatex_${String(req.body?.titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`
    });
  } catch(e) {
    console.error('Error Excel:', e.message);
    res.json({ ok: false, error: e.message });
  }
});

// ─── HEALTH ───────────────────────────────────────────────────────
app.get('/api/health', async (req, res) => {
  res.json({ ok: true, db: !!pool, model: ANTHROPIC_MODEL, provider: 'anthropic' });
});

app.get('/manifest.json', (req, res) => res.json({ name: 'AVIVA', short_name: 'AVIVA', start_url: '/', display: 'standalone', theme_color: '#1a1f16' }));

app.get('*', (req, res) => {
  const index = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(index)) return res.sendFile(index);
  res.send('AVIVA ONLINE');
});

initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE — Puerto ${PORT}`);
    console.log(`🤖 Modelo: ${ANTHROPIC_MODEL}`);
    console.log(`🔑 API Key: ${ANTHROPIC_API_KEY ? 'Configurada ✅' : 'FALTA ❌'}`);
  });
});
