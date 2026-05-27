const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const multer = require('multer');
const { spawn } = require('child_process');
const xlsx = require('xlsx');
const pdfParse = require('pdf-parse');
const { Pool } = require('pg');

const app = express();

app.use(express.json({ limit: '50mb' }));

app.use((req, res, next) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
const DATABASE_URL = process.env.DATABASE_URL || '';
const AI_PROVIDER = process.env.AI_PROVIDER || 'openai';
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';

let pool = null;
let qadDataCache = {};
let qadLastUpdate = null;

function getPgSslConfig() {
  if (!DATABASE_URL) return false;
  if (DATABASE_URL.includes('localhost') || DATABASE_URL.includes('127.0.0.1')) return false;
  return { rejectUnauthorized: false };
}

async function initDB() {
  if (!DATABASE_URL) {
    console.log('⚠️ Sin DATABASE_URL — usando memoria temporal');
    return;
  }

  try {
    pool = new Pool({
      connectionString: DATABASE_URL,
      ssl: getPgSslConfig(),
    });

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

    const baseUsers = [
      { username: 'JRZ123', password: 'JACOBOrz4646', role: 'user' },
      { username: 'JRR234', password: 'JACOBOrr8989', role: 'user' },
      { username: 'SRR456', password: 'SIMONrr0202', role: 'user' },
      { username: 'SRB0707', password: 'SALOrb0909', role: 'user' },
      { username: 'SISTEMAS1900', password: 'SISTEMASviva2026', role: 'admin' },
      { username: 'LEO2026', password: 'LEOg1986', role: 'user' },
      { username: 'ISMAEL_VENTAS', password: 'Vivatex2026', role: 'user' },
      { username: 'ISRAEL_ACABADO', password: 'Vivatex2026', role: 'user' },
      { username: 'MEMO_TEJIDO', password: 'Vivatex2026', role: 'user' },
      { username: 'CARLOSM_H', password: 'Vivatex2026', role: 'user' },
      { username: 'MARTIN_CONTA', password: 'Vivatex2026', role: 'user' },
      { username: 'VENTAS_MONICA', password: 'Vivatex2026', role: 'user' },
      { username: 'VENTAS_EDGAR', password: 'Vivatex2026', role: 'user' },
      { username: 'VENTAS_AMELIA', password: 'Vivatex2026', role: 'user' },
      { username: 'VENTAS_JORGE', password: 'Vivatex2026', role: 'user' },
      { username: 'GTE_ACABADO', password: 'Vivatex2026', role: 'user' },
      { username: 'GTE_TEJIDO', password: 'Vivatex2026', role: 'user' },
      { username: 'GTE_HILATURA', password: 'Vivatex2026', role: 'user' },
      { username: 'COBRANZA_CLAUDIA', password: 'Vivatex2026', role: 'user' },
      { username: 'COBRANZA_OLIVIA', password: 'Vivatex2026', role: 'user' },
      { username: 'ADMON_LUCY', password: 'Vivatex2026', role: 'user' },
    ];

    for (const u of baseUsers) {
      await pool.query(`
        INSERT INTO usuarios (username, password, role, active)
        VALUES ($1, $2, $3, true)
        ON CONFLICT (username) DO NOTHING
      `, [u.username, u.password, u.role]);
    }

    console.log('✅ PostgreSQL conectado');
  } catch (e) {
    console.error('❌ Error conectando PostgreSQL:', e.message);
    pool = null;
  }
}

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatValue(value) {
  if (value === null || value === undefined || value === '') return 'NO DISPONIBLE EN QAD';

  if (typeof value === 'number') {
    return Number.isInteger(value)
      ? value.toLocaleString('es-MX')
      : value.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  return String(value).replace(/\n/g, ' ').replace(/\|/g, '/').trim();
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const cleaned = String(value || '')
    .replace(/[$,]/g, '')
    .replace(/\s/g, '')
    .replace(/[^\d.-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function rowsToTable(rows, maxRows = 50) {
  if (!rows || rows.length === 0) return 'No hay registros disponibles.';

  const firstRow = rows.find(r => r && typeof r === 'object');
  if (!firstRow) return 'No hay registros disponibles.';

  const cols = Object.keys(firstRow);
  if (cols.length === 0) return 'No hay columnas disponibles.';

  const limited = rows.slice(0, maxRows);
  const header = `| ${cols.join(' | ')} |`;
  const separator = `| ${cols.map(() => '---').join(' | ')} |`;

  const lines = limited.map(row =>
    `| ${cols.map(col => formatValue(row[col])).join(' | ')} |`
  );

  return [header, separator, ...lines].join('\n');
}

function getImportantKeywords(text) {
  const stopWords = new Set([
    'que', 'como', 'cual', 'cuales', 'dame', 'dime', 'muestra', 'genera',
    'generar', 'hacer', 'haz', 'para', 'por', 'con', 'sin', 'los', 'las',
    'del', 'una', 'uno', 'este', 'esta', 'hay', 'tiene', 'pueden', 'quiero',
    'necesito', 'favor', 'reporte', 'tabla', 'excel', 'pdf', 'info',
    'informacion', 'datos', 'todo', 'todos', 'toda', 'todas', 'top',
    'mayor', 'menor', 'mas', 'menos', 'archivo', 'archivos', 'qad',
    'sobre', 'de', 'la', 'el', 'en', 'y', 'o', 'a', 'un', 'al', 'me',
    'lo', 'explica', 'explicame', 'detalla', 'detalle', 'cliente',
    'clientes', 'vendedor', 'vendedores'
  ]);

  return normalizeText(text)
    .split(' ')
    .map(w => w.trim())
    .filter(w => w.length > 2 && !stopWords.has(w));
}

function getAllRowsFromCache(cache) {
  const all = [];
  Object.keys(cache || {}).forEach(key => {
    const c = cache[key];
    if (!c || !Array.isArray(c.data)) return;

    c.data.forEach(row => {
      if (!row || typeof row !== 'object') return;
      all.push({
        sourceKey: key,
        filename: c.filename || '',
        sheet: c.sheet || '',
        updatedAt: c.updatedAt || '',
        row,
      });
    });
  });
  return all;
}

function rowScore(row, keywords) {
  const text = normalizeText(JSON.stringify(row));
  let score = 0;

  for (const kw of keywords) {
    if (text.includes(kw)) score += 50;
    const words = text.split(' ');
    if (words.some(w => w.startsWith(kw) || kw.startsWith(w))) score += 10;
  }

  return score;
}

function getBestNumericColumn(rows, preferredWords) {
  const first = rows.find(r => r && typeof r === 'object');
  if (!first) return null;

  const cols = Object.keys(first);
  const normalizedPreferred = preferredWords.map(normalizeText);

  let best = null;
  let bestScore = -1;

  for (const col of cols) {
    const ncol = normalizeText(col);
    let score = 0;

    for (const w of normalizedPreferred) {
      if (ncol.includes(w)) score += 10;
    }

    const numericCount = rows.slice(0, 30).filter(r => Math.abs(parseNumber(r[col])) > 0).length;
    score += numericCount;

    if (score > bestScore) {
      bestScore = score;
      best = col;
    }
  }

  return best;
}

function formatQADFilesTable(cache) {
  const keys = Object.keys(cache || {});
  if (keys.length === 0) return 'No tengo archivos QAD cargados actualmente.';

  const rows = keys.map(k => {
    const c = cache[k] || {};
    const registros = Array.isArray(c.data) ? c.data.length : 0;
    const fecha = c.updatedAt ? new Date(c.updatedAt).toLocaleString('es-MX') : '-';

    let uso = 'Datos generales';
    const name = normalizeText(`${c.filename || ''} ${c.sheet || ''}`);

    if (name.includes('cliente') || name.includes('cxc') || name.includes('cartera') || name.includes('saldo')) uso = 'Clientes / cartera / saldos';
    else if (name.includes('venta') || name.includes('factur')) uso = 'Ventas / facturación';
    else if (name.includes('pedido') || name.includes('orden')) uso = 'Pedidos / órdenes';
    else if (name.includes('inventario') || name.includes('almacen') || name.includes('stock')) uso = 'Inventario / almacén';
    else if (name.includes('produccion') || name.includes('tejido') || name.includes('acabado') || name.includes('merma')) uso = 'Producción / mermas';
    else if (name.includes('pdf')) uso = 'PDF procesado';

    return {
      Archivo: c.filename || '-',
      Hoja: c.sheet || '-',
      Registros: registros,
      'Uso probable': uso,
      'Última actualización': fecha,
    };
  });

  return `## Archivos QAD cargados

${rowsToTable(rows, 200)}

Estos son los archivos disponibles en PostgreSQL para análisis.`;
}

function buildDeterministicQADReply(cache, message) {
  const msgNorm = normalizeText(message);
  const keywords = getImportantKeywords(message);
  const allRows = getAllRowsFromCache(cache);

  if (allRows.length === 0) {
    return 'No tengo registros QAD cargados actualmente.';
  }

  const asksFiles =
    msgNorm.includes('que archivos') ||
    msgNorm.includes('archivos qad') ||
    msgNorm.includes('que datos tienes') ||
    msgNorm.includes('que tienes cargado') ||
    msgNorm.includes('archivos cargados');

  if (asksFiles) {
    return formatQADFilesTable(cache);
  }

  const wantsClient =
    msgNorm.includes('cliente') ||
    msgNorm.includes('clientes') ||
    msgNorm.includes('cartera') ||
    msgNorm.includes('saldo') ||
    msgNorm.includes('cxc') ||
    msgNorm.includes('vencido');

  const wantsVendor =
    msgNorm.includes('vendedor') ||
    msgNorm.includes('vendedores');

  const wantsReport =
    msgNorm.includes('reporte') ||
    msgNorm.includes('tabla') ||
    msgNorm.includes('detalle') ||
    msgNorm.includes('detallado') ||
    msgNorm.includes('informacion');

  const wantsTop10 =
    msgNorm.includes('top 10') ||
    msgNorm.includes('top10') ||
    msgNorm.includes('principales');

  if (!wantsClient && !wantsVendor && !wantsReport) {
    return null;
  }

  let matches = [];

  if (keywords.length > 0) {
    matches = allRows
      .map(item => ({ ...item, score: rowScore(item.row, keywords) }))
      .filter(item => item.score > 0)
      .sort((a, b) => b.score - a.score);
  }

  if (matches.length === 0 && (wantsClient || wantsVendor || wantsReport)) {
    matches = allRows
      .filter(item => {
        const source = normalizeText(`${item.filename} ${item.sheet} ${JSON.stringify(item.row)}`);
        if (wantsClient) return source.includes('cliente') || source.includes('saldo') || source.includes('cartera') || source.includes('cxc');
        if (wantsVendor) return source.includes('vendedor') || source.includes('ventas') || source.includes('factura');
        return true;
      })
      .map(item => ({ ...item, score: 1 }));
  }

  if (matches.length === 0) {
    return `## Resultado

No encontré registros que coincidan directamente con tu consulta en QAD.

## Archivos disponibles

${formatQADFilesTable(cache)}`;
  }

  let rowsForReport = matches.map(m => m.row);

  if (wantsTop10) {
    const sortCol = getBestNumericColumn(rowsForReport, ['saldo', 'total', 'importe', 'venta', 'vencido', 'monto']);
    if (sortCol) {
      rowsForReport = [...rowsForReport].sort((a, b) => parseNumber(b[sortCol]) - parseNumber(a[sortCol])).slice(0, 10);
    } else {
      rowsForReport = rowsForReport.slice(0, 10);
    }
  } else {
    rowsForReport = rowsForReport.slice(0, 50);
  }

  const sourceSummaryMap = {};
  matches.forEach(m => {
    const key = `${m.filename} / ${m.sheet}`;
    sourceSummaryMap[key] = (sourceSummaryMap[key] || 0) + 1;
  });

  const sourcesTable = Object.keys(sourceSummaryMap).map(k => ({
    Archivo: k,
    Coincidencias: sourceSummaryMap[k],
  }));

  return `## Resumen ejecutivo

Se encontraron **${matches.length} registros reales en QAD** relacionados con tu consulta.  
La tabla de abajo se genera directamente desde PostgreSQL/QAD, sin inventar campos ni completar información manualmente.

## Fuentes revisadas

${rowsToTable(sourcesTable, 20)}

## Tabla principal con datos QAD

${rowsToTable(rowsForReport, wantsTop10 ? 10 : 50)}

## Observaciones

- Todos los datos mostrados vienen de QAD.
- Si un campo aparece como **NO DISPONIBLE EN QAD**, significa que está vacío o no existe en el registro original.
- No se inventaron vendedores, clientes, saldos, fechas ni importes.
- Si necesitas más registros, pide: **"muéstrame más resultados"** o **"exporta esto a Excel"**.`;
}

async function saveQADToDB(key, filename, sheetName, data) {
  if (!pool) return;

  try {
    await pool.query(`
      INSERT INTO qad_data (sheet_key, filename, sheet_name, data, updated_at)
      VALUES ($1, $2, $3, $4, NOW())
      ON CONFLICT (sheet_key)
      DO UPDATE SET filename=$2, sheet_name=$3, data=$4, updated_at=NOW()
    `, [key, filename, sheetName, JSON.stringify(data)]);
  } catch (e) {
    console.error('Error guardando QAD en DB:', e.message);
  }
}

async function loadQADFromDB() {
  if (!pool) return {};

  try {
    const result = await pool.query('SELECT * FROM qad_data ORDER BY updated_at DESC');
    const cache = {};

    result.rows.forEach(row => {
      let data = row.data;

      if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch { data = []; }
      }

      if (!Array.isArray(data)) data = data ? [data] : [];

      cache[row.sheet_key] = {
        data,
        filename: row.filename,
        sheet: row.sheet_name,
        updatedAt: row.updated_at,
      };
    });

    return cache;
  } catch (e) {
    console.error('Error cargando QAD desde DB:', e.message);
    return {};
  }
}

async function clearQADFromDB() {
  if (!pool) return;
  try {
    await pool.query('DELETE FROM qad_data');
  } catch (e) {
    console.error('Error limpiando QAD DB:', e.message);
  }
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();

    if (!allowed.includes(ext)) {
      return cb(new Error('Formato no permitido: ' + ext + '. Solo .xlsx, .xls, .csv'));
    }

    cb(null, true);
  },
});

const uploadPDF = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.pdf') return cb(new Error('Solo se aceptan archivos PDF'));
    cb(null, true);
  },
});

function parseFile(buffer, filename) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const result = {};

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      result[sheetName] = xlsx.utils.sheet_to_json(sheet, { defval: '' });
    });

    return result;
  } catch (e) {
    console.error('Error parseando archivo:', filename, e.message);
    return null;
  }
}

app.post('/api/qad/upload', (req, res, next) => {
  upload.array('files', 20)(req, res, err => {
    if (err) return res.status(400).json({ error: err.message });
    next();
  });
}, async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: 'No se recibieron archivos o formato no permitido' });
  }

  let totalSheets = 0;

  for (const file of req.files) {
    const parsed = parseFile(file.buffer, file.originalname);
    if (!parsed) continue;

    const baseName = path.basename(file.originalname, path.extname(file.originalname));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);

    for (const sheet of Object.keys(parsed)) {
      const key = `${baseName}_${sheet}_${timestamp}`
        .replace(/\s+/g, '_')
        .replace(/[^a-zA-Z0-9_-]/g, '_')
        .toLowerCase();

      const entry = {
        data: parsed[sheet],
        filename: file.originalname,
        sheet,
        uploadDate: timestamp,
        updatedAt: new Date().toISOString(),
      };

      qadDataCache[key] = entry;
      await saveQADToDB(key, file.originalname, sheet, parsed[sheet]);
      totalSheets++;
    }
  }

  qadLastUpdate = new Date().toISOString();

  console.log(`📊 QAD cargado: ${req.files.length} archivo(s), ${totalSheets} hoja(s)`);

  res.json({
    ok: true,
    files: req.files.length,
    sheets: totalSheets,
    updatedAt: qadLastUpdate,
    persistent: !!pool,
  });
});

app.post('/api/qad/upload-pdf', uploadPDF.single('pdf'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No se recibió archivo PDF' });

  try {
    let pdfData = [];

    try {
      const parsed = await pdfParse(req.file.buffer);
      const pdfText = parsed.text || '';
      const lines = pdfText
        .split('\n')
        .map(l => l.trim())
        .filter(l => l.length > 2);

      pdfData = lines.map(line => ({ texto: line }));
      console.log(`📄 PDF procesado: ${req.file.originalname} — ${lines.length} líneas`);
    } catch (pdfErr) {
      console.warn('Error leyendo PDF:', pdfErr.message);
      pdfData = [{ texto: '[PDF no legible: ' + req.file.originalname + ']' }];
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const key = 'pdf_' + req.file.originalname.replace(/[^a-z0-9]/gi, '_') + '_' + timestamp;

    qadDataCache[key] = {
      filename: req.file.originalname,
      sheet: 'PDF',
      data: pdfData,
      updatedAt: new Date().toISOString(),
    };

    await saveQADToDB(key, req.file.originalname, 'PDF', pdfData);

    res.json({
      ok: true,
      message: `PDF procesado: ${req.file.originalname} (${pdfData.length} líneas extraídas)`,
      persistent: !!pool,
    });
  } catch (e) {
    console.error('Error PDF:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/qad/status', async (req, res) => {
  const cache = pool ? await loadQADFromDB() : qadDataCache;
  const keys = Object.keys(cache);

  res.json({
    hasData: keys.length > 0,
    sheets: keys,
    lastUpdate: qadLastUpdate,
    totalRecords: keys.reduce((sum, k) => sum + (cache[k]?.data?.length || 0), 0),
    persistent: !!pool,
  });
});

app.delete('/api/qad/clear', async (req, res) => {
  qadDataCache = {};
  qadLastUpdate = null;
  await clearQADFromDB();
  res.json({ ok: true });
});

app.post('/api/conversations/save', async (req, res) => {
  const { username, convId, title, messages } = req.body || {};
  if (!username || !convId) return res.status(400).json({ error: 'Faltan datos' });

  try {
    if (pool) {
      await pool.query(`
        INSERT INTO conversations (username, conv_id, title, messages, updated_at)
        VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (username, conv_id)
        DO UPDATE SET title=$3, messages=$4, updated_at=NOW()
      `, [username, convId, title || 'Conversación', JSON.stringify(messages || [])]);
    }

    res.json({ ok: true });
  } catch (e) {
    console.error('Error guardando conversación:', e.message);
    res.status(500).json({ error: 'Error guardando conversación' });
  }
});

app.get('/api/conversations/:username', async (req, res) => {
  const { username } = req.params;

  try {
    if (!pool) return res.json({ conversations: [] });

    const result = await pool.query(`
      SELECT conv_id, title, messages, created_at, updated_at
      FROM conversations
      WHERE username=$1
      ORDER BY updated_at DESC
      LIMIT 50
    `, [username]);

    res.json({ conversations: result.rows });
  } catch (e) {
    console.error('Error cargando conversaciones:', e.message);
    res.json({ conversations: [] });
  }
});

app.delete('/api/conversations/:username/:convId', async (req, res) => {
  const { username, convId } = req.params;

  try {
    if (pool) {
      await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [username, convId]);
    }

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: 'Error eliminando conversación' });
  }
});

function generateBasicExcelFallback(data) {
  const wb = xlsx.utils.book_new();

  const portada = [
    ['GRUPO VIVATEX S.A. DE C.V.'],
    ['AVIVA — Sistema de Inteligencia Empresarial'],
    [],
    ['Reporte', data.titulo || 'Reporte Vivatex'],
    ['Subtítulo', data.subtitulo || ''],
    ['Usuario', data.usuario || 'Usuario'],
    ['Periodo', data.periodo || ''],
    ['Fecha', new Date().toLocaleString('es-MX')],
  ];

  const wsPortada = xlsx.utils.aoa_to_sheet(portada);
  xlsx.utils.book_append_sheet(wb, wsPortada, 'Portada');

  for (const hoja of data.hojas || []) {
    const nombre = String(hoja.nombre || 'Datos').substring(0, 31);
    const columnas = hoja.columnas || [];
    const filas = hoja.filas || [];
    const aoa = [columnas, ...filas];

    if (hoja.totales) {
      aoa.push([]);
      aoa.push(hoja.totales);
    }

    const ws = xlsx.utils.aoa_to_sheet(aoa);
    xlsx.utils.book_append_sheet(wb, ws, nombre);
  }

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.post('/api/excel/generate', async (req, res) => {
  const { titulo, usuario, hojas, subtitulo, periodo } = req.body || {};

  if (!hojas || !Array.isArray(hojas) || hojas.length === 0) {
    return res.status(400).json({ error: 'Sin datos para Excel' });
  }

  const excelPayload = {
    titulo: titulo || 'Reporte Vivatex',
    subtitulo: subtitulo || '',
    usuario: usuario || 'Usuario',
    periodo: periodo || new Date().toLocaleDateString('es-MX', { year: 'numeric', month: 'long' }),
    hojas,
  };

  const filename = `Vivatex_${String(titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`;

  try {
    const pyPath = path.join(__dirname, 'generate_excel.py');

    if (!fs.existsSync(pyPath)) {
      const buffer = generateBasicExcelFallback(excelPayload);
      return res.json({ ok: true, base64: buffer.toString('base64'), filename, fallback: true });
    }

    const tmpPath = path.join(os.tmpdir(), `vivatex_${Date.now()}.xlsx`);
    const py = spawn('python3', [pyPath, tmpPath]);

    let output = '';
    let errOut = '';

    py.stdin.write(JSON.stringify(excelPayload));
    py.stdin.end();

    py.stdout.on('data', d => output += d.toString());
    py.stderr.on('data', d => errOut += d.toString());

    py.on('error', err => {
      console.error('Error iniciando Python:', err.message);
      const buffer = generateBasicExcelFallback(excelPayload);
      return res.json({
        ok: true,
        base64: buffer.toString('base64'),
        filename,
        fallback: true,
        warning: 'Python no disponible; se generó Excel básico',
      });
    });

    py.on('close', code => {
      try {
        if (code !== 0) {
          console.error('Python Excel falló:', errOut);
          const buffer = generateBasicExcelFallback(excelPayload);
          return res.json({
            ok: true,
            base64: buffer.toString('base64'),
            filename,
            fallback: true,
            warning: 'Python/openpyxl falló; se generó Excel básico',
          });
        }

        const parsed = JSON.parse(output.trim() || '{}');

        if (!parsed.success) {
          const buffer = generateBasicExcelFallback(excelPayload);
          return res.json({
            ok: true,
            base64: buffer.toString('base64'),
            filename,
            fallback: true,
            warning: parsed.error || 'Python no generó archivo',
          });
        }

        const fileBuffer = fs.readFileSync(tmpPath);

        try { fs.unlinkSync(tmpPath); } catch {}

        return res.json({
          ok: true,
          base64: fileBuffer.toString('base64'),
          filename,
          fallback: false,
        });
      } catch (e) {
        console.error('Error procesando Excel:', e.message);
        const buffer = generateBasicExcelFallback(excelPayload);
        return res.json({
          ok: true,
          base64: buffer.toString('base64'),
          filename,
          fallback: true,
          warning: 'Error procesando Excel profesional; se generó Excel básico',
        });
      }
    });
  } catch (e) {
    console.error('Error Excel:', e.message);
    const buffer = generateBasicExcelFallback(excelPayload);
    return res.json({ ok: true, base64: buffer.toString('base64'), filename, fallback: true });
  }
});

app.post('/api/chat', async (req, res) => {
  const { messages, system, username } = req.body || {};

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'Faltan mensajes' });
  }

  if (!OPENAI_API_KEY) {
    return res.status(500).json({
      error: 'Falta OPENAI_API_KEY',
      detail: 'Configura OPENAI_API_KEY en Railway > Variables',
    });
  }

  try {
    const cache = pool ? await loadQADFromDB() : qadDataCache;
    const lastMsg = messages[messages.length - 1]?.content || '';

    const directReply = buildDeterministicQADReply(cache, lastMsg);
    if (directReply) {
      return res.json({ reply: directReply });
    }

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        messages: [
          {
            role: 'system',
            content: `${system || ''}

Eres AVIVA, asistente empresarial de Grupo Vivatex. No inventes datos. Si no tienes QAD suficiente, dilo claramente.`
          },
          ...messages.map(m => ({
            role: m.role === 'assistant' ? 'assistant' : 'user',
            content: typeof m.content === 'string' ? m.content : JSON.stringify(m.content || ''),
          })),
        ],
        temperature: 0,
        max_tokens: 2048,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error('OpenAI API error:', response.status, errText);
      return res.status(response.status).json({
        error: 'Error en API de IA',
        detail: errText,
        status: response.status,
        model: OPENAI_MODEL,
      });
    }

    const data = await response.json();
    const reply = data.choices?.[0]?.message?.content || 'Sin respuesta';

    res.json({ reply });
  } catch (err) {
    console.error('Error en /api/chat:', err);
    res.status(500).json({
      error: 'Error interno del servidor',
      detail: err.message,
    });
  }
});

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

const ONE_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=',
  'base64'
);

app.get('/icons/icon-192.png', (req, res) => {
  res.setHeader('Content-Type', 'image/png');
  res.send(ONE_PIXEL_PNG);
});

app.get('/icons/icon-512.png', (req, res) => {
  res.setHeader('Content-Type', 'image/png');
  res.send(ONE_PIXEL_PNG);
});

app.get('*', (req, res) => {
  const indexPath = path.join(__dirname, 'public', 'index.html');

  if (fs.existsSync(indexPath)) {
    return res.sendFile(indexPath);
  }

  res.status(404).send('AVIVA: public/index.html no encontrado');
});

initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE PORT ${PORT}`);
    console.log(`🤖 Provider: ${AI_PROVIDER}`);
    console.log(`🧠 Model: ${OPENAI_MODEL}`);
    console.log(`🗄️ PostgreSQL: ${pool ? 'conectado' : 'no conectado'}`);
  });
});
