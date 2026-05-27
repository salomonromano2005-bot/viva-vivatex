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
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';
const AI_PROVIDER = process.env.AI_PROVIDER || 'openai';

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

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
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

function getSearchTerms(message) {
  const stopWords = new Set([
    'que', 'como', 'cual', 'cuales', 'dame', 'dime', 'muestra', 'genera',
    'generar', 'hacer', 'haz', 'para', 'por', 'con', 'sin', 'los', 'las',
    'del', 'una', 'uno', 'este', 'esta', 'hay', 'tiene', 'tienes', 'pueden',
    'quiero', 'necesito', 'favor', 'reporte', 'tabla', 'excel', 'pdf',
    'info', 'informacion', 'datos', 'todo', 'todos', 'toda', 'todas',
    'top', 'mayor', 'menor', 'mas', 'menos', 'archivo', 'archivos',
    'qad', 'sobre', 'de', 'la', 'el', 'en', 'y', 'o', 'a', 'un',
    'al', 'me', 'lo', 'explica', 'explicame', 'detalla', 'detalle',
    'cliente', 'clientes', 'vendedor', 'vendedores', 'principal',
    'principales', 'registros'
  ]);

  return normalizeText(message)
    .split(' ')
    .map(w => w.trim())
    .filter(w => w.length > 2 && !stopWords.has(w));
}

function htmlTable(rows, maxRows = 60) {
  if (!rows || rows.length === 0) return '<p><strong>No hay registros disponibles.</strong></p>';

  const first = rows.find(r => r && typeof r === 'object');
  if (!first) return '<p><strong>No hay registros disponibles.</strong></p>';

  const allCols = Object.keys(first);
  if (allCols.length === 0) return '<p><strong>No hay columnas disponibles.</strong></p>';

  const priorityWords = [
    'cliente', 'codigo', 'código', 'nombre', 'razon', 'razón', 'vendedor',
    'saldo', 'vencido', 'corriente', 'credito', 'crédito', 'limite', 'límite',
    'pedido', 'factura', 'fecha', 'importe', 'total', 'cantidad', 'inventario',
    'producto', 'descripcion', 'descripción', 'cartera', 'cxc'
  ];

  const priorityCols = allCols.filter(c => {
    const nc = normalizeText(c);
    return priorityWords.some(w => nc.includes(normalizeText(w)));
  });

  const otherCols = allCols.filter(c => !priorityCols.includes(c));
  const cols = [...priorityCols, ...otherCols].slice(0, 18);
  const limited = rows.slice(0, maxRows);

  return `
<div style="overflow-x:auto;width:100%;margin:14px 0;">
<table style="border-collapse:collapse;width:100%;font-size:13px;">
<thead>
<tr>
${cols.map(c => `<th style="border:1px solid #b7c9ad;padding:8px;background:#eaf4e4;text-align:left;">${escapeHtml(c)}</th>`).join('')}
</tr>
</thead>
<tbody>
${limited.map(row => `
<tr>
${cols.map(c => `<td style="border:1px solid #d6e3cf;padding:8px;vertical-align:top;">${escapeHtml(formatValue(row[c]))}</td>`).join('')}
</tr>
`).join('')}
</tbody>
</table>
</div>
${rows.length > maxRows ? `<p><em>Mostrando ${maxRows} de ${rows.length} registros encontrados.</em></p>` : ''}
`;
}

function textRowsToObjects(text) {
  const lines = String(text || '')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  return lines.slice(0, 1000).map((line, i) => ({
    linea: i + 1,
    texto: line,
  }));
}

function normalizeDataPayload(data) {
  if (!data) return [];

  if (Array.isArray(data)) return data;

  if (typeof data === 'string') {
    try {
      const parsed = JSON.parse(data);
      return normalizeDataPayload(parsed);
    } catch {
      return textRowsToObjects(data);
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

async function tableExists(tableName) {
  if (!pool) return false;

  try {
    const r = await pool.query(
      `SELECT EXISTS (
        SELECT FROM information_schema.tables 
        WHERE table_schema='public' AND table_name=$1
      ) AS exists`,
      [tableName]
    );
    return !!r.rows[0]?.exists;
  } catch {
    return false;
  }
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
  const cache = {};
  if (!pool) return cache;

  try {
    const hasQadData = await tableExists('qad_data');

    if (hasQadData) {
      const result = await pool.query('SELECT * FROM qad_data ORDER BY updated_at DESC');

      result.rows.forEach(row => {
        const data = normalizeDataPayload(row.data);
        const key = row.sheet_key || `qad_data_${row.id}`;

        cache[key] = {
          data,
          filename: row.filename || 'qad_data',
          sheet: row.sheet_name || 'QAD',
          updatedAt: row.updated_at,
          sourceTable: 'qad_data',
        };
      });
    }

    const hasQadFiles = await tableExists('qad_files');

    if (hasQadFiles) {
      const result = await pool.query('SELECT * FROM qad_files ORDER BY created_at DESC');

      result.rows.forEach(row => {
        const raw = row.data || row.content || row.text || row.file_content || '';
        const data = normalizeDataPayload(raw);
        const key = `qad_files_${row.id}`;

        cache[key] = {
          data,
          filename: row.filename || row.name || 'qad_files',
          sheet: 'QAD_FILE',
          updatedAt: row.created_at || row.updated_at,
          sourceTable: 'qad_files',
        };
      });
    }

    return cache;
  } catch (e) {
    console.error('Error cargando QAD desde DB:', e.message);
    return {};
  }
}

async function getQADDiagnostics() {
  if (!pool) {
    return {
      postgres: false,
      qad_data_exists: false,
      qad_data_rows: 0,
      qad_files_exists: false,
      qad_files_rows: 0,
    };
  }

  const diag = {
    postgres: true,
    qad_data_exists: false,
    qad_data_rows: 0,
    qad_files_exists: false,
    qad_files_rows: 0,
  };

  try {
    diag.qad_data_exists = await tableExists('qad_data');
    if (diag.qad_data_exists) {
      const r = await pool.query('SELECT COUNT(*)::int AS count FROM qad_data');
      diag.qad_data_rows = r.rows[0]?.count || 0;
    }

    diag.qad_files_exists = await tableExists('qad_files');
    if (diag.qad_files_exists) {
      const r = await pool.query('SELECT COUNT(*)::int AS count FROM qad_files');
      diag.qad_files_rows = r.rows[0]?.count || 0;
    }
  } catch (e) {
    diag.error = e.message;
  }

  return diag;
}

async function clearQADFromDB() {
  if (!pool) return;
  try {
    if (await tableExists('qad_data')) await pool.query('DELETE FROM qad_data');
    qadDataCache = {};
  } catch (e) {
    console.error('Error limpiando QAD DB:', e.message);
  }
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
        sourceTable: c.sourceTable || '',
        row,
      });
    });
  });

  return all;
}

function rowScore(row, terms) {
  const text = normalizeText(JSON.stringify(row));
  let score = 0;

  for (const term of terms) {
    if (text.includes(term)) score += 100;
    const words = text.split(' ');
    if (words.some(w => w.startsWith(term) || term.startsWith(w))) score += 20;
  }

  return score;
}

function bestNumericColumn(rows) {
  if (!rows || rows.length === 0) return null;
  const first = rows.find(r => r && typeof r === 'object');
  if (!first) return null;

  const preferred = ['saldo', 'total', 'importe', 'venta', 'vencido', 'monto', 'cartera'];
  const cols = Object.keys(first);

  let best = null;
  let bestScore = -1;

  for (const col of cols) {
    const nc = normalizeText(col);
    let score = 0;

    preferred.forEach(p => {
      if (nc.includes(p)) score += 30;
    });

    rows.slice(0, 50).forEach(r => {
      if (Math.abs(parseNumber(r[col])) > 0) score += 1;
    });

    if (score > bestScore) {
      bestScore = score;
      best = col;
    }
  }

  return best;
}

function formatQADFilesTable(cache) {
  const keys = Object.keys(cache || {});

  if (keys.length === 0) {
    return '<p>No tengo archivos QAD cargados actualmente.</p>';
  }

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
      Fuente: c.sourceTable || '-',
      'Uso probable': uso,
      'Última actualización': fecha,
    };
  });

  return `
<h2>Archivos QAD cargados</h2>
<p>Estos son los archivos disponibles directamente en PostgreSQL/QAD.</p>
${htmlTable(rows, 200)}
`;
}

async function buildNoRecordsMessage() {
  const d = await getQADDiagnostics();

  return `
<h2>No encontré registros QAD cargados</h2>
<p>Diagnóstico real de PostgreSQL:</p>
${htmlTable([{
    'PostgreSQL conectado': d.postgres ? 'SÍ' : 'NO',
    'Tabla qad_data existe': d.qad_data_exists ? 'SÍ' : 'NO',
    'Registros en qad_data': d.qad_data_rows,
    'Tabla qad_files existe': d.qad_files_exists ? 'SÍ' : 'NO',
    'Registros en qad_files': d.qad_files_rows,
    'Error': d.error || '-',
  }], 1)}
<p>Si antes sí había QAD, no lo borres. Mándame esta pantalla y revisamos la tabla exacta donde quedó guardado.</p>
`;
}

async function buildDeterministicQADReply(cache, message) {
  const msgNorm = normalizeText(message);
  const terms = getSearchTerms(message);
  const allItems = getAllRowsFromCache(cache);

  if (allItems.length === 0) {
    return await buildNoRecordsMessage();
  }

  const asksFiles =
    msgNorm.includes('que archivos') ||
    msgNorm.includes('archivos qad') ||
    msgNorm.includes('que datos tienes') ||
    msgNorm.includes('que tienes cargado') ||
    msgNorm.includes('archivos cargados') ||
    msgNorm.includes('registros en qad');

  if (asksFiles) return formatQADFilesTable(cache);

  const isQADQuestion =
    msgNorm.includes('cliente') ||
    msgNorm.includes('clientes') ||
    msgNorm.includes('vendedor') ||
    msgNorm.includes('vendedores') ||
    msgNorm.includes('saldo') ||
    msgNorm.includes('cartera') ||
    msgNorm.includes('cxc') ||
    msgNorm.includes('venta') ||
    msgNorm.includes('ventas') ||
    msgNorm.includes('pedido') ||
    msgNorm.includes('pedidos') ||
    msgNorm.includes('inventario') ||
    msgNorm.includes('produccion') ||
    msgNorm.includes('producción') ||
    msgNorm.includes('reporte') ||
    msgNorm.includes('tabla');

  if (!isQADQuestion) return null;

  let matches = [];

  if (terms.length > 0) {
    matches = allItems
      .map(item => ({ ...item, score: rowScore(item.row, terms) }))
      .filter(item => item.score > 0)
      .sort((a, b) => b.score - a.score);
  }

  if (matches.length === 0 && terms.length === 0) {
    matches = allItems.slice(0, 300).map(item => ({ ...item, score: 1 }));
  }

  if (matches.length === 0) {
    return `
<h2>Resultado</h2>
<p>No encontré coincidencias exactas en QAD para tu consulta.</p>
<h3>Archivos disponibles</h3>
${formatQADFilesTable(cache)}
`;
  }

  let reportRows = matches.map(m => m.row);

  const wantsTop =
    msgNorm.includes('top') ||
    msgNorm.includes('principal') ||
    msgNorm.includes('principales') ||
    msgNorm.includes('mayor') ||
    msgNorm.includes('mayores');

  if (wantsTop) {
    const numericCol = bestNumericColumn(reportRows);
    if (numericCol) {
      reportRows = [...reportRows]
        .sort((a, b) => parseNumber(b[numericCol]) - parseNumber(a[numericCol]))
        .slice(0, 10);
    } else {
      reportRows = reportRows.slice(0, 10);
    }
  } else {
    reportRows = reportRows.slice(0, 80);
  }

  const sourceMap = {};
  matches.forEach(m => {
    const key = `${m.filename} / ${m.sheet} / ${m.sourceTable}`;
    sourceMap[key] = (sourceMap[key] || 0) + 1;
  });

  const sourceRows = Object.keys(sourceMap).map(k => ({
    Fuente: k,
    Coincidencias: sourceMap[k],
  }));

  return `
<h2>Resumen ejecutivo</h2>
<p>Encontré <strong>${matches.length}</strong> registros reales en QAD relacionados con tu consulta.</p>
<p>La siguiente información fue extraída directamente de PostgreSQL/QAD. No se inventaron clientes, vendedores, importes, fechas ni saldos.</p>

<h2>Fuentes revisadas</h2>
${htmlTable(sourceRows, 20)}

<h2>Tabla principal con información QAD</h2>
${htmlTable(reportRows, wantsTop ? 10 : 80)}

<h2>Observaciones</h2>
<ul>
  <li>La tabla se generó directamente desde QAD/PostgreSQL.</li>
  <li>Si un campo aparece como <strong>NO DISPONIBLE EN QAD</strong>, significa que el campo viene vacío o no existe en el registro original.</li>
  <li>No se completó manualmente ningún dato.</li>
  <li>Si necesitas más registros, pide: <strong>muéstrame más resultados</strong> o <strong>exporta esto a Excel</strong>.</li>
</ul>
`;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (!allowed.includes(ext)) return cb(new Error('Formato no permitido: ' + ext + '. Solo .xlsx, .xls, .csv'));
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
  if (!req.files || req.files.length === 0) return res.status(400).json({ error: 'No se recibieron archivos o formato no permitido' });

  let totalSheets = 0;

  for (const file of req.files) {
    const parsed = parseFile(file.buffer, file.originalname);
    if (!parsed) continue;

    const baseName = path.basename(file.originalname, path.extname(file.originalname));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);

    for (const sheet of Object.keys(parsed)) {
      const key = `${baseName}_${sheet}_${timestamp}`.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_-]/g, '_').toLowerCase();

      qadDataCache[key] = {
        data: parsed[sheet],
        filename: file.originalname,
        sheet,
        uploadDate: timestamp,
        updatedAt: new Date().toISOString(),
      };

      await saveQADToDB(key, file.originalname, sheet, parsed[sheet]);
      totalSheets++;
    }
  }

  qadLastUpdate = new Date().toISOString();

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
      pdfData = String(parsed.text || '')
        .split('\n')
        .map(l => l.trim())
        .filter(l => l.length > 2)
        .map((line, i) => ({ linea: i + 1, texto: line }));
    } catch (pdfErr) {
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
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/qad/status', async (req, res) => {
  const cache = pool ? await loadQADFromDB() : qadDataCache;
  const keys = Object.keys(cache);
  const diag = await getQADDiagnostics();

  res.json({
    hasData: keys.length > 0,
    sheets: keys,
    lastUpdate: qadLastUpdate,
    totalRecords: keys.reduce((sum, k) => sum + (cache[k]?.data?.length || 0), 0),
    persistent: !!pool,
    diagnostics: diag,
  });
});

app.delete('/api/qad/clear', async (req, res) => {
  qadDataCache = {};
  qadLastUpdate = null;
  await clearQADFromDB();
  res.json({ ok: true });
});

app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ ok: false, error: 'Faltan datos' });

  try {
    if (!pool) return res.status(500).json({ ok: false, error: 'Base de datos no disponible' });

    const result = await pool.query(
      'SELECT id, username, role, active FROM usuarios WHERE UPPER(username)=$1 AND password=$2',
      [String(username).trim().toUpperCase(), password]
    );

    if (result.rows.length === 0) return res.json({ ok: false, error: 'Usuario o contraseña incorrectos' });

    const user = result.rows[0];
    if (!user.active) return res.json({ ok: false, error: 'Usuario inactivo' });

    res.json({ ok: true, user });
  } catch (e) {
    res.status(500).json({ ok: false, error: 'Error del servidor' });
  }
});

app.get('/api/health', async (req, res) => {
  const diag = await getQADDiagnostics();

  res.json({
    ok: true,
    provider: AI_PROVIDER,
    openai: !!OPENAI_API_KEY,
    db: !!pool,
    model: OPENAI_MODEL,
    time: new Date().toISOString(),
    diagnostics: diag,
  });
});

app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body || {};

  if (!messages || !Array.isArray(messages) || messages.length === 0) return res.status(400).json({ error: 'Faltan mensajes' });

  try {
    const cache = pool ? await loadQADFromDB() : qadDataCache;
    const lastMsg = messages[messages.length - 1]?.content || '';

    const directReply = await buildDeterministicQADReply(cache, lastMsg);
    if (directReply) return res.json({ reply: directReply });

    if (!OPENAI_API_KEY) {
      return res.status(500).json({
        error: 'Falta OPENAI_API_KEY',
        detail: 'Configura OPENAI_API_KEY en Railway > Variables',
      });
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
            content: `${system || ''}\n\nEres AVIVA. No inventes datos. Si no hay datos QAD suficientes, dilo claramente.`,
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
    res.json({ reply: data.choices?.[0]?.message?.content || 'Sin respuesta' });
  } catch (err) {
    console.error('Error en /api/chat:', err);
    res.status(500).json({ error: 'Error interno del servidor', detail: err.message });
  }
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
    res.json({ conversations: [] });
  }
});

app.delete('/api/conversations/:username/:convId', async (req, res) => {
  const { username, convId } = req.params;

  try {
    if (pool) await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [username, convId]);
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
    ['Usuario', data.usuario || 'Usuario'],
    ['Fecha', new Date().toLocaleString('es-MX')],
  ];

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(portada), 'Portada');

  for (const hoja of data.hojas || []) {
    const nombre = String(hoja.nombre || 'Datos').substring(0, 31);
    const columnas = hoja.columnas || [];
    const filas = hoja.filas || [];
    const ws = xlsx.utils.aoa_to_sheet([columnas, ...filas]);
    xlsx.utils.book_append_sheet(wb, ws, nombre);
  }

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.post('/api/excel/generate', async (req, res) => {
  const { titulo, usuario, hojas } = req.body || {};

  if (!hojas || !Array.isArray(hojas) || hojas.length === 0) return res.status(400).json({ error: 'Sin datos para Excel' });

  const filename = `Vivatex_${String(titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`;

  const buffer = generateBasicExcelFallback({
    titulo: titulo || 'Reporte Vivatex',
    usuario: usuario || 'Usuario',
    hojas,
  });

  res.json({
    ok: true,
    base64: buffer.toString('base64'),
    filename,
    fallback: true,
  });
});

app.get('/manifest.json', (req, res) => {
  res.setHeader('Content-Type', 'application/manifest+json');
  res.json({
    name: 'AVIVA — Vivatex IA',
    short_name: 'AVIVA',
    start_url: '/',
    scope: '/',
    display: 'standalone',
    background_color: '#1E6B2E',
    theme_color: '#1E6B2E',
    lang: 'es-MX',
    icons: [],
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
  if (fs.existsSync(indexPath)) return res.sendFile(indexPath);
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
