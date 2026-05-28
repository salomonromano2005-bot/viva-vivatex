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

// ─── PERMISOS POR USUARIO ─────────────────────────────────────────
const USER_PERMISSIONS = {
  'JRZ123':          { type: 'full',      panel: false, restrict: [] },
  'JRR234':          { type: 'full',      panel: false, restrict: [] },
  'SRR456':          { type: 'full',      panel: false, restrict: [] },
  'SRB0707':         { type: 'full',      panel: false, restrict: [] },
  'SISTEMAS1900':    { type: 'admin',     panel: true,  restrict: [] },
  'LEO2026':         { type: 'no_conta',  panel: false, restrict: ['contabilidad','activos','pasivos','balance','estado de resultados','cuentas contables','debe haber','mayor general','poliza','polizas'] },
  'ISMAEL_VENTAS':   { type: 'ventas',    panel: false, restrict: [], allow: ['ventas','produccion','pedidos','almacen','bodega','tela acabada','vendedores'] },
  'ISRAEL_ACABADO':  { type: 'no_conta',  panel: false, restrict: ['contabilidad','activos','pasivos','balance','estado de resultados','cuentas contables','debe haber','mayor general','poliza','polizas'] },
  'MEMO_TEJIDO':     { type: 'no_conta',  panel: false, restrict: ['contabilidad','activos','pasivos','balance','estado de resultados','cuentas contables','debe haber','mayor general','poliza','polizas'] },
  'CARLOSM_H':       { type: 'no_conta',  panel: false, restrict: ['contabilidad','activos','pasivos','balance','estado de resultados','cuentas contables','debe haber','mayor general','poliza','polizas'] },
  'MARTIN_CONTA':    { type: 'full',      panel: false, restrict: [] },
  'VENTAS_MONICA':   { type: 'vendedor',  panel: false, restrict: [], vendedor: 'MONICA CACERES',  allow: ['ventas','pedidos','atrasos','clientes','bodega','almacen','tela acabada'] },
  'VENTAS_EDGAR':    { type: 'vendedor',  panel: false, restrict: [], vendedor: 'EDGAR ZARATE',    allow: ['ventas','pedidos','atrasos','clientes','bodega','almacen','tela acabada'] },
  'VENTAS_AMELIA':   { type: 'vendedor',  panel: false, restrict: [], vendedor: 'AMELIA ARAGON',   allow: ['ventas','pedidos','atrasos','clientes','bodega','almacen','tela acabada'] },
  'VENTAS_JORGE':    { type: 'vendedor',  panel: false, restrict: [], vendedor: 'JORGE',           allow: ['ventas','pedidos','atrasos','clientes','bodega','almacen','tela acabada'] },
  'ADMON_LUCY':      { type: 'full',      panel: false, restrict: [] },
};

function getPerms(username) {
  return USER_PERMISSIONS[String(username || '').toUpperCase()] || { type: 'full', panel: false, restrict: [] };
}

function buildPermissionContext(username) {
  const p = getPerms(username);
  let ctx = '';

  if (p.type === 'no_conta') {
    ctx = `RESTRICCIÓN: Este usuario NO puede ver información contable. Si pregunta sobre contabilidad, activos, pasivos, balance general, estado de resultados o cuentas contables, responde: "No tienes acceso a información contable. Consulta con el área de Contabilidad." Para todo lo demás, responde normalmente.`;
  } else if (p.type === 'vendedor') {
    ctx = `RESTRICCIÓN: Este usuario es vendedor (${p.vendedor}). SOLO puede ver:
- Sus propias ventas (filtrar por vendedor = ${p.vendedor})
- Sus propios clientes y pedidos
- Información de bodega/almacén general
Si pregunta por datos de OTROS vendedores o información confidencial de otros usuarios, responde: "Solo puedes ver tu propia información de ventas y tus clientes."
Cuando muestres datos de ventas, filtra SOLO los registros donde el vendedor sea ${p.vendedor}.`;
  } else if (p.type === 'ventas') {
    ctx = `RESTRICCIÓN: Este usuario puede ver ventas, producción, pedidos, almacén y bodega. No puede ver información contable ni financiera confidencial.`;
  } else if (p.type === 'admin') {
    ctx = `ACCESO TOTAL: Administrador del sistema. Puede ver toda la información.`;
  } else {
    ctx = `ACCESO COMPLETO: Puede ver toda la información operativa de la empresa.`;
  }

  return ctx;
}

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
    messages JSONB DEFAULT '[]', created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW(), UNIQUE(username, conv_id)
  )`);
  await pool.query(`CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL, role VARCHAR(50) DEFAULT 'user',
    active BOOLEAN DEFAULT true, created_at TIMESTAMP DEFAULT NOW()
  )`);
  await pool.query(`CREATE TABLE IF NOT EXISTS qad_pdfs (
    id SERIAL PRIMARY KEY,
    filename VARCHAR(500) UNIQUE NOT NULL,
    display_name VARCHAR(500),
    pdf_base64 TEXT NOT NULL,
    created_at TIMESTAMP DEFAULT NOW()
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
    ['ADMON_LUCY','Vivatex2026','user'],
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

async function tableExists(name) {
  if (!pool) return false;
  const r = await pool.query(`SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_schema='public' AND table_name=$1) AS exists`, [name]);
  return !!r.rows[0]?.exists;
}

// ─── MEMORIA ──────────────────────────────────────────────────────
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
    await pool.query(`INSERT INTO user_memory (username,memory,updated_at) VALUES ($1,$2,NOW()) ON CONFLICT (username) DO UPDATE SET memory=$2,updated_at=NOW()`, [username, memory]);
  } catch(e) {}
}

// ─── BÚSQUEDA PARCIAL DE NOMBRES ─────────────────────────────────
// Extrae todos los nombres únicos de clientes/vendedores de QAD
async function getAllNames() {
  if (!pool) return [];
  try {
    // Buscar en hojas de CxC y ventas los nombres únicos
    const r = await pool.query(`
      SELECT DISTINCT elem->>'Cliente' as nombre FROM qad_data, jsonb_array_elements(data) AS elem
      WHERE elem->>'Cliente' IS NOT NULL AND length(elem->>'Cliente') > 2
      UNION
      SELECT DISTINCT elem->>'Nombre' as nombre FROM qad_data, jsonb_array_elements(data) AS elem
      WHERE elem->>'Nombre' IS NOT NULL AND length(elem->>'Nombre') > 2
      UNION
      SELECT DISTINCT elem->>'CLIENTE' as nombre FROM qad_data, jsonb_array_elements(data) AS elem
      WHERE elem->>'CLIENTE' IS NOT NULL AND length(elem->>'CLIENTE') > 2
      LIMIT 500
    `);
    return r.rows.map(r => r.nombre).filter(Boolean);
  } catch(e) { return []; }
}

// Encuentra el nombre más parecido al que escribió el usuario
function fuzzyMatch(query, names) {
  if (!query || !names.length) return null;
  const q = norm(query);
  
  // Coincidencia exacta primero
  const exact = names.find(n => norm(n) === q);
  if (exact) return exact;
  
  // Coincidencia parcial — el query está contenido en el nombre o viceversa
  const partial = names.filter(n => {
    const nn = norm(n);
    return nn.includes(q) || q.includes(nn) || 
           q.split(' ').some(w => w.length > 3 && nn.includes(w));
  });
  
  if (partial.length === 1) return partial[0];
  if (partial.length > 1) {
    // Retornar el más corto (más específico)
    return partial.sort((a, b) => a.length - b.length)[0];
  }
  
  return null;
}

// ─── DETECCIÓN DE HOJAS RELEVANTES ───────────────────────────────
function detectRelevantSheets(userMessage, sources) {
  const msg = norm(userMessage);
  const mappings = [
    { keys: ['cartera','saldo','cliente','cxc','cobrar','deuda','vencido','antiguedad','adeudo','atrasado','moroso','cobranza','debe','deben','cuanto debe','factura pendiente'], patterns: ['cxc','saldo','cliente','cartera','cobrar','antiguedad'] },
    { keys: ['venta','factura','remision','ingreso','vendedor','pedido venta'], patterns: ['venta','factura','remision','ingreso','pedido'] },
    { keys: ['inventario','stock','existencia','almacen','metros disponibles','bodega','tela acabada'], patterns: ['inventario','stock','existencia','almacen'] },
    { keys: ['proveedor','cxp','pagar proveedor','compra'], patterns: ['proveedor','cxp','pagar','compra'] },
    { keys: ['produccion','tejido metros','acabado metros','manufactura'], patterns: ['produccion','manufactura','tejido','acabado'] },
    { keys: ['especificacion','ficha tecnica','gramaje','composicion'], patterns: ['especificacion','ficha','tecnica'] },
  ];

  let relevantPatterns = [];
  for (const map of mappings) {
    if (map.keys.some(k => msg.includes(k))) {
      relevantPatterns = map.patterns;
      break;
    }
  }

  if (!relevantPatterns.length) return sources;
  const relevant = sources.filter(src => {
    const n = norm(`${src.filename} ${src.sheet_name}`);
    return relevantPatterns.some(p => n.includes(p));
  });
  return relevant.length > 0 ? relevant : sources;
}

// ─── BÚSQUEDA QAD ─────────────────────────────────────────────────
async function searchQAD(userMessage, maxRows = 300) {
  if (!pool) return { rows: [], total: 0, sourceInfo: '' };

  try {
    const allSources = await pool.query(
      `SELECT sheet_key, filename, sheet_name, updated_at, jsonb_array_length(data) as count FROM qad_data ORDER BY updated_at DESC`
    );
    if (!allSources.rows.length) return { rows: [], total: 0, sourceInfo: '' };

    const total = allSources.rows.reduce((s, r) => s + (parseInt(r.count) || 0), 0);
    const relevantSources = detectRelevantSheets(userMessage, allSources.rows);

    // Búsqueda parcial de nombres
    const allNames = await getAllNames();
    const msgWords = norm(userMessage).split(' ').filter(w => w.length > 3);
    let resolvedNames = [];
    for (const word of msgWords) {
      const match = fuzzyMatch(word, allNames);
      if (match) resolvedNames.push(match);
    }
    if (resolvedNames.length) {
      console.log(`Nombres resueltos: ${resolvedNames.join(', ')}`);
    }

    let allMatches = [];

    for (const src of relevantSources) {
      const srcCount = parseInt(src.count) || 0;
      const limit = srcCount <= 500 ? srcCount : 250;

      const r = await pool.query(
        `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT $2`,
        [src.sheet_key, limit]
      );
      r.rows.forEach(row => allMatches.push({ ...row.row, _src: `${src.filename}/${src.sheet_name}` }));
    }

    // Si hay nombres resueltos, filtrar primero por esos nombres
    if (resolvedNames.length && allMatches.length) {
      const filtered = allMatches.filter(row => {
        const rowText = norm(JSON.stringify(row));
        return resolvedNames.some(name => rowText.includes(norm(name)));
      });
      if (filtered.length > 0) allMatches = filtered;
    }

    const sourceInfo = relevantSources.map(s =>
      `• ${s.filename} (${s.sheet_name}) — ${s.count} registros`
    ).join('\n');

    return { rows: allMatches.slice(0, maxRows), total, sourceInfo, resolvedNames };
  } catch(e) {
    console.error('searchQAD error:', e.message);
    return { rows: [], total: 0, sourceInfo: '' };
  }
}

// ─── CONTEXTO QAD ─────────────────────────────────────────────────
async function buildQADContext(userMessage = '') {
  try {
    if (!pool || !(await tableExists('qad_data'))) return { hasData: false, text: 'No hay datos QAD.' };
    const countR = await pool.query('SELECT COUNT(*) as c FROM qad_data');
    if (parseInt(countR.rows[0].c) === 0) return { hasData: false, text: 'No hay archivos QAD cargados.' };

    const { rows, total, sourceInfo, resolvedNames } = await searchQAD(userMessage, 300);
    if (!rows.length) return { hasData: false, text: 'No encontré datos para esta consulta.' };

    const bySrc = {};
    rows.forEach(row => {
      const src = row._src || 'QAD';
      if (!bySrc[src]) bySrc[src] = [];
      const clean = { ...row };
      delete clean._src;
      bySrc[src].push(clean);
    });

    let ctx = `DATOS QAD — ${total.toLocaleString('es-MX')} registros totales\n`;
    if (resolvedNames?.length) ctx += `Búsqueda resuelta para: ${resolvedNames.join(', ')}\n`;
    ctx += `Fuentes:\n${sourceInfo}\n\n`;
    ctx += `REGISTROS (${rows.length}):\n\n`;

    Object.keys(bySrc).forEach(src => {
      const srcRows = bySrc[src];
      if (!srcRows.length) return;
      const cols = Object.keys(srcRows[0]);
      ctx += `=== ${src} ===\n`;
      ctx += cols.join(' | ') + '\n';
      srcRows.forEach(row => { ctx += cols.map(c => fmt(row[c])).join(' | ') + '\n'; });
      ctx += '\n';
    });

    if (total > rows.length) ctx += `\n(Mostrando ${rows.length} de ${total} registros.)`;
    return { hasData: true, text: ctx, shown: rows.length, total };
  } catch(e) {
    console.error('buildQADContext error:', e.message);
    return { hasData: false, text: 'Error cargando datos QAD.' };
  }
}

// ─── ALERTAS DE CARTERA VENCIDA ───────────────────────────────────
async function getAlertas() {
  const alertas = [];
  if (!pool) return alertas;

  try {
    // Buscar hojas de CxC
    const sources = await pool.query(
      `SELECT sheet_key FROM qad_data WHERE lower(filename) LIKE '%cxc%' OR lower(filename) LIKE '%cartera%' OR lower(sheet_name) LIKE '%saldo%' OR lower(sheet_name) LIKE '%cliente%'`
    );

    for (const src of sources.rows) {
      const r = await pool.query(
        `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT 500`,
        [src.sheet_key]
      );

      for (const { row } of r.rows) {
        // Buscar columnas de saldo/vencido
        const keys = Object.keys(row);
        const saldoKey = keys.find(k => /saldo|vencido|importe|total/i.test(k));
        const clienteKey = keys.find(k => /cliente|nombre|razon/i.test(k));
        const diasKey = keys.find(k => /dias|días|vencimiento/i.test(k));

        if (!saldoKey || !clienteKey) continue;

        const saldo = parseFloat(String(row[saldoKey] || '').replace(/[$,\s]/g, '')) || 0;
        const dias = parseInt(row[diasKey] || '0') || 0;
        const cliente = String(row[clienteKey] || '').trim();

        if (!cliente) continue;

        // Alerta si saldo vencido > 50,000 o días > 60
        if (saldo > 50000 || dias > 60) {
          alertas.push({
            tipo: dias > 90 ? 'critico' : dias > 60 ? 'alto' : 'medio',
            cliente,
            saldo: saldo.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }),
            dias,
            mensaje: dias > 90
              ? `${cliente} tiene ${dias} días vencidos — REQUIERE ATENCIÓN INMEDIATA`
              : `${cliente} tiene saldo vencido de ${saldo.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}`
          });
        }
      }
    }

    // Ordenar por criticidad
    alertas.sort((a, b) => {
      const order = { critico: 0, alto: 1, medio: 2 };
      return (order[a.tipo] || 3) - (order[b.tipo] || 3);
    });

  } catch(e) {
    console.error('Error alertas:', e.message);
  }

  return alertas.slice(0, 20);
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
  await pool.query(`INSERT INTO qad_data (sheet_key,filename,sheet_name,data,updated_at) VALUES ($1,$2,$3,$4,NOW()) ON CONFLICT (sheet_key) DO UPDATE SET filename=$2,sheet_name=$3,data=$4,updated_at=NOW()`,
    [key, filename, sheet, JSON.stringify(data)]);
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
        const rawText = String(p.text || '').trim();

        // Guardar PDF original en base64 para descarga directa
        const pdfBase64 = file.buffer.toString('base64');
        const displayName = path.basename(file.originalname, '.pdf')
          .replace(/^FT[_\s-]*/i, '')
          .replace(/[_-]/g, ' ')
          .trim()
          .toUpperCase();

        if (pool) {
          await pool.query(
            `INSERT INTO qad_pdfs (filename, display_name, pdf_base64, created_at)
             VALUES ($1, $2, $3, NOW())
             ON CONFLICT (filename) DO UPDATE SET display_name=$2, pdf_base64=$3, created_at=NOW()`,
            [file.originalname, displayName, pdfBase64]
          );
        }

        // También guardar texto para búsquedas
        const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 2);
        let data = [];

        if (lines.length > 0) {
          const tabularLines = lines.filter(l => /\s{2,}/.test(l));
          const isTabular = tabularLines.length > lines.length * 0.3 && lines.length > 3;

          if (isTabular) {
            let headerIdx = 0;
            for (let i = 0; i < Math.min(10, lines.length); i++) {
              if (/\s{2,}/.test(lines[i])) { headerIdx = i; break; }
            }
            const headers = lines[headerIdx].split(/\s{2,}/).map(h => h.trim()).filter(Boolean);
            if (headers.length >= 2) {
              for (let i = headerIdx + 1; i < lines.length; i++) {
                const parts = lines[i].split(/\s{2,}/).map(p2 => p2.trim());
                if (parts.length >= 2) {
                  const row = {};
                  headers.forEach((h, idx) => { row[h] = parts[idx] || ''; });
                  data.push(row);
                }
              }
            }
          }

          if (!data.length) {
            data = lines.map((l, i) => ({ linea: i + 1, contenido: l, archivo: file.originalname }));
          }
        }

        if (data.length) {
          const key = `${file.originalname}_PDF`.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 200);
          await saveQAD(key, file.originalname, 'PDF', data);
        }

        console.log(`✅ PDF: ${file.originalname} — guardado (${Math.round(pdfBase64.length/1024)}KB)`);
        sheets++; files++;
      } catch(e) { console.error('Error PDF:', file.originalname, e.message); }
    }
  }
  res.json({ ok: true, files, sheets });
});

// ─── ENDPOINT PDF ────────────────────────────────────────────────
// Buscar PDF por nombre de tela
app.get('/api/pdf/search/:query', async (req, res) => {
  if (!pool) return res.status(404).json({ found: false });
  try {
    const query = req.params.query.trim();
    const normQuery = query.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9\s]/g, ' ').trim();
    
    // Buscar por nombre display o filename
    const r = await pool.query(
      `SELECT id, filename, display_name FROM qad_pdfs 
       WHERE lower(display_name) LIKE $1 OR lower(filename) LIKE $1
       ORDER BY length(display_name) ASC LIMIT 5`,
      [`%${normQuery}%`]
    );

    if (!r.rows.length) return res.json({ found: false, query });

    // Tomar el más corto (más específico)
    const best = r.rows[0];
    res.json({ found: true, id: best.id, filename: best.filename, displayName: best.display_name });
  } catch(e) {
    console.error('Error buscando PDF:', e.message);
    res.status(500).json({ found: false });
  }
});

// Servir PDF por ID
app.get('/api/pdf/:id', async (req, res) => {
  if (!pool) return res.status(404).send('Sin base de datos');
  try {
    const r = await pool.query('SELECT filename, display_name, pdf_base64 FROM qad_pdfs WHERE id=$1', [req.params.id]);
    if (!r.rows.length) return res.status(404).send('PDF no encontrado');
    
    const { filename, pdf_base64 } = r.rows[0];
    const buf = Buffer.from(pdf_base64, 'base64');
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `inline; filename="${filename}"`,
      'Content-Length': buf.length
    });
    res.send(buf);
  } catch(e) {
    console.error('Error sirviendo PDF:', e.message);
    res.status(500).send('Error');
  }
});

// Listar todos los PDFs disponibles
app.get('/api/pdfs/list', async (req, res) => {
  if (!pool) return res.json({ pdfs: [] });
  try {
    const r = await pool.query('SELECT id, filename, display_name, created_at FROM qad_pdfs ORDER BY display_name ASC');
    res.json({ pdfs: r.rows });
  } catch(e) {
    res.json({ pdfs: [] });
  }
});

app.get('/api/qad/status', async (req, res) => {
  if (!pool) return res.json({ hasData: false, totalRecords: 0, sheets: [] });
  try {
    const r = await pool.query('SELECT sheet_key,filename,sheet_name,jsonb_array_length(data) as count,updated_at FROM qad_data ORDER BY updated_at DESC');
    const total = r.rows.reduce((s, row) => s + (parseInt(row.count) || 0), 0);
    res.json({ hasData: total > 0, totalRecords: total, sheets: r.rows });
  } catch(e) { res.json({ hasData: false, totalRecords: 0, sheets: [] }); }
});

app.delete('/api/qad/clear', async (req, res) => {
  if (pool) await pool.query('DELETE FROM qad_data');
  res.json({ ok: true });
});

// ─── ALERTAS ──────────────────────────────────────────────────────
app.get('/api/alertas', async (req, res) => {
  const alertas = await getAlertas();
  res.json({ alertas });
});

// ─── PERMISOS ─────────────────────────────────────────────────────
app.get('/api/permisos/:username', (req, res) => {
  const p = getPerms(req.params.username);
  res.json({ perms: p });
});

// ─── USUARIOS ─────────────────────────────────────────────────────
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};
  if (!pool) return res.status(500).json({ ok: false, error: 'Sin DB' });
  try {
    const r = await pool.query('SELECT id,username,role,active FROM usuarios WHERE UPPER(username)=$1 AND password=$2',
      [String(username || '').toUpperCase(), password]);
    if (!r.rows.length) return res.json({ ok: false, error: 'Usuario o contraseña incorrectos' });
    if (!r.rows[0].active) return res.json({ ok: false, error: 'Usuario inactivo' });
    const perms = getPerms(r.rows[0].username);
    res.json({ ok: true, user: { ...r.rows[0], hasPanel: perms.panel, perms } });
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
    await pool.query(`INSERT INTO usuarios (username,password,role,active) VALUES ($1,$2,$3,true) ON CONFLICT (username) DO UPDATE SET password=$2,role=$3,active=true`,
      [username.toUpperCase(), password, role || 'user']);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: false, error: e.message }); }
});

app.delete('/api/usuarios/:username', async (req, res) => {
  if (!pool) return res.json({ ok: false });
  try { await pool.query('UPDATE usuarios SET active=false WHERE UPPER(username)=$1', [req.params.username.toUpperCase()]); res.json({ ok: true }); }
  catch(e) { res.json({ ok: false }); }
});

app.get('/api/memory/:username', async (req, res) => {
  res.json({ memory: await getUserMemory(req.params.username) });
});
app.post('/api/memory/:username', async (req, res) => {
  await saveUserMemory(req.params.username, req.body?.memory || '');
  res.json({ ok: true });
});

// ─── CHAT ─────────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages, username } = req.body || {};
  if (!ANTHROPIC_API_KEY) return res.json({ reply: '⚠️ Falta ANTHROPIC_API_KEY en Railway.' });

  try {
    const lastUserMsg = (messages || []).filter(m => m.role === 'user').slice(-1)[0]?.content || '';
    const perms = getPerms(username);

    // Verificar restricciones ANTES de llamar a Claude
    const msgNorm = norm(lastUserMsg);
    if (perms.restrict?.length) {
      const blocked = perms.restrict.some(r => msgNorm.includes(norm(r)));
      if (blocked) {
        return res.json({ reply: '🚫 No tienes acceso a esa información. Si necesitas datos contables, contacta al área de Contabilidad.' });
      }
    }

    if (perms.type === 'vendedor' && perms.allow?.length) {
      const allowed = perms.allow.some(a => msgNorm.includes(norm(a)));
      const isGeneral = ['hola','ayuda','que puedes','como','gracias'].some(g => msgNorm.includes(g));
      if (!allowed && !isGeneral && msgNorm.length > 10) {
        // No bloquear, pero sí filtrar — dejar que Claude maneje con el contexto
      }
    }

    const [qadCtx, userMemory] = await Promise.all([
      buildQADContext(lastUserMsg),
      getUserMemory(username || '')
    ]);

    const permContext = buildPermissionContext(username);

    const systemPrompt = `Eres AVIVA, analista IA de Grupo Vivatex S.A. de C.V. NUNCA menciones Claude ni Anthropic.

${permContext}

REGLAS:
1. NUNCA inventes datos — SOLO usa registros QAD del contexto
2. Si no encuentras el dato: "No encontré [X] en los datos disponibles."
3. SIEMPRE usa tabla Markdown para datos tabulares
4. Copia nombres y números EXACTAMENTE como aparecen
5. Ordena de mayor a menor por saldo/importe
6. Incluye fila TOTAL al final de tablas numéricas
7. Después de la tabla: 2-3 líneas de análisis ejecutivo

FORMATO TABLA:
| Cliente | Saldo | Días |
|---------|-------|------|
| NOMBRE | $1,234 | 90 |
| **TOTAL** | **$X,XXX** | |

PARA EXCEL al final:
%%EXCEL%%{"titulo":"Nombre","hojas":[{"nombre":"Hoja","columnas":["Col1","Col2"],"filas":[["val",123]],"totales":["TOTAL",123]}]}%%EXCEL%%

${userMemory ? `MEMORIA:\n${userMemory}\n` : ''}

DATOS QAD:
${qadCtx.text}

Usuario: ${username} | ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}`;

    let cleanMessages = (messages || [])
      .filter(m => (m.role === 'user' || m.role === 'assistant') && String(m.content || '').trim())
      .slice(-6)
      .map(m => ({ role: m.role, content: String(m.content || '').substring(0, 1500) }));

    while (cleanMessages.length && cleanMessages[0].role !== 'user') cleanMessages.shift();
    if (!cleanMessages.length) cleanMessages = [{ role: 'user', content: lastUserMsg || 'Hola' }];

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 2048, system: systemPrompt, messages: cleanMessages })
    });

    const data = await response.json();
    if (!response.ok) {
      if (response.status === 429) return res.json({ reply: '⏳ AVIVA está muy ocupada. Espera 15 segundos e intenta de nuevo.' });
      return res.json({ reply: `Error: ${data?.error?.message || response.status}` });
    }

    let reply = data.content?.[0]?.text || 'Sin respuesta.';
    let excelData = null;
    const excelMatch = reply.match(/%%EXCEL%%([\s\S]*?)%%EXCEL%%/);
    if (excelMatch) {
      try { excelData = JSON.parse(excelMatch[1].trim()); reply = reply.replace(/%%EXCEL%%[\s\S]*?%%EXCEL%%/, '').trim(); }
      catch(e) {}
    }

    const memTriggers = ['recuerda','mi nombre','llamame','soy el','soy la','trabajo en','mi area'];
    if (memTriggers.some(t => msgNorm.includes(t))) {
      const cur = await getUserMemory(username || '');
      await saveUserMemory(username || '', (cur + `\n[${new Date().toLocaleDateString('es-MX')}] ${lastUserMsg.substring(0, 150)}`).substring(0, 1500).trim());
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
function generateProfessionalExcel(data) {
  const wb = xlsx.utils.book_new();
  const { titulo = 'Reporte', subtitulo = '', periodo = '', usuario = '', hojas = [] } = data;
  hojas.forEach(hoja => {
    const { nombre = 'Datos', columnas = [], filas = [], totales = null } = hoja;
    const wsData = [];
    wsData.push([`GRUPO VIVATEX S.A. DE C.V.  ·  ${titulo.toUpperCase()}`]);
    wsData.push([`${subtitulo || nombre}   |   ${periodo || new Date().toLocaleDateString('es-MX')}   |   Usuario: ${usuario}   |   ${new Date().toLocaleString('es-MX')}`]);
    wsData.push([]);
    wsData.push(columnas);
    filas.forEach(fila => wsData.push(fila));
    if (totales) { wsData.push([]); wsData.push(totales); }
    const ws = xlsx.utils.aoa_to_sheet(wsData);
    const colWidths = columnas.map((col, i) => {
      let maxW = String(col).length + 4;
      filas.forEach(fila => { const l = String(fila[i] ?? '').length; if (l > maxW) maxW = l; });
      return { wch: Math.min(Math.max(maxW + 2, 12), 45) };
    });
    ws['!cols'] = colWidths;
    if (columnas.length > 1) {
      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: columnas.length - 1 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: columnas.length - 1 } },
      ];
    }
    xlsx.utils.book_append_sheet(wb, ws, String(nombre).slice(0, 31));
  });
  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.post('/api/excel/generate', async (req, res) => {
  try {
    const buf = generateProfessionalExcel(req.body || {});
    res.json({ ok: true, base64: buf.toString('base64'), filename: `Vivatex_${String(req.body?.titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx` });
  } catch(e) { res.json({ ok: false, error: e.message }); }
});

app.get('/api/health', async (req, res) => res.json({ ok: true, db: !!pool, model: ANTHROPIC_MODEL }));
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
