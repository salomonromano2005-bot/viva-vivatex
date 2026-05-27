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
    messages JSONB DEFAULT '[]', created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW(), UNIQUE(username, conv_id)
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

// ─── DETECCIÓN DE HOJAS RELEVANTES ───────────────────────────────
function detectRelevantSheets(userMessage, sources) {
  const msg = norm(userMessage);
  const mappings = [
    { keys: ['cartera','saldo','cliente','cxc','cobrar','deuda','vencido','antiguedad','adeudo','pago atrasado','atrasado','moroso','cobranza','debe','deben'], patterns: ['cxc','saldo','cliente','cartera','cobrar','antiguedad'] },
    { keys: ['venta','factura','remision','ingreso','pedido vendido'], patterns: ['venta','factura','remision','ingreso','pedido'] },
    { keys: ['inventario','stock','existencia','almacen','metros disponibles','disponible'], patterns: ['inventario','stock','existencia','almacen'] },
    { keys: ['proveedor','cxp','pagar proveedor','compra','abastecimiento'], patterns: ['proveedor','cxp','pagar','compra'] },
    { keys: ['produccion','manufactura','tejido metros','acabado metros','fabricacion'], patterns: ['produccion','manufactura','tejido','acabado'] },
    { keys: ['especificacion','ficha tecnica','gramaje','composicion','tela tecnica'], patterns: ['especificacion','ficha','tecnica'] },
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

// ─── BÚSQUEDA QAD INTELIGENTE ─────────────────────────────────────
async function searchQAD(userMessage, maxRows = 300) {
  if (!pool) return { rows: [], total: 0, sources: [], sourceInfo: '' };

  try {
    const allSources = await pool.query(
      `SELECT sheet_key, filename, sheet_name, updated_at, jsonb_array_length(data) as count FROM qad_data ORDER BY updated_at DESC`
    );
    if (!allSources.rows.length) return { rows: [], total: 0, sources: [], sourceInfo: '' };

    const total = allSources.rows.reduce((s, r) => s + (parseInt(r.count) || 0), 0);

    // Detectar hojas relevantes
    const relevantSources = detectRelevantSheets(userMessage, allSources.rows);
    console.log(`Hojas relevantes (${relevantSources.length}): ${relevantSources.map(s => s.filename).join(', ')}`);

    let allMatches = [];

    for (const src of relevantSources) {
      // Límite por hoja según tamaño — hojas pequeñas van completas
      const srcCount = parseInt(src.count) || 0;
      const limit = srcCount <= 500 ? srcCount : 200;

      const r = await pool.query(
        `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT $2`,
        [src.sheet_key, limit]
      );
      r.rows.forEach(row => allMatches.push({ ...row.row, _src: `${src.filename}/${src.sheet_name}` }));
    }

    // Info de fuentes para el contexto
    const sourceInfo = relevantSources.map(s =>
      `• ${s.filename} (${s.sheet_name}) — ${s.count} registros`
    ).join('\n');

    return { rows: allMatches.slice(0, maxRows), total, sources: relevantSources, sourceInfo };
  } catch(e) {
    console.error('searchQAD error:', e.message);
    return { rows: [], total: 0, sources: [], sourceInfo: '' };
  }
}

// ─── CONSTRUIR CONTEXTO QAD COMPACTO ─────────────────────────────
// Convierte los datos en texto compacto para no exceder tokens
async function buildQADContext(userMessage = '') {
  try {
    if (!pool || !(await tableExists('qad_data'))) {
      return { hasData: false, text: 'No hay datos QAD cargados.' };
    }

    const countR = await pool.query('SELECT COUNT(*) as c FROM qad_data');
    if (parseInt(countR.rows[0].c) === 0) {
      return { hasData: false, text: 'No hay archivos QAD. El administrador debe subirlos desde Panel de Sistemas → Archivos QAD.' };
    }

    const { rows, total, sourceInfo } = await searchQAD(userMessage, 300);

    if (!rows.length) {
      return { hasData: false, text: 'No encontré datos para esta consulta en QAD.' };
    }

    // Agrupar por fuente
    const bySrc = {};
    rows.forEach(row => {
      const src = row._src || 'QAD';
      if (!bySrc[src]) bySrc[src] = [];
      const clean = { ...row };
      delete clean._src;
      bySrc[src].push(clean);
    });

    let ctx = `DATOS QAD — ${total.toLocaleString('es-MX')} registros totales\n`;
    ctx += `Fuentes consultadas:\n${sourceInfo}\n\n`;
    ctx += `REGISTROS PARA ESTA CONSULTA (${rows.length}):\n\n`;

    Object.keys(bySrc).forEach(src => {
      const srcRows = bySrc[src];
      if (!srcRows.length) return;
      const cols = Object.keys(srcRows[0]);
      ctx += `=== ${src} ===\n`;
      ctx += cols.join(' | ') + '\n';
      srcRows.forEach(row => {
        ctx += cols.map(c => fmt(row[c])).join(' | ') + '\n';
      });
      ctx += '\n';
    });

    if (total > rows.length) {
      ctx += `\n(Mostrando ${rows.length} de ${total} registros. Para ver más, especifica cliente, fecha o período exacto.)`;
    }

    return { hasData: true, text: ctx, shown: rows.length, total };
  } catch(e) {
    console.error('buildQADContext error:', e.message);
    return { hasData: false, text: 'Error cargando datos QAD.' };
  }
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
        const data = String(p.text || '').split('\n').filter(Boolean).map((x, i) => ({ linea: i + 1, texto: x }));
        const key = `${file.originalname}_PDF`.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 200);
        await saveQAD(key, file.originalname, 'PDF', data);
        sheets++; files++;
      } catch(e) { console.error('Error PDF:', e.message); }
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
    const r = await pool.query('SELECT id,username,role,active FROM usuarios WHERE UPPER(username)=$1 AND password=$2',
      [String(username || '').toUpperCase(), password]);
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
    await pool.query(`INSERT INTO usuarios (username,password,role,active) VALUES ($1,$2,$3,true) ON CONFLICT (username) DO UPDATE SET password=$2,role=$3,active=true`,
      [username.toUpperCase(), password, role || 'user']);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: false, error: e.message }); }
});

app.delete('/api/usuarios/:username', async (req, res) => {
  if (!pool) return res.json({ ok: false });
  try {
    await pool.query('UPDATE usuarios SET active=false WHERE UPPER(username)=$1', [req.params.username.toUpperCase()]);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: false }); }
});

// ─── MEMORIA ──────────────────────────────────────────────────────
app.get('/api/memory/:username', async (req, res) => {
  const mem = await getUserMemory(req.params.username);
  res.json({ memory: mem });
});

app.post('/api/memory/:username', async (req, res) => {
  await saveUserMemory(req.params.username, req.body?.memory || '');
  res.json({ ok: true });
});

// ─── CHAT CON CLAUDE ──────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages, username } = req.body || {};
  if (!ANTHROPIC_API_KEY) return res.json({ reply: '⚠️ Falta ANTHROPIC_API_KEY en Railway.' });

  try {
    const lastUserMsg = (messages || []).filter(m => m.role === 'user').slice(-1)[0]?.content || '';

    // Cargar QAD y memoria en paralelo
    const [qadCtx, userMemory] = await Promise.all([
      buildQADContext(lastUserMsg),
      getUserMemory(username || '')
    ]);

    // Sistema compacto para no exceder tokens
    const systemPrompt = `Eres AVIVA, analista IA de Grupo Vivatex S.A. de C.V. NUNCA menciones Claude ni Anthropic.

REGLAS:
1. NUNCA inventes datos — SOLO usa los registros QAD que recibes
2. Si el cliente/dato no aparece en el contexto: "No encontré [X] en los datos disponibles. Verifica en QAD."
3. SIEMPRE responde con tabla Markdown cuando hay datos tabulares
4. Copia nombres y números EXACTAMENTE como aparecen
5. Ordena de mayor a menor por saldo/importe
6. Incluye fila TOTAL al final de tablas numéricas
7. Después de la tabla: 2-3 líneas de análisis ejecutivo

FORMATO TABLA:
| Cliente | Saldo | Días |
|---------|-------|------|
| NOMBRE EXACTO | $1,234 | 90 |
| **TOTAL** | **$X,XXX** | |

PARA EXCEL — al final agrega:
%%EXCEL%%{"titulo":"Nombre","hojas":[{"nombre":"Hoja","columnas":["Col1","Col2"],"filas":[["val",123]],"totales":["TOTAL",123]}]}%%EXCEL%%

${userMemory ? `MEMORIA USUARIO:\n${userMemory}\n` : ''}

DATOS QAD:
${qadCtx.text}

Usuario: ${username || 'Usuario'} | ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}`;

    // Mensajes limpios — máximo 6 turnos para no exceder tokens
    let cleanMessages = (messages || [])
      .filter(m => (m.role === 'user' || m.role === 'assistant') && String(m.content || '').trim())
      .slice(-6)
      .map(m => ({ role: m.role, content: String(m.content || '').substring(0, 1500) }));

    while (cleanMessages.length && cleanMessages[0].role !== 'user') cleanMessages.shift();
    if (!cleanMessages.length) cleanMessages = [{ role: 'user', content: lastUserMsg || 'Hola' }];

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: ANTHROPIC_MODEL,
        max_tokens: 2048,
        system: systemPrompt,
        messages: cleanMessages
      })
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Anthropic error:', data?.error?.message);
      // Si es rate limit, esperar y dar mensaje amigable
      if (response.status === 429) {
        return res.json({ reply: '⏳ AVIVA está procesando muchas solicitudes en este momento. Por favor espera 10 segundos e intenta de nuevo.' });
      }
      return res.json({ reply: `Error: ${data?.error?.message || response.status}` });
    }

    let reply = data.content?.[0]?.text || 'Sin respuesta.';

    // Extraer Excel si hay
    let excelData = null;
    const excelMatch = reply.match(/%%EXCEL%%([\s\S]*?)%%EXCEL%%/);
    if (excelMatch) {
      try {
        excelData = JSON.parse(excelMatch[1].trim());
        reply = reply.replace(/%%EXCEL%%[\s\S]*?%%EXCEL%%/, '').trim();
      } catch(e) {}
    }

    // Guardar memoria si el usuario comparte info personal
    const memTriggers = ['recuerda','mi nombre','llámame','soy el','soy la','trabajo en','mi área','soy gerente','soy director'];
    if (memTriggers.some(t => norm(lastUserMsg).includes(norm(t)))) {
      const currentMem = await getUserMemory(username || '');
      const newMem = (currentMem + `\n[${new Date().toLocaleDateString('es-MX')}] ${lastUserMsg.substring(0, 150)}`).substring(0, 1500);
      await saveUserMemory(username || '', newMem.trim());
    }

    res.json({ reply, excelData });

  } catch(e) {
    console.error('Error /api/chat:', e.message);
    res.status(500).json({ reply: 'Error interno. Intenta de nuevo en unos segundos.' });
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
  try {
    await pool.query('DELETE FROM conversations WHERE username=$1 AND conv_id=$2', [req.params.username, req.params.convId]);
    res.json({ ok: true });
  } catch(e) { res.json({ ok: true }); }
});

// ─── EXCEL PROFESIONAL ────────────────────────────────────────────
function generateProfessionalExcel(data) {
  const wb = xlsx.utils.book_new();
  const { titulo = 'Reporte', subtitulo = '', periodo = '', usuario = '', hojas = [] } = data;

  hojas.forEach(hoja => {
    const { nombre = 'Datos', columnas = [], filas = [], totales = null } = hoja;
    const wsData = [];

    // Fila 1: Header empresa
    wsData.push([`GRUPO VIVATEX S.A. DE C.V.  ·  ${titulo.toUpperCase()}`]);
    // Fila 2: Meta
    wsData.push([`${subtitulo || nombre}   |   ${periodo || new Date().toLocaleDateString('es-MX')}   |   Usuario: ${usuario}   |   Generado: ${new Date().toLocaleString('es-MX')}`]);
    // Fila 3: Espacio
    wsData.push([]);
    // Fila 4: Headers
    wsData.push(columnas);
    // Filas de datos
    filas.forEach(fila => wsData.push(fila));
    // Fila totales
    if (totales) { wsData.push([]); wsData.push(totales); }

    const ws = xlsx.utils.aoa_to_sheet(wsData);

    // Anchos automáticos
    const colWidths = columnas.map((col, i) => {
      let maxW = String(col).length + 4;
      filas.forEach(fila => { const l = String(fila[i] ?? '').length; if (l > maxW) maxW = l; });
      return { wch: Math.min(Math.max(maxW + 2, 12), 45) };
    });
    ws['!cols'] = colWidths;

    // Merge header empresa
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
