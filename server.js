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
    filename VARCHAR(500), sheet_name VARCHAR(255), data JSONB, updated_at TIMESTAMP DEFAULT NOW()
  )`);
  await pool.query(`CREATE TABLE IF NOT EXISTS conversations (
    id SERIAL PRIMARY KEY, username VARCHAR(100) NOT NULL, conv_id VARCHAR(100) NOT NULL,
    title VARCHAR(500), messages JSONB DEFAULT '[]',
    created_at TIMESTAMP DEFAULT NOW(), updated_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(username, conv_id)
  )`);
  await pool.query(`CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL, role VARCHAR(50) DEFAULT 'user',
    active BOOLEAN DEFAULT true, created_at TIMESTAMP DEFAULT NOW()
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

function fmt(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return v.toLocaleString('es-MX', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  return String(v).replace(/\n/g, ' ').trim();
}

function norm(v) {
  return String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
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

// ─── BÚSQUEDA INTELIGENTE EN POSTGRESQL ───────────────────────────
// Busca directamente en JSONB usando PostgreSQL — mucho más eficiente
async function searchQADPostgres(userMessage, maxRows = 200) {
  if (!pool) return { rows: [], total: 0, sources: [] };

  const msg = norm(userMessage);
  
  // Detectar palabras clave importantes (ignorar palabras comunes)
  const stopWords = new Set(['dame','reporte','tabla','de','del','la','el','los','las','un','una','en','que','tienes','por','para','con','me','mi','mis','tu','sus','hay','son','mas','top','cinco','diez','todos','todas','cuales','cual','como','cuando','donde']);
  const keywords = msg.split(' ').filter(w => w.length > 2 && !stopWords.has(w));

  try {
    // Obtener info de todas las hojas disponibles
    const sheetsInfo = await pool.query('SELECT sheet_key, filename, sheet_name, updated_at, jsonb_array_length(data) as count FROM qad_data ORDER BY updated_at DESC');
    
    let totalRows = 0;
    sheetsInfo.rows.forEach(s => { totalRows += parseInt(s.count) || 0; });

    const sources = sheetsInfo.rows.map(s => ({
      key: s.sheet_key,
      filename: s.filename,
      sheet: s.sheet_name,
      count: s.count,
      updated: s.updated_at
    }));

    // Si no hay keywords específicas, traer muestra representativa de cada hoja
    if (keywords.length === 0) {
      let allRows = [];
      for (const src of sources) {
        const r = await pool.query(
          `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT 50`,
          [src.key]
        );
        r.rows.forEach(row => allRows.push({ ...row.row, _source: src.filename + '/' + src.sheet }));
      }
      return { rows: allRows.slice(0, maxRows), total: totalRows, sources };
    }

    // Buscar con keywords en JSONB — búsqueda de texto completo
    let allMatches = [];
    
    for (const src of sources) {
      // Construir query de búsqueda JSONB
      const searchConditions = keywords.map((_, i) => 
        `lower(data_row::text) LIKE $${i + 2}`
      ).join(' OR ');
      
      const query = `
        SELECT data_row as row
        FROM qad_data,
        jsonb_array_elements(data) AS data_row
        WHERE sheet_key = $1
        AND (${searchConditions})
        LIMIT $${keywords.length + 2}
      `;
      
      const params = [src.key, ...keywords.map(k => `%${k}%`), Math.ceil(maxRows / sources.length) + 20];
      
      try {
        const r = await pool.query(query, params);
        r.rows.forEach(row => {
          allMatches.push({ ...row.row, _source: src.filename + '/' + src.sheet });
        });
      } catch(e) {
        // Si falla la búsqueda JSONB, traer muestra
        const fallback = await pool.query(
          `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT 30`,
          [src.key]
        );
        fallback.rows.forEach(row => {
          allMatches.push({ ...row.row, _source: src.filename + '/' + src.sheet });
        });
      }
    }

    // Si no encontró nada con keywords, traer muestra general
    if (allMatches.length === 0) {
      for (const src of sources.slice(0, 3)) {
        const r = await pool.query(
          `SELECT jsonb_array_elements(data) as row FROM qad_data WHERE sheet_key = $1 LIMIT 60`,
          [src.key]
        );
        r.rows.forEach(row => allMatches.push({ ...row.row, _source: src.filename + '/' + src.sheet }));
      }
    }

    return { rows: allMatches.slice(0, maxRows), total: totalRows, sources };

  } catch(e) {
    console.error('Error búsqueda QAD:', e.message);
    return { rows: [], total: 0, sources: [] };
  }
}

async function buildQADContext(userMessage = '') {
  try {
    if (!pool) return { hasData: false, text: 'Sin conexión a base de datos.' };

    const exists = await tableExists('qad_data');
    if (!exists) return { hasData: false, text: 'No hay datos QAD cargados.' };

    const countResult = await pool.query('SELECT COUNT(*) as total FROM qad_data');
    if (parseInt(countResult.rows[0].total) === 0) {
      return { hasData: false, text: 'No hay datos QAD cargados. El administrador debe subir archivos Excel/CSV desde el Panel de Sistemas.' };
    }

    const { rows, total, sources } = await searchQADPostgres(userMessage, 200);

    if (!rows.length) return { hasData: false, text: 'No encontré datos relacionados con tu consulta en QAD.' };

    // Construir resumen de fuentes
    let contextText = `DATOS QAD — ${total} registros totales en sistema\n`;
    contextText += `Archivos disponibles:\n`;
    sources.forEach(s => {
      const fecha = s.updated ? new Date(s.updated).toLocaleString('es-MX') : 'N/A';
      contextText += `  • ${s.filename} / ${s.sheet} — ${s.count} registros — Actualizado: ${fecha}\n`;
    });
    contextText += `\nRegistros relevantes para esta consulta (${rows.length} de ${total}):\n\n`;

    // Agrupar por fuente para mejor presentación
    const bySrc = {};
    rows.forEach(row => {
      const src = row._source || 'QAD';
      if (!bySrc[src]) bySrc[src] = [];
      const cleanRow = { ...row };
      delete cleanRow._source;
      bySrc[src].push(cleanRow);
    });

    Object.keys(bySrc).forEach(src => {
      const srcRows = bySrc[src];
      if (!srcRows.length) return;
      contextText += `=== ${src} ===\n`;
      const cols = Object.keys(srcRows[0]);
      contextText += cols.join(' | ') + '\n';
      contextText += cols.map(() => '---').join(' | ') + '\n';
      srcRows.forEach(row => {
        contextText += cols.map(c => fmt(row[c])).join(' | ') + '\n';
      });
      contextText += '\n';
    });

    if (total > rows.length) {
      contextText += `\n⚠️ Mostrando ${rows.length} de ${total} registros totales. Para ver más, especifica el cliente, vendedor o período exacto.`;
    }

    return { hasData: true, text: contextText, total, shown: rows.length };

  } catch(e) {
    console.error('Error buildQADContext:', e.message);
    return { hasData: false, text: 'Error al cargar datos QAD: ' + e.message };
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
  await pool.query(`INSERT INTO qad_data (sheet_key,filename,sheet_name,data,updated_at) VALUES ($1,$2,$3,$4,NOW()) ON CONFLICT (sheet_key) DO UPDATE SET filename=$2,sheet_name=$3,data=$4,updated_at=NOW()`, [key, filename, sheet, JSON.stringify(data)]);
}

app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  let sheets = 0, files = 0;
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
        sheets++; files++;
      } catch(e) { console.error('PDF error:', e.message); }
    }
  }
  res.json({ ok: true, files, sheets });
});

app.get('/api/qad/status', async (req, res) => {
  if (!pool) return res.json({ hasData: false, totalRecords: 0, sheets: [] });
  try {
    const r = await pool.query('SELECT sheet_key, filename, sheet_name, jsonb_array_length(data) as count FROM qad_data');
    const total = r.rows.reduce((sum, row) => sum + (parseInt(row.count) || 0), 0);
    res.json({ hasData: total > 0, totalRecords: total, sheets: r.rows.map(r => r.sheet_key) });
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
  const r = await pool.query('SELECT id,username,role,active FROM usuarios WHERE UPPER(username)=$1 AND password=$2', [String(username || '').toUpperCase(), password]);
  if (!r.rows.length) return res.json({ ok: false, error: 'Usuario o contraseña incorrectos' });
  if (!r.rows[0].active) return res.json({ ok: false, error: 'Usuario inactivo' });
  res.json({ ok: true, user: r.rows[0] });
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

// ─── CHAT CON CLAUDE ──────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages, system, username } = req.body || {};

  if (!ANTHROPIC_API_KEY) {
    return res.json({ reply: '⚠️ Falta configurar ANTHROPIC_API_KEY en Railway.' });
  }

  try {
    const lastUserMsg = (messages || []).filter(m => m.role === 'user').slice(-1)[0]?.content || '';
    const qadCtx = await buildQADContext(lastUserMsg);

    const systemPrompt = `Eres AVIVA, la analista de inteligencia artificial de Grupo Vivatex S.A. de C.V.

IDENTIDAD: Tu nombre es AVIVA. Nunca menciones Claude, Anthropic ni OpenAI. Trabajas exclusivamente para Grupo Vivatex. Hablas español mexicano natural, tuteas al usuario, eres directa y profesional.

REGLAS ABSOLUTAS:
1. NUNCA inventes clientes, vendedores, productos ni cantidades
2. SOLO usa datos del contexto QAD que recibes
3. Si el dato no está disponible di: "Ese dato no está en la información que tengo. Verifica directamente en QAD."
4. NUNCA hagas gráficas ASCII con caracteres como █ ▓ ░
5. NUNCA muestres JSON crudo en tu respuesta
6. Muestra SIEMPRE el listado COMPLETO de los registros que tienes — nunca omitas filas
7. SIEMPRE responde con tablas Markdown cuando hay datos tabulares
8. COPIA EXACTO — nombres de clientes, vendedores, productos y códigos deben copiarse EXACTAMENTE como aparecen en los datos. NUNCA traduzcas, abrevies, corrijas ortografía ni cambies mayúsculas/minúsculas. Si dice "VENTAS_EDGAR" escribe "VENTAS_EDGAR", no "Edgar Zárate" ni ninguna variación.
9. NÚMEROS EXACTOS — copia los valores numéricos exactamente como están en los datos, sin redondear ni aproximar.

FORMATO OBLIGATORIO PARA DATOS:
Cuando tengas datos de clientes, ventas, cartera, inventario, etc., SIEMPRE usa tabla Markdown:

| Cliente | Saldo Total | Días Vencidos | Vendedor |
|---------|-------------|---------------|---------|
| EMPRESA SA DE CV | $1,234,567 | 90 | Juan Pérez |

- Ordena de MAYOR a MENOR por saldo o importe
- Incluye fila de TOTALES al final
- Usa **negritas** para los valores más importantes
- Usa ## para títulos de sección

DATOS QAD DISPONIBLES AHORA:
${qadCtx.hasData ? qadCtx.text : qadCtx.text}

Usuario: ${username || 'Usuario'}
Fecha: ${new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })}`;

    const cleanMessages = (messages || [])
      .filter(m => m.role === 'user' || m.role === 'assistant')
      .slice(-6)
      .map(m => ({ role: m.role, content: String(m.content || '').substring(0, 2000) }));

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

    const data = await response.json();

    if (!response.ok) {
      console.error('Anthropic error:', JSON.stringify(data));
      return res.json({ reply: `Error de API: ${data?.error?.message || response.status}` });
    }

    res.json({ reply: data.content?.[0]?.text || 'Sin respuesta.' });

  } catch (e) {
    console.error('Error /api/chat:', e.message);
    res.status(500).json({ reply: 'Error interno. Intenta de nuevo.' });
  }
});

// ─── CONVERSACIONES ───────────────────────────────────────────────
app.post('/api/conversations/save', async (req, res) => {
  const { username, convId, title, messages } = req.body || {};
  if (!pool || !username || !convId) return res.json({ ok: true });
  try {
    await pool.query(`INSERT INTO conversations (username,conv_id,title,messages,updated_at) VALUES ($1,$2,$3,$4,NOW()) ON CONFLICT (username,conv_id) DO UPDATE SET title=$3,messages=$4,updated_at=NOW()`, [username, convId, title || 'Conversación', JSON.stringify(messages || [])]);
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

// ─── EXCEL ────────────────────────────────────────────────────────
app.post('/api/excel/generate', async (req, res) => {
  const wb = xlsx.utils.book_new();
  const { titulo, hojas } = req.body || {};
  for (const h of hojas || []) {
    const ws = xlsx.utils.aoa_to_sheet([h.columnas || [], ...(h.filas || [])]);
    xlsx.utils.book_append_sheet(wb, ws, String(h.nombre || 'Datos').slice(0, 31));
  }
  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.json({ ok: true, base64: buf.toString('base64'), filename: `${String(titulo || 'Reporte').replace(/[^a-zA-Z0-9_]/g, '_')}.xlsx` });
});

// ─── HEALTH ───────────────────────────────────────────────────────
app.get('/api/health', async (req, res) => {
  res.json({ ok: true, db: !!pool, model: ANTHROPIC_MODEL, provider: 'anthropic' });
});

app.get('/manifest.json', (req, res) => res.json({ name: 'AVIVA', short_name: 'AVIVA', start_url: '/', display: 'standalone' }));

app.get('*', (req, res) => {
  const index = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(index)) return res.sendFile(index);
  res.send('AVIVA ONLINE');
});

// ─── INICIO ───────────────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE — Puerto ${PORT}`);
    console.log(`🤖 Modelo: ${ANTHROPIC_MODEL}`);
    console.log(`🔑 API Key: ${ANTHROPIC_API_KEY ? 'Configurada ✅' : 'FALTA ❌'}`);
  });
});
