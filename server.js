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

app.get('/api/health', async (req, res) => {
  res.json({
    ok: true,
    provider: AI_PROVIDER,
    openai: !!OPENAI_API_KEY,
    db: !!pool,
    model: OPENAI_MODEL,
    time: new Date().toISOString(),
  });
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

    res.json({
      ok: true,
      user: {
        id: user.id,
        username: user.username,
        role: user.role,
        active: user.active,
      },
    });
  } catch (e) {
    console.error('Error login:', e.message);
    res.status(500).json({ ok: false, error: 'Error del servidor' });
  }
});

app.get('/api/usuarios', async (req, res) => {
  try {
    if (!pool) return res.json({ users: [] });
    const result = await pool.query('SELECT id, username, role, active FROM usuarios ORDER BY id');
    res.json({ users: result.rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/usuarios', async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'Faltan datos' });

  try {
    if (!pool) return res.status(500).json({ error: 'Sin DB' });

    await pool.query(`
      INSERT INTO usuarios (username, password, role, active)
      VALUES ($1, $2, $3, true)
      ON CONFLICT (username)
      DO UPDATE SET password=$2, role=$3, active=true
    `, [String(username).trim().toUpperCase(), password, role || 'user']);

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/usuarios/:username', async (req, res) => {
  const username = String(req.params.username || '').trim().toUpperCase();

  if (username === 'SISTEMAS1900') {
    return res.status(400).json({ error: 'No puedes eliminar este usuario' });
  }

  try {
    if (!pool) return res.status(500).json({ error: 'Sin DB' });
    await pool.query('UPDATE usuarios SET active=false WHERE UPPER(username)=$1', [username]);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

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
        try {
          data = JSON.parse(data);
        } catch {
          data = [];
        }
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

        try {
          fs.unlinkSync(tmpPath);
        } catch {}

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

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function rowsToTable(rows) {
  if (!rows || rows.length === 0) return '(sin datos)';

  const firstRow = rows.find(r => r && typeof r === 'object');
  if (!firstRow) return '(sin datos)';

  const cols = Object.keys(firstRow);
  if (cols.length === 0) return '(sin datos)';

  const header = `| ${cols.join(' | ')} |`;
  const separator = `| ${cols.map(() => '---').join(' | ')} |`;

  const lines = rows.map(row =>
    `| ${cols.map(col => {
      const v = row[col];

      if (v === null || v === undefined || v === '') return '-';

      return String(v)
        .replace(/\n/g, ' ')
        .replace(/\|/g, '/')
        .trim();
    }).join(' | ')} |`
  );

  return [header, separator, ...lines].join('\n');
}

function getImportantKeywords(text) {
  const stopWords = new Set([
    'que', 'como', 'cual', 'cuales', 'dame', 'dime', 'muestra', 'genera',
    'hacer', 'haz', 'para', 'por', 'con', 'sin', 'los', 'las', 'del',
    'una', 'uno', 'este', 'esta', 'hay', 'tiene', 'pueden', 'quiero',
    'necesito', 'favor', 'reporte', 'tabla', 'excel', 'pdf', 'info',
    'informacion', 'datos', 'todo', 'todos', 'toda', 'todas', 'top',
    'mayor', 'menor', 'mas', 'menos', 'cliente', 'clientes', 'saldo',
    'saldos', 'ventas', 'venta', 'archivo', 'archivos', 'qad', 'sobre',
    'de', 'la', 'el', 'en', 'y', 'o', 'a', 'un', 'al', 'me', 'lo',
    'explica', 'explicame', 'detalla', 'detalle'
  ]);

  return normalizeText(text)
    .split(' ')
    .map(w => w.trim())
    .filter(w => w.length > 2 && !stopWords.has(w));
}

function formatQADFilesTable(cache) {
  const keys = Object.keys(cache || {});

  if (keys.length === 0) {
    return 'No tengo archivos QAD cargados actualmente.';
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

    return `| ${c.filename || '-'} | ${c.sheet || '-'} | ${registros} | ${uso} | ${fecha} |`;
  }).join('\n');

  return `## Archivos QAD cargados

| Archivo | Hoja | Registros | Uso probable | Última actualización |
|---|---|---:|---|---|
${rows}

Estos son los archivos disponibles en PostgreSQL para análisis. Puedes pedirme reportes por clientes, cartera, saldos, ventas, inventario, pedidos, producción o proveedores.`;
}

function buildQADContext(cache, messages) {
  const keys = Object.keys(cache || {});
  if (keys.length === 0) return '';

  const lastMsg = messages[messages.length - 1]?.content || '';
  const lastMsgNorm = normalizeText(lastMsg);
  const keywords = getImportantKeywords(lastMsg);

  console.log('🔎 Keywords detectadas:', keywords);

  const categories = {
    cliente: ['cliente', 'clientes', 'cartera', 'cobrar', 'saldo', 'credito', 'deuda', 'vencido', 'antiguedad', 'cxc', 'cobranza'],
    venta: ['venta', 'ventas', 'factura', 'facturacion', 'ingreso', 'factur'],
    vendedor: ['vendedor', 'vendedores'],
    pedido: ['pedido', 'pedidos', 'orden', 'ordenes'],
    proveedor: ['proveedor', 'proveedores', 'cxp'],
    produccion: ['produccion', 'hilatura', 'tejido', 'acabado', 'merma'],
    inventario: ['inventario', 'stock', 'almacen', 'tela', 'producto', 'especif'],
  };

  const msgCats = new Set();

  Object.entries(categories).forEach(([cat, words]) => {
    if (words.some(w => lastMsgNorm.includes(w))) msgCats.add(cat);
  });

  const includeAll = msgCats.size === 0;

  const sheetIndex = keys.map(k => {
    const c = cache[k];

    return `- ${c.filename} / Hoja: ${c.sheet}: ${Array.isArray(c.data) ? c.data.length : '?'} registros`;
  }).join('\n');

  const results = [];

  for (const key of keys) {
    const c = cache[key];
    if (!c || !Array.isArray(c.data) || c.data.length === 0) continue;

    const rows = c.data;
    const fileNorm = normalizeText(`${c.filename} ${c.sheet} ${key}`);

    const fileCats = new Set();

    Object.entries(categories).forEach(([cat, words]) => {
      if (words.some(w => fileNorm.includes(w))) fileCats.add(cat);
    });

    const catMatch = includeAll || [...msgCats].some(cat => fileCats.has(cat));

    const matchedRows = [];

    for (const row of rows) {
      const rowText = normalizeText(JSON.stringify(row));
      let score = 0;

      for (const kw of keywords) {
        if (rowText.includes(kw)) score += 30;

        const words = rowText.split(' ');

        if (words.some(w => w.startsWith(kw) || kw.startsWith(w))) score += 10;
      }

      if (catMatch) score += 5;

      if (score > 0) matchedRows.push({ row, score });
    }

    matchedRows.sort((a, b) => b.score - a.score);

    let rowsToInclude = [];

    if (keywords.length > 0) {
      rowsToInclude = matchedRows.slice(0, 80).map(r => r.row);
    } else if (catMatch) {
      rowsToInclude = rows.slice(0, 80);
    }

    if (rowsToInclude.length > 0) {
      results.push({
        key,
        score: matchedRows.reduce((s, r) => s + r.score, 0) + (catMatch ? 1000 : 0),
        text: `
### ARCHIVO: ${c.filename}
### HOJA: ${c.sheet}
### TOTAL REGISTROS EN ARCHIVO: ${rows.length}
### REGISTROS ENVIADOS A AVIVA: ${rowsToInclude.length}
### COINCIDENCIAS DIRECTAS: ${matchedRows.length}

${rowsToTable(rowsToInclude)}
`
      });
    }
  }

  results.sort((a, b) => b.score - a.score);

  const MAX_CHARS = 45000;
  let totalChars = 0;
  const selected = [];
  let selectedCount = 0;

  for (const r of results) {
    if (selectedCount >= 6) break;

    const trimmedText = r.text.substring(0, 8000);

    if (totalChars + trimmedText.length > MAX_CHARS) break;

    selected.push(trimmedText);
    totalChars += trimmedText.length;
    selectedCount++;
  }

  if (selected.length === 0) {
    const samples = keys.slice(0, 5).map(k => {
      const c = cache[k];

      return `
### ARCHIVO: ${c.filename}
### HOJA: ${c.sheet}
### TOTAL REGISTROS: ${c.data.length}

${rowsToTable((c.data || []).slice(0, 15))}
`;
    }).join('\n');

    return `
===== DATOS QAD =====

ARCHIVOS DISPONIBLES:
${sheetIndex}

No hubo coincidencias exactas con la pregunta. Aquí hay una muestra de datos disponibles:

${samples}
`;
  }

  return `
===== DATOS QAD DISPONIBLES =====

ARCHIVOS DISPONIBLES:
${sheetIndex}

===== DATOS RELEVANTES PARA ESTA CONSULTA =====

${selected.join('\n')}

===== INSTRUCCIONES DE USO DE DATOS =====

- Usa únicamente los datos anteriores.
- Cuando haya registros, responde con tabla markdown.
- No respondas en una sola frase.
- Da resumen ejecutivo, tabla, análisis y observaciones.
- Mantén nombres, códigos, importes y fechas exactamente como aparecen.
`;
}

function buildPermissionContext(username) {
  const cleanUsername = String(username || '').trim().toUpperCase();

  const NO_CONTABLE_USERS = [
    'LEO2026',
    'GTE_ACABADO',
    'GTE_TEJIDO',
    'GTE_HILATURA',
    'ISRAEL_ACABADO',
    'MEMO_TEJIDO',
    'CARLOSM_H',
  ];

  const VENDEDOR_USERS = {
    VENTAS_MONICA: 'Monica Caceres',
    VENTAS_EDGAR: 'Edgar Zarate',
    VENTAS_AMELIA: 'Amelia Aragon',
    VENTAS_JORGE: 'Jorge Mejia',
  };

  if (NO_CONTABLE_USERS.includes(cleanUsername)) {
    return `\n\nRESTRICCIONES DE ACCESO:
- No mostrar información contable: activos, pasivos, balance general, estado de resultados, cuentas contables, debe, haber, patrimonio.
- Si preguntan esos temas responde: "No tienes acceso a información contable. Consulta con el área de sistemas."
- Sí puedes mostrar facturación, ventas, producción, inventario, clientes, proveedores, pedidos, cobranza y datos operativos.`;
  }

  if (VENDEDOR_USERS[cleanUsername]) {
    const vend = VENDEDOR_USERS[cleanUsername];

    return `\n\nRESTRICCIONES DE ACCESO:
- Solo puedes mostrar información del vendedor: ${vend}.
- Filtra siempre los datos para mostrar únicamente ventas, pedidos, atrasos, clientes y bodega/almacén relacionados con ${vend}.
- Si preguntan sobre otros vendedores responde: "Solo tienes acceso a tu información de ventas."
- Nunca muestres datos de otros vendedores.`;
  }

  if (cleanUsername === 'ISMAEL_VENTAS') {
    return `\n\nACCESOS HABILITADOS:
- Puede ver ventas de todos los vendedores, producción, pedidos con y sin cliente, almacén de tela acabada y bodega.`;
  }

  return '';
}

function buildExecutivePrompt() {
  return `
==============================
MODO DE RESPUESTA AVIVA PREMIUM
==============================

Eres AVIVA, la inteligencia empresarial de Grupo Vivatex.

Tu estilo debe parecerse a un analista corporativo senior:
- claro
- profesional
- detallado
- ordenado
- concreto
- visualmente limpio

REGLAS OBLIGATORIAS:

1. Si hay datos QAD, NO respondas genérico.
2. Si hay datos QAD, SIEMPRE usa tablas markdown.
3. Para cualquier consulta de clientes, saldos, cartera, ventas, pedidos, inventario o producción, responde con esta estructura:

## Resumen ejecutivo
Explica qué encontraste.

## Tabla de datos encontrados
Incluye tabla markdown con columnas importantes.

## Análisis
Explica puntos relevantes, riesgos, saldos, fechas, atrasos, cantidades o patrones.

## Observaciones
Aclara si faltan campos, si hay varias coincidencias o si los datos parecen incompletos.

## Siguiente acción sugerida
Sugiere qué revisar, exportar o confirmar.

4. Si el usuario pregunta por un cliente específico:
- Busca nombre, código, razón social o coincidencias parciales.
- Muestra todos los registros encontrados.
- No ocultes nombres.
- No inventes campos.
- Si un campo no aparece, escribe "No disponible".

5. Si hay muchos registros:
- Muestra los más relevantes.
- Ordena por saldo, fecha, importe, cliente, pedido o relevancia.
- Di cuántos registros se revisaron y cuántos se muestran.

6. Las tablas deben escribirse en markdown correcto:
| Columna | Columna |
|---|---|
| Dato | Dato |

7. Nunca respondas solo "no tengo información" si hay archivos QAD disponibles.
8. No seas demasiado breve.
9. No inventes números.
10. Mantén exactamente nombres, códigos, fechas e importes como aparecen.
`;
}

app.post('/api/chat', async (req, res) => {
  const { messages, system, username } = req.body || {};

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'Faltan mensajes' });
  }

  if (!system) {
    return res.status(400).json({ error: 'Falta system prompt' });
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
    const lastMsgNorm = normalizeText(lastMsg);

    if (
      lastMsgNorm.includes('que archivos') ||
      lastMsgNorm.includes('archivos qad') ||
      lastMsgNorm.includes('que datos tienes') ||
      lastMsgNorm.includes('que tienes cargado') ||
      lastMsgNorm.includes('archivos cargados')
    ) {
      return res.json({
        reply: formatQADFilesTable(cache),
      });
    }

    const qadContext = buildQADContext(cache, messages);
    const permContext = buildPermissionContext(username);
    const executivePrompt = buildExecutivePrompt();

    const dataSize = qadContext.length;
    const sheetsIncluded = (qadContext.match(/### ARCHIVO:/g) || []).length;

    console.log(`Chat request: user="${username || '-'}" provider="${AI_PROVIDER}" model="${OPENAI_MODEL}" qadContext=${dataSize} chars, sheets=${sheetsIncluded}`);

    const fullSystem = `
${String(system)}

${executivePrompt}

${qadContext}

${permContext}
`;

    const openaiMessages = [
      { role: 'system', content: fullSystem },
      ...messages.map(m => ({
        role: m.role === 'assistant' ? 'assistant' : 'user',
        content: typeof m.content === 'string' ? m.content : JSON.stringify(m.content || ''),
      })),
    ];

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        messages: openaiMessages,
        temperature: 0.05,
        max_tokens: 4096,
        presence_penalty: 0,
        frequency_penalty: 0,
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
