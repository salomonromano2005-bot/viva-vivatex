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

// ============================================================
const PORT = process.env.PORT || 3000;
const DATABASE_URL = process.env.DATABASE_URL || '';

const AI_PROVIDER = process.env.AI_PROVIDER || 'openai';

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1';

const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || '';
// ============================================================

let pool = null;
let qadDataCache = {};
let qadLastUpdate = null;

// ============================================================
// DB
// ============================================================

function getPgSslConfig() {
  if (!DATABASE_URL) return false;

  if (
    DATABASE_URL.includes('localhost') ||
    DATABASE_URL.includes('127.0.0.1')
  ) {
    return false;
  }

  return {
    rejectUnauthorized: false,
  };
}

async function initDB() {
  if (!DATABASE_URL) {
    console.log('⚠️ Sin DATABASE_URL');
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

    console.log('✅ PostgreSQL conectado');
  } catch (e) {
    console.error('❌ PostgreSQL error:', e.message);
  }
}

// ============================================================
// HEALTH
// ============================================================

app.get('/api/health', async (req, res) => {
  res.json({
    ok: true,
    provider: AI_PROVIDER,
    openai: !!OPENAI_API_KEY,
    anthropic: !!ANTHROPIC_API_KEY,
    db: !!pool,
    model: OPENAI_MODEL,
  });
});

// ============================================================
// LOGIN
// ============================================================

app.post('/api/login', async (req, res) => {
  const { username, password } = req.body || {};

  try {
    const result = await pool.query(
      'SELECT id, username, role, active FROM usuarios WHERE UPPER(username)=UPPER($1) AND password=$2',
      [username, password]
    );

    if (result.rows.length === 0) {
      return res.json({
        ok: false,
        error: 'Usuario o contraseña incorrectos',
      });
    }

    const user = result.rows[0];

    if (!user.active) {
      return res.json({
        ok: false,
        error: 'Usuario inactivo',
      });
    }

    return res.json({
      ok: true,
      user,
    });
  } catch (e) {
    console.error(e);

    return res.status(500).json({
      ok: false,
      error: 'Error login',
    });
  }
});

// ============================================================
// MULTER
// ============================================================

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
});

// ============================================================
// PARSE EXCEL
// ============================================================

function parseFile(buffer) {
  const workbook = xlsx.read(buffer, {
    type: 'buffer',
  });

  const result = {};

  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];

    result[sheetName] = xlsx.utils.sheet_to_json(sheet, {
      defval: '',
    });
  });

  return result;
}

// ============================================================
// QAD UPLOAD
// ============================================================

app.post('/api/qad/upload', upload.array('files', 20), async (req, res) => {
  try {
    let totalSheets = 0;

    for (const file of req.files) {
      const parsed = parseFile(file.buffer);

      const baseName = path.basename(
        file.originalname,
        path.extname(file.originalname)
      );

      for (const sheet of Object.keys(parsed)) {
        const key = `${baseName}_${sheet}`
          .replace(/\s+/g, '_')
          .toLowerCase();

        qadDataCache[key] = {
          data: parsed[sheet],
          filename: file.originalname,
          sheet,
        };

        totalSheets++;
      }
    }

    qadLastUpdate = new Date().toISOString();

    return res.json({
      ok: true,
      sheets: totalSheets,
    });
  } catch (e) {
    console.error(e);

    return res.status(500).json({
      error: 'Error cargando archivos',
    });
  }
});

// ============================================================
// PDF
// ============================================================

app.post('/api/qad/upload-pdf', upload.single('pdf'), async (req, res) => {
  try {
    const parsed = await pdfParse(req.file.buffer);

    const key = `pdf_${Date.now()}`;

    qadDataCache[key] = {
      data: [{ texto: parsed.text }],
      filename: req.file.originalname,
      sheet: 'PDF',
    };

    return res.json({
      ok: true,
    });
  } catch (e) {
    console.error(e);

    return res.status(500).json({
      error: 'Error PDF',
    });
  }
});

// ============================================================
// STATUS
// ============================================================

app.get('/api/qad/status', async (req, res) => {
  const keys = Object.keys(qadDataCache);

  return res.json({
    hasData: keys.length > 0,
    sheets: keys,
    total: keys.length,
    updatedAt: qadLastUpdate,
  });
});

// ============================================================
// CONTEXT
// ============================================================

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function rowsToTable(rows) {
  if (!rows || rows.length === 0) return '';

  const first = rows[0];

  const cols = Object.keys(first);

  const header = cols.join(' | ');

  const lines = rows.slice(0, 100).map(r =>
    cols.map(c => String(r[c] || '')).join(' | ')
  );

  return [header, ...lines].join('\n');
}

function buildQADContext(cache) {
  const keys = Object.keys(cache);

  if (keys.length === 0) return '';

  const sections = [];

  for (const key of keys) {
    const c = cache[key];

    sections.push(`
### ${c.filename} / ${c.sheet}

${rowsToTable(c.data)}
`);
  }

  return `
===== DATOS QAD =====

${sections.join('\n')}
`;
}

// ============================================================
// OPENAI CHAT
// ============================================================

app.post('/api/chat', async (req, res) => {
  const { messages, system } = req.body || {};

  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({
      error: 'Mensajes inválidos',
    });
  }

  try {
    const qadContext = buildQADContext(qadDataCache);

    const fullSystem = `
${system}

${qadContext}
`;

    console.log('🤖 Provider:', AI_PROVIDER);
    console.log('📊 QAD chars:', qadContext.length);

    if (AI_PROVIDER === 'openai') {
      if (!OPENAI_API_KEY) {
        return res.status(500).json({
          error: 'Falta OPENAI_API_KEY',
        });
      }

      const openaiMessages = [
        {
          role: 'system',
          content: fullSystem,
        },
        ...messages,
      ];

      const response = await fetch(
        'https://api.openai.com/v1/chat/completions',
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${OPENAI_API_KEY}`,
          },
          body: JSON.stringify({
            model: OPENAI_MODEL,
            messages: openaiMessages,
            temperature: 0.2,
            max_tokens: 4096,
          }),
        }
      );

      if (!response.ok) {
        const errText = await response.text();

        console.error('❌ OpenAI error:', errText);

        return res.status(response.status).json({
          error: 'Error OpenAI',
          detail: errText,
        });
      }

      const data = await response.json();

      const reply =
        data.choices?.[0]?.message?.content || 'Sin respuesta';

      return res.json({
        reply,
      });
    }

    return res.status(500).json({
      error: 'Provider inválido',
    });
  } catch (e) {
    console.error('❌ Chat error:', e);

    return res.status(500).json({
      error: 'Error interno',
      detail: e.message,
    });
  }
});

// ============================================================
// FRONTEND
// ============================================================

app.get('*', (req, res) => {
  const indexPath = path.join(__dirname, 'public', 'index.html');

  if (fs.existsSync(indexPath)) {
    return res.sendFile(indexPath);
  }

  return res.send('AVIVA ONLINE');
});

// ============================================================
// START
// ============================================================

initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`🚀 AVIVA ONLINE PORT ${PORT}`);
    console.log(`🤖 Provider: ${AI_PROVIDER}`);
    console.log(`🧠 Model: ${OPENAI_MODEL}`);
  });
});
