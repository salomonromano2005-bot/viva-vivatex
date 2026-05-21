# 🌿 VIVA — Asistente Inteligente Vivatex

Sistema de IA para Grupo Vivatex S.C. de C.V., conectado a QAD Enterprise.

---

## 📁 Estructura del proyecto

```
viva/
├── public/
│   └── index.html       ← Toda la interfaz (chat + admin)
├── server.js            ← Backend (Node.js + Express)
├── package.json
└── README.md
```

---

## 🔑 PASO 1 — Obtener API Key de Claude (Anthropic)

1. Ve a: https://console.anthropic.com
2. Crea una cuenta o inicia sesión
3. Ve a "API Keys" → "Create Key"
4. Copia la key (empieza con `sk-ant-...`)
5. Guárdala, la necesitarás en el Paso 3

---

## 🚀 PASO 2 — Opciones de despliegue (elige una)

### Opción A: Railway (MÁS FÁCIL — recomendada) ⭐
**Gratis para empezar, links automáticos**

1. Ve a https://railway.app y crea cuenta gratuita
2. Haz clic en "New Project" → "Deploy from GitHub"
3. Sube el código a GitHub primero:
   - Ve a https://github.com → New repository → "viva-vivatex"
   - Sube todos los archivos de esta carpeta
4. Railway detecta automáticamente que es Node.js
5. En Railway → Variables → agrega:
   ```
   ANTHROPIC_API_KEY = sk-ant-TU_KEY_AQUI
   QAD_URL          = https://TU_SERVIDOR_QAD.com
   QAD_TOKEN        = TU_TOKEN_QAD
   QAD_COMPANY      = VIVATEX
   ```
6. Railway te da un link automático tipo:
   `https://viva-vivatex.up.railway.app`

---

### Opción B: Render (también gratuito)

1. Ve a https://render.com
2. New → Web Service → conecta tu repo de GitHub
3. Build Command: `npm install`
4. Start Command: `node server.js`
5. Agrega las variables de entorno igual que Railway
6. Tu link: `https://viva-vivatex.onrender.com`

---

### Opción C: Servidor propio (si tienen uno en Vivatex)

```bash
# En el servidor, ejecutar:
git clone TU_REPO
cd viva
npm install

# Configurar variables de entorno:
export ANTHROPIC_API_KEY="sk-ant-TU_KEY"
export QAD_URL="http://QAD_SERVER_IP"
export QAD_TOKEN="TU_TOKEN_QAD"
export QAD_COMPANY="VIVATEX"
export PORT=3000

# Iniciar:
node server.js

# Para que corra siempre (con PM2):
npm install -g pm2
pm2 start server.js --name viva
pm2 startup
pm2 save
```

---

## 🌐 PASO 3 — Dominio personalizado (opcional)

Si tienen `vivatex.com` o quieren `viva.vivatex.com`:

1. En tu proveedor de dominio (GoDaddy, Namecheap, etc.)
2. Agrega un registro CNAME:
   - Nombre: `viva`
   - Valor: tu link de Railway/Render
3. En Railway/Render → Settings → Custom Domain → agrega `viva.vivatex.com`
4. En ~5 minutos funciona: `https://viva.vivatex.com`

---

## 📱 PASO 4 — Instalar como app en dispositivos

### iPhone / iPad:
1. Abrir Safari → ir al link de VIVA
2. Tocar el botón compartir (cuadro con flecha)
3. "Agregar a pantalla de inicio"
4. ¡Aparece el ícono de VIVA como app!

### Android:
1. Abrir Chrome → ir al link de VIVA
2. Menú (⋮) → "Agregar a pantalla de inicio"
3. ¡Listo!

### Windows / Mac (computadora):
1. Abrir Chrome o Edge → ir al link
2. Barra de direcciones → ícono de instalar (⊕)
3. "Instalar VIVA"
4. Aparece como aplicación de escritorio

---

## ⚙️ Conexión con QAD

La IA detecta automáticamente qué módulo consultar según la pregunta:

| Pregunta sobre... | Módulo QAD consultado |
|---|---|
| Ventas / órdenes / pedidos | `/api/sales/summary` |
| Inventario / existencias / stock | `/api/inventory/status` |
| Compras / proveedores | `/api/purchase/orders` |
| Clientes | `/api/customer/list` |
| Finanzas / balance | `/api/finance/summary` |

**Nota para el equipo de Sistemas:**
Los endpoints anteriores son genéricos. Deben ajustarlos en `server.js` (función `fetchQADData`) según la estructura real de la API de QAD de Vivatex.

---

## 👥 Usuarios incluidos

| Usuario | Rol | Acceso |
|---|---|---|
| JRZ123 | Usuario | Solo chat |
| JRR234 | Usuario | Solo chat |
| SRR456 | Usuario | Solo chat |
| SRB0707 | Usuario | Solo chat |
| LEO2026 | Usuario | Solo chat |
| SISTEMAS1900 | **Admin** | Chat + Panel de Sistemas |

Los usuarios se pueden gestionar desde el Panel de Sistemas (sin tocar código).

---

## 🆘 Soporte

Cualquier problema técnico, el usuario SISTEMAS1900 tiene acceso al Panel de Sistemas donde puede:
- Ver logs de actividad
- Agregar/quitar/editar usuarios
- Configurar o actualizar credenciales de QAD
- Ver el estado de la conexión
