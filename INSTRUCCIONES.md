# 🚗 Sólido Auto Servicio — Página Web

Archivo único `index.html` listo para producción. Sin npm, sin build, sin dependencias.

---

## ⚙️ Configuración inicial

Abre `index.html` y edita el bloque `CONFIG` al inicio del archivo:

```js
const CONFIG = {
  API_BASE:   "https://crm-automotriz-production.up.railway.app", // ← Tu URL de Railway
  WHATSAPP:   "18097122027",       // Número con código de país, sin +
  TELEFONO:   "809-712-2027",
  DIRECCION:  "Santo Domingo, República Dominicana",
  MAPS_EMBED: "...",               // URL del iframe de Google Maps
  HORARIO: [
    { dia: "Lunes – Viernes", hora: "7:30 AM – 6:00 PM" },
    { dia: "Sábados",         hora: "8:00 AM – 2:00 PM" },
    { dia: "Domingos",        hora: "Cerrado" },
  ],
};
```

### Cómo obtener el embed de Google Maps
1. Ve a maps.google.com y busca tu dirección exacta
2. Clic en "Compartir" → "Incorporar mapa"
3. Copia solo el valor del atributo `src="..."` del iframe
4. Pégalo en `CONFIG.MAPS_EMBED`

---

## 🖼️ Logo

Coloca tu logo en la misma carpeta que `index.html` con el nombre `logo.png`.

Si el logo no existe, la web funciona igual (simplemente no muestra imagen).

---

## 🚀 Opciones de despliegue

### Opción 1 — Netlify (Gratis, recomendado)
1. Ve a [netlify.com](https://netlify.com) y crea una cuenta gratis
2. Arrastra la carpeta `solido-web` al área de deploy
3. Listo. Obtienes un enlace como `solidoauto.netlify.app`
4. Puedes conectar tu propio dominio (ej: `solidoautoservicio.com`)

### Opción 2 — Vercel (Gratis)
1. Ve a [vercel.com](https://vercel.com)
2. "Add New Project" → sube la carpeta
3. Deploy en segundos

### Opción 3 — Hosting tradicional (cPanel)
1. Compra un dominio (ej: en GoDaddy o Namecheap)
2. Sube la carpeta `solido-web` por FTP a `public_html`
3. La web queda accesible en tu dominio

---

## 🔗 Secciones conectadas a tu CRM

| Sección | Endpoint |
|---|---|
| Consulta de vehículo | `GET /vehiculo-historial/placa/:placa` |
| Menú cafetería | `GET /cafeteria/productos` |

Ambos endpoints ya existen en tu backend de Railway.

---

## 📱 WhatsApp

El botón flotante de WhatsApp y todos los enlaces de contacto se generan automáticamente desde `CONFIG.WHATSAPP`.

Para personalizar el mensaje inicial edita la función `waUrl()` en el script.

---

## ✏️ Personalización rápida

- **Colores**: Busca `:root` en el CSS y cambia `--blue`, `--orange`, etc.
- **Estadísticas del hero** ("+500 clientes"): Busca `stats-bar` en el HTML
- **Servicios**: Busca `serv-card` en el HTML y edita los bloques
- **Redes sociales** (footer): Cambia los `href` de los íconos 📘📸 al final del HTML

---

## 🌐 CORS

Si al consultar la placa aparece error de conexión, asegúrate de que tu backend de Railway tenga CORS habilitado para el dominio de la web. En `server.mjs` debe estar:

```js
app.use(cors({ origin: "*" }));
// o más específico:
app.use(cors({ origin: "https://tudominio.com" }));
```
