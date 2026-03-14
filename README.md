# MOSexpress v2.0 - Sistema POS PWA Multi-Zona 🛒

Bienvenido al repositorio oficial de **MOSexpress**, un Sistema de Punto de Venta (POS) diseñado como una Aplicación Web Progresiva (PWA) optimizada para móviles, tablets y escritorio.

## 📌 ¿De qué trata este proyecto?
MOSexpress fue diseñado bajo la arquitectura *Rompefilas v2.0* para solucionar los cuellos de botella en la atención de clientes. Es un sistema **100% Frontend** escrito en HTML, Tailwind CSS y Vue.js 3, que se conecta directamente a **Google Sheets** a través de Google Apps Script (GAS) convirtiéndolo en un robusto backend gratuito.

El objetivo principal de este documento (el `README.md`) es **guiarte a ti (o a cualquier otro desarrollador)** sobre cómo funciona el proyecto, cómo instalarlo y cómo actualizarlo en el futuro. ¡Es el manual oficial de tu código!

---

## ✨ Características Principales

1. **Arquitectura Modular (Bottom Navigation)**: 
   - **Módulo POS**: Catálogo y cobro con UX proactiva.
   - **Módulo CAJA**: Control estricto de turnos, historial de ventas (Notas, Boletas, Facturas) y anulación/reimpresión de tickets.
   - **Módulo HERRAMIENTAS**: Generador e impresión silenciosa de membretes y etiquetas de precios (80mm x 300mm).
2. **Offline-First (PWA)**: Registra un `Service Worker` para que el catálogo de productos y la aplicación carguen de forma inmediata, incluso si el internet falla.  
3. **Escáner Integrado**: Permite usar la cámara de cualquier celular (vía `html5-qrcode`) para escanear códigos de barras.
4. **Middleware de Caja**: Bloqueo inteligente del POS. Nadie puede realizar una venta sin antes establecer un saldo inicial y registrar una "Caja Abierta" en Google Sheets.
5. **PrintNode API**: Impresión RAW *silenciosa* (sin ventanas emergentes de impresoras) para emitir comprobantes profesionales a cualquier ticketera térmica local conectada a internet.

---

## 🛠️ Tecnologías Usadas

- **Frontend**: 
  - [Vue.js 3](https://vuejs.org/) (Reactividad y Estado)
  - [Tailwind CSS](https://tailwindcss.com/) (Diseño responsivo y componentes de UI)
  - HTML5 & JavaScript (Vanilla, todo contenido en *index.html*).
- **Backend & Base de Datos**: 
  - [Google Apps Script (GAS)](https://developers.google.com/apps-script)
  - Google Sheets (Estructuras: PRODUCTOS, PROMOCIONES, CLIENTES, CAJAS, VENTAS_CABECERA, etc.)
- **Hardware Integrations**:
  - `html5-qrcode` para lectura óptica de barras y QRs.
  - PrintNode (Comandos nativos ESC/POS e impresión local remota).

---

## 🚀 Despliegue en GitHub Pages (Guía Rápida)

Ya que este proyecto no requiere un servidor NodeJS, el despliegue es completamente estático:

1. Modifica la variable `API_URL` dentro del bloque `<script>` en el `index.html` con tu endpoint publicado de Google Apps Script.
2. Ingresa tu API Key en la variable `PRINTNODE_API_KEY`.
3. Sube todos los archivos de esta carpeta a la rama `main` o `master` de tu repositorio público en GitHub.
4. En Settings > Pages, activa la rama `main` y guarda.
5. ¡Listo! Accede a `https://[tu-usuario].github.io/mosexpress` y la App solicitará los permisos PWA para instalarse en tu teléfono.

---

## 📂 Estructura de Archivos

Al consolidar la aplicación, nos hemos quedado con la estructura más limpia y profesional posible:

- `index.html`: **El Corazón**. Contiene todo el esqueleto UI, la lógica, el enrutado, los estilos de Tailwind y los componentes de Vue.js. (Sustituye al antiguo `app.js`).
- `manifest.json`: Identidad de la App para ser instalable (Iconos, colores de navegador, pantalla completa).
- `sw.js`: Archivo Service Worker de caché sin conexión.
- `code.gs`: **Backup del Backend**. Es una copia local del código que vive actualmente funcionando en tu servidor de Google Apps Script.
- `README.md`: Este manual guía.

---

## 💡 ¿Siempre debo actualizar este README?

**Sí, es una excelente práctica.** 
Cada vez que agregues un módulo nuevo importante (por ejemplo: "Módulo de Devoluciones", "Integración con SUNAT para Facturación Electrónica en Producción", etc.), puedes venir a este archivo y añadir la documentación.

Eso te ayudará a que, si retomas el proyecto en 6 meses, o contratas a un equipo de desarrolladores para que lo extiendan, este README les diga exactamente qué hace la App y no tengan que empezar desde cero intentando adivinar el código.
