# MOSexpress v3.0 - Sistema POS PWA Multi-Zona 🛒

Bienvenido al repositorio oficial de **MOSexpress**, un Sistema de Punto de Venta (POS) diseñado como una Aplicación Web Progresiva (PWA) optimizada para móviles, tablets y escritorio.

## 📌 ¿De qué trata este proyecto?
MOSexpress fue diseñado bajo la arquitectura *Rompefilas v3.0 PROFESIONAL* para solucionar cuellos de botella en la atención de clientes. Es un sistema **100% Frontend** (HTML, Tailwind CSS, Vue.js 3) conectado a **Google Sheets** a través de Google Apps Script (GAS).

---

## 🏗️ Esquema de Base de Datos (Google Sheets)

Para que el sistema funcione correctamente, tu Google Sheets (conectado a `code.gs`) debe tener exactamente las siguientes pestañas con sus respectivas columnas:

1. **PRODUCTO_BASE**: `SKU_Base`, `Nombre`, `Categoria`, `Cod_Tributo`, `IGV_Porcentaje`, `Cod_SUNAT`.
2. **PRESENTACIONES**: `Cod_Barras`, `SKU_Base`, `Empaque`, `Factor`, `Precio_Venta`.
3. **EQUIVALENCIAS**: `Cod_Alias`, `Cod_Barras_Real`.
4. **PROMOCIONES**: `SKU_Base`, `Tipo_Promo`, `Cant_Min`, `Valor_Promo`.
5. **ZONAS_CONFIG**: `Zona_ID`, `Estacion_Nombre`, `PrintNode_ID`, `Serie_Nota`, `Serie_Boleta`, `Serie_Factura`.
6. **DISPOSITIVOS**: `ID_Dispositivo`, `Nombre_Equipo`, `Estado (ACTIVO/INACTIVO)`.
7. **CAJAS**: `ID_Caja`, `Vendedor`, `Estacion`, `Fecha_Apertura`, `Monto_Inicial`, `Estado (ABIERTA/CERRADA)`, `Monto_Final`.
8. **VENTAS_CABECERA**: `ID_Venta`, `Fecha`, `Vendedor`, `Estacion`, `Cliente_Doc`, `Cliente_Nom`, `Total`, `Tipo_Doc_Metodo`, `Correlativo_CPE`, `ID_Caja`, `ID_Dispositivo`, `Status_Envio`.
9. **VENTAS_DETALLE**: `ID_Venta`, `SKU`, `Nombre`, `Cantidad`, `Precio_Unit`, `Subtotal`.

---

## ✨ Características y Lógica Core

1. **🔒 Módulo de Seguridad (Dispositivos)**: 
   - Al iniciar, la App genera una Huella Digital (`crypto.randomUUID()`) almacenada en `localStorage` como `mosexpress_deviceId`.
   - Si este ID no existe como `ACTIVO` en la tabla `DISPOSITIVOS`, la App bloquea completamente la interfaz de usuario.
   
2. **📡 Resiliencia y Sincronización (Offline-First)**:
   - **Cola de Espera Local:** Toda venta procesada se almacena en el celular (`pendingSales` en localStorage) con estado `pending`.
   - **Background Sync:** Una rutina iterativa (`syncPendientes()`) corre en background **cada 5 minutos** (o al detectar retorno de red) empujando las ventas no enviadas al servidor.
   - **Tickets Provisionales:** Cero interrupciones. Si no hay red, la vendedora imprime una *PREVENTA PROVISIONAL* y sigue despachando.

3. **📱 Arquitectura Modular UX**:
   - **Módulo CAJA:** Gestión de Turno (Apertura y Cierre con Tickets Z). Acordeones para revisar y reimprimir comprobantes.
   - **Módulo POS:** Interfaz "Split Screen" (Dual panel). Catálogo Dinámico con promociones aplicadas en tiempo real.
   - **Módulo HERRAMIENTAS:** Impresión ciega de membretes de anaquel de precios.

4. **🖨️ Integración PrintNode**:
   - Uso de la API pura de PrintNode con inyección de código binario plano RAW (`ESC/POS` como `\x1b\x61\x01`) para cortes, códigos CODE128 y texto de doble altura en impresoras térmicas remotas.

---

## 🚀 Despliegue en GitHub Pages

Al carecer de NodeJS backend, el despliegue es 100% estático y gratuito:

1. Configura `API_URL` en el `<script>` de `index.html` apuntando a la URL Executable del Script GAS.
2. Ingresa la `PRINTNODE_API_KEY` válida.
3. Haz Commit + Push al *branch main* de GitHub.
4. Habilita Settings > Pages (Deploy from branch main).
5. Comparte la URL web. La app requerirá autenticar tu **Dispositivo** al entrar por primera vez.

---

## 💼 Flujos Comerciales Relevantes Incorporados
*   **Multimétodo de Pago**: Efectivo vs Virtual/Yape (con autocompletado de monto).
*   **Calculadora Vuelto**: Input veloz de billetes (10, 20, 50, 100).
*   **Ahorro Promo**: Etiqueta visual que detalla en línea cuánto ahorró el cliente de un producto en específico por una promo unitaria o grupal.
*   **MiddleWare de Caja**: Impide venta si el empleado no inicia turno primero.
*   **Cierre de Caja Operativo X/Z**: Cuadre de montos Virtual y Efectivo en físico para despilfarro 0%.

`¡Tu PWA Oficial Rompefilas v3.0 ha llegado a Producción!`
