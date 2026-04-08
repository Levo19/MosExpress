# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What is MOSexpress

MOSexpress is a **100% frontend PWA (Progressive Web App)** Point-of-Sale system built with Vue.js 3 + Tailwind CSS. It has no Node.js backend — all server logic lives in **Google Apps Script (GAS)** connected to a Google Sheets database.

## Project Structure

```
index.html        ← Entire frontend: Vue 3 app, all modules, all logic (~2000+ lines)
sw.js             ← Service Worker for offline/PWA caching (cache name: mosexpress-cache-v16)
manifest.json     ← PWA manifest
gas/Code.gs       ← Google Apps Script backend (deployed as Web App)
```

## Architecture

### Frontend (`index.html`)
Single-file Vue 3 app (no build step, CDN imports). All logic is in one `<script>` block with a single `createApp({...})` instance. Key state variables:

- `currentModule` — controls which panel is shown: `'CAJA'`, `'POS'`, or `'HERRAMIENTAS'`
- `dispositivoAutorizado` — `null` (loading) / `true` / `false` — gates the entire UI
- `config` — persisted in `localStorage` under `mosexpress_config`; contains `vendedor`, `estacion`, `completado`
- `pendingSales` — array of offline-queued sales, persisted in `localStorage`
- `cajaAbierta` / `idCajaActual` — cash register session state

### Backend (`gas/Code.gs`)
Google Apps Script deployed as a Web App. Entry points:
- `doGet(e)` — handles `accion=descargar` (catalog download) and `accion=verificar_dispositivo`
- `doPost(e)` — handles `tipoEvento=APERTURA_CAJA`, `tipoEvento=CIERRE_CAJA`, `accion=imprimir`, and default (sale registration)

The `PRINTNODE_API_KEY` must be stored in Script Properties (not in code). The GAS acts as a proxy so the key never reaches the browser.

### Communication Flow
1. On app start → `GET ?accion=verificar_dispositivo&id=<deviceId>`
2. On first load → `GET ?accion=descargar` to fetch catalog (PRODUCTO_BASE, PRESENTACIONES, EQUIVALENCIAS, PROMOCIONES, ZONAS_CONFIG, CLIENTES_FRECUENTES)
3. On sale → `POST` with full sale payload; if offline, queued in `pendingSales` and retried every 5 minutes via `syncPendientes()`
4. On print → `POST` with `accion=imprimir`, `printerId`, and base64-encoded ESC/POS `content`

### Device Authorization
The app generates a `crypto.randomUUID()` on first launch and stores it as `mosexpress_deviceId` in `localStorage`. This ID must be registered as `ACTIVO` in the `DISPOSITIVOS` sheet. If not, the UI is fully blocked.

## Deployment

No build process. Deployment is:
1. Set `API_URL` in `index.html` to the GAS Web App Executable URL
2. Commit + push to `main` branch on GitHub
3. GitHub Pages serves the static files

To update the PWA cache version (force all devices to reload): bump `CACHE_NAME` in `sw.js` (e.g., `mosexpress-cache-v18`).

## Google Sheets Schema

Required tabs: `PRODUCTO_BASE`, `PRESENTACIONES`, `EQUIVALENCIAS`, `PROMOCIONES`, `ZONAS_CONFIG`, `DISPOSITIVOS`, `CAJAS`, `VENTAS_CABECERA`, `VENTAS_DETALLE`, `CLIENTES_FRECUENTES`
New required: `CORRELATIVOS` (auto-created by GAS on first sale)
Optional: `LOG_IMPRESIONES`

### Schema changes (NubeFact preparation)

**PRODUCTO_BASE** — new columns (source of truth for IGV/medida per product):
- `Tipo_IGV` — 1=Gravado (default), 2=Exonerado, 3=Inafecto
- `Unidad_Medida` — "NIU" (default/unidad), "KGM" (kg), "ZZ" (servicio), etc.
- `Cod_SUNAT` — already existed, now flows through to items payload

**VENTAS_CABECERA** — col 16 added: `Tipo_Doc_Cliente` (0=sin doc, 1=DNI, 6=RUC)

**VENTAS_DETALLE** — cols 8-10 added (values come from PRODUCTO_BASE via carrito):
- Col 8: `Valor_Unitario` — precio sin IGV (precio/1.18 si Gravado, precio si Exonerado/Inafecto)
- Col 9: `Tipo_IGV` — viene de PRODUCTO_BASE.Tipo_IGV (no hardcodeado)
- Col 10: `Unidad_Medida` — viene de PRODUCTO_BASE.Unidad_Medida (no hardcodeado)

**CLIENTES_FRECUENTES** — col 5 added: `Direccion` (dirección fiscal del RUC)

**CORRELATIVOS** (nueva hoja) — Cabeceras: `Serie | Siguiente`
- GAS la crea automáticamente la primera vez. Permite correlativo O(1) en vez de O(n).
- Agregar manualmente la hoja antes del primer deploy para evitar el fallback lento.

### NubeFact fields in ventaBase payload
```
header.total_gravada    — suma de subtotales gravados / 1.18
header.total_igv        — gravada × 0.18
header.total_exonerada  — suma de subtotales exonerados (sin IGV)
header.total_inafecta   — suma de subtotales inafectos (sin IGV)
header.porcentaje_igv   — 18
header.moneda           — 1 (PEN)
header.cliente.tipo     — 0/1/6 (auto-detectado del largo del doc)
header.cliente.direccion — dirección fiscal (llenada por API DNI/RUC)
items[].valor_unitario  — desde PRODUCTO_BASE.Tipo_IGV: Gravado→precio/1.18, otros→precio
items[].tipo_igv        — desde PRODUCTO_BASE.Tipo_IGV
items[].unidad_de_medida — desde PRODUCTO_BASE.Unidad_Medida
items[].cod_sunat       — desde PRODUCTO_BASE.Cod_SUNAT
```

### Data flow for IGV fields
```
PRODUCTO_BASE (Tipo_IGV, Unidad_Medida, Cod_SUNAT)
  → mapearCatalogoLocal (tipoIgv, unidadMedida, codSunat)
  → elegirPresentacion / agregarCarritoObj
  → carrito item (c.tipoIgv, c.unidadMedida, c.codSunat)
  → itemsParaImprimir (tipo_igv, unidad_de_medida, cod_sunat)
  → ventaBase → GAS → VENTAS_DETALLE cols 8-10
```

## Key Behaviors to Preserve

- **Offline-first**: Sales must always be saved to `pendingSales` (localStorage) before attempting network sync. Never block the sale flow on network availability.
- **Correlativo generation**: `obtenerSiguienteCorrelativoRapido()` uses `CORRELATIVOS` sheet (O(1), with LockService). Falls back to legacy `obtenerSiguienteCorrelativo()` (O(n) scan) if the sheet doesn't exist.
- **PrintNode integration**: GAS handles printing internally within `procesarVenta` (single round-trip). Browser uses `mandarImpresionPrintNode` only as offline/fallback. The API key lives in Script Properties.
- **Dark mode**: Implemented via CSS class `dark-mode` on the document root, not Tailwind's dark variant.
- **direccion from DNI/RUC API**: `consultarCliente` returns `direccion` for RUC lookups (APISPeru). Empty string for DNI (RENIEC doesn't expose address). Stored in `CLIENTES_FRECUENTES.Direccion` col.
