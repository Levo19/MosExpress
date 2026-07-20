# CLAUDE.md

Guía para Claude Code al trabajar en este repositorio.

## Qué es MosExpress

PWA de punto de venta (POS) 100% frontend: Vue 3 + Tailwind (CDN, sin build step).
**Backend: 100% Supabase (cero GAS, cero fallback).** Todo el server-side vive en
Postgres (RPCs por esquema `me.*` / `mos.*`) y Edge Functions (`mint-me`, `imprimir`,
`emitir-cpe`, `ia`). Los SQL numerados están en `C:\Users\ISO\ProyectoMOS\supabase\`.

## Estructura

```
index.html    ← toda la app: Vue 3, 11 scripts inline (el grande es el setup())
sw.js         ← Service Worker (VERSION + push FCM; NO intercepta *.supabase.co)
radio.html    ← pantalla publicitaria SmartTV (ES5, GET RPC public.radio_productos)
version.json  ← el auto-update compara contra `var V` del index
manifest.json ← PWA manifest
```

Assets compartidos servidos desde el repo MOS (levo19.github.io/MOS/assets/):
`device-auth.js`, `seguridad-modal.js`, `extensor-horario.js`, `membrete-modal.js`,
`adhesivo/preview.js` — se versionan con `?v=` pineado en index.html.

## Arquitectura clave

- **Auth de dispositivo**: UUID en `localStorage.mosexpress_deviceId`, validado contra
  `mos.dispositivos` vía device-auth.js + Edge `mint-me` (JWT claim app='mosExpress').
  Sin token no hay RPCs (fail-closed, sin fallback).
- **Ventas**: RPCs `me.*` directas (escritura directa, correlativo pre-reservado server-side).
  Offline-first: `pendingSales` en localStorage, sync al volver la red.
- **Impresión**: Edge `imprimir` (ESC/POS + PrintNode). El estado real de impresora viene
  de `mePollEstadoImpresora` (webhook PrintNode vía RPC).
- **CPE (NubeFact)**: Edge `emitir-cpe`; reglas IGV/unidad vienen del catálogo canónico.
- **Catálogo**: `catalogo_pos_rls` + poller de `mos.catalogo_version()` (money-safe:
  no re-descarga con venta en curso).
- **Seguridad/horarios**: `SeguridadSystem.enforceBoot()` (anti-F5 con cache local
  `seg_fuera_cache_mosExpress`) + pantalla única de bloqueo con candado interactivo.
  Extensión de horario = `forzar_horario_hasta` (SQL 535; jamás usar
  `desbloqueo_temporal_hasta`, ese campo re-suspende al vencer).

## Reglas de trabajo OBLIGATORIAS

1. **Ediciones por STRINGS EXACTOS** con assert de unicidad (jamás regex glotones
   sobre bloques con llaves). Un regex `[\s\S]*?` ya tumbó la app una vez.
2. **Validar tras editar**: extraer los 11 scripts inline y `node --check` cada uno;
   verificar que el template no use claves fuera del return del setup (pantalla blanca).
3. **Ritual de versión en CADA deploy** (los 3 son obligatorios o hay bucle de update):
   - `sw.js` → `const VERSION`
   - `version.json` → `version`
   - `index.html` → `var V` (línea ~10)
4. **Dinero**: todo cálculo usa `_money()` (redondeo al asignar/acumular, no solo al
   mostrar). El formatter de tickets es `_moneyFmt` (devuelve string) — no confundir.
5. **Diálogos nativos prohibidos** (prompt/confirm/alert) → usar los modales del sistema.
6. Vue 3: el template NO ve claves `_xxx` (aliasear sin underscore en el return).
7. Deploy = commit + push a `main` (GitHub Pages, repo Levo19/MosExpress). No hay clasp.

## Verificación

Suites de BD (tx+ROLLBACK) en `ProyectoMOS/supabase/_test_*.js` (node + pg).
Tests de navegador real: `ProyectoMOS/browsercheck/check.js` con escenarios JSON
(sembrar SIEMPRE los devices fijos TEST-CLAUDE para no ensuciar `mos.dispositivos`).
