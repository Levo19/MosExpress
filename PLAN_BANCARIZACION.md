# PLAN — Bancarización en ME (pagos ≥ S/ 2,000)

> Documento vivo. Estado: **IMPLEMENTADO (F1–F4) + review 500x + fixes. LOCAL, sin desplegar.**
> Falta: prueba con boleta REAL/de-prueba antes de usar en caja (emite CPE fiscal) → luego deploy.
> Toda modificación futura actualiza este archivo.

## ESTADO (2026-07-16)
- ✅ **F1** Config MOS ▸ 🏦 Bancarios (bancos/Yape/Plin + check imprime + QR + límite) · `me.get_medios_cobro` (SQL 503/504) — probado (UI + round-trip).
- ✅ **F2** Modal gate ≥ límite (A/B/⟵) en `procesarPago` — compila, ME arranca.
- ✅ **F4a** Opción A: fuerza VIRTUAL + imprime ticket de cuenta — verificado correcto (500x).
- ✅ **F3/F4b** Opción B: `partirCarritoBanc` (validado) + emisión N CPE con blindaje de fallo parcial (para + deja faltantes en carrito) + guard doble-tap.
- ✅ **Review 500x**: happy-path correcto; 1 hueco fiscal (#1) + 2 endurecimientos (#2,#4) corregidos.
- ⏳ **Pendiente**: probar A y B con boleta de prueba (emite fiscal real) → luego deploy MOS+ME.

Archivos tocados (local, uncommitted): `ProyectoMOS/js/app.js` `js/api.js` `index.html` + `supabase/503,504.sql` (aplicado, aditivo/inerte); `MosExpress/index.html`.

## 0. Grounding (código real verificado)

- **ME emite CPE fiscal REAL hoy.** Flags en `mos.config` todos en `'1'`: `ME_CPE_DIRECTO`, `ME_ESCRITURA_DIRECTA`, `ME_IMPRESION_DIRECTA`. Emisión: `_crearCPEDirecto(ventaBase, localId)` (MosExpress/index.html:17266) → RPC `me.crear_cpe_directo` (mintea correlativo atómico + resuelve serie desde `mos.series_documentales`) → Edge `emitir-cpe` (NubeFact, token en secret) → `set_cpe_nf` (QR/hash/estado).
- **Gancho de bancarización YA existe**: `granTotal.value >= 2000 && metodo==='EFECTIVO'` en ~19098 (hoy un `confirm()` nativo → se reemplaza por modal; regla: prohibido confirm/alert).
- **Carrito**: `carrito.value` (ref). Item: `{sku, skuBase, factor, codBarras, nombre, codSunat, tipoIgv, unidadMedida('KGM'=granel), precioUnitario, cantidad(kg si granel), subtotal, segmentosPrecio, segmentoAplicado, segmentoAjustePct, ...}`. Total: `granTotal` = Σ subtotal.
- **Formas de pago**: `EFECTIVO / VIRTUAL / MIXTO / CREDITO / POR_COBRAR` (VIRTUAL agrupa Yape/Plin/transferencia; no hay método bancario discreto).
- **Doc**: `pago.tipoDoc` = `NOTA_DE_VENTA / BOLETA / FACTURA`. Cliente en `pago.docCliente/nombreCliente/direccionCliente`.
- **Granel/tramos**: `_meCalcPrecioGranel(precioCanonico, gramos, segmentos)` (18268) + `calcularImportes(item)` (18298) → setea `subtotal`, `segmentoAplicado`, `segmentoAjustePct`. Precio efectivo del tramo comprado = `item.subtotal / item.cantidad`.
- **Config cross-app**: molde `me.get_tarjeta_config()` (lee claves de `mos.config`, anon) y `me.empresa_fiscal` (cacheado en ME por `_cargarEmpresaFiscal`). Datos de empresa emisora en `fac.config` (RUC 20610714057, INVERSIONES MOS EIRL). **No hay datos bancarios de la empresa** (solo de proveedores) → campo virgen.
- **Config MOS**: módulo `#view-config`, tabs `['infra','personal','categorias','notifs']`, `renderCfgTab`. Escritura config: `API.post('setConfig',{clave,valor})` → `mos.set_config` (cero-GAS). Kit UI `.fac-*` + patrón `facRenderConfig`/`facGuardarConfig`, gate `_esAdminOMaster()`.
- **Impresión**: `imprimirTicketVenta(correlativo, sync, nfData)` (19264, ESC/POS W=48) + `mandarImpresionPrintNode(printerId, titulo, b64ESC(txt))` (13479) / Edge `imprimir`.

## 1. Regla tributaria (qué dispara la alerta)

- **Umbral: pago ≥ S/ 2,000** (o US$ 500) por **operación** → obliga medio de pago bancario (Ley 28194, vigente desde 01/04/2022; antes S/ 3,500).
- **Aplica a BOLETA y FACTURA** (no depende del tipo de CPE, sino del monto).
- **El voucher puede ser en varios medios** (ej. 7 Yapes de S/ 500, incluso de distintas cuentas): lo que importa es que el TOTAL quede cubierto con medios de pago y sea trazable. El CPE siempre dice el total real.
- **Split (opción B) = riesgo de "fraccionamiento"** ante SUNAT (mira la operación real). Se ofrece como decisión informada; la vía limpia es bancarizar (opción A).

## 2. FASE 1 — Config MOS: "Medios de cobro de la empresa" [riesgo bajo]

**UI**: nueva sección en `#view-config` (card en Infra o tab nuevo `bancarios`), kit `.fac-*`, gate `_esAdminOMaster()`.
- Lista editable de **medios**: botones "+ Cuenta bancaria", "+ Yape", "+ Plin". Cada medio:
  - **Banco**: `{banco, nroCuenta, cci, titular, moneda}` + **☑ imprime**.
  - **Yape/Plin**: `{titular, telefono, qrUrl?}` + **☑ imprime**.
- **☑ imprime** por medio (varios pueden imprimirse). Botón quitar por medio. Botón Guardar.
- **QR (Yape/Plin)**: se **sube imagen** (la que da tu app Yape/Plin) → Storage → `qrUrl`. Se muestra en pantalla y/o se imprime.

**Persistencia**: 1 clave en `mos.config`: **`EMPRESA_MEDIOS_COBRO`** = JSON:
```json
[{ "id":"m1","tipo":"banco","banco":"BCP","nroCuenta":"191-...","cci":"002191...","titular":"INVERSIONES MOS EIRL","moneda":"PEN","imprime":true },
 { "id":"m2","tipo":"yape","titular":"...","telefono":"9xxxxxxxx","qrUrl":"https://.../yape.png","imprime":true }]
```
Guardado con `API.post('setConfig',{clave:'EMPRESA_MEDIOS_COBRO', valor: JSON.stringify(arr)})`. **Sin cambios en api.js/SQL de escritura** (usa `mos.set_config`). El nombre de clave no contiene pin/secret/token → `config_publico` la expone; igual se lee por RPC dedicada (Fase 1b).

**Fase 1b — Lectura en ME**: RPC nueva `me.get_medios_cobro()` (clon de `me.get_tarjeta_config`) que devuelve `EMPRESA_MEDIOS_COBRO` (+ los `EMPRESA_*` fiscales por conveniencia). Grant `anon, authenticated`. ME: `_cargarMediosCobro()` (clon de `_cargarEmpresaFiscal`) → `mediosCobro` ref + cache localStorage.

## 3. FASE 2 — Modal de bancarización (gate ≥ S/ 2,000) [riesgo medio]

Reemplaza el `confirm()` de ~19098. **Modal Vue optimista** (sin confirm/alert; sonido+visual). Es un **gate previo a la emisión**: nada se emite hasta elegir A o B.
- **Dispara** si `granTotal >= 2000` y `tipoDoc ∈ {BOLETA, FACTURA}` y el pago NO es ya 100% bancarizado (si el operador ya eligió VIRTUAL total, se puede saltar el gate — es decisión: por ahora **siempre** mostrar el modal si ≥2000 para dejar registro de la decisión).
- **Contenido**:
  - Título: "Pago de S/ X — requiere bancarización (≥ 2,000)".
  - **[A] 💳 Bancarizar** → pago **VIRTUAL obligatorio**; imprime ticket de cuenta; 1 CPE.
  - **[B] ✂️ Partir en varios CPE (< 2,000)** → parte la venta; cada CPE paga efectivo/virtual como siempre.
  - **⟵ flecha volver** (sin "Cancelar"): cierra el modal, vuelve al carrito, **no emite nada**.
- Preview de la partición en el propio modal (opción B): muestra los N tickets y sus totales antes de confirmar.

## 4. FASE 3 — Opción A (bancarizar) [riesgo medio]

- **Fuerza** `pago.metodo = 'VIRTUAL'` (bloquea efectivo para esta venta).
- Emite **1 CPE** (flujo existente `_crearCPEDirecto`), forma pago VIRTUAL.
- **Imprime ticket extra "DATOS PARA TRANSFERIR"** (nuevo builder `_ticketMediosCobro(medios, monto, ref)`, ESC/POS W=48):
  - Encabezado empresa (razón social + RUC).
  - Por cada medio con `imprime:true`: banco/cuenta/CCI/titular (texto) · Yape/Plin: titular + teléfono.
  - Monto: S/ X · Ref: serie-correlativo del CPE.
  - **QR Yape/Plin**: se **muestra en pantalla** (en el modal/pantalla del cajero) para que el cliente escanee; el ticket lleva el texto + "escanea el QR en pantalla". *(No se imprime la imagen QR en v1 — rasterizar imagen a ESC/POS es complejo; el QR en pantalla es más práctico.)*
- El cliente transfiere; queda registrado el medio (VIRTUAL).

## 5. FASE 4 — Opción B (partir en N CPE) [riesgo alto · fiscal]

**Algoritmo puro** `_partirCarrito(items, limite=2000, pasoGranelKg=0.1)` → devuelve `subCarritos[]` (cada uno con sus items/fragmentos y `total < limite`):
- Recorre items **en orden**. Mantiene `subActual` con su acumulado.
- Por ítem:
  - **NIU (unidades)**: `caben = floor((limite - acumulado) / precioUnit)`.
    - `caben >= cantidad` → mete completo.
    - `0 < caben < cantidad` → mete `caben` unidades acá; abre sub nuevo; sigue con `cantidad-caben`.
    - `caben == 0` → cierra sub (si tiene algo), abre nuevo, reintenta.
  - **KGM (granel)**: precio efectivo `pe = item.subtotal / item.cantidad` (**tramo congelado**). `kgCaben = floor((limite - acumulado) / pe / pasoGranel) * pasoGranel` (pasos 1kg→0.5→0.1). Mete `kgCaben` con precio `pe` (subtotal = kgCaben*pe); resto al sub siguiente. **El % de tramo (segmentoAjustePct) viaja con cada fragmento** (no revive tramo menor).
  - **Borde**: si 1 unidad o el paso mínimo de granel ya ≥ limite → ese fragmento no baja del umbral → marca `requiereBancarizacion:true` en el resultado → el modal avisa "no se puede partir bajo 2,000, hay que bancarizar".
- Cada fragmento **preserva los campos fiscales** del item (sku, skuBase, codSunat, tipoIgv, unidadMedida, factor) y su precio efectivo.
- **Invariante de dinero**: `Σ totales(subCarritos) === granTotal` (control anti-drift de redondeo; si hay diferencia de centavos por redondeo de IGV, ajustar el último fragmento).

**Emisión N CPE**:
- Por cada `subCarrito`: arma `ventaBase` (mismo cliente, mismo tipoDoc, sus items) y llama `_crearCPEDirecto` con **`ref_local` único** (`<localId>-p1`, `-p2`, …) → **idempotente** (reintento no duplica). Cada uno mintea su propio correlativo/serie.
- Forma de pago de cada sub-CPE = la que eligió el operador (efectivo o virtual); cada uno < 2,000.
- **Emisión secuencial con progreso** (modal "emitiendo 2/3…"); si uno falla, los ya emitidos quedan (reales), se permite **reintentar los faltantes** (idempotente). No se pierde ni se duplica.
- Imprime **N tickets** (uno por CPE).

## 6. Idempotencia y seguridad fiscal (obligatorio)

- Nada se emite hasta confirmar A/B (el modal es el gate).
- N-CPE: N `ref_local` distintos → reintento seguro (dedup server por ref_local, patrón existente).
- Money-safe: `Σ sub == total`; IGV recomputado por sub-CPE; ajuste de centavos en el último.
- No tocar los flags fiscales; ME ya emite real.

## 7. Caveats (documentados, decisión del dueño)

- Opción B = fraccionamiento → riesgo SUNAT. Opción A (bancarizar) es la limpia y va primero.
- Emisión es fiscal REAL → **probar en boletas/facturas de prueba antes de usar en caja**. No se activa nada nuevo; se usa el path vivo.

## 8. Orden de implementación

1. **F1** Config MOS (medios de cobro) + `me.get_medios_cobro` + `_cargarMediosCobro` en ME.
2. **F3** Algoritmo `_partirCarrito` (función pura, con pruebas unitarias: caso ajinomoto 20×150, multi-producto, granel con tramo, borde ≥límite).
3. **F2** Modal bancarización (gate ≥2000, A/B, ⟵, preview).
4. **F4a** Opción A (forzar VIRTUAL + ticket de cuenta).
5. **F4b** Opción B (emisión N CPE con idempotencia + progreso + reimpresión).
6. Revisión 500x (bugs, dinero, duplicados, obsoletos) + pruebas.

## 9bis. Afinamientos del review 100x (huecos cerrados)

- **Tope estricto**: cada sub-CPE debe quedar **< S/ 2,000** (no "≤"). Margen configurable `LIMITE_BANCARIZACION` en `mos.config` (default 2000; el corte usa `< 2000.00`).
- **Cap de particiones**: si la venta requiere **más de N sub-CPE** (default 6) para bajar del umbral (típico en granel partido en pedacitos), el modal **avisa**: "se necesitarían X documentos — mejor bancarizar". Evita el caso "100 boletas de 50g" (impráctico + muy expuesto + 100 emisiones NubeFact).
- **Solo BOLETA/FACTURA**: el gate NO dispara en `NOTA_DE_VENTA` (NV no es CPE fiscal). *(Decisión del dueño: si quiere incluir NV, es 1 línea.)*
- **MIXTO en split**: si el operador tenía `MIXTO`, cada sub-CPE se cobra con el método simple que corresponda (o se pide elegir efectivo/virtual una vez para todo el split). No se reparte el mixto por documento (complejidad innecesaria).
- **Gate y pago ya-virtual**: si el operador YA eligió `VIRTUAL` al 100% (venta ya bancarizada), igual se muestra el modal para dejar registro, con la opción A pre-seleccionada. *(Decisión: se puede saltar si molesta.)*
- **QR v1 vs v2**: v1 = QR Yape/Plin **en pantalla** para escanear + datos en texto en el ticket. v2 (si lo pides) = rasterizar la imagen QR al ticket ESC/POS.
- **Independencia de los N CPE**: cada sub-CPE es un documento fiscal independiente (se puede anular/reimprimir por separado con el flujo existente).
- **Redondeo IGV**: al partir, la suma de sub-totales debe igualar el total original; si el redondeo de IGV por documento genera diferencia de céntimos, se ajusta el ÚLTIMO fragmento (invariante `Σ = total`).

## 9. Sobre el QR (confirmado)

No existe QR "universal" que transfiera a una cuenta bancaria con solo el CCI. El QR útil es el de **Yape/Plin** (imagen de tu app; con interoperabilidad 2023+ lo paga también Plin/banco). Diseño: **cuenta bancaria = texto (cuenta/CCI); Yape/Plin = imagen QR opcional que subes** y se muestra en pantalla para escanear.
