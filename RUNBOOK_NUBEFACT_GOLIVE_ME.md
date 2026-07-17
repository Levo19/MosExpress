# RUNBOOK — Go-live NubeFact producción (MosExpress)

Estado al 2026-07-17: **todo cableado**. Para pasar de demo a producción SOLO hay que pegar el token real.
El resto (emisión, anulación con reversa de pago, reconciliación, auto-baja) ya está desplegado y encendido.

---

## Paso único para producción: pegar el token real

```bash
supabase secrets set \
  NUBEFACT_RUTA="https://api.nubefact.com/api/v1/<TU-UUID-DE-PRODUCCION>" \
  NUBEFACT_TOKEN="<TU-TOKEN-DE-PRODUCCION>" \
  NUBEFACT_RUC="20610714057" \
  --project-ref rzbzdeipbtqkzjqdchqk
```

- Los secrets se leen en cada request → **no hace falta redeploy** de las Edge.
- `NUBEFACT_RUC` ya está seteado; solo cámbialo si el RUC emisor cambia.

## Validación inmediata (hazla apenas pegues el token)

1. En ME, emite **una BOLETA de prueba barata** (S/ 2).
2. El ticket debe salir con **QR de SUNAT** (no solo correlativo).
3. En 1 hora (o corriendo la reconciliación) debe pasar a **EMITIDO** con `nf_hash` y `nf_aceptada_sunat=true`.
4. Verifica en BD:
   ```sql
   select correlativo, nf_estado, nf_aceptada_sunat, (nf_hash<>'') tiene_hash
   from me.ventas where tipo_doc='BOLETA' order by created_at desc limit 3;
   ```
   Debe verse `EMITIDO / true / true`. Si sale `PENDIENTE` + "no existe" al consultar → el token/ruta no
   está registrando (revisar credenciales antes de operar en serio).

## Estado de flags (ya configurados — no tocar salvo emergencia)

| Flag | Valor | Rol |
|---|---|---|
| `ME_CPE_DIRECTO` | 1 | emitir CPE directo (cero-GAS) |
| `ME_ESCRITURA_DIRECTA` | 1 | escritura de ventas directa |
| `ME_IMPRESION_DIRECTA` | 1 | impresión por Edge |
| `ME_ANULAR_DIRECTO` | 1 | anulación directa (reversa de pago) |
| `CPE_RECON_ON` | 1 | reconciliación + **auto-baja** horaria |
| `FAC_CPE_DIRECTO` | 0 | otra vía (fac.*), inerte a propósito |

## Cómo funciona "Anular comprobante" (CPE)

Botón **📋 Anular comprobante** en el menú admin del ticket (BOLETA/FACTURA, MASTER-only):

1. **Pago:** SIEMPRE y al momento → `forma_pago='ANULADO'` (sale de caja) + repone stock + descuenta pickup.
   Atómico, idempotente, offline-safe (si no hay red, encola). → **la caja cuadra siempre.**
2. **SUNAT (automático):**
   - CPE **EMITIDO** → comunica la baja ya (`BAJA_ACEPTADA/SOLICITADA`).
   - CPE **PENDIENTE** → `ANULADO_PEND_BAJA`; la reconciliación manda la baja **sola** apenas SUNAT acepta.
   - CPE **RECHAZADO** → `ANULADO` (no hay nada que dar de baja).
   - **Sin internet** → basta con `forma_pago='ANULADO'`; la reconciliación (señal maestra) cierra el fiscal.

Piezas: Edge `emitir-cpe` op=baja · Edge `reconciliar-cpe` (auto-baja) · `me.cpe_recon_candidatos` +
`me.set_cpe_nf` (SQL 505) · front `adminConfirmarBaja` (reusa `me.anular_venta_directo` para el pago).

## Reconciliación / auto-baja manual (debug)

```bash
# secret del cron: select decrypted_secret from vault.decrypted_secrets where name='cpe_cron_secret';
curl -s -X POST "https://rzbzdeipbtqkzjqdchqk.supabase.co/functions/v1/reconciliar-cpe" \
  -H "Content-Type: application/json" -H "x-cpe-cron: <SECRET>" \
  -d '{"dias":45,"limite":90}'
# → {ok, revisados, emitidos, rechazados, bajas, agendadas, sin_cambio, detalle[]}
```
El cron `cpe-reconciliar` ya corre cada hora (minuto 23).
