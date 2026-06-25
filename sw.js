// ============================================================
// MOSexpress — Service Worker
// Cambia VERSION en cada deploy para invalidar caché
// ============================================================

// ── Firebase Cloud Messaging (background push) ─────────────
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey:            'AIzaSyA_gfynRxAmlbGgHWoioaj5aeaxnnywP88',
  projectId:         'proyectomos-push',
  messagingSenderId: '328735199478',
  appId:             '1:328735199478:web:947f338ae9716a7c049cd7'
});

const _fcmMsg = firebase.messaging();
_fcmMsg.onBackgroundMessage(payload => {
  // Comandos data-only (audio_start, audio_stop, gps_locate) → reenviar al cliente sin notificación
  if (payload.data && payload.data.action) {
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clients => {
      clients.forEach(c => c.postMessage({ type: 'mos_command', data: payload.data }));
    });
    return; // no mostrar notificación visible
  }
  const title = payload.notification?.title || 'MosExpress';
  const body  = payload.notification?.body  || '';
  // [Mensajería] Preservar idNotif/mensajeId para deep-link al hacer click.
  const pd = payload.data || {};
  self.registration.showNotification(title, {
    body,
    icon:    'https://levo19.github.io/MOS/icon-192.png',
    badge:   'https://levo19.github.io/MOS/icon-192.png',
    tag:     'me-push',
    vibrate: [200, 100, 200],
    data:    { idNotif: pd.idNotif || '', mensajeId: pd.mensajeId || (pd.extra && pd.extra.mensajeId) || null }
  });
});

// ── [Mensajería] Click en notificación → enfocar app + deep-link a la bandeja ──
self.addEventListener('notificationclick', event => {
  const d = (event.notification && event.notification.data) || {};
  event.notification.close();
  event.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const cmd = { type: 'mos_command', data: { action: 'me_deeplink', mensajeId: d.mensajeId || null } };
    if (all.length > 0) {
      const c = all[0];
      try { await c.focus(); } catch(_) {}
      try { c.postMessage(cmd); } catch(_) {}
    } else {
      // App cerrada: abrir + el polling de la bandeja levanta el badge igual
      try { await self.clients.openWindow('./'); } catch(_) {}
    }
  })());
});

// v2.8.24 — auth de dispositivos DIRECTO a Supabase (mos.verificar_dispositivo,
//           REST anon, app:'mosExpress') reusando la config Supabase que ME ya
//           tiene. Igual que WH. Doble-check + fallback a GAS intactos en device-auth.js v1.0.22.
// v2.8.32 — auto-refresco del catalogo: poller de mos.catalogo_version() money-safe
//           (solo visible, difiere si hay venta en curso, re-descarga sin reload).
// v2.8.38 — money-safety: idempotency key estable para guias manuales (idGuiaSnap en confirmarGuia
//           viaja en el payload; GAS registrarGuia/registrarGuiaAbierta respetan data.idGuia) →
//           los reintentos de _postGuiaBackground NO crean guias duplicadas → el cierre NO dobla stock.
// v2.8.40 — revision senior 40x ciclo guias: (1) reset duro del fill del hold-to-confirm tras un
//           cierre que falla (ya no queda la barra verde fantasma); (2) :key/seq en el banner undo
//           → la barra de 4s reinicia en borrados consecutivos; (3) intent-map TTL en el merge-guard
//           → la REAPERTURA optimista deja de ser revertida por un refresh disparado por otra accion
//           (simetrico con el cierre). Money-safe: el backend cerrar/reabrir sigue idempotente con lock.
const VERSION = '2.8.65';
const CACHE   = 'mosexpress-v' + VERSION;
const ASSETS  = [
  './',
  './index.html',
  './radio.html',
  './manifest.json',
  './version.json',
  'https://unpkg.com/vue@3.4.21/dist/vue.global.prod.js',
  'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

// ── Instalar: cachear secuencial con reporte de progreso + skipWaiting ──
// postMessage al cliente por cada asset → banner muestra barra real.
// skipWaiting al final: el SW nuevo se activa de inmediato cuando termina
// de instalar (combinado con clients.claim en activate, toma control de
// las pestañas abiertas sin necesidad de cerrar todo). Antes esperábamos
// que el usuario cerrara todo → updates se atascaban días. Cambio para
// que pushes lleguen a los cajeros al primer refresh.
self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    const total = ASSETS.length;
    let done = 0;
    async function _broadcast(payload) {
      const cs = await self.clients.matchAll({ includeUncontrolled: true, type: 'window' });
      cs.forEach(c => { try { c.postMessage(payload); } catch(_){} });
    }
    await _broadcast({ type: 'sw-install-progress', done: 0, total, version: VERSION });
    // Timeout duro por asset — si la red está lenta o el CDN se cuelga,
    // no dejamos que el install se atore eternamente.
    const _withTimeout = (p, ms, label) => Promise.race([
      p,
      new Promise((_, rej) => setTimeout(() => rej(new Error('timeout ' + label)), ms))
    ]);
    for (const url of ASSETS) {
      try {
        await _withTimeout(cache.add(new Request(url, { cache: 'no-store' })), 15000, url);
      } catch (err) { console.warn('[SW ME] No se pudo cachear:', url, err); }
      done++;
      await _broadcast({ type: 'sw-install-progress', done, total, version: VERSION });
    }
    await _broadcast({ type: 'sw-install-done', total, version: VERSION });
    // Activar de inmediato (clients.claim en activate toma las pestañas abiertas)
    self.skipWaiting();
  })());
});

// ── Activar: borrar cachés viejos y reclamar clientes ───────
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// ── Fetch: estrategia híbrida (network-first crítico, cache-first assets) ─
// [v2.5.53] Network-first con timeout 2.5s para HTML/JS críticos. Resuelve
// el dolor histórico de "deployé v.X pero el SW sirve v.X-2 cacheado por
// horas". Ahora cuando deployo, en el siguiente refresh la versión nueva
// llega de inmediato (siempre que haya red — si offline, fallback a cache).
// Para imágenes/fonts/manifest seguimos cache-first (cambian poco y mejora
// performance percibida en arranque offline).
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // No interceptar GAS ni PrintNode (delegamos al fetch nativo)
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('printnode.com')) return;

  // version.json: siempre desde red (detecta nuevas versiones rápido)
  if (url.pathname.endsWith('version.json')) {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    return;
  }

  const path = url.pathname;
  const esCritico =
    path === '/' ||
    path.endsWith('/') ||
    path.endsWith('.html') ||
    path.endsWith('.js') ||
    path.endsWith('manifest.json');

  if (esCritico && url.origin === self.location.origin) {
    // Network-first con timeout 2.5s → cache fallback
    e.respondWith((async () => {
      try {
        const netPromise = fetch(e.request);
        const timeout = new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 2500));
        const res = await Promise.race([netPromise, timeout]);
        if (res && res.status === 200 && res.type !== 'opaque') {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone)).catch(() => {});
        }
        return res;
      } catch(_) {
        const cached = await caches.match(e.request);
        if (cached) return cached;
        // Último recurso: red sin timeout
        return fetch(e.request).catch(() => Response.error());
      }
    })());
    return;
  }

  // Cache-first para assets estáticos (imágenes, fonts, CDN externos cacheados)
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (!res || res.status !== 200) return res;
        if (res.type !== 'basic' && res.type !== 'cors') return res;
        if (e.request.method !== 'GET') return res;   // [FIX] Cache.put solo soporta GET (HEAD/POST lanzan)
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone)).catch(() => {});  // defensivo: nunca uncaught
        return res;
      }).catch(() => Response.error());
    })
  );
});

// ── Mensaje SKIP_WAITING desde la app ───────────────────────
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});
