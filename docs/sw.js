// Service Worker — CRM Ti'Piedade
const CACHE = 'tipiedade-crm-v1';

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache =>
      cache.addAll(['/tipiedade-leads/', '/tipiedade-leads/index.html'])
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(clients.claim());
});

self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});

// ── Notificações push ────────────────────────────────────────────
self.addEventListener('push', e => {
  const data = e.data?.json() || {};
  e.waitUntil(
    self.registration.showNotification(data.title || "CRM Ti'Piedade", {
      body:    data.body  || "Tens leads para visitar hoje.",
      icon:    '/tipiedade-leads/icons/icon-192.png',
      badge:   '/tipiedade-leads/icons/icon-192.png',
      tag:     data.tag   || 'crm-notif',
      data:    data.url   || '/tipiedade-leads/',
      actions: data.actions || [
        { action: 'ver', title: '📋 Ver leads' },
        { action: 'dispensar', title: 'Dispensar' }
      ],
      vibrate:   [200, 100, 200],
      requireInteraction: true
    })
  );
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  if (e.action === 'dispensar') return;
  e.waitUntil(
    clients.matchAll({ type: 'window' }).then(wins => {
      if (wins.length) return wins[0].focus();
      return clients.openWindow(e.notification.data || '/tipiedade-leads/');
    })
  );
});

// ── Notificação agendada (verificar diariamente) ──────────────────
self.addEventListener('periodicsync', e => {
  if (e.tag === 'crm-daily') {
    e.waitUntil(verificarNotificacaoDiaria());
  }
});

async function verificarNotificacaoDiaria() {
  const hora = new Date().getHours();
  if (hora < 8 || hora > 19) return;
  await self.registration.showNotification("CRM Ti'Piedade — Bom dia!", {
    body: "Tens leads para visitar hoje. Vê as que estão prontas para visita.",
    icon: '/tipiedade-leads/icons/icon-192.png',
    tag:  'crm-daily',
    data: '/tipiedade-leads/',
    requireInteraction: false
  });
}
