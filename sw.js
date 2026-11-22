/* Service Worker – מקבל התראות מהשרת ומציג אותן בטלפון גם כשהאתר סגור */
var API = 'https://script.google.com/macros/s/AKfycbxx1JL_72prP_jSFRQzfS3mdFycX1QDMK7Udg58ApQ7if3GptH0eQyUbRFwSElW9VU/exec';
var ICON = 'icon-192.png';

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });

function getToken() {
  return caches.open('cb-auth').then(function (c) { return c.match('token'); }).then(function (r) { return r ? r.text() : null; });
}

self.addEventListener('push', function (e) {
  e.waitUntil(getToken().then(function (token) {
    if (!token) return [];
    return fetch(API, { method: 'POST', body: JSON.stringify({ action: 'notices', token: token }) })
      .then(function (r) { return r.json(); })
      .then(function (j) { return j.ok ? (j.notices || []) : []; })
      .catch(function () { return []; });
  }).then(function (list) {
    if (!list.length) {
      return self.registration.showNotification('שריון רכבי מחלקה', {
        body: 'יש עדכון חדש באתר', icon: ICON, badge: ICON, tag: 'cb-generic', dir: 'rtl', lang: 'he', data: { view: 'mine' }
      });
    }
    return Promise.all(list.map(function (n) {
      return self.registration.showNotification(n.title || 'שריון רכבי מחלקה', {
        body: n.text, icon: ICON, badge: ICON, tag: n.id, dir: 'rtl', lang: 'he', data: { view: n.view || 'mine' }
      });
    }));
  }));
});

self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  var view = (e.notification.data && e.notification.data.view) || 'mine';
  var url = new URL('./?view=' + view, self.registration.scope).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
    for (var i = 0; i < list.length; i++) {
      if (list[i].url.indexOf(self.registration.scope) === 0) {
        list[i].postMessage({ type: 'open', view: view });
        return list[i].focus();
      }
    }
    return self.clients.openWindow(url);
  }));
});
