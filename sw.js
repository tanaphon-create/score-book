// แคชให้เปิดเร็วและใช้ออฟไลน์ได้: หน้าเว็บ = ลองเน็ตก่อน (ได้เวอร์ชันใหม่ทันที) ไม่ได้ค่อยใช้ที่เก็บไว้ ; ฟอนต์/ไลบรารี = ใช้ที่เก็บไว้ก่อน
const CACHE = 'scorebook-v2';
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'icon.svg', 'manifest.webmanifest']).catch(()=>{}))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.hostname === 'script.google.com' || u.hostname.endsWith('googleusercontent.com')) return; // การซิงก์ไม่แคช
  const same = u.origin === location.origin;
  if (same) {
    e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(CACHE).then(c => c.put(r, cp)); return res; }).catch(() => caches.match(r, {ignoreSearch:true})));
  } else {
    e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => { const cp = res.clone(); caches.open(CACHE).then(c => c.put(r, cp)); return res; })));
  }
});
