const CACHE='gesms-v10-rc2-20260723';
const STATIC=['./','./index.html','./data.js','./service-data.js','./office-data.js','./office-excel.js','./prayer-excel.js','./jszip.min.js','./finance-data.js','./manifest.webmanifest','./logo.png','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(STATIC)));self.skipWaiting();});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));self.clients.claim();});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  const u=new URL(e.request.url);
  const dynamic=/\.(xlsx|csv|json)$/i.test(u.pathname)||u.pathname.endsWith('/index.html')||u.pathname.endsWith('/');
  if(dynamic){e.respondWith(fetch(e.request,{cache:'no-store'}).catch(()=>caches.match(e.request)));return;}
  e.respondWith(caches.match(e.request).then(hit=>hit||fetch(e.request).then(r=>{if(r.ok){const copy=r.clone();caches.open(CACHE).then(c=>c.put(e.request,copy));}return r;})));
});
