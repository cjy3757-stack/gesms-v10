const CACHE='gesms-v10-6-stats-dashboard-20260724-1';
const STATIC=['./','./index.html','./data.js','./service-data.js','./office-data.js','./office-excel.js','./prayer-excel.js','./jszip.min.js','./member-excel.js','./stats-v10.6.js','./finance-data.js','./manifest.webmanifest','./logo.png','./icon-192.png','./icon-512.png'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(STATIC)));self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.map(key=>key===CACHE?Promise.resolve():caches.delete(key)))));self.clients.claim();});
self.addEventListener('fetch',event=>{
 if(event.request.method!=='GET')return;
 const url=new URL(event.request.url);
 const alwaysFresh=/\.(xlsx|csv|json)$/i.test(url.pathname)||['/','/index.html','/member-excel.js','/stats-v10.6.js','/office-excel.js','/office-data.js','/prayer-excel.js','/sw.js'].some(x=>url.pathname.endsWith(x));
 if(alwaysFresh){
   event.respondWith(fetch(event.request,{cache:'no-store'}).then(response=>{
     if(response.ok){const copy=response.clone();caches.open(CACHE).then(cache=>cache.put(event.request,copy));}
     return response;
   }).catch(()=>caches.match(event.request)));
   return;
 }
 event.respondWith(caches.match(event.request).then(hit=>hit||fetch(event.request).then(response=>{
   if(response.ok){const copy=response.clone();caches.open(CACHE).then(cache=>cache.put(event.request,copy));}
   return response;
 })));
});
