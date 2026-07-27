const CACHE='gesms-v10-9-3-v31-clean-source-20260728-1';
const STATIC=['./','./index.html','./version.js','./data.js','./service-data.js','./office-data.js','./office-excel.js','./prayer-excel-v10.8.js','./jszip.min.js','./member-excel.js','./stats-v10.7.3.js','./finance-data.js','./manifest.webmanifest','./logo.png','./icon-192.png','./icon-512.png'];
self.addEventListener('install',event=>{event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(STATIC)));self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.map(key=>key===CACHE?Promise.resolve():caches.delete(key)))));self.clients.claim();});
self.addEventListener('fetch',event=>{
 if(event.request.method!=='GET')return;
 const url=new URL(event.request.url);
 // Excel은 반드시 GitHub의 최신 원본만 사용합니다. 이전 캐시로 되돌아가지 않습니다.
 if(/\.xlsx$/i.test(url.pathname)){
   event.respondWith(fetch(event.request,{cache:'no-store'}));
   return;
 }
 const alwaysFresh=/\.(csv|json)$/i.test(url.pathname)||['/','/index.html','/member-excel.js','/stats-v10.7.3.js','/office-excel.js','/office-data.js','/prayer-excel-v10.8.js','/sw.js'].some(x=>url.pathname.endsWith(x));
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
