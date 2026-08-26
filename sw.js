const CACHE='gesms-v11-1-4-project-completion-rate-20260826';
const STATIC=['./','./index.html','./version.js','./data.js','./service-data.js','./office-data.js','./prayer-static-data.js','./office-excel.js','./prayer-excel-v10.8.js','./jszip.min.js','./member-excel.js','./stats-v10.7.3.js','./finance-data.js','./manifest.webmanifest','./logo.png','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(STATIC)));self.skipWaiting();});
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
async function networkFirst(req){const c=await caches.open(CACHE);try{const r=await fetch(req,{cache:'no-store'});if(r&&r.ok)c.put(req,r.clone());return r}catch(e){return (await c.match(req,{ignoreSearch:true}))||c.match('./index.html')}}
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  const u=new URL(e.request.url);
  if(e.request.mode==='navigate'){e.respondWith(networkFirst(e.request));return;}

  const dynamicData=/\.(xlsx|csv|json)$/i.test(u.pathname)||
    /\/(data|service-data|office-data|prayer-static-data|finance-data|version)\.js$/i.test(u.pathname);

  if(dynamicData){
    e.respondWith(
      fetch(e.request,{cache:'no-store'})
        .then(async r=>{
          if(r&&r.ok)(await caches.open(CACHE)).put(e.request,r.clone());
          return r;
        })
        .catch(()=>caches.match(e.request,{ignoreSearch:true}))
    );
    return;
  }

  e.respondWith(
    caches.match(e.request,{ignoreSearch:true})
      .then(r=>r||fetch(e.request).then(async x=>{
        if(x&&x.ok)(await caches.open(CACHE)).put(e.request,x.clone());
        return x;
      }))
  );
});
