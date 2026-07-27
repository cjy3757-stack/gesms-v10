const CACHE='gesms-v10-9-6-mobile-excel-recovery-20260728-1';
const STATIC=[
 './','./index.html','./version.js','./data.js','./service-data.js','./office-data.js','./office-excel.js',
 './prayer-excel-v10.8.js','./jszip.min.js','./member-excel.js','./stats-v10.7.3.js','./finance-data.js',
 './manifest.webmanifest','./logo.png','./icon-192.png','./icon-512.png',
 './공지사항.xlsx','./교회일정.xlsx','./notice.xlsx','./events.xlsx',
 './GESMS_전도영혼구원_표준자료양식_V3.1_버전자동연동.xlsx','./기도_식사_전도_봉사일정.xlsx'
];
self.addEventListener('install',event=>{
 event.waitUntil(caches.open(CACHE).then(async cache=>{
   for(const item of STATIC){try{await cache.add(item)}catch(e){/* 선택 파일 누락은 설치를 막지 않음 */}}
 }));
 self.skipWaiting();
});
self.addEventListener('activate',event=>{
 event.waitUntil(caches.keys().then(keys=>Promise.all(keys.map(key=>key===CACHE?Promise.resolve():caches.delete(key)))));
 self.clients.claim();
});
async function networkFirst(request){
 const cache=await caches.open(CACHE);
 try{
   const response=await fetch(request,{cache:'no-store'});
   if(response&&response.ok)await cache.put(request,response.clone());
   return response;
 }catch(error){
   const cached=await cache.match(request,{ignoreSearch:true});
   if(cached)return cached;
   throw error;
 }
}
self.addEventListener('fetch',event=>{
 if(event.request.method!=='GET')return;
 const url=new URL(event.request.url);
 if(/\.(xlsx|csv|json)$/i.test(url.pathname)){
   event.respondWith(networkFirst(event.request));
   return;
 }
 const alwaysFresh=['/','/index.html','/version.js','/member-excel.js','/stats-v10.7.3.js','/office-excel.js','/office-data.js','/prayer-excel-v10.8.js','/sw.js'].some(x=>url.pathname.endsWith(x));
 if(alwaysFresh){event.respondWith(networkFirst(event.request));return;}
 event.respondWith(caches.match(event.request,{ignoreSearch:true}).then(hit=>hit||fetch(event.request).then(async response=>{
   if(response&&response.ok){const cache=await caches.open(CACHE);await cache.put(event.request,response.clone())}
   return response;
 })));
});
