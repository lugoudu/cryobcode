const CACHE='cbpwa-6abb5fc9';
const ASSETS=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png','./apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil((async()=>{const c=await caches.open(CACHE);await c.addAll(ASSETS);await self.skipWaiting()})())});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{const ks=await caches.keys();await Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)));await self.clients.claim()})())});
// 缓存优先（打开永远秒开、断网可用）+ 后台静默更新（联网时拉新版，下次打开生效）
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  e.respondWith((async()=>{
    const c=await caches.open(CACHE);
    const cached=await c.match(e.request,{ignoreSearch:e.request.mode==='navigate'});
    const refresh=fetch(e.request).then(res=>{
      if(res&&(res.status===200||res.type==='opaque')) c.put(e.request,res.clone());
      return res;
    }).catch(()=>null);
    if(cached) return cached;
    const res=await refresh;
    if(res) return res;
    return c.match('./index.html');
  })());
});