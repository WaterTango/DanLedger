const C='ledger-v27';
const A=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png','fonts/dm-sans-latin-400-normal.woff2','fonts/dm-sans-latin-500-normal.woff2','fonts/dm-sans-latin-600-normal.woff2','fonts/manrope-latin-400-normal.woff2','fonts/manrope-latin-500-normal.woff2','fonts/manrope-latin-600-normal.woff2','fonts/inter-latin-400-normal.woff2','fonts/inter-latin-500-normal.woff2','fonts/inter-latin-600-normal.woff2','fonts/poppins-latin-400-normal.woff2','fonts/poppins-latin-500-normal.woff2','fonts/poppins-latin-600-normal.woff2','fonts/manrope-latin-700-normal.woff2','fonts/poppins-latin-700-normal.woff2'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(A)));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))));self.clients.claim()});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET'||new URL(e.request.url).origin!==location.origin)return;
  e.respondWith(fetch(e.request).then(r=>{const cp=r.clone();caches.open(C).then(c=>c.put(e.request,cp));return r}).catch(()=>caches.match(e.request).then(m=>m||caches.match('index.html'))));
});
