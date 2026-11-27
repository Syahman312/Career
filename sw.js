const C='ladder-v9',A=['./','index.html','manifest.json','icon.svg','icon-192.png','icon-512.png','apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(A)));self.skipWaiting()});
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>clients.claim())));
const put=(q,r)=>{if(r&&r.ok){const p=r.clone();caches.open(C).then(c=>c.put(q,p))}return r};
const old=q=>caches.match(q,{ignoreSearch:true}).then(r=>r||caches.match('index.html'));
self.addEventListener('fetch',e=>{const q=e.request,u=new URL(q.url);if(q.method!=='GET'||u.origin!==location.origin)return;
// icons and the PDF reader rarely change: serve from cache, fetch only if missing
if(/\.(png|svg|mjs)$/.test(u.pathname)){e.respondWith(caches.match(q,{ignoreSearch:true}).then(r=>r||fetch(q).then(x=>put(q,x))));return}
// the app itself: newest from the network, but fall back to the saved copy after 3 s on a slow connection
e.respondWith(new Promise(ok=>{let done=0;const f=fetch(q).then(r=>put(q,r));
const t=setTimeout(()=>old(q).then(r=>{if(r&&!done){done=1;ok(r)}}),3000);
f.then(r=>{clearTimeout(t);if(!done){done=1;ok(r)}}).catch(()=>{clearTimeout(t);if(!done){done=1;ok(old(q))}})}))});
