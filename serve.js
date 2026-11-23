// Tiny static server for local testing: node serve.js [root] [port]
const http=require('http'),fs=require('fs'),path=require('path');
const root=path.resolve(process.argv[2]||'.'),port=+process.env.PORT||+process.argv[3]||8765;
const T={'.html':'text/html; charset=utf-8','.js':'text/javascript','.json':'application/json','.png':'image/png','.svg':'image/svg+xml','.mjs':'text/javascript'};
http.createServer((q,r)=>{let p=decodeURIComponent(q.url.split('?')[0]);if(p.endsWith('/'))p+='index.html';
const f=path.join(root,path.normalize(p));if(f!==root&&!f.startsWith(root+path.sep)){r.writeHead(403);return r.end()}
fs.readFile(f,(e,d)=>{if(e){r.writeHead(404);return r.end('not found')}r.writeHead(200,{'Content-Type':T[path.extname(f)]||'application/octet-stream','Cache-Control':'no-cache'});r.end(d)})}).listen(port,()=>console.log('on '+port));
