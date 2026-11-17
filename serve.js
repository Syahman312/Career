// Tiny static server for local testing.
const http=require('http'),fs=require('fs'),path=require('path');
const root=process.argv[2],port=+process.env.PORT||+process.argv[3]||8765;
const T={'.html':'text/html','.js':'text/javascript','.json':'application/json','.png':'image/png'};
http.createServer((q,r)=>{let p=decodeURIComponent(q.url.split('?')[0]);if(p.endsWith('/'))p+='index.html';
const f=path.join(root,path.normalize(p));if(!f.startsWith(root)){r.writeHead(403);return r.end()}
fs.readFile(f,(e,d)=>{if(e){r.writeHead(404);return r.end('not found')}r.writeHead(200,{'Content-Type':T[path.extname(f)]||'application/octet-stream'});r.end(d)})}).listen(port,()=>console.log('on '+port));
