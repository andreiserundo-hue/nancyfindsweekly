// Nancy Finds - live dashboard server. Local: node server.js -> http://localhost:4321
const http=require('http'), https=require('https'), fs=require('fs'), path=require('path');
let meta=null; try{meta=require('./meta_ads');}catch(e){} // direct Meta = failover source
const DIR=__dirname, FX=7.8, PORT=process.env.PORT||4321;
const {SHOP_TOKEN,SHOP,GLUED_KEY,WS}=require('./config');
// Optional password gate (set DASH_USER & DASH_PASS env vars on the host). Off locally if unset.
const AUTH=(process.env.DASH_USER&&process.env.DASH_PASS)?('Basic '+Buffer.from(process.env.DASH_USER+':'+process.env.DASH_PASS).toString('base64')):null;

const IMGDIR=DIR+'/creative_imgs'; if(!fs.existsSync(IMGDIR))fs.mkdirSync(IMGDIR);

// ---- shared task board (persisted to tasks.json) ----
const TASKS_FILE=DIR+'/tasks.json';
const DEFAULT_TASKS=[{id:'seed-duo',text:'Test the New Blossom Duo Landing Page https://home.nancyfinds.com/duo/',done:false,ts:Date.now()}];
function loadTasks(){try{return JSON.parse(fs.readFileSync(TASKS_FILE,'utf8'));}catch(e){return DEFAULT_TASKS.map(t=>({...t}));}}
function saveTasks(t){try{fs.writeFileSync(TASKS_FILE,JSON.stringify(t,null,2));}catch(e){}}
function readBody(req){return new Promise(r=>{let b='';req.on('data',c=>{b+=c;if(b.length>1e5)req.destroy();});req.on('end',()=>{try{r(JSON.parse(b||'{}'));}catch(e){r({});}});});}
// ---- shared "Situation Room" signals (synced context across all tabs) ----
const SIG_FILE=DIR+'/signals.json';
const DEFAULT_SIGNALS=[
 {id:'backlog',label:'Fulfillment backlog — paid spend throttled',status:'active',impact:'risk',note:'Spend is intentionally throttled while the shipment backlog clears (returns / chargeback risk). Resolve this when the backlog is cleared and all tabs update.',ts:Date.now(),updated:Date.now()},
 {id:'duo-lp',label:'Blossom Duo landing page — test pending',status:'active',impact:'test',taskId:'seed-duo',note:'Not tested yet. Linked to the "Test the New Blossom Duo Landing Page" board task — this auto-resolves when that checkbox is ticked. https://home.nancyfinds.com/duo/',ts:Date.now(),updated:Date.now()}
];
function loadSignals(){try{return JSON.parse(fs.readFileSync(SIG_FILE,'utf8'));}catch(e){return DEFAULT_SIGNALS.map(s=>({...s}));}}
function saveSignals(s){try{fs.writeFileSync(SIG_FILE,JSON.stringify(s,null,2));}catch(e){}}

// ---- suggestion overrides (edits / dismissals, persisted) ----
const SUG_FILE=DIR+'/suggestions.json';
function loadSug(){try{return JSON.parse(fs.readFileSync(SUG_FILE,'utf8'));}catch(e){return {};}}
function saveSug(o){try{fs.writeFileSync(SUG_FILE,JSON.stringify(o,null,2));}catch(e){}}

// ---- link preview (OpenGraph scrape, cached) ----
const previewCache={};
function decodeEnt(s){return (s||'').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#0?39;/g,"'").replace(/&#x27;/gi,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ');}
function metaTag(body,prop){const re=new RegExp('<meta[^>]+(?:property|name)=["\\\']'+prop+'["\\\'][^>]*>','i');const m=body.match(re);if(!m)return null;const c=m[0].match(/content=["\']([^"\']*)["\']/i);return c?decodeEnt(c[1]):null;}
function fetchUrl(url,depth){return new Promise((resolve,reject)=>{if((depth||0)>4)return reject(new Error('too many redirects'));const lib=url.startsWith('https')?https:http;const req=lib.get(url,{headers:{'User-Agent':'Mozilla/5.0 (NancyFinds link preview)'}},r=>{if(r.statusCode>=300&&r.statusCode<400&&r.headers.location){r.resume();let next;try{next=new URL(r.headers.location,url).toString();}catch(e){return reject(e);}return fetchUrl(next,(depth||0)+1).then(resolve,reject);}let b='',len=0;r.on('data',c=>{len+=c.length;if(len>6e5){req.destroy();}else b+=c;});r.on('end',()=>resolve({body:b,finalUrl:url}));});req.on('error',reject);req.setTimeout(8000,()=>req.destroy(new Error('timeout')));});}
async function linkPreview(url){if(previewCache[url]&&Date.now()-previewCache[url].t<6*3600000)return previewCache[url].d;const {body,finalUrl}=await fetchUrl(url);let image=metaTag(body,'og:image')||metaTag(body,'twitter:image');if(image){try{image=new URL(image,finalUrl).toString();}catch(e){}}let title=metaTag(body,'og:title')||(body.match(/<title[^>]*>([\s\S]*?)<\/title>/i)||[])[1]||'';let desc=metaTag(body,'og:description')||metaTag(body,'description')||'';let host='';try{host=new URL(finalUrl).host.replace(/^www\./,'');}catch(e){}const d={url:finalUrl,title:decodeEnt(title).trim().slice(0,120),desc:decodeEnt(desc).trim().slice(0,220),image:image||null,host};previewCache[url]={t:Date.now(),d};return d;}

// ---- editable scorecard benchmarks (persisted) ----
const BENCH_FILE=DIR+'/benchmarks.json';
function loadBench(){try{return JSON.parse(fs.readFileSync(BENCH_FILE,'utf8'));}catch(e){return null;}}
function saveBench(b){try{fs.writeFileSync(BENCH_FILE,JSON.stringify(b,null,2));}catch(e){}}

function isoHK(d){return new Date(d.getTime()+8*3600*1000).toISOString().slice(0,10);}
function addDay(s){const d=new Date(s+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+1);return d.toISOString().slice(0,10);} // Glued end-date is EXCLUSIVE (next-day 00:00)
function num(x){return parseFloat(x||'0');}
function nP(t){const s=(t||'').toLowerCase();if(s.includes('blossom')||s.includes('duo'))return'Blossom Duo';if(s.includes('stem'))return'Stem';if(s.includes('sprout'))return'Sprout';if(s.includes('bloom'))return'Bloom';if(s.includes('eclipse'))return'Eclipse';if(s.includes('vine'))return'Vine';if(s.includes('pebble'))return'Pebble Rose';if(s.includes('dew'))return'Dew';return t||'Unknown';}

// ---- Glued MCP ----
function gluedRpc(body){return new Promise((res,rej)=>{const data=JSON.stringify(body);const req=https.request({hostname:'app.glued.me',path:'/mcp',method:'POST',headers:{'Authorization':'Bearer '+GLUED_KEY,'Content-Type':'application/json','Accept':'application/json, text/event-stream','Content-Length':Buffer.byteLength(data)}},r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>{let j=null;const m=b.match(/data:\s*(\{[\s\S]*\})/);if(m){try{j=JSON.parse(m[1]);}catch(e){}}else{try{j=JSON.parse(b);}catch(e){}}res(j);});});req.on('error',rej);req.write(data);req.end();});}
let gluedReady=false;
async function gluedInit(){if(gluedReady)return;await gluedRpc({jsonrpc:"2.0",id:1,method:"initialize",params:{protocolVersion:"2024-11-05",capabilities:{},clientInfo:{name:"nancy",version:"1"}}});gluedReady=true;}
async function gluedCall(name,args){const r=await gluedRpc({jsonrpc:"2.0",id:Date.now(),method:"tools/call",params:{name,arguments:args}});if(r&&r.result&&r.result.content){const t=r.result.content.map(c=>c.text||'').join('\n');try{return JSON.parse(t);}catch(e){return t;}}return null;}
async function gReport(group,sd,ed){const j=await gluedCall('query_ad_report',{workspace_id:WS,start_date:sd,end_date:ed,date_range:"custom",group_by:group,metrics:["spend","roas","purchase_count"],sort_metric:"spend",sort_direction:"desc",limit:50});return (j&&j.rows)?j.rows.map(x=>({name:x.entity_name,spend:x.metrics.spend,roas:x.metrics.roas||0,buys:x.metrics.purchase_count||0,status:x.status})):[];}
function dlImg(url,dest,depth){return new Promise((res,rej)=>{if((depth||0)>4)return rej();https.get(url,{headers:{'User-Agent':'Mozilla/5.0'}},r=>{if(r.statusCode>=300&&r.statusCode<400&&r.headers.location){r.resume();return dlImg(r.headers.location,dest,(depth||0)+1).then(res,rej);}if(r.statusCode!==200){r.resume();return rej();}const f=fs.createWriteStream(dest);r.pipe(f);f.on('finish',()=>f.close(()=>res(dest)));}).on('error',rej);});}
async function gCreatives(sd,ed){const j=await gluedCall('query_ad_report',{workspace_id:WS,start_date:sd,end_date:ed,date_range:"custom",group_by:"creative",metrics:["spend","roas","purchase_count"],sort_metric:"spend",sort_direction:"desc",limit:24});if(!j||!j.rows)return [];const out=[];for(const r of j.rows){const name=(r.entity_name||'').split(' 2026-')[0];let img=null;const u=r.asset_url||r.thumbnail_url;if(u){const ext=(r.asset_url&&/\.(png|jpe?g)/i.test(r.asset_url.split('?')[0]))?r.asset_url.split('?')[0].match(/\.(png|jpe?g)/i)[1]:'jpg';const fn=(r.image_hash||name.replace(/[^a-z0-9]/gi,'_'))+'.'+ext;if(!fs.existsSync(IMGDIR+'/'+fn)){try{await dlImg(u,IMGDIR+'/'+fn);}catch(e){try{if(r.thumbnail_url)await dlImg(r.thumbnail_url,IMGDIR+'/'+fn);}catch(e2){}}}if(fs.existsSync(IMGDIR+'/'+fn))img='creative_imgs/'+fn;}
  const bd=(r.ad_breakdown&&r.ad_breakdown[0])||{};
  out.push({name,roas:r.metrics.roas||0,spend:r.metrics.spend,buys:r.metrics.buys||r.metrics.purchase_count||0,img,status:r.status,type:r.creative_type,campaign:(bd.campaign_name||'').replace(/Nancy-Testing-2_/,''),adset:(bd.ad_set_name||'').replace(/Nancy-Testing-2_/,''),lp:bd.landing_page_url||'',cpa:bd.metrics&&bd.metrics.cpa});}return out;}

// ---- Shopify ----
function shopAll(url){return new Promise((res,rej)=>{let all=[];(function go(u){const p=new URL(u);https.request({hostname:p.hostname,path:p.pathname+p.search,headers:{'X-Shopify-Access-Token':SHOP_TOKEN}},r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>{const d=JSON.parse(b);if(d.orders)all=all.concat(d.orders);const l=r.headers['link'];if(l&&l.includes('rel="next"')){const n=l.match(/<([^>]+)>;\s*rel="next"/);if(n)go(n[1]);else res(all);}else res(all);});}).on('error',rej).end();})(url);});}
let refundCache=null;
async function refunds(){if(refundCache)return refundCache;const ords=await shopAll(`https://${SHOP}/admin/api/2024-01/orders.json?status=any&limit=250&created_at_min=2026-03-01T00:00:00%2B08:00&fields=id,refunds`);const by={};ords.forEach(o=>(o.refunds||[]).forEach(r=>{const amt=(r.transactions||[]).filter(t=>t.kind==='refund'&&t.status==='success').reduce((a,t)=>a+num(t.amount),0);if(amt>0){const d=isoHK(new Date(r.created_at));by[d]=(by[d]||0)+amt;}}));refundCache=by;setTimeout(()=>refundCache=null,3600000);return by;}
function hkHM(iso){return new Date(new Date(iso).getTime()+8*3600*1000).toISOString().slice(11,16);}
async function shopRange(sd,ed,st,et){let orders=await shopAll(`https://${SHOP}/admin/api/2024-01/orders.json?status=any&limit=250&created_at_min=${sd}T00:00:00%2B08:00&created_at_max=${ed}T23:59:59%2B08:00&fields=id,created_at,total_price,current_total_price,line_items,shipping_address,billing_address`);
  const timed = st&&et;
  if(timed) orders=orders.filter(o=>{const hm=hkHM(o.created_at);return hm>=st&&hm<et;});
  const rb=await refunds();const ret=timed?0:Object.entries(rb).filter(([d])=>d>=sd&&d<=ed).reduce((a,[,v])=>a+v,0);const gross=orders.reduce((a,o)=>a+num(o.total_price),0);const prod={},country={},daily={};orders.forEach(o=>{(o.line_items||[]).forEach(li=>{const n=nP(li.title);if(!prod[n])prod[n]={units:0,rev:0};prod[n].units+=(li.quantity||1);prod[n].rev+=num(li.price)*(li.quantity||1);});const c=(o.shipping_address&&o.shipping_address.country)||(o.billing_address&&o.billing_address.country)||'Unknown';if(!country[c])country[c]={orders:0,net:0};country[c].orders++;country[c].net+=num(o.current_total_price);const dd=isoHK(new Date(o.created_at));if(!daily[dd])daily[dd]={orders:0,net:0};daily[dd].orders++;daily[dd].net+=num(o.current_total_price);});return {orders:orders.length,gross,returns:ret,net:gross-ret,retPct:gross>0?ret/gross*100:0,aov:orders.length?gross/orders.length:0,products:prod,countries:country,daily};}

// ---- campaign hourly (time-windowed) ----
async function gCampHourly(sd,edEx,st,et){const j=await gluedCall('query_campaign_hourly',{workspace_id:WS,date_range:'custom',start_date:sd,end_date:edEx,start_time:st,end_time:et,metrics:['spend','revenue','roas','purchase_count']});if(!j||!j.rows)return [];return j.rows.map(r=>({name:r.campaign_name,spend:(r.metrics&&r.metrics.spend)||0,roas:(r.metrics&&r.metrics.roas)||0,buys:(r.metrics&&r.metrics.purchase_count)||0,status:(r.metadata&&r.metadata.status)||''}));}

// ---- range bundle (cached) ----
const cache={};
const safe=p=>p.then(x=>x).catch(()=>null);
async function getRange(sd,ed,st,et){const timed=!!(st&&et);const k=sd+'_'+ed+'_'+(st||'')+'_'+(et||'');if(cache[k]&&Date.now()-cache[k].t<1800000)return cache[k].d;
  let glueUp=true; try{await gluedInit();}catch(e){glueUp=false;}
  const edEx=addDay(ed); /* Glued end exclusive */ let source='glued';
  // campaigns â€” Glued primary, Meta direct failover (day-level)
  let campaigns = glueUp ? await safe(timed?gCampHourly(sd,edEx,st,et):gReport('campaign',sd,edEx)) : null;
  if(campaigns===null){ if(meta){ try{ campaigns=await meta.insights('campaign',sd,ed); source='meta (failover)'; }catch(e){ campaigns=[]; source='unavailable'; } } else { campaigns=[]; source='unavailable'; } }
  const [landing,creatives,shop]=await Promise.all([ glueUp?safe(gReport('landing_page',sd,edEx)).then(x=>x||[]):Promise.resolve([]), glueUp?safe(gCreatives(sd,edEx)).then(x=>x||[]):Promise.resolve([]), shopRange(sd,ed,st,et) ]);
  const spendHKD=campaigns.reduce((a,c)=>a+c.spend,0);const metaRevHKD=campaigns.reduce((a,c)=>a+c.spend*c.roas,0);const d={start:sd,end:ed,st:st||null,et:et||null,timed,source,fx:FX,campaigns,landing,creatives,shop,meta:{spendHKD,spendUSD:spendHKD/FX,metaRoas:spendHKD>0?metaRevHKD/spendHKD:0,buys:campaigns.reduce((a,c)=>a+c.buys,0)},trueRoas:(spendHKD/FX)>0?shop.net/(spendHKD/FX):0};cache[k]={t:Date.now(),d};return d;}

// ---- server ----
const MIME={'.html':'text/html','.js':'text/javascript','.css':'text/css','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.json':'application/json'};
http.createServer(async (req,res)=>{
  const u=new URL(req.url,'http://x');
  if(AUTH && (req.headers.authorization||'')!==AUTH){res.writeHead(401,{'WWW-Authenticate':'Basic realm="Nancy Finds Weekly"'});return res.end('Authentication required');}
  try{
    if(u.pathname==='/'||u.pathname==='/index.html'){const h=fs.readFileSync(DIR+'/index.html');res.writeHead(200,{'Content-Type':'text/html'});return res.end(h);}
    if(u.pathname.startsWith('/creative_imgs/')){const fp=DIR+decodeURIComponent(u.pathname);if(fs.existsSync(fp)){res.writeHead(200,{'Content-Type':MIME[path.extname(fp)]||'application/octet-stream'});return res.end(fs.readFileSync(fp));}res.writeHead(404);return res.end();}
    if(u.pathname==='/api/range'){const sd=u.searchParams.get('start'),ed=u.searchParams.get('end'),st=u.searchParams.get('st')||'',et=u.searchParams.get('et')||'';const d=await getRange(sd,ed,st,et);res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(d));}
    if(u.pathname==='/api/hourly'){const date=u.searchParams.get('date');await gluedInit();const j=await gluedCall('query_campaign_hourly',{workspace_id:WS,date_range:'custom',start_date:date,end_date:addDay(date),metrics:['spend','revenue','roas','purchase_count']});res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(j||{rows:[]}));}
    if(u.pathname==='/api/pptx'){const now=new Date();const yd=isoHK(new Date(now-86400000)),l7=isoHK(new Date(now-7*86400000)),l30=isoHK(new Date(now-30*86400000));const [r30,r7,ry]=await Promise.all([getRange(l30,yd),getRange(l7,yd),getRange(yd,yd)]);const out=DIR+'/Nancy Finds - Weekly (live).pptx';await require('./build_live_deck')(r30,r7,ry,out);const buf=fs.readFileSync(out);res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation','Content-Disposition':'attachment; filename="Nancy Finds - Weekly.pptx"'});return res.end(buf);}
    if(u.pathname==='/api/tasks'){
      if(req.method==='GET'){res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(loadTasks()));}
      if(req.method==='POST'){const bd=await readBody(req);const text=(bd.text||'').toString().trim().slice(0,200);if(!text){res.writeHead(400,{'Content-Type':'application/json'});return res.end('{"error":"empty"}');}const task={id:'t'+Date.now().toString(36)+Math.floor(Math.random()*1e4).toString(36),text,done:false,ts:Date.now()};const t=loadTasks();t.unshift(task);saveTasks(t);res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(task));}
      if(req.method==='PATCH'){const bd=await readBody(req);const id=u.searchParams.get('id');const t=loadTasks().map(x=>x.id===id?{...x,done:!!bd.done}:x);saveTasks(t);res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
      if(req.method==='DELETE'){const id=u.searchParams.get('id');saveTasks(loadTasks().filter(x=>x.id!==id));res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
    }
    if(u.pathname==='/api/signals'){
      if(req.method==='GET'){res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(loadSignals()));}
      if(req.method==='POST'){const bd=await readBody(req);let sigs=loadSignals();
        if(bd.id&&sigs.find(s=>s.id===bd.id)){sigs=sigs.map(s=>{if(s.id!==bd.id)return s;const n={...s};if(bd.label!==undefined)n.label=(''+bd.label).slice(0,160);if(bd.note!==undefined)n.note=(''+bd.note).slice(0,400);if(bd.impact!==undefined)n.impact=bd.impact;if(bd.status!==undefined)n.status=bd.status;n.updated=Date.now();return n;});}
        else{const id=bd.id||('s'+Date.now().toString(36)+Math.floor(Math.random()*1e3).toString(36));sigs.unshift({id,label:(''+(bd.label||'')).slice(0,160),note:(''+(bd.note||'')).slice(0,400),impact:bd.impact||'info',status:bd.status||'active',ts:Date.now(),updated:Date.now()});}
        saveSignals(sigs);res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
      if(req.method==='DELETE'){const id=u.searchParams.get('id');saveSignals(loadSignals().filter(s=>s.id!==id));res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
    }
    if(u.pathname==='/api/suggestions'){
      if(req.method==='GET'){res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(loadSug()));}
      if(req.method==='POST'){const bd=await readBody(req);const key=(bd.key||'').toString();if(key){const o=loadSug();const cur=o[key]||{};if(bd.text!==undefined)cur.text=(bd.text||'').toString().slice(0,400);if(bd.hidden!==undefined)cur.hidden=!!bd.hidden;o[key]=cur;saveSug(o);}res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
      if(req.method==='DELETE'){const key=u.searchParams.get('key');if(key){const o=loadSug();delete o[key];saveSug(o);}else saveSug({});res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
    }
    if(u.pathname==='/api/preview'){const url=u.searchParams.get('url')||'';if(!/^https?:\/\//i.test(url)){res.writeHead(400,{'Content-Type':'application/json'});return res.end('{"error":"bad url"}');}try{const d=await linkPreview(url);res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(d));}catch(e){res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify({url,error:String(e&&e.message||e)}));}}
    if(u.pathname==='/api/benchmarks'){
      if(req.method==='GET'){res.writeHead(200,{'Content-Type':'application/json'});return res.end(JSON.stringify(loadBench()||{}));}
      if(req.method==='PUT'){const bd=await readBody(req);saveBench(bd);res.writeHead(200,{'Content-Type':'application/json'});return res.end('{"ok":true}');}
    }
    res.writeHead(404);res.end('not found');
  }catch(e){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({error:String(e&&e.message||e)}));}
}).listen(PORT,()=>console.log('Nancy Finds dashboard â†’ http://localhost:'+PORT));
