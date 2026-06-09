// Nancy Finds Ã¢â‚¬â€ Weekly auto-refresh: pulls Glued (Meta) + Shopify, writes a tabbed dashboard.html
const https = require('https');
const fs = require('fs');
const DIR = "C:/Users/andre/Nancy Finds Reporting";
const FX = 7.8;

const {SHOP_TOKEN,SHOP,GLUED_KEY,WS}=require('./config');




function isoHK(d){return new Date(d.getTime()+8*3600*1000).toISOString().slice(0,10);}
const now = new Date();
const today = isoHK(now);
const yest = isoHK(new Date(now.getTime()-1*86400000));
const l7s = isoHK(new Date(now.getTime()-7*86400000));
const l30s = isoHK(new Date(now.getTime()-30*86400000));
const RANGES = { L30D:[l30s,yest], L7D:[l7s,yest], Yesterday:[yest,yest] };

function gluedRpc(body){return new Promise((res,rej)=>{const data=JSON.stringify(body);const req=https.request({hostname:'app.glued.me',path:'/mcp',method:'POST',headers:{'Authorization':'Bearer '+GLUED_KEY,'Content-Type':'application/json','Accept':'application/json, text/event-stream','Content-Length':Buffer.byteLength(data)}},r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>{let j=null;const m=b.match(/data:\s*(\{[\s\S]*\})/);if(m){try{j=JSON.parse(m[1]);}catch(e){}}else{try{j=JSON.parse(b);}catch(e){}}res(j);});});req.on('error',rej);req.write(data);req.end();});}
async function gluedCall(name,args){const r=await gluedRpc({jsonrpc:"2.0",id:Date.now(),method:"tools/call",params:{name,arguments:args}});if(r&&r.result&&r.result.content){const t=r.result.content.map(c=>c.text||'').join('\n');try{return JSON.parse(t);}catch(e){return t;}}return null;}
async function gluedReport(group,sd,ed){const j=await gluedCall('query_ad_report',{workspace_id:WS,start_date:sd,end_date:ed,date_range:"custom",group_by:group,metrics:["spend","roas","purchase_count"],sort_metric:"spend",sort_direction:"desc",limit:50});return (j&&j.rows)?j.rows.map(x=>({name:x.entity_name,spend:x.metrics.spend,roas:x.metrics.roas||0,buys:x.metrics.purchase_count||0,status:x.status})):[];}

function shopAll(url){return new Promise((res,rej)=>{let all=[];(function go(u){const p=new URL(u);https.request({hostname:p.hostname,path:p.pathname+p.search,headers:{'X-Shopify-Access-Token':SHOP_TOKEN}},r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>{const d=JSON.parse(b);if(d.orders)all=all.concat(d.orders);const l=r.headers['link'];if(l&&l.includes('rel="next"')){const n=l.match(/<([^>]+)>;\s*rel="next"/);if(n)go(n[1]);else res(all);}else res(all);});}).on('error',rej).end();})(url);});}
function num(x){return parseFloat(x||'0');}
function nP(t){const s=(t||'').toLowerCase();if(s.includes('blossom')||s.includes('duo'))return'Blossom Duo';if(s.includes('stem'))return'Stem';if(s.includes('sprout'))return'Sprout';if(s.includes('bloom'))return'Bloom';if(s.includes('eclipse'))return'Eclipse';if(s.includes('vine'))return'Vine';if(s.includes('pebble'))return'Pebble Rose';if(s.includes('dew'))return'Dew';return t||'Unknown';}
function downloadImg(url,dest,depth){return new Promise((res,rej)=>{if((depth||0)>4)return rej(new Error('redirects'));https.get(url,{headers:{'User-Agent':'Mozilla/5.0'}},r=>{if(r.statusCode>=300&&r.statusCode<400&&r.headers.location){r.resume();return downloadImg(r.headers.location,dest,(depth||0)+1).then(res,rej);}if(r.statusCode!==200){r.resume();return rej(new Error('status '+r.statusCode));}const f=fs.createWriteStream(dest);r.pipe(f);f.on('finish',()=>f.close(()=>res(dest)));}).on('error',rej);});}
async function gluedCreatives(sd,ed){
  const j=await gluedCall('query_ad_report',{workspace_id:WS,start_date:sd,end_date:ed,date_range:"custom",group_by:"creative",metrics:["spend","roas","purchase_count"],sort_metric:"spend",sort_direction:"desc",limit:16});
  if(!j||!j.rows)return [];
  if(!fs.existsSync(DIR+'/creative_imgs'))fs.mkdirSync(DIR+'/creative_imgs');
  const out=[];
  for(const r of j.rows){
    const name=(r.entity_name||'').split(' 2026-')[0];
    let img=null; const url=r.asset_url||r.thumbnail_url;
    if(url){const ext=(r.asset_url&&/\.(png|jpe?g)/i.test(r.asset_url.split('?')[0]))?r.asset_url.split('?')[0].match(/\.(png|jpe?g)/i)[1]:'jpg';
      const fn='creative_imgs/'+(r.image_hash||name.replace(/[^a-z0-9]/gi,'_'))+'.'+ext;
      try{await downloadImg(url,DIR+'/'+fn);img=fn;}catch(e){try{if(r.thumbnail_url){await downloadImg(r.thumbnail_url,DIR+'/'+fn);img=fn;}}catch(e2){}}}
    out.push({name,roas:r.metrics.roas||0,spend:r.metrics.spend,buys:r.metrics.purchase_count||0,img,status:r.status});
  }
  return out;
}

(async()=>{
  console.log('Refreshing... ranges:', JSON.stringify(RANGES));
  await gluedRpc({jsonrpc:"2.0",id:1,method:"initialize",params:{protocolVersion:"2024-11-05",capabilities:{},clientInfo:{name:"nancy-dash",version:"1"}}});
  const meta = {};
  for(const [lbl,[sd,ed]] of Object.entries(RANGES)){
    meta[lbl] = { campaigns: await gluedReport('campaign',sd,ed), landing: await gluedReport('landing_page',sd,ed), ads: await gluedReport('ad',sd,ed) };
  }
  const orders = await shopAll(`https://${SHOP}/admin/api/2024-01/orders.json?status=any&limit=250&created_at_min=${l30s}T00:00:00%2B08:00&created_at_max=${today}T23:59:59%2B08:00&fields=id,created_at,total_price,current_total_price,line_items,shipping_address,billing_address`);
  const ords2 = await shopAll(`https://${SHOP}/admin/api/2024-01/orders.json?status=any&limit=250&created_at_min=2026-03-01T00:00:00%2B08:00&fields=id,refunds`);
  const refByDate={}; ords2.forEach(o=>(o.refunds||[]).forEach(r=>{const amt=(r.transactions||[]).filter(t=>t.kind==='refund'&&t.status==='success').reduce((a,t)=>a+num(t.amount),0);if(amt>0){const d=isoHK(new Date(r.created_at));refByDate[d]=(refByDate[d]||0)+amt;}}));
  const refSum=(s,e)=>Object.entries(refByDate).filter(([d])=>d>=s&&d<=e).reduce((a,[,v])=>a+v,0);
  function shopPeriod(sd,ed){const inP=orders.filter(o=>{const d=isoHK(new Date(o.created_at));return d>=sd&&d<=ed;});const gross=inP.reduce((a,o)=>a+num(o.total_price),0);const ret=refSum(sd,ed);return {orders:inP.length,gross,returns:ret,net:gross-ret,retPct:gross>0?ret/gross*100:0};}
  const shop={}; for(const [lbl,[sd,ed]] of Object.entries(RANGES)) shop[lbl]=shopPeriod(sd,ed);
  const prod={},country={},daily={};
  orders.forEach(o=>{(o.line_items||[]).forEach(li=>{const n=nP(li.title);if(!prod[n])prod[n]={units:0,rev:0};prod[n].units+=(li.quantity||1);prod[n].rev+=num(li.price)*(li.quantity||1);});const c=(o.shipping_address&&o.shipping_address.country)||(o.billing_address&&o.billing_address.country)||'Unknown';if(!country[c])country[c]={orders:0,net:0};country[c].orders++;country[c].net+=num(o.current_total_price);const dd=isoHK(new Date(o.created_at));if(!daily[dd])daily[dd]={orders:0,net:0};daily[dd].orders++;daily[dd].net+=num(o.current_total_price);});
  console.log('Pulling creatives + images...');
  const creatives={ L30D: await gluedCreatives(RANGES.L30D[0],RANGES.L30D[1]), L7D: await gluedCreatives(RANGES.L7D[0],RANGES.L7D[1]) };
  const data={generated:new Date().toISOString(),ranges:RANGES,fx:FX,meta,shop,products:prod,countries:country,daily,creatives};
  fs.writeFileSync(DIR+'/dashboard_data.json',JSON.stringify(data,null,2));
  buildHtml(data);
  console.log('Done -> dashboard.html');
})().catch(e=>console.log('ERR',e.message));

function buildHtml(d){
  const DATA = JSON.stringify(d);
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Nancy Finds Ã¢â‚¬â€ Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Calibri,'Segoe UI',sans-serif;background:#F4EEF1;color:#1A1430;display:flex;min-height:100vh}
/* sidebar */
.nav{width:230px;background:#1A1430;color:#fff;padding:22px 0;flex-shrink:0;position:sticky;top:0;height:100vh}
.brand{font-family:Georgia,serif;font-size:20px;font-weight:700;padding:0 22px 4px;color:#fff}
.brand small{display:block;font-family:Calibri;font-size:10px;letter-spacing:2px;color:#C44B6E;font-weight:700;margin-top:2px}
.navitem{padding:11px 22px;font-size:14px;color:#BCB5C2;cursor:pointer;border-left:3px solid transparent;display:flex;align-items:center;gap:10px}
.navitem:hover{background:#2A2046;color:#fff}
.navitem.active{background:#2A2046;color:#fff;border-left-color:#C44B6E;font-weight:700}
.navitem .ic{width:18px;text-align:center}
.gen{position:absolute;bottom:16px;left:22px;right:22px;font-size:10px;color:#6b6480;font-style:italic}
/* main */
.main{flex:1;padding:28px 34px;overflow-y:auto}
.page{display:none}.page.show{display:block}
h1{font-family:Georgia,serif;font-size:28px;margin-bottom:2px}
.cap{color:#8A8392;font-size:12px;margin-bottom:18px}
.toggle{display:inline-flex;background:#fff;border:1px solid #E7DFE4;border-radius:10px;overflow:hidden;margin-bottom:16px}
.toggle button{border:0;background:#fff;padding:8px 16px;font-size:13px;cursor:pointer;color:#8A8392;font-weight:600}
.toggle button.on{background:#C44B6E;color:#fff}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
.kpi{background:#1A1430;border-radius:12px;padding:16px 18px}
.kpi .v{font-family:Georgia,serif;font-size:28px;color:#fff;font-weight:700}
.kpi .l{color:#C44B6E;font-size:11px;text-transform:uppercase;letter-spacing:1px}
.kpi .n{color:#BCB5C2;font-size:11px;font-style:italic;margin-top:2px}
.card{background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 1px 6px rgba(0,0,0,.06);margin-bottom:18px}
h2{font-family:Georgia,serif;font-size:18px;color:#8E2F4C;margin-bottom:2px}
.sub2{color:#8A8392;font-size:11px;margin-bottom:12px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:right;color:#8A8392;font-size:10px;text-transform:uppercase;letter-spacing:.4px;padding:7px 9px;border-bottom:2px solid #F4EEF1}
td{text-align:right;padding:8px 9px;border-bottom:1px solid #F4EEF1}
th:first-child,td.lh{text-align:left}
tr:nth-child(even) td{background:#FBF7F9}
.peach td{background:#FCE4D6 !important}
.bar{height:14px;background:#C44B6E;border-radius:3px;display:inline-block;vertical-align:middle}
.barbg{background:#F4EEF1;border-radius:3px;flex:1}
.pill{display:inline-block;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:700;color:#fff}
.dot{font-size:18px;vertical-align:middle}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.note{background:#F7E9EF;border-radius:10px;padding:12px 14px;font-size:12.5px;color:#1A1430;margin-top:6px;display:flex;align-items:center;gap:10px}
.note img{width:38px;height:38px;object-fit:cover;border-radius:6px;flex-shrink:0}
.cgrid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
.ccard{background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.08)}
.cbody{padding:11px 13px}
.cname{font-size:12px;font-weight:700;color:#1A1430;min-height:30px;line-height:1.25}
.crow{display:flex;justify-content:space-between;align-items:center;margin:6px 0 4px}
.croas{font-family:Georgia,serif;font-size:20px;font-weight:700}
.cmeta{font-size:11px;color:#8A8392}
@media(max-width:1100px){.cgrid{grid-template-columns:repeat(2,1fr)}}
</style></head><body>
<div class="nav">
  <div class="brand">Nancy Finds<small>WEEKLY DASHBOARD</small></div>
  <div style="margin-top:18px">
    <div class="navitem active" data-p="overview"><span class="ic">Ã¢â€”Ë†</span>Overview</div>
    <div class="navitem" data-p="scorecard"><span class="ic">Ã¢â€“Â¦</span>Scorecard</div>
    <div class="navitem" data-p="campaigns"><span class="ic">Ã°Å¸â€œÂ£</span>Campaigns</div>
    <div class="navitem" data-p="landing"><span class="ic">Ã°Å¸â€â€”</span>Landing Pages</div>
    <div class="navitem" data-p="creatives"><span class="ic">Ã°Å¸Å½Â¨</span>Creatives</div>
    <div class="navitem" data-p="products"><span class="ic">Ã°Å¸Å’Â¹</span>Products</div>
    <div class="navitem" data-p="countries"><span class="ic">Ã°Å¸Å’Å½</span>Countries</div>
  </div>
  <div class="gen" id="gen"></div>
</div>
<div class="main" id="main"></div>
<script>
const D=${DATA};
const FX=D.fx;
const usd=n=>'$'+Math.round(n).toLocaleString();
const usdH=n=>'$'+Math.round(n/FX).toLocaleString();
const rc=v=>v>=1.75?'#1F9D6B':(v>=1.2?'#E0922F':'#DB3B3B');
const dot=(v,g,b)=>'<span class="dot" style="color:'+(v>=b?'#1F9D6B':(v>=g?'#E0922F':'#DB3B3B'))+'">Ã¢â€”Â</span>';
const tier=v=>v>=2.1?['SCALE','#1F9D6B']:(v>=1.6?['HOLD','#E0922F']:['CUT','#DB3B3B']);
const clean=n=>(n||'').replace(/Nancy-Testing-2_/,'').replace(/https?:\\/\\/(www\\.)?/,'');
const DAYS={L30D:30,L7D:7,Yesterday:1};
function sc(lbl){const m=D.meta[lbl],s=D.shop[lbl],days=DAYS[lbl];const spH=m.campaigns.reduce((a,c)=>a+c.spend,0);const spU=spH/FX;const rev=m.campaigns.reduce((a,c)=>a+c.spend*c.roas,0);return{rev:s.net/days,spend:spU/days,orders:s.orders/days,retPct:s.retPct,metaRoas:spH>0?rev/spH:0,trueRoas:spU>0?s.net/spU:0,spendPct:s.net>0?spU/s.net*100:0};}
const P={L30D:sc('L30D'),L7D:sc('L7D'),Yesterday:sc('Yesterday')};
function gen(){return new Date(D.generated).toLocaleString();}
document.getElementById('gen').innerHTML='Generated '+gen()+'<br>L30D '+D.ranges.L30D[0]+'Ã¢â€ â€™'+D.ranges.L30D[1];

// ---------- pages ----------
function overview(){
  const k=P.L7D;
  return '<h1>Overview</h1><div class="cap">Last 7 days snapshot Ã‚Â· auto-pulled from Glued + Shopify</div>'+
  '<div class="kpis">'+
  kpi('Rev/day (net)',usd(k.rev),'L7D Ã‚Â· target $1,000')+
  kpi('True ROAS',k.trueRoas.toFixed(2)+'x',(k.trueRoas>=1.5?'on target':'below 1.5 floor'))+
  kpi('Meta ROAS',k.metaRoas.toFixed(2)+'x','Glued attributed')+
  kpi('Return %',k.retPct.toFixed(0)+'%','target <3%')+
  '</div>'+
  '<div class="grid2"><div class="card"><h2>Daily Net Revenue (L30D)</h2><div class="sub2">Shopify, USD</div>'+dailyBars()+'</div>'+
  '<div class="card"><h2>Green Shoots vs Watch-outs</h2><div class="sub2">where the forward play is</div>'+greenShoots()+'</div></div>';
}
function dailyBars(){const days=Object.entries(D.daily).sort((a,b)=>a[0]<b[0]?-1:1);const max=Math.max(...days.map(x=>x[1].net),1);return '<div style="display:flex;align-items:flex-end;gap:3px;height:160px;margin-top:8px">'+days.map(([d,v])=>'<div title="'+d+': $'+Math.round(v.net)+'" style="flex:1;background:#C44B6E;border-radius:2px 2px 0 0;height:'+Math.max(2,v.net/max*150)+'px"></div>').join('')+'</div><div class="sub2" style="margin-top:6px">'+days[0][0].slice(5)+' Ã¢â€ â€™ '+days[days.length-1][0].slice(5)+'</div>';}
function greenShoots(){const lp7=D.meta.L7D.landing.filter(x=>x.spend>0).sort((a,b)=>b.roas-a.roas);const top=lp7[0];
  const cre=((D.creatives&&D.creatives.L7D)||[]).filter(x=>x.roas>1.5).sort((a,b)=>b.roas-a.roas).slice(0,2);
  let h='<div class="note" style="background:#E4F0EC">Ã°Å¸Å¸Â¢ <b>Best LP (L7D):</b>&nbsp; '+(top?clean(top.name).substring(0,38)+' Ã¢â‚¬â€ '+top.roas.toFixed(2)+'x':'Ã¢â‚¬â€')+'</div>';
  cre.forEach(a=>h+='<div class="note" style="background:#E4F0EC">'+(a.img?'<img src="'+a.img+'">':'')+'<span>Ã°Å¸Å¸Â¢ <b>Winning creative:</b> '+clean(a.name).substring(0,30)+' Ã¢â‚¬â€ '+a.roas.toFixed(0)+'x</span></div>');
  h+='<div class="note">Ã°Å¸â€Â´ <b>Watch:</b> all campaigns below 1.5x L7D; spend throttled for the fulfillment backlog (returns '+P.L7D.retPct.toFixed(0)+'%).</div>';
  return h;}
function kpi(l,v,n){return '<div class="kpi"><div class="l">'+l+'</div><div class="v">'+v+'</div><div class="n">'+n+'</div></div>';}

function scorecard(){
  const row=(label,fmt,key,circ,g,b,peach)=>'<tr class="'+(peach?'peach':'')+'"><td class="lh">'+label+'</td>'+['L30D','L7D','Yesterday'].map(k=>{const v=P[k][key];return '<td>'+fmt(v)+(circ?' '+dot(v,g,b):'')+'</td>';}).join('')+'<td style="color:#8A8392">'+ (key==='rev'?'$1,000/day':key==='metaRoas'||key==='trueRoas'?'1.5Ã¢â‚¬â€œ2.0':key==='retPct'?'<3%':'Ã¢â‚¬â€')+'</td></tr>';
  return '<h1>Overall Scorecard</h1><div class="cap">Daily averages, USD Ã‚Â· Ã¢â€”Â green Ã¢â€°Â¥ Better Ã‚Â· amber Ã¢â€°Â¥ Good Ã‚Â· red below target</div><div class="card"><table>'+
  '<tr><th>Metric</th><th>L30D avg</th><th>L7D avg</th><th>Yesterday</th><th>Target</th></tr>'+
  row('Revenue / day',usd,'rev',true,1000,1500,true)+
  row('Meta Spend / day',usd,'spend',false)+
  row('Spend % of rev',v=>v.toFixed(0)+'%','spendPct',false)+
  row('Orders / day',v=>v.toFixed(1),'orders',false)+
  row('Return %',v=>v.toFixed(1)+'%','retPct',false)+
  row('Meta ROAS',v=>v.toFixed(2),'metaRoas',true,1.5,1.75,true)+
  row('True ROAS (net)',v=>v.toFixed(2),'trueRoas',true,1.5,1.75,true)+
  '</table></div>';
}
let cR='L7D',lR='L7D',crR='L7D';
function rangeToggle(cur,fn){return '<div class="toggle">'+['L30D','L7D','Yesterday'].map(r=>'<button class="'+(r===cur?'on':'')+'" onclick="'+fn+'(\\''+r+'\\')">'+r+'</button>').join('')+'</div>';}
function campaigns(){const rows=D.meta[cR].campaigns.filter(c=>c.spend>0).map(c=>{const t=tier(c.roas);return '<tr><td class="lh">'+clean(c.name).substring(0,52)+'</td><td>'+usdH(c.spend)+'</td><td style="color:'+rc(c.roas)+';font-weight:700">'+c.roas.toFixed(2)+'</td><td>'+c.buys+'</td><td><span class="pill" style="background:'+t[1]+'">'+t[0]+'</span></td></tr>';}).join('');
  return '<h1>Campaign Analysis</h1><div class="cap">Tiered by ROAS Ã‚Â· SCALE Ã¢â€°Â¥2.1 Ã‚Â· HOLD 1.6Ã¢â‚¬â€œ2.0 Ã‚Â· CUT Ã¢â€°Â¤1.5</div>'+rangeToggle(cR,'setC')+'<div class="card"><table><tr><th>Campaign</th><th>Spend</th><th>ROAS</th><th>Buys</th><th>Tier</th></tr>'+(rows||'<tr><td class="lh">No spend in range</td></tr>')+'</table></div>';}
function landing(){const all=D.meta[lR].landing.filter(c=>c.spend>0);const aura=all.filter(x=>/aura/i.test(x.name));const bridge=all.filter(x=>!/aura/i.test(x.name));
  const tbl=(arr)=>'<table><tr><th>Page</th><th>Spend</th><th>ROAS</th><th>Buys</th></tr>'+(arr.map(c=>'<tr><td class="lh">'+clean(c.name).substring(0,46)+'</td><td>'+usdH(c.spend)+'</td><td style="color:'+rc(c.roas)+';font-weight:700">'+c.roas.toFixed(2)+'</td><td>'+c.buys+'</td></tr>').join('')||'<tr><td class="lh">none</td></tr>')+'</table>';
  return '<h1>Landing Pages</h1><div class="cap">Bridge pages (PDP) + Aura advertorials</div>'+rangeToggle(lR,'setL')+'<div class="grid2"><div class="card"><h2>Website / Bridge Pages</h2><div class="sub2">per-product PDP bridges</div>'+tbl(bridge)+'</div><div class="card"><h2>Aura Advertorials</h2><div class="sub2">listicle LPs</div>'+tbl(aura)+'</div></div>';}
function creatives(){const list=(D.creatives&&D.creatives[crR])?D.creatives[crR]:[];
  const cards=list.map(c=>{const col=rc(c.roas);const t=tier(c.roas);
    const im=c.img?'<img src="'+c.img+'" style="width:100%;height:190px;object-fit:cover;display:block">':'<div style="width:100%;height:190px;background:#2A2046;display:flex;align-items:center;justify-content:center;color:#8A8392;font-size:11px">no image</div>';
    return '<div class="ccard">'+im+'<div class="cbody"><div class="cname">'+clean(c.name).substring(0,40)+'</div>'+
      '<div class="crow"><span class="croas" style="color:'+col+'">'+c.roas.toFixed(2)+'x</span><span class="pill" style="background:'+t[1]+'">'+t[0]+'</span></div>'+
      '<div class="cmeta">'+usdH(c.spend)+' spend Ã‚Â· '+c.buys+' buys</div></div></div>';}).join('');
  return '<h1>Creatives Ã¢â‚¬â€ Ads Analysis</h1><div class="cap">Top ads by spend Ã‚Â· live images pulled from Meta via Glued</div>'+rangeToggle(crR,'setCr')+'<div class="cgrid">'+(cards||'<div class="cap">No creatives in range</div>')+'</div>';}
function products(){const sale=['Blossom Duo','Vine','Eclipse','Bloom','Stem','Sprout'];const tot=Object.values(D.products).reduce((a,p)=>a+p.units,0);const max=Math.max(...sale.map(n=>(D.products[n]||{units:0}).units),1);
  const rows=sale.map(n=>{const p=D.products[n]||{units:0,rev:0};return '<tr><td class="lh">'+n+'</td><td><div style="display:flex;align-items:center;gap:8px"><div class="barbg"><div class="bar" style="width:'+(p.units/max*100)+'%"></div></div><b>'+p.units+'</b></div></td><td>'+usd(p.rev)+'</td><td>'+Math.round(p.units/tot*100)+'%</td></tr>';}).join('');
  return '<h1>Orders by Product</h1><div class="cap">Saleable suction line (L30D)</div><div class="card"><table><tr><th>Product</th><th>Units</th><th>Revenue</th><th>% Units</th></tr>'+rows+'</table></div>';}
function countries(){const co=Object.entries(D.countries).sort((a,b)=>b[1].net-a[1].net).slice(0,8);const tot=Object.values(D.countries).reduce((a,c)=>a+c.net,0);const max=co[0]?co[0][1].net:1;
  const rows=co.map(([c,v])=>'<tr><td class="lh">'+c+'</td><td>'+v.orders+'</td><td><div style="display:flex;align-items:center;gap:8px"><div class="barbg"><div class="bar" style="width:'+(v.net/max*100)+'%"></div></div><b>'+usd(v.net)+'</b></div></td><td>'+Math.round(v.net/tot*100)+'%</td></tr>').join('');
  return '<h1>Sales by Country</h1><div class="cap">Net revenue, L30D</div><div class="card"><table><tr><th>Country</th><th>Orders</th><th>Revenue</th><th>% Rev</th></tr>'+rows+'</table></div>';}

const PAGES={overview,scorecard,campaigns,landing,creatives,products,countries};
function render(p){document.getElementById('main').innerHTML=PAGES[p]();}
function setC(r){cR=r;render('campaigns');} function setL(r){lR=r;render('landing');} function setCr(r){crR=r;render('creatives');}
document.querySelectorAll('.navitem').forEach(n=>n.onclick=()=>{document.querySelectorAll('.navitem').forEach(x=>x.classList.remove('active'));n.classList.add('active');render(n.dataset.p);});
render('overview');
</script></body></html>`;
  fs.writeFileSync(DIR+'/dashboard.html',html);
}
