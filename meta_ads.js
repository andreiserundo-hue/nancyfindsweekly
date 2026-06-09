// Direct Meta Marketing API connector (SECONDARY source / failover to Glued)
// Setup: paste a token with ads_read into  meta_token.txt  (one line). Then: node meta_ads.js test
// Ad account is auto-known from Glued. To use a different one, edit ACCOUNT below.
const https = require('https');
const fs = require('fs');
const path = require('path');
const API = 'v21.0';
const ACCOUNT = 'act_943074264794521'; // Nancy Finds (from Glued workspace)
const TOKEN_FILE = path.join(__dirname, 'meta_token.txt');

function token(){
  if(process.env.META_TOKEN) return process.env.META_TOKEN.trim();
  if(!fs.existsSync(TOKEN_FILE)) throw new Error('No Meta token - set META_TOKEN env var or paste into meta_token.txt');
  return fs.readFileSync(TOKEN_FILE,'utf8').trim();
}
function gget(p){return new Promise((res,rej)=>{https.get('https://graph.facebook.com/'+API+p,r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>{let j;try{j=JSON.parse(b);}catch(e){return rej(new Error('bad json: '+b.slice(0,200)));}if(j.error)return rej(new Error(j.error.message));res(j);});}).on('error',rej);});}

// level: 'campaign' | 'adset' | 'ad' ; since/until = YYYY-MM-DD (until is INCLUSIVE in Meta time_range)
async function insights(level, since, until){
  const t = token();
  const fields = 'campaign_name,adset_name,ad_name,spend,purchase_roas,actions,impressions,clicks';
  let url = `/${ACCOUNT}/insights?level=${level}&time_range={'since':'${since}','until':'${until}'}&fields=${fields}&limit=200&access_token=${t}`;
  let rows=[]; let page=await gget(encodeURI(url));
  rows=rows.concat(page.data||[]);
  while(page.paging&&page.paging.next){const u=page.paging.next.replace('https://graph.facebook.com/'+API,'');page=await gget(u);rows=rows.concat(page.data||[]);}
  // normalize → {name, spend, roas, buys}
  return rows.map(r=>{
    const name=r.campaign_name||r.adset_name||r.ad_name||'(unknown)';
    const roas=(r.purchase_roas&&r.purchase_roas[0])?parseFloat(r.purchase_roas[0].value):0;
    const buys=(r.actions||[]).filter(a=>a.action_type==='purchase'||a.action_type==='omni_purchase').reduce((s,a)=>s+parseFloat(a.value||0),0);
    return {name, spend:parseFloat(r.spend||0), roas, buys};
  });
}
module.exports = { insights, ACCOUNT };

// CLI test:  node meta_ads.js test
if(require.main===module && process.argv[2]==='test'){
  (async()=>{
    try{
      const rows=await insights('campaign','2026-06-02','2026-06-08');
      console.log('Meta API OK — '+rows.length+' campaigns (L7D):');
      rows.sort((a,b)=>b.spend-a.spend).slice(0,8).forEach(r=>console.log('  '+r.name.slice(0,40)+' | spend '+Math.round(r.spend)+' | roas '+r.roas.toFixed(2)+' | buys '+r.buys));
    }catch(e){console.log('Meta API not ready:', e.message);}
  })();
}
