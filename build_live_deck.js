// Compact weekly deck generated from LIVE dashboard data bundles
const pptxgen = require("pptxgenjs");
const C={dark:"1A1430",card:"2A2046",berry:"C44B6E",berryDk:"8E2F4C",pinkLt:"F4EEF1",pinkBox:"F7E9EF",white:"FFFFFF",ink:"1A1430",grayTx:"8A8392",grayLt:"BCB5C2",red:"DB3B3B",amber:"E0922F",green:"1F9D6B",rule:"E7DFE4",zebra:"FBF7F9",band:"F3DCE4"};
const FH="Georgia",FB="Calibri",FX=7.8;
const usd=n=>'$'+Math.round(n).toLocaleString(), usdH=n=>'$'+Math.round(n/FX).toLocaleString();
const rcol=v=>v>=1.75?C.green:(v>=1.2?C.amber:C.red);
const clean=n=>(n||'').replace(/Nancy-Testing-2_/,'').replace(/https?:\/\/(www\.)?/,'');
const tier=v=>v>=2.1?['SCALE',C.green]:(v>=1.6?['HOLD',C.amber]:['CUT',C.red]);

module.exports=function(l30,l7,yd,outPath){
  const pptx=new pptxgen();pptx.defineLayout({name:"W",width:13.333,height:7.5});pptx.layout="W";
  const S=pptx.ShapeType,W=pptx.ChartType;let P=0;
  function head(s,k,t){s.addText(k,{x:0.5,y:0.28,w:12,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:2});s.addText(t,{x:0.5,y:0.55,w:11,h:0.6,fontFace:FH,fontSize:28,bold:true,color:C.ink});s.addText("LIVE Â· "+new Date().toLocaleDateString(),{x:9.8,y:0.5,w:3,h:0.5,align:"right",fontFace:FB,fontSize:9,color:C.grayTx});}
  function pg(s){P++;s.addText(String(P),{x:12.7,y:7.05,w:0.4,h:0.3,fontSize:9,color:C.grayLt,align:"right",fontFace:FB});}
  function tbl(s,x,y,cols,cw,rows,rh){let cx=x;cols.forEach((c,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.38,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(c,{x:cx+0.06,y,w:cw[i]-0.1,h:0.38,fontFace:FB,fontSize:9,bold:true,color:C.grayTx,align:i===0?"left":"center",valign:"middle"});cx+=cw[i];});let yy=y+0.38;rows.forEach((r,ri)=>{cx=x;if(ri%2===1)s.addShape(S.rect,{x,y:yy,w:cw.reduce((a,b)=>a+b,0),h:rh,fill:{color:C.zebra},line:{type:"none"}});r.forEach((cell,i)=>{const o=typeof cell==='object'?cell:{t:cell};s.addText(String(o.t),{x:cx+0.05,y:yy,w:cw[i]-0.08,h:rh,fontFace:FB,fontSize:o.fs||11,bold:o.b,color:o.c||C.ink,align:i===0?"left":"center",valign:"middle"});cx+=cw[i];});yy+=rh;});return yy;}

  // 1 Scorecard
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"OVERALL Â· SHOPIFY Ã— META","Scorecard");
    const m=(r,d,key)=>{const sp=r.meta.spendUSD;switch(key){case'rev':return r.shop.net/d;case'spend':return sp/d;case'ord':return r.shop.orders/d;case'ret':return r.shop.retPct;case'meta':return r.meta.metaRoas;case'true':return r.trueRoas;}};
    const D=[[l30,30],[l7,7],[yd,1]];
    const rows=[
      ['Revenue / day',{t:usd(m(l30,30,'rev')),b:1},{t:usd(m(l7,7,'rev')),b:1},{t:usd(m(yd,1,'rev')),b:1},'$1,000','$1,500','$2,000'],
      ['Meta Spend / day',usd(m(l30,30,'spend')),usd(m(l7,7,'spend')),usd(m(yd,1,'spend')),'$667','$857','$1,000'],
      ['Orders / day',m(l30,30,'ord').toFixed(1),m(l7,7,'ord').toFixed(1),m(yd,1,'ord').toFixed(1),'17','25','33'],
      ['Return %',m(l30,30,'ret').toFixed(1)+'%',m(l7,7,'ret').toFixed(1)+'%',m(yd,1,'ret').toFixed(1)+'%','<4%','<3%','<2%'],
      ['Meta ROAS',{t:m(l30,30,'meta').toFixed(2),b:1,c:rcol(m(l30,30,'meta'))},{t:m(l7,7,'meta').toFixed(2),b:1,c:rcol(m(l7,7,'meta'))},{t:m(yd,1,'meta').toFixed(2),b:1,c:rcol(m(yd,1,'meta'))},'1.50','1.75','2.00'],
      ['True ROAS (net)',{t:m(l30,30,'true').toFixed(2),b:1,c:rcol(m(l30,30,'true'))},{t:m(l7,7,'true').toFixed(2),b:1,c:rcol(m(l7,7,'true'))},{t:m(yd,1,'true').toFixed(2),b:1,c:rcol(m(yd,1,'true'))},'1.50','1.75','2.00'],
    ];
    tbl(s,0.5,1.5,['METRIC','L30D','L7D','YEST','GOOD','BETTER','BEST'],[3.0,1.45,1.45,1.55,1.3,1.3,1.3],rows,0.62);
    pg(s);})();
  // 2 Campaigns (L7D)
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"CAMPAIGN ANALYSIS Â· L7D","Campaigns");
    const rows=l7.campaigns.filter(c=>c.spend>0).map(c=>{const t=tier(c.roas);return [clean(c.name).substring(0,52),usdH(c.spend),{t:c.roas.toFixed(2),b:1,c:rcol(c.roas)},c.buys,{t:t[0],b:1,c:t[1]}];});
    tbl(s,0.5,1.5,['CAMPAIGN','SPEND','ROAS','BUYS','TIER'],[5.5,1.7,1.4,1.4,1.6],rows.length?rows:[['No spend','','','','']],0.55);pg(s);})();
  // 3 Landing pages (L7D)
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"LANDING PAGES Â· L7D","Landing Pages");
    const rows=l7.landing.filter(c=>c.spend>0).map(c=>[clean(c.name).substring(0,50),usdH(c.spend),{t:c.roas.toFixed(2),b:1,c:rcol(c.roas)},c.buys]);
    tbl(s,0.5,1.5,['PAGE','SPEND','ROAS','BUYS'],[6.5,1.8,1.6,1.6],rows.length?rows:[['none','','','']],0.5);pg(s);})();
  // 4 Creatives (L7D images)
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"CREATIVES Â· L7D (with purchases)","Creatives");
    const list=l7.creatives.filter(c=>c.buys>0).sort((a,b)=>b.roas-a.roas).slice(0,5);
    const iw=2.32,ig=0.18,ix=0.5,iy=1.55,ih=3.4;const stc={SCALE:C.green,WATCH:C.amber,CUT:C.red};
    list.forEach((c,i)=>{const x=ix+i*(iw+ig);const t=tier(c.roas);s.addShape(S.roundRect,{x:x-0.02,y:iy-0.02,w:iw+0.04,h:4.9,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.05});
      if(c.img){try{s.addImage({path:require('path').join(__dirname,c.img),x:x+0.12,y:iy+0.12,w:iw-0.24,h:ih,sizing:{type:"contain",w:iw-0.24,h:ih}});}catch(e){}}
      s.addText(clean(c.name).substring(0,34),{x:x+0.12,y:iy+ih+0.16,w:iw-0.24,h:0.5,fontFace:FB,fontSize:9.5,bold:true,color:C.ink});
      s.addText([{text:usdH(c.spend)+'  ',options:{color:C.grayTx,fontSize:9}},{text:c.roas.toFixed(2)+'x',options:{color:t[1],bold:true,fontSize:13}}],{x:x+0.12,y:iy+ih+0.62,w:iw-0.24,h:0.3,fontFace:FB});
      s.addText(c.buys+' buys',{x:x+0.12,y:iy+ih+0.92,w:iw-0.24,h:0.3,fontFace:FB,fontSize:9,color:C.grayTx});});
    if(!list.length)s.addText("No converting creatives in L7D",{x:0.5,y:3,w:12,h:0.5,fontFace:FB,fontSize:14,color:C.grayTx});pg(s);})();
  // 5 Products
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"WHAT'S SELLING Â· L30D","Orders by Product");
    const sale=['Blossom Duo','Vine','Eclipse','Bloom','Stem','Sprout'];const pr=l30.shop.products;const tot=Object.values(pr).reduce((a,p)=>a+p.units,0)||1;
    const rows=sale.map(n=>{const p=pr[n]||{units:0,rev:0};return [n,p.units,usd(p.rev),Math.round(p.units/tot*100)+'%'];});
    tbl(s,0.5,1.5,['PRODUCT','UNITS','REVENUE','% UNITS'],[4,2,2.5,2],rows,0.55);pg(s);})();
  // 6 Countries
  (()=>{const s=pptx.addSlide();s.background={color:C.white};head(s,"GEOGRAPHY Â· L30D","Sales by Country");
    const co=Object.entries(l30.shop.countries).sort((a,b)=>b[1].net-a[1].net).slice(0,8);const tot=Object.values(l30.shop.countries).reduce((a,c)=>a+c.net,0)||1;
    const rows=co.map(([c,v])=>[c,v.orders,usd(v.net),Math.round(v.net/tot*100)+'%']);
    tbl(s,0.5,1.5,['COUNTRY','ORDERS','REVENUE','% REV'],[4,2,2.5,2],rows,0.5);pg(s);})();
  // 7 Strategy
  (()=>{const s=pptx.addSlide();s.background={color:C.dark};
    s.addText("STRATEGY & NEXT STEPS",{x:0.6,y:0.5,w:12,h:0.4,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:2});
    s.addText("Based on live L7D performance",{x:0.6,y:0.85,w:12,h:0.5,fontFace:FH,fontSize:26,bold:true,color:C.white});
    const scale=l7.creatives.filter(c=>c.buys>0&&c.roas>=1.75).sort((a,b)=>b.roas-a.roas).slice(0,3).map(c=>'Scale '+clean(c.name).substring(0,30)+' ('+c.roas.toFixed(1)+'x) â†’ ABO creative-test');
    const lpw=l7.landing.filter(l=>l.spend>0&&l.roas>=1.5).sort((a,b)=>b.roas-a.roas)[0];if(lpw)scale.push('Shift budget to '+clean(lpw.name).substring(0,28)+' ('+lpw.roas.toFixed(2)+'x)');
    const fix=l7.campaigns.filter(c=>c.spend>0&&c.roas<1.2).map(c=>'Fix/cut '+clean(c.name).substring(0,30)+' ('+c.roas.toFixed(2)+'x)');
    const struct=['Clear fulfillment backlog (returns '+l7.shop.retPct.toFixed(0)+'% L7D)','Escalate 2nd Meta Business Manager / ad account','Refresh the Aura Scream LP + competitor reframe'];
    const col=(title,c,arr,x)=>{s.addText(title,{x,y:1.7,w:5.8,h:0.3,fontFace:FB,fontSize:12,bold:true,color:c,charSpacing:1});arr.forEach((t,i)=>s.addText([{text:"â€º ",options:{color:c,bold:true}},{text:t,options:{color:C.pinkLt}}],{x,y:2.1+i*0.55,w:5.8,h:0.5,fontFace:FB,fontSize:12,valign:"middle"}));};
    col("â–² DOUBLE DOWN",C.green,scale.length?scale:['â€”'],0.6);
    col("â–  FIX / CUT",C.red,fix.length?fix:['â€”'],6.9);
    s.addText("â¬¢ STRUCTURAL",{x:0.6,y:4.9,w:12,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:1});
    struct.forEach((t,i)=>s.addText([{text:"â€º ",options:{color:C.berry,bold:true}},{text:t,options:{color:C.pinkLt}}],{x:0.6,y:5.3+i*0.5,w:12,h:0.45,fontFace:FB,fontSize:12.5,valign:"middle"}));
    pg(s);})();

  return pptx.writeFile({fileName:outPath}).then(()=>outPath);
};
