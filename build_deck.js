const pptxgen = require("pptxgenjs");
const fs = require("fs");
const DIR = "C:/Users/andre/Nancy Finds Reporting";
const CRE = DIR + "/Creatives";
const LP = DIR + "/LandingPages";
const shop = JSON.parse(fs.readFileSync(DIR + "/shopify_jun9.json", "utf8"));
const FX = 7.8;
const P30=shop.periods['L30D (May 10-Jun 8)'], P7=shop.periods['L7D (Jun 2-Jun 8)'], PY=shop.periods['Yesterday (Jun 8)'];

// palette
const C={dark:"1A1430",card:"2A2046",card2:"4A3D66",berry:"C44B6E",berryDk:"8E2F4C",pinkLt:"F4EEF1",pinkBox:"F7E9EF",white:"FFFFFF",ink:"1A1430",grayTx:"8A8392",grayLt:"BCB5C2",red:"DB3B3B",amber:"E0922F",green:"1F9D6B",rule:"E7DFE4",zebra:"FBF7F9",band:"F3DCE4",hlG:"DCEFE4",hlY:"FBEFD6",hlB:"F6DCE4"};
const FH="Georgia",FB="Calibri",OWNER="OWNER: ANDREI",THRU="DATA THRU 2026-06-09";
const pptx=new pptxgen();pptx.defineLayout({name:"W",width:13.333,height:7.5});pptx.layout="W";pptx.author="Nancy Finds";pptx.title="Nancy Finds â€” Weekly Review";
const W=pptx.ChartType,S=pptx.ShapeType;
function kicker(s,t){s.addText(t,{x:0.5,y:0.28,w:12.3,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:2});}
function title(s,t,c){s.addText(t,{x:0.5,y:0.55,w:9.5,h:0.6,fontFace:FH,fontSize:30,bold:true,color:c||C.ink});}
function meta(s){s.addText([{text:THRU+"\n",options:{fontSize:9,color:C.grayTx}},{text:OWNER,options:{fontSize:9,bold:true,color:C.berryDk}}],{x:9.8,y:0.5,w:3.0,h:0.6,align:"right",fontFace:FB});}
function pageno(s,n){s.addText(String(n),{x:12.7,y:7.05,w:0.4,h:0.3,fontSize:9,color:C.grayLt,align:"right",fontFace:FB});}
function callout(s,y,h,parts){s.addShape(S.roundRect,{x:0.5,y,w:12.33,h,fill:{color:C.pinkBox},line:{type:"none"},rectRadius:0.05});s.addText(parts,{x:0.7,y:y+0.08,w:11.9,h:h-0.16,fontFace:FB,fontSize:11.5,valign:"middle",lineSpacingMultiple:1.05});}
function roasCol(v){return v>=1.75?C.green:(v>=1.2?C.amber:C.red);}
function usd(n){return "$"+Math.round(n).toLocaleString();}
let P=0;

// 1 GOAL
(()=>{const s=pptx.addSlide();P++;s.background={color:C.dark};
  const prodTot=Object.values(shop.products).reduce((a,p)=>a+p.units,0);
  const bdPct=Math.round(shop.products['Blossom Duo'].units/prodTot*100);
  const usOrdPct=Math.round(shop.countries['United States'].orders/shop.periods['Lifetime (May 1-Jun 8)'].orders*100);
  s.addText("NANCY FINDS Â· PAID ACQUISITION",{x:0.6,y:0.5,w:12,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:3});
  s.addText("The Goal â€” Own These Numbers",{x:0.6,y:0.85,w:12,h:0.7,fontFace:FH,fontSize:34,bold:true,color:C.white});
  const ph=[{t:"PHASE 1 Â· FLOOR",b:"$1,000",u:"revenue / day",s:"ROAS â‰¥ 1.50 Â· never breach"},{t:"PHASE 2 Â· HEALTHY",b:"$1,000",u:"revenue / day",s:"ROAS â‰¥ 1.75 Â· efficient scale"},{t:"PHASE 3 Â· GREAT",b:"$2,000",u:"revenue / day",s:"ROAS â‰¥ 2.00 Â· scale volume"}];
  const cw=3.85,gp=0.32,x0=0.6,y0=1.95;
  ph.forEach((p,i)=>{const x=x0+i*(cw+gp);s.addShape(S.roundRect,{x,y:y0,w:cw,h:2.15,fill:{color:C.card},line:{color:i===0?C.berry:C.card2,width:i===0?1.5:1},rectRadius:0.08});
    s.addText(p.t,{x:x+0.25,y:y0+0.2,w:cw-0.5,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:1});
    s.addText(p.b,{x:x+0.25,y:y0+0.55,w:cw-0.5,h:0.8,fontFace:FH,fontSize:44,bold:true,color:C.white});
    s.addText(p.u,{x:x+0.25,y:y0+1.42,w:cw-0.5,h:0.3,fontFace:FB,fontSize:13,color:C.grayLt});
    s.addText(p.s,{x:x+0.25,y:y0+1.72,w:cw-0.5,h:0.3,fontFace:FB,fontSize:11,italic:true,color:C.pinkLt});});
  s.addText("WHERE WE ARE TODAY",{x:0.6,y:4.45,w:12,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:2});
  const trueRoas30=(P30.net/(77319.61/FX)).toFixed(2);
  const now=[{k:"Top product",v:"Blossom Duo",n:bdPct+"% of units",sm:true},{k:"Top market",v:"United States",n:usOrdPct+"% of orders",sm:true},{k:"30-day True ROAS",v:trueRoas30+"x",n:"net Â· vs 1.50 floor"},{k:"Bright spot",v:"nancy-bloom",n:"1.64x L7D LP",sm:true,grn:true}];
  const nw=2.92,ng=0.18,nx=0.6,ny=4.8;
  now.forEach((n,i)=>{const x=nx+i*(nw+ng);s.addShape(S.roundRect,{x,y:ny,w:nw,h:1.5,fill:{color:n.grn?C.berryDk:C.card2},line:{type:"none"},rectRadius:0.06});
    s.addText(n.v,{x:x+0.2,y:ny+0.18,w:nw-0.4,h:0.55,fontFace:FH,fontSize:n.sm?22:28,bold:true,color:C.white,valign:"middle"});
    s.addText(n.k,{x:x+0.2,y:ny+0.78,w:nw-0.4,h:0.3,fontFace:FB,fontSize:12,color:C.pinkLt});
    s.addText(n.n,{x:x+0.2,y:ny+1.08,w:nw-0.4,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayLt});});
  s.addText("Spend is throttled while the fulfillment backlog clears. The forward play: scale the green shoots â€” nancy-bloom + the new editorial creatives â€” to rebuild toward $1,000/day.",{x:0.6,y:6.62,w:12,h:0.4,fontFace:FB,fontSize:12,italic:true,color:C.berry});
  pageno(s,P);})();

// 2 SCORECARD (Jeanine style)
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  const HB="6FA8DC",peach="FCE4D6",blu="DEEAF6",brd="BFBFBF";
  kicker(s,"OVERALL Â· SHOPIFY Ã— META SCORECARD");title(s,"Overall Scorecard â€” Nancy");meta(s);
  s.addText([{text:"â— ",options:{color:C.amber,fontSize:20}},{text:"Mid     ",options:{color:C.ink,fontSize:14}},{text:"â— ",options:{color:C.red,fontSize:20}},{text:"Bad     ",options:{color:C.ink,fontSize:14}},{text:"â— ",options:{color:C.green,fontSize:20}},{text:"Good",options:{color:C.ink,fontSize:14}}],{x:7.8,y:1.12,w:5.0,h:0.4,fontFace:FB,align:"right",valign:"middle"});
  s.addText("Daily averages, USD Â· Net = Shopify sales after refunds Â· True ROAS = Net Ã· Meta spend.",{x:0.5,y:1.18,w:7.2,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const cols=["OVERALL","L30D AVG","L7D AVG","YESTERDAY","GOOD TARGET","BETTER TARGET","BEST TARGET"];
  const cw=[2.6,1.5,1.5,1.55,1.5,1.5,1.5];
  const rows=[
    {l:"Revenue / day",v:["$211","$21","$38"],t:["$1,000","$1,500","$2,000"],band:peach,circ:true,g:1000,b:1500},
    {l:"Meta Spend / day",v:["$330","$51","$56"],t:["$667","$857","$1,000"],band:blu,circ:false},
    {l:"Meta Spend % (of rev)",v:["157%","243%","147%"],t:["<67%","<57%","<50%"],band:"FFFFFF",circ:false},
    {l:"Orders / day",v:["4.0","1.0","1"],t:["17","25","33"],band:"FFFFFF",circ:false},
    {l:"Return %",v:["14.4%","33.3%","0%"],t:["<4%","<3%","<2%"],band:"FFFFFF",circ:false},
    {l:"Meta ROAS",v:["0.50","0.35","0.69"],t:["1.50","1.75","2.00"],band:peach,circ:true,g:1.5,b:1.75},
    {l:"True ROAS (net)",v:["0.64","0.41","0.68"],t:["1.50","1.75","2.00"],band:peach,circ:true,g:1.5,b:1.75},
  ];
  function dot(v,g,b){const n=parseFloat(String(v).replace(/[^0-9.]/g,""));return n>=b?C.green:(n>=g?C.amber:C.red);}
  let x=0.5,y=1.6,cx=x;const hh=0.44,rh=0.5;
  cols.forEach((c,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:hh,fill:{color:HB},line:{color:brd,width:1}});s.addText(c,{x:cx+0.05,y,w:cw[i]-0.08,h:hh,fontFace:FB,fontSize:9.5,bold:true,color:"1A1430",align:"center",valign:"middle"});cx+=cw[i];});
  y+=hh;
  rows.forEach(r=>{cx=x;
    s.addShape(S.rect,{x:cx,y,w:cw[0],h:rh,fill:{color:r.band},line:{color:brd,width:1}});
    s.addText(r.l,{x:cx+0.06,y,w:cw[0]-0.1,h:rh,fontFace:FB,fontSize:11,bold:true,color:C.ink,align:"center",valign:"middle"});cx+=cw[0];
    for(let i=0;i<3;i++){s.addShape(S.rect,{x:cx,y,w:cw[1+i],h:rh,fill:{color:r.band},line:{color:brd,width:1}});
      if(r.circ){const dc=dot(r.v[i],r.g,r.b);s.addText([{text:r.v[i]+" ",options:{color:C.ink,bold:true,fontSize:12}},{text:"â—",options:{color:dc,fontSize:18}}],{x:cx+0.04,y,w:cw[1+i]-0.08,h:rh,fontFace:FB,align:"center",valign:"middle"});}
      else s.addText(r.v[i],{x:cx+0.04,y,w:cw[1+i]-0.08,h:rh,fontFace:FB,fontSize:11.5,color:C.ink,align:"center",valign:"middle"});cx+=cw[1+i];}
    for(let i=0;i<3;i++){s.addShape(S.rect,{x:cx,y,w:cw[4+i],h:rh,fill:{color:"FFFFFF"},line:{color:brd,width:1}});s.addText(r.t[i],{x:cx+0.04,y,w:cw[4+i]-0.08,h:rh,fontFace:FB,fontSize:11,color:C.grayTx,align:"center",valign:"middle"});cx+=cw[4+i];}
    y+=rh;});
  s.addText([{text:"â€º  ",options:{color:C.berry,bold:true}},{text:"Spend throttled hard (L7D $51/day) during the backlog; True ROAS 0.41x L7D as refunds (33%) ate thin sales.",options:{italic:true,color:C.grayTx}}],{x:0.5,y:y+0.18,w:12.3,h:0.3,fontFace:FB,fontSize:11});
  s.addText([{text:"â€º  ",options:{color:C.berry,bold:true}},{text:"Action plan: clear the backlog, then re-scale through the proven LP/creatives (nancy-bloom, new editorial) to recover ROAS toward target.",options:{italic:true,color:C.grayTx}}],{x:0.5,y:y+0.5,w:12.3,h:0.3,fontFace:FB,fontSize:11});
  pageno(s,P);})();

// 3 ORDERS BY PRODUCT
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"WHAT'S SELLING Â· THE TURNING POINT");title(s,"Orders by Product");meta(s);
  s.addText("Units sold (May 1 â€“ Jun 8). Top 3 highlighted. Focus narrowed to the suction line.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const saleable=["Blossom Duo","Vine","Eclipse","Bloom","Stem","Sprout"];const pr=shop.products;
  const tot=Object.values(pr).reduce((a,p)=>a+p.units,0);
  const sale=saleable.map(n=>({name:n,units:pr[n]?pr[n].units:0,revenue:pr[n]?pr[n].revenue:0})).sort((a,b)=>b.units-a.units);
  s.addText("SALEABLE LINE â€” UNITS SOLD",{x:0.5,y:1.55,w:6,h:0.3,fontFace:FB,fontSize:10,bold:true,color:C.berryDk,charSpacing:0.5});
  s.addChart(W.bar,[{name:"Units",labels:sale.map(p=>p.name),values:sale.map(p=>p.units)}],{x:0.5,y:1.9,w:5.8,h:3.5,barDir:"bar",chartColors:[C.berry],showValue:true,dataLabelFontSize:11,dataLabelFontFace:FB,dataLabelColor:C.ink,dataLabelPosition:"outEnd",catAxisLabelFontSize:11,catAxisLabelFontFace:FB,valAxisHidden:true,valAxisMaxVal:85,showLegend:false,barGapWidthPct:40});
  const head=["SALEABLE PRODUCT","UNITS","REVENUE","% UNITS"];const cw=[2.5,1.1,1.6,1.2];let x=6.55,y=1.9,cx=x;const top3=[C.hlG,C.hlY,C.hlB];
  head.forEach((h,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.36,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(h,{x:cx+0.06,y,w:cw[i]-0.1,h:0.36,fontFace:FB,fontSize:9,bold:true,color:C.grayTx,align:i===0?"left":"center",valign:"middle"});cx+=cw[i];});
  y+=0.36;
  sale.forEach((p,ri)=>{cx=x;const bg=ri<3?top3[ri]:(ri%2===1?C.zebra:null);if(bg)s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h:0.42,fill:{color:bg},line:{type:"none"}});
    s.addText(p.name,{x:cx+0.06,y,w:cw[0]-0.1,h:0.42,fontFace:FB,fontSize:11,bold:ri<3,color:ri<3?C.berryDk:C.ink,valign:"middle"});cx+=cw[0];
    s.addText(String(p.units),{x:cx+0.04,y,w:cw[1]-0.08,h:0.42,fontFace:FB,fontSize:11,bold:true,color:C.ink,align:"center",valign:"middle"});cx+=cw[1];
    s.addText(usd(p.revenue),{x:cx+0.04,y,w:cw[2]-0.08,h:0.42,fontFace:FB,fontSize:11,color:C.ink,align:"center",valign:"middle"});cx+=cw[2];
    s.addText(Math.round(p.units/tot*100)+"%",{x:cx+0.04,y,w:cw[3]-0.08,h:0.42,fontFace:FB,fontSize:11,color:C.grayTx,align:"center",valign:"middle"});y+=0.42;});
  const su=sale.reduce((a,p)=>a+p.units,0),sr=sale.reduce((a,p)=>a+p.revenue,0);
  s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h:0.42,fill:{color:C.pinkBox},line:{type:"none"}});cx=x;
  s.addText("Saleable total",{x:cx+0.06,y,w:cw[0]-0.1,h:0.42,fontFace:FB,fontSize:11,bold:true,color:C.berryDk,valign:"middle"});cx+=cw[0];
  s.addText(String(su),{x:cx+0.04,y,w:cw[1]-0.08,h:0.42,fontFace:FB,fontSize:11,bold:true,color:C.berryDk,align:"center",valign:"middle"});cx+=cw[1];
  s.addText(usd(sr),{x:cx+0.04,y,w:cw[2]-0.08,h:0.42,fontFace:FB,fontSize:11,bold:true,color:C.berryDk,align:"center",valign:"middle"});cx+=cw[2];
  s.addText(Math.round(su/tot*100)+"%",{x:cx+0.04,y,w:cw[3]-0.08,h:0.42,fontFace:FB,fontSize:11,bold:true,color:C.berryDk,align:"center",valign:"middle"});
  s.addShape(S.roundRect,{x:0.5,y:5.5,w:12.33,h:0.62,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.05});
  s.addText([{text:"RESERVED (not yet saleable):  ",options:{bold:true,color:C.berryDk}},{text:"Mushroom Cloud, Pebble Rose, Baby Rabbit, Dew + others â€” demand exists, held until stocked/approved; orders can be reserved.",options:{color:C.ink}}],{x:0.7,y:5.56,w:11.9,h:0.5,fontFace:FB,fontSize:10.5,valign:"middle"});
  callout(s,6.22,0.95,[{text:"WHY THIS MATTERS:  ",options:{bold:true,color:C.berryDk}},{text:"Blossom Duo ("+sale[0].units+" units / "+Math.round(sale[0].units/tot*100)+"%) is the runaway hero. The suction line = ",options:{color:C.ink}},{text:su+" of "+tot+" units ("+Math.round(su/tot*100)+"%)",options:{bold:true,color:C.ink}},{text:". Concentrating spend on this line â€” not the full catalog â€” was the turning point.",options:{color:C.ink}}]);
  pageno(s,P);})();

// 4 COUNTRY
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"GEOGRAPHY Â· WHERE SALES COME FROM");title(s,"Sales by Country");meta(s);
  s.addText("Shopify orders & net revenue by destination (May 1 â€“ Jun 8). Top 3 highlighted. The US carries the store.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const co=Object.entries(shop.countries).map(([c,v])=>({c,orders:v.orders,net:v.net})).sort((a,b)=>b.net-a.net);const top=co.slice(0,8);
  const totNet=co.reduce((a,c)=>a+c.net,0);const totOrd=co.reduce((a,c)=>a+c.orders,0);
  s.addText("REVENUE BY COUNTRY (USD)",{x:0.5,y:1.55,w:6,h:0.3,fontFace:FB,fontSize:10,bold:true,color:C.berryDk,charSpacing:0.5});
  s.addChart(W.bar,[{name:"Revenue",labels:top.map(c=>c.c),values:top.map(c=>+c.net.toFixed(0))}],{x:0.5,y:1.9,w:6.0,h:3.7,barDir:"bar",chartColors:[C.berry],showValue:true,dataLabelFontSize:9,dataLabelFontFace:FB,dataLabelColor:C.ink,dataLabelPosition:"outEnd",catAxisLabelFontSize:10,catAxisLabelFontFace:FB,valAxisHidden:true,valAxisMaxVal:5000,showLegend:false,barGapWidthPct:35});
  const head=["COUNTRY","ORDERS","REVENUE","% REV"];const cw=[2.4,1.2,1.6,1.1];let x=6.8,y=1.9,cx=x;const top3=[C.hlG,C.hlY,C.hlB];
  head.forEach((h,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.36,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(h,{x:cx+0.06,y,w:cw[i]-0.1,h:0.36,fontFace:FB,fontSize:9,bold:true,color:C.grayTx,align:i===0?"left":"center",valign:"middle"});cx+=cw[i];});
  y+=0.36;
  top.forEach((c,ri)=>{cx=x;const bg=ri<3?top3[ri]:(ri%2===1?C.zebra:null);if(bg)s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h:0.4,fill:{color:bg},line:{type:"none"}});
    s.addText(c.c,{x:cx+0.06,y,w:cw[0]-0.1,h:0.4,fontFace:FB,fontSize:11,bold:ri<3,color:ri<3?C.berryDk:C.ink,valign:"middle"});cx+=cw[0];
    s.addText(String(c.orders),{x:cx+0.04,y,w:cw[1]-0.08,h:0.4,fontFace:FB,fontSize:11,color:C.ink,align:"center",valign:"middle"});cx+=cw[1];
    s.addText(usd(c.net),{x:cx+0.04,y,w:cw[2]-0.08,h:0.4,fontFace:FB,fontSize:11,bold:ri<3,color:C.ink,align:"center",valign:"middle"});cx+=cw[2];
    s.addText(Math.round(c.net/totNet*100)+"%",{x:cx+0.04,y,w:cw[3]-0.08,h:0.4,fontFace:FB,fontSize:11,color:C.grayTx,align:"center",valign:"middle"});y+=0.4;});
  callout(s,5.85,1.15,[{text:"READ:  ",options:{bold:true,color:C.berryDk}},{text:"The US is ",options:{color:C.ink}},{text:Math.round(co[0].orders/totOrd*100)+"% of orders and "+Math.round(co[0].net/totNet*100)+"% of revenue",options:{bold:true,color:C.ink}},{text:" â€” campaigns target 'Top 15 Countries' but demand concentrates in the US, with Australia & UK secondary. Weight budget & fulfillment toward the US, and test US-only ad sets for efficiency.",options:{color:C.ink}}]);
  pageno(s,P);})();

// 5 CAMPAIGN ANALYSIS (finalized thresholds + budget note)
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"CAMPAIGN ANALYSIS Â· SEGMENT BY PERFORMANCE");title(s,"Campaign Analysis");meta(s);
  s.addText("Active campaigns only, tiered by L7D ROAS.  â— SCALE â‰¥2.1   â— HOLD 1.6â€“2.0   â— CUT â‰¤1.5.  Spend USD.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  function tier(v){return v>=2.1?{c:C.green,l:"SCALE"}:(v>=1.6?{c:C.amber,l:"HOLD"}:{c:C.red,l:"CUT"});}
  const camp=[{n:"Aura Â· Rose Toys Purchase CBO_14May",sp:12710.50,r30:0.34,r7:0.48,ry:1.33,b30:11,b7:2},{n:"NancyFinds Â· Rose Toys Purchase CBO_14May",sp:56408.90,r30:0.61,r7:0.23,ry:0.00,b30:70,b7:1}];
  const head=["CAMPAIGN","SPEND L30D","ROAS 30D","ROAS 7D","ROAS YEST","BUYS 30D","BUYS 7D","TIER"];
  const cw=[4.0,1.45,1.2,1.2,1.25,1.1,1.0,1.1];let x=0.5,y=1.7,cx=x;
  head.forEach((h,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.4,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(h,{x:cx+0.06,y,w:cw[i]-0.1,h:0.4,fontFace:FB,fontSize:9,bold:true,color:C.grayTx,align:i===0||i===7?"left":"center",valign:"middle",charSpacing:0.2});cx+=cw[i];});
  y+=0.4;
  camp.forEach((c,ri)=>{cx=x;const h=0.85;const tt=tier(c.r7);if(ri%2===1)s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h,fill:{color:C.zebra},line:{type:"none"}});
    s.addText(c.n,{x:cx+0.07,y,w:cw[0]-0.12,h,fontFace:FB,fontSize:11,bold:true,color:C.ink,valign:"middle"});cx+=cw[0];
    s.addText(usd(c.sp/FX),{x:cx+0.04,y,w:cw[1]-0.08,h,fontFace:FB,fontSize:11.5,color:C.ink,align:"center",valign:"middle"});cx+=cw[1];
    [c.r30,c.r7,c.ry].forEach((v,vi)=>{s.addText(v.toFixed(2),{x:cx+0.04,y,w:cw[2+vi]-0.08,h,fontFace:FB,fontSize:13,bold:true,color:roasCol(v),align:"center",valign:"middle"});cx+=cw[2+vi];});
    s.addText(String(c.b30),{x:cx+0.04,y,w:cw[5]-0.08,h,fontFace:FB,fontSize:12,color:C.ink,align:"center",valign:"middle"});cx+=cw[5];
    s.addText(String(c.b7),{x:cx+0.04,y,w:cw[6]-0.08,h,fontFace:FB,fontSize:12,color:C.ink,align:"center",valign:"middle"});cx+=cw[6];
    s.addShape(S.ellipse,{x:cx+0.14,y:y+h/2-0.13,w:0.26,h:0.26,fill:{color:tt.c},line:{type:"none"}});
    s.addText(tt.l,{x:cx+0.42,y,w:cw[7]-0.46,h,fontFace:FB,fontSize:9.5,bold:true,color:tt.c,valign:"middle"});y+=h;});
  s.addText("2 campaigns paused (Rose Toys Conversion 01May, Dew Leads 26May) â€” both <0.1x, excluded.  Â·  Current budget: HK$200/day each campaign.",{x:0.5,y:y+0.12,w:12.3,h:0.3,fontFace:FB,fontSize:9.5,italic:true,color:C.grayTx});
  callout(s,y+0.5,1.5,[{text:"THE READ:  ",options:{bold:true,color:C.berryDk}},
    {text:"Both active campaigns are below the 1.5 CUT line on L7D â€” Aura 0.48x, NancyFinds 0.23x â€” though ",options:{color:C.ink}},
    {text:"Aura flickered to 1.33x yesterday",options:{bold:true,color:C.green}},
    {text:". The real green shoots are a layer down: the nancy-bloom bridge (1.64x L7D) and two new editorial creatives. Next: rebuild the active campaigns around those winners and the biird ABO test before adding budget.",options:{color:C.ink}}]);
  pageno(s,P);})();

// 6 BRIDGE PAGES
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"LANDING PAGES Â· WEBSITE BRIDGE PAGES");title(s,"Website / Bridge Pages");meta(s);
  s.addText("Bridge pages per purchasable product â€” built because Meta rejects creatives pointing straight to the PDP. Links open live (30D = May 10â€“Jun 8, 7D = Jun 2â€“Jun 8).",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const head=["BRIDGE PAGE","SPEND 30D","ROAS 30D","ROAS 7D","BUYS 30D","STATUS"];const cw=[5.0,1.5,1.25,1.25,1.2,1.13];let x=0.5,y=1.65,cx=x;
  head.forEach((h,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.4,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(h,{x:cx+0.07,y,w:cw[i]-0.12,h:0.4,fontFace:FB,fontSize:9,bold:true,color:C.grayTx,align:i===0||i===5?"left":"center",valign:"middle",charSpacing:0.2});cx+=cw[i];});
  y+=0.4;
  // name, spend30(HKD), r30, r7(or null), buys30, status
  const rows=[["home.nancyfinds.com/blossom-duo",53010,0.60,0.00,62,"WATCH"],["home.nancyfinds.com/nancy-bloom",1688,1.24,1.64,6,"SCALE"],["home.nancyfinds.com/stem",3491,0.32,0.00,3,"CUT"],["home.nancyfinds.com/eclipse",1059,0.00,null,0,"CUT"],["home.nancyfinds.com/vine",1523,0.00,0.00,0,"CUT"],["home.nancyfinds.com/sprout",646,0.00,null,0,"CUT"]];
  const stc={SCALE:C.green,WATCH:C.amber,CUT:C.red};
  rows.forEach((r,ri)=>{cx=x;const h=0.5;if(ri%2===1)s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h,fill:{color:C.zebra},line:{type:"none"}});
    s.addText(r[0],{x:cx+0.07,y,w:cw[0]-0.12,h,fontFace:FB,fontSize:11,bold:true,color:C.berry,valign:"middle"});cx+=cw[0];
    s.addText(usd(r[1]/FX),{x:cx+0.04,y,w:cw[1]-0.08,h,fontFace:FB,fontSize:11,color:C.ink,align:"center",valign:"middle"});cx+=cw[1];
    s.addText(r[2].toFixed(2),{x:cx+0.04,y,w:cw[2]-0.08,h,fontFace:FB,fontSize:11,bold:true,color:roasCol(r[2]),align:"center",valign:"middle"});cx+=cw[2];
    s.addText(r[3]===null?"â€”":r[3].toFixed(2),{x:cx+0.04,y,w:cw[3]-0.08,h,fontFace:FB,fontSize:11,bold:r[3]!==null,color:r[3]===null?C.grayLt:roasCol(r[3]),align:"center",valign:"middle"});cx+=cw[3];
    s.addText(String(r[4]),{x:cx+0.04,y,w:cw[4]-0.08,h,fontFace:FB,fontSize:11,color:C.ink,align:"center",valign:"middle"});cx+=cw[4];
    s.addText(r[5],{x:cx+0.07,y,w:cw[5]-0.12,h,fontFace:FB,fontSize:9.5,bold:true,color:stc[r[5]],valign:"middle"});y+=h;});
  callout(s,y+0.2,1.5,[{text:"WHY BRIDGE PAGES:  ",options:{bold:true,color:C.berryDk}},{text:"Meta rejects ads linking directly to explicit PDPs, so every product gets a compliant bridge page in between â€” this unlocked paid traffic to the catalog.  ",options:{color:C.ink}},{text:"nancy-bloom is the new winner â€” 1.24x L30D and 1.64x L7D",options:{bold:true,color:C.green}},{text:"; blossom-duo carries volume (62 buys) but cooled to 0.00x L7D. Next: shift budget toward nancy-bloom and the Aura Scream funnel.",options:{color:C.ink}}]);
  pageno(s,P);})();

// 7 AURA LANDING PAGES
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"LANDING PAGES Â· AURA ADVERTORIALS");title(s,"Aura Landing Pages");meta(s);
  s.addText("The two Aura listicle advertorials â€” performance + how they look.  30D = May 10â€“Jun 8 Â· 7D = Jun 2â€“Jun 8.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const iw=5.3,ih=2.55;
  s.addImage({path:LP+"/aura scream lp.png",x:0.5,y:1.6,w:iw,h:ih,sizing:{type:"contain",w:iw,h:ih}});
  s.addText([{text:"Scream LP   ",options:{bold:true,color:C.berryDk,fontSize:10}},{text:"aura.today/landing-rose-toys-scream",options:{italic:true,color:C.berry,fontSize:9}}],{x:0.5,y:4.18,w:iw,h:0.25,fontFace:FB});
  s.addImage({path:LP+"/aura rose toys lp.png",x:0.5,y:4.5,w:iw,h:ih,sizing:{type:"contain",w:iw,h:ih}});
  s.addText([{text:"Best Rose Toys 2026 LP   ",options:{bold:true,color:C.berryDk,fontSize:10}},{text:"aura.today/landing-best-rose-toys-2026",options:{italic:true,color:C.berry,fontSize:9}}],{x:0.5,y:7.08,w:iw,h:0.25,fontFace:FB});
  const head=["AURA LP","SPEND 30D","SPEND 7D","ROAS 30D","ROAS 7D","BUYS 30D"];const cw=[2.0,0.98,0.96,0.96,0.88,0.82];let x=6.1,y=1.6,cx=x;
  head.forEach((h,i)=>{s.addShape(S.rect,{x:cx,y,w:cw[i],h:0.42,fill:{color:C.pinkLt},line:{type:"none"}});s.addText(h,{x:cx+0.05,y,w:cw[i]-0.08,h:0.42,fontFace:FB,fontSize:8.5,bold:true,color:C.grayTx,align:i===0?"left":"center",valign:"middle"});cx+=cw[i];});
  y+=0.42;
  const rows=[["Rose Toys Scream",6857,1345.89,0.56,0.49,10],["Best Rose Toys 2026",5940,null,0.08,null,1]];
  rows.forEach((r,ri)=>{cx=x;const h=0.62;if(ri%2===1)s.addShape(S.rect,{x,y,w:cw.reduce((a,b)=>a+b,0),h,fill:{color:C.zebra},line:{type:"none"}});
    s.addText(r[0],{x:cx+0.06,y,w:cw[0]-0.1,h,fontFace:FB,fontSize:10.5,bold:true,color:C.ink,valign:"middle"});cx+=cw[0];
    s.addText(usd(r[1]/FX),{x:cx+0.03,y,w:cw[1]-0.06,h,fontFace:FB,fontSize:10.5,color:C.ink,align:"center",valign:"middle"});cx+=cw[1];
    s.addText(r[2]===null?"â€”":usd(r[2]/FX),{x:cx+0.03,y,w:cw[2]-0.06,h,fontFace:FB,fontSize:10.5,color:r[2]===null?C.grayLt:C.ink,align:"center",valign:"middle"});cx+=cw[2];
    s.addText(r[3].toFixed(2),{x:cx+0.03,y,w:cw[3]-0.06,h,fontFace:FB,fontSize:11,bold:true,color:roasCol(r[3]),align:"center",valign:"middle"});cx+=cw[3];
    s.addText(r[4]===null?"â€”":r[4].toFixed(2),{x:cx+0.03,y,w:cw[4]-0.06,h,fontFace:FB,fontSize:11,bold:r[4]!==null,color:r[4]===null?C.grayLt:roasCol(r[4]),align:"center",valign:"middle"});cx+=cw[4];
    s.addText(String(r[5]),{x:cx+0.03,y,w:cw[5]-0.06,h,fontFace:FB,fontSize:10.5,color:C.ink,align:"center",valign:"middle"});y+=h;});
  s.addShape(S.roundRect,{x:6.1,y:3.2,w:6.73,h:3.05,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.05});
  s.addText("THE READ",{x:6.3,y:3.32,w:6.4,h:0.3,fontFace:FB,fontSize:10,bold:true,color:C.berry,charSpacing:1});
  s.addText([{text:"The Scream advertorial",options:{bold:true,color:C.berryDk}},{text:" carried the funnel mid-May (peaked ~1.79x) but ",options:{color:C.ink}},{text:"cooled to 0.49x L7D",options:{bold:true,color:C.red}},{text:" â€” it needs a creative/offer refresh. The ",options:{color:C.ink}},{text:"\"Best Rose Toys 2026\"",options:{bold:true,color:C.ink}},{text:" version bled at 0.08x and was cut ~May 22.\n\nBoth follow the Hello Nancy listicle playbook; Scream's visceral hook still converts best of the two. Pair it with the new editorial creatives to re-lift ROAS.",options:{color:C.ink}}],{x:6.3,y:3.6,w:6.4,h:2.55,fontFace:FB,fontSize:11.5,valign:"top",lineSpacingMultiple:1.05});
  pageno(s,P);})();

// 8 LANDING PAGE STUDY & REFRAME
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"FUNNEL STUDY Â· THE TWO ADVERTORIALS");title(s,"Landing Page Study & Reframe");meta(s);
  const cols=[{tag:"KILLED ~MAY 22",head:'"Best Rose Toys 2026"',url:"aura.today/landing-best-rose-toys-2026",hook:'"The 8 Rose Toys Worth Knowing in 2026"',angle:"Curiosity / category education",roas:"0.08x  (30D)",col:C.red,pts:["Neutral, intellectual tone","Broke scent-trail from 'scream' ad copy","No add-to-carts, no purchases for 3 days","Bled $762 at 0.08x"]},{tag:"KEPT & REFRAMED",head:'"Rose Toys Scream"',url:"aura.today/landing-rose-toys-scream",hook:'"8 Rose Toys So Good You\'ll Scream Out Loud"',angle:"Visceral / outcome-led ('scream' Ã—11)",roas:"0.56x  (30D)",col:C.amber,pts:["Matches winning ad copy exactly","Blossom Duo #1 â€” 'which to actually buy'","Competitors reframed as beginner/starter","Peaked 1.79x; cooled to 0.49x L7D â€” refresh"]}];
  cols.forEach((c,i)=>{const x=0.5+i*6.25;s.addShape(S.roundRect,{x,y:1.35,w:6.0,h:4.0,fill:{color:C.pinkLt},line:{color:c.col,width:1.5},rectRadius:0.06});
    s.addShape(S.roundRect,{x:x+0.25,y:1.55,w:2.6,h:0.34,fill:{color:c.col},line:{type:"none"},rectRadius:0.04});
    s.addText(c.tag,{x:x+0.25,y:1.55,w:2.6,h:0.34,fontFace:FB,fontSize:10,bold:true,color:C.white,align:"center",valign:"middle"});
    s.addText(c.roas,{x:x+3.0,y:1.5,w:2.75,h:0.42,fontFace:FH,fontSize:22,bold:true,color:c.col,align:"right"});
    s.addText(c.head,{x:x+0.25,y:2.0,w:5.5,h:0.4,fontFace:FH,fontSize:19,bold:true,color:C.ink});
    s.addText(c.url,{x:x+0.25,y:2.42,w:5.5,h:0.25,fontFace:FB,fontSize:9.5,italic:true,color:C.berry});
    s.addText([{text:"Hook  ",options:{bold:true,color:C.grayTx,fontSize:9}},{text:c.hook+"\n",options:{color:C.ink,fontSize:10.5}},{text:"Angle  ",options:{bold:true,color:C.grayTx,fontSize:9}},{text:c.angle,options:{color:C.ink,fontSize:10.5}}],{x:x+0.25,y:2.72,w:5.5,h:0.85,fontFace:FB,lineSpacingMultiple:1.1});
    c.pts.forEach((p,pi)=>s.addText([{text:"â€º  ",options:{color:c.col,bold:true}},{text:p,options:{color:C.ink}}],{x:x+0.25,y:3.62+pi*0.42,w:5.5,h:0.4,fontFace:FB,fontSize:10.5,valign:"middle"}));});
  s.addShape(S.roundRect,{x:0.5,y:5.55,w:12.33,h:1.5,fill:{color:C.dark},line:{type:"none"},rectRadius:0.06});
  s.addText("THE REFRAME DIRECTIVE â€” LANDED",{x:0.75,y:5.7,w:12,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:1});
  s.addText([{text:"Competitors are now boxed as lesser options: ",options:{color:C.pinkLt}},{text:"Satisfyer = 'Best for First-Timers,' Lyra = 'starter / travel toy,' The Rose = style-over-substance",options:{color:C.white,bold:true}},{text:" â€” while ",options:{color:C.pinkLt}},{text:"Blossom Duo owns 'which one should I actually buy.'",options:{color:C.white,bold:true}},{text:"  Watch item: competitors still rank #2 & #4 above 3 Nancy products â€” next lever is demoting them below.",options:{color:C.grayLt,italic:true}}],{x:0.75,y:6.0,w:11.85,h:1.0,fontFace:FB,fontSize:11.5,valign:"top",lineSpacingMultiple:1.05});
  pageno(s,P);})();

// 9 CREATIVE / ADS ANALYSIS (new heroes + 2 new images)
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"DOUBLE DOWN VS CUT");title(s,"Creative â€” Ads Analysis");meta(s);
  s.addText("Big-spend studio shots vs the new tiny-budget winners (USD Â· 30D = May 10â€“Jun 8, recent = L7D).",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const heroes=[
    {img:CRE+"/Blossom Duo_Image_20_capoff_windowsill_May-14.jpg",name:"Blossom Duo Â· Windowsill",spend:"$2,391 (30D)",roas:"0.73x",buys:"23",st:"CUT"},
    {img:CRE+"/01_he_loves_it_too_May-16.jpg",name:"01 Â· He Loves It Too",spend:"$495 (30D)",roas:"0.22x",buys:"2",st:"CUT"},
    {img:CRE+"/02_May-18.jpg",name:"02 Â· So Good You'll Scream",spend:"$120 (30D)",roas:"2.02x",buys:"4",st:"WATCH"},
    {img:CRE+"/aura_05_stays_with_you_May-16.jpg",name:"aura_05 Â· Stays With You â˜…NEW",spend:"<$1 (L7D)",roas:"104x",buys:"1",st:"SCALE"},
    {img:CRE+"/Bloom_Image_angle31a_fortune_cookie_Jun-02.jpg",name:"Bloom Â· Fortune Cookie â˜…NEW",spend:"<$1 (L7D)",roas:"120x",buys:"1",st:"SCALE"},
  ];
  const stc={SCALE:C.green,WATCH:C.amber,CUT:C.red};const iw=2.32,ig=0.18,ix=0.5,iy=1.6,ih=3.0;
  heroes.forEach((h,i)=>{const x=ix+i*(iw+ig);s.addShape(S.roundRect,{x:x-0.02,y:iy-0.02,w:iw+0.04,h:4.6,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.05});
    s.addImage({path:h.img,x:x+0.12,y:iy+0.12,w:iw-0.24,h:ih,sizing:{type:"cover",w:iw-0.24,h:ih}});
    s.addText(h.name,{x:x+0.12,y:iy+ih+0.18,w:iw-0.24,h:0.5,fontFace:FB,fontSize:9.5,bold:true,color:C.ink});
    s.addText([{text:h.spend+"  ",options:{color:C.grayTx,fontSize:9}},{text:h.roas,options:{color:stc[h.st],bold:true,fontSize:12}}],{x:x+0.12,y:iy+ih+0.66,w:iw-0.24,h:0.3,fontFace:FB});
    s.addShape(S.roundRect,{x:x+0.12,y:iy+ih+1.0,w:1.5,h:0.3,fill:{color:stc[h.st]},line:{type:"none"},rectRadius:0.04});
    s.addText(h.st+" Â· "+h.buys+" buys",{x:x+0.12,y:iy+ih+1.0,w:1.9,h:0.3,fontFace:FB,fontSize:8.5,bold:true,color:C.white,valign:"middle",align:"left"});});
  callout(s,6.35,0.85,[{text:"THE PATTERN:  ",options:{bold:true,color:C.berryDk}},{text:"The big studio shots ate most spend at sub-0.75x. The new editorial / lifestyle creatives (Stays With You, Fortune Cookie) hit 100x+ on pennies. ",options:{color:C.ink}},{text:"Feed budget to the new editorial winners via the biird ABO test.",options:{bold:true,color:C.ink}}]);
  pageno(s,P);})();

// 10 HELLO NANCY -> NANCY FINDS
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"PLAYBOOK TRANSFER Â· PROVEN WINNERS");title(s,"Hello Nancy â†’ Nancy Finds");meta(s);
  s.addText("We didn't start from zero. We ported what already won on the sister brand (Hello Nancy) into Nancy Finds.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  s.addShape(S.roundRect,{x:0.5,y:1.6,w:4.3,h:4.0,fill:{color:C.card},line:{type:"none"},rectRadius:0.06});
  s.addText("HELLO NANCY",{x:0.7,y:1.78,w:3.9,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:2});
  s.addText("Sister brand Â· proven at scale",{x:0.7,y:2.08,w:3.9,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayLt});
  ["Winning on-creative text hooks","Rose-toy ad creatives","Advertorial / listicle structure","Bridge-page funnel idea"].forEach((t,i)=>s.addText([{text:"âœ“  ",options:{color:C.berry,bold:true}},{text:t,options:{color:C.pinkLt}}],{x:0.7,y:2.5+i*0.62,w:3.9,h:0.55,fontFace:FB,fontSize:12,valign:"middle"}));
  s.addShape(S.rightArrow,{x:4.95,y:3.3,w:1.05,h:0.7,fill:{color:C.berry},line:{type:"none"}});
  s.addText("ADAPTED",{x:4.85,y:4.05,w:1.25,h:0.3,fontFace:FB,fontSize:9,bold:true,color:C.berryDk,align:"center"});
  s.addShape(S.roundRect,{x:6.15,y:1.6,w:6.65,h:4.0,fill:{color:C.pinkLt},line:{color:C.berry,width:1.5},rectRadius:0.06});
  s.addText("NANCY FINDS",{x:6.4,y:1.78,w:6.2,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berryDk,charSpacing:2});
  s.addText("New brand Â· winners transplanted",{x:6.4,y:2.08,w:6.2,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const ad=[["\"So Good, You'll Scream\"","On-creative winner â†’ top hook + the LP theme"],["\"He Loves It Too\"","On-creative winner â†’ 16 purchases (2nd-highest)"],["Rose Ads","Hello Nancy rose creatives re-shot for Nancy Finds"],["Bridge Pages","Funnel idea ported â†’ lowered Meta rejection risk"]];
  ad.forEach((a,i)=>{const yy=2.5+i*0.72;s.addText(a[0],{x:6.4,y:yy,w:6.2,h:0.3,fontFace:FH,fontSize:14,bold:true,italic:true,color:C.berryDk});s.addText(a[1],{x:6.4,y:yy+0.3,w:6.2,h:0.3,fontFace:FB,fontSize:10.5,color:C.ink});});
  callout(s,5.8,1.1,[{text:"WHY IT WORKED:  ",options:{bold:true,color:C.berryDk}},{text:"Porting battle-tested copy, creative and the bridge-page funnel skipped months of testing â€” the 'scream' hook and Rose ads were already validated on Hello Nancy. ",options:{color:C.ink}},{text:"Nancy Finds inherited a proven playbook on day one; the job now is fitting it to Nancy Finds' catalog + Meta's stricter rules (see Bottlenecks).",options:{color:C.ink}}]);
  pageno(s,P);})();

// 11 PIGGYBACK
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"GROWTH LEVER Â· CROSS-BRAND PLACEMENT");title(s,"Piggyback a Proven Article");meta(s);
  s.addText("Hello Nancy's suction-toy advertorial already ranks & converts. Snug a Nancy Finds link onto it â€” as a special mention, not a forced #1.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  s.addShape(S.roundRect,{x:0.5,y:1.6,w:5.5,h:4.0,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.06});
  s.addText("THE ASSET (Hello Nancy)",{x:0.72,y:1.75,w:5.1,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berryDk,charSpacing:1});
  s.addText('"7 Suction Toys So Good, You\'ll Want to Cheat"',{x:0.72,y:2.05,w:5.1,h:0.55,fontFace:FH,fontSize:15,bold:true,italic:true,color:C.ink});
  s.addText("zenify.today/article/7-best-suction-toys-for-women",{x:0.72,y:2.62,w:5.1,h:0.25,fontFace:FB,fontSize:9.5,italic:true,color:C.berry});
  [["1","Lem â€” Hello Nancy",true],["2","Namii 2 â€” Biird",false],["3","Sona 2 â€” LELO",false],["4","Avo â€” Hello Nancy",true],["5","Satisfyer Pro 2",false],["6","Aer â€” Dame",false],["7","The Rose",false]].forEach((r,i)=>{const yy=3.0+i*0.36;s.addText([{text:r[0]+"  ",options:{bold:true,color:r[2]?C.berry:C.grayLt}},{text:r[1],options:{color:r[2]?C.ink:C.grayTx,bold:r[2]}}],{x:0.72,y:yy,w:5.1,h:0.34,fontFace:FB,fontSize:10.5,valign:"middle"});});
  s.addText("All buy-links â†’ hellonancy.com. No Nancy Finds product on it yet.",{x:0.72,y:5.32,w:5.1,h:0.25,fontFace:FB,fontSize:9,italic:true,color:C.grayTx});
  s.addShape(S.roundRect,{x:6.3,y:1.6,w:6.5,h:4.0,fill:{color:C.dark},line:{type:"none"},rectRadius:0.06});
  s.addText("THE MOVE â€” EDITOR'S WILDCARD",{x:6.55,y:1.75,w:6,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:1});
  s.addImage({path:CRE+"/02_May-18.jpg",x:6.55,y:2.15,w:1.85,h:2.95,sizing:{type:"cover",w:1.85,h:2.95}});
  s.addText("Blossom Duo",{x:8.6,y:2.2,w:4.0,h:0.4,fontFace:FH,fontSize:20,bold:true,color:C.white});
  s.addText("Nancy Finds Â· special mention",{x:8.6,y:2.62,w:4.0,h:0.25,fontFace:FB,fontSize:10,italic:true,color:C.grayLt});
  s.addText([{text:'"The rose-shaped one that didn\'t fit the ranking but kept stealing the spotlight. Where Lem wins on raw air-pulse power, Blossom Duo plays a different game â€” a magnetic cap that flips between suction and a tapping mode, in a pocket-sized rose. The design-forward pick to watch."',options:{color:C.pinkLt,italic:true}}],{x:8.6,y:2.95,w:4.05,h:1.7,fontFace:FB,fontSize:11,valign:"top",lineSpacingMultiple:1.05});
  s.addShape(S.roundRect,{x:8.6,y:4.7,w:2.6,h:0.34,fill:{color:C.berry},line:{type:"none"},rectRadius:0.04});
  s.addText("â†’ nancyfinds.com",{x:8.6,y:4.7,w:2.6,h:0.34,fontFace:FB,fontSize:10,bold:true,color:C.white,align:"center",valign:"middle"});
  callout(s,5.8,1.2,[{text:"WHY SPECIAL MENTION (not top 7):  ",options:{bold:true,color:C.berryDk}},{text:"Ranking Blossom Duo at/above Lem would cannibalize Hello Nancy's hero and dent credibility. A distinct-lane wildcard â€” 'rose + dual mechanism,' which Lem doesn't offer â€” is ",options:{color:C.ink}},{text:"additive, not competitive",options:{bold:true,color:C.ink}},{text:": it captures rose-seekers, sends them to Nancy Finds, and leaves Lem's #1 volume intact.",options:{color:C.ink}}]);
  pageno(s,P);})();

// 12 BOTTLENECKS
(()=>{const s=pptx.addSlide();P++;s.background={color:C.white};
  kicker(s,"WHAT'S HOLDING US BACK");title(s,"Bottlenecks & Limitations");meta(s);
  s.addText("Four structural constraints. The first is why spend is throttled right now; the rest cap how fast we can scale.",{x:0.5,y:1.15,w:12.3,h:0.3,fontFace:FB,fontSize:10,italic:true,color:C.grayTx});
  const items=[{num:"1",sev:"ACTIVE",sc:C.red,title:"Fulfillment / shipment backlog",body:"Packaging is delayed and past orders are still pending shipment, so we keep new orders minimal until it clears.",impl:"Implication: revenue capped on purpose. Scaling now â†’ unshipped orders â†’ refunds â†’ chargebacks â†’ Shopify flags the store."},{num:"2",sev:"ONGOING",sc:C.amber,title:"Meta creative rejection risk",body:"Meta rejects product uploads outright, so creatives use Rose-Flower imagery + only non-explicit rose toys.",impl:"Implication: bridge pages lowered â€” but didn't eliminate â€” rejection; narrows creative variety and slows testing speed."},{num:"3",sev:"BLOCKER",sc:C.red,title:"Single ad account Â· no new BM",body:"One ad account only. Attempts to create a new Business Manager / Portfolio fail with errors â€” across the whole team.",impl:"Implication: no failover, no parallel scaling headroom, all risk in one asset. A single ban halts all paid acquisition."},{num:"4",sev:"DATA",sc:C.amber,title:"Attribution & manual reporting",body:"Meta under-credits Shopify revenue; Glued data is still pulled by hand each time.",impl:"Implication: true ROAS only visible by stitching Shopify + Glued manually; risks lag in spotting a floor miss."}];
  const cw=6.0,ch=2.35,gx=0.33,gy=0.25,x0=0.5,y0=1.6;
  items.forEach((it,i)=>{const col=i%2,row=Math.floor(i/2),x=x0+col*(cw+gx),y=y0+row*(ch+gy);
    s.addShape(S.roundRect,{x,y,w:cw,h:ch,fill:{color:C.pinkLt},line:{color:C.rule,width:1},rectRadius:0.06});
    s.addShape(S.ellipse,{x:x+0.25,y:y+0.25,w:0.55,h:0.55,fill:{color:C.berry},line:{type:"none"}});
    s.addText(it.num,{x:x+0.25,y:y+0.25,w:0.55,h:0.55,fontFace:FH,fontSize:22,bold:true,color:C.white,align:"center",valign:"middle"});
    s.addShape(S.roundRect,{x:x+cw-1.5,y:y+0.28,w:1.25,h:0.32,fill:{color:it.sc},line:{type:"none"},rectRadius:0.04});
    s.addText(it.sev,{x:x+cw-1.5,y:y+0.28,w:1.25,h:0.32,fontFace:FB,fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
    s.addText(it.title,{x:x+0.95,y:y+0.22,w:cw-2.5,h:0.6,fontFace:FH,fontSize:15,bold:true,color:C.ink,valign:"middle"});
    s.addText(it.body,{x:x+0.27,y:y+0.92,w:cw-0.54,h:0.7,fontFace:FB,fontSize:10.5,color:C.ink,valign:"top",lineSpacingMultiple:1.0});
    s.addText([{text:"â†’ ",options:{color:it.sc,bold:true}},{text:it.impl,options:{color:C.berryDk,italic:true}}],{x:x+0.27,y:y+1.55,w:cw-0.54,h:0.72,fontFace:FB,fontSize:10,valign:"top",lineSpacingMultiple:1.0});});
  pageno(s,P);})();

// 13 SUMMARY
(()=>{const s=pptx.addSlide();P++;s.background={color:C.dark};
  s.addText("NANCY FINDS Â· SUMMARY",{x:0.6,y:0.4,w:12,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:3});
  s.addText("What We Did Â· What Worked Â· What's Next",{x:0.6,y:0.72,w:12.2,h:0.6,fontFace:FH,fontSize:28,bold:true,color:C.white});
  const colW=3.9,gap=0.22,x0=0.6,y0=1.55,h=3.5;
  s.addShape(S.roundRect,{x:x0,y:y0,w:colW,h,fill:{color:C.card},line:{type:"none"},rectRadius:0.06});
  s.addText("WHAT WE DID",{x:x0+0.25,y:y0+0.16,w:colW-0.5,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:1});
  ["Adapted Hello Nancy winner text, creatives, LPs & the bridge-page idea","Built the 2 Aura listicles (Scream + Best 2026)","Used Rose-Flower imagery + only non-explicit rose toys â€” Meta rejects product uploads","Bridge pages lowered (not eliminated) Meta rejection","Focused spend on the suction line (Blossom Duo = top seller)"].forEach((t,i)=>s.addText([{text:"â€º  ",options:{color:C.berry,bold:true}},{text:t,options:{color:C.pinkLt}}],{x:x0+0.25,y:y0+0.55+i*0.575,w:colW-0.5,h:0.55,fontFace:FB,fontSize:10.5,valign:"top",lineSpacingMultiple:0.98}));
  const x1=x0+colW+gap;s.addShape(S.roundRect,{x:x1,y:y0,w:colW,h,fill:{color:C.card},line:{type:"none"},rectRadius:0.06});
  s.addText("WORKED / DIDN'T",{x:x1+0.25,y:y0+0.16,w:colW-0.5,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.berry,charSpacing:1});
  s.addText([{text:"WORKED\n",options:{color:C.green,bold:true,fontSize:11}},{text:"Bridge pages Â· nancy-bloom LP (1.64x L7D) Â· new editorial creatives (100x+) Â· suction-line focus Â· Hello Nancy playbook",options:{color:C.pinkLt,fontSize:10.5}}],{x:x1+0.25,y:y0+0.55,w:colW-0.5,h:1.5,fontFace:FB,valign:"top",lineSpacingMultiple:1.0});
  s.addText([{text:"DIDN'T\n",options:{color:C.red,bold:true,fontSize:11}},{text:"Direct-to-PDP ads (Meta rejects) Â· 'Best 2026' listicle (0.08x) Â· studio product shots Â· backlog refunds (33% L7D)",options:{color:C.pinkLt,fontSize:10.5}}],{x:x1+0.25,y:y0+2.05,w:colW-0.5,h:1.4,fontFace:FB,valign:"top",lineSpacingMultiple:1.0});
  const x2=x1+colW+gap;s.addShape(S.roundRect,{x:x2,y:y0,w:colW,h,fill:{color:C.berryDk},line:{type:"none"},rectRadius:0.06});
  s.addText("WHAT'S NEXT",{x:x2+0.25,y:y0+0.16,w:colW-0.5,h:0.3,fontFace:FB,fontSize:12,bold:true,color:C.white,charSpacing:1});
  ["Clear the shipment backlog first â€” stop the refund leak","Re-scale via winners: nancy-bloom LP + new editorial creatives","Refresh the Aura Scream LP (cooled to 0.49x)","Escalate the Meta BM / ad-account blocker","Run biird's ABO creative-testing setup (spec below)"].forEach((t,i)=>s.addText([{text:"â€º  ",options:{color:C.white,bold:true}},{text:t,options:{color:C.pinkLt}}],{x:x2+0.25,y:y0+0.55+i*0.56,w:colW-0.5,h:0.55,fontFace:FB,fontSize:10.5,valign:"top",lineSpacingMultiple:0.96}));
  s.addShape(S.roundRect,{x:0.6,y:5.25,w:12.2,h:1.05,fill:{color:C.card2},line:{type:"none"},rectRadius:0.06});
  s.addText("biird ABO TEST SETUP",{x:0.8,y:5.35,w:4,h:0.3,fontFace:FB,fontSize:11,bold:true,color:C.berry,charSpacing:1});
  s.addText([{text:"Each ad set = ",options:{color:C.pinkLt}},{text:"HK$200/day with ~20 creatives",options:{bold:true,color:C.white}},{text:". Kill any creative with ",options:{color:C.pinkLt}},{text:"0 purchases in 3 days",options:{bold:true,color:C.white}},{text:"; let winners run to ",options:{color:C.pinkLt}},{text:"5â€“7 days, then kill if ROAS < 1.5",options:{bold:true,color:C.white}},{text:".  Run across both domains â€” direct-to-website and via the Aura Scream LP.",options:{color:C.pinkLt}}],{x:0.8,y:5.65,w:11.8,h:0.6,fontFace:FB,fontSize:11.5,valign:"top",lineSpacingMultiple:1.05});
  s.addText("DECISIONS NEEDED:  backlog clear-by date Â· approve post-backlog scale-up Â· greenlight biird ABO test Â· escalate Meta BM blocker",{x:0.6,y:6.5,w:12.2,h:0.5,fontFace:FB,fontSize:11,italic:true,color:C.grayLt});
  pageno(s,P);})();

pptx.writeFile({fileName:DIR+"/Nancy Finds - Weekly Review.pptx"}).then(f=>console.log("WROTE:",f,"| slides:",P));
