import {validateSitePolygon,polygonCentroidLocalProjection} from './worker_algo.mjs';
import {writeFileSync} from 'fs';
// deterministic PRNG
let s=123456789; const rnd=()=>{s=(s*1103515245+12345)&0x7fffffff; return s/0x7fffffff;};
const out=[];
for(let t=0;t<5000;t++){
  const n=3+Math.floor(rnd()*40);
  const lat0=(rnd()*160)-80, lon0=(rnd()*360)-180;
  const pts=[];
  for(let i=0;i<n;i++){
    // star-shaped polygon => never self-intersecting
    const a=2*Math.PI*i/n, r=0.0005+rnd()*0.05;
    pts.push([(lat0+r*Math.cos(a)).toFixed(8),(lon0+r*Math.sin(a)*1.1).toFixed(8)]);
  }
  if(t%7===0) pts.push([...pts[0]]);            // explicitly closed ring
  if(t%11===0) pts.splice(2,0,[...pts[1]]);      // consecutive duplicate
  out.push({id:t,coords:pts.map(p=>p.join(",")).join(";")});
}
writeFileSync('parity_cases.json',JSON.stringify(out));
const res=[];
for(const c of out){
  try{
    const poly=validateSitePolygon(c.coords.split(";").map(p=>p.split(",").map(Number)));
    const g=polygonCentroidLocalProjection(poly);
    res.push({id:c.id,ok:true,n:poly.length,lat8:g.latitude.toFixed(8),lon8:g.longitude.toFixed(8),area2:g.areaM2.toFixed(2)});
  }catch(e){res.push({id:c.id,ok:false,error:String(e.message||e)});}
}
writeFileSync('parity_worker.json',JSON.stringify(res,null,0));
console.log("worker done:",res.filter(r=>r.ok).length,"ok /",res.length);
