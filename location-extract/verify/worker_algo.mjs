// EXACT copy of k12/worker_k12.js lines 555-601 (validateSitePolygon + polygonCentroidLocalProjection)
function fixed8(v){return Number(v).toFixed(8);}
function validCoordinate(lat,lon){return Number.isFinite(lat)&&Number.isFinite(lon)&&lat>=-90&&lat<=90&&lon>=-180&&lon<=180;}
export function validateSitePolygon(input){
  if(!Array.isArray(input)) throw new Error("SITE polygon must be an array.");
  if(input.length<3||input.length>1000) throw new Error("SITE polygon requires 3 to 1000 vertices.");
  const out=[];
  for(const item of input){
    const lat=Number(Array.isArray(item)?item[0]:item?.latitude);
    const lon=Number(Array.isArray(item)?item[1]:item?.longitude);
    if(!validCoordinate(lat,lon)) throw new Error("SITE polygon contains an invalid coordinate.");
    const p=[Number(fixed8(lat)),Number(fixed8(lon))];
    const last=out[out.length-1];
    if(!last||last[0]!==p[0]||last[1]!==p[1]) out.push(p);
  }
  if(out.length>1&&out[0][0]===out[out.length-1][0]&&out[0][1]===out[out.length-1][1]) out.pop();
  if(out.length<3) throw new Error("SITE polygon has fewer than three distinct vertices.");
  const distinct=new Set(out.map(p=>`${p[0]},${p[1]}`));
  if(distinct.size<3) throw new Error("SITE polygon has fewer than three distinct vertices.");
  const centroid=polygonCentroidLocalProjection(out);
  if(!Number.isFinite(centroid.areaM2)||centroid.areaM2<0.01) throw new Error("SITE polygon has zero or negligible area.");
  return out;
}
export function polygonCentroidLocalProjection(points){
  const meanLat=points.reduce((s,p)=>s+p[0],0)/points.length;
  const meanLon=points.reduce((s,p)=>s+p[1],0)/points.length;
  const r=6378137;
  const cosLat=Math.cos(meanLat*Math.PI/180);
  if(Math.abs(cosLat)<1e-12) throw new Error("Polygon is too close to a pole for this centroid projection.");
  const xy=points.map(p=>[(p[1]-meanLon)*Math.PI/180*r*cosLat,(p[0]-meanLat)*Math.PI/180*r]);
  let twiceArea=0,cx6a=0,cy6a=0;
  for(let i=0;i<xy.length;i++){
    const a=xy[i],b=xy[(i+1)%xy.length];
    const cross=a[0]*b[1]-b[0]*a[1];
    twiceArea+=cross; cx6a+=(a[0]+b[0])*cross; cy6a+=(a[1]+b[1])*cross;
  }
  if(Math.abs(twiceArea)<1e-9) throw new Error("Polygon area is zero.");
  const cx=cx6a/(3*twiceArea), cy=cy6a/(3*twiceArea);
  return {latitude:meanLat+(cy/r)*180/Math.PI, longitude:meanLon+(cx/(r*cosLat))*180/Math.PI, areaM2:Math.abs(twiceArea)/2};
}
