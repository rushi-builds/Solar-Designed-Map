import {validateSitePolygon,polygonCentroidLocalProjection} from './worker_algo.mjs';
// "lat,lon;lat,lon;..." exactly as DRAWING_DATA!Coordinates stores it (map saves 8 decimals)
export const CASES = [
 ["sonwadi_site","17.99038747,74.43525012;17.99138747,74.43525012;17.99138747,74.43625012;17.99038747,74.43625012"],
 ["closed_ring_dup","18.18043580,74.61004963;18.18143580,74.61004963;18.18143580,74.61104963;18.18043580,74.61004963"],
 ["consecutive_dup","18.52043000,73.85674000;18.52043000,73.85674000;18.52143000,73.85774000;18.52093000,73.85874000"],
 ["triangle","19.07609000,72.87766000;19.08109123,72.88166456;19.07209789,72.88066123"],
 ["long_12pt","17.45012345,75.12012345;17.45098765,75.12098765;17.45177777,75.12200000;17.45100000,75.12333333;17.45022222,75.12400000;17.44955555,75.12300000;17.44900000,75.12188888;17.44966666,75.12077777;17.45000001,75.12000002;17.45066667,75.12066668;17.45133334,75.12133335;17.45050000,75.12050000"],
 ["southern_hemi","-23.55052000,-46.63330800;-23.55152000,-46.63330800;-23.55152000,-46.63230800;-23.55052000,-46.63230800"],
 ["tiny_1m","18.00000000,74.00000000;18.00000900,74.00000000;18.00000900,74.00000900;18.00000000,74.00000900"]
];
export function runWorker(){
  return CASES.map(([name,coords])=>{
    try{
      const poly=validateSitePolygon(coords.split(";").map(p=>p.split(",").map(Number)));
      const c=polygonCentroidLocalProjection(poly);
      return {name,ok:true,n:poly.length,lat:c.latitude,lon:c.longitude,area:c.areaM2,
              lat8:c.latitude.toFixed(8),lon8:c.longitude.toFixed(8),area2:c.areaM2.toFixed(2),
              vertices:poly.map(p=>p[0].toFixed(8)+","+p[1].toFixed(8)).join(";")};
    }catch(e){return {name,ok:false,error:String(e.message||e)};}
  });
}
if(process.argv[2]==='run'){console.log(JSON.stringify(runWorker(),null,1));}
