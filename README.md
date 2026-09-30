<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard Saber 11.º — I.E. La Esperanza</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<style>
:root{--bg:#f4f6fb;--card:#fff;--tx:#1e293b;--mu:#64748b;--pr:#14532d;--ac:#0f766e;--bd:#e2e8f0}
@media(prefers-color-scheme:dark){:root{--bg:#0f172a;--card:#1e293b;--tx:#e2e8f0;--mu:#94a3b8;--pr:#166534;--bd:#334155}}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--tx)}
header{background:linear-gradient(120deg,#14532d,#0f766e);color:#fff;padding:20px 28px}
header h1{margin:0;font-size:26px}header p{margin:4px 0;opacity:.9}
.chips span{display:inline-block;background:#ffffff26;border-radius:20px;padding:4px 14px;margin:6px 8px 0 0;font-weight:600}
.wrap{display:flex;min-height:calc(100vh - 150px)}
nav{width:230px;background:var(--card);border-right:1px solid var(--bd);padding:12px;flex-shrink:0}
nav button{display:block;width:100%;text-align:left;border:0;background:none;padding:11px 12px;border-radius:8px;font-size:14px;color:var(--tx);cursor:pointer;margin-bottom:2px}
nav button.on,nav button:hover{background:#0f766e22;font-weight:700}
main{flex:1;padding:22px;min-width:0}section{display:none}section.on{display:block}
h2{margin:0 0 14px}h3{margin:0 0 10px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:18px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:16px;box-shadow:0 1px 3px #0001;margin-bottom:16px}
.grid .card{margin:0;border-top:5px solid var(--k,#0f766e)}
.k{font-size:12px;color:var(--mu);font-weight:700;letter-spacing:.04em}.v{font-size:32px;font-weight:800;margin-top:4px}.s{font-size:12px;color:var(--mu)}
.two{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:16px}
.br{display:grid;grid-template-columns:150px 1fr 60px;gap:10px;align-items:center;margin:8px 0;font-size:14px}
.tr{background:#94a3b833;border-radius:8px;height:22px;overflow:hidden}.fi{height:100%;border-radius:8px}
table{width:100%;border-collapse:collapse;font-size:14px}th,td{padding:8px;border-bottom:1px solid var(--bd);text-align:left}
th{color:var(--mu);font-size:12px}th.so{cursor:pointer}.tw{overflow-x:auto}
.pill{display:inline-block;border-radius:12px;padding:2px 10px;color:#fff;font-weight:700;font-size:13px}
.i{display:inline-block;position:relative;width:16px;height:16px;line-height:16px;text-align:center;border-radius:50%;background:var(--mu);color:#fff;font-size:11px;font-style:normal;cursor:help;margin-left:5px}
.i:hover::after{content:attr(data-tip);position:absolute;left:20px;top:-6px;width:230px;background:#0f172a;color:#fff;padding:8px 10px;border-radius:8px;font-size:12px;font-weight:400;letter-spacing:0;z-index:9}
button.b,select,input{font:inherit;padding:8px 12px;border-radius:8px;border:1px solid var(--bd);background:var(--card);color:var(--tx)}
button.b{background:#0f766e;color:#fff;border:0;cursor:pointer;font-weight:700}button.big{font-size:18px;padding:16px 26px}
.tools{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}.pg{display:flex;gap:8px;align-items:center;margin-top:10px}
.ok{border-left:5px solid #16a34a}.wn{border-left:5px solid #f97316}
#modal{display:none;position:fixed;inset:0;background:#0009;align-items:center;justify-content:center;z-index:20}#modal>div{background:var(--card);border-radius:14px;padding:22px;width:min(480px,92vw)}
li{margin-bottom:8px}
@media(max-width:800px){.wrap{flex-direction:column}nav{width:100%;display:flex;overflow-x:auto}nav button{white-space:nowrap;width:auto}.br{grid-template-columns:100px 1fr 46px}}
</style></head><body>
<header><h1>Institución Educativa La Esperanza</h1><p>Resultados Saber 11° — Oficial Institución Educativa La Esperanza</p>
<div class="chips"><span>Grado 11.º</span><span id="hn"></span><span id="hsrc"></span></div></header>
<div class="wrap"><nav id="nav"></nav><main>
<section id="resumen"></section><section id="areas"></section><section id="dist"></section><section id="est"></section>
<section id="fort"></section><section id="evol"></section><section id="concl"></section>
<section id="pdf"><h2>📄 Informe PDF institucional</h2><div class="card"><p>Genera un informe de 8 páginas con los <b>mismos datos, cálculos, gráficos y conclusiones</b> que ves en este dashboard (una sola función de análisis alimenta ambos).</p>
<button class="b big" onclick="makePDF()">📄 GENERAR INFORME PDF</button></div>
<div class="card"><h3>Cargar otro archivo Excel</h3><p class="s">La app ya trae los datos del archivo oficial. Si actualizas el Excel, cárgalo aquí y todo (dashboard y PDF) se recalcula. Funciona sin servidor.</p><input type="file" id="f" accept=".xlsx,.xls"><p id="ferr" class="s"></p></div></section>
</main></div>
<div id="modal" onclick="this.style.display='none'"><div id="mb" onclick="event.stopPropagation()"></div></div>
<script>
const AREAS=['Lectura Crítica','Matemáticas','Sociales y Ciudadanas','Ciencias Naturales','Inglés'];
const SH=['Lectura','Matemáticas','Sociales','Naturales','Inglés'];
const EMB=[["BERROCAL ARRIETA LUIS MATEO",74,80,67,100,64,395],["APARICIO MELENDREZ ANGELA MARÍA",59,69,53,68,45,304.615385],["RUIZ RODRIGUEZ VALENTINA",59,62,58,65,51,301.153846],["MARIMON TUIRAN SARA ENA",51,66,55,64,46,290],["ARRIETA GONZALEZ LORENA",55,65,53,59,51,287.307692],["MENDEZ SIERRA EIDY LUZ",52,65,37,67,45,272.307692],["JARAMILLO CAMAÑO JUAN PABLO",51,51,47,57,45,255],["GASPAR SIERRA EMILY VALENTINA",55,44,44,50,40,238.076923],["RUIZ ESTRADA JUAN DAVID",43,43,55,40,40,224.230769],["ORTIZ MADERA CAMILO ANDRES",41,48,41,48,40,220.769231],["MORENO SANCHEZ EMANUEL",34,53,40,47,29,211.923077]];
/* Referencia histórica: informe PDF "Saber Once 2025" (I.E. La Esperanza, 20 evaluados en 2025) */
const HG={2016:239,2017:226,2018:226,2019:229,2020:212,2021:214,2022:229,2023:215,2024:248,2025:227};
const H25=[48,46,43,47,41]; // Lectura, Mat, Soc, Nat, Ing (promedios 2025 del informe)
const COL=['#dc2626','#eab308','#f97316','#2563eb','#16a34a'];
const gc=v=>v<200?COL[0]:v<250?COL[1]:v<300?COL[2]:v<400?COL[3]:COL[4];
const ac=v=>v<40?COL[0]:v<60?COL[1]:COL[4];
const rnd=x=>{if(x<0)return -rnd(-x);const i=Math.floor(x+1e-9);return x-i>0.5+1e-9?i+1:i};
const f1=x=>String(rnd(x)),pc=(a,n)=>rnd(100*a/n)+' %';
const $=id=>document.getElementById(id);
const tip=t=>`<i class="i" data-tip="${t}">i</i>`;
let M,SRC='Archivo oficial integrado';

/* ============ ANÁLISIS ÚNICO (Dashboard y PDF) ============ */
function pct(s,p){const i=(s.length-1)*p,l=Math.floor(i);return s[l]+(s[Math.min(l+1,s.length-1)]-s[l])*(i-l)}
function stats(a){const n=a.length,s=[...a].sort((x,y)=>x-y),m=a.reduce((x,y)=>x+y,0)/n;
 const sd=n>1?Math.sqrt(a.reduce((x,y)=>x+(y-m)**2,0)/(n-1)):0;
 return{n,mean:m,sd,min:s[0],max:s[n-1],range:s[n-1]-s[0],med:pct(s,.5),p10:pct(s,.1),p25:pct(s,.25),p50:pct(s,.5),p75:pct(s,.75),p90:pct(s,.9)}}
function analyze(raw,total,missing){
 const R=raw,n=R.length,names=new Set(),dup=R.length-new Set(R.map(r=>r.name.trim().toUpperCase())).size;
 const g=stats(R.map(r=>r.g)),ar=AREAS.map((nm,i)=>{const v=R.map(r=>r.a[i]),s=stats(v);
  return{nm,sh:SH[i],...s,lo:v.filter(x=>x<40).length,mid:v.filter(x=>x>=40&&x<60).length,hi:v.filter(x=>x>=60).length}});
 const bands=[['<200',x=>x<200,COL[0]],['200–249',x=>x>=200&&x<250,COL[1]],['250–299',x=>x>=250&&x<300,COL[2]],['300–349',x=>x>=300&&x<350,COL[3]],['≥350',x=>x>=350,COL[4]]]
  .map(([l,f,c])=>({l,c,f,n:R.filter(r=>f(r.g)).length}));
 const top=(fn,l,ic)=>{const v=R.map(fn),mx=Math.max(...v);return{l,ic,val:mx,who:R.filter((r,i)=>v[i]===mx).map(r=>r.name)}};
 const tops=[top(r=>r.g,'Mayor puntaje global','🏆'),...[0,1,2,3,4].map(i=>top(r=>r.a[i],'Mayor '+SH[i],['📚','🧮','🌎','🔬','🇬🇧'][i]))];
 const rk=[...ar].sort((a,b)=>b.mean-a.mean),best=rk[0],worst=rk[rk.length-1];
 const n300=R.filter(r=>r.g>=300).length,n250=R.filter(r=>r.g<250).length;
 const bigB=[...bands].sort((a,b)=>b.n-a.n)[0];
 const gs=R.map(r=>r.g).sort((a,b)=>b-a);
 const find=[`El grupo de grado 11.º obtuvo un promedio global de ${f1(g.mean)} puntos. La mediana fue de ${f1(g.med)} puntos y los resultados se ubicaron entre ${rnd(g.min)} y ${rnd(g.max)} puntos.`,
  `El área con mayor promedio fue ${best.nm} (${f1(best.mean)}), mientras que ${worst.nm} presentó el menor promedio (${f1(worst.mean)}).`,
  `${n300} estudiantes alcanzaron 300 puntos o más y ${n250} estudiantes se ubicaron por debajo de 250 puntos.`];
 const mxl=Math.max(...ar.map(a=>a.lo)),lowA={lo:mxl,nm:ar.filter(a=>a.lo===mxl).map(a=>a.nm).join(', ')};
 const d=g.mean-g.med;
 const hy=Math.max(...Object.keys(HG));
 const concl=[
  `El promedio global es ${f1(g.mean)} y la mediana ${f1(g.med)}; ${Math.abs(d)<g.sd*.15?'ambos valores son cercanos, lo que indica una distribución central sin sesgo marcado':d>0?'el promedio supera a la mediana, lo que sugiere influencia de puntajes altos':'la mediana supera al promedio, lo que sugiere influencia de puntajes bajos'}.`,
  `La dispersión es de ${f1(g.sd)} puntos (desviación estándar) con un rango de ${f1(g.range)} puntos entre el mínimo (${rnd(g.min)}) y el máximo (${rnd(g.max)}).`,
  `${best.nm} es el área con mayor promedio (${f1(best.mean)}) y ${worst.nm} la de menor promedio (${f1(worst.mean)}); la diferencia es de ${f1(best.mean-worst.mean)} puntos.`,
  `${n300} de ${n} estudiantes (${pc(n300,n)}) alcanzan 300 puntos o más; ${n250} (${pc(n250,n)}) están por debajo de 250.`,
  `El rango de ${bigB.l} concentra la mayor cantidad de estudiantes: ${bigB.n} (${pc(bigB.n,n)}).`,
  `El puntaje global más alto (${rnd(gs[0])}) supera en ${f1(gs[0]-gs[1])} puntos al segundo mejor (${rnd(gs[1])}).`,
  lowA.lo>0?`Áreas con más estudiantes por debajo de 40 puntos: ${lowA.nm} (${lowA.lo} estudiante${lowA.lo>1?'s':''} en cada una, ${pc(lowA.lo,n)}).`:`Ningún estudiante obtuvo menos de 40 puntos en alguna de las cinco áreas.`,
  `Referencia histórica: el promedio global de este archivo (${f1(g.mean)}) frente al promedio 2025 del informe institucional (${HG[hy]}) difiere en ${f1(g.mean-HG[hy])} puntos; se trata de grupos distintos (${n} vs. 20 evaluados), por lo que es una referencia y no una evolución.`];
 return{R,n,g,ar,bands,tops,rk,best,worst,n300,n250,find,concl,q:{total,valid:n,missing,dup},bigB,
  sf:rk.slice(0,2),op:rk.slice(-2).reverse()}}

/* ============ LECTURA DEL EXCEL ============ */
function parseWB(wb){
 const order=[...wb.SheetNames].sort((a,b)=>(/DATOS/i.test(b)?1:0)-(/DATOS/i.test(a)?1:0));
 for(const sn of order){const A=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null});
  const hi=A.findIndex(r=>r.some(c=>/nombre/i.test(c||''))&&r.some(c=>/mat/i.test(c||'')));if(hi<0)continue;
  const H=A[hi].map(c=>String(c||'')),ix=re=>H.findIndex(h=>re.test(h));
  const c={n:ix(/nombre/i),a:[ix(/lec/i),ix(/mat/i),ix(/soc/i),ix(/nat/i),ix(/ingl/i)],g:ix(/^\s*puntaje(\s*global.*)?$/i)>=0?ix(/^\s*puntaje(\s*global.*)?$/i):ix(/puntaje.*global/i)};
  if(c.a.includes(-1)||c.g<0)continue;
  let total=0,miss=0;const R=[];
  for(const r of A.slice(hi+1)){const nm=r[c.n];if(!nm||/promedio|desviaci/i.test(nm))continue;total++;
   const a=c.a.map(i=>parseFloat(r[i])),g=parseFloat(r[c.g]);
   if(a.some(isNaN)||isNaN(g)){miss++;continue}R.push({name:String(nm).trim(),a,g})}
  if(R.length)return{R,total,miss,sn}}
 throw Error('No se encontró una tabla con nombres, cinco áreas y puntaje global.')}

/* ============ RENDER ============ */
const bar=(l,v,max,c,txt)=>`<div class="br"><span>${l}</span><div class="tr"><div class="fi" style="width:${Math.max(0,v/max*100)}%;background:${c}"></div></div><b>${txt??f1(v)}</b></div>`;
function render(){
 const m=M;$('hn').textContent=m.n+' estudiantes evaluados';$('hsrc').textContent='Fuente: '+SRC;
 const N=[['resumen','🏠 RESUMEN'],['areas','📊 RESULTADOS POR ÁREA'],['dist','📈 DISTRIBUCIÓN'],['est','👥 ESTUDIANTES'],['fort','🎯 FORTALEZAS Y FORTALECIMIENTO'],['evol','📈 REFERENCIA HISTÓRICA'],['concl','🧾 CONCLUSIONES Y METODOLOGÍA'],['pdf','📄 INFORME PDF']];
 $('nav').innerHTML=N.map(([k,t])=>`<button data-k="${k}" onclick="go('${k}')">${t}</button>`).join('');
 const g=m.g,K=(k,v,s,c,t)=>`<div class="card" style="--k:${c||'#0f766e'}"><div class="k">${k}${t?tip(t):''}</div><div class="v">${v}</div><div class="s">${s||''}</div></div>`;
 $('resumen').innerHTML=`<h2>Resumen ejecutivo</h2><div class="grid">
 ${K('👥 ESTUDIANTES',m.n,'evaluados')}${K('🎯 PROMEDIO GLOBAL',f1(g.mean),'puntos',gc(g.mean),'Suma de los puntajes globales dividida entre el número de estudiantes.')}
 ${K('📊 MEDIANA',f1(g.med),'puntos',gc(g.med),'Valor central cuando los resultados se ordenan de menor a mayor: divide al grupo en dos mitades.')}
 ${K('📐 DESV. ESTÁNDAR',f1(g.sd),'puntos',null,'Indica qué tan dispersos están los resultados alrededor del promedio del grupo (versión muestral).')}
 ${K('⬆️ MÁXIMO',rnd(g.max),'puntos',gc(g.max))}${K('⬇️ MÍNIMO',rnd(g.min),'puntos',gc(g.min))}
 ${K('🟢 ≥300',m.n300,pc(m.n300,m.n)+' del grupo',COL[3],'Estudiantes con 300 puntos o más en el puntaje global.')}${K('🟠 <250',m.n250,pc(m.n250,m.n)+' del grupo',COL[1],'Estudiantes con menos de 250 puntos en el puntaje global.')}</div>
 <div class="card"><h3>Hallazgos principales</h3>${m.find.map(t=>`<p>${t}</p>`).join('')}</div>
 <div class="card ${m.q.missing||m.q.dup?'wn':'ok'}"><h3>Calidad de los datos</h3><div class="grid"><div><b>${m.n}</b><br><span class="s">Estudiantes analizados</span></div><div><b>${m.q.valid}</b><br><span class="s">Registros válidos (de ${m.q.total})</span></div><div><b>${m.q.missing}</b><br><span class="s">Con datos faltantes</span></div><div><b>${m.q.dup}</b><br><span class="s">Duplicados detectados</span></div></div></div>`;
 $('areas').innerHTML=`<h2>Resultados por área</h2><div class="two"><div class="card"><h3>Promedio por área (0–100)</h3>${m.ar.map(a=>bar(a.nm,a.mean,100,ac(a.mean))).join('')}</div>
 <div class="card"><h3>Tabla comparativa</h3><div class="tw"><table><tr><th>Área</th><th>Promedio</th><th>Mín.</th><th>Máx.</th></tr>${m.ar.map(a=>`<tr><td>${a.nm}</td><td><b>${f1(a.mean)}</b></td><td>${a.min}</td><td>${a.max}</td></tr>`).join('')}</table></div></div></div>
 <div class="card"><h3>Distribución por rangos de puntaje de área ${tip('Cada estudiante se cuenta según su puntaje en el área: rojo (<40), amarillo (40–59), verde (≥60).')}</h3><div class="tw"><table><tr><th>Área</th><th>🔴 &lt;40</th><th>🟡 40–59</th><th>🟢 ≥60</th></tr>${m.ar.map(a=>`<tr><td>${a.nm}</td><td>${a.lo} · ${pc(a.lo,m.n)}</td><td>${a.mid} · ${pc(a.mid,m.n)}</td><td>${a.hi} · ${pc(a.hi,m.n)}</td></tr>`).join('')}</table></div></div>`;
 const mx=Math.max(...m.bands.map(b=>b.n),1);
 $('dist').innerHTML=`<h2>¿Dónde se concentra el grupo?</h2><div class="card">${m.bands.map(b=>bar(b.l,b.n,mx,b.c,`${b.n} · ${pc(b.n,m.n)}`)).join('')}<p><b>Lectura:</b> el rango <b>${m.bigB.l}</b> reúne la mayor cantidad de estudiantes (${m.bigB.n}, ${pc(m.bigB.n,m.n)}); ${m.n300} estudiantes tienen 300 puntos o más y ${m.n250} tienen menos de 250.</p></div>
 <h2>Perfil estadístico del grupo</h2><div class="grid">${[['Promedio',g.mean,'Suma de los puntajes dividida entre el número de estudiantes.'],['Mediana',g.med,'Valor central cuando los resultados se ordenan de menor a mayor.'],['Desv. estándar',g.sd,'Qué tan dispersos están los resultados alrededor del promedio.'],['Mínimo',g.min,'Menor puntaje global.'],['Máximo',g.max,'Mayor puntaje global.'],['Rango',g.range,'Diferencia entre el máximo y el mínimo.'],['P10',g.p10,'Aproximadamente el 10 % de los estudiantes está en o por debajo de este valor.'],['P25',g.p25,'Aproximadamente el 25 % de los estudiantes está en o por debajo de este valor.'],['P50',g.p50,'Aproximadamente el 50 % está en o por debajo de este valor (coincide con la mediana).'],['P75',g.p75,'Aproximadamente el 75 % está en o por debajo de este valor.'],['P90',g.p90,'Aproximadamente el 90 % está en o por debajo de este valor.']].map(([k,v,t])=>K(k,f1(v),'',gc(v),t).replace(/style="--k:[^"]*"/,'')).join('')}</div>`;
 $('est').innerHTML=`<h2>Resultados individuales</h2><div class="card"><div class="tools"><input id="q" placeholder="🔍 Buscar estudiante…" oninput="pgN=0;tbl()"><select id="fl" onchange="pgN=0;tbl()"><option value="">Todos los puntajes</option>${m.bands.map((b,i)=>`<option value="${i}">Global ${b.l}</option>`).join('')}</select></div><div class="tw" id="tb"></div><div class="pg" id="pg"></div></div>
 <h2>Mayores resultados</h2><div class="grid">${m.tops.map(t=>`<div class="card" style="--k:#eab308"><div class="k">${t.ic} ${t.l.toUpperCase()}</div><div class="v">${f1(t.val).replace(',0','')}</div><div class="s">${t.who.join('<br>')}</div></div>`).join('')}</div>`;
 sortK=6;sortD=-1;pgN=0;tbl();
 const li=a=>a.map(x=>`<li><b>${x.nm}</b>: promedio ${f1(x.mean)}; ${x.hi} estudiantes con ≥60 y ${x.lo} con &lt;40.</li>`).join('');
 $('fort').innerHTML=`<h2>Análisis académico</h2><div class="two"><div class="card ok"><h3>💪 Fortalezas observadas</h3><ul>${li(m.sf)}</ul></div><div class="card wn"><h3>🌱 Oportunidades de fortalecimiento</h3><ul>${li(m.op)}</ul></div></div><div class="card"><h3>Ranking de áreas</h3>${m.rk.map(a=>bar(a.nm,a.mean,100,ac(a.mean))).join('')}</div>
 <div class="card s">No se muestra la sección de competencias: el archivo no contiene información desagregada por competencias.</div>`;
 const ys=Object.keys(HG),pts=[...ys.map(y=>[y,HG[y]]),['2026',+g.mean.toFixed(1)]],hm=Math.max(...pts.map(p=>p[1]))+15;
 $('evol').innerHTML=`<h2>Referencia histórica institucional</h2><div class="card wn"><p>El Excel no contiene simulacros ni varias aplicaciones, por eso <b>no hay evolución dentro del archivo</b>. Como referencia se muestran los promedios globales 2016–2025 del <i>Informe de prueba Saber Once 2025</i> (I.E. La Esperanza) junto al promedio del grupo actual. Son cohortes distintas: no equivale a una evolución de un mismo grupo.</p></div>
 <div class="card"><h3>Promedio global por año</h3>${pts.map(([y,v])=>bar(y,v,hm,y==='2026'?'#0f766e':gc(v),f1(v))).join('')}</div>
 <div class="card"><h3>Promedio por área: 2025 (informe) vs. grupo actual</h3><div class="tw"><table><tr><th>Área</th><th>2025</th><th>Actual</th><th>Diferencia</th></tr>${m.ar.map((a,i)=>`<tr><td>${a.nm}</td><td>${H25[i]}</td><td>${f1(a.mean)}</td><td>${(a.mean-H25[i]>=0?'+':'')+f1(a.mean-H25[i])}</td></tr>`).join('')}</table></div></div>`;
 $('concl').innerHTML=`<h2>Conclusiones</h2><div class="card"><ol>${m.concl.map(c=>`<li>${c}</li>`).join('')}</ol></div><h2>Metodología</h2><div class="card">${metod().map(t=>`<p>${t}</p>`).join('')}</div>`;
 go(sessionStorageSafe()||'resumen');
}
const metod=()=>[`<b>Fuente:</b> ${SRC}. Estudiantes analizados: ${M.n} de ${M.q.total} registros (${M.q.missing} con datos faltantes excluidos; ${M.q.dup} duplicados).`,
 `<b>Áreas:</b> ${AREAS.join(', ')}. El puntaje global es el valor original del archivo (no se recalcula).`,
 `<b>Estadística:</b> promedio aritmético, mediana, desviación estándar muestral (n−1) y percentiles por interpolación lineal (equivalente a PERCENTILE.INC de Excel). Se usan siempre los valores originales.`,
 `<b>Redondeo:</b> todos los valores mostrados (puntajes, promedios, desviación, percentiles, diferencias y porcentajes) se redondean al entero: decimal &gt; 0,5 sube al entero siguiente; ≤ 0,5 se conserva el entero inferior (57,6→58; 57,5→57). Es solo visual: los cálculos usan los valores originales, que no se alteran.`,
 `<b>Competencias:</b> no disponibles en el archivo. <b>Histórico/simulacros:</b> no disponibles en el archivo; la referencia 2016–2025 proviene del informe PDF institucional 2025.`];
function sessionStorageSafe(){return null}
function go(k){document.querySelectorAll('section').forEach(s=>s.classList.toggle('on',s.id===k));document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.k===k));scrollTo(0,0)}

/* ============ TABLA ============ */
let sortK=6,sortD=-1,pgN=0;const PS=10;
function tbl(){const q=($('q').value||'').toUpperCase(),fl=$('fl').value;
 let L=M.R.filter(r=>r.name.toUpperCase().includes(q)&&(fl===''||M.bands[+fl].f(r.g)));
 const val=r=>sortK===0?r.name:sortK===6?r.g:r.a[sortK-1];
 L.sort((a,b)=>(val(a)>val(b)?1:-1)*sortD);const pages=Math.max(1,Math.ceil(L.length/PS));pgN=Math.min(pgN,pages-1);
 const hd=['Estudiante',...SH,'Global'];
 $('tb').innerHTML=`<table><tr>${hd.map((h,i)=>`<th class="so" onclick="sortK=${i};sortD=sortK==${i}&&sortD==-1?1:-1;tbl()">${h}${i==sortK?(sortD<0?' ▼':' ▲'):''}</th>`).join('')}</tr>${L.slice(pgN*PS,pgN*PS+PS).map(r=>`<tr style="cursor:pointer" onclick="modal(${M.R.indexOf(r)})"><td>${r.name}</td>${r.a.map(x=>`<td><span style="color:${ac(x)};font-weight:700">${rnd(x)}</span></td>`).join('')}<td><span class="pill" style="background:${gc(r.g)}">${rnd(r.g)}</span></td></tr>`).join('')}</table>`;
 $('pg').innerHTML=`<button class="b" onclick="pgN=Math.max(0,pgN-1);tbl()">‹</button> Página ${pgN+1} de ${pages} · ${L.length} estudiantes <button class="b" onclick="pgN=Math.min(${pages-1},pgN+1);tbl()">›</button>`}
function modal(i){const r=M.R[i];$('mb').innerHTML=`<h3>${r.name}</h3><p>Global: <span class="pill" style="background:${gc(r.g)}">${rnd(r.g)}</span> · Posición ${[...M.R].sort((a,b)=>b.g-a.g).indexOf(r)+1} de ${M.n}</p>${r.a.map((x,k)=>bar(SH[k],x,100,ac(x),rnd(x))).join('')}<p class="s">Barra: puntaje del área. Línea de referencia: promedio del grupo ${M.ar.map((a,k)=>SH[k]+' '+f1(a.mean)).join(' · ')}</p><button class="b" onclick="$('modal').style.display='none'">Cerrar</button>`;$('modal').style.display='flex'}

/* ============ PDF (usa M, el mismo modelo) ============ */
const rgb=h=>[1,3,5].map(i=>parseInt(h.substr(i,2),16));
const T=s=>String(s).replace(/≥/g,'>=').replace(/≤/g,'<=').replace(/–/g,'-').replace(/[^\x00-\xFF]/g,'');
function makePDF(){
 if(!window.jspdf){alert('No se pudo cargar la librería PDF (se requiere conexión la primera vez).');return}
 const D=new jspdf.jsPDF({unit:'mm',format:'a4'}),m=M,g=m.g;let pgs=0;
 const col=(c,t='f')=>t==='f'?D.setFillColor(...rgb(c)):D.setTextColor(...rgb(c));
 const tx=(s,x,y,sz=10,b=false,c='#1e293b',o={})=>{D.setFont('helvetica',b?'bold':'normal');D.setFontSize(sz);col(c,'t');D.text(T(s),x,y,o)};
 const para=(s,x,y,w,sz=10)=>{D.setFont('helvetica','normal');D.setFontSize(sz);col('#1e293b','t');const L=D.splitTextToSize(T(s),w);D.text(L,x,y);return y+L.length*sz*.42+2};
 const page=(t,sub)=>{if(pgs++)D.addPage();col('#14532d');D.rect(0,0,210,22,'F');tx('Institución Educativa La Esperanza',12,10,15,true,'#fff');tx('Resultados Saber 11° — Grado 11.º · '+m.n+' estudiantes evaluados',12,17,9,false,'#fff');
  tx(t,12,34,17,true,'#14532d');if(sub)tx(sub,12,40,9,false,'#64748b');tx('Fuente: '+SRC+' · Página '+pgs,12,290,8,false,'#64748b')};
 const bars=(it,x,y,max,w=110)=>{it.forEach((b,i)=>{const yy=y+i*9;tx(b.l,x,yy+5,9);col('#e2e8f0');D.roundedRect(x+42,yy,w,6,1,1,'F');col(b.c);if(b.v>0)D.roundedRect(x+42,yy,Math.max(1,b.v/max*w),6,1,1,'F');tx(b.t,x+42+w+3,yy+5,9,true)});return y+it.length*9+3};
 const tab=(hd,rows,x,y,cw)=>{let cx=x;col('#e2e8f0');D.rect(x,y-5,cw.reduce((a,b)=>a+b,0),7,'F');hd.forEach((h,i)=>{tx(h,cx+1,y,8.5,true);cx+=cw[i]});y+=7;
  rows.forEach(r=>{cx=x;r.forEach((c,i)=>{tx(c,cx+1,y,8.5);cx+=cw[i]});y+=6.5});return y};
 // P1
 page('Resumen ejecutivo','Informe institucional de resultados Saber 11.º');
 const kp=[['ESTUDIANTES',m.n,'#0f766e'],['PROMEDIO GLOBAL',f1(g.mean),gc(g.mean)],['MEDIANA',f1(g.med),gc(g.med)],['DESV. ESTANDAR',f1(g.sd),'#64748b'],['MAXIMO',rnd(g.max),gc(g.max)],['MINIMO',rnd(g.min),gc(g.min)],['>=300 PUNTOS',m.n300,COL[3]],['<250 PUNTOS',m.n250,COL[1]]];
 kp.forEach(([k,v,c],i)=>{const x=12+(i%4)*47,y=48+Math.floor(i/4)*30;col('#f1f5f9');D.roundedRect(x,y,44,26,2,2,'F');col(c);D.rect(x,y,44,2,'F');tx(k,x+3,y+8,7.5,true,'#64748b');tx(String(v),x+3,y+20,20,true,c)});
 let y=115;tx('Hallazgos principales',12,y,13,true,'#14532d');y+=7;m.find.forEach(t=>y=para(t,12,y,186,10.5)+2);
 y+=4;tx('Calidad de los datos',12,y,12,true,'#14532d');y=para(`Estudiantes analizados: ${m.n} · Registros validos: ${m.q.valid} de ${m.q.total} · Datos faltantes: ${m.q.missing} · Duplicados detectados: ${m.q.dup}.`,12,y+6,186);
 // P2
 page('Resultados por área','Promedio de cada área (escala 0–100)');
 y=bars(m.ar.map(a=>({l:a.nm,v:a.mean,c:ac(a.mean),t:f1(a.mean)})),12,48,100,90);
 y=tab(['Área','Promedio','Mín.','Máx.','<40','40-59','>=60'],m.ar.map(a=>[a.nm,f1(a.mean),a.min,a.max,`${a.lo} (${pc(a.lo,m.n)})`,`${a.mid} (${pc(a.mid,m.n)})`,`${a.hi} (${pc(a.hi,m.n)})`]),12,y+8,[52,22,14,14,28,28,28]);
 para(`${m.best.nm} tiene el mayor promedio (${f1(m.best.mean)}) y ${m.worst.nm} el menor (${f1(m.worst.mean)}).`,12,y+6,186);
 // P3
 page('Distribución global','¿Dónde se concentra el grupo?');const mx=Math.max(...m.bands.map(b=>b.n),1);
 y=bars(m.bands.map(b=>({l:b.l,v:b.n,c:b.c,t:`${b.n} · ${pc(b.n,m.n)}`})),12,48,mx,90);
 para(`El rango ${m.bigB.l} reúne la mayor cantidad de estudiantes (${m.bigB.n}, ${pc(m.bigB.n,m.n)}). ${m.n300} estudiantes tienen 300 puntos o más y ${m.n250} tienen menos de 250.`,12,y+6,186,10.5);
 tx('Escala institucional de colores: rojo <200 · amarillo 200-249 · naranja 250-299 · azul 300-399 · verde >=400',12,y+26,8.5,false,'#64748b');
 // P4
 page('Percentiles y dispersión','Perfil estadístico del puntaje global');
 y=tab(['Indicador','Valor','Significado'],[['Promedio',g.mean,'Suma de puntajes / número de estudiantes'],['Mediana',g.med,'Valor central: divide al grupo en dos mitades'],['Desv. estándar',g.sd,'Dispersión alrededor del promedio'],['Mínimo',g.min,'Menor puntaje'],['Máximo',g.max,'Mayor puntaje'],['Rango',g.range,'Máximo menos mínimo'],['P10',g.p10,'~10 % del grupo en o por debajo'],['P25',g.p25,'~25 % del grupo en o por debajo'],['P50',g.p50,'~50 % del grupo en o por debajo'],['P75',g.p75,'~75 % del grupo en o por debajo'],['P90',g.p90,'~90 % del grupo en o por debajo']].map(r=>[r[0],f1(r[1]),r[2]]),12,50,[38,28,120]);
 para('Percentiles calculados por interpolación lineal sobre los puntajes originales del archivo.',12,y+6,186,9);
 // P5
 page('Estudiantes destacados','Mayores resultados globales y por área');
 y=tab(['Categoría','Puntaje','Estudiante(s)'],m.tops.map(t=>[t.l,f1(t.val).replace(',0',''),t.who.join('; ')]),12,50,[46,22,118]);
 y+=6;tx('Resultados individuales (orden por puntaje global)',12,y,11,true,'#14532d');
 tab(['Estudiante',...SH,'Global'],[...m.R].sort((a,b)=>b.g-a.g).map(r=>[r.name.slice(0,30),...r.a.map(rnd),rnd(r.g)]),12,y+8,[62,20,24,20,20,18,22]);
 // P6
 page('Fortalezas y oportunidades de fortalecimiento','Análisis descriptivo basado en promedios por área');
 y=48;tx('Fortalezas observadas',12,y,12,true,'#16a34a');y+=7;m.sf.forEach(a=>y=para(`${a.nm}: promedio ${f1(a.mean)}; ${a.hi} estudiantes con >=60 y ${a.lo} con <40.`,12,y,186)+1);
 y+=6;tx('Oportunidades de fortalecimiento',12,y,12,true,'#f97316');y+=7;m.op.forEach(a=>y=para(`${a.nm}: promedio ${f1(a.mean)}; ${a.hi} estudiantes con >=60 y ${a.lo} con <40.`,12,y,186)+1);
 y+=6;tx('Competencias',12,y,12,true,'#14532d');para('El archivo no contiene información desagregada por competencias; esta sección se omite.',12,y+6,186);
 // P7 histórico
 page('Referencia histórica','Informe Saber Once 2025 (cohortes distintas; no es evolución de un mismo grupo)');
 const ys=Object.keys(HG),pts=[...ys.map(k=>[k,HG[k]]),['2026',+g.mean.toFixed(1)]],hm=Math.max(...pts.map(p=>p[1]))+15;
 y=bars(pts.map(([k,v])=>({l:k,v,c:k==='2026'?'#0f766e':gc(v),t:f1(v)})),12,48,hm,100);
 tab(['Área','2025','Actual','Diferencia'],m.ar.map((a,i)=>[a.nm,H25[i],f1(a.mean),(a.mean-H25[i]>=0?'+':'')+f1(a.mean-H25[i])]),12,y+10,[60,26,26,30]);
 // P8
 page('Conclusiones y metodología');y=48;m.concl.forEach((c,i)=>y=para(`${i+1}. ${c}`,12,y,186,10)+1);
 y+=4;tx('Metodología',12,y,12,true,'#14532d');y+=6;metod().forEach(t=>y=para(t.replace(/<[^>]+>/g,'').replace(/&gt;/g,'>'),12,y,186,9)+1);
 D.save('Informe_Saber11_La_Esperanza.pdf')}

/* ============ ARRANQUE ============ */
function load(R,total,miss){M=analyze(R,total,miss);render()}
load(EMB.map(r=>({name:r[0],a:r.slice(1,6),g:r[6]})),EMB.length,0);
$('f').onchange=e=>{const fl=e.target.files[0];if(!fl)return;const rd=new FileReader();rd.onload=ev=>{try{const p=parseWB(XLSX.read(ev.target.result,{type:'array'}));SRC=fl.name+' (hoja '+p.sn+')';load(p.R,p.total,p.miss);go('pdf');$('ferr').textContent='✔ Archivo cargado: '+p.R.length+' estudiantes.'}catch(er){$('ferr').textContent='⚠ '+er.message}};rd.readAsArrayBuffer(fl)};
</script></body></html>
