const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const state={workbook:null,file:null,sheets:{},sheetName:'',rawRows:[],rows:[],headers:[],indicatorCols:[],filtered:[],page:1,pageSize:25,search:'',filterCol:'',filterVal:'',groupCol:''};
const YES=new Set(['si','sí','s','x','1','true','yes','y','ok','✓','✔','marcado','cumple']);
const NO=new Set(['no','n','0','false','-','✗','✕','no cumple']);
const norm=v=>String(v??'').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
const isYes=v=>YES.has(norm(v)); const isNo=v=>NO.has(norm(v));
const fmt=n=>new Intl.NumberFormat('es-CR').format(n||0);
function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2500)}
function safeName(s){return String(s||'archivo').replace(/[^a-z0-9áéíóúüñ _-]/gi,'').trim().replace(/\s+/g,'_').slice(0,80)}
function uniqueHeaders(arr){const used={};return arr.map((h,i)=>{let x=String(h??'').trim()||`Columna ${i+1}`;used[x]=(used[x]||0)+1;return used[x]>1?`${x} (${used[x]})`:x})}
function detectHeaderRow(matrix){let best=0,bestScore=-1;for(let r=0;r<Math.min(matrix.length,25);r++){const row=matrix[r]||[];const vals=row.map(v=>String(v??'').trim()).filter(Boolean);if(vals.length<2)continue;const uniq=new Set(vals.map(v=>norm(v))).size;const alpha=vals.filter(v=>/[a-záéíóúñ]/i.test(v)).length;const score=vals.length*2+uniq+alpha*1.5-r*.25;if(score>bestScore){bestScore=score;best=r}}return best}
function parseSheet(ws){const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,blankrows:false});if(!matrix.length)return {headers:[],rows:[]};const hr=detectHeaderRow(matrix);const headers=uniqueHeaders(matrix[hr]||[]);const rows=matrix.slice(hr+1).filter(r=>r.some(v=>String(v??'').trim()!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));return {headers,rows,headerRow:hr+1}}
function detectIndicators(rows,headers){return headers.filter(h=>{let yes=0,no=0,other=0,blank=0;for(const r of rows){const v=r[h];if(String(v??'').trim()===''){blank++;continue} if(isYes(v))yes++;else if(isNo(v))no++;else other++}const nonblank=yes+no+other;if(!nonblank||yes===0)return false;const binaryRatio=(yes+no)/nonblank;const sparseBinary=(yes+no)/(rows.length||1);return binaryRatio>=.78 && sparseBinary>=.08 && other<=Math.max(2,nonblank*.22)})}
function normalizeRows(rows,headers,indicators){return rows.map((r,idx)=>{const o={__row:idx+1};for(const h of headers){let v=String(r[h]??'').trim();if(indicators.includes(h)){o[h]=isYes(v)?'Sí':(isNo(v)||v===''?'No':v)}else{o[h]=v===''?'Sin dato':v} }return o})}
function categoricalColumns(rows,headers){return headers.filter(h=>{const vals=new Set(rows.map(r=>r[h]).filter(v=>v&&v!=='Sin dato'));return vals.size>=2&&vals.size<=60})}
function processCurrentSheet(){const p=state.sheets[state.sheetName];state.rawRows=p.rows;state.headers=p.headers;state.indicatorCols=detectIndicators(p.rows,p.headers);state.rows=normalizeRows(p.rows,p.headers,state.indicatorCols);state.page=1;state.search='';state.filterCol='';state.filterVal='';$('#searchInput').value='';setupSelectors();applyFilters();updateMeta()}
function setupSelectors(){const cats=categoricalColumns(state.rows,state.headers);const group=$('#groupSelect'), fc=$('#filterColumnSelect');group.innerHTML='';fc.innerHTML='<option value="">Sin filtro específico</option>';const preferred=cats.find(h=>/region|deleg|unidad|oper|tipo|grupo|provincia|canton|cantón|distrito|responsable/i.test(h))||cats[0]||'';cats.forEach(h=>{group.add(new Option(h,h));fc.add(new Option(h,h))});state.groupCol=preferred;if(preferred)group.value=preferred;updateFilterValues()}
function updateFilterValues(){const s=$('#filterValueSelect');s.innerHTML='<option value="">Todos</option>';state.filterCol=$('#filterColumnSelect').value;state.filterVal='';if(!state.filterCol){s.disabled=true;return}s.disabled=false;[...new Set(state.rows.map(r=>r[state.filterCol]))].sort((a,b)=>String(a).localeCompare(String(b),'es')).forEach(v=>s.add(new Option(v,v)))}
function applyFilters(){const q=norm($('#searchInput').value);state.search=q;state.filterCol=$('#filterColumnSelect').value;state.filterVal=$('#filterValueSelect').value;state.groupCol=$('#groupSelect').value;state.filtered=state.rows.filter(r=>{if(state.filterCol&&state.filterVal&&r[state.filterCol]!==state.filterVal)return false;if(q){return state.headers.some(h=>norm(r[h]).includes(q))}return true});state.page=1;renderAll()}
function counts(rows=state.filtered){let y=0,n=0;for(const r of rows)for(const h of state.indicatorCols){if(r[h]==='Sí')y++;else if(r[h]==='No')n++}return {y,n,total:y+n,pct:(y+n)?y/(y+n)*100:0}}
function updateMeta(){const p=state.sheets[state.sheetName];$('#fileName').textContent=state.file.name;$('#fileMeta').textContent=`Hoja: ${state.sheetName} · encabezados detectados en fila ${p.headerRow} · ${fmt(state.rows.length)} registros · ${fmt(state.headers.length)} columnas`;$('#generatedAt').textContent=`Generado: ${new Date().toLocaleString('es-CR')}`}
function renderKPIs(){const c=counts();$('#kpiRows').textContent=fmt(state.filtered.length);$('#kpiRowsSub').textContent=`de ${fmt(state.rows.length)}`;$('#kpiYes').textContent=fmt(c.y);$('#kpiYesSub').textContent='indicadores positivos';$('#kpiNo').textContent=fmt(c.n);$('#kpiPct').textContent=`${c.pct.toFixed(1)}%`;$('#kpiIndicatorSub').textContent=`${fmt(state.indicatorCols.length)} indicadores detectados`}
function renderInsights(){const results=state.indicatorCols.map(h=>{let y=0,n=0;state.filtered.forEach(r=>{if(r[h]==='Sí')y++;else if(r[h]==='No')n++});return {h,y,n,p:(y+n)?y/(y+n)*100:0}}).sort((a,b)=>b.p-a.p);if(!results.length){$('#insightTitle').textContent='No se detectaron columnas Sí/No';$('#insightText').textContent='La tabla se cargó correctamente, pero ninguna columna cumple todavía el patrón de indicador binario.'}else{const best=results[0],worst=results.at(-1);$('#insightTitle').textContent=`Mayor cumplimiento: ${best.h}`;$('#insightText').textContent=`${best.p.toFixed(1)}% de respuestas “Sí”. El indicador con menor resultado es “${worst.h}” con ${worst.p.toFixed(1)}%.`}
let filled=0,total=state.filtered.length*state.headers.length;state.filtered.forEach(r=>state.headers.forEach(h=>{if(r[h]!=='Sin dato')filled++}));const qp=total?filled/total*100:0;$('#qualityPct').textContent=`${qp.toFixed(0)}%`;$('#qualityBar').style.width=`${qp}%`;$('#qualityText').textContent=qp>=95?'Muy buena integridad de datos.':qp>=80?'Integridad adecuada; existen algunos campos sin dato.':'Conviene revisar campos descriptivos incompletos.'}
function plotLayout(extra={}){return {margin:{l:55,r:20,t:18,b:48},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{family:'Inter, system-ui',color:getComputedStyle(document.body).getPropertyValue('--text').trim(),size:11},showlegend:true,legend:{orientation:'h',y:-.13},...extra}}
const config={displayModeBar:false,responsive:true,locale:'es'};
function renderCharts(){}

function normKey(v){
  return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .trim().toUpperCase().replace(/\s+/g,' ');
}
function uniqueStatusBy(field){
  const ind=state.indicators[0];
  const map=new Map();
  if(!field||!ind) return map;
  state.filtered.forEach(r=>{
    const key=normKey(r[field]);
    if(!key) return;
    const yes=isYes(r[ind]);
    if(!map.has(key)) map.set(key,yes);
    else map.set(key,map.get(key)||yes); // if any covered row is Sí, the grouped unit is Sí
  });
  return map;
}
function renderCantonSummary(){
  const cantonH=state.headers.find(h=>/cant[oó]n/i.test(h));
  const m=uniqueStatusBy(cantonH);
  const total=m.size;
  const yes=[...m.values()].filter(Boolean).length;
  const no=total-yes;
  const pct=total?yes/total*100:0;
  const y=document.getElementById('cantonYes'),n=document.getElementById('cantonNo'),
        p=document.getElementById('cantonPct'),t=document.getElementById('cantonTotal');
  if(y)y.textContent=yes;if(n)n.textContent=no;if(p)p.textContent=pct.toFixed(1)+'%';
  if(t)t.textContent=total+' cantones únicos';
}

function renderTable(){const thead=$('#dataTable thead'),tbody=$('#dataTable tbody');thead.innerHTML='<tr>'+state.headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr>';const start=(state.page-1)*state.pageSize, rows=state.filtered.slice(start,start+state.pageSize);tbody.innerHTML=rows.map(r=>'<tr>'+state.headers.map(h=>{const v=r[h];if(state.indicatorCols.includes(h)&&['Sí','No'].includes(v))return `<td><span class="chip ${v==='Sí'?'yes':'no'}">${v}</span></td>`;return `<td title="${esc(v)}">${esc(v)}</td>`}).join('')+'</tr>').join('');const pages=Math.max(1,Math.ceil(state.filtered.length/state.pageSize));$('#pageInfo').textContent=`Página ${state.page} de ${pages}`;$('#prevPage').disabled=state.page<=1;$('#nextPage').disabled=state.page>=pages;$('#tableInfo').textContent=`${fmt(state.filtered.length)} registros visibles · ${fmt(state.headers.length)} columnas`}
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function renderAll(){renderKPIs();renderInsights();renderCharts();renderCantonSummary();renderTable();updateMeta()}
async function loadFile(file){if(typeof XLSX==='undefined'){toast('No se pudo cargar la librería de Excel. Revisa tu conexión a Internet.');return}$('#loading').classList.remove('hidden');try{const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:'array',cellDates:true});state.workbook=wb;state.file=file;state.sheets={};for(const name of wb.SheetNames){state.sheets[name]=parseSheet(wb.Sheets[name])}const valid=wb.SheetNames.filter(n=>state.sheets[n].rows.length&&state.sheets[n].headers.length);if(!valid.length)throw new Error('No se encontró una tabla utilizable en el archivo.');valid.sort((a,b)=>state.sheets[b].rows.length-state.sheets[a].rows.length);state.sheetName=valid[0];const sel=$('#sheetSelect');sel.innerHTML='';valid.forEach(n=>sel.add(new Option(`${n} (${state.sheets[n].rows.length})`,n)));sel.value=state.sheetName;processCurrentSheet();$('#uploadView').classList.add('hidden');$('#dashboard').classList.remove('hidden');$('#newFileBtn').classList.remove('hidden');setTimeout(()=>window.dispatchEvent(new Event('resize')),50);toast(`Archivo leído: ${fmt(state.rows.length)} registros`)}catch(e){console.error(e);toast(e.message||'No se pudo leer el archivo.')}finally{$('#loading').classList.add('hidden')}}
async function captureDashboard(format){
  try{
    showLoading(true,'Preparando listado…');
    const rows=state.filtered;
    const cantonH=state.headers.find(h=>/cant[oó]n/i.test(h));
    const delH=state.headers.find(h=>/delegaci[oó]n/i.test(h));
    const regH=state.headers.find(h=>/direcci[oó]n.*regional|regi[oó]n/i.test(h));
    const indH=state.indicators[0] || state.headers.find(h=>/recibida|estado|cumpl/i.test(h));
    if(!delH||!indH) throw new Error('No se encontraron Delegación e indicador');

    const data=rows.map(r=>({
      region: regH?String(r[regH]??'').trim():'',
      canton: cantonH?String(r[cantonH]??'').trim():'',
      delegacion:String(r[delH]??'').trim(),
      estado:isYes(r[indH])?'SÍ':'NO'
    }));

    const W=1500, margin=70, titleH=150, rowH=48;
    const H=Math.max(700,titleH+data.length*rowH+90);
    const canvas=document.createElement('canvas');
    canvas.width=W;canvas.height=H;
    const ctx=canvas.getContext('2d');
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,W,H);
    ctx.fillStyle='#0b1f36';ctx.font='700 34px Arial';
    ctx.fillText('CONTROL DE OPERACIONES',margin,55);
    ctx.font='18px Arial';ctx.fillStyle='#516276';
    ctx.fillText(`Archivo: ${state.file.name}`,margin,88);
    const yes=data.filter(x=>x.estado==='SÍ').length;
    ctx.fillText(`Registros: ${data.length}   •   Sí: ${yes}   •   No: ${data.length-yes}`,margin,118);

    const cols=[
      {k:'region',t:'Dirección Regional',x:margin,w:420},
      {k:'canton',t:'Cantón',x:490,w:300},
      {k:'delegacion',t:'Delegación',x:790,w:450},
      {k:'estado',t:'Estado',x:1240,w:190}
    ];
    ctx.fillStyle='#0e4961';ctx.fillRect(margin,145,W-margin*2,54);
    ctx.font='700 18px Arial';ctx.fillStyle='#ffffff';
    cols.forEach(c=>ctx.fillText(c.t,c.x+12,179));

    data.forEach((r,i)=>{
      const y=199+i*rowH;
      ctx.fillStyle=i%2?'#f7f9fc':'#eef4f8';ctx.fillRect(margin,y,W-margin*2,rowH);
      ctx.strokeStyle='#d6e0e8';ctx.beginPath();ctx.moveTo(margin,y+rowH);ctx.lineTo(W-margin,y+rowH);ctx.stroke();
      ctx.font='16px Arial';ctx.fillStyle='#172b3f';
      cols.slice(0,3).forEach(c=>{
        let txt=r[c.k]||'—';
        while(ctx.measureText(txt).width>c.w-24 && txt.length>3) txt=txt.slice(0,-2);
        if(txt!==r[c.k] && txt!=='—') txt+='…';
        ctx.fillText(txt,c.x+12,y+30);
      });
      const yesRow=r.estado==='SÍ';
      ctx.fillStyle=yesRow?'#d8f3e7':'#fde2e5';
      ctx.beginPath();ctx.roundRect(1270,y+9,82,30,15);ctx.fill();
      ctx.fillStyle=yesRow?'#087a4b':'#c92e43';ctx.font='700 16px Arial';
      ctx.fillText(r.estado,1295,y+30);
    });

    const mime=format==='jpg'?'image/jpeg':'image/png';
    const ext=format==='jpg'?'jpg':'png';
    const a=document.createElement('a');
    a.href=canvas.toDataURL(mime,.95);
    a.download=`${safeName(state.file.name.replace(/\.[^.]+$/,''))}_cantones_delegaciones.${ext}`;
    document.body.appendChild(a);a.click();a.remove();
    toast(`Listado ${ext.toUpperCase()} descargado`);
  }catch(err){
    console.error(err);toast('No se pudo generar el listado.');
  }finally{showLoading(false)}
}
async function exportPlot(id,format){
  try{
    showLoading(true,'Preparando gráfico…');
    const ext=format==='jpg'?'jpeg':format;
    await Plotly.downloadImage('overallChart',{format:ext,width:1400,height:1000,filename:`${safeName(state.file.name.replace(/\.[^.]+$/,''))}_circular`});
    toast(`Gráfico ${format.toUpperCase()} descargado`);
  }catch(err){console.error(err);toast('No se pudo descargar el gráfico.');}
  finally{showLoading(false)}
}
function exportCsv(){const aoa=[state.headers,...state.filtered.map(r=>state.headers.map(h=>r[h]))];const ws=XLSX.utils.aoa_to_sheet(aoa);const csv=XLSX.utils.sheet_to_csv(ws,{FS:';',RS:'\n'});const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8;'});downloadData(URL.createObjectURL(blob),`${safeName(state.file.name.replace(/\.[^.]+$/,''))}_filtrado.csv`)}
$('#chooseBtn').onclick=()=>$('#fileInput').click();$('#dropZone').onclick=e=>{if(e.target.id!=='chooseBtn')$('#fileInput').click()};$('#dropZone').onkeydown=e=>{if(e.key==='Enter'||e.key===' ')$('#fileInput').click()};$('#fileInput').onchange=e=>e.target.files[0]&&loadFile(e.target.files[0]);
['dragenter','dragover'].forEach(ev=>$('#dropZone').addEventListener(ev,e=>{e.preventDefault();$('#dropZone').classList.add('drag')}));['dragleave','drop'].forEach(ev=>$('#dropZone').addEventListener(ev,e=>{e.preventDefault();$('#dropZone').classList.remove('drag')}));$('#dropZone').addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)loadFile(f)});
$('#newFileBtn').onclick=()=>$('#fileInput').click();$('#searchInput').addEventListener('input',applyFilters);$('#sheetSelect').onchange=e=>{state.sheetName=e.target.value;processCurrentSheet()};$('#groupSelect').onchange=applyFilters;$('#filterColumnSelect').onchange=()=>{updateFilterValues();applyFilters()};$('#filterValueSelect').onchange=applyFilters;$('#clearFiltersBtn').onclick=()=>{$('#searchInput').value='';$('#filterColumnSelect').value='';updateFilterValues();applyFilters()};
$('#prevPage').onclick=()=>{if(state.page>1){state.page--;renderTable()}};$('#nextPage').onclick=()=>{if(state.page<Math.ceil(state.filtered.length/state.pageSize)){state.page++;renderTable()}};$('#exportCsvBtn').onclick=exportCsv;$('#printBtn').onclick=()=>window.print();$$('[data-capture]').forEach(b=>b.onclick=()=>captureDashboard(b.dataset.capture));$$('[data-chart]').forEach(b=>b.onclick=()=>exportPlot(b.dataset.chart,b.dataset.format));
$('#themeBtn').onclick=()=>{const d=document.documentElement;d.dataset.theme=d.dataset.theme==='dark'?'light':'dark';renderCharts()};
