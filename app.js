let rawRows=[];
let charts={};
const INR=new Intl.NumberFormat('en-IN',{style:'currency',currency:'INR',maximumFractionDigits:2});
const compactINR=new Intl.NumberFormat('en-IN',{style:'currency',currency:'INR',notation:'compact',maximumFractionDigits:1});
const el=id=>document.getElementById(id);

function money(v){return INR.format(Number(v)||0)}
function compactMoney(v){return compactINR.format(Number(v)||0)}
function parseDate(v){
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==='number' && window.XLSX && XLSX.SSF) { const o=XLSX.SSF.parse_date_code(v); if(o) return new Date(o.y,o.m-1,o.d); }
  const s=String(v??'').trim(); if(!s) return null;
  const d=new Date(s); if(!isNaN(d)) return new Date(d.getFullYear(),d.getMonth(),d.getDate());
  const m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/); if(m) return new Date(+m[3],+m[2]-1,+m[1]);
  return null;
}
function dateKey(d){return d.toISOString().slice(0,10)}
function fmtDate(d){return d.toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'})}
function monthKey(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`}
function monthLabel(key){const [y,m]=key.split('-').map(Number);return new Date(y,m-1,1).toLocaleDateString('en-IN',{month:'long',year:'numeric'})}

function normalize(rows){
  return rows.map(r=>{
    const get=(names)=>{for(const n of names){if(r[n]!==undefined) return r[n]} return ''};
    const d=parseDate(get(['Date','DATE','date']));
    const amount=Number(get(['Amount (INR)','Amount','amount','AMOUNT']))||0;
    return {date:d,description:String(get(['Expense Description','Expense Name','Description','Expense Description '])||'').trim(),type:String(get(['Expense Type','Category','Type'])||'Uncategorised').trim()||'Uncategorised',amount,payment:String(get(['Payment Method','Payment','Mode'])||'Not specified').trim()||'Not specified',notes:String(get(['Notes','Note'])||'').trim()};
  }).filter(x=>x.date && x.description && x.amount>0);
}

function loadWorkbook(data,name='Excel file'){
  const wb=XLSX.read(data,{type:'array',cellDates:true});
  let chosen=wb.SheetNames.find(s=>/expense/i.test(s))||wb.SheetNames[0];
  const ws=wb.Sheets[chosen];
  rawRows=normalize(XLSX.utils.sheet_to_json(ws,{defval:''}));
  rawRows.sort((a,b)=>a.date-b.date);
  el('fileStatus').textContent=`${name} · ${rawRows.length} expenses`;
  el('subhead').textContent=`${chosen} · ${rawRows.length} valid expense rows loaded locally in your browser.`;
  populateFilters(); render();
}

function populateFilters(){
  const cats=[...new Set(rawRows.map(r=>r.type))].sort();
  const pays=[...new Set(rawRows.map(r=>r.payment))].sort();
  const months=[...new Set(rawRows.map(r=>monthKey(r.date)))].sort().reverse();
  el('categorySelect').innerHTML='<option value="ALL">All categories</option>'+cats.map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  el('paymentSelect').innerHTML='<option value="ALL">All payment methods</option>'+pays.map(x=>`<option>${escapeHtml(x)}</option>`).join('');
  el('monthSelect').innerHTML=months.length?months.map(x=>`<option value="${x}">${monthLabel(x)}</option>`).join(''):'<option value="ALL">No data</option>';
}

function filtered(){
  const month=el('monthSelect').value;
  const cat=el('categorySelect').value;
  const pay=el('paymentSelect').value;
  const q=el('searchInput').value.toLowerCase().trim();
  return rawRows.filter(r=>
    (month==='ALL'||monthKey(r.date)===month) &&
    (cat==='ALL'||r.type===cat) &&
    (pay==='ALL'||r.payment===pay) &&
    (!q||`${r.description} ${r.type} ${r.payment} ${r.notes}`.toLowerCase().includes(q))
  );
}

function render(){
  const rows=filtered();
  const month=el('monthSelect').value;
  const total=rows.reduce((s,r)=>s+r.amount,0);
  const max=rows.reduce((a,b)=>b.amount>a.amount?b:a,{amount:0});
  const activeDays=new Set(rows.map(r=>dateKey(r.date))).size;
  el('totalSpend').textContent=money(total);
  el('txnCount').textContent=rows.length.toLocaleString('en-IN');
  el('dailyAvg').textContent=money(activeDays?total/activeDays:0);
  el('largestExpense').textContent=money(max.amount);
  el('largestLabel').textContent=max.description||'—';
  el('periodLabel').textContent=month==='ALL'?'All loaded data':monthLabel(month);
  renderCharts(rows);
  renderTables(rows);
}

function renderCharts(rows){
  Object.values(charts).forEach(c=>c?.destroy());
  const byDay=new Map(), byCat=new Map();
  rows.forEach(r=>{const dk=dateKey(r.date);byDay.set(dk,(byDay.get(dk)||0)+r.amount);byCat.set(r.type,(byCat.get(r.type)||0)+r.amount)});
  const days=[...byDay.keys()].sort();
  charts.daily=new Chart(el('dailyChart'),{type:'line',data:{labels:days.map(d=>new Date(d+'T00:00:00').toLocaleDateString('en-IN',{day:'2-digit',month:'short'})),datasets:[{data:days.map(d=>byDay.get(d)),borderWidth:2,fill:true,tension:.28,pointRadius:2}]},options:baseChart({scales:{y:{ticks:{callback:v=>compactMoney(v)}},x:{grid:{display:false}}}})});
  const cats=[...byCat.entries()].sort((a,b)=>b[1]-a[1]);
  charts.category=new Chart(el('categoryChart'),{type:'doughnut',data:{labels:cats.map(x=>x[0]),datasets:[{data:cats.map(x=>x[1]),borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:'#c9d1df',boxWidth:12,font:{size:11}}},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${money(ctx.raw)}`}}}}});
  const top=cats.slice(0,8);
  charts.bar=new Chart(el('barChart'),{type:'bar',data:{labels:top.map(x=>x[0]),datasets:[{data:top.map(x=>x[1]),borderRadius:7}]},options:baseChart({indexAxis:'y',scales:{x:{ticks:{callback:v=>compactMoney(v)}},y:{grid:{display:false}}}})});
}
function baseChart(extra={}){return {responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${money(ctx.raw)}`}}},scales:{x:{ticks:{color:'#8f9bb0'},grid:{color:'#1c2435'}},y:{ticks:{color:'#8f9bb0'},grid:{color:'#1c2435'}}},...extra};}

function renderTables(rows){
  const top=[...rows].sort((a,b)=>b.amount-a.amount).slice(0,10);
  el('topExpensesBody').innerHTML=top.length?top.map(r=>`<tr><td>${fmtDate(r.date)}</td><td>${escapeHtml(r.description)}</td><td><span class="pill">${escapeHtml(r.type)}</span></td><td class="right amount">${money(r.amount)}</td></tr>`).join(''):`<tr><td colspan="4" class="empty">No matching expenses.</td></tr>`;
  const recent=[...rows].sort((a,b)=>b.date-a.date||b.amount-a.amount).slice(0,100);
  el('logSummary').textContent=`${rows.length.toLocaleString('en-IN')} rows · showing ${recent.length}`;
  el('logBody').innerHTML=recent.length?recent.map(r=>`<tr><td>${fmtDate(r.date)}</td><td>${escapeHtml(r.description)}</td><td>${escapeHtml(r.type)}</td><td>${escapeHtml(r.payment)}</td><td class="right amount">${money(r.amount)}</td><td>${escapeHtml(r.notes)}</td></tr>`).join(''):`<tr><td colspan="6" class="empty">No matching expenses.</td></tr>`;
}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));}

el('fileInput').addEventListener('change',async e=>{const f=e.target.files?.[0];if(f) loadWorkbook(await f.arrayBuffer(),f.name)});
['monthSelect','categorySelect','paymentSelect'].forEach(id=>el(id).addEventListener('change',render));
el('searchInput').addEventListener('input',render);
el('clearBtn').addEventListener('click',()=>{rawRows=[];el('fileInput').value='';el('fileStatus').textContent='No file loaded';el('subhead').textContent='Load your Excel file to turn your expense log into a live dashboard.';el('monthSelect').innerHTML='<option value="ALL">No data</option>';el('categorySelect').innerHTML='<option value="ALL">All categories</option>';el('paymentSelect').innerHTML='<option value="ALL">All payment methods</option>';['totalSpend','txnCount','dailyAvg','largestExpense'].forEach(id=>el(id).textContent=id==='txnCount'?'0':'₹0');el('largestLabel').textContent='—';Object.values(charts).forEach(c=>c?.destroy());charts={};renderTables([])});
const dz=el('dropzone');['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));dz.addEventListener('drop',async e=>{const f=e.dataTransfer.files?.[0];if(f){el('fileInput').value='';loadWorkbook(await f.arrayBuffer(),f.name)}});
renderTables([]);
