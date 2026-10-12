"""Shared helpers and the HTML page shell used by every financial-year dashboard page."""
import datetime as dt
import json

TOTAL_SQFT = 469655

# First matching keyword wins; anything else is "Others".
EXPENSE_CATEGORIES = [
    ("cornerstone", "HouseKeeping"),
    ("force security", "Security"),
    ("veeranna", "Water Tanker"),
    ("gowthami", "STP/WTP"),
    ("star product", "Consumables"),
    ("keerthi", "Diesel"),
    ("staff salary", "Staff Salary"),
    ("bescom", "Electricity (BESCOM)"),
    ("bwssb", "Water (BWSSB)"),
    ("creative", "Gardening"),
    ("nature facility", "Gardening"),
    ("ecosphere", "Garbage"),
    ("narasimha", "Garbage"),
    ("pool vibes", "Swimming Pool"),
]


def num(value):
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").replace("₹", "").strip())
    except ValueError:
        return 0.0


def fmt_date(value):
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%d-%b-%Y")
    return "" if value is None else str(value)


def categorize_expense(name):
    low = str(name or "").lower()
    for key, label in EXPENSE_CATEGORIES:
        if key in low:
            return label
    return "Others"


PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Society Audit Dashboard</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;margin:0;background:#f4f6f9;color:#222}
header{background:#1f3a5f;color:#fff;padding:14px 24px;display:flex;align-items:center;gap:16px}
header h1{margin:0;font-size:20px}
header select{font-size:14px;padding:4px 8px;border-radius:4px}
.tabs{display:flex;flex-wrap:wrap;gap:4px;padding:10px 24px 0;background:#e3e8ef}
.tab{padding:8px 18px;border:0;background:#cfd7e3;cursor:pointer;border-radius:6px 6px 0 0;font-size:14px}
.tab.active{background:#fff;font-weight:600}
.tab.flag{border-bottom:3px solid #e15759}
.panel{display:none;padding:20px 24px}.panel.active{display:block}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px}
.card{background:#fff;border-radius:8px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.12);overflow:auto}
.card h2{margin:0 0 12px;font-size:16px}
.wide{grid-column:1/-1}
.chart{display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.legend div{display:flex;align-items:center;gap:6px;font-size:13px;margin:3px 0}
.sw{width:12px;height:12px;border-radius:2px;display:inline-block}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:6px 8px;border-bottom:1px solid #e5e9f0;text-align:left}
th{background:#eef2f7}td.n,th.n{text-align:right}
tr.total td{font-weight:700;border-top:2px solid #999}
tr.bad td{background:#fde8e8}
tr.warn td{background:#fff6cc}
td.red{color:#c62828;font-weight:700}
tr.ok td{background:#eaf6ea}
.bar{display:grid;grid-template-columns:150px 1fr 110px;gap:8px;align-items:center;margin:6px 0;font-size:13px}
.bar .fill{height:16px;border-radius:3px}
.empty{color:#888;font-style:italic}
.issues{margin:0 0 12px;padding-left:18px;color:#a12622;font-size:13px}
.okmsg{color:#2e7d32;font-size:13px;margin:0 0 12px}
</style></head><body>
<header><h1>Society Audit Dashboard</h1><select id="fy"></select></header>
<div class="tabs" id="tabs"></div><div id="panels"></div>
<script>
const DATASETS=__DATA__;let DATA=[];const SQFT=__SQFT__;
const COLORS=["#4e79a7","#f28e2b","#59a14f","#e15759","#76b7b2","#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac"];
const STATUS_COLORS={"Matched":"#59a14f","Partially matched":"#edc948","Unmatched":"#e15759","Pending":"#c62828","Not expense / no income entry":"#ef9a9a","Review":"#ff8a80"};
const inr=v=>Number(v).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const tabs=document.getElementById("tabs"),panels=document.getElementById("panels");
const card=(t,b,w)=>`<div class="card ${w?'wide':''}"><h2>${t}</h2>${b}</div>`;
function pie(items,colorOf,fmt){
  const total=items.reduce((a,i)=>a+i.value,0);
  if(!total)return '<div class="empty">No data</div>';
  const R=90,C=100;let a=-Math.PI/2,svg='';
  items.forEach((it,i)=>{
    const col=colorOf(it,i),f=it.value/total;
    if(f>=0.9999){svg+=`<circle cx="${C}" cy="${C}" r="${R}" fill="${col}"><title>${esc(it.label)}</title></circle>`;return;}
    const b=a+f*2*Math.PI,x1=C+R*Math.cos(a),y1=C+R*Math.sin(a),x2=C+R*Math.cos(b),y2=C+R*Math.sin(b);
    svg+=`<path d="M${C},${C} L${x1},${y1} A${R},${R} 0 ${f>0.5?1:0} 1 ${x2},${y2} Z" fill="${col}"><title>${esc(it.label)}: ${fmt(it.value)}</title></path>`;a=b;
  });
  const legend=items.map((it,i)=>`<div><span class="sw" style="background:${colorOf(it,i)}"></span>${esc(it.label)}: <b>${fmt(it.value)}</b> (${(it.value*100/total).toFixed(1)}%)</div>`).join('');
  return `<div class="chart"><svg width="200" height="200" viewBox="0 0 200 200">${svg}</svg><div class="legend">${legend}</div></div>`;
}
function bars(items,noTotal){
  if(!items.length)return '<div class="empty">No data</div>';
  const max=Math.max(...items.map(i=>i.value));
  return items.map((it,i)=>`<div class="bar"><span>${esc(it.label)}</span><div><div class="fill" style="width:${it.value*100/max}%;background:${COLORS[i%COLORS.length]}"></div></div><span style="text-align:right">&#8377; ${inr(it.value)}</span></div>`).join('')
    +(noTotal?'':`<div class="bar"><b>Total</b><span></span><b style="text-align:right">&#8377; ${inr(items.reduce((a,i)=>a+i.value,0))}</b></div>`);
}
function table(cols,rows,empty){
  if(!rows.length)return `<div class="empty">${empty}</div>`;
  return '<table><tr>'+cols.map(c=>`<th class="${c.n||c.r?'n':''}">${c.h}</th>`).join('')+'</tr>'
    +rows.map(r=>`<tr class="${r.cls||''}">`+cols.map(c=>`<td class="${c.n||c.r?'n':''}${c.red&&r[c.k]?' red':''}">${c.n?(r[c.k]?inr(r[c.k]):''):esc(r[c.k]??'')}</td>`).join('')+'</tr>').join('')+'</table>';
}
function debitsPanel(){
  const items=DATA.map(m=>({label:m.label,value:m.debits.net}));
  const rows=DATA.map(m=>({month:m.label,...m.debits}));
  rows.forEach(r=>r.rate=(r.net/SQFT).toFixed(4));
  const cols=[{h:"Month",k:"month"},{h:"All debits",k:"gross",n:1},{h:"FD transactions (excluded)",k:"fd",n:1},{h:"Sponsorship payout (excluded)",k:"reimbursement",n:1},{h:"Net debited",k:"net",n:1},{h:`Rate per sqft (&#8377;, ${SQFT.toLocaleString("en-IN")} sqft)`,k:"rate",r:1}];
  return '<div class="grid"><div class="card wide"><h2>Money debited from account per month (excluding FD transactions and sponsorship payouts)</h2>'+bars(items,true)+'</div>'
    +'<div class="card wide"><h2>Debit breakup</h2>'+table(cols,rows,"No data")+'</div></div>';
}
function show(i){[...tabs.children].forEach((t,j)=>t.classList.toggle("active",i===j));[...panels.children].forEach((p,j)=>p.classList.toggle("active",i===j));}
function addTab(label,html,flag){
  const idx=tabs.children.length;
  const b=document.createElement("button");b.className="tab"+(flag?" flag":"");b.textContent=label;b.onclick=()=>show(idx);tabs.appendChild(b);
  const p=document.createElement("div");p.className="panel";p.innerHTML=html;panels.appendChild(p);
}
const fySel=document.getElementById("fy");
Object.keys(DATASETS).forEach(k=>{const o=document.createElement("option");o.value=o.textContent=k;fySel.appendChild(o);});
function render(label){
  DATA=DATASETS[label];tabs.innerHTML="";panels.innerHTML="";
  if(!DATA.length){panels.innerHTML='<div class="panel active">No data found.</div>';return;}
  renderYear();
}
fySel.onchange=()=>render(fySel.value);
__BODY__
render(fySel.value);
</script></body></html>
"""


def build_page(datasets, body_js):
    """datasets: {label: [month dicts]} ordered newest year first; the first one is shown initially."""
    html = PAGE_TEMPLATE.replace("__SQFT__", str(TOTAL_SQFT)).replace("__BODY__", body_js)
    # Data goes in last so user text can never collide with the placeholders above.
    return html.replace("__DATA__", json.dumps(datasets).replace("</", "<\\/"))
