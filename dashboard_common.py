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
:root{color-scheme:light;--ink:#203535;--muted:#647b79;--deep:#163f3d;--teal:#16877f;--mint:#dff3ef;--coral:#ef765f;--gold:#e9b949;--paper:#fff;--canvas:#edf4f2;--line:#d7e5e1;--red-bg:#fde7e3;--yellow-bg:#fff3cf}
*{box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;margin:0;background:var(--canvas);color:var(--ink);line-height:1.45}
header{background:linear-gradient(110deg,var(--deep),#206b64);color:#fff;padding:16px 24px;display:flex;align-items:center;gap:16px;border-bottom:4px solid var(--coral)}
header h1{margin:0;font-size:20px;flex:1}
header select{font-size:14px;padding:8px 34px 8px 10px;border:1px solid #91c7bd;border-radius:5px;background:#fff;color:var(--deep);font-weight:650;max-width:100%}
.current-section{display:none}
.dashboard-shell{display:grid;grid-template-columns:230px minmax(0,1fr);min-height:calc(100vh - 68px)}
.sidebar{position:sticky;top:0;height:calc(100vh - 68px);overflow-y:auto;background:#d7e9e5;border-right:1px solid #c1d9d3;padding:14px 10px}
.sidebar-title{margin:2px 8px 10px;color:#52726d;font-size:11px;font-weight:750;text-transform:uppercase}
.tabs{display:flex;flex-direction:column;gap:4px;padding:0;background:transparent;overflow:visible}
.tab{width:100%;text-align:left;padding:10px 12px;border:0;border-left:3px solid transparent;background:transparent;color:#294b47;cursor:pointer;border-radius:4px;font-size:13px;font-weight:600;white-space:normal;line-height:1.3}
.tab:hover{background:#c5ddd7}
.tab.active{background:var(--paper);border-left-color:var(--coral);color:var(--deep);box-shadow:0 1px 4px rgba(27,71,66,.08)}
.tab.flag{box-shadow:inset 3px 0 #d84439}
.panels{min-width:0}
.nav-toggle,.nav-scrim{display:none}
.panel{display:none;padding:20px 24px}.panel.active{display:block}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,420px),1fr));gap:16px;align-items:start}
.card{background:var(--paper);border:1px solid var(--line);border-top:3px solid var(--teal);border-radius:7px;padding:16px;box-shadow:0 3px 12px rgba(27,71,66,.07);overflow:auto;min-width:0}
.card:nth-child(3n+2){border-top-color:var(--coral)}.card:nth-child(3n){border-top-color:var(--gold)}
.card h2{margin:0 0 12px;font-size:16px;line-height:1.35;color:var(--deep)}
.wide{grid-column:1/-1}
.chart{display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.chart svg{flex:0 0 auto}
.legend div{display:flex;align-items:center;gap:6px;font-size:13px;margin:5px 0;color:#35524f}
.sw{width:12px;height:12px;border-radius:3px;display:inline-block;flex:0 0 auto}
table{border-collapse:collapse;width:100%;font-size:13px;min-width:480px}
th,td{padding:8px 9px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:#e6f1ee;color:#315450;font-size:12px}td.n,th.n{text-align:right;white-space:nowrap}
tr:nth-child(even):not(.bad):not(.warn):not(.ok) td{background:#f8fbfa}
tr.total td{font-weight:700;border-top:2px solid #83aaa2;background:#edf5f2}
tr.bad td{background:var(--red-bg)}
tr.warn td{background:var(--yellow-bg)}
td.red{color:#c62828;font-weight:700}
tr.ok td{background:#e4f4e8}
.bar{display:grid;grid-template-columns:minmax(90px,150px) minmax(60px,1fr) minmax(92px,110px);gap:8px;align-items:center;margin:8px 0;font-size:13px}
.bar .fill{height:14px;border-radius:3px;background-image:linear-gradient(90deg,rgba(255,255,255,.2),transparent)}
.empty{color:var(--muted);font-style:italic}
.issues{margin:0 0 12px;padding:10px 12px;color:#a12622;background:var(--red-bg);border-left:3px solid #d84439;font-size:13px}
.okmsg{color:#246b40;background:#e4f4e8;border-left:3px solid #4d9b62;font-size:13px;margin:0 0 12px;padding:10px 12px}
@media(max-width:640px){
  header{padding:11px 12px;gap:9px;flex-wrap:wrap}
  header h1{font-size:17px;flex:1 1 auto;min-width:0}
  header select{width:auto;max-width:42%;padding:8px}
  .current-section{display:block;flex-basis:100%;order:4;padding:5px 9px;border-left:3px solid var(--gold);background:rgba(255,255,255,.1);font-size:13px;font-weight:650;line-height:1.3}
  .dashboard-shell{display:block;min-height:calc(100vh - 64px)}
  .nav-toggle{display:inline-flex;align-items:center;justify-content:center;width:40px;height:40px;flex:0 0 40px;border:1px solid #77aaa0;border-radius:5px;background:#fff;color:var(--deep);font-size:22px;line-height:1;cursor:pointer}
  .sidebar{position:fixed;z-index:20;top:0;bottom:0;left:0;width:min(82vw,310px);height:100dvh;transform:translateX(-105%);transition:transform .2s ease;box-shadow:8px 0 24px rgba(20,55,51,.18);padding:18px 12px}
  body.nav-open .sidebar{transform:translateX(0)}
  .nav-scrim{position:fixed;inset:0;z-index:19;background:rgba(17,43,40,.45)}
  body.nav-open .nav-scrim{display:block}
  .sidebar-title{font-size:12px;margin:4px 10px 14px}
  .tabs{gap:5px}
  .tab{min-height:44px;padding:11px 12px;font-size:14px}
  .panel{padding:12px 10px}
  .grid{grid-template-columns:minmax(0,1fr);gap:10px}
  .card{padding:13px}
  .card h2{font-size:15px}
  .chart{justify-content:center;gap:8px}
  .chart svg{width:156px;height:156px}
  .legend{width:100%}
  .bar{grid-template-columns:minmax(82px,112px) minmax(40px,1fr) minmax(86px,98px);gap:6px;font-size:12px}
  table{font-size:12px}
  th,td{padding:7px}
}
</style></head><body>
<header><button class="nav-toggle" id="navToggle" type="button" aria-label="Open dashboard sections" aria-expanded="false">&#9776;</button><h1>Society Audit Dashboard</h1><select id="fy" aria-label="Financial year"></select><div class="current-section" id="currentSection" aria-live="polite"></div></header>
<div class="dashboard-shell"><aside class="sidebar" id="sidebar"><div class="sidebar-title">Dashboard sections</div><nav class="tabs" id="tabs" aria-label="Dashboard sections"></nav></aside><button class="nav-scrim" id="navScrim" type="button" aria-label="Close dashboard sections"></button><main class="panels" id="panels"></main></div>
<script>
const DATASETS=__DATA__;let DATA=[];const SQFT=__SQFT__;
const COLORS=["#4e79a7","#f28e2b","#59a14f","#e15759","#76b7b2","#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac"];
const STATUS_COLORS={"Matched":"#59a14f","Partially matched":"#edc948","Unmatched":"#e15759","Pending":"#c62828","Not expense / no income entry":"#ef9a9a","Review":"#ff8a80"};
const inr=v=>Number(v).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const tabs=document.getElementById("tabs"),panels=document.getElementById("panels");
const navToggle=document.getElementById("navToggle");
function closeNav(){document.body.classList.remove("nav-open");navToggle.setAttribute("aria-expanded","false");}
navToggle.onclick=()=>{const open=document.body.classList.toggle("nav-open");navToggle.setAttribute("aria-expanded",String(open));};
document.getElementById("navScrim").onclick=closeNav;
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
  const hasLiftAmc=DATA.some(m=>m.summary&&m.summary.lift_amc);
  const rows=DATA.map(m=>({month:m.label,gross:m.debits.gross,exclusions:m.debits.exclusions.map(x=>`${x.label}: INR ${inr(x.amount)}`).join("; ")||"None",net:m.debits.net,rate:(m.debits.net/SQFT).toFixed(4),rateWithAmc:((m.debits.net+(m.summary?.lift_amc||0))/SQFT).toFixed(4)}));
  const cols=[{h:"Month",k:"month"},{h:"All debits",k:"gross",n:1},{h:"Exclusions",k:"exclusions"},{h:"Net debited",k:"net",n:1},{h:`Rate per sqft (&#8377;, ${SQFT.toLocaleString("en-IN")} sqft)`,k:"rate",r:1}];
  if(hasLiftAmc)cols.push({h:`Rate per sqft + lift AMC (&#8377;, ${SQFT.toLocaleString("en-IN")} sqft)`,k:"rateWithAmc",r:1});
  return '<div class="grid"><div class="card wide"><h2>Money debited from account per month, less listed exclusions</h2>'+bars(items,true)+'</div>'
    +'<div class="card wide"><h2>Debit breakup</h2>'+table(cols,rows,"No data")+'</div></div>';
}
function show(i){[...tabs.children].forEach((t,j)=>{t.classList.toggle("active",i===j);t.setAttribute("aria-current",i===j?"page":"false");});[...panels.children].forEach((p,j)=>p.classList.toggle("active",i===j));document.getElementById("currentSection").textContent="Viewing: "+(tabs.children[i]?.textContent||"");closeNav();}
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
