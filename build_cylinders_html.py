#!/usr/bin/env python3
"""
Build qsk50_cylinders.html — cylinder exhaust temperature analysis page.
Format matches the PowerPoint slide: bank charts, deviation bar, interbank delta.
Usage: python3 build_cylinders_html.py
"""

import io, json, re, pathlib, warnings
import pandas as pd

warnings.filterwarnings('ignore')

EXTRACT   = pathlib.Path('/home/user/NTE200/.qsk50_extracted')
PLOTLY_JS = pathlib.Path('/usr/local/lib/python3.11/dist-packages/plotly/package_data/plotly.min.js')
OUTPUT    = pathlib.Path('/home/user/NTE200/qsk50_cylinders.html')

MAX_ROWS = 600   # max time points per truck

RU_MONTHS = {
    'янв':'Jan','фев':'Feb','мар':'Mar','апр':'Apr','май':'May','июн':'Jun',
    'июл':'Jul','авг':'Aug','сен':'Sep','окт':'Oct','ноя':'Nov','дек':'Dec',
}
A_BANK = [1,3,5,7,9,11,13,15]   # odd = ECM144
B_BANK = [2,4,6,8,10,12,14,16]  # even = ECM1


def ru_date(d):
    for ru, en in RU_MONTHS.items():
        d = d.replace(ru, en)
    return d


def open_file(fp):
    with open(fp, 'rb') as fh:
        bom = fh.read(3)
    enc_order = ('utf-8-sig', 'cp1251') if bom[:3] == b'\xef\xbb\xbf' else ('cp1251', 'utf-8-sig', 'utf-8')
    for enc in enc_order:
        try:
            with open(fp, encoding=enc, errors='replace') as fh:
                raw = fh.read()
            if 'Дата' in raw or '"Date"' in raw or 'Date,' in raw:
                return raw.splitlines(keepends=True), enc
        except Exception:
            pass
    return None, None


def find_cyl_col(cols, n):
    """Find exhaust temperature column for cylinder n."""
    for c in cols:
        # Russian: "Датчик температуры отработавших газов цилиндра N (°C)- ИдентификаторXXX"
        if f'отработавших газов цилиндра {n} (°C)' in c:
            return c
        # English: "Exhaust Temperature Sensor Cylinder N (°C)- IdentifierXXX"
        if f'Exhaust Temperature Sensor Cylinder {n} (' in c:
            return c
    return None


def read_csv(fp):
    lines, enc = open_file(fp)
    if lines is None:
        return None

    hdr = None
    for i, line in enumerate(lines[:80]):
        s = line.strip()
        if s.startswith('"Date"') or s.startswith('Date,') or \
           s.startswith('"Дата"') or s.startswith('Дата,'):
            hdr = i
            break
    if hdr is None:
        return None

    try:
        df = pd.read_csv(io.StringIO(''.join(lines[hdr:])), low_memory=False)
    except Exception:
        return None

    if len(df) < 10:
        return None

    all_cols  = list(df.columns)
    date_col  = 'Date'  if 'Date'  in all_cols else 'Дата'
    time_col  = 'Time'  if 'Time'  in all_cols else 'Время'

    # Find cylinder columns
    cyl_cols = {}
    for n in range(1, 17):
        c = find_cyl_col(all_cols, n)
        if c:
            cyl_cols[n] = c

    if len(cyl_cols) < 8:
        return None   # Need at least half the cylinders

    keep = [date_col, time_col] + list(cyl_cols.values())
    df_sub = df[keep].copy()

    # Numeric conversion
    for col in cyl_cols.values():
        try:
            conv = df_sub[col].astype(str).str.replace(',', '.', regex=False).str.strip()
            num  = pd.to_numeric(conv, errors='coerce')
            if num.notna().sum() > len(df_sub) * 0.25:
                df_sub[col] = num.round(0)
        except Exception:
            pass

    # Timestamps
    ts_list = []
    for _, row in df_sub.iterrows():
        d = ru_date(str(row[date_col]))
        t = str(row[time_col])[:8]
        try:
            dt = pd.to_datetime(f"{d} {t}", dayfirst=True, errors='coerce')
            ts_list.append(dt.isoformat()[:19] if pd.notna(dt) else None)
        except Exception:
            ts_list.append(None)
    df_sub['_ts'] = ts_list
    df_sub = df_sub[df_sub['_ts'].notna()].copy()

    if len(df_sub) < 10:
        return None

    # Downsample
    step = max(1, len(df_sub) // MAX_ROWS)
    df_sub = df_sub.iloc[::step].reset_index(drop=True)

    ts = df_sub['_ts'].tolist()

    # Build cyl data: {1: [vals], 2: [vals], ...}
    cyls = {}
    for n, col in cyl_cols.items():
        series = df_sub[col]
        vals = [None if pd.isna(v) else int(v) for v in series]
        cyls[n] = vals

    # Date string from first timestamp
    date_str = ts[0][:10].replace('-', '.') if ts else ''
    # Reformat dd.mm.yyyy (ISO is yyyy-mm-dd)
    parts = date_str.split('.')
    if len(parts) == 3:
        date_str = f"{parts[2]}.{parts[1]}.{parts[0]}"

    return {
        'ts':   ts,
        'cyls': cyls,
        'date': date_str,
        'n_cyls': len(cyl_cols),
    }


def truck_num(fname):
    for pat in [r'№\s*(\d+)', r'\bN(\d+)\b', r'\((\d+)\)',
                r'[\s_-](\d{2,3})[\s_.( ]']:
        m = re.search(pat, fname)
        if m:
            return int(m.group(1))
    m = re.search(r'DML-\d+-\d+\s+(\d+)', fname)
    if m:
        return int(m.group(1))
    return None


def build_data():
    db, order = {}, []
    for fp in sorted(EXTRACT.glob('*.csv')):
        name = fp.name
        if 'production' in name.lower() or 'speed' in name.lower():
            continue
        truck = truck_num(name)
        if truck is None:
            continue
        model = 'NTE200' if ('NTE' in name or 'DML' in name) else '730E'
        key   = f"{model}_{truck}"
        print(f"  {name} … ", end='', flush=True)
        result = read_csv(fp)
        if result is None:
            print("no cyl data")
            continue
        print(f"{len(result['ts'])} rows, {result['n_cyls']} cyls")
        if key not in db:
            db[key] = {**result, 'truck': truck, 'model': model, 'file': name}
            order.append(key)
        else:
            db[key]['ts'].extend(result['ts'])
            for n, vals in result['cyls'].items():
                if n in db[key]['cyls']:
                    db[key]['cyls'][n].extend(vals)
                else:
                    db[key]['cyls'][n] = vals
    return db, order


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>QSK50 — Температура цилиндров</title>
<script>__PLOTLY__</script>
<style>
:root{
  --bg:#0d1520;--bg2:#111e2d;--card:#162030;--brd:rgba(80,140,200,0.13);
  --tx:#cce0f5;--tx2:#7aaabb;--tx3:#3d6080;
  --acc:#00d4aa;--red:#ff4060;--amb:#ffb830;--grn:#2ecc71;
  --blu:#5b8fff;--pur:#b06bff;--org:#ff7c3a;
  --a-color:#ff7c3a;--b-color:#5b8fff;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--tx);font-size:13px;display:flex;flex-direction:column}

.hdr{display:flex;align-items:center;gap:14px;padding:0 18px;height:50px;border-bottom:1px solid var(--brd);background:rgba(8,14,22,0.97);flex-shrink:0}
.hdr-title{font-weight:800;font-size:15px;color:#fff;letter-spacing:-.3px}
.hdr-title span{color:var(--tx3);font-weight:400;font-size:10px;margin-left:8px;vertical-align:middle}
.hdr-date{font-size:11px;color:var(--amb);font-weight:600;margin-left:auto;font-family:'SF Mono',monospace}

.truck-row{display:flex;align-items:center;gap:6px;padding:6px 18px;border-bottom:1px solid var(--brd);background:var(--bg2);flex-shrink:0;flex-wrap:wrap;min-height:38px}
.tr-label{font-size:10px;color:var(--tx3);font-weight:700;text-transform:uppercase;letter-spacing:.6px;margin-right:4px}
.pill{display:flex;align-items:center;gap:5px;padding:4px 11px;border:1px solid var(--brd);border-radius:20px;font-size:11px;color:var(--tx2);cursor:pointer;transition:all .14s;user-select:none}
.pill:hover{border-color:var(--acc);color:var(--acc)}
.pill.act{background:rgba(0,212,170,0.1);border-color:var(--acc);color:var(--acc);font-weight:600}
.pill .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.no-data-msg{font-size:11px;color:var(--tx3)}

.main{flex:1;overflow-y:auto;overflow-x:hidden;padding:12px}
.main::-webkit-scrollbar{width:5px}
.main::-webkit-scrollbar-thumb{background:var(--brd);border-radius:3px}

.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.card{background:var(--card);border:1px solid var(--brd);border-radius:10px;padding:10px;overflow:hidden}
.card-title{font-size:11px;font-weight:700;color:var(--tx2);margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:8px}
.tag{padding:2px 7px;border-radius:4px;font-size:9px;font-weight:700;letter-spacing:.4px}
.tag-a{background:rgba(255,124,58,0.2);color:var(--org)}
.tag-b{background:rgba(91,143,255,0.2);color:var(--blu)}

.stats-box{display:flex;gap:18px;padding:6px 10px;background:rgba(255,255,255,0.04);border-radius:6px;border:1px solid var(--brd);margin-top:6px;flex-wrap:wrap}
.stat{text-align:center}
.stat .sv{font-size:16px;font-weight:700;font-family:'SF Mono',monospace;color:var(--acc)}
.stat .sl{font-size:9px;color:var(--tx3);text-transform:uppercase;letter-spacing:.5px;margin-top:1px}
.stat.warn .sv{color:var(--amb)}
.stat.crit .sv{color:var(--red)}

.conclusion{background:linear-gradient(90deg,rgba(255,124,58,0.08),rgba(91,143,255,0.08));border:1px solid rgba(255,124,58,0.25);border-radius:8px;padding:10px 16px;font-size:11px;color:var(--tx);line-height:1.8;margin-top:2px}
.conclusion strong{color:#fff;font-size:12px}
.c-pill{display:inline-block;padding:1px 8px;border-radius:4px;font-weight:700;font-size:11px;margin:0 2px}
.c-a{background:rgba(255,124,58,0.2);color:var(--org)}
.c-b{background:rgba(91,143,255,0.2);color:var(--blu)}
.c-warn{background:rgba(255,184,48,0.2);color:var(--amb)}
.c-ok{background:rgba(46,204,113,0.2);color:var(--grn)}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-title">⚙ QSK50 <span>Cummins MCRS · Анализ температуры цилиндров</span></div>
  <div class="hdr-date" id="hdr-date">–</div>
</div>

<div class="truck-row" id="truck-row">
  <span class="tr-label">Самосвал:</span>
  <span class="no-data-msg" id="no-data-msg">нет данных</span>
</div>

<div class="main">
  <div class="row2">
    <div class="card">
      <div class="card-title">
        <span class="tag tag-a">А-банк</span>
        Нечётные цилиндры — ECM144
      </div>
      <div id="ch-a-bank"></div>
      <div class="stats-box" id="stats-a"></div>
    </div>
    <div class="card">
      <div class="card-title">
        <span class="tag tag-b">В-банк</span>
        Чётные цилиндры — ECM1
      </div>
      <div id="ch-b-bank"></div>
      <div class="stats-box" id="stats-b"></div>
    </div>
  </div>

  <div class="card" style="margin-bottom:12px">
    <div class="card-title">∆ каждого цилиндра от среднего по своему банку</div>
    <div id="ch-deviation"></div>
  </div>

  <div class="card" style="margin-bottom:12px">
    <div class="card-title">Межбанковая дельта: В-банк avg − А-банк avg (по времени)</div>
    <div id="ch-delta"></div>
  </div>

  <div class="conclusion" id="conclusion">Выберите самосвал</div>
</div>

<script>
// ── EMBEDDED DATA ──────────────────────────────────────────────────────────
const TRUCKS = __DATA__;
const TRUCK_ORDER = __ORDER__;

// ── CONFIG ─────────────────────────────────────────────────────────────────
const A_CYLS = [1,3,5,7,9,11,13,15];
const B_CYLS = [2,4,6,8,10,12,14,16];
const COLORS_A = ['#ff7c3a','#ffb830','#e55c3a','#cc4020','#ff9860','#f0c000','#dd7060','#aa3010'];
const COLORS_B = ['#5b8fff','#00bfff','#7066ee','#9955dd','#3399ff','#44aaff','#8877ff','#bb77ee'];
const TRUCK_COLORS = ['#00d4aa','#5b8fff','#ffb830','#ff4060','#b06bff','#00bfff','#2ecc71','#ff7c3a'];

const PLY_CFG = {responsive:true,displaylogo:false,modeBarButtonsToRemove:['select2d','lasso2d']};
const PLY_BASE = {
  paper_bgcolor:'#162030',plot_bgcolor:'rgba(10,20,35,0.7)',
  font:{color:'#cce0f5',family:'Segoe UI,sans-serif',size:10},
  margin:{l:48,r:14,t:28,b:36},
  xaxis:{gridcolor:'rgba(80,140,200,0.07)',zerolinecolor:'rgba(80,140,200,0.12)',color:'#3d6080'},
  yaxis:{gridcolor:'rgba(80,140,200,0.07)',zerolinecolor:'rgba(80,140,200,0.12)',color:'#3d6080'},
  legend:{bgcolor:'rgba(0,0,0,0)',borderwidth:0,font:{size:9}},
};
function pl(id,data,lay={}){
  const L=Object.assign({},PLY_BASE,lay);
  if(lay.xaxis)L.xaxis=Object.assign({},PLY_BASE.xaxis,lay.xaxis);
  if(lay.yaxis)L.yaxis=Object.assign({},PLY_BASE.yaxis,lay.yaxis);
  Plotly.react(id,data,L,PLY_CFG);
}

// ── MATH ────────────────────────────────────────────────────────────────────
function nn(a){return a.filter(v=>v!==null&&!isNaN(v));}
function mean(a){const b=nn(a);return b.length?b.reduce((s,v)=>s+v,0)/b.length:null;}
function maxv(a){const b=nn(a);return b.length?Math.max(...b):null;}
function minv(a){const b=nn(a);return b.length?Math.min(...b):null;}
function rowMeans(cyls,keys){
  const n=cyls[keys[0]].length;
  return Array.from({length:n},(_,i)=>{
    const vals=keys.map(k=>cyls[k]?cyls[k][i]:null).filter(v=>v!==null&&!isNaN(v));
    return vals.length?vals.reduce((s,v)=>s+v,0)/vals.length:null;
  });
}
function timeSecs(ts){
  if(!ts||!ts.length)return [];
  const t0=new Date(ts[0]).getTime();
  return ts.map(t=>t?Math.round((new Date(t).getTime()-t0)/1000):null);
}
function ewma(arr,alpha=0.06){let s=null;return arr.map(v=>{if(v!==null){s=s===null?v:alpha*v+(1-alpha)*s;return Math.round(s);}return s;});}

// ── TRUCK PILLS ─────────────────────────────────────────────────────────────
let activeTruck = null;

function buildPills(){
  const row = document.getElementById('truck-row');
  row.querySelectorAll('.pill').forEach(p=>p.remove());
  const noMsg = document.getElementById('no-data-msg');
  if(noMsg) noMsg.style.display = TRUCK_ORDER.length?'none':'';
  TRUCK_ORDER.forEach((key,i)=>{
    const d = TRUCKS[key];
    if(!d) return;
    const pill = document.createElement('div');
    pill.className = 'pill' + (i===0?' act':'');
    pill.dataset.key = key;
    const dot = document.createElement('span');
    dot.className = 'dot';
    dot.style.background = TRUCK_COLORS[i%TRUCK_COLORS.length];
    const lbl = document.createElement('span');
    lbl.textContent = d.model + ' №' + d.truck;
    pill.appendChild(dot); pill.appendChild(lbl);
    pill.onclick = () => {
      document.querySelectorAll('.pill').forEach(p=>p.classList.remove('act'));
      pill.classList.add('act');
      activeTruck = key;
      render(key);
    };
    row.appendChild(pill);
  });
  if(TRUCK_ORDER.length){
    activeTruck = TRUCK_ORDER[0];
  }
}

// ── STATS BOX ───────────────────────────────────────────────────────────────
function renderStats(id, avg, maxVal, delta){
  const el = document.getElementById(id);
  if(!el) return;
  const avgC = avg>=560?'crit':avg>=520?'warn':'';
  const dC = delta>=200?'crit':delta>=100?'warn':'';
  el.innerHTML = `
    <div class="stat ${avgC}"><div class="sv">${Math.round(avg)}°C</div><div class="sl">Avg банка</div></div>
    <div class="stat ${avgC}"><div class="sv">${Math.round(maxVal)}°C</div><div class="sl">Max</div></div>
    <div class="stat ${dC}"><div class="sv">${Math.round(delta)}°C</div><div class="sl">∆ внутри банка</div></div>
  `;
}

// ── MAIN RENDER ─────────────────────────────────────────────────────────────
function render(key){
  const d = TRUCKS[key];
  if(!d){ return; }

  const ts = d.ts;
  const cyls = d.cyls;
  const secs = timeSecs(ts);

  // Header date
  document.getElementById('hdr-date').textContent =
    (d.model||'') + ' №' + (d.truck||'') + ' | ' + (d.date||ts[0]?.slice(0,10)||'');

  // Available banks
  const aCyls = A_CYLS.filter(n=>cyls[n]);
  const bCyls = B_CYLS.filter(n=>cyls[n]);

  // Bank averages (per time step)
  const aAvgSeries = aCyls.length ? rowMeans(cyls, aCyls) : [];
  const bAvgSeries = bCyls.length ? rowMeans(cyls, bCyls) : [];

  // ── Chart 1: A-bank ────────────────────────────────────────────────────
  if(aCyls.length){
    const traces = aCyls.map((n,i)=>({
      type:'scattergl', x:secs, y:cyls[n],
      name:'Ц'+n, line:{color:COLORS_A[i%COLORS_A.length],width:1},
      opacity:0.75, showlegend:true,
    }));
    traces.push({type:'scattergl',x:secs,y:ewma(aAvgSeries),name:'avg А',
      line:{color:'#fff',width:2.5,dash:'dot'},showlegend:true});
    pl('ch-a-bank', traces, {
      height:270,
      shapes:[
        {type:'line',x0:0,x1:1,xref:'paper',y0:500,y1:500,line:{color:'rgba(255,184,48,0.5)',dash:'dot',width:1.5}},
        {type:'line',x0:0,x1:1,xref:'paper',y0:550,y1:550,line:{color:'rgba(255,64,96,0.5)',dash:'dot',width:1.5}},
      ],
      annotations:[
        {x:0.01,xref:'paper',y:500,text:'500°C',showarrow:false,font:{color:'rgba(255,184,48,0.7)',size:9},xanchor:'left'},
        {x:0.01,xref:'paper',y:550,text:'550°C',showarrow:false,font:{color:'rgba(255,64,96,0.7)',size:9},xanchor:'left'},
      ],
      yaxis:{title:'EGT °C',range:[300,650]},
      xaxis:{title:'Время, сек'},
    });
    const aAvgMean  = mean(aAvgSeries);
    const aAllVals  = aCyls.flatMap(n=>nn(cyls[n]));
    const aDelta    = (maxv(aAllVals)||0) - (minv(aAllVals)||0);
    renderStats('stats-a', aAvgMean||0, maxv(aAllVals)||0, aDelta);
  }

  // ── Chart 2: B-bank ────────────────────────────────────────────────────
  if(bCyls.length){
    const traces = bCyls.map((n,i)=>({
      type:'scattergl', x:secs, y:cyls[n],
      name:'Ц'+n, line:{color:COLORS_B[i%COLORS_B.length],width:1},
      opacity:0.75, showlegend:true,
    }));
    traces.push({type:'scattergl',x:secs,y:ewma(bAvgSeries),name:'avg В',
      line:{color:'#fff',width:2.5,dash:'dot'},showlegend:true});
    pl('ch-b-bank', traces, {
      height:270,
      shapes:[
        {type:'line',x0:0,x1:1,xref:'paper',y0:500,y1:500,line:{color:'rgba(255,184,48,0.5)',dash:'dot',width:1.5}},
        {type:'line',x0:0,x1:1,xref:'paper',y0:550,y1:550,line:{color:'rgba(255,64,96,0.5)',dash:'dot',width:1.5}},
      ],
      annotations:[
        {x:0.01,xref:'paper',y:500,text:'500°C',showarrow:false,font:{color:'rgba(255,184,48,0.7)',size:9},xanchor:'left'},
        {x:0.01,xref:'paper',y:550,text:'550°C',showarrow:false,font:{color:'rgba(255,64,96,0.7)',size:9},xanchor:'left'},
      ],
      yaxis:{title:'EGT °C',range:[300,650]},
      xaxis:{title:'Время, сек'},
    });
    const bAvgMean  = mean(bAvgSeries);
    const bAllVals  = bCyls.flatMap(n=>nn(cyls[n]));
    const bDelta    = (maxv(bAllVals)||0) - (minv(bAllVals)||0);
    renderStats('stats-b', bAvgMean||0, maxv(bAllVals)||0, bDelta);
  }

  // ── Chart 3: Per-cylinder deviation from bank avg ──────────────────────
  {
    const allCyls = [...aCyls, ...bCyls].sort((a,b)=>a-b);
    const aAvgMean = mean(aAvgSeries)||0;
    const bAvgMean = mean(bAvgSeries)||0;
    const deviations = allCyls.map(n=>{
      const v = mean(cyls[n])||0;
      const bankAvg = A_CYLS.includes(n) ? aAvgMean : bAvgMean;
      return {n, dev: Math.round(v - bankAvg), bank: A_CYLS.includes(n)?'A':'B', avg:v};
    });

    pl('ch-deviation', [{
      type:'bar',
      x: deviations.map(d=>'Ц'+d.n),
      y: deviations.map(d=>d.dev),
      text: deviations.map(d=>(d.dev>=0?'+':'')+d.dev+'°'),
      textposition:'outside',
      textfont:{size:10},
      marker:{
        color: deviations.map(d=>{
          const ab=Math.abs(d.dev);
          const base=d.bank==='A'?'255,124,58':'91,143,255';
          const op=ab>60?'0.95':ab>30?'0.75':'0.55';
          return `rgba(${base},${op})`;
        }),
        line:{color:deviations.map(d=>d.bank==='A'?'#ff7c3a':'#5b8fff'),width:1.5},
      },
      hovertemplate:'Ц%{x}: %{y:+d}°C от avg банка<extra></extra>',
    }], {
      height:200,
      shapes:[
        {type:'line',x0:-0.5,x1:15.5,y0:30,y1:30,line:{color:'rgba(255,184,48,0.5)',dash:'dot',width:1.2}},
        {type:'line',x0:-0.5,x1:15.5,y0:-30,y1:-30,line:{color:'rgba(255,184,48,0.5)',dash:'dot',width:1.2}},
        {type:'line',x0:-0.5,x1:15.5,y0:60,y1:60,line:{color:'rgba(255,64,96,0.5)',dash:'dot',width:1.2}},
        {type:'line',x0:-0.5,x1:15.5,y0:-60,y1:-60,line:{color:'rgba(255,64,96,0.5)',dash:'dot',width:1.2}},
      ],
      annotations:[
        {x:15.5,y:30,text:'±30°',showarrow:false,font:{color:'rgba(255,184,48,0.7)',size:9},xanchor:'left'},
        {x:15.5,y:60,text:'±60°',showarrow:false,font:{color:'rgba(255,64,96,0.7)',size:9},xanchor:'left'},
      ],
      xaxis:{tickfont:{size:10}},
      yaxis:{title:'∆°C от avg банка',zeroline:true,zerolinecolor:'rgba(255,255,255,0.2)',zerolinewidth:1.5},
      margin:{l:48,r:40,t:14,b:36},
    });
  }

  // ── Chart 4: Interbank delta ───────────────────────────────────────────
  if(aAvgSeries.length && bAvgSeries.length){
    const delta = secs.map((_,i)=>{
      const a=aAvgSeries[i], b=bAvgSeries[i];
      return (a!==null&&b!==null) ? Math.round(b-a) : null;
    });
    const deltaMean = mean(delta);
    const deltaSign = deltaMean!==null ? (deltaMean>0?'В':'А') : '–';
    const posDelta = delta.map(v=>v!==null&&v>0?v:0);
    const negDelta = delta.map(v=>v!==null&&v<0?v:0);

    pl('ch-delta', [
      {type:'scattergl',x:secs,y:posDelta,name:'В-банк горячее',fill:'tozeroy',
        fillcolor:'rgba(91,143,255,0.18)',line:{color:'rgba(91,143,255,0.7)',width:1.5}},
      {type:'scattergl',x:secs,y:negDelta,name:'А-банк горячее',fill:'tozeroy',
        fillcolor:'rgba(255,124,58,0.18)',line:{color:'rgba(255,124,58,0.7)',width:1.5}},
      {type:'scattergl',x:secs,y:ewma(delta,0.05),name:'EWMA ∆',
        line:{color:'#ffb830',width:2.5},showlegend:true},
    ], {
      height:200,
      shapes:[
        {type:'line',x0:0,x1:1,xref:'paper',y0:20,y1:20,line:{color:'rgba(91,143,255,0.4)',dash:'dot',width:1.2}},
        {type:'line',x0:0,x1:1,xref:'paper',y0:-20,y1:-20,line:{color:'rgba(255,124,58,0.4)',dash:'dot',width:1.2}},
      ],
      annotations:[
        {x:0.02,xref:'paper',y:20,text:'+20°',showarrow:false,font:{color:'rgba(91,143,255,0.7)',size:9},xanchor:'left'},
        {x:0.02,xref:'paper',y:-20,text:'-20°',showarrow:false,font:{color:'rgba(255,124,58,0.7)',size:9},xanchor:'left'},
      ],
      yaxis:{title:'∆°C (В–А)',zeroline:true,zerolinecolor:'rgba(255,255,255,0.25)',zerolinewidth:1.5},
      xaxis:{title:'Время, сек'},
    });

    // ── Conclusion ─────────────────────────────────────────────────────
    const aAllVals  = aCyls.flatMap(n=>nn(cyls[n]));
    const bAllVals  = bCyls.flatMap(n=>nn(cyls[n]));
    const aDelta    = (maxv(aAllVals)||0)-(minv(aAllVals)||0);
    const bDelta    = (maxv(bAllVals)||0)-(minv(bAllVals)||0);
    const aAvgMean  = mean(aAvgSeries)||0;
    const bAvgMean  = mean(bAvgSeries)||0;

    // Find max-deviation cylinder
    const allCyls = [...aCyls, ...bCyls];
    let maxDevCyl=null, maxDevVal=0;
    allCyls.forEach(n=>{
      const v=mean(cyls[n])||0;
      const bankAvg=A_CYLS.includes(n)?aAvgMean:bAvgMean;
      const dev=Math.abs(v-bankAvg);
      if(dev>maxDevVal){maxDevVal=dev;maxDevCyl={n,dev:Math.round(v-bankAvg)};}
    });

    const hotter = deltaMean!==null ? (deltaMean>0?'В-банк горячее А-банка':'А-банк горячее В-банка') : '';
    const hotterAmt = deltaMean!==null ? Math.abs(Math.round(deltaMean)) : 0;
    const maxDevStr = maxDevCyl ? `Цил.<b>${maxDevCyl.n}</b> отклонение <b class="${Math.abs(maxDevCyl.dev)>60?'c-pill c-a':'c-pill c-ok'}">${maxDevCyl.dev>0?'+':''}${maxDevCyl.dev}°</b> от avg банка` : '';

    document.getElementById('conclusion').innerHTML = `
      <strong>ВЫВОД:</strong>
      <span class="c-pill ${deltaMean>=0?'c-b':'c-a'}">${hotter}</span> на <b>${hotterAmt}°С</b> &nbsp;|&nbsp;
      Разброс: <span class="c-pill c-a">A=${Math.round(aDelta)}°С</span> <span class="c-pill c-b">B=${Math.round(bDelta)}°С</span> &nbsp;|&nbsp;
      ${maxDevStr}
    `;
  }
}

// ── INIT ────────────────────────────────────────────────────────────────────
window.addEventListener('load', function(){
  buildPills();
  if(activeTruck) render(activeTruck);
});
</script>
</body>
</html>
"""


def build_html(db_data, truck_order):
    plotly_src = PLOTLY_JS.read_text(encoding='utf-8')
    data_json  = json.dumps(db_data,     ensure_ascii=False, separators=(',', ':'))
    order_json = json.dumps(truck_order, ensure_ascii=False)

    html = HTML_TEMPLATE
    html = html.replace('__PLOTLY__', plotly_src)
    html = html.replace('__DATA__',   data_json)
    html = html.replace('__ORDER__',  order_json)
    return html


def main():
    print('=== QSK50 Cylinder Temperature Dashboard Builder ===\n')
    db_data, truck_order = build_data()
    if not db_data:
        print('ERROR: No cylinder data found!')
        return

    total_rows = sum(len(v['ts']) for v in db_data.values())
    data_kb = len(json.dumps(db_data, separators=(',', ':')).encode()) // 1024
    print(f'\n{len(db_data)} trucks, {total_rows} rows, data JSON: {data_kb} KB')

    print('Building HTML …')
    html = build_html(db_data, truck_order)
    OUTPUT.write_text(html, encoding='utf-8')
    size_mb = OUTPUT.stat().st_size / 1024 / 1024
    print(f'Written: {OUTPUT}  ({size_mb:.1f} MB)')


if __name__ == '__main__':
    main()
