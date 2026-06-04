#!/usr/bin/env python3
"""
Build qsk50_cylinders.html v2 — cylinder EGT + pressure analysis.
Features: sidebar truck list, date selector, file loading, comparison tab, pressure tab.
"""

import io, json, re, pathlib, warnings
import pandas as pd

warnings.filterwarnings('ignore')

EXTRACT   = pathlib.Path('/home/user/NTE200/.qsk50_extracted')
PLOTLY_JS = pathlib.Path('/usr/local/lib/python3.11/dist-packages/plotly/package_data/plotly.min.js')
OUTPUT    = pathlib.Path('/home/user/NTE200/qsk50_cylinders.html')

MAX_ROWS = 600

RU_MONTHS = {
    'янв':'Jan','фев':'Feb','мар':'Mar','апр':'Apr','май':'May','июн':'Jun',
    'июл':'Jul','авг':'Aug','сен':'Sep','окт':'Oct','ноя':'Nov','дек':'Dec',
}
A_BANK = [1,3,5,7,9,11,13,15]
B_BANK = [2,4,6,8,10,12,14,16]

PRESS_PATTERNS = {
    'oil':    ['Давление масла (кПа)', 'Engine Oil Pressure (kPa)'],
    'crank':  ['Давление картерных газов', 'Crankcase Pressure (kPa)'],
    'boost':  ['Давление во впускном коллекторе (кПа)- Идентификатор0',
               'Intake Manifold Pressure (kPa)- Identifier0'],
    'oil_t':  ['Температура масла (°C)', 'Engine Oil Temperature (°C)'],
    'cool_t': ['Датчик температуры (°C)- Идентификатор0',
               'Engine Coolant Temperature (°C)'],
    'rpm':    ['Частота вращения двигателя (об/мин)', 'Engine Speed (RPM)- Identifier0'],
    'load':   ['Относительная нагрузка (Проценты)- Идентификатор0',
               'Percent Load (percent)- Identifier0'],
}


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


def find_col(cols, *patterns):
    for pat in patterns:
        for c in cols:
            if pat in c:
                return c
    return None


def find_cyl_col(cols, n):
    for c in cols:
        if f'отработавших газов цилиндра {n} (°C)' in c:
            return c
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

    all_cols = list(df.columns)
    date_col = 'Date' if 'Date' in all_cols else 'Дата'
    time_col = 'Time' if 'Time' in all_cols else 'Время'

    cyl_cols = {}
    for n in range(1, 17):
        c = find_cyl_col(all_cols, n)
        if c:
            cyl_cols[n] = c

    if len(cyl_cols) < 8:
        return None

    press_cols = {}
    for key, patterns in PRESS_PATTERNS.items():
        c = find_col(all_cols, *patterns)
        if c:
            press_cols[key] = c

    keep = [date_col, time_col] + list(cyl_cols.values()) + list(press_cols.values())
    df_sub = df[keep].copy()

    for col in keep[2:]:
        try:
            conv = df_sub[col].astype(str).str.replace(',', '.', regex=False).str.strip()
            num  = pd.to_numeric(conv, errors='coerce')
            if num.notna().sum() > len(df_sub) * 0.2:
                df_sub[col] = num.round(1)
        except Exception:
            pass

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

    step = max(1, len(df_sub) // MAX_ROWS)
    df_sub = df_sub.iloc[::step].reset_index(drop=True)

    ts = df_sub['_ts'].tolist()

    cyls = {}
    for n, col in cyl_cols.items():
        s = df_sub[col]
        cyls[n] = [None if pd.isna(v) else int(v) for v in s]

    press = {}
    for key, col in press_cols.items():
        s = df_sub[col]
        press[key] = [None if pd.isna(v) else round(float(v), 1) for v in s]

    # Date display string
    raw_date = ts[0][:10] if ts else ''
    parts = raw_date.split('-')
    date_str = f"{parts[2]}.{parts[1]}.{parts[0]}" if len(parts) == 3 else raw_date

    return {
        'ts':    ts,
        'cyls':  cyls,
        'press': press,
        'date':  date_str,
        'n_cyls': len(cyl_cols),
        'n_press': len(press_cols),
    }


def truck_num(fname):
    for pat in [r'№\s*(\d+)', r'\bN(\d+)\b', r'\((\d+)\)', r'[\s_-](\d{2,3})[\s_.( ]']:
        m = re.search(pat, fname)
        if m:
            return int(m.group(1))
    m = re.search(r'DML-\d+-\d+\s+(\d+)', fname)
    if m:
        return int(m.group(1))
    return None


def build_data():
    # Structure: TRUCKS[key] = {truck, model, sessions:[{date,ts,cyls,press},...]}
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
            print("skip")
            continue
        print(f"{len(result['ts'])} rows, {result['n_cyls']} cyls, {result['n_press']} press")
        session = {
            'date':  result['date'],
            'ts':    result['ts'],
            'cyls':  result['cyls'],
            'press': result['press'],
            'file':  name,
        }
        if key not in db:
            db[key] = {'truck': truck, 'model': model, 'sessions': [session]}
            order.append(key)
        else:
            # Add as new session (don't merge — keep dates separate)
            db[key]['sessions'].append(session)
    return db, order


# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>QSK50 — Цилиндры и давление</title>
<script>__PLOTLY__</script>
<style>
:root{
  --bg:#0b1622;--bg2:#0f1e2e;--sb:#0d1926;--card:#112030;--brd:rgba(60,120,180,0.14);
  --tx:#c8dff0;--tx2:#6a9ab0;--tx3:#334d60;
  --acc:#00d4aa;--red:#ff4060;--amb:#ffb830;--grn:#2ecc71;
  --blu:#5b8fff;--pur:#b06bff;--org:#ff7c3a;
  --a-col:#ff7c3a;--b-col:#5b8fff;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--tx);font-size:13px;display:grid;grid-template-columns:230px 1fr;grid-template-rows:48px 1fr;height:100vh}

/* ── HEADER ── */
.hdr{grid-column:1/-1;grid-row:1;display:flex;align-items:center;gap:12px;padding:0 16px;background:rgba(6,12,22,0.98);border-bottom:1px solid var(--brd);z-index:10}
.logo{font-weight:800;font-size:14px;color:#fff;letter-spacing:-.3px}
.logo span{color:var(--tx3);font-weight:400;font-size:10px;margin-left:8px}
.hdr-info{font-size:11px;color:var(--amb);font-family:'SF Mono',monospace;margin-left:auto}

/* ── SIDEBAR ── */
.sidebar{grid-column:1;grid-row:2;background:var(--sb);border-right:1px solid var(--brd);display:flex;flex-direction:column;overflow:hidden}
.sb-section{padding:10px 12px;border-bottom:1px solid var(--brd);flex-shrink:0}
.sb-label{font-size:9px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.7px;margin-bottom:7px;display:flex;align-items:center;justify-content:space-between}
.sb-label button{font-size:9px;color:var(--acc);background:none;border:1px solid rgba(0,212,170,0.3);border-radius:4px;padding:1px 6px;cursor:pointer}
.sb-label button:hover{background:rgba(0,212,170,0.08)}

.truck-list{flex:1;overflow-y:auto;padding:6px 0}
.truck-list::-webkit-scrollbar{width:3px}
.truck-list::-webkit-scrollbar-thumb{background:var(--brd);border-radius:2px}
.truck-item{display:flex;align-items:center;gap:8px;padding:7px 12px;cursor:pointer;transition:background .12s;border-left:3px solid transparent;user-select:none;position:relative}
.truck-item:hover{background:rgba(255,255,255,0.04)}
.truck-item.active{background:rgba(0,212,170,0.07);border-left-color:var(--acc)}
.truck-item.cmp-on{background:rgba(91,143,255,0.07)}
.truck-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.truck-name{font-size:12px;font-weight:600;flex:1}
.truck-date{font-size:9px;color:var(--tx3);margin-top:1px}
.cmp-chk{width:14px;height:14px;accent-color:var(--blu);flex-shrink:0;cursor:pointer}

.date-pills{display:flex;flex-wrap:wrap;gap:5px}
.date-pill{padding:3px 9px;border:1px solid var(--brd);border-radius:12px;font-size:10px;color:var(--tx2);cursor:pointer;transition:all .12s}
.date-pill:hover{border-color:var(--acc);color:var(--acc)}
.date-pill.active{background:rgba(0,212,170,0.12);border-color:var(--acc);color:var(--acc);font-weight:600}

.file-zone{border:1px dashed rgba(60,120,180,0.35);border-radius:7px;padding:8px;text-align:center;cursor:pointer;transition:all .15s;position:relative}
.file-zone:hover{border-color:var(--acc);background:rgba(0,212,170,0.04)}
.file-zone input{position:absolute;inset:0;opacity:0;width:100%;height:100%;cursor:pointer}
.file-zone-icon{font-size:18px;margin-bottom:3px}
.file-zone-txt{font-size:10px;color:var(--tx3);line-height:1.4}
.file-status{font-size:9px;color:var(--grn);margin-top:4px;font-weight:600}

/* ── MAIN AREA ── */
.main{grid-column:2;grid-row:2;display:flex;flex-direction:column;overflow:hidden}
.tabs{display:flex;border-bottom:1px solid var(--brd);background:var(--bg2);flex-shrink:0}
.tbtn{padding:10px 16px;border:none;border-bottom:2px solid transparent;background:none;color:var(--tx3);cursor:pointer;font-size:11px;font-weight:500;white-space:nowrap;transition:all .12s}
.tbtn:hover{color:var(--tx);background:rgba(255,255,255,0.03)}
.tbtn.act{color:var(--acc);border-bottom-color:var(--acc);background:rgba(0,212,170,0.04)}
.content{flex:1;overflow-y:auto;overflow-x:hidden;padding:12px}
.content::-webkit-scrollbar{width:5px}
.content::-webkit-scrollbar-thumb{background:var(--brd);border-radius:3px}

.panel{display:none;flex-direction:column;gap:12px}
.panel.show{display:flex}

/* ── CARDS / LAYOUT ── */
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.card{background:var(--card);border:1px solid var(--brd);border-radius:10px;padding:10px;overflow:hidden}
.card-title{font-size:10px;font-weight:700;color:var(--tx2);margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:8px}
.tag{padding:2px 7px;border-radius:4px;font-size:9px;font-weight:700;letter-spacing:.3px}
.tag-a{background:rgba(255,124,58,.2);color:var(--org)}
.tag-b{background:rgba(91,143,255,.2);color:var(--blu)}

.stats-row{display:flex;gap:14px;padding:6px 10px;background:rgba(255,255,255,.03);border-radius:6px;border:1px solid var(--brd);margin-top:6px;flex-wrap:wrap}
.stat{text-align:center}
.stat .sv{font-size:15px;font-weight:700;font-family:'SF Mono',monospace;color:var(--acc)}
.stat .sl{font-size:9px;color:var(--tx3);text-transform:uppercase;letter-spacing:.4px;margin-top:1px}
.stat.warn .sv{color:var(--amb)} .stat.crit .sv{color:var(--red)}

.conclusion{background:linear-gradient(90deg,rgba(255,124,58,.07),rgba(91,143,255,.07));border:1px solid rgba(255,124,58,.2);border-radius:8px;padding:10px 16px;font-size:11px;line-height:1.9}
.conclusion strong{color:#fff;font-size:12px}
.cp{display:inline-block;padding:1px 7px;border-radius:4px;font-weight:700;font-size:11px;margin:0 2px}
.cp-a{background:rgba(255,124,58,.2);color:var(--org)} .cp-b{background:rgba(91,143,255,.2);color:var(--blu)}
.cp-ok{background:rgba(46,204,113,.2);color:var(--grn)} .cp-w{background:rgba(255,184,48,.2);color:var(--amb)} .cp-r{background:rgba(255,64,96,.2);color:var(--red)}

/* empty state */
.empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:60px;color:var(--tx3);text-align:center}
.empty .ei{font-size:48px;opacity:.2}
.empty p{font-size:12px}

/* Сравнение tab */
.cmp-legend{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:6px}
.cmp-leg{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--tx2)}
.cmp-dot{width:10px;height:10px;border-radius:50%}

/* Давление tab */
.press-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.press-wide{grid-column:1/-1}
.thres-legend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:6px}
.tl-item{display:flex;align-items:center;gap:5px;font-size:9px;color:var(--tx3)}
.tl-box{width:22px;height:4px;border-radius:2px}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="logo">⚙ QSK50 <span>Цилиндры · Давление · Сравнение — Cummins MCRS</span></div>
  <div class="hdr-info" id="hdr-info">–</div>
</div>

<!-- SIDEBAR -->
<div class="sidebar">
  <!-- Truck list header -->
  <div class="sb-section">
    <div class="sb-label">
      🚛 Самосвалы
      <button onclick="toggleCmpMode()">⊞ Сравнение</button>
    </div>
  </div>
  <div class="truck-list" id="truck-list"></div>

  <!-- Date selector -->
  <div class="sb-section" id="date-section" style="display:none">
    <div class="sb-label">📅 Дата</div>
    <div class="date-pills" id="date-pills"></div>
  </div>

  <!-- File loader -->
  <div class="sb-section">
    <div class="sb-label">📂 Добавить файлы</div>
    <div class="file-zone" onclick="document.getElementById('file-in').click()">
      <div class="file-zone-icon">📊</div>
      <div class="file-zone-txt">Log Files CSV<br>Один или несколько файлов</div>
      <div class="file-status" id="file-status"></div>
      <input type="file" id="file-in" accept=".csv" multiple style="display:none" onchange="loadFiles(this.files)">
    </div>
  </div>
</div>

<!-- MAIN -->
<div class="main">
  <div class="tabs">
    <button class="tbtn act" onclick="showTab('cylinders',this)">🔩 Цилиндры</button>
    <button class="tbtn" onclick="showTab('compare',this)">📊 Сравнение</button>
    <button class="tbtn" onclick="showTab('pressure',this)">💧 Давление</button>
  </div>
  <div class="content">

    <!-- TAB: CYLINDERS -->
    <div id="tab-cylinders" class="panel show">
      <div class="row2">
        <div class="card">
          <div class="card-title"><span class="tag tag-a">А-банк</span>Нечётные цилиндры — ECM144</div>
          <div id="ch-a"></div>
          <div class="stats-row" id="stats-a"></div>
        </div>
        <div class="card">
          <div class="card-title"><span class="tag tag-b">В-банк</span>Чётные цилиндры — ECM1</div>
          <div id="ch-b"></div>
          <div class="stats-row" id="stats-b"></div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">∆ каждого цилиндра от среднего по своему банку</div>
        <div id="ch-dev"></div>
      </div>
      <div class="card">
        <div class="card-title">Межбанковая дельта: В-банк avg − А-банк avg (по времени)</div>
        <div id="ch-delta"></div>
      </div>
      <div class="conclusion" id="conclusion">← Выберите самосвал слева</div>
    </div>

    <!-- TAB: COMPARE -->
    <div id="tab-compare" class="panel">
      <div id="cmp-empty" class="empty" style="display:none">
        <div class="ei">📊</div>
        <p>Отметьте ✓ несколько самосвалов в панели слева для сравнения</p>
      </div>
      <div id="cmp-legend-box" class="cmp-legend"></div>
      <div class="card">
        <div class="card-title">Средняя температура по цилиндрам (по выбранным самосвалам)</div>
        <div id="ch-cmp-cyls"></div>
      </div>
      <div class="row2">
        <div class="card">
          <div class="card-title">А-банк avg vs В-банк avg по самосвалам</div>
          <div id="ch-cmp-banks"></div>
        </div>
        <div class="card">
          <div class="card-title">Межбанковая дельта (В–А) по самосвалам</div>
          <div id="ch-cmp-delta"></div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Сводная таблица сравнения</div>
        <div id="ch-cmp-table"></div>
      </div>
    </div>

    <!-- TAB: PRESSURE -->
    <div id="tab-pressure" class="panel">
      <div class="press-grid">
        <div class="card press-wide">
          <div class="card-title">Давление масла (кПа)</div>
          <div class="thres-legend">
            <div class="tl-item"><div class="tl-box" style="background:rgba(46,204,113,.5)"></div>Норма 350–650</div>
            <div class="tl-item"><div class="tl-box" style="background:rgba(255,184,48,.5)"></div>Внимание &lt;350</div>
            <div class="tl-item"><div class="tl-box" style="background:rgba(255,64,96,.5)"></div>Критично &lt;300</div>
          </div>
          <div id="ch-p-oil"></div>
        </div>
        <div class="card">
          <div class="card-title">Давление картерных газов (кПа)</div>
          <div class="thres-legend">
            <div class="tl-item"><div class="tl-box" style="background:rgba(46,204,113,.5)"></div>Норма 0–3</div>
            <div class="tl-item"><div class="tl-box" style="background:rgba(255,184,48,.5)"></div>Внимание 3–7</div>
            <div class="tl-item"><div class="tl-box" style="background:rgba(255,64,96,.5)"></div>Критично &gt;7</div>
          </div>
          <div id="ch-p-crank"></div>
        </div>
        <div class="card">
          <div class="card-title">Давление наддува — впускной коллектор (кПа)</div>
          <div class="thres-legend">
            <div class="tl-item"><div class="tl-box" style="background:rgba(46,204,113,.5)"></div>Норма 200–350</div>
            <div class="tl-item"><div class="tl-box" style="background:rgba(255,184,48,.5)"></div>Внимание &lt;180</div>
          </div>
          <div id="ch-p-boost"></div>
        </div>
        <div class="card">
          <div class="card-title">T масла + T охл. жидкости (°C)</div>
          <div id="ch-p-temps"></div>
        </div>
        <div class="card press-wide">
          <div class="card-title">Обороты (об/мин) + Нагрузка (%)</div>
          <div id="ch-p-rpm"></div>
        </div>
      </div>
      <div class="conclusion" id="press-summary">← Выберите самосвал слева</div>
    </div>

  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════════════════
// PRE-LOADED DATA
// ═══════════════════════════════════════════════════════════════════
const TRUCKS = __DATA__;
const TRUCK_ORDER = __ORDER__;

const A_CYLS = [1,3,5,7,9,11,13,15];
const B_CYLS = [2,4,6,8,10,12,14,16];
const COLORS_A = ['#ff7c3a','#ffb830','#e55c3a','#cc4020','#ff9860','#f0c000','#dd7060','#aa3010'];
const COLORS_B = ['#5b8fff','#00bfff','#7066ee','#9955dd','#3399ff','#44aaff','#8877ff','#bb77ee'];
const TC = ['#00d4aa','#5b8fff','#ffb830','#ff4060','#b06bff','#00bfff','#2ecc71','#ff7c3a','#ff6eb4','#a3e635'];

const PLY_CFG = {responsive:true,displaylogo:false,modeBarButtonsToRemove:['select2d','lasso2d']};
const PLY = {
  paper_bgcolor:'#112030',plot_bgcolor:'rgba(8,18,32,0.7)',
  font:{color:'#c8dff0',family:'Segoe UI,sans-serif',size:10},
  margin:{l:48,r:14,t:26,b:36},
  xaxis:{gridcolor:'rgba(60,120,180,0.07)',zerolinecolor:'rgba(60,120,180,0.12)',color:'#334d60'},
  yaxis:{gridcolor:'rgba(60,120,180,0.07)',zerolinecolor:'rgba(60,120,180,0.12)',color:'#334d60'},
  legend:{bgcolor:'rgba(0,0,0,0)',borderwidth:0,font:{size:9}},
};
function pl(id,data,lay={}){
  const L=Object.assign({},PLY,lay);
  if(lay.xaxis)L.xaxis=Object.assign({},PLY.xaxis,lay.xaxis);
  if(lay.yaxis)L.yaxis=Object.assign({},PLY.yaxis,lay.yaxis);
  if(lay.yaxis2)L.yaxis2=Object.assign({showgrid:false,color:'#334d60'},lay.yaxis2);
  Plotly.react(id,data,L,PLY_CFG);
}

// ═══ STATE ═══════════════════════════════════════════════════════════
let activeTruck = null;   // truck key
let activeSession = {};   // truckKey → sessionIndex
let compareMode = false;
let compareSet = new Set();
let currentTab = 'cylinders';

// ═══ MATH ════════════════════════════════════════════════════════════
function nn(a){return(a||[]).filter(v=>v!==null&&!isNaN(v));}
function mean(a){const b=nn(a);return b.length?b.reduce((s,v)=>s+v,0)/b.length:null;}
function maxv(a){const b=nn(a);return b.length?Math.max(...b):null;}
function minv(a){const b=nn(a);return b.length?Math.min(...b):null;}
function rowMeans(cyls,keys){
  if(!cyls||!keys.length)return [];
  const n=(cyls[keys[0]]||[]).length;
  return Array.from({length:n},(_,i)=>{
    const vals=keys.map(k=>cyls[k]?cyls[k][i]:null).filter(v=>v!==null&&!isNaN(v));
    return vals.length?vals.reduce((s,v)=>s+v,0)/vals.length:null;
  });
}
function timeSecs(ts){
  if(!ts||!ts.length)return[];
  const t0=new Date(ts[0]).getTime();
  return ts.map(t=>t?Math.round((new Date(t).getTime()-t0)/1000):null);
}
function ewma(arr,a=0.06){let s=null;return arr.map(v=>{if(v!=null){s=s==null?v:a*v+(1-a)*s;return Math.round(s*10)/10;}return s;});}

// ═══ SESSION HELPERS ════════════════════════════════════════════════
function getSession(key){
  const d=TRUCKS[key];if(!d)return null;
  const si=activeSession[key]||0;
  return d.sessions[si]||d.sessions[0]||null;
}
function sessionLabel(key){
  const d=TRUCKS[key];if(!d)return key;
  return d.model+' №'+d.truck;
}
function sessionDateLabel(key){
  const s=getSession(key);return s?s.date:'?';
}

// ═══ SIDEBAR ════════════════════════════════════════════════════════
function buildSidebar(){
  const list=document.getElementById('truck-list');
  list.innerHTML='';
  TRUCK_ORDER.forEach((key,i)=>{
    const d=TRUCKS[key];if(!d)return;
    const item=document.createElement('div');
    item.className='truck-item'+(activeTruck===key?' active':'');
    item.id='ti-'+key;
    item.onclick=(e)=>{if(e.target.classList.contains('cmp-chk'))return;selectTruck(key);};

    const dot=document.createElement('span');
    dot.className='truck-dot';dot.style.background=TC[i%TC.length];
    const info=document.createElement('div');info.style.flex='1';info.style.minWidth=0;
    const nm=document.createElement('div');nm.className='truck-name';nm.textContent=d.model+' №'+d.truck;
    const dt=document.createElement('div');dt.className='truck-date';dt.id='td-'+key;
    dt.textContent=sessionDateLabel(key)+(d.sessions.length>1?' (+'+( d.sessions.length-1)+')':'');
    info.appendChild(nm);info.appendChild(dt);

    const chk=document.createElement('input');
    chk.type='checkbox';chk.className='cmp-chk';chk.id='chk-'+key;
    chk.style.display=compareMode?'':'none';
    chk.checked=compareSet.has(key);
    chk.onchange=()=>{if(chk.checked)compareSet.add(key);else compareSet.delete(key);if(currentTab==='compare')renderCompare();};

    item.appendChild(dot);item.appendChild(info);item.appendChild(chk);
    list.appendChild(item);
  });
}

function toggleCmpMode(){
  compareMode=!compareMode;
  document.querySelectorAll('.cmp-chk').forEach(c=>c.style.display=compareMode?'':'none');
  if(compareMode&&currentTab==='cylinders')showTab('compare',document.querySelectorAll('.tbtn')[1]);
}

function selectTruck(key){
  activeTruck=key;
  document.querySelectorAll('.truck-item').forEach(el=>el.classList.remove('active'));
  const el=document.getElementById('ti-'+key);if(el)el.classList.add('active');
  buildDatePills(key);
  renderCurrentTab();
  updateHeader();
}

function buildDatePills(key){
  const d=TRUCKS[key];
  const sec=document.getElementById('date-section');
  const pills=document.getElementById('date-pills');
  if(!d||d.sessions.length<=1){sec.style.display='none';return;}
  sec.style.display='';
  pills.innerHTML='';
  d.sessions.forEach((s,i)=>{
    const p=document.createElement('div');
    p.className='date-pill'+(i===(activeSession[key]||0)?' active':'');
    p.textContent=s.date;
    p.onclick=()=>{
      activeSession[key]=i;
      pills.querySelectorAll('.date-pill').forEach(x=>x.classList.remove('active'));
      p.classList.add('active');
      document.getElementById('td-'+key).textContent=s.date+(d.sessions.length>1?' (+'+( d.sessions.length-1)+')':'');
      renderCurrentTab();updateHeader();
    };
    pills.appendChild(p);
  });
}

function updateHeader(){
  if(!activeTruck)return;
  const s=getSession(activeTruck);
  const d=TRUCKS[activeTruck];
  document.getElementById('hdr-info').textContent=
    (d?d.model+' №'+d.truck:'')+(s?' | '+s.date:'');
}

// ═══ TAB SWITCHING ══════════════════════════════════════════════════
function showTab(name,btn){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('show'));
  document.querySelectorAll('.tbtn').forEach(b=>b.classList.remove('act'));
  document.getElementById('tab-'+name).classList.add('show');
  btn.classList.add('act');
  currentTab=name;
  renderCurrentTab();
}
function renderCurrentTab(){
  if(currentTab==='cylinders')renderCylinders();
  else if(currentTab==='compare')renderCompare();
  else if(currentTab==='pressure')renderPressure();
}

// ═══ CHART HELPERS ══════════════════════════════════════════════════
function refLines(vals,colors){
  return vals.map((v,i)=>({type:'line',x0:0,x1:1,xref:'paper',y0:v,y1:v,line:{color:colors[i]||'#888',dash:'dot',width:1.5}}));
}
function refAnnots(vals,labels,colors){
  return vals.map((v,i)=>({x:0.01,xref:'paper',y:v,text:labels[i],showarrow:false,font:{color:colors[i]||'#888',size:9},xanchor:'left'}));
}
function statsHTML(avg,max,delta){
  const ac=avg>=560?'crit':avg>=520?'warn':'';
  const dc=delta>=200?'crit':delta>=100?'warn':'';
  return `<div class="stat ${ac}"><div class="sv">${Math.round(avg)}°C</div><div class="sl">Avg банка</div></div>
          <div class="stat ${ac}"><div class="sv">${Math.round(max)}°C</div><div class="sl">Max</div></div>
          <div class="stat ${dc}"><div class="sv">${Math.round(delta)}°C</div><div class="sl">∆ внутри</div></div>`;
}

// ═══ TAB: CYLINDERS ═════════════════════════════════════════════════
function renderCylinders(){
  if(!activeTruck||!TRUCKS[activeTruck]){
    ['ch-a','ch-b','ch-dev','ch-delta'].forEach(id=>Plotly.purge(id));
    document.getElementById('conclusion').innerHTML='← Выберите самосвал слева';
    return;
  }
  const s=getSession(activeTruck);if(!s)return;
  const {ts,cyls}=s;
  const secs=timeSecs(ts);
  const aCyls=A_CYLS.filter(n=>cyls[n]);
  const bCyls=B_CYLS.filter(n=>cyls[n]);
  const aAvg=aCyls.length?rowMeans(cyls,aCyls):[];
  const bAvg=bCyls.length?rowMeans(cyls,bCyls):[];

  // A-bank chart
  if(aCyls.length){
    const tr=aCyls.map((n,i)=>({type:'scattergl',x:secs,y:cyls[n],name:'Ц'+n,line:{color:COLORS_A[i%8],width:1},opacity:.8}));
    tr.push({type:'scattergl',x:secs,y:ewma(aAvg),name:'avg А',line:{color:'#fff',width:2.5,dash:'dot'}});
    pl('ch-a',tr,{height:260,yaxis:{title:'EGT °C',range:[250,650]},xaxis:{title:'сек'},
      shapes:refLines([500,550],['rgba(255,184,48,.5)','rgba(255,64,96,.5)']),
      annotations:refAnnots([500,550],['500°C','550°C'],['rgba(255,184,48,.7)','rgba(255,64,96,.7)'])});
    const av=mean(aAvg)||0,al=aCyls.flatMap(n=>nn(cyls[n]));
    document.getElementById('stats-a').innerHTML=statsHTML(av,maxv(al)||0,(maxv(al)||0)-(minv(al)||0));
  }

  // B-bank chart
  if(bCyls.length){
    const tr=bCyls.map((n,i)=>({type:'scattergl',x:secs,y:cyls[n],name:'Ц'+n,line:{color:COLORS_B[i%8],width:1},opacity:.8}));
    tr.push({type:'scattergl',x:secs,y:ewma(bAvg),name:'avg В',line:{color:'#fff',width:2.5,dash:'dot'}});
    pl('ch-b',tr,{height:260,yaxis:{title:'EGT °C',range:[250,650]},xaxis:{title:'сек'},
      shapes:refLines([500,550],['rgba(255,184,48,.5)','rgba(255,64,96,.5)']),
      annotations:refAnnots([500,550],['500°C','550°C'],['rgba(255,184,48,.7)','rgba(255,64,96,.7)'])});
    const bv=mean(bAvg)||0,bl=bCyls.flatMap(n=>nn(cyls[n]));
    document.getElementById('stats-b').innerHTML=statsHTML(bv,maxv(bl)||0,(maxv(bl)||0)-(minv(bl)||0));
  }

  // Deviation bar
  const aAvgM=mean(aAvg)||0,bAvgM=mean(bAvg)||0;
  const allC=[...aCyls,...bCyls].sort((a,b)=>a-b);
  const devs=allC.map(n=>{
    const v=mean(cyls[n])||0;
    const bk=A_CYLS.includes(n)?aAvgM:bAvgM;
    return{n,d:Math.round(v-bk),bank:A_CYLS.includes(n)?'A':'B'};
  });
  pl('ch-dev',[{type:'bar',x:devs.map(d=>'Ц'+d.n),y:devs.map(d=>d.d),
    text:devs.map(d=>(d.d>=0?'+':'')+d.d+'°'),textposition:'outside',textfont:{size:9},
    marker:{color:devs.map(d=>{const ab=Math.abs(d.d),base=d.bank==='A'?'255,124,58':'91,143,255',op=ab>60?.95:ab>30?.7:.5;return `rgba(${base},${op})`;})},
  }],{height:190,
    shapes:[...refLines([30,-30,60,-60],['rgba(255,184,48,.4)','rgba(255,184,48,.4)','rgba(255,64,96,.4)','rgba(255,64,96,.4)'])],
    yaxis:{title:'∆°C',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5},
    xaxis:{tickfont:{size:9}},margin:{l:48,r:40,t:14,b:36}});

  // Delta time series
  const delta=secs.map((_,i)=>{
    const a=aAvg[i],b=bAvg[i];return(a!=null&&b!=null)?Math.round(b-a):null;});
  const dm=mean(delta);
  pl('ch-delta',[
    {type:'scattergl',x:secs,y:delta.map(v=>v!=null&&v>0?v:0),name:'В>А',fill:'tozeroy',fillcolor:'rgba(91,143,255,.15)',line:{color:'rgba(91,143,255,.6)',width:1}},
    {type:'scattergl',x:secs,y:delta.map(v=>v!=null&&v<0?v:0),name:'А>В',fill:'tozeroy',fillcolor:'rgba(255,124,58,.15)',line:{color:'rgba(255,124,58,.6)',width:1}},
    {type:'scattergl',x:secs,y:ewma(delta,.05),name:'EWMA',line:{color:'#ffb830',width:2.5}},
  ],{height:190,yaxis:{title:'∆°C (В−А)',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5},xaxis:{title:'сек'},
    shapes:refLines([20,-20],['rgba(91,143,255,.35)','rgba(255,124,58,.35)']),
    annotations:refAnnots([20,-20],['+20°','-20°'],['rgba(91,143,255,.6)','rgba(255,124,58,.6)'])});

  // Conclusion
  const hotBank=dm!=null?(dm>0?'В-банк':'А-банк'):'?',hotAmt=dm!=null?Math.abs(Math.round(dm)):0;
  const aSpread=Math.round((maxv(aCyls.flatMap(n=>nn(cyls[n])))||0)-(minv(aCyls.flatMap(n=>nn(cyls[n])))||0));
  const bSpread=Math.round((maxv(bCyls.flatMap(n=>nn(cyls[n])))||0)-(minv(bCyls.flatMap(n=>nn(cyls[n])))||0));
  let maxDevC=null,maxDevV=0;
  allC.forEach(n=>{const v=mean(cyls[n])||0;const bk=A_CYLS.includes(n)?aAvgM:bAvgM;const dv=Math.abs(v-bk);if(dv>maxDevV){maxDevV=dv;maxDevC={n,d:Math.round(v-bk)};}});
  const riskC=maxDevC&&Math.abs(maxDevC.d)>60?'cp-r':maxDevC&&Math.abs(maxDevC.d)>30?'cp-w':'cp-ok';
  document.getElementById('conclusion').innerHTML=`
    <strong>ВЫВОД:</strong>
    <span class="cp ${dm>=0?'cp-b':'cp-a'}">${hotBank} горячее</span> на <b>${hotAmt}°С</b> &nbsp;|&nbsp;
    Разброс: <span class="cp cp-a">A=${aSpread}°</span><span class="cp cp-b">B=${bSpread}°</span> &nbsp;|&nbsp;
    ${maxDevC?`Цил.<b>${maxDevC.n}</b>: <span class="cp ${riskC}">${maxDevC.d>0?'+':''}${maxDevC.d}°</span> от avg банка`:''}`;
}

// ═══ TAB: COMPARE ═══════════════════════════════════════════════════
function renderCompare(){
  const keys=[...compareSet].filter(k=>TRUCKS[k]);
  const empty=document.getElementById('cmp-empty');
  if(keys.length<2){
    empty.style.display='';
    ['ch-cmp-cyls','ch-cmp-banks','ch-cmp-delta','ch-cmp-table'].forEach(id=>Plotly.purge(id));
    document.getElementById('cmp-legend-box').innerHTML='';
    return;
  }
  empty.style.display='none';

  // Legend
  const legBox=document.getElementById('cmp-legend-box');
  legBox.innerHTML=keys.map((k,i)=>`<div class="cmp-leg"><div class="cmp-dot" style="background:${TC[i%TC.length]}"></div>${sessionLabel(k)} — ${sessionDateLabel(k)}</div>`).join('');

  // Per-cylinder avg for each selected truck
  const allCyls=[...A_CYLS,...B_CYLS].sort((a,b)=>a-b);
  const trCylAvgs=keys.map(k=>{
    const s=getSession(k);if(!s)return{};
    const out={};allCyls.forEach(n=>{out[n]=mean(s.cyls[n]);});
    return out;
  });

  // Chart: grouped bar per cylinder
  const cylTraces=keys.map((k,i)=>({
    type:'bar',name:sessionLabel(k)+' '+sessionDateLabel(k),
    x:allCyls.map(n=>'Ц'+n),y:allCyls.map(n=>trCylAvgs[i][n]!=null?Math.round(trCylAvgs[i][n]):null),
    marker:{color:TC[i%TC.length]},
  }));
  pl('ch-cmp-cyls',cylTraces,{height:280,barmode:'group',
    yaxis:{title:'°C'},xaxis:{tickfont:{size:9}},
    shapes:refLines([500,550],['rgba(255,184,48,.35)','rgba(255,64,96,.35)'])});

  // Bank comparison
  const bankTraces=[{
    type:'bar',name:'А-банк avg',
    x:keys.map(k=>sessionLabel(k)),
    y:keys.map((k,i)=>Math.round(mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),
    marker:{color:TC.slice(0,keys.length).map(c=>c+'cc')},offsetgroup:0,
  },{
    type:'bar',name:'В-банк avg',
    x:keys.map(k=>sessionLabel(k)),
    y:keys.map((k,i)=>Math.round(mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),
    marker:{color:TC.slice(0,keys.length)},offsetgroup:1,
  }];
  pl('ch-cmp-banks',bankTraces,{height:220,barmode:'group',yaxis:{title:'°C'},
    shapes:refLines([500,550],['rgba(255,184,48,.35)','rgba(255,64,96,.35)'])});

  // Interbank delta bar
  const idTraces=[{
    type:'bar',
    x:keys.map(k=>sessionLabel(k)),
    y:keys.map((k,i)=>{
      const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      return Math.round(b-a);
    }),
    marker:{color:keys.map((k,i)=>{
      const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      return(b-a)>=0?'rgba(91,143,255,.8)':'rgba(255,124,58,.8)';
    })},
    text:keys.map((k,i)=>{
      const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;
      return(b-a>0?'+':'')+Math.round(b-a)+'°';
    }),textposition:'outside',
  }];
  pl('ch-cmp-delta',idTraces,{height:220,yaxis:{title:'В−А °C',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5}});

  // Summary table
  const tHead=['Самосвал','Дата','A-avg °C','B-avg °C','Разброс A','Разброс B','∆(B−A)','Макс цил.'];
  const tRows=keys.map((k,i)=>{
    const s=getSession(k);const ca=trCylAvgs[i];
    const aAvgM=mean(A_CYLS.map(n=>ca[n]).filter(v=>v!=null))||0;
    const bAvgM=mean(B_CYLS.map(n=>ca[n]).filter(v=>v!=null))||0;
    const aVals=A_CYLS.map(n=>ca[n]).filter(v=>v!=null);
    const bVals=B_CYLS.map(n=>ca[n]).filter(v=>v!=null);
    let maxCyl=null,maxDev=0;
    allCyls.forEach(n=>{if(ca[n]==null)return;const bk=A_CYLS.includes(n)?aAvgM:bAvgM;const dv=Math.abs(ca[n]-bk);if(dv>maxDev){maxDev=dv;maxCyl=n;}});
    return[sessionLabel(k),s?s.date:'?',Math.round(aAvgM),Math.round(bAvgM),
      Math.round((maxv(aVals)||0)-(minv(aVals)||0)),
      Math.round((maxv(bVals)||0)-(minv(bVals)||0)),
      (bAvgM-aAvgM>0?'+':'')+Math.round(bAvgM-aAvgM),
      maxCyl?'Ц'+maxCyl+' (∆'+Math.round(ca[maxCyl]-(A_CYLS.includes(maxCyl)?aAvgM:bAvgM))+'°)':'—'];
  });
  pl('ch-cmp-table',[{type:'table',
    header:{values:tHead,fill:{color:'#0f1e2e'},font:{color:'#00d4aa',size:10},align:'left'},
    cells:{values:tHead.map((_,ci)=>tRows.map(r=>r[ci])),
      fill:{color:'#112030'},font:{color:'#c8dff0',size:10},align:'left'},
  }],{height:Math.max(180,60+keys.length*28),margin:{l:4,r:4,t:4,b:4}});
}

// ═══ TAB: PRESSURE ══════════════════════════════════════════════════
function renderPressure(){
  if(!activeTruck||!TRUCKS[activeTruck]){
    document.getElementById('press-summary').innerHTML='← Выберите самосвал слева';
    return;
  }
  const s=getSession(activeTruck);if(!s)return;
  const {ts,press}=s;
  const secs=timeSecs(ts);
  const d=TRUCKS[activeTruck];
  const lbl=d.model+' №'+d.truck+' | '+s.date;
  const ACC='#00d4aa',AMB='#ffb830',RED='#ff4060',BLU='#5b8fff',ORG='#ff7c3a';

  function tsChart(id,key,color,title,h=200){
    const v=press&&press[key];
    if(!v||nn(v).length<5){pl(id,[],{height:h,title:{text:'Нет данных: '+title,font:{color:'#334d60',size:10}}});return;}
    pl(id,[
      {type:'scattergl',x:secs,y:v,name:'Факт',line:{color:color,width:1},opacity:.55},
      {type:'scattergl',x:secs,y:ewma(v,.06),name:'EWMA',line:{color:color,width:2.5}},
    ],{height:h,yaxis:{title:title},xaxis:{title:'сек'}});
  }

  // Oil pressure with thresholds
  {const v=press&&press.oil;
   if(v&&nn(v).length>4){
     pl('ch-p-oil',[
       {type:'scattergl',x:secs,y:v,name:'P масла',line:{color:ACC,width:1},opacity:.5},
       {type:'scattergl',x:secs,y:ewma(v),name:'EWMA',line:{color:ACC,width:2.5}},
     ],{height:220,yaxis:{title:'кПа'},xaxis:{title:'сек'},
       shapes:[
         {type:'rect',x0:0,x1:1,xref:'paper',y0:350,y1:650,fillcolor:'rgba(46,204,113,.04)',line:{width:0}},
         {type:'line',x0:0,x1:1,xref:'paper',y0:350,y1:350,line:{color:'rgba(255,184,48,.5)',dash:'dot',width:1.5}},
         {type:'line',x0:0,x1:1,xref:'paper',y0:300,y1:300,line:{color:'rgba(255,64,96,.5)',dash:'dot',width:1.5}},
       ],
       annotations:[
         {x:.01,xref:'paper',y:350,text:'⚠ 350',showarrow:false,font:{color:AMB,size:9},xanchor:'left'},
         {x:.01,xref:'paper',y:300,text:'🔴 300',showarrow:false,font:{color:RED,size:9},xanchor:'left'},
       ]});
   }else pl('ch-p-oil',[],{height:220,title:{text:'Нет данных давления масла',font:{color:'#334d60',size:10}}});}

  // Crankcase with thresholds
  {const v=press&&press.crank;
   if(v&&nn(v).length>4){
     pl('ch-p-crank',[
       {type:'scattergl',x:secs,y:v,name:'P картера',line:{color:RED,width:1},opacity:.5},
       {type:'scattergl',x:secs,y:ewma(v,.04),name:'EWMA',line:{color:RED,width:2.5}},
     ],{height:220,yaxis:{title:'кПа'},xaxis:{title:'сек'},
       shapes:[
         {type:'rect',x0:0,x1:1,xref:'paper',y0:0,y1:3,fillcolor:'rgba(46,204,113,.06)',line:{width:0}},
         {type:'line',x0:0,x1:1,xref:'paper',y0:3,y1:3,line:{color:'rgba(255,184,48,.5)',dash:'dot',width:1.5}},
         {type:'line',x0:0,x1:1,xref:'paper',y0:7,y1:7,line:{color:'rgba(255,64,96,.5)',dash:'dot',width:1.5}},
       ],
       annotations:[
         {x:.01,xref:'paper',y:3,text:'⚠ 3 кПа',showarrow:false,font:{color:AMB,size:9},xanchor:'left'},
         {x:.01,xref:'paper',y:7,text:'🔴 7 кПа',showarrow:false,font:{color:RED,size:9},xanchor:'left'},
       ]});
   }else pl('ch-p-crank',[],{height:220,title:{text:'Нет данных давления картера',font:{color:'#334d60',size:10}}});}

  // Boost pressure
  {const v=press&&press.boost;
   if(v&&nn(v).length>4){
     pl('ch-p-boost',[
       {type:'scattergl',x:secs,y:v,name:'Наддув',line:{color:BLU,width:1},opacity:.5},
       {type:'scattergl',x:secs,y:ewma(v),name:'EWMA',line:{color:BLU,width:2.5}},
     ],{height:220,yaxis:{title:'кПа'},xaxis:{title:'сек'},
       shapes:[
         {type:'rect',x0:0,x1:1,xref:'paper',y0:200,y1:350,fillcolor:'rgba(46,204,113,.04)',line:{width:0}},
         {type:'line',x0:0,x1:1,xref:'paper',y0:180,y1:180,line:{color:'rgba(255,184,48,.5)',dash:'dot',width:1.5}},
       ],
       annotations:[{x:.01,xref:'paper',y:180,text:'⚠ 180',showarrow:false,font:{color:AMB,size:9},xanchor:'left'}]});
   }else pl('ch-p-boost',[],{height:220,title:{text:'Нет данных наддува',font:{color:'#334d60',size:10}}});}

  // Temps dual
  {const ot=press&&press.oil_t,ct=press&&press.cool_t;
   const tr=[];
   if(ot&&nn(ot).length>4)tr.push({type:'scattergl',x:secs,y:ewma(ot),name:'T масла',line:{color:ORG,width:2}});
   if(ct&&nn(ct).length>4)tr.push({type:'scattergl',x:secs,y:ewma(ct),name:'T охл.',line:{color:ACC,width:2}});
   if(tr.length)pl('ch-p-temps',tr,{height:220,yaxis:{title:'°C'},xaxis:{title:'сек'},
     shapes:[
       {type:'line',x0:0,x1:1,xref:'paper',y0:105,y1:105,line:{color:'rgba(255,184,48,.4)',dash:'dot',width:1.2}},
       {type:'line',x0:0,x1:1,xref:'paper',y0:95,y1:95,line:{color:'rgba(46,204,113,.4)',dash:'dot',width:1.2}},
     ]});
   else pl('ch-p-temps',[],{height:220,title:{text:'Нет данных температур',font:{color:'#334d60',size:10}}});}

  // RPM + Load dual axis
  {const rpm=press&&press.rpm,ld=press&&press.load;
   const tr=[];
   if(rpm&&nn(rpm).length>4)tr.push({type:'scattergl',x:secs,y:ewma(rpm,.1),name:'RPM',line:{color:ACC,width:2},yaxis:'y'});
   if(ld&&nn(ld).length>4)tr.push({type:'scattergl',x:secs,y:ewma(ld,.1),name:'Нагрузка %',line:{color:AMB,width:2},yaxis:'y2'});
   if(tr.length)pl('ch-p-rpm',tr,{height:200,
     yaxis:{title:'RPM'},yaxis2:{title:'Нагрузка %',side:'right',overlaying:'y'},xaxis:{title:'сек'}});
   else pl('ch-p-rpm',[],{height:200,title:{text:'Нет данных оборотов',font:{color:'#334d60',size:10}}});}

  // Pressure summary
  const avgOil=mean(press&&press.oil);
  const avgCrank=mean(press&&press.crank);
  const avgBoost=mean(press&&press.boost);
  const oilC=avgOil!=null?(avgOil<300?'cp-r':avgOil<350?'cp-w':'cp-ok'):'';
  const crankC=avgCrank!=null?(avgCrank>7?'cp-r':avgCrank>3?'cp-w':'cp-ok'):'';
  const boostC=avgBoost!=null?(avgBoost<180?'cp-r':avgBoost<200?'cp-w':'cp-ok'):'';
  document.getElementById('press-summary').innerHTML=`
    <strong>${lbl}</strong> &nbsp;|&nbsp;
    P масла: <span class="cp ${oilC}">${avgOil!=null?Math.round(avgOil)+' кПа':'н/д'}</span> &nbsp;|&nbsp;
    P картера: <span class="cp ${crankC}">${avgCrank!=null?Math.round(avgCrank*10)/10+' кПа':'н/д'}</span> &nbsp;|&nbsp;
    Наддув: <span class="cp ${boostC}">${avgBoost!=null?Math.round(avgBoost)+' кПа':'н/д'}</span>`;
}

// ═══ FILE LOADING (browser CSV) ═════════════════════════════════════
const RU_MON={янв:'Jan',фев:'Feb',мар:'Mar',апр:'Apr',май:'May',июн:'Jun',июл:'Jul',авг:'Aug',сен:'Sep',окт:'Oct',ноя:'Nov',дек:'Dec'};
function ruDate2(s){for(const[r,e] of Object.entries(RU_MON))s=s.replace(r,e);return s;}

function splitCSV(line){
  const f=[];let q=false,s='';
  for(const c of line){if(c==='"'){q=!q;}else if(c===','&&!q){f.push(s.trim());s='';}else s+=c;}
  f.push(s.trim());return f;
}

function parseINSITE(text){
  const lines=text.split(/\r?\n/);
  let hdr=-1;
  for(let i=0;i<Math.min(80,lines.length);i++){
    const l=lines[i].trim();
    if(l.startsWith('"Date"')||l.startsWith('Date,')||l.startsWith('"Дата"')||l.startsWith('Дата,')){hdr=i;break;}
  }
  if(hdr<0)return null;
  const cols=splitCSV(lines[hdr]).map(c=>c.replace(/^"|"$/g,''));
  const dateC=cols.includes('Date')?'Date':'Дата';
  const timeC=cols.includes('Time')?'Time':'Время';
  const rows=[];
  for(let i=hdr+1;i<lines.length;i++){
    if(!lines[i].trim())continue;
    const fs=splitCSV(lines[i]).map(c=>c.replace(/^"|"$/g,''));
    if(fs.length<3)continue;
    const obj={};cols.forEach((c,j)=>obj[c]=fs[j]||'');rows.push(obj);
  }
  return{cols,rows,dateC,timeC};
}

function findCylColBrowser(cols,n){
  return cols.find(c=>c.includes('отработавших газов цилиндра '+n+' (°C)')||c.includes('Exhaust Temperature Sensor Cylinder '+n+' ('))||null;
}
function findPressBrowser(cols,patterns){
  for(const p of patterns){const c=cols.find(x=>x.includes(p));if(c)return c;}return null;
}

function processINSITE(text,fname){
  const parsed=parseINSITE(text);if(!parsed)return null;
  const{cols,rows,dateC,timeC}=parsed;
  const cylCols={};for(let n=1;n<=16;n++){const c=findCylColBrowser(cols,n);if(c)cylCols[n]=c;}
  if(Object.keys(cylCols).length<8)return null;
  const pressPat={
    oil:['Давление масла (кПа)','Engine Oil Pressure (kPa)'],
    crank:['Давление картерных','Crankcase Pressure (kPa)'],
    boost:['Давление во впускном коллекторе (кПа)- Идентификатор0','Intake Manifold Pressure (kPa)- Identifier0'],
    oil_t:['Температура масла (°C)','Engine Oil Temperature (°C)'],
    cool_t:['Датчик температуры (°C)- Идентификатор0','Engine Coolant Temperature (°C)'],
    rpm:['Частота вращения двигателя','Engine Speed (RPM)'],
    load:['Относительная нагрузка','Percent Load'],
  };
  const pressC={};for(const[k,p]of Object.entries(pressPat)){const c=findPressBrowser(cols,p);if(c)pressC[k]=c;}

  // Parse, downsample
  const step=Math.max(1,Math.floor(rows.length/600));
  const out={ts:[],cyls:{},press:{}};
  for(let n=1;n<=16;n++)if(cylCols[n])out.cyls[n]=[];
  for(const k of Object.keys(pressC))out.press[k]=[];

  for(let i=0;i<rows.length;i+=step){
    const r=rows[i];
    let d=ruDate2(r[dateC]||''),t=(r[timeC]||'').substring(0,8);
    try{const dt=new Date(d+' '+t);if(isNaN(dt))continue;out.ts.push(dt.toISOString().substring(0,19));}catch{continue;}
    for(let n=1;n<=16;n++){if(!cylCols[n])continue;const v=parseFloat((r[cylCols[n]]||'').replace(',','.'));out.cyls[n].push(isNaN(v)?null:Math.round(v));}
    for(const[k,c] of Object.entries(pressC)){const v=parseFloat((r[c]||'').replace(',','.'));out.press[k].push(isNaN(v)?null:Math.round(v*10)/10);}
  }
  if(out.ts.length<10)return null;

  // Date string
  const raw=out.ts[0]?.substring(0,10)||'';
  const p=raw.split('-');const dateStr=p.length===3?`${p[2]}.${p[1]}.${p[0]}`:raw;

  // Truck number
  let truck=null;
  for(const pat of[/№\s*(\d+)/,/N(\d+)\b/,/[\s_-](\d{2,3})[\s_.(]/]){const m=fname.match(pat);if(m){truck=+m[1];break;}}
  if(!truck)truck=Math.floor(Math.random()*900)+10;
  const model=fname.includes('NTE')||fname.includes('DML')?'NTE200':'730E';
  const key=model+'_'+truck;

  return{key,truck,model,session:{date:dateStr,ts:out.ts,cyls:out.cyls,press:out.press,file:fname}};
}

function loadFiles(files){
  let loaded=0;
  const total=files.length;
  Array.from(files).forEach(file=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const result=processINSITE(e.target.result,file.name);
        if(result){
          const{key,truck,model,session}=result;
          if(!TRUCKS[key]){TRUCKS[key]={truck,model,sessions:[session]};TRUCK_ORDER.push(key);}
          else TRUCKS[key].sessions.push(session);
          loaded++;
          document.getElementById('file-status').textContent=`Загружено: ${loaded}/${total}`;
          buildSidebar();
          if(!activeTruck){selectTruck(key);}
        }
      }catch(ex){console.warn(file.name,ex);}
    };
    reader.readAsText(file,'windows-1251');
  });
}

// ═══ INIT ════════════════════════════════════════════════════════════
window.addEventListener('load',function(){
  if(TRUCK_ORDER.length){
    activeTruck=TRUCK_ORDER[0];
    buildSidebar();
    buildDatePills(activeTruck);
    updateHeader();
    renderCylinders();
  }else{
    buildSidebar();
  }
});
</script>
</body>
</html>
"""


def build_html(db_data, truck_order):
    plotly_src = PLOTLY_JS.read_text(encoding='utf-8')
    data_json  = json.dumps(db_data,     ensure_ascii=False, separators=(',', ':'))
    order_json = json.dumps(truck_order, ensure_ascii=False)
    html = HTML.replace('__PLOTLY__', plotly_src) \
               .replace('__DATA__',   data_json) \
               .replace('__ORDER__',  order_json)
    return html


def main():
    print('=== QSK50 Cylinders v2 Builder ===\n')
    db_data, truck_order = build_data()
    if not db_data:
        print('ERROR: No data found!')
        return
    total = sum(sum(len(s['ts']) for s in v['sessions']) for v in db_data.values())
    kb = len(json.dumps(db_data, separators=(',', ':')).encode()) // 1024
    print(f'\n{len(db_data)} trucks, {total} rows, data JSON: {kb} KB')
    html = build_html(db_data, truck_order)
    OUTPUT.write_text(html, encoding='utf-8')
    mb = OUTPUT.stat().st_size / 1024 / 1024
    print(f'Written: {OUTPUT}  ({mb:.1f} MB)')


if __name__ == '__main__':
    main()
