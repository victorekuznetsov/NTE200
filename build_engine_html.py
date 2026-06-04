#!/usr/bin/env python3
"""
Build self-contained QSK50 engine dashboard HTML with embedded INSITE data.
Reads CSV files from .qsk50_extracted/, outputs qsk50_engine_data.html.
Usage:  python3 build_engine_html.py
"""

import io, json, re, pathlib, warnings
import pandas as pd

warnings.filterwarnings('ignore')

EXTRACT   = pathlib.Path('/home/user/NTE200/.qsk50_extracted')
TEMPLATE  = pathlib.Path('/home/user/NTE200/qsk50_engine.html')
PLOTLY_JS = pathlib.Path('/usr/local/lib/python3.11/dist-packages/plotly/package_data/plotly.min.js')
OUTPUT    = pathlib.Path('/home/user/NTE200/qsk50_engine_data.html')

MAX_ROWS = 900  # max data points per truck

# Column patterns to keep — English + Russian
KEEP_PATTERNS = [
    # English
    'Engine Speed (RPM)',
    'Engine Coolant Temperature',
    'Engine Oil Temperature',
    'Engine Oil Pressure',
    'Average Exhaust Temperature',
    'Exhaust Temperature Sensor Cylinder',
    'Instantaneous Fuel Rate',
    'Fuel Flow Rate Commanded',
    'Crankcase Pressure (kPa)',
    'Intake Manifold Pressure (kPa)',
    'Intake Manifold Air Temperature',
    'Intake Manifold Temperature Sensor',
    'Percent Load',
    'Battery Voltage (V)',
    'Fuel Rail Pressure',
    'Fuel Pump Drive',
    'Barometric Air Pressure (kPa)',
    'Coolant Pressure Sensor (kPa)',
    # Russian
    'Частота вращения',
    'Датчик температуры (°C)',
    'Температура масла (°C)',
    'Давление масла (кПа)',
    'Давление картерных',
    'Средняя температура отработавших',
    'отработавших газов цилиндра',
    'Мгновенный расход топлива',
    'Давление во впускном коллекторе',
    'Относительная нагрузка',
    'Напряжение аккумуляторной батареи',
    'Заданное давление в общем топливопроводе',
    'Измеренное давление в общем топливопроводе',
    'Рабочий цикл привода топливного насоса',
    'Заданный ток привода топливного насоса',
    'Температура воздуха во впускном коллекторе',
    'Датчик температуры впускного коллектора',
    'Атмосферное давление (кПа)',
    'левый ряд',
    'правый ряд',
]

RU_MONTHS = {
    'янв': 'Jan', 'фев': 'Feb', 'мар': 'Mar', 'апр': 'Apr',
    'май': 'May', 'июн': 'Jun', 'июл': 'Jul', 'авг': 'Aug',
    'сен': 'Sep', 'окт': 'Oct', 'ноя': 'Nov', 'дек': 'Dec',
}


def ru_date(d: str) -> str:
    for ru, en in RU_MONTHS.items():
        d = d.replace(ru, en)
    return d


def is_relevant(col: str) -> bool:
    return any(p in col for p in KEEP_PATTERNS)


def open_csv(fp: pathlib.Path):
    """
    Detect encoding and return (lines, enc).
    Strategy: cp1251 first (INSITE 8.x default), then UTF-8 for newer exports.
    """
    # Check for UTF-8 BOM
    with open(fp, 'rb') as fh:
        bom = fh.read(3)
    if bom[:3] == b'\xef\xbb\xbf':
        enc_order = ('utf-8-sig', 'cp1251')
    else:
        # Try cp1251 first — it's the default for Russian INSITE exports
        enc_order = ('cp1251', 'utf-8-sig', 'utf-8')

    for enc in enc_order:
        try:
            with open(fp, encoding=enc, errors='replace') as fh:
                raw = fh.read()
            if 'Дата' in raw or '"Date"' in raw or 'Date,' in raw:
                return raw.splitlines(keepends=True), enc
        except Exception:
            continue
    return None, None


def read_csv(fp: pathlib.Path):
    """
    Parse INSITE CSV.
    Returns dict with columnar data:
      {cols: [...], ts: [...], data: {col: [vals...]}, truck, model, file}
    or None on failure.
    """
    lines, enc = open_csv(fp)
    if lines is None:
        return None

    # Find header row
    hdr = None
    for i, line in enumerate(lines[:80]):
        s = line.strip()
        if (s.startswith('"Date"') or s.startswith('Date,') or
                s.startswith('"Дата"') or s.startswith('Дата,')):
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

    # Select relevant columns
    keep = [c for c in all_cols if c not in (date_col, time_col) and is_relevant(c)]
    if not keep:
        return None

    df_sub = df[[date_col, time_col] + keep].copy()

    # Numeric conversion (comma-decimal)
    for col in keep:
        try:
            conv = df_sub[col].astype(str).str.replace(',', '.', regex=False).str.strip()
            num  = pd.to_numeric(conv, errors='coerce')
            if num.notna().sum() > len(df_sub) * 0.25:
                df_sub[col] = num.round(1)
        except Exception:
            pass

    # Parse timestamps
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

    if len(df_sub) < 5:
        return None

    # Downsample
    step = max(1, len(df_sub) // MAX_ROWS)
    df_sub = df_sub.iloc[::step].reset_index(drop=True)

    # Build columnar format
    ts = df_sub['_ts'].tolist()

    data = {}
    for col in keep:
        series = df_sub[col]
        if pd.api.types.is_numeric_dtype(series):
            vals = [None if pd.isna(v) else round(float(v), 1) for v in series]
        else:
            vals = [None if pd.isna(v) else str(v) for v in series]
        data[col] = vals

    return {
        'cols': keep,   # column names (for JS detection)
        'ts':   ts,     # ISO timestamp strings
        'data': data,   # col → [values]
    }


def truck_num(fname: str):
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
    db_data, truck_order = {}, []
    csv_files = sorted(EXTRACT.glob('*.csv'))
    print(f"CSV files: {len(csv_files)}")

    for fp in csv_files:
        name = fp.name
        if 'production' in name.lower() or 'speed' in name.lower():
            continue
        truck = truck_num(name)
        if truck is None:
            print(f"  SKIP (no truck#): {name}")
            continue

        model = 'NTE200' if ('NTE' in name or 'DML' in name) else '730E'
        key   = f"{model}_{truck}"

        print(f"  {name} … ", end='', flush=True)
        result = read_csv(fp)

        if result is None:
            print("no data")
            continue

        n_rows = len(result['ts'])
        n_cols = len(result['cols'])
        print(f"{n_rows} rows, {n_cols} cols")

        if key not in db_data:
            db_data[key] = {**result, 'truck': truck, 'model': model, 'file': name}
            truck_order.append(key)
        else:
            # Merge additional records
            db_data[key]['ts'].extend(result['ts'])
            for col, vals in result['data'].items():
                if col in db_data[key]['data']:
                    db_data[key]['data'][col].extend(vals)
                else:
                    db_data[key]['data'][col] = vals

    return db_data, truck_order


# ── JS COMPATIBILITY SHIM ─────────────────────────────────────────────────────
# The existing JS uses row-based access: getCol(rows, col)
# We'll inject this shim that converts columnar → row-based at load time.
JS_SHIM = r"""
// ═══ PRE-LOADED DATA (build_engine_html.py) ═══
(function(){
  var _PRE = __DATA_JSON__;
  var _ORDER = __ORDER_JSON__;

  for(var i=0;i<_ORDER.length;i++){
    var key=_ORDER[i], d=_PRE[key];
    if(!d) continue;

    // Convert columnar → row-based for compatibility with existing getCol() calls
    var nRows=d.ts.length;
    var rows=new Array(nRows);
    for(var r=0;r<nRows;r++){
      var obj={};
      obj._dt=new Date(d.ts[r]);
      for(var ci=0;ci<d.cols.length;ci++){
        obj[d.cols[ci]]=d.data[d.cols[ci]][r];
      }
      rows[r]=obj;
    }

    DB[key]={rows:rows,cols:d.cols,truck:d.truck,model:d.model,file:d.file};
    truckOrder.push(key);
  }
})();

window.addEventListener('load',function(){
  rebuildPills();
  updateKPIs();
  renderOverview();
});
"""

OLD_INIT = (
    "window.addEventListener('load',()=>{\n"
    "  // Show initial empty state\n"
    "  renderOverview();\n"
    "  document.getElementById('no-trucks').style.display='';\n"
    "});"
)


def patch_html(template: str, db_data: dict, truck_order: list) -> str:
    # ── 1. Inline Plotly.js ──────────────────────────────────────────────────
    plotly_src = PLOTLY_JS.read_text(encoding='utf-8')
    template = template.replace(
        '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>',
        f'<script>{plotly_src}</script>'
    )

    # ── 2. Remove PapaParse CDN ──────────────────────────────────────────────
    template = template.replace(
        '<script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>',
        ''
    )

    # ── 3. Replace loadbar with status banner ────────────────────────────────
    n_t = len(db_data)
    n_r = sum(len(v['ts']) for v in db_data.values())
    banner = (
        '<div class="loadbar">'
        '<div class="lb-title">⚙ QSK50 · INSITE Log Files</div>'
        f'<div class="status ok" id="status">'
        f'Данные загружены: {n_t} самосвалов · {n_r:,} записей</div>'
        '</div>'
    )
    lb_s = template.find('<div class="loadbar">')
    if lb_s >= 0:
        depth, pos, lb_e = 0, lb_s, lb_s
        while pos < len(template):
            if template[pos:pos+4] == '<div':
                depth += 1
            elif template[pos:pos+6] == '</div>':
                depth -= 1
                if depth == 0:
                    lb_e = pos + 6
                    break
            pos += 1
        template = template[:lb_s] + banner + template[lb_e:]

    # ── 4. Fix colsCyls for English column names (Cylinder N () ─────────────
    template = template.replace(
        "c.includes('Cyl'+n+' Exhaust'))",
        "c.includes('Cyl'+n+' Exhaust')||c.includes('Cylinder '+n+' ('))"
    )

    # ── 5. Serialize data ────────────────────────────────────────────────────
    data_json  = json.dumps(db_data,     ensure_ascii=False, separators=(',', ':'))
    order_json = json.dumps(truck_order, ensure_ascii=False)

    shim = JS_SHIM.replace('__DATA_JSON__',  data_json) \
                  .replace('__ORDER_JSON__', order_json)

    # ── 6. Replace init block ────────────────────────────────────────────────
    if OLD_INIT in template:
        template = template.replace(OLD_INIT, shim)
    else:
        last = template.rfind('</script>')
        template = template[:last] + '\n' + shim + '\n' + template[last:]

    return template


def main():
    print('=== QSK50 Engine Dashboard Builder ===\n')

    db_data, truck_order = build_data()
    if not db_data:
        print('ERROR: No data loaded!')
        return

    total_rows = sum(len(v['ts']) for v in db_data.values())
    data_bytes = len(json.dumps(db_data, separators=(',', ':')).encode())
    print(f'\nLoaded {len(db_data)} trucks, {total_rows} rows')
    print(f'Data JSON: {data_bytes/1024:.0f} KB')

    print('\nPatching HTML template …')
    template = TEMPLATE.read_text(encoding='utf-8')
    result   = patch_html(template, db_data, truck_order)

    OUTPUT.write_text(result, encoding='utf-8')
    size_mb = OUTPUT.stat().st_size / 1024 / 1024
    print(f'Written: {OUTPUT}  ({size_mb:.1f} MB)')
    print('\nDone! Open qsk50_engine_data.html in a browser.')


if __name__ == '__main__':
    main()
