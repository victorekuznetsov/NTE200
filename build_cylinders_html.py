#!/usr/bin/env python3
"""Build qsk50_cylinders.html — comprehensive QSK50 dashboard."""

import io, json, re, pathlib, warnings
import pandas as pd

warnings.filterwarnings('ignore')

EXTRACT   = pathlib.Path('/home/user/NTE200/.qsk50_extracted')
PLOTLY_JS = pathlib.Path('/usr/local/lib/python3.11/dist-packages/plotly/package_data/plotly.min.js')
OUTPUT    = pathlib.Path('/home/user/NTE200/qsk50_cylinders.html')

MAX_ROWS = 800

RU_MONTHS = {
    'янв':'Jan','фев':'Feb','мар':'Mar','апр':'Apr','май':'May','июн':'Jun',
    'июл':'Jul','авг':'Aug','сен':'Sep','окт':'Oct','ноя':'Nov','дек':'Dec',
}
A_BANK = [1,3,5,7,9,11,13,15]
B_BANK = [2,4,6,8,10,12,14,16]

PARAM_PATTERNS = {
    'rpm':         ['Engine Speed (RPM)- Identifier0', 'Частота вращения двигателя (об/мин)- Идентификатор0'],
    'load':        ['Percent Load (percent)- Identifier0', 'Относительная нагрузка (Проценты)- Идентификатор0'],
    'accel':       ['Percent Accelerator Pedal or Lever (percent)- Identifier0', 'Процент действия педали или рычага акселератора (Проценты)- Идентификатор0'],
    'oil_p':       ['Engine Oil Pressure (kPa)- Identifier0', 'Давление масла (кПа)- Идентификатор0'],
    'crank_p':     ['Crankcase Pressure (kPa)- Identifier0', 'Давление картерных газов (кПа)- Идентификатор0'],
    'boost':       ['Intake Manifold Pressure (kPa)- Identifier0', 'Давление во впускном коллекторе (кПа)- Идентификатор0'],
    'boost1':      ['Intake Manifold Pressure Sensor 1 (kPa)- Identifier0', 'Датчик давления во впускном коллекторе 1 (кПа)- Идентификатор0'],
    'boost2':      ['Intake Manifold Pressure Sensor 2 (kPa)- Identifier1', 'Датчик давления во впускном коллекторе 2 (кПа)- Идентификатор1'],
    'rail_meas':   ['Fuel Rail Pressure Measured (bar)- Identifier0', 'Измеренное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
    'rail_cmd':    ['Fuel Rail Pressure Commanded (bar)- Identifier0', 'Заданное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
    'pump_oil_p':  ['Fuel Pump Oil Pressure (kPa)- Identifier0', 'Fuel Pump Oil Pressure (кПа)- Идентификатор0'],
    'pre_filt_p':  ['Pre Oil Filter Pressure (kPa)- Identifier1', 'Давление в предварительном масляном фильтре (кПа)- Идентификатор1'],
    'filt_dp':     ['Oil Filter Differential Pressure Sensor (Calculated) (kPa)- Identifier0', 'датчик перепада давления в масляном фильтре (расчетное значение) (кПа)- Идентификатор0'],
    'cool_p':      ['Coolant Pressure Sensor (kPa)- Identifier0', 'Датчик давления охлаждающей жидкости (кПа)- Идентификатор0'],
    'baro_p':      ['Barometric Air Pressure (kPa)- Identifier0', 'Атмосферное давление (кПа)- Идентификатор0'],
    'fuel_reg_p':  ['Fuel Regulator Intake Pressure (kPa)- Identifier0', 'Давление на входе топливного регулятора (кПа)- Идентификатор0'],
    'cool_t':      ['Engine Coolant Temperature (°C)- Identifier0', 'Датчик температуры (°C)- Идентификатор0'],
    'oil_t':       ['Engine Oil Temperature (°C)- Identifier0', 'Температура масла (°C)- Идентификатор0'],
    'intake_t':    ['Intake Manifold Air Temperature (°C)- Identifier0', 'Температура воздуха во впускном коллекторе (°C)- Идентификатор0'],
    'im_t1':       ['Intake Manifold Temperature Sensor 1 (°C)- Identifier0', 'Датчик температуры впускного коллектора 1 (°C)- Идентификатор0'],
    'im_t2':       ['Intake Manifold Temperature Sensor 2 (°C)- Identifier0', 'Датчик температуры впускного коллектора 2 (°C)- Идентификатор0'],
    'im_t3':       ['Intake Manifold Temperature Sensor 3 (°C)- Identifier0', 'Датчик температуры впускного коллектора 3 (°C)- Идентификатор0'],
    'im_t4':       ['Intake Manifold Temperature Sensor 4 (°C)- Identifier0', 'Датчик температуры впускного коллектора 4 (°C)- Идентификатор0'],
    'fuel_t':      ['Fuel Temperature Sensor (°C)- Identifier0', 'Датчик температуры топлива (°C)- Идентификатор0'],
    'ecm_t2':      ['ECM Temperature 2 (°C)- Identifier1', 'Температура 2 модуля ECM (°C)- Идентификатор1'],
    'ecm_t3':      ['ECM Temperature 3 (°C)- Identifier144', 'Температура 3 модуля ECM (°C)- Идентификатор144'],
    'avg_egt_l':   ['Average Exhaust Temperature - Left Bank (Calculated) (°C)- Identifier144', 'Средняя температура отработавших газов - левый ряд (расчетное значение) (°C)- Идентификатор144'],
    'avg_egt_r':   ['Average Exhaust Temperature - Right Bank (Calculated) (°C)- Identifier1', 'Средняя температура отработавших газов - правый ряд (расчетное значение) (°C)- Идентификатор1'],
    'fuel_flow':   ['Fuel Flow Rate Commanded (L/hr)- Identifier0', 'Управляемый расход топлива (л/час)- Идентификатор0'],
    'fuel_inst':   ['Instantaneous Fuel Rate (L/hr)- Identifier0', 'Мгновенный расход топлива (л/час)- Идентификатор0'],
    'fuel_dos':    ['Instantaneous Dosing Fuel Rate (L/hr)- Identifier0', 'Мгновенный расход топлива через систему дозирования (л/час)- Идентификатор0'],
    'pump_mA_cmd': ['Fuel Pump Actuator Commanded Current (mA)- Identifier0', 'Заданный ток привода топливного насоса (мА)- Идентификатор0'],
    'pump_mA_meas':['Fuel Pump Actuator Measured Current (mA)- Identifier0', 'Измеренный ток привода топливного насоса (мА)- Идентификатор0'],
    'pump_dc':     ['Fuel Pump Actuator Duty Cycle (percent)- Identifier0', 'Рабочий цикл привода топливного насоса (Проценты)- Идентификатор0'],
    'lift_dc':     ['Electric Fuel Lift Pump Duty Cycle (percent)- Identifier0', 'Рабочий цикл электрического подкачивающего топливного насоса (Проценты)- Идентификатор0'],
    'water_fuel':  ['Water In Fuel State- Identifier0', 'Состояние обнаружения воды в топливе- Идентификатор0'],
    'fan_cmd':     ['Fan Control Command (percent)- Identifier0', 'Команда управления вентилятором (Проценты)- Идентификатор0'],
    'battery':     ['Battery Voltage (V)- Identifier0', 'Напряжение аккумуляторной батареи (В)- Идентификатор0'],
    'amber_lamp':  ['Amber Warning Lamp Status- Identifier0', 'Состояние желтого сигнального индикатора- Идентификатор0'],
    'red_lamp':    ['Red Stop Lamp Status- Identifier0', 'Состояние красного индикатора останова- Идентификатор0'],
    'cool_level':  ['Engine Coolant Level- Identifier0', 'Уровень охлаждающей жидкости- Идентификатор0'],
}


def ru_date(d):
    for ru, en in RU_MONTHS.items():
        d = d.replace(ru, en)
    return d


def open_csv(fp):
    with open(fp, 'rb') as fh:
        bom = fh.read(3)
    enc_order = ('utf-8-sig', 'cp1251') if bom[:3] == b'\xef\xbb\xbf' else ('cp1251', 'utf-8-sig', 'utf-8')
    for enc in enc_order:
        try:
            with open(fp, encoding=enc, errors='replace') as fh:
                raw = fh.read()
            if 'Дата' in raw or 'Date' in raw:
                return raw.splitlines(keepends=True), enc
        except Exception:
            pass
    return None, None


def find_header_row(lines):
    for i, line in enumerate(lines[:80]):
        s = line.strip()
        if (s.startswith('"Date"') or s.startswith('Date,') or
                s.startswith('"Дата"') or s.startswith('Дата,')):
            return i
    return None


def find_col(cols, *patterns):
    """Exact substring match, try each pattern in order."""
    for pat in patterns:
        for c in cols:
            if pat in c:
                return c
    return None


def find_cyl_col(cols, n):
    for c in cols:
        if f'Exhaust Temperature Sensor Cylinder {n} (' in c:
            return c
        if f'отработавших газов цилиндра {n} (°C)' in c:
            return c
    return None


def parse_val(series):
    """Convert a pandas Series with comma decimals to float list."""
    conv = series.astype(str).str.replace(',', '.', regex=False).str.strip()
    num = pd.to_numeric(conv, errors='coerce')
    return num


def read_csv(fp):
    lines, enc = open_csv(fp)
    if lines is None:
        return None

    hdr = find_header_row(lines)
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
    if date_col not in all_cols or time_col not in all_cols:
        return None

    # Cylinder columns
    cyl_cols = {}
    for n in range(1, 17):
        c = find_cyl_col(all_cols, n)
        if c:
            cyl_cols[n] = c

    if len(cyl_cols) < 8:
        return None

    # Parameter columns
    param_cols = {}
    for key, patterns in PARAM_PATTERNS.items():
        c = find_col(all_cols, *patterns)
        if c:
            param_cols[key] = c

    # Parse numeric values
    keep = [date_col, time_col] + list(cyl_cols.values()) + list(set(param_cols.values()))
    df_sub = df[[c for c in keep if c in df.columns]].copy()

    for col in list(cyl_cols.values()) + list(set(param_cols.values())):
        if col not in df_sub.columns:
            continue
        try:
            num = parse_val(df_sub[col])
            if num.notna().sum() > len(df_sub) * 0.1:
                df_sub[col] = num
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

    if len(df_sub) < 10:
        return None

    # Downsample
    step = max(1, len(df_sub) // MAX_ROWS)
    df_sub = df_sub.iloc[::step].reset_index(drop=True)

    ts = df_sub['_ts'].tolist()

    # Build cylinder data
    cyls = {}
    for n, col in cyl_cols.items():
        s = df_sub[col]
        cyls[str(n)] = [None if pd.isna(v) else int(round(float(v))) for v in s]

    # Build parameter data — handle boolean/string cols (water_fuel, lamps, cool_level)
    BOOL_KEYS = {'water_fuel', 'amber_lamp', 'red_lamp', 'cool_level'}
    p = {}
    for key, col in param_cols.items():
        if col not in df_sub.columns:
            continue
        s = df_sub[col]
        if key in BOOL_KEYS:
            # Convert On/Off/Low/High/etc to 0/1
            def bool_val(v):
                sv = str(v).strip().lower()
                if sv in ('on', 'yes', 'true', '1', 'low', 'active', 'detected'):
                    return 1
                if sv in ('off', 'no', 'false', '0', 'ok', 'normal', 'not detected'):
                    return 0
                try:
                    return int(float(sv.replace(',', '.')))
                except Exception:
                    return None
            p[key] = [bool_val(v) for v in s]
        else:
            p[key] = [None if pd.isna(v) else round(float(v), 2) for v in s]

    # Date display string
    raw_date = ts[0][:10] if ts else ''
    parts = raw_date.split('-')
    date_str = f"{parts[2]}.{parts[1]}.{parts[0]}" if len(parts) == 3 else raw_date

    return {
        'ts':   ts,
        'cyls': cyls,
        'p':    p,
        'date': date_str,
        'n_cyls': len(cyl_cols),
        'n_params': len(param_cols),
    }


def truck_num(fname):
    for pat in [r'№\s*(\d+)', r'\bN(\d+)\b', r'\((\d+)\)', r'[\s_-](\d{2,3})[\s_.( ]']:
        m = re.search(pat, fname)
        if m:
            return int(m.group(1))
    return None


def build_data():
    db, order = {}, []
    for fp in sorted(EXTRACT.glob('*.csv')):
        name = fp.name
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
        print(f"{len(result['ts'])} rows, {result['n_cyls']} cyls, {result['n_params']} params")
        session = {
            'date': result['date'],
            'ts':   result['ts'],
            'cyls': result['cyls'],
            'p':    result['p'],
            'file': name,
        }
        if key not in db:
            db[key] = {'truck': truck, 'model': model, 'sessions': [session]}
            order.append(key)
        else:
            db[key]['sessions'].append(session)
    return db, order


# ─────────────────────────────────────────────────────────────────────────────
# HTML TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>QSK50 — Дашборд INSITE</title>
<script>__PLOTLY__</script>
<style>
:root{
  --bg:#293136;--bg2:#232d34;--sb:#1e262c;--card:#232d34;--brd:rgba(62,240,175,0.12);
  --tx:#ffffff;--tx2:#B3BDDA;--tx3:#80868B;
  --acc:#3EF0AF;--red:#ff4060;--amb:#FBAE40;--grn:#3EF0AF;
  --blu:#7E83FA;--pur:#B668E4;--org:#FBAE40;
  --a-col:#FBAE40;--b-col:#7E83FA;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:Arial,system-ui,sans-serif;background:var(--bg);color:var(--tx);font-size:13px;display:grid;grid-template-columns:230px 1fr;grid-template-rows:60px 1fr;height:100vh}

/* HEADER */
.hdr{grid-column:1/-1;grid-row:1;display:flex;align-items:center;gap:12px;padding:0 16px;background:rgba(20,27,32,0.98);border-bottom:1px solid var(--brd);z-index:10}
.hdr-left{display:flex;flex-direction:column;gap:2px}
.logo{font-weight:800;font-size:14px;color:#fff;letter-spacing:-.3px;white-space:nowrap}
.logo span{color:var(--tx3);font-weight:400;font-size:10px;margin-left:8px}
.contact-notice{font-size:10px;color:var(--tx2);font-style:italic}
.hdr-info{font-size:11px;color:var(--amb);font-family:'Courier New',monospace;margin-left:auto;white-space:nowrap}

/* RANGE BAR */
.range-bar{display:flex;align-items:center;gap:6px;padding:5px 12px;background:var(--bg2);border-bottom:1px solid var(--brd);flex-shrink:0;flex-wrap:wrap}
.rng-label{font-size:10px;color:var(--tx3)}
.rng-btn{padding:3px 10px;border:1px solid var(--brd);border-radius:12px;background:none;color:var(--tx3);cursor:pointer;font-size:10px;transition:all .12s}
.rng-btn:hover{border-color:var(--acc);color:var(--acc)}
.rng-btn.act{background:rgba(62,240,175,0.12);border-color:var(--acc);color:var(--acc);font-weight:700}
.rng-slider-wrap{display:flex;align-items:center;gap:5px;margin-left:8px}
.rng-slider-wrap input[type=range]{width:110px;accent-color:var(--acc)}
.rng-val{font-size:9px;color:var(--tx2);font-family:monospace;min-width:28px}

/* SIDEBAR */
.sidebar{grid-column:1;grid-row:2;background:var(--sb);border-right:1px solid var(--brd);display:flex;flex-direction:column;overflow:hidden}
.sb-section{padding:8px 10px;border-bottom:1px solid var(--brd);flex-shrink:0}
.sb-label{font-size:9px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.7px;margin-bottom:6px;display:flex;align-items:center;justify-content:space-between}
.sb-label button{font-size:9px;color:var(--acc);background:none;border:1px solid rgba(62,240,175,0.3);border-radius:4px;padding:1px 6px;cursor:pointer}
.sb-label button:hover{background:rgba(62,240,175,0.08)}
.sb-label button.active{background:rgba(62,240,175,0.15);border-color:var(--acc)}

.truck-list{flex:1;overflow-y:auto;padding:4px 0}
.truck-list::-webkit-scrollbar{width:3px}
.truck-list::-webkit-scrollbar-thumb{background:var(--brd);border-radius:2px}
.truck-item{display:flex;align-items:center;gap:7px;padding:6px 10px;cursor:pointer;transition:background .12s;border-left:3px solid transparent;user-select:none}
.truck-item:hover{background:rgba(255,255,255,0.04)}
.truck-item.active{background:rgba(62,240,175,0.07);border-left-color:var(--acc)}
.truck-item.cmp-sel{background:rgba(126,131,250,0.06);border-left-color:var(--blu)}
.truck-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.truck-info{flex:1;min-width:0}
.truck-name{font-size:11px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.truck-date{font-size:9px;color:var(--tx3);margin-top:1px}
.cmp-chk{width:13px;height:13px;accent-color:var(--blu);flex-shrink:0;cursor:pointer;display:none}
.cmp-chk.show{display:block}

/* Date pills */
.date-pills{display:flex;flex-wrap:wrap;gap:4px}
.date-pill{padding:2px 8px;border:1px solid var(--brd);border-radius:10px;font-size:9px;color:var(--tx2);cursor:pointer;transition:all .12s}
.date-pill:hover{border-color:var(--acc);color:var(--acc)}
.date-pill.active{background:rgba(62,240,175,0.12);border-color:var(--acc);color:var(--acc);font-weight:600}

/* Calendar panel */
.cal-panel{max-height:0;overflow:hidden;transition:max-height .3s ease}
.cal-panel.open{max-height:300px;overflow-y:auto}
.cal-grid{display:grid;gap:2px;margin-top:4px}
.cal-head{font-size:8px;color:var(--tx3);padding:2px 3px;text-align:center}
.cal-truck-name{font-size:8px;color:var(--tx2);padding:2px 3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cal-cell{padding:2px 4px;border-radius:3px;font-size:8px;text-align:center;cursor:default;color:var(--tx3)}
.cal-cell.has-data{background:rgba(62,240,175,0.15);color:var(--acc);cursor:pointer;font-weight:600}
.cal-cell.has-data:hover{background:rgba(62,240,175,0.28)}
.cal-cell.active-date{outline:1px solid var(--blu);background:rgba(126,131,250,0.18);color:var(--blu)}

/* File upload */
.file-zone{border:1px dashed rgba(60,120,180,0.35);border-radius:6px;padding:7px;text-align:center;cursor:pointer;transition:all .15s;position:relative}
.file-zone:hover{border-color:var(--acc);background:rgba(62,240,175,0.04)}
.file-zone-icon{font-size:16px;margin-bottom:2px}
.file-zone-txt{font-size:9px;color:var(--tx3);line-height:1.4}
.file-status{font-size:9px;color:var(--grn);margin-top:3px;font-weight:600}

/* MAIN */
.main{grid-column:2;grid-row:2;display:flex;flex-direction:column;overflow:hidden}
.tabs{display:flex;border-bottom:1px solid var(--brd);background:var(--bg2);flex-shrink:0;overflow-x:auto}
.tabs::-webkit-scrollbar{height:2px}
.tbtn{padding:9px 14px;border:none;border-bottom:2px solid transparent;background:none;color:var(--tx3);cursor:pointer;font-size:11px;font-weight:500;white-space:nowrap;transition:all .12s}
.tbtn:hover{color:var(--tx);background:rgba(255,255,255,0.03)}
.tbtn.act{color:var(--acc);border-bottom-color:var(--acc);background:rgba(62,240,175,0.04)}
.content{flex:1;overflow-y:auto;overflow-x:hidden;padding:10px}
.content::-webkit-scrollbar{width:5px}
.content::-webkit-scrollbar-thumb{background:var(--brd);border-radius:3px}

.panel{display:none;flex-direction:column;gap:10px}
.panel.show{display:flex}

/* Cards */
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.full{grid-column:1/-1}
.card{background:var(--card);border:1px solid var(--brd);border-radius:9px;padding:9px;overflow:hidden}
.card-title{font-size:10px;font-weight:700;color:var(--tx2);margin-bottom:7px;text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.tag{padding:2px 6px;border-radius:3px;font-size:9px;font-weight:700;letter-spacing:.3px}
.tag-a{background:rgba(255,124,58,.2);color:var(--org)}
.tag-b{background:rgba(126,131,250,.2);color:var(--blu)}
.thres-badges{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:5px}
.tb-item{display:flex;align-items:center;gap:4px;font-size:9px;color:var(--tx3)}
.tb-box{width:18px;height:3px;border-radius:2px}

.stats-row{display:flex;gap:12px;padding:5px 8px;background:rgba(255,255,255,.025);border-radius:5px;border:1px solid var(--brd);margin-top:5px;flex-wrap:wrap}
.stat{text-align:center}
.stat .sv{font-size:14px;font-weight:700;font-family:'SF Mono',monospace;color:var(--acc)}
.stat .sl{font-size:8px;color:var(--tx3);text-transform:uppercase;letter-spacing:.4px;margin-top:1px}
.stat.warn .sv{color:var(--amb)}
.stat.crit .sv{color:var(--red)}

.conclusion{background:linear-gradient(90deg,rgba(255,124,58,.06),rgba(126,131,250,.06));border:1px solid rgba(255,124,58,.18);border-radius:7px;padding:9px 14px;font-size:11px;line-height:1.9}
.conclusion strong{color:#fff;font-size:12px}
.cp{display:inline-block;padding:1px 6px;border-radius:4px;font-weight:700;font-size:11px;margin:0 2px}
.cp-a{background:rgba(255,124,58,.2);color:var(--org)}
.cp-b{background:rgba(126,131,250,.2);color:var(--blu)}
.cp-ok{background:rgba(46,204,113,.2);color:var(--grn)}
.cp-w{background:rgba(255,184,48,.2);color:var(--amb)}
.cp-r{background:rgba(255,64,96,.2);color:var(--red)}

.empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:50px;color:var(--tx3);text-align:center}
.empty .ei{font-size:44px;opacity:.2}
.empty p{font-size:12px}

/* Compare sub-tabs */
.sub-tabs{display:flex;gap:4px;margin-bottom:8px;flex-wrap:wrap}
.stbtn{padding:4px 10px;border:1px solid var(--brd);border-radius:5px;background:none;color:var(--tx3);cursor:pointer;font-size:10px;transition:all .12s}
.stbtn:hover{color:var(--tx);border-color:var(--tx3)}
.stbtn.act{background:rgba(62,240,175,0.1);border-color:var(--acc);color:var(--acc)}
.sub-panel{display:none;flex-direction:column;gap:10px}
.sub-panel.show{display:flex}

.cmp-legend{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:5px}
.cmp-leg{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--tx2)}
.cmp-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-left">
    <div class="logo">⚙ QSK50 <span>Цилиндры · Давление · Температуры · Топливо · Режимы</span></div>
    <div class="contact-notice">При наличии вопросов и замечаний обращаться к Кузнецову Виктору Евгеньевича</div>
  </div>
  <div class="hdr-info" id="hdr-info">–</div>
</div>

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sb-section">
    <div class="sb-label">
      🚛 Самосвалы
      <button id="cmp-btn" onclick="toggleCmpMode()">⊞ Сравнение</button>
    </div>
  </div>
  <div class="truck-list" id="truck-list"></div>

  <div class="sb-section" id="date-section" style="display:none">
    <div class="sb-label">📅 Дата сессии</div>
    <div class="date-pills" id="date-pills"></div>
  </div>

  <div class="sb-section">
    <div class="sb-label">
      🗓 Календарь
      <button onclick="toggleCal()">▸</button>
    </div>
    <div class="cal-panel" id="cal-panel">
      <div id="cal-grid"></div>
    </div>
  </div>

  <div class="sb-section">
    <div class="sb-label">📂 Добавить CSV</div>
    <div class="file-zone" onclick="document.getElementById('file-in').click()">
      <div class="file-zone-icon">📊</div>
      <div class="file-zone-txt">INSITE Log CSV</div>
      <div class="file-status" id="file-status"></div>
      <input type="file" id="file-in" accept=".csv" multiple style="display:none" onchange="loadFiles(this.files)">
    </div>
  </div>
</div>

<!-- MAIN -->
<div class="main">
  <div class="tabs">
    <button class="tbtn act" onclick="showTab('cylinders',this)">🌡️ Цилиндры</button>
    <button class="tbtn" onclick="showTab('pressure',this)">🔵 Давление</button>
    <button class="tbtn" onclick="showTab('temps',this)">🌡️ Температуры</button>
    <button class="tbtn" onclick="showTab('fuel',this)">⛽ Топливо</button>
    <button class="tbtn" onclick="showTab('modes',this)">⚙️ Режимы</button>
    <button class="tbtn" onclick="showTab('compare',this)">📊 Сравнение</button>
  </div>
  <div class="range-bar">
    <span class="rng-label">⏱ Период:</span>
    <button class="rng-btn act" onclick="applyRange(null,this)">Всё</button>
    <button class="rng-btn" onclick="applyRange(300,this)">5 мин</button>
    <button class="rng-btn" onclick="applyRange(600,this)">10 мин</button>
    <button class="rng-btn" onclick="applyRange(1800,this)">30 мин</button>
    <button class="rng-btn" onclick="applyRange(3600,this)">1 ч</button>
    <div class="rng-slider-wrap">
      <input type="range" id="rng-slider" min="0" max="100" value="100" oninput="applyRangeSlider(+this.value)">
      <span class="rng-val" id="rng-val">100%</span>
    </div>
  </div>
  <div class="content">

    <!-- TAB: CYLINDERS -->
    <div id="tab-cylinders" class="panel show">
      <div class="grid2">
        <div class="card">
          <div class="card-title"><span class="tag tag-a">А-банк</span>Нечётные цилиндры (ECM144)</div>
          <div id="ch-a"></div>
          <div class="stats-row" id="stats-a"></div>
        </div>
        <div class="card">
          <div class="card-title"><span class="tag tag-b">В-банк</span>Чётные цилиндры (ECM1)</div>
          <div id="ch-b"></div>
          <div class="stats-row" id="stats-b"></div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">∆ каждого цилиндра от среднего по банку</div>
        <div id="ch-dev"></div>
      </div>
      <div class="card">
        <div class="card-title">Межбанковая дельта: В-банк avg − А-банк avg</div>
        <div id="ch-delta"></div>
      </div>
      <div class="conclusion" id="conclusion">← Выберите самосвал слева</div>
    </div>

    <!-- TAB: PRESSURE -->
    <div id="tab-pressure" class="panel">
      <div class="grid2">
        <div class="card full">
          <div class="card-title">Давление масла + пред. фильтр + перепад [кПа]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 350 кПа</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 300 кПа</div>
          </div>
          <div id="ch-p-oil"></div>
        </div>
        <div class="card">
          <div class="card-title">Картерные газы [кПа]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 3 кПа</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 7 кПа</div>
          </div>
          <div id="ch-p-crank"></div>
        </div>
        <div class="card">
          <div class="card-title">Наддув: boost + boost1 + boost2 [кПа]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 200 кПа</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 180 кПа</div>
          </div>
          <div id="ch-p-boost"></div>
        </div>
        <div class="card">
          <div class="card-title">Давление топливной рейки: измеренное vs заданное [бар]</div>
          <div id="ch-p-rail"></div>
        </div>
        <div class="card">
          <div class="card-title">Давление масла топливного насоса [кПа]</div>
          <div id="ch-p-pump"></div>
        </div>
        <div class="card">
          <div class="card-title">Атм. давление + охл. жидкость + регулятор [кПа]</div>
          <div id="ch-p-misc"></div>
        </div>
      </div>
    </div>

    <!-- TAB: TEMPERATURES -->
    <div id="tab-temps" class="panel">
      <div class="grid2">
        <div class="card">
          <div class="card-title">Охлаждающая жидкость [°C]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(46,204,113,.6)"></div>Норма 82–95°C</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 95–105°C</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 >105°C</div>
          </div>
          <div id="ch-t-cool"></div>
        </div>
        <div class="card">
          <div class="card-title">Масло [°C]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 105°C</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 115°C</div>
          </div>
          <div id="ch-t-oil"></div>
        </div>
        <div class="card">
          <div class="card-title">Впускной коллектор 4 датчика [°C]</div>
          <div id="ch-t-im4"></div>
        </div>
        <div class="card">
          <div class="card-title">Температура топлива [°C]</div>
          <div id="ch-t-fuel"></div>
        </div>
        <div class="card">
          <div class="card-title">Средняя ТОГ: левый банк + правый банк [°C]</div>
          <div id="ch-t-avg"></div>
        </div>
        <div class="card">
          <div class="card-title">ECM температура 2 + 3 [°C]</div>
          <div id="ch-t-ecm"></div>
        </div>
      </div>
    </div>

    <!-- TAB: FUEL -->
    <div id="tab-fuel" class="panel">
      <div class="grid2">
        <div class="card">
          <div class="card-title">Расход топлива: мгн. + управл. + дозирование [л/час]</div>
          <div id="ch-f-flow"></div>
        </div>
        <div class="card">
          <div class="card-title">Рейка: измеренное vs заданное давление [бар]</div>
          <div id="ch-f-rail"></div>
        </div>
        <div class="card">
          <div class="card-title">Ток насоса: задан. vs изм. [мА]</div>
          <div id="ch-f-curr"></div>
        </div>
        <div class="card">
          <div class="card-title">Рабочий цикл насоса + лифт-насос [%]</div>
          <div id="ch-f-dc"></div>
        </div>
        <div class="card full">
          <div class="card-title">Вода в топливе (0=нет, 1=обнаружена)</div>
          <div id="ch-f-water"></div>
        </div>
      </div>
    </div>

    <!-- TAB: MODES -->
    <div id="tab-modes" class="panel">
      <div class="grid2">
        <div class="card full">
          <div class="card-title">Обороты [RPM] + Нагрузка [%]</div>
          <div id="ch-m-rpm"></div>
        </div>
        <div class="card">
          <div class="card-title">Педаль акселератора [%]</div>
          <div id="ch-m-accel"></div>
        </div>
        <div class="card">
          <div class="card-title">Вентилятор [%]</div>
          <div id="ch-m-fan"></div>
        </div>
        <div class="card">
          <div class="card-title">Напряжение АКБ [В]</div>
          <div class="thres-badges">
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,184,48,.6)"></div>⚠ 22В</div>
            <div class="tb-item"><div class="tb-box" style="background:rgba(255,64,96,.6)"></div>🔴 26В макс.</div>
          </div>
          <div id="ch-m-bat"></div>
        </div>
        <div class="card full">
          <div class="card-title">Лампы аварий: Amber / Red + Уровень ОЖ</div>
          <div id="ch-m-lamps"></div>
        </div>
      </div>
    </div>

    <!-- TAB: COMPARE -->
    <div id="tab-compare" class="panel">
      <div id="cmp-empty" class="empty" style="display:none">
        <div class="ei">📊</div>
        <p>Отметьте ✓ несколько самосвалов в панели слева</p>
      </div>
      <div id="cmp-legend-box" class="cmp-legend"></div>
      <div class="sub-tabs">
        <button class="stbtn act" onclick="showSub('cmp-cyls',this)">🌡️ Цилиндры</button>
        <button class="stbtn" onclick="showSub('cmp-press',this)">🔵 Давление</button>
        <button class="stbtn" onclick="showSub('cmp-temps',this)">🌡️ Температуры</button>
        <button class="stbtn" onclick="showSub('cmp-fuel',this)">⛽ Топливо</button>
        <button class="stbtn" onclick="showSub('cmp-modes',this)">⚙️ Режимы</button>
        <button class="stbtn" onclick="showSub('cmp-cal',this)">🗓 Календарь</button>
      </div>

      <div id="sub-cmp-cyls" class="sub-panel show">
        <div class="card">
          <div class="card-title">Средняя EGT по цилиндрам (все выбранные)</div>
          <div id="ch-cmp-cyls"></div>
        </div>
        <div class="grid2">
          <div class="card">
            <div class="card-title">А-банк avg vs В-банк avg</div>
            <div id="ch-cmp-banks"></div>
          </div>
          <div class="card">
            <div class="card-title">Межбанковая дельта (В−А)</div>
            <div id="ch-cmp-delta"></div>
          </div>
        </div>
        <div class="card">
          <div class="card-title">Сводная таблица</div>
          <div id="ch-cmp-table"></div>
        </div>
      </div>

      <div id="sub-cmp-press" class="sub-panel">
        <div class="grid2">
          <div class="card"><div class="card-title">Давление масла avg [кПа]</div><div id="ch-cmp-oil"></div></div>
          <div class="card"><div class="card-title">Наддув avg [кПа]</div><div id="ch-cmp-boost"></div></div>
          <div class="card"><div class="card-title">Картерные газы avg [кПа]</div><div id="ch-cmp-crank"></div></div>
          <div class="card"><div class="card-title">Рейка измеренная avg [бар]</div><div id="ch-cmp-rail"></div></div>
        </div>
      </div>

      <div id="sub-cmp-temps" class="sub-panel">
        <div class="grid2">
          <div class="card"><div class="card-title">Охл. жидкость avg [°C]</div><div id="ch-cmp-cool"></div></div>
          <div class="card"><div class="card-title">Масло avg [°C]</div><div id="ch-cmp-oil-t"></div></div>
          <div class="card"><div class="card-title">Avg ТОГ левый банк [°C]</div><div id="ch-cmp-egt-l"></div></div>
          <div class="card"><div class="card-title">Avg ТОГ правый банк [°C]</div><div id="ch-cmp-egt-r"></div></div>
        </div>
      </div>

      <div id="sub-cmp-fuel" class="sub-panel">
        <div class="grid2">
          <div class="card"><div class="card-title">Расход мгновенный avg [л/час]</div><div id="ch-cmp-fuel-inst"></div></div>
          <div class="card"><div class="card-title">Ток насоса cmd avg [мА]</div><div id="ch-cmp-pump-ma"></div></div>
        </div>
      </div>

      <div id="sub-cmp-modes" class="sub-panel">
        <div class="grid2">
          <div class="card"><div class="card-title">RPM avg</div><div id="ch-cmp-rpm"></div></div>
          <div class="card"><div class="card-title">Нагрузка avg [%]</div><div id="ch-cmp-load"></div></div>
        </div>
      </div>

      <div id="sub-cmp-cal" class="sub-panel">
        <div class="card">
          <div class="card-title">Флит-календарь: самосвалы × даты (цвет = кол-во точек)</div>
          <div id="ch-cmp-heatmap"></div>
        </div>
      </div>
    </div>

  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════════════════
// DATA
// ═══════════════════════════════════════════════════════════════════
const TRUCKS = __DATA__;
const TRUCK_ORDER = __ORDER__;

const A_CYLS = [1,3,5,7,9,11,13,15];
const B_CYLS = [2,4,6,8,10,12,14,16];
const COLORS_A = ['#ff7c3a','#ffb830','#e55c3a','#cc4020','#ff9860','#f0c000','#dd7060','#aa3010'];
const COLORS_B = ['#5b8fff','#00bfff','#7066ee','#9955dd','#3399ff','#44aaff','#8877ff','#bb77ee'];
const TC = ['#3EF0AF','#7E83FA','#FBAE40','#ff4060','#B668E4','#38bdf8','#34d399','#fb923c',
            '#ff6eb4','#a3e635','#00bfff','#f0c000','#c084fc','#3CF2AE','#f472b6','#facc15'];

const PLY_CFG = {responsive:true,displaylogo:false,modeBarButtonsToRemove:['select2d','lasso2d']};
const PLY_BASE = {
  paper_bgcolor:'#232d34',plot_bgcolor:'rgba(30,38,44,0.9)',
  font:{color:'#ffffff',family:'Arial,sans-serif',size:10},
  margin:{l:50,r:14,t:22,b:36},
  xaxis:{gridcolor:'rgba(62,240,175,0.07)',zerolinecolor:'rgba(62,240,175,0.1)',color:'#80868B'},
  yaxis:{gridcolor:'rgba(62,240,175,0.07)',zerolinecolor:'rgba(62,240,175,0.1)',color:'#80868B'},
  legend:{bgcolor:'rgba(0,0,0,0)',borderwidth:0,font:{size:9}},
};

// ═══ TIME RANGE ════════════════════════════════════════════════════
let activeRangeSec = null;  // null = show all
function applyRange(secs, btn){
  activeRangeSec = secs;
  document.querySelectorAll('.rng-btn').forEach(b=>b.classList.remove('act'));
  if(btn) btn.classList.add('act');
  if(!secs){ const sl=document.getElementById('rng-slider'); if(sl)sl.value=100; const rv=document.getElementById('rng-val'); if(rv)rv.textContent='100%'; }
  syncRangeToCharts();
}
function applyRangeSlider(pct){
  document.getElementById('rng-val').textContent=pct+'%';
  document.querySelectorAll('.rng-btn').forEach(b=>b.classList.remove('act'));
  if(pct>=100){activeRangeSec=null;const rb=document.querySelector('.rng-btn');if(rb)rb.classList.add('act');syncRangeToCharts();return;}
  const s=activeTruck?getSession(activeTruck):null;
  const maxSec=(s&&s.ts&&s.ts.length>1)?timeSecs(s.ts).at(-1)||3600:3600;
  activeRangeSec=Math.round((pct/100)*maxSec);
  syncRangeToCharts();
}
function syncRangeToCharts(){
  const panel=document.querySelector('.panel.show');
  if(!panel)return;
  panel.querySelectorAll('div[id^="ch-"]').forEach(div=>{
    try{
      if(activeRangeSec===null)Plotly.relayout(div,{'xaxis.autorange':true});
      else Plotly.relayout(div,{'xaxis.range':[0,activeRangeSec]});
    }catch(e){}
  });
}

function pl(id,data,lay){
  const L=Object.assign({},PLY_BASE,lay||{});
  if(lay&&lay.xaxis)L.xaxis=Object.assign({},PLY_BASE.xaxis,lay.xaxis);
  if(lay&&lay.yaxis)L.yaxis=Object.assign({},PLY_BASE.yaxis,lay.yaxis);
  if(lay&&lay.yaxis2)L.yaxis2=Object.assign({showgrid:false,color:'#80868B'},lay.yaxis2);
  if(activeRangeSec!==null){L.xaxis=L.xaxis||{};L.xaxis.range=[0,activeRangeSec];L.xaxis.autorange=false;}
  Plotly.react(id,data,L,PLY_CFG);
}
function plEmpty(id,msg,h){pl(id,[],{height:h||180,title:{text:msg||'Нет данных',font:{color:'#80868B',size:10}}});}

// ═══ STATE ════════════════════════════════════════════════════════
let activeTruck=null;
let activeSession={};
let compareMode=false;
let compareSet=new Set();
let currentTab='cylinders';
let currentSub='cmp-cyls';

// ═══ MATH ══════════════════════════════════════════════════════════
function nn(a){return(a||[]).filter(v=>v!==null&&v!==undefined&&!isNaN(v));}
function mean(a){const b=nn(a);return b.length?b.reduce((s,v)=>s+v,0)/b.length:null;}
function maxv(a){const b=nn(a);return b.length?Math.max(...b):null;}
function minv(a){const b=nn(a);return b.length?Math.min(...b):null;}
function rowMeans(cyls,keys){
  if(!cyls||!keys.length)return[];
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
function ewma(arr,a=0.06){
  let s=null;
  return arr.map(v=>{if(v!=null&&!isNaN(v)){s=s==null?v:a*v+(1-a)*s;return Math.round(s*10)/10;}return s;});
}
function mkLine(y,color,dash='dot'){return{type:'line',x0:0,x1:1,xref:'paper',y0:y,y1:y,line:{color,dash,width:1.5}};}
function mkZone(y0,y1,color){return{type:'rect',x0:0,x1:1,xref:'paper',y0,y1,fillcolor:color,line:{width:0}};}

// ═══ SESSION ═══════════════════════════════════════════════════════
function getSession(key){
  const d=TRUCKS[key];if(!d)return null;
  const si=activeSession[key]||0;
  return d.sessions[si]||d.sessions[0]||null;
}
function truckLabel(key){
  const d=TRUCKS[key];if(!d)return key;
  const s=getSession(key);
  return d.model+' №'+d.truck+(s?' | '+s.date:'');
}
function truckColor(key){return TC[TRUCK_ORDER.indexOf(key)%TC.length];}

// ═══ ACTIVE KEYS (single or compare) ══════════════════════════════
function getActiveKeys(){
  if(compareMode&&compareSet.size>=2)return[...compareSet].filter(k=>TRUCKS[k]);
  return activeTruck&&TRUCKS[activeTruck]?[activeTruck]:[];
}

// ═══ GENERIC TIME-SERIES TRACE BUILDER ════════════════════════════
function tsTraces(paramKey,defColor,withEWMA){
  if(withEWMA===undefined)withEWMA=true;
  const keys=getActiveKeys();const traces=[];
  keys.forEach(k=>{
    const s=getSession(k);if(!s||!s.p)return;
    const v=s.p[paramKey];if(!v||nn(v).length<5)return;
    const secs=timeSecs(s.ts);
    const color=keys.length>1?truckColor(k):defColor;
    const name=keys.length>1?truckLabel(k):'Факт';
    traces.push({type:'scattergl',x:secs,y:v,name,line:{color,width:1.2},opacity:keys.length>1?.7:.55});
    if(withEWMA&&keys.length===1)
      traces.push({type:'scattergl',x:secs,y:ewma(v),name:'EWMA',line:{color,width:2.5}});
  });
  return traces;
}
function tsChart(id,paramKey,yTitle,defColor,h,shapes){
  const traces=tsTraces(paramKey,defColor);
  if(!traces.length){plEmpty(id,'Нет данных: '+yTitle,h||180);return;}
  pl(id,traces,{height:h||200,yaxis:{title:yTitle},xaxis:{title:'сек'},shapes:shapes||[]});
}
function tsChart2(id,key1,key2,name1,name2,c1,c2,yTitle,h,shapes){
  const keys=getActiveKeys();const traces=[];
  keys.forEach(k=>{
    const s=getSession(k);if(!s||!s.p)return;
    const v1=s.p[key1],v2=s.p[key2];
    const secs=timeSecs(s.ts);
    const pfx=keys.length>1?(truckLabel(k)+' '):''
    if(v1&&nn(v1).length>4)traces.push({type:'scattergl',x:secs,y:v1,name:pfx+name1,line:{color:c1,width:keys.length>1?1:2},opacity:keys.length>1?.7:.8});
    if(v2&&nn(v2).length>4)traces.push({type:'scattergl',x:secs,y:v2,name:pfx+name2,line:{color:c2,width:keys.length>1?1:2},opacity:keys.length>1?.7:.8});
  });
  if(!traces.length){plEmpty(id,'Нет данных',h||180);return;}
  pl(id,traces,{height:h||200,yaxis:{title:yTitle},xaxis:{title:'сек'},shapes:shapes||[]});
}

// ═══ SIDEBAR ═══════════════════════════════════════════════════════
function buildSidebar(){
  const list=document.getElementById('truck-list');
  list.innerHTML='';
  TRUCK_ORDER.forEach((key,i)=>{
    const d=TRUCKS[key];if(!d)return;
    const item=document.createElement('div');
    item.className='truck-item'+(activeTruck===key?' active':'')+(compareSet.has(key)?' cmp-sel':'');
    item.id='ti-'+key;
    item.onclick=e=>{if(e.target.tagName==='INPUT')return;selectTruck(key);};

    const dot=document.createElement('span');
    dot.className='truck-dot';dot.style.background=TC[i%TC.length];

    const info=document.createElement('div');info.className='truck-info';
    const nm=document.createElement('div');nm.className='truck-name';nm.textContent=d.model+' №'+d.truck;
    const dt=document.createElement('div');dt.className='truck-date';dt.id='td-'+key;
    const s=getSession(key);
    dt.textContent=(s?s.date:'?')+(d.sessions.length>1?' (+'+( d.sessions.length-1)+')':'');
    info.appendChild(nm);info.appendChild(dt);

    const chk=document.createElement('input');
    chk.type='checkbox';chk.className='cmp-chk'+(compareMode?' show':'');
    chk.checked=compareSet.has(key);
    chk.onchange=()=>{
      if(chk.checked)compareSet.add(key);else compareSet.delete(key);
      item.className='truck-item'+(activeTruck===key?' active':'')+(compareSet.has(key)?' cmp-sel':'');
      if(currentTab==='compare')renderCompare();
    };

    item.appendChild(dot);item.appendChild(info);item.appendChild(chk);
    list.appendChild(item);
  });
}

function toggleCmpMode(){
  compareMode=!compareMode;
  document.querySelectorAll('.cmp-chk').forEach(c=>compareMode?c.classList.add('show'):c.classList.remove('show'));
  document.getElementById('cmp-btn').classList.toggle('active',compareMode);
  if(compareMode&&currentTab!=='compare'){
    const btn=document.querySelectorAll('.tbtn')[5];
    showTab('compare',btn);
  }else if(!compareMode){
    renderCurrentTab();
  }
}

function selectTruck(key){
  activeTruck=key;
  document.querySelectorAll('.truck-item').forEach(el=>el.classList.remove('active'));
  const el=document.getElementById('ti-'+key);if(el)el.classList.add('active');
  buildDatePills(key);
  updateCal();
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
      updateCal();
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
    (d?d.model+' №'+d.truck:'')+(s?' | '+s.date+' | '+nn(s.p&&s.p.rpm||[]).length+' точек':'');
}

// ═══ CALENDAR ═══════════════════════════════════════════════════════
function toggleCal(){
  const p=document.getElementById('cal-panel');
  p.classList.toggle('open');
  if(p.classList.contains('open'))buildCalGrid();
}
function buildCalGrid(){
  // Collect all unique dates across all trucks
  const allDates=new Set();
  TRUCK_ORDER.forEach(k=>{
    const d=TRUCKS[k];if(!d)return;
    d.sessions.forEach(s=>allDates.add(s.date));
  });
  const dates=[...allDates].sort();
  const nCols=dates.length+1;
  const grid=document.getElementById('cal-grid');
  grid.style.display='grid';
  grid.style.gridTemplateColumns='70px '+dates.map(()=>'1fr').join(' ');

  let html='<div class="cal-head"></div>';
  dates.forEach(d=>{html+=`<div class="cal-head">${d.slice(0,5)}</div>`;});
  TRUCK_ORDER.forEach(k=>{
    const d=TRUCKS[k];if(!d)return;
    html+=`<div class="cal-truck-name">${d.model} №${d.truck}</div>`;
    dates.forEach(dt=>{
      const si=d.sessions.findIndex(s=>s.date===dt);
      if(si<0){html+=`<div class="cal-cell">–</div>`;}
      else{
        const active=activeTruck===k&&(activeSession[k]||0)===si;
        const pts=nn(d.sessions[si].p&&d.sessions[si].p.rpm||[]).length;
        html+=`<div class="cal-cell has-data${active?' active-date':''}" onclick="selectTruckDate('${k}',${si})">${pts}</div>`;
      }
    });
  });
  grid.innerHTML=html;
}
function selectTruckDate(key,si){
  activeSession[key]=si;
  selectTruck(key);
  buildCalGrid();
}
function updateCal(){
  const p=document.getElementById('cal-panel');
  if(p.classList.contains('open'))buildCalGrid();
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
function showSub(name,btn){
  document.querySelectorAll('.sub-panel').forEach(p=>p.classList.remove('show'));
  document.querySelectorAll('.stbtn').forEach(b=>b.classList.remove('act'));
  document.getElementById('sub-'+name).classList.add('show');
  btn.classList.add('act');
  currentSub=name;
  renderCompare();
}
function renderCurrentTab(){
  if(currentTab==='cylinders')renderCylinders();
  else if(currentTab==='pressure')renderPressure();
  else if(currentTab==='temps')renderTemps();
  else if(currentTab==='fuel')renderFuel();
  else if(currentTab==='modes')renderModes();
  else if(currentTab==='compare')renderCompare();
}

// ═══ STATS ROW ══════════════════════════════════════════════════════
function statsHTML(avg,max,delta){
  const ac=avg>=560?'crit':avg>=520?'warn':'';
  const dc=delta>=200?'crit':delta>=100?'warn':'';
  return `<div class="stat ${ac}"><div class="sv">${Math.round(avg)}°C</div><div class="sl">Avg банка</div></div>
          <div class="stat ${ac}"><div class="sv">${Math.round(max)}°C</div><div class="sl">Max</div></div>
          <div class="stat ${dc}"><div class="sv">${Math.round(delta)}°C</div><div class="sl">∆ внутри</div></div>`;
}

// ═══ TAB: CYLINDERS ═════════════════════════════════════════════════
function renderCylinders(){
  const keys=getActiveKeys();
  if(!keys.length){
    ['ch-a','ch-b','ch-dev','ch-delta'].forEach(id=>Plotly.purge(id));
    document.getElementById('conclusion').innerHTML='← Выберите самосвал слева';
    return;
  }

  if(keys.length===1){
    const key=keys[0];
    const s=getSession(key);if(!s)return;
    const{ts,cyls}=s;
    const secs=timeSecs(ts);
    const aCyls=A_CYLS.filter(n=>cyls[String(n)]);
    const bCyls=B_CYLS.filter(n=>cyls[String(n)]);
    const aAvg=aCyls.length?rowMeans(cyls,aCyls.map(String)):[];
    const bAvg=bCyls.length?rowMeans(cyls,bCyls.map(String)):[];
    const thShapes=[mkLine(500,'rgba(255,184,48,.5)'),mkLine(550,'rgba(255,64,96,.5)')];

    if(aCyls.length){
      const tr=aCyls.map((n,i)=>({type:'scattergl',x:secs,y:cyls[String(n)],name:'Ц'+n,line:{color:COLORS_A[i%8],width:1},opacity:.8}));
      tr.push({type:'scattergl',x:secs,y:ewma(aAvg),name:'avg А',line:{color:'#fff',width:2.5,dash:'dot'}});
      pl('ch-a',tr,{height:250,yaxis:{title:'EGT °C',range:[250,650]},xaxis:{title:'сек'},shapes:thShapes});
      const aVals=aCyls.flatMap(n=>nn(cyls[String(n)]));
      document.getElementById('stats-a').innerHTML=statsHTML(mean(aAvg)||0,maxv(aVals)||0,(maxv(aVals)||0)-(minv(aVals)||0));
    }
    if(bCyls.length){
      const tr=bCyls.map((n,i)=>({type:'scattergl',x:secs,y:cyls[String(n)],name:'Ц'+n,line:{color:COLORS_B[i%8],width:1},opacity:.8}));
      tr.push({type:'scattergl',x:secs,y:ewma(bAvg),name:'avg В',line:{color:'#fff',width:2.5,dash:'dot'}});
      pl('ch-b',tr,{height:250,yaxis:{title:'EGT °C',range:[250,650]},xaxis:{title:'сек'},shapes:thShapes});
      const bVals=bCyls.flatMap(n=>nn(cyls[String(n)]));
      document.getElementById('stats-b').innerHTML=statsHTML(mean(bAvg)||0,maxv(bVals)||0,(maxv(bVals)||0)-(minv(bVals)||0));
    }

    // Deviation bar
    const aAvgM=mean(aAvg)||0,bAvgM=mean(bAvg)||0;
    const allC=[...aCyls,...bCyls].sort((a,b)=>a-b);
    const devs=allC.map(n=>{
      const v=mean(cyls[String(n)])||0;
      const bk=A_CYLS.includes(n)?aAvgM:bAvgM;
      return{n,d:Math.round(v-bk),bank:A_CYLS.includes(n)?'A':'B'};
    });
    pl('ch-dev',[{type:'bar',x:devs.map(d=>'Ц'+d.n),y:devs.map(d=>d.d),
      text:devs.map(d=>(d.d>=0?'+':'')+d.d+'°'),textposition:'outside',textfont:{size:9},
      marker:{color:devs.map(d=>{const ab=Math.abs(d.d),base=d.bank==='A'?'255,124,58':'91,143,255',op=ab>60?.95:ab>30?.7:.45;return`rgba(${base},${op})`;})},
    }],{height:185,
      shapes:[mkLine(30,'rgba(255,184,48,.4)'),mkLine(-30,'rgba(255,184,48,.4)'),mkLine(60,'rgba(255,64,96,.4)'),mkLine(-60,'rgba(255,64,96,.4)')],
      yaxis:{title:'∆°C',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5},
      xaxis:{tickfont:{size:9}},margin:{l:48,r:40,t:14,b:36}});

    // Delta time
    const delta=secs.map((_,i)=>{
      const a=aAvg[i],b=bAvg[i];return(a!=null&&b!=null)?Math.round(b-a):null;});
    const dm=mean(delta);
    pl('ch-delta',[
      {type:'scattergl',x:secs,y:delta.map(v=>v!=null&&v>0?v:0),name:'В>А',fill:'tozeroy',fillcolor:'rgba(126,131,250,.15)',line:{color:'rgba(126,131,250,.6)',width:1}},
      {type:'scattergl',x:secs,y:delta.map(v=>v!=null&&v<0?v:0),name:'А>В',fill:'tozeroy',fillcolor:'rgba(255,124,58,.15)',line:{color:'rgba(255,124,58,.6)',width:1}},
      {type:'scattergl',x:secs,y:ewma(delta,.05),name:'EWMA',line:{color:'#ffb830',width:2.5}},
    ],{height:185,yaxis:{title:'∆°C (В−А)',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5},xaxis:{title:'сек'},
      shapes:[mkLine(20,'rgba(126,131,250,.35)'),mkLine(-20,'rgba(255,124,58,.35)')]});

    // Conclusion
    const hotBank=dm!=null?(dm>0?'В-банк':'А-банк'):'?',hotAmt=dm!=null?Math.abs(Math.round(dm)):0;
    const aValsAll=aCyls.flatMap(n=>nn(cyls[String(n)]));
    const bValsAll=bCyls.flatMap(n=>nn(cyls[String(n)]));
    const aSpread=Math.round((maxv(aValsAll)||0)-(minv(aValsAll)||0));
    const bSpread=Math.round((maxv(bValsAll)||0)-(minv(bValsAll)||0));
    let maxDevC=null,maxDevV=0;
    allC.forEach(n=>{const v=mean(cyls[String(n)])||0;const bk=A_CYLS.includes(n)?aAvgM:bAvgM;const dv=Math.abs(v-bk);if(dv>maxDevV){maxDevV=dv;maxDevC={n,d:Math.round(v-bk)};}});
    const riskC=maxDevC&&Math.abs(maxDevC.d)>60?'cp-r':maxDevC&&Math.abs(maxDevC.d)>30?'cp-w':'cp-ok';
    document.getElementById('conclusion').innerHTML=`
      <strong>ВЫВОД: ${truckLabel(keys[0])}</strong><br>
      <span class="cp ${dm>=0?'cp-b':'cp-a'}">${hotBank} горячее</span> на <b>${hotAmt}°С</b> &nbsp;|&nbsp;
      Разброс: <span class="cp cp-a">A=${aSpread}°</span><span class="cp cp-b">B=${bSpread}°</span> &nbsp;|&nbsp;
      ${maxDevC?`Цил.<b>${maxDevC.n}</b>: <span class="cp ${riskC}">${maxDevC.d>0?'+':''}${maxDevC.d}°</span> от avg банка`:''}`;

  }else{
    // Compare mode: grouped EGT bar per truck
    const allCyls=[...A_CYLS,...B_CYLS].sort((a,b)=>a-b);
    const trCylAvgs=keys.map(k=>{
      const s=getSession(k);if(!s)return{};
      const out={};allCyls.forEach(n=>{out[n]=mean(s.cyls&&s.cyls[String(n)]);});
      return out;
    });
    const cylTraces=keys.map((k,i)=>({
      type:'bar',name:truckLabel(k),
      x:allCyls.map(n=>'Ц'+n),
      y:allCyls.map(n=>trCylAvgs[i][n]!=null?Math.round(trCylAvgs[i][n]):null),
      marker:{color:truckColor(k)},
    }));
    pl('ch-a',cylTraces,{height:260,barmode:'group',
      yaxis:{title:'°C'},xaxis:{tickfont:{size:9}},
      shapes:[mkLine(500,'rgba(255,184,48,.35)'),mkLine(550,'rgba(255,64,96,.35)')]});
    plEmpty('ch-b','(см. график слева — режим сравнения)',200);
    // Banks
    const bankX=keys.map(k=>truckLabel(k));
    pl('ch-dev',[
      {type:'bar',name:'А-банк',x:bankX,y:keys.map((_,i)=>Math.round(mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),marker:{color:'rgba(255,124,58,.8)'}},
      {type:'bar',name:'В-банк',x:bankX,y:keys.map((_,i)=>Math.round(mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),marker:{color:'rgba(126,131,250,.8)'}},
    ],{height:185,barmode:'group',yaxis:{title:'°C'},shapes:[mkLine(500,'rgba(255,184,48,.25)'),mkLine(550,'rgba(255,64,96,.25)')]});
    pl('ch-delta',[{type:'bar',
      x:bankX,
      y:keys.map((_,i)=>{const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;return Math.round(b-a);}),
      marker:{color:keys.map((_,i)=>{const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;return(b-a)>=0?'rgba(126,131,250,.8)':'rgba(255,124,58,.8)';})}
    }],{height:185,yaxis:{title:'В−А °C',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5}});
    document.getElementById('conclusion').innerHTML='<strong>Режим сравнения:</strong> выбрано '+keys.length+' самосвалов';
  }
}

// ═══ TAB: PRESSURE ══════════════════════════════════════════════════
function renderPressure(){
  // Oil pressure + pre-filter + DP
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      const col=keys.length>1?truckColor(k):'#00d4aa';
      if(s.p.oil_p&&nn(s.p.oil_p).length>4){
        traces.push({type:'scattergl',x:secs,y:s.p.oil_p,name:pfx+'P масла',line:{color:col,width:keys.length>1?1:1.2},opacity:.55});
        if(keys.length===1)traces.push({type:'scattergl',x:secs,y:ewma(s.p.oil_p),name:'EWMA масло',line:{color:col,width:2.5}});
      }
      if(keys.length===1){
        if(s.p.pre_filt_p&&nn(s.p.pre_filt_p).length>4)traces.push({type:'scattergl',x:secs,y:s.p.pre_filt_p,name:'Пред.фильтр',line:{color:'#5b8fff',width:2}});
        if(s.p.filt_dp&&nn(s.p.filt_dp).length>4)traces.push({type:'scattergl',x:secs,y:s.p.filt_dp,name:'∆P фильтра',line:{color:'#ffb830',width:2,dash:'dash'}});
      }
    });
    if(!traces.length){plEmpty('ch-p-oil','Нет данных давления масла',200);return;}
    pl('ch-p-oil',traces,{height:200,yaxis:{title:'кПа'},xaxis:{title:'сек'},
      shapes:[mkZone(350,650,'rgba(46,204,113,.025)'),mkLine(350,'rgba(255,184,48,.5)'),mkLine(300,'rgba(255,64,96,.5)')]});
  })();

  // Crankcase
  tsChart('ch-p-crank','crank_p','кПа','#ff4060',200,[
    mkZone(0,3,'rgba(46,204,113,.04)'),mkLine(3,'rgba(255,184,48,.5)'),mkLine(7,'rgba(255,64,96,.5)')]);

  // Boost
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const col=keys.length>1?truckColor(k):'#5b8fff';
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.boost&&nn(s.p.boost).length>4){
        traces.push({type:'scattergl',x:secs,y:s.p.boost,name:pfx+'boost',line:{color:col,width:keys.length>1?1:1.2},opacity:.55});
        if(keys.length===1)traces.push({type:'scattergl',x:secs,y:ewma(s.p.boost),name:'EWMA',line:{color:col,width:2.5}});
      }
      if(keys.length===1){
        if(s.p.boost1&&nn(s.p.boost1).length>4)traces.push({type:'scattergl',x:secs,y:s.p.boost1,name:'boost1',line:{color:'#00bfff',width:1.5,dash:'dash'}});
        if(s.p.boost2&&nn(s.p.boost2).length>4)traces.push({type:'scattergl',x:secs,y:s.p.boost2,name:'boost2',line:{color:'#9955dd',width:1.5,dash:'dot'}});
      }
    });
    if(!traces.length){plEmpty('ch-p-boost','Нет данных наддува',200);return;}
    pl('ch-p-boost',traces,{height:200,yaxis:{title:'кПа'},xaxis:{title:'сек'},
      shapes:[mkZone(200,350,'rgba(46,204,113,.03)'),mkLine(200,'rgba(255,184,48,.5)'),mkLine(180,'rgba(255,64,96,.5)')]});
  })();

  // Rail meas vs cmd
  tsChart2('ch-p-rail','rail_meas','rail_cmd','Изм.','Зад.','#00d4aa','#ffb830','бар',200);

  // Pump oil pressure
  tsChart('ch-p-pump','pump_oil_p','кПа','#b06bff',200);

  // Misc
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.baro_p&&nn(s.p.baro_p).length>4)traces.push({type:'scattergl',x:secs,y:s.p.baro_p,name:pfx+'Атм.',line:{color:'#a3e635',width:1.5}});
      if(s.p.cool_p&&nn(s.p.cool_p).length>4)traces.push({type:'scattergl',x:secs,y:s.p.cool_p,name:pfx+'ОЖ',line:{color:'#00d4aa',width:1.5}});
      if(s.p.fuel_reg_p&&nn(s.p.fuel_reg_p).length>4)traces.push({type:'scattergl',x:secs,y:s.p.fuel_reg_p,name:pfx+'Регулятор',line:{color:'#ff7c3a',width:1.5}});
    });
    if(!traces.length){plEmpty('ch-p-misc','Нет данных',200);return;}
    pl('ch-p-misc',traces,{height:200,yaxis:{title:'кПа'},xaxis:{title:'сек'}});
  })();
}

// ═══ TAB: TEMPERATURES ══════════════════════════════════════════════
function renderTemps(){
  tsChart('ch-t-cool','cool_t','°C','#00d4aa',200,[
    mkZone(82,95,'rgba(46,204,113,.04)'),mkZone(95,105,'rgba(255,184,48,.04)'),
    mkLine(95,'rgba(255,184,48,.5)'),mkLine(105,'rgba(255,64,96,.5)')]);

  tsChart('ch-t-oil','oil_t','°C','#ff7c3a',200,[
    mkLine(105,'rgba(255,184,48,.5)'),mkLine(115,'rgba(255,64,96,.5)')]);

  // IM4 temps
  (()=>{
    const keys=getActiveKeys();const traces=[];
    const imColors=['#00d4aa','#5b8fff','#ffb830','#ff4060'];
    const imKeys=['im_t1','im_t2','im_t3','im_t4'];
    const imNames=['IM1','IM2','IM3','IM4'];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      imKeys.forEach((pk,i)=>{
        if(s.p[pk]&&nn(s.p[pk]).length>4)
          traces.push({type:'scattergl',x:secs,y:s.p[pk],name:pfx+imNames[i],line:{color:keys.length>1?truckColor(k):imColors[i],width:1.5},opacity:.8});
      });
    });
    if(!traces.length){plEmpty('ch-t-im4','Нет данных IM',200);return;}
    pl('ch-t-im4',traces,{height:200,yaxis:{title:'°C'},xaxis:{title:'сек'}});
  })();

  tsChart('ch-t-fuel','fuel_t','°C','#ffb830',200);

  tsChart2('ch-t-avg','avg_egt_l','avg_egt_r','Лев. банк','Пр. банк','#ff7c3a','#5b8fff','°C',200);
  tsChart2('ch-t-ecm','ecm_t2','ecm_t3','ECM T2','ECM T3','#00d4aa','#b06bff','°C',200);
}

// ═══ TAB: FUEL ══════════════════════════════════════════════════════
function renderFuel(){
  // Flow 3 lines
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.fuel_inst&&nn(s.p.fuel_inst).length>4)traces.push({type:'scattergl',x:secs,y:s.p.fuel_inst,name:pfx+'Мгн.',line:{color:'#00d4aa',width:1.5}});
      if(s.p.fuel_flow&&nn(s.p.fuel_flow).length>4)traces.push({type:'scattergl',x:secs,y:s.p.fuel_flow,name:pfx+'Упр.',line:{color:'#5b8fff',width:1.5,dash:'dash'}});
      if(s.p.fuel_dos&&nn(s.p.fuel_dos).length>4)traces.push({type:'scattergl',x:secs,y:s.p.fuel_dos,name:pfx+'Доз.',line:{color:'#ffb830',width:1.5,dash:'dot'}});
    });
    if(!traces.length){plEmpty('ch-f-flow','Нет данных расхода',200);return;}
    pl('ch-f-flow',traces,{height:200,yaxis:{title:'л/час'},xaxis:{title:'сек'}});
  })();

  tsChart2('ch-f-rail','rail_meas','rail_cmd','Изм.','Зад.','#00d4aa','#ffb830','бар',200);
  tsChart2('ch-f-curr','pump_mA_cmd','pump_mA_meas','Зад.','Изм.','#5b8fff','#ff4060','мА',200);
  tsChart2('ch-f-dc','pump_dc','lift_dc','Насос','Лифт','#00d4aa','#b06bff','%',200);

  // Water (step chart)
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const col=keys.length>1?truckColor(k):'#ff4060';
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.water_fuel&&nn(s.p.water_fuel).length>4)
        traces.push({type:'scatter',x:secs,y:s.p.water_fuel,name:pfx+'Вода',mode:'lines',line:{color:col,width:2,shape:'hv'},fill:'tozeroy',fillcolor:col.replace(/[^#\d]/,'')+'22'});
    });
    if(!traces.length){plEmpty('ch-f-water','Нет данных (вода в топливе)',120);return;}
    pl('ch-f-water',traces,{height:120,yaxis:{title:'0/1',range:[-0.1,1.5],tickvals:[0,1],ticktext:['Нет','Обнаружена']},xaxis:{title:'сек'}});
  })();
}

// ═══ TAB: MODES ═════════════════════════════════════════════════════
function renderModes(){
  // RPM + Load dual-axis
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const col=keys.length>1?truckColor(k):'#00d4aa';
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.rpm&&nn(s.p.rpm).length>4){
        traces.push({type:'scattergl',x:secs,y:keys.length===1?ewma(s.p.rpm,.1):s.p.rpm,name:pfx+'RPM',line:{color:col,width:2},yaxis:'y'});
      }
      if(s.p.load&&nn(s.p.load).length>4){
        traces.push({type:'scattergl',x:secs,y:keys.length===1?ewma(s.p.load,.1):s.p.load,name:pfx+'Нагрузка%',line:{color:'#ffb830',width:2,dash:keys.length>1?'dot':'solid'},yaxis:'y2'});
      }
    });
    if(!traces.length){plEmpty('ch-m-rpm','Нет данных оборотов',200);return;}
    pl('ch-m-rpm',traces,{height:200,yaxis:{title:'RPM'},yaxis2:{title:'Нагрузка %',side:'right',overlaying:'y'},xaxis:{title:'сек'}});
  })();

  tsChart('ch-m-accel','accel','%','#a3e635',180);
  tsChart('ch-m-fan','fan_cmd','%','#38bdf8',180);
  tsChart('ch-m-bat','battery','В','#ffb830',180,[mkLine(22,'rgba(255,184,48,.5)'),mkLine(26,'rgba(255,64,96,.5)')]);

  // Lamps step chart
  (()=>{
    const keys=getActiveKeys();const traces=[];
    keys.forEach(k=>{
      const s=getSession(k);if(!s||!s.p)return;
      const secs=timeSecs(s.ts);
      const pfx=keys.length>1?(truckLabel(k)+' '):'';
      if(s.p.amber_lamp&&nn(s.p.amber_lamp).length>4)
        traces.push({type:'scatter',x:secs,y:s.p.amber_lamp,name:pfx+'Amber',mode:'lines',line:{color:'#ffb830',width:2,shape:'hv'}});
      if(s.p.red_lamp&&nn(s.p.red_lamp).length>4)
        traces.push({type:'scatter',x:secs,y:s.p.red_lamp.map(v=>v!=null?v*1.5:null),name:pfx+'Red',mode:'lines',line:{color:'#ff4060',width:2,shape:'hv'}});
      if(s.p.cool_level&&nn(s.p.cool_level).length>4)
        traces.push({type:'scatter',x:secs,y:s.p.cool_level.map(v=>v!=null?v*0.8:null),name:pfx+'ОЖ уровень',mode:'lines',line:{color:'#00d4aa',width:2,shape:'hv'}});
    });
    if(!traces.length){plEmpty('ch-m-lamps','Нет данных ламп',120);return;}
    pl('ch-m-lamps',traces,{height:120,yaxis:{title:'0/1',range:[-0.1,2]},xaxis:{title:'сек'}});
  })();
}

// ═══ TAB: COMPARE ═══════════════════════════════════════════════════
function renderCompare(){
  const keys=[...compareSet].filter(k=>TRUCKS[k]);
  const empty=document.getElementById('cmp-empty');
  if(keys.length<2){empty.style.display='';document.getElementById('cmp-legend-box').innerHTML='';return;}
  empty.style.display='none';

  const legBox=document.getElementById('cmp-legend-box');
  legBox.innerHTML=keys.map((k,i)=>`<div class="cmp-leg"><div class="cmp-dot" style="background:${truckColor(k)}"></div>${truckLabel(k)}</div>`).join('');

  const allCyls=[...A_CYLS,...B_CYLS].sort((a,b)=>a-b);

  if(currentSub==='cmp-cyls'){
    const trCylAvgs=keys.map(k=>{
      const s=getSession(k);if(!s)return{};
      const out={};allCyls.forEach(n=>{out[n]=mean(s.cyls&&s.cyls[String(n)]);});return out;
    });
    const cylTraces=keys.map((k,i)=>({
      type:'bar',name:truckLabel(k),
      x:allCyls.map(n=>'Ц'+n),
      y:allCyls.map(n=>trCylAvgs[i][n]!=null?Math.round(trCylAvgs[i][n]):null),
      marker:{color:truckColor(k)},
    }));
    pl('ch-cmp-cyls',cylTraces,{height:280,barmode:'group',yaxis:{title:'°C'},xaxis:{tickfont:{size:9}},
      shapes:[mkLine(500,'rgba(255,184,48,.35)'),mkLine(550,'rgba(255,64,96,.35)')]});

    const bankX=keys.map(k=>truckLabel(k));
    pl('ch-cmp-banks',[
      {type:'bar',name:'А-банк',x:bankX,y:keys.map((_,i)=>Math.round(mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),marker:{color:'rgba(255,124,58,.8)'}},
      {type:'bar',name:'В-банк',x:bankX,y:keys.map((_,i)=>Math.round(mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0)),marker:{color:'rgba(126,131,250,.8)'}},
    ],{height:220,barmode:'group',yaxis:{title:'°C'},shapes:[mkLine(500,'rgba(255,184,48,.25)')]});

    pl('ch-cmp-delta',[{type:'bar',
      x:bankX,
      y:keys.map((_,i)=>{const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;return Math.round(b-a);}),
      marker:{color:keys.map((_,i)=>{const a=mean(A_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;const b=mean(B_CYLS.map(n=>trCylAvgs[i][n]).filter(v=>v!=null))||0;return(b-a)>=0?'rgba(126,131,250,.8)':'rgba(255,124,58,.8)';})}
    }],{height:220,yaxis:{title:'В−А °C',zeroline:true,zerolinecolor:'rgba(255,255,255,.2)',zerolinewidth:1.5}});

    const tHead=['Самосвал','Дата','A avg°','B avg°','∆A','∆B','B−A','Max цил.'];
    const tRows=keys.map((k,i)=>{
      const s=getSession(k);const ca=trCylAvgs[i];
      const aAvgM=mean(A_CYLS.map(n=>ca[n]).filter(v=>v!=null))||0;
      const bAvgM=mean(B_CYLS.map(n=>ca[n]).filter(v=>v!=null))||0;
      const aVals=A_CYLS.map(n=>ca[n]).filter(v=>v!=null);
      const bVals=B_CYLS.map(n=>ca[n]).filter(v=>v!=null);
      let maxCyl=null,maxDev=0;
      allCyls.forEach(n=>{if(ca[n]==null)return;const bk=A_CYLS.includes(n)?aAvgM:bAvgM;const dv=Math.abs(ca[n]-bk);if(dv>maxDev){maxDev=dv;maxCyl=n;}});
      return[truckLabel(k),s?s.date:'?',Math.round(aAvgM),Math.round(bAvgM),
        Math.round((maxv(aVals)||0)-(minv(aVals)||0)),
        Math.round((maxv(bVals)||0)-(minv(bVals)||0)),
        (bAvgM-aAvgM>0?'+':'')+Math.round(bAvgM-aAvgM),
        maxCyl?'Ц'+maxCyl+' (∆'+Math.round(ca[maxCyl]-(A_CYLS.includes(maxCyl)?aAvgM:bAvgM))+'°)':'—'];
    });
    pl('ch-cmp-table',[{type:'table',
      header:{values:tHead,fill:{color:'#1e262c'},font:{color:'#3EF0AF',size:10},align:'left'},
      cells:{values:tHead.map((_,ci)=>tRows.map(r=>r[ci])),fill:{color:'#232d34'},font:{color:'#ffffff',size:10},align:'left'},
    }],{height:Math.max(180,60+keys.length*28),margin:{l:4,r:4,t:4,b:4}});

  }else if(currentSub==='cmp-press'){
    function cmpBar(id,pkey,unit){
      const vals=keys.map(k=>{const s=getSession(k);return s&&s.p&&s.p[pkey]?Math.round((mean(s.p[pkey])||0)*10)/10:null;});
      pl(id,[{type:'bar',x:keys.map(k=>truckLabel(k)),y:vals,marker:{color:keys.map(k=>truckColor(k))}}],
        {height:200,yaxis:{title:unit},xaxis:{tickfont:{size:8}}});
    }
    cmpBar('ch-cmp-oil','oil_p','кПа');
    cmpBar('ch-cmp-boost','boost','кПа');
    cmpBar('ch-cmp-crank','crank_p','кПа');
    cmpBar('ch-cmp-rail','rail_meas','бар');

  }else if(currentSub==='cmp-temps'){
    function cmpBar(id,pkey,unit){
      const vals=keys.map(k=>{const s=getSession(k);return s&&s.p&&s.p[pkey]?Math.round((mean(s.p[pkey])||0)*10)/10:null;});
      pl(id,[{type:'bar',x:keys.map(k=>truckLabel(k)),y:vals,marker:{color:keys.map(k=>truckColor(k))}}],
        {height:200,yaxis:{title:unit},xaxis:{tickfont:{size:8}}});
    }
    cmpBar('ch-cmp-cool','cool_t','°C');
    cmpBar('ch-cmp-oil-t','oil_t','°C');
    cmpBar('ch-cmp-egt-l','avg_egt_l','°C');
    cmpBar('ch-cmp-egt-r','avg_egt_r','°C');

  }else if(currentSub==='cmp-fuel'){
    function cmpBar(id,pkey,unit){
      const vals=keys.map(k=>{const s=getSession(k);return s&&s.p&&s.p[pkey]?Math.round((mean(s.p[pkey])||0)*10)/10:null;});
      pl(id,[{type:'bar',x:keys.map(k=>truckLabel(k)),y:vals,marker:{color:keys.map(k=>truckColor(k))}}],
        {height:200,yaxis:{title:unit},xaxis:{tickfont:{size:8}}});
    }
    cmpBar('ch-cmp-fuel-inst','fuel_inst','л/час');
    cmpBar('ch-cmp-pump-ma','pump_mA_cmd','мА');

  }else if(currentSub==='cmp-modes'){
    function cmpBar(id,pkey,unit){
      const vals=keys.map(k=>{const s=getSession(k);return s&&s.p&&s.p[pkey]?Math.round((mean(s.p[pkey])||0)*10)/10:null;});
      pl(id,[{type:'bar',x:keys.map(k=>truckLabel(k)),y:vals,marker:{color:keys.map(k=>truckColor(k))}}],
        {height:200,yaxis:{title:unit},xaxis:{tickfont:{size:8}}});
    }
    cmpBar('ch-cmp-rpm','rpm','RPM');
    cmpBar('ch-cmp-load','load','%');

  }else if(currentSub==='cmp-cal'){
    // Fleet heatmap
    const allDates=new Set();
    TRUCK_ORDER.forEach(k=>{const d=TRUCKS[k];if(d)d.sessions.forEach(s=>allDates.add(s.date));});
    const dates=[...allDates].sort();
    const truckNames=TRUCK_ORDER.map(k=>{const d=TRUCKS[k];return d?d.model+' №'+d.truck:'?';});
    const z=TRUCK_ORDER.map(k=>{
      const d=TRUCKS[k];if(!d)return dates.map(()=>0);
      return dates.map(dt=>{
        const s=d.sessions.find(ss=>ss.date===dt);
        return s?nn(s.p&&s.p.rpm||[]).length:0;
      });
    });
    pl('ch-cmp-heatmap',[{type:'heatmap',x:dates,y:truckNames,z,
      colorscale:[[0,'#0f1e30'],[0.001,'rgba(62,240,175,0.3)'],[1,'#00d4aa']],
      hoverongaps:false,showscale:true,
    }],{height:Math.max(200,TRUCK_ORDER.length*24+60),
      margin:{l:100,r:60,t:20,b:60},
      xaxis:{tickangle:-30},yaxis:{autorange:'reversed'}});
  }
}

// ═══ FILE LOADING ════════════════════════════════════════════════════
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
function findCylColBr(cols,n){
  return cols.find(c=>c.includes('Exhaust Temperature Sensor Cylinder '+n+' (')||c.includes('отработавших газов цилиндра '+n+' (°C)'))||null;
}
const PARAM_PATS_BR = __PARAM_PATS_BROWSER__;
function findParamColBr(cols,patterns){
  for(const p of patterns){const c=cols.find(x=>x.includes(p));if(c)return c;}return null;
}
function processINSITE(text,fname){
  const parsed=parseINSITE(text);if(!parsed)return null;
  const{cols,rows,dateC,timeC}=parsed;
  const cylCols={};for(let n=1;n<=16;n++){const c=findCylColBr(cols,n);if(c)cylCols[n]=c;}
  if(Object.keys(cylCols).length<8)return null;
  const paramCols={};
  for(const[k,pats]of Object.entries(PARAM_PATS_BR)){const c=findParamColBr(cols,pats);if(c)paramCols[k]=c;}
  const BOOL_KEYS=new Set(['water_fuel','amber_lamp','red_lamp','cool_level']);
  function boolVal(v){const sv=v.trim().toLowerCase();if(['on','yes','true','1','low','active','detected'].includes(sv))return 1;if(['off','no','false','0','ok','normal','not detected'].includes(sv))return 0;const n=parseFloat(sv.replace(',','.'));return isNaN(n)?null:Math.round(n);}
  const step=Math.max(1,Math.floor(rows.length/800));
  const out={ts:[],cyls:{},p:{}};
  for(let n=1;n<=16;n++)if(cylCols[n])out.cyls[String(n)]=[];
  for(const k of Object.keys(paramCols))out.p[k]=[];
  for(let i=0;i<rows.length;i+=step){
    const r=rows[i];
    let d=ruDate2(r[dateC]||''),t=(r[timeC]||'').substring(0,8);
    try{const dt=new Date(d+' '+t);if(isNaN(dt))continue;out.ts.push(dt.toISOString().substring(0,19));}catch{continue;}
    for(let n=1;n<=16;n++){if(!cylCols[n])continue;const v=parseFloat((r[cylCols[n]]||'').replace(',','.'));out.cyls[String(n)].push(isNaN(v)?null:Math.round(v));}
    for(const[k,c]of Object.entries(paramCols)){const raw=(r[c]||'').replace(',','.');const v=BOOL_KEYS.has(k)?boolVal(r[c]||''):parseFloat(raw);out.p[k].push(isNaN(v)?null:BOOL_KEYS.has(k)?v:Math.round(v*100)/100);}
  }
  if(out.ts.length<10)return null;
  const raw=out.ts[0]?.substring(0,10)||'';const pp=raw.split('-');const dateStr=pp.length===3?`${pp[2]}.${pp[1]}.${pp[0]}`:raw;
  let truck=null;
  for(const pat of[/№\s*(\d+)/,/N(\d+)\b/,/[\s_-](\d{2,3})[\s_.(]/]){const m=fname.match(pat);if(m){truck=+m[1];break;}}
  if(!truck)truck=Math.floor(Math.random()*900)+100;
  const model=fname.includes('NTE')||fname.includes('DML')?'NTE200':'730E';
  const key=model+'_'+truck;
  return{key,truck,model,session:{date:dateStr,ts:out.ts,cyls:out.cyls,p:out.p,file:fname}};
}
function loadFiles(files){
  let loaded=0;const total=files.length;
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
          if(!activeTruck)selectTruck(key);
        }
      }catch(ex){console.warn(file.name,ex);}
    };
    reader.readAsText(file,'windows-1251');
  });
}

// ═══ INIT ════════════════════════════════════════════════════════════
window.addEventListener('load',function(){
  buildSidebar();
  if(TRUCK_ORDER.length){
    activeTruck=TRUCK_ORDER[0];
    const el=document.getElementById('ti-'+activeTruck);if(el)el.classList.add('active');
    buildDatePills(activeTruck);
    updateHeader();
    renderCylinders();
  }
});
</script>
</body>
</html>
"""


def build_html(db_data, truck_order):
    plotly_src = PLOTLY_JS.read_text(encoding='utf-8')

    # Build browser-side pattern map (same patterns as Python)
    param_pats_browser = {k: v for k, v in PARAM_PATTERNS.items()}
    param_pats_json = json.dumps(param_pats_browser, ensure_ascii=False, separators=(',', ':'))

    data_json  = json.dumps(db_data,     ensure_ascii=False, separators=(',', ':'))
    order_json = json.dumps(truck_order, ensure_ascii=False)

    html = HTML.replace('__PLOTLY__', plotly_src) \
               .replace('__DATA__',   data_json) \
               .replace('__ORDER__',  order_json) \
               .replace('__PARAM_PATS_BROWSER__', param_pats_json)
    return html


def main():
    print('=== QSK50 Comprehensive Dashboard Builder ===\n')
    db_data, truck_order = build_data()
    if not db_data:
        print('ERROR: No data found!')
        return
    total_rows = sum(sum(len(s['ts']) for s in v['sessions']) for v in db_data.values())
    total_params = sum(sum(len(s.get('p', {})) for s in v['sessions']) for v in db_data.values())
    kb = len(json.dumps(db_data, separators=(',', ':')).encode()) // 1024
    print(f'\n{len(db_data)} trucks, {total_rows} rows, {total_params} param-sets, data JSON: {kb} KB')

    # Verify JSON
    data_str = json.dumps(db_data, ensure_ascii=False)
    try:
        json.loads(data_str)
        print('JSON validation: OK')
    except Exception as e:
        print(f'JSON validation FAILED: {e}')
        return

    html = build_html(db_data, truck_order)
    OUTPUT.write_text(html, encoding='utf-8')
    mb = OUTPUT.stat().st_size / 1024 / 1024
    print(f'Written: {OUTPUT}  ({mb:.1f} MB)')

    # Check for unreplaced tokens
    bad = [t for t in ['__DATA__', '__ORDER__', '__PLOTLY__', '__PARAM_PATS_BROWSER__'] if t in html]
    if bad:
        print(f'WARNING: Unreplaced tokens: {bad}')
    else:
        print('Token replacement: OK')

    print(f'\nSummary:')
    for key, d in sorted(db_data.items()):
        for s in d['sessions']:
            n_p = len(s.get('p', {}))
            n_cyls = len(s.get('cyls', {}))
            print(f'  {key}: {s["date"]} — {len(s["ts"])} pts, {n_cyls} cyls, {n_p} params')


if __name__ == '__main__':
    main()
