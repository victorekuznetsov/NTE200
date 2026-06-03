#!/usr/bin/env python3
"""
QSK50 Predictive Analytics Dashboard
Cummins QSK50 | NTE200 & KOMATSU 730E | Горная Евразия
Запуск: streamlit run qsk50_dashboard.py
"""

import os, re, io
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import scipy.stats as sci_stats
import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QSK50 Predictive Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.big-title{font-size:1.8rem;font-weight:700;color:#0d2b45;line-height:1.1}
.sub-title{font-size:.85rem;color:#555;margin-top:2px}
div[data-testid="stMetric"]{background:#f4f6f9;border-radius:8px;padding:10px 14px}
.stTabs [data-baseweb="tab"]{font-size:.8rem;padding:6px 12px}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
RU_MONTHS = {
    'янв':'01','фев':'02','мар':'03','апр':'04','май':'05','июн':'06',
    'июл':'07','авг':'08','сен':'09','окт':'10','ноя':'11','дек':'12',
    'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
    'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12',
}

THR = {
    'coolant': {'ok':(82,95),'warn':(95,105),'crit':105},
    'oil_p':   {'ok':(350,700),'warn':(300,350),'crit':300},
    'oil_t':   {'ok':(85,105),'warn':(105,115),'crit':115},
    'exhaust': {'ok':(380,520),'warn':(520,560),'crit':560},
    'ex_del':  {'ok':(0,30),  'warn':(30,60),  'crit':60},
    'crank':   {'ok':(0,3),   'warn':(3,7),    'crit':7},
    'boost':   {'ok':(220,350),'warn':(180,220),'crit':180},
}

CYL_LABELS = ['1L/1','2L/3','3L/5','4L/7','5L/9','6L/11','7L/13','8L/15',
               '1R/2','2R/4','3R/6','4R/8','5R/10','6R/12','7R/14','8R/16']

COLORS = {'ok':'#28a745','warn':'#ffc107','crit':'#dc3545','info':'#1e88e5'}

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_ru_datetime(date_s, time_s):
    try:
        d = str(date_s).strip()
        t = str(time_s).strip().split('.')[0]
        parts = re.split(r'[-/.]', d)
        if len(parts) == 3:
            day, mon, yr = parts
            m = RU_MONTHS.get(mon.lower().strip(), None)
            if m:
                return pd.to_datetime(f"{yr}-{m}-{day.zfill(2)} {t}", errors='coerce')
        return pd.to_datetime(f"{d} {t}", dayfirst=True, errors='coerce')
    except:
        return pd.NaT


def find_col(cols, *keywords):
    """Find first column matching any keyword (case-insensitive)."""
    for kw in keywords:
        for c in cols:
            if kw.lower() in c.lower():
                return c
    return None


# Bilingual (English INSITE / Russian DML) column aliases
COL_ALIASES = {
    'rpm':      ['Engine Speed (RPM)- Identifier0',
                 'Частота вращения двигателя (об/мин)- Идентификатор0'],
    'tcool':    ['Engine Coolant Temperature (',
                 'Температура охлаждающей жидкости двигателя (°C)'],
    'toil':     ['Engine Oil Temperature (',
                 'Температура масла (°C)- Идентификатор0'],
    'poil':     ['Engine Oil Pressure (kPa)',
                 'Давление масла (кПа)- Идентификатор0'],
    'texh_avg': ['Average Exhaust Temperature (Calculated)',
                 'Средняя температура отработавших газов (расчетное значение)'],
    'texh_L':   ['Average Exhaust Temperature - Left Bank',
                 'Средняя температура отработавших газов - левый ряд'],
    'texh_R':   ['Average Exhaust Temperature - Right Bank',
                 'Средняя температура отработавших газов - правый ряд'],
    'fuel':     ['Instantaneous Fuel Rate',
                 'Мгновенный расход топлива (л/час)- Идентификатор0'],
    'crank':    ['Crankcase Pressure (kPa)',
                 'Давление картерных газов (кПа)'],
    'boost':    ['Intake Manifold Pressure (kPa)- Identifier0',
                 'Давление во впускном коллекторе (кПа)- Идентификатор0'],
    'frail_m':  ['Fuel Rail Pressure Measured (bar)- Identifier0',
                 'Измеренное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
    'frail_c':  ['Fuel Rail Pressure Commanded (bar)- Identifier0',
                 'Заданное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
    'load':     ['Percent Load',
                 'Относительная нагрузка (Проценты)'],
    'batt':     ['Battery Voltage (V)- Identifier0',
                 'Напряжение аккумуляторной батареи (В)- Идентификатор0'],
    'tamb':     ['Intake Manifold Air Temperature',
                 'Температура воздуха во впускном коллекторе (°C)'],
    'amber':    ['Amber Warning Lamp', 'Желтая сигнальная лампа',
                 'Желтый сигнал'],
    'red':      ['Red Stop Lamp', 'Красная сигнальная лампа',
                 'Красный сигнал'],
    'water':    ['Water In Fuel State', 'Наличие воды в топливе'],
}


def get_col(cols, key):
    """Resolve a logical parameter to an actual column name (EN or RU)."""
    return find_col(cols, *COL_ALIASES.get(key, []))


def cyl_exhaust_cols(cols):
    """Per-cylinder exhaust temperature columns (EN + RU), sorted by number."""
    out = [c for c in cols
           if ('Exhaust Temperature Sensor Cylinder' in c
               or 'температуры отработавших газов цилиндра' in c.lower())
           and 'voltage' not in c.lower()
           and 'напряжение' not in c.lower()]

    def cyl_no(name):
        m = re.search(r'(?:Cylinder|цилиндра)\s*(\d+)', name, re.IGNORECASE)
        return int(m.group(1)) if m else 0

    return sorted(out, key=cyl_no)


def intake_temp_cols(cols):
    """Intake manifold temperature sensor columns (EN + RU)."""
    return [c for c in cols
            if ('Intake Manifold Temperature Sensor' in c
                or 'температуры впускного коллектора' in c.lower())
            and 'voltage' not in c.lower()
            and 'напряжение' not in c.lower()]


def cyl_label(name, prefix='Цил.'):
    """Short cylinder label from EN or RU column name."""
    m = re.search(r'(?:Cylinder|цилиндра)\s*(\d+)', name, re.IGNORECASE)
    return f'{prefix}{m.group(1)}' if m else name[:10]


def score_param(val, ok_lo, ok_hi, warn_hi, higher_is_bad=True):
    """Return 0-100 sub-score for a parameter."""
    if pd.isna(val):
        return 70  # unknown → moderate penalty
    if higher_is_bad:
        if val <= ok_hi:  return 100
        if val <= warn_hi: return 60
        return 20
    else:
        if val >= ok_lo:  return 100
        if val >= warn_hi: return 60
        return 20


def status_icon(val, ok_max, warn_max, higher_bad=True):
    if pd.isna(val): return "⚪"
    if higher_bad:
        if val <= ok_max:   return "🟢"
        if val <= warn_max: return "🟡"
        return "🔴"
    else:
        if val >= ok_max:   return "🟢"
        if val >= warn_max: return "🟡"
        return "🔴"


def ewma(series, span=20):
    return series.ewm(span=span, adjust=False).mean()


def anomaly_zscore(series, window=60, thr=3.0):
    rm = series.rolling(window, min_periods=5).mean()
    rs = series.rolling(window, min_periods=5).std()
    z  = (series - rm) / (rs + 1e-9)
    return z.abs() > thr


def linear_forecast(series, steps=120):
    v = series.dropna()
    if len(v) < 10:
        return None, None, None, None
    x = np.arange(len(v))
    slope, intercept, r, p, se = sci_stats.linregress(x, v.values)
    fx = np.arange(len(v), len(v)+steps)
    fc = slope*fx + intercept
    n  = len(v)
    t  = sci_stats.t.ppf(0.975, n-2)
    se2= se * np.sqrt(1/n + (fx - x.mean())**2 / ((x - x.mean())**2).sum())
    return fc, fc - t*se2, fc + t*se2, slope


# ── Data discovery & archive extraction ───────────────────────────────────────

EXTRACT_DIR = '.qsk50_extracted'


@st.cache_data(show_spinner="Распаковка архивов 7z…")
def extract_archives(folder):
    """Extract any *.7z archives in folder into a cache subdir (once)."""
    folder = Path(folder)
    cache  = folder / EXTRACT_DIR
    archives = list(folder.glob('*.7z'))
    if not archives:
        return str(cache) if cache.exists() else None
    try:
        import py7zr
    except ImportError:
        return None
    cache.mkdir(exist_ok=True)
    for arc in archives:
        marker = cache / (arc.stem + '.done')
        if marker.exists():
            continue
        try:
            with py7zr.SevenZipFile(arc, mode='r') as z:
                z.extractall(path=cache)
            marker.write_text('ok')
        except Exception:
            pass
    return str(cache)


def find_data_files(folder, pattern):
    """Recursively find files matching pattern in folder + extraction cache.
    Deduplicated by filename."""
    roots = [Path(folder)]
    cache = Path(folder) / EXTRACT_DIR
    if cache.exists():
        roots.append(cache)
    # legacy dev location
    legacy = Path('/tmp/nte_data')
    if legacy.exists():
        roots.append(legacy)
    seen = {}
    for root in roots:
        if not root.exists():
            continue
        for fp in root.rglob(pattern):
            if fp.is_file() and fp.name not in seen:
                seen[fp.name] = fp
    return list(seen.values())


# ── Data loaders ──────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_insite_csvs(folder):
    """Return dict {label: DataFrame} for all INSITE CSV files found."""
    results = {}
    csv_files = find_data_files(folder, '*.csv')

    for fpath in csv_files:
        fname = fpath.name
        if 'production' in fname.lower() or 'speed' in fname.lower():
            continue

        # truck number — handle both №43 style and DML "... 82 ..." style
        m = re.search(r'[№#N]\s*(\d+)', fname)
        if not m:
            m = re.search(r'DML-\d+-\d+\s+(\d+)', fname)
        truck = int(m.group(1)) if m else None
        truck_type = '730E' if '730' in fname else 'NTE200'

        for enc in ('cp1251', 'latin-1', 'utf-8'):
            try:
                with open(fpath, 'r', encoding=enc, errors='replace') as f:
                    lines = f.readlines()

                # header may be English ("Date") or Russian ("Дата")
                hdr_idx = next(
                    (i for i, l in enumerate(lines)
                     if l.startswith('"Date"') or l.startswith('Date,')
                     or l.startswith('"Дата"') or l.startswith('Дата,')), None
                )
                if hdr_idx is None:
                    continue

                # Build DataFrame from non-blank lines starting at the header.
                # INSITE files use \r\r\n endings producing blank lines, so we
                # parse the cleaned text directly to keep header/data aligned.
                data_lines = [l for l in lines[hdr_idx:] if l.strip()]
                if len(data_lines) < 6:
                    continue
                df = pd.read_csv(io.StringIO(''.join(data_lines)),
                                 low_memory=False)
                df = df.dropna(how='all')
                if len(df) < 5:
                    continue

                # normalise Russian date/time headers to Date/Time
                df = df.rename(columns={'Дата': 'Date', 'Время': 'Time'})

                # datetime
                if 'Date' in df.columns and 'Time' in df.columns:
                    df['_dt'] = df.apply(
                        lambda r: parse_ru_datetime(r['Date'], r['Time']), axis=1)
                    df.sort_values('_dt', inplace=True)

                # numeric conversion (comma decimal) — pandas 3.x reads text as
                # StringDtype, so convert any non-numeric column.
                for col in df.columns:
                    if col in ('Date', 'Time', '_dt'):
                        continue
                    if pd.api.types.is_numeric_dtype(df[col]):
                        continue
                    try:
                        converted = (df[col].astype(str)
                                     .str.replace(',', '.', regex=False)
                                     .str.strip())
                        num = pd.to_numeric(converted, errors='coerce')
                        if num.notna().sum() > len(df) * 0.3:
                            df[col] = num
                    except:
                        pass

                df['_truck'] = truck
                df['_type']  = truck_type
                df['_file']  = fname
                label = f"{truck_type} №{truck} — {fname[:30]}"
                results[label] = df
                break
            except:
                continue

    return results


@st.cache_data(show_spinner=False)
def load_ecm_faults(folder):
    records = []
    for fp in sorted(find_data_files(folder, '*ECM0*.xlsx')):
            m = re.match(r'(\d+)_ECM', fp.name)
            if not m: continue
            truck = int(m.group(1))
            dm = re.search(r'ECM0_(\S+?)\.xlsx', fp.name)
            date_s = dm.group(1) if dm else ''
            try:
                xl = pd.ExcelFile(fp)
                if 'Fault Codes' in xl.sheet_names:
                    df = pd.read_excel(fp, sheet_name='Fault Codes', header=6)
                    df.columns = (['Sr_No','Fault_Code','Status','Counts','Lamp',
                                   'Description','First','Last','_x',
                                   'Since_First','Since_Last'] + ['_']*10)[:len(df.columns)]
                    df = df[df['Fault_Code'].astype(str).str.match(r'^\d+$', na=False)]
                    df['Truck']      = truck
                    df['Date']       = date_s
                    df['Fault_Code'] = pd.to_numeric(df['Fault_Code'], errors='coerce')
                    df['Counts']     = pd.to_numeric(df['Counts'], errors='coerce')
                    records.append(df[['Truck','Date','Fault_Code','Status',
                                       'Counts','Lamp','Description','First','Last']])
                if 'Engine Protection' in xl.sheet_names:
                    ep = pd.read_excel(fp, sheet_name='Engine Protection', header=6)
                    ep.columns = (['Description','Fault_Code','Value','Unit',
                                   'Duration','ECM_Time','ECM_RT','Since'] + ['_']*10)[:len(ep.columns)]
                    ep = ep[ep['Fault_Code'].astype(str).str.match(r'^\d+$', na=False)]
                    ep['Truck'] = truck; ep['Date'] = date_s
                    records.append(ep[['Truck','Date','Fault_Code','Description',
                                       'Value','Unit','Duration']].assign(
                        Status='Protection', Counts=1, Lamp='Red', First='', Last=''))
            except: pass
    return pd.concat(records, ignore_index=True) if records else pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_ge_events(folder):
    records = []
    for fp in sorted(find_data_files(folder, 'GE_*.xlsx')):
            m = re.match(r'GE_(\d+)_', fp.name)
            if not m: continue
            truck = int(m.group(1))
            try:
                xl = pd.ExcelFile(fp)
                sn = xl.sheet_names[0]
                # header row 7 (0-indexed) → row 8 in 1-indexed
                df = pd.read_excel(fp, sheet_name=sn, header=7)
                # rename columns
                col_map = {df.columns[i]: n for i, n in enumerate(
                    ['_idx','Type','Event_No','Sub_ID','Name','Sub_Name',
                     'Time','Restriction','TruckSpd','State'])
                    if i < len(df.columns)}
                df = df.rename(columns=col_map)
                df = df[df['Type'].notna() &
                        (df['Type'].astype(str).str.strip() != 'nan') &
                        (df['Type'].astype(str).str.strip() != 'Type')]
                df['Truck']    = truck
                df['Time']     = pd.to_datetime(df['Time'], errors='coerce')
                df['Event_No'] = pd.to_numeric(df['Event_No'], errors='coerce')
                df['TruckSpd'] = pd.to_numeric(df['TruckSpd'], errors='coerce')
                records.append(df)
            except: pass
    return pd.concat(records, ignore_index=True) if records else pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_payload(folder):
    records = []
    for fp in sorted(find_data_files(folder, 'Весовая_*.xlsx')):
            m = re.search(r'_(\d+)_', fp.name)
            if not m: continue
            truck = int(m.group(1))
            try:
                raw = pd.read_excel(fp, header=0)
                cols = [str(c) for c in raw.iloc[0]]
                df   = raw.iloc[1:].copy()
                df.columns = cols
                df['Truck'] = truck
                num_cols = ['FinalPayload','StaticPayload','LoadPercentage',
                            'BucketCount','DistanceEmpty','DistanceLoaded',
                            'LF_TKPH','RF_TKPH','LR_TKPH','RR_TKPH',
                            'LdMaxSpeed','EmMaxSpeed','FuelLevel','FrameTorque',
                            'MaxPlusTorque','MaxMinusTorque',
                            'LeftFrontLoadWeight','RightFrontLoadWeight',
                            'LeftRearLoadWeight','RightRearLoadWeight',
                            'StopLoadedTime','MovingLoadedTime',
                            'MovingEmptyTime','StopEmptyTime',
                            'LoadingTime','DumpingTime','Slope']
                for c in num_cols:
                    if c in df.columns:
                        df[c] = pd.to_numeric(df[c], errors='coerce')
                if all(c in df.columns for c in ['ear','Month','Day','Hour','Minute']):
                    df['Year'] = pd.to_numeric(df['ear'], errors='coerce') + 2000
                    for c in ['Month','Day','Hour','Minute']:
                        df[c] = pd.to_numeric(df[c], errors='coerce')
                    df['_dt'] = pd.to_datetime(
                        dict(year=df['Year'],month=df['Month'],
                             day=df['Day'],hour=df['Hour'],minute=df['Minute']),
                        errors='coerce')
                records.append(df)
            except: pass
    return pd.concat(records, ignore_index=True) if records else pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_maintenance(folder):
    matches = find_data_files(folder, 'ОТЧЕТ Полюс Магадан*.xlsx')
    if not matches:
        return pd.DataFrame()
    fp = matches[0]
    try:
        df = pd.read_excel(fp)
        df.columns = (['Механик','Дата','Смена','Оборудование','Гараж_№',
                       'Время_останова','Время_запуска','Простой_td',
                       'Простой_full','Работа','Мото_часы',
                       'Примечание','План_запуска'] + ['_']*5)[:len(df.columns)]
        df['Дата']     = pd.to_datetime(df['Дата'], errors='coerce')
        df['Гараж_№']  = pd.to_numeric(df['Гараж_№'], errors='coerce')
        df['Мото_часы']= pd.to_numeric(df['Мото_часы'], errors='coerce')

        def td_hours(v):
            try:
                s = str(v)
                if 'days' in s:
                    p = s.split(); d=float(p[0]); h=p[2].split(':')
                    return d*24 + float(h[0]) + float(h[1])/60
                if ':' in s:
                    h = s.split(':')
                    return float(h[0]) + float(h[1])/60
                return float(s)
            except: return np.nan

        df['Простой_ч'] = df['Простой_td'].apply(td_hours)
        df['Код_ошибки']= df['Работа'].astype(str).str.extract(r'Ошибк\w*\s+(\d+)')

        def cat(w):
            w = str(w).lower()
            if re.search(r'то-?\d', w): return 'ТО'
            if 'гбц' in w or 'головк' in w: return 'Замена ГБЦ'
            if 'клапан' in w or 'вып.' in w: return 'Клапаны'
            if 'форсунк' in w: return 'Форсунки'
            if 'долив' in w: return 'Доливка масла'
            if 'диагност' in w: return 'Диагностика'
            if 'ошибк' in w: return 'Ошибка ЭБУ'
            if 'замена' in w: return 'Замена детали'
            if 'ремонт' in w: return 'Ремонт'
            if 'течь' in w: return 'Утечка'
            if 'проводк' in w or 'эл.' in w: return 'Электрика'
            if 'забор' in w or 'проб' in w: return 'Отбор проб'
            if 'осмотр' in w or 'провер' in w: return 'Осмотр'
            return 'Прочее'

        df['Категория'] = df['Работа'].apply(cat)
        return df.dropna(subset=['Дата'])
    except Exception as e:
        st.error(f"Ошибка загрузки ТО: {e}")
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_gbc(folder):
    matches = find_data_files(folder, 'ГБЦ ремонты*.xlsx')
    if not matches:
        return pd.DataFrame()
    fp = matches[0]
    try:
        df = pd.read_excel(fp, header=0)
        n_cyl = len(df.columns) - 5
        df.columns = (['Самосвал','Дата','МЧ','ДВС'] +
                      [f'Цил_{i}' for i in range(n_cyl)] +
                      ['Примечание'])[:len(df.columns)]
        df['Дата']    = pd.to_datetime(df['Дата'], errors='coerce')
        df['МЧ']      = pd.to_numeric(df['МЧ'], errors='coerce')
        df['Самосвал']= pd.to_numeric(df['Самосвал'], errors='coerce')
        return df.dropna(subset=['Самосвал'])
    except:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_oil(folder):
    matches = find_data_files(folder, 'Доливки*.xlsx')
    if not matches:
        return pd.DataFrame()
    fp = matches[0]
    try:
        df = pd.read_excel(fp, header=5)
        df.columns = (['Дата','Оборудование','Гараж_№',
                       'Вид_ГСМ','Объём_л','Мото_часы'] + ['_']*5)[:len(df.columns)]
        df['Дата']    = pd.to_datetime(df['Дата'], errors='coerce')
        df['Гараж_№'] = pd.to_numeric(df['Гараж_№'], errors='coerce')
        df['Объём_л'] = pd.to_numeric(df['Объём_л'], errors='coerce')
        df['Мото_часы']= pd.to_numeric(df['Мото_часы'], errors='coerce')
        return df.dropna(subset=['Дата','Объём_л'])
    except:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_production(folder):
    matches = find_data_files(folder, 'speed and production*.csv')
    if not matches:
        return pd.DataFrame()
    fp = matches[0]
    for enc in ('utf-8','cp1251','latin-1'):
        try:
            return pd.read_csv(fp, encoding=enc, encoding_errors='replace')
        except: pass
    return pd.DataFrame()


# ── Tab: Fleet Overview ───────────────────────────────────────────────────────

def tab_fleet(folder, maint, sel_trucks, faults, oil_df):
    st.subheader("📊 Обзор состояния парка")

    if maint.empty:
        st.warning("Нет данных ТО (ОТЧЕТ Полюс Магадан.xlsx)")
        return

    m = maint[maint['Гараж_№'].isin(sel_trucks)] if sel_trucks else maint
    f = faults[faults['Truck'].isin(sel_trucks)] if (not faults.empty and sel_trucks) else faults
    o = oil_df[oil_df['Гараж_№'].isin(sel_trucks)] if (not oil_df.empty and sel_trucks) else oil_df

    # KPI row
    c1,c2,c3,c4,c5 = st.columns(5)
    n_trucks = int(m['Гараж_№'].nunique())
    total_down = m['Простой_ч'].sum()
    avg_down   = m['Простой_ч'].mean()
    active_faults = 0 if faults.empty else int((faults['Status']=='Active').sum())
    oil_total  = 0 if oil_df.empty else int(o['Объём_л'].sum())
    c1.metric("🚛 Самосвалов", n_trucks)
    c2.metric("⏱️ Суммарный простой, ч", f"{total_down:.0f}")
    c3.metric("⏱️ Ср. простой на событие, ч", f"{avg_down:.1f}")
    c4.metric("🔴 Активных кодов ЭБУ", active_faults)
    c5.metric("🛢️ Доливки масла, л", oil_total)

    st.markdown("---")
    col1, col2 = st.columns(2)

    # 1. Downtime by truck (bar)
    with col1:
        down_by_truck = (m.groupby('Гараж_№')['Простой_ч'].sum()
                          .reset_index().sort_values('Простой_ч', ascending=False))
        fig = px.bar(down_by_truck, x='Гараж_№', y='Простой_ч',
                     color='Простой_ч', color_continuous_scale='RdYlGn_r',
                     title='Суммарный простой по самосвалам, ч',
                     labels={'Гараж_№':'Самосвал','Простой_ч':'Простой, ч'})
        fig.update_layout(height=320, margin=dict(t=40,b=20))
        st.plotly_chart(fig, use_container_width=True)

    # 2. Pareto of work categories
    with col2:
        cat_cnt = m['Категория'].value_counts().reset_index()
        cat_cnt.columns = ['Категория','Кол-во']
        cat_cnt['Накопл.%'] = (cat_cnt['Кол-во'].cumsum() /
                                cat_cnt['Кол-во'].sum() * 100)
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=cat_cnt['Категория'], y=cat_cnt['Кол-во'],
                    name='Кол-во событий', marker_color='#1e88e5',
                    secondary_y=False)
        fig.add_scatter(x=cat_cnt['Категория'], y=cat_cnt['Накопл.%'],
                        name='Накопл.%', mode='lines+markers',
                        marker_color='#e53935', secondary_y=True)
        fig.update_layout(title='Парето причин ТО/ремонтов', height=320,
                          margin=dict(t=40,b=20))
        fig.update_yaxes(title_text='Кол-во', secondary_y=False)
        fig.update_yaxes(title_text='%', secondary_y=True, range=[0,105])
        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    # 3. Heatmap: truck × week downtime
    with col3:
        m2 = m.copy()
        m2['Неделя'] = m2['Дата'].dt.to_period('W').astype(str)
        heat = (m2.groupby(['Гараж_№','Неделя'])['Простой_ч'].sum()
                  .reset_index())
        if not heat.empty:
            pv = heat.pivot(index='Гараж_№', columns='Неделя', values='Простой_ч').fillna(0)
            fig = px.imshow(pv, color_continuous_scale='RdYlGn_r',
                            title='Тепловая карта простоев (самосвал × неделя)',
                            labels={'color':'Простой, ч'})
            fig.update_layout(height=350, margin=dict(t=40,b=20))
            st.plotly_chart(fig, use_container_width=True)

    # 4. Pie of categories
    with col4:
        fig = px.pie(cat_cnt.head(10), names='Категория', values='Кол-во',
                     title='Структура работ ТО/ремонт', hole=0.4,
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(height=350, margin=dict(t=40,b=20))
        st.plotly_chart(fig, use_container_width=True)

    col5, col6 = st.columns(2)

    # 5. Timeline of cumulative downtime
    with col5:
        ts = (m.groupby('Дата')['Простой_ч'].sum()
                .cumsum().reset_index())
        fig = px.area(ts, x='Дата', y='Простой_ч',
                      title='Накопленный простой парка, ч',
                      labels={'Простой_ч':'Накопл. простой, ч'})
        fig.update_layout(height=300, margin=dict(t=40,b=20))
        st.plotly_chart(fig, use_container_width=True)

    # 6. MTBF per truck
    with col6:
        mtbf_data = []
        for tr, grp in m.groupby('Гараж_№'):
            events = grp.sort_values('Дата')
            gaps   = events['Дата'].diff().dt.total_seconds() / 3600
            mtbf   = gaps.mean()
            mtbf_data.append({'Самосвал': tr, 'MTBF_ч': mtbf})
        mtbf_df = pd.DataFrame(mtbf_data).dropna()
        if not mtbf_df.empty:
            mtbf_df = mtbf_df.sort_values('MTBF_ч')
            avg_mtbf = mtbf_df['MTBF_ч'].mean()
            fig = px.bar(mtbf_df, x='Самосвал', y='MTBF_ч',
                         title='MTBF по самосвалам (среднее время между событиями ТО), ч',
                         color='MTBF_ч', color_continuous_scale='RdYlGn')
            fig.add_hline(y=avg_mtbf, line_dash='dash', line_color='red',
                          annotation_text=f'Среднее: {avg_mtbf:.1f}ч')
            fig.update_layout(height=300, margin=dict(t=40,b=20))
            st.plotly_chart(fig, use_container_width=True)

    col7, col8 = st.columns(2)

    # 7. Error codes from maintenance
    with col7:
        err = m['Код_ошибки'].dropna().value_counts().reset_index()
        err.columns = ['Код','Кол-во']
        if not err.empty:
            fig = px.bar(err.head(15), x='Код', y='Кол-во',
                         title='Коды ошибок ЭБУ из журнала ТО',
                         color='Кол-во', color_continuous_scale='Reds')
            fig.update_layout(height=300, margin=dict(t=40,b=20))
            st.plotly_chart(fig, use_container_width=True)

    # 8. Active ECM faults table
    with col8:
        if not faults.empty:
            active = f[f['Status']=='Active'][['Truck','Fault_Code','Lamp','Counts','Description']]
            if not active.empty:
                st.markdown("**🔴 Активные коды неисправностей ЭБУ**")
                st.dataframe(active.sort_values('Counts', ascending=False),
                             use_container_width=True, height=280)
            else:
                st.info("Активных кодов ЭБУ не обнаружено")

    # 9. Oil consumption by truck
    if not oil_df.empty and not o.empty:
        st.markdown("---")
        st.markdown("**🛢️ Доливки моторного масла по самосвалам**")
        co1, co2 = st.columns(2)
        with co1:
            oil_truck = o.groupby('Гараж_№')['Объём_л'].sum().reset_index()
            oil_truck.columns = ['Самосвал','Доливки_л']
            oil_truck = oil_truck.sort_values('Доливки_л', ascending=False)
            fig = px.bar(oil_truck, x='Самосвал', y='Доливки_л',
                         title='Суммарный объём доливок масла, л',
                         color='Доливки_л', color_continuous_scale='Oranges')
            fig.update_layout(height=280, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)
        with co2:
            oil_date = o.groupby('Дата')['Объём_л'].sum().reset_index()
            fig = px.area(oil_date, x='Дата', y='Объём_л',
                          title='Динамика доливок масла по дням, л',
                          color_discrete_sequence=['#ff7043'])
            fig.update_layout(height=280, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)


# ── Tab: Engine Parameters ────────────────────────────────────────────────────

def tab_engine(insite_dict):
    st.subheader("🔩 Параметры двигателя (INSITE данные)")

    if not insite_dict:
        st.warning("CSV файлы INSITE не найдены в указанной папке. "
                   "Убедитесь, что папка содержит файлы *.csv с данными INSITE Professional.")
        st.info("Ожидаются файлы вида: `NTE200A N43 16.05.2026.csv` или `DML-*.csv`")
        return

    file_key = st.selectbox("Выберите файл (самосвал / сессия):", list(insite_dict.keys()))
    df = insite_dict[file_key].copy()
    cols = df.columns.tolist()

    # identify key columns (bilingual EN/RU)
    COL = {k: get_col(cols, k) for k in COL_ALIASES}
    CYL_COLS = cyl_exhaust_cols(cols)
    IMT_COLS = intake_temp_cols(cols)

    dt = df['_dt'] if '_dt' in df.columns else pd.Series(range(len(df)))

    # metadata
    truck   = df['_truck'].iloc[0] if '_truck' in df.columns else '—'
    ttype   = df['_type'].iloc[0]  if '_type'  in df.columns else '—'
    n_rows  = len(df)
    dur_min = (dt.max() - dt.min()).total_seconds()/60 if pd.api.types.is_datetime64_any_dtype(dt) else n_rows
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Самосвал", f"{ttype} №{truck}")
    c2.metric("Точек данных", n_rows)
    c3.metric("Продолжительность, мин", f"{dur_min:.1f}")
    if COL['rpm'] and df[COL['rpm']].notna().any():
        c4.metric("Обороты (avg), RPM", f"{df[COL['rpm']].mean():.0f}")
    st.markdown("---")

    # 1. RPM
    if COL['rpm']:
        fig = go.Figure()
        fig.add_scatter(x=dt, y=df[COL['rpm']], name='RPM', line=dict(color='#1565c0'))
        fig.add_scatter(x=dt, y=ewma(df[COL['rpm']], 30),
                        name='EWMA', line=dict(color='#ff6f00', dash='dash'))
        fig.update_layout(title='Обороты двигателя (RPM)', height=280,
                          margin=dict(t=40,b=10), yaxis_title='RPM')
        st.plotly_chart(fig, use_container_width=True)

    col_a, col_b = st.columns(2)

    # 2. Coolant temp
    with col_a:
        if COL['tcool']:
            fig = go.Figure()
            fig.add_scatter(x=dt, y=df[COL['tcool']], name='T охл.', line=dict(color='#0288d1'))
            fig.add_hrect(y0=82, y1=95, fillcolor='green', opacity=0.08, line_width=0)
            fig.add_hrect(y0=95, y1=105,fillcolor='orange',opacity=0.08, line_width=0)
            fig.add_hrect(y0=105,y1=130, fillcolor='red',  opacity=0.08, line_width=0)
            fig.add_hline(y=95, line_dash='dot', line_color='orange', annotation_text='95°C')
            fig.add_hline(y=105,line_dash='dot', line_color='red',    annotation_text='105°C')
            fig.update_layout(title='Температура охлаждающей жидкости, °C',
                              height=280, margin=dict(t=40,b=10), yaxis_title='°C')
            st.plotly_chart(fig, use_container_width=True)

    # 3. Oil temp + pressure (dual axis)
    with col_b:
        if COL['toil'] and COL['poil']:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_scatter(x=dt, y=df[COL['toil']], name='T масла °C',
                            line=dict(color='#e65100'), secondary_y=False)
            fig.add_scatter(x=dt, y=df[COL['poil']], name='P масла кПа',
                            line=dict(color='#6a1b9a'), secondary_y=True)
            fig.update_layout(title='Температура и давление масла', height=280,
                              margin=dict(t=40,b=10))
            fig.update_yaxes(title_text='°C',  secondary_y=False)
            fig.update_yaxes(title_text='кПа', secondary_y=True)
            st.plotly_chart(fig, use_container_width=True)

    col_c, col_d = st.columns(2)

    # 4. Exhaust avg L/R banks
    with col_c:
        fig = go.Figure()
        if COL['texh_L']:
            fig.add_scatter(x=dt, y=df[COL['texh_L']], name='Лев. банк',
                            line=dict(color='#c62828'))
        if COL['texh_R']:
            fig.add_scatter(x=dt, y=df[COL['texh_R']], name='Прав. банк',
                            line=dict(color='#283593'))
        if COL['texh_avg']:
            fig.add_scatter(x=dt, y=df[COL['texh_avg']], name='Среднее',
                            line=dict(color='#2e7d32', dash='dash'))
        fig.add_hrect(y0=520, y1=600, fillcolor='orange', opacity=0.08, line_width=0)
        fig.add_hrect(y0=560, y1=700, fillcolor='red',    opacity=0.08, line_width=0)
        fig.update_layout(title='Средняя температура выхлопа, °C',
                          height=280, margin=dict(t=40,b=10), yaxis_title='°C')
        st.plotly_chart(fig, use_container_width=True)

    # 5. Per-cylinder exhaust heatmap
    with col_d:
        if CYL_COLS:
            # sort by cylinder number
            cyl_data = df[CYL_COLS].copy()
            cyl_data.columns = [cyl_label(c, 'Цил') for c in CYL_COLS]
            cyl_data = cyl_data.apply(pd.to_numeric, errors='coerce')
            # downsample for heatmap
            step = max(1, len(cyl_data)//80)
            z    = cyl_data.iloc[::step].T.values
            x_ax = list(range(0, len(cyl_data), step))
            y_ax = cyl_data.columns.tolist()
            fig  = go.Figure(go.Heatmap(z=z, x=x_ax, y=y_ax,
                                        colorscale='RdYlGn_r', colorbar_title='°C'))
            fig.update_layout(title='Тепловая карта T выхлопа по цилиндрам',
                              height=280, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    col_e, col_f = st.columns(2)

    # 6. Cylinder deviation from mean (bar at last valid point)
    with col_e:
        if CYL_COLS:
            last_valid = df[CYL_COLS].dropna(how='all').iloc[-5:].mean()
            avg_all    = last_valid.mean()
            delta      = last_valid - avg_all
            delta.index = [cyl_label(c) for c in CYL_COLS]
            colors_bar = ['#dc3545' if abs(v)>40 else
                          '#ffc107' if abs(v)>20 else '#28a745'
                          for v in delta.values]
            fig = go.Figure(go.Bar(x=delta.index, y=delta.values,
                                   marker_color=colors_bar))
            fig.add_hline(y=0, line_color='black')
            fig.add_hline(y=40,  line_dash='dot', line_color='red')
            fig.add_hline(y=-40, line_dash='dot', line_color='red')
            fig.update_layout(title='Отклонение T цилиндров от среднего, °C',
                              height=280, margin=dict(t=40,b=10), yaxis_title='Δ°C')
            st.plotly_chart(fig, use_container_width=True)

    # 7. Fuel rate
    with col_f:
        if COL['fuel']:
            fig = go.Figure()
            fig.add_scatter(x=dt, y=df[COL['fuel']], name='Расход л/ч',
                            line=dict(color='#f57f17', width=1), opacity=0.5)
            fig.add_scatter(x=dt, y=ewma(df[COL['fuel']], 20),
                            name='EWMA', line=dict(color='#e65100', width=2))
            fig.update_layout(title='Мгновенный расход топлива, л/ч',
                              height=280, margin=dict(t=40,b=10), yaxis_title='л/ч')
            st.plotly_chart(fig, use_container_width=True)

    col_g, col_h = st.columns(2)

    # 8. Boost + crankcase
    with col_g:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        if COL['boost']:
            fig.add_scatter(x=dt, y=df[COL['boost']], name='Наддув кПа',
                            line=dict(color='#00838f'), secondary_y=False)
        if COL['crank']:
            fig.add_scatter(x=dt, y=df[COL['crank']], name='Картер кПа',
                            line=dict(color='#6a1b9a'), secondary_y=True)
            fig.add_hrect(y0=3, y1=7, fillcolor='orange', opacity=0.1,
                          secondary_y=True, line_width=0)
            fig.add_hrect(y0=7, y1=30, fillcolor='red', opacity=0.1,
                          secondary_y=True, line_width=0)
        fig.update_layout(title='Давление наддува и картерных газов',
                          height=280, margin=dict(t=40,b=10))
        fig.update_yaxes(title_text='Наддув кПа', secondary_y=False)
        fig.update_yaxes(title_text='Картер кПа', secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)

    # 9. Fuel rail pressure commanded vs measured
    with col_h:
        if COL['frail_m'] and COL['frail_c']:
            fig = go.Figure()
            fig.add_scatter(x=dt, y=df[COL['frail_c']], name='Задание (bar)',
                            line=dict(color='#1565c0', dash='dash'))
            fig.add_scatter(x=dt, y=df[COL['frail_m']], name='Измерение (bar)',
                            line=dict(color='#c62828'))
            fig.update_layout(title='Давление топлива в рампе, bar',
                              height=280, margin=dict(t=40,b=10), yaxis_title='bar')
            st.plotly_chart(fig, use_container_width=True)

    col_i, col_j = st.columns(2)

    # 10. Intake manifold temps (4 sensors)
    with col_i:
        if IMT_COLS:
            fig = go.Figure()
            palette = ['#1e88e5','#e53935','#43a047','#fb8c00']
            for i, c in enumerate(IMT_COLS[:4]):
                lbl = f'Датчик {i+1}'
                fig.add_scatter(x=dt, y=df[c], name=lbl,
                                line=dict(color=palette[i%4]))
            fig.update_layout(title='Температура воздуха на впуске (4 датчика), °C',
                              height=280, margin=dict(t=40,b=10), yaxis_title='°C')
            st.plotly_chart(fig, use_container_width=True)

    # 11. Engine load + battery
    with col_j:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        if COL['load']:
            fig.add_scatter(x=dt, y=df[COL['load']], name='Нагрузка %',
                            line=dict(color='#00695c'), secondary_y=False)
        if COL['batt']:
            fig.add_scatter(x=dt, y=df[COL['batt']], name='АКБ В',
                            line=dict(color='#f57f17'), secondary_y=True)
        fig.update_layout(title='Нагрузка двигателя и напряжение АКБ',
                          height=280, margin=dict(t=40,b=10))
        fig.update_yaxes(title_text='Нагрузка %', secondary_y=False)
        fig.update_yaxes(title_text='Напряжение В', secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)

    # 12. Statistics table
    st.markdown("---")
    st.markdown("**📋 Статистическая сводка**")
    stat_cols = {k: v for k, v in COL.items()
                 if v and df[v].dtype in [np.float64, np.int64, float, int]}
    if CYL_COLS:
        stat_cols.update({f'Цил.{i+1}': c for i, c in enumerate(CYL_COLS)})
    stat_rows = []
    for name, col in stat_cols.items():
        s = df[col].dropna()
        if len(s) > 0:
            stat_rows.append({
                'Параметр': col[:55],
                'Min': f"{s.min():.2f}",
                'Max': f"{s.max():.2f}",
                'Среднее': f"{s.mean():.2f}",
                'Std': f"{s.std():.2f}",
                'Последнее': f"{s.iloc[-1]:.2f}"
            })
    if stat_rows:
        st.dataframe(pd.DataFrame(stat_rows), use_container_width=True)

    # Export
    if st.button("💾 Экспорт данных в CSV"):
        csv_buf = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Скачать CSV", csv_buf,
                           file_name=f"engine_data_{truck}.csv",
                           mime='text/csv')


# ── Tab: Predictive Analytics ─────────────────────────────────────────────────

def tab_predictive(insite_dict):
    st.subheader("🔮 Предиктивный анализ")

    if not insite_dict:
        st.warning("Нет данных INSITE CSV для анализа")
        return

    file_key = st.selectbox("Файл для анализа:", list(insite_dict.keys()),
                            key='pred_sel')
    df  = insite_dict[file_key].copy()
    dt  = df['_dt'] if '_dt' in df.columns else pd.Series(range(len(df)))
    cols= df.columns.tolist()
    truck = df['_truck'].iloc[0] if '_truck' in df.columns else '—'

    tcool = get_col(cols, 'tcool')
    toil  = get_col(cols, 'toil')
    poil  = get_col(cols, 'poil')
    texh  = get_col(cols, 'texh_avg')
    fuel  = get_col(cols, 'fuel')
    crank = get_col(cols, 'crank')
    load  = get_col(cols, 'load')

    CYL_COLS = cyl_exhaust_cols(cols)

    # ─ Health Score
    sub_scores = {}
    if tcool and df[tcool].notna().any():
        v = df[tcool].mean()
        sub_scores['Охлаждение'] = score_param(v, 82, 95, 105, True)
    if poil and df[poil].notna().any():
        v = df[poil].mean()
        sub_scores['Давл. масла'] = score_param(v, 350, 350, 300, False)
    if crank and df[crank].notna().any():
        v = df[crank].mean()
        sub_scores['Картер'] = score_param(v, 0, 3, 7, True)
    if CYL_COLS:
        cyl_std = df[CYL_COLS].std(axis=1).mean()
        sub_scores['Баланс цил.'] = score_param(cyl_std, 0, 30, 60, True)
    if fuel and load and df[fuel].notna().any():
        # fuel per percent load
        fl = (df[fuel] / (df[load].replace(0, np.nan))).mean()
        sub_scores['Расход/нагр.'] = max(20, 100 - (fl / 4) * 5) if not np.isnan(fl) else 70

    health = np.mean(list(sub_scores.values())) if sub_scores else 70

    col1, col2 = st.columns([1, 2])
    with col1:
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=health,
            delta={'reference': 75, 'increasing': {'color': 'green'},
                   'decreasing': {'color': 'red'}},
            gauge={'axis': {'range': [0, 100]},
                   'bar': {'color': '#1e88e5'},
                   'steps': [
                       {'range': [0,  50], 'color': '#ffebee'},
                       {'range': [50, 75], 'color': '#fff8e1'},
                       {'range': [75,100], 'color': '#e8f5e9'}],
                   'threshold': {'line': {'color':'red','width':3},
                                 'thickness': 0.8, 'value': 50}},
            title={'text': f"Индекс здоровья ДВС №{truck}"}
        ))
        fig_gauge.update_layout(height=280, margin=dict(t=40,b=10,l=20,r=20))
        st.plotly_chart(fig_gauge, use_container_width=True)

    with col2:
        if sub_scores:
            bar_df = pd.DataFrame({'Подсистема': list(sub_scores.keys()),
                                   'Балл': list(sub_scores.values())})
            clr = ['#dc3545' if v<50 else '#ffc107' if v<75 else '#28a745'
                   for v in bar_df['Балл']]
            fig = go.Figure(go.Bar(x=bar_df['Балл'], y=bar_df['Подсистема'],
                                   orientation='h', marker_color=clr,
                                   text=bar_df['Балл'].round(0), textposition='outside'))
            fig.update_layout(title='Оценки по подсистемам (0-100)',
                              height=280, margin=dict(t=40,b=10), xaxis_range=[0,110])
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ─ Anomaly detection on coolant temp
    if tcool:
        st.markdown("**🚨 Обнаружение аномалий — температура охлаждающей жидкости**")
        s = df[tcool].dropna()
        if len(s) > 20:
            anom = anomaly_zscore(s, window=min(60, len(s)//3))
            dt_s = dt.loc[s.index]
            fig  = go.Figure()
            fig.add_scatter(x=dt_s, y=s, name='T охл.', line=dict(color='#0288d1', width=1))
            fig.add_scatter(x=dt_s, y=ewma(s,20), name='EWMA',
                            line=dict(color='#01579b', dash='dash'))
            anom_idx = s[anom]
            fig.add_scatter(x=dt_s.loc[anom_idx.index], y=anom_idx,
                            mode='markers', name='Аномалия',
                            marker=dict(color='red', size=8, symbol='x'))
            fig.update_layout(height=280, margin=dict(t=20,b=10), yaxis_title='°C')
            st.plotly_chart(fig, use_container_width=True)
            n_anom = anom.sum()
            if n_anom > 0:
                st.warning(f"⚠️ Обнаружено {n_anom} аномальных точек (Z > 3σ)")

    col_p1, col_p2 = st.columns(2)

    # ─ Oil pressure trend + forecast
    with col_p1:
        if poil:
            s = df[poil].dropna()
            if len(s) >= 10:
                fc, ci_lo, ci_hi, slope = linear_forecast(s, steps=min(120, len(s)//2))
                x_hist = list(range(len(s)))
                x_fc   = list(range(len(s), len(s)+len(fc)))
                fig = go.Figure()
                fig.add_scatter(x=x_hist, y=s.values, name='P масла кПа',
                                line=dict(color='#6a1b9a'))
                if fc is not None:
                    fig.add_scatter(x=x_fc, y=fc, name='Прогноз',
                                    line=dict(color='#e65100', dash='dash'))
                    fig.add_scatter(x=x_fc+x_fc[::-1],
                                    y=list(ci_hi)+list(ci_lo[::-1]),
                                    fill='toself', fillcolor='rgba(230,81,0,0.1)',
                                    line=dict(width=0), name='Доверит. интервал')
                    trend_dir = "↑" if slope > 0 else "↓"
                    fig.add_hline(y=THR['oil_p']['crit'], line_dash='dot',
                                  line_color='red', annotation_text='Критично 300кПа')
                    fig.update_layout(
                        title=f'Тренд давления масла {trend_dir} '
                              f'({slope*60:.2f} кПа/мин)',
                        height=300, margin=dict(t=40,b=10), yaxis_title='кПа')
                st.plotly_chart(fig, use_container_width=True)

    # ─ Crankcase trend + forecast
    with col_p2:
        if crank:
            s = df[crank].dropna()
            if len(s) >= 10:
                fc, ci_lo, ci_hi, slope = linear_forecast(s, steps=min(120, len(s)//2))
                x_hist = list(range(len(s)))
                x_fc   = list(range(len(s), len(s)+len(fc)))
                fig = go.Figure()
                fig.add_scatter(x=x_hist, y=s.values, name='P картера кПа',
                                line=dict(color='#6a1b9a'))
                fig.add_hrect(y0=3, y1=7,  fillcolor='orange', opacity=0.1, line_width=0)
                fig.add_hrect(y0=7, y1=30, fillcolor='red',    opacity=0.1, line_width=0)
                if fc is not None:
                    fig.add_scatter(x=x_fc, y=fc, name='Прогноз',
                                    line=dict(color='#e65100', dash='dash'))
                    fig.add_scatter(x=x_fc+x_fc[::-1],
                                    y=list(ci_hi)+list(ci_lo[::-1]),
                                    fill='toself', fillcolor='rgba(230,81,0,0.1)',
                                    line=dict(width=0), name='Доверит. интервал')
                fig.update_layout(title='Давление картерных газов + прогноз',
                                  height=300, margin=dict(t=40,b=10), yaxis_title='кПа')
                st.plotly_chart(fig, use_container_width=True)

    # ─ Cylinder imbalance radar
    if CYL_COLS and len(CYL_COLS) >= 8:
        st.markdown("**🔍 Радарная диаграмма температур цилиндров**")
        last_temps = df[CYL_COLS].dropna(how='all').iloc[-5:].mean()
        cyl_names  = [cyl_label(c) for c in CYL_COLS]
        avg_t = last_temps.mean()
        fig = go.Figure()
        fig.add_scatterpolar(r=last_temps.values, theta=cyl_names,
                             fill='toself', name='T цилиндров',
                             line_color='#c62828')
        fig.add_scatterpolar(r=[avg_t]*len(cyl_names), theta=cyl_names,
                             fill='toself', name='Среднее',
                             line_color='#1565c0', line_dash='dash',
                             fillcolor='rgba(21,101,192,0.05)')
        fig.update_layout(polar=dict(radialaxis=dict(
                              range=[max(0, avg_t-80), avg_t+80])),
                          title='Радарная карта T выхлопа по цилиндрам (°C)',
                          height=380, margin=dict(t=50,b=20))
        st.plotly_chart(fig, use_container_width=True)

    # ─ Fuel rate EWMA + anomaly
    if fuel:
        s = df[fuel].dropna()
        if len(s) > 20:
            ew = ewma(s, 20)
            anom = anomaly_zscore(s, window=min(60, len(s)//3))
            dt_s = dt.loc[s.index]
            fig  = go.Figure()
            fig.add_scatter(x=dt_s, y=s, name='Расход л/ч', opacity=0.4,
                            line=dict(color='#f57f17', width=1))
            fig.add_scatter(x=dt_s, y=ew, name='EWMA тренд',
                            line=dict(color='#e65100', width=2))
            if anom.any():
                fig.add_scatter(x=dt_s.loc[s[anom].index], y=s[anom],
                                mode='markers', name='Аномалия',
                                marker=dict(color='red', size=7))
            fig.update_layout(title='Расход топлива: тренд EWMA и аномалии',
                              height=280, margin=dict(t=40,b=10), yaxis_title='л/ч')
            st.plotly_chart(fig, use_container_width=True)

    # ─ Risk summary table
    st.markdown("---")
    st.markdown("**🗂️ Таблица риска по параметрам**")
    risk_rows = []
    param_map = {
        'Т охл. жидкости, °C':  (tcool,  THR['coolant']['ok'][1], THR['coolant']['crit'], True),
        'Давл. масла, кПа':     (poil,   THR['oil_p']['ok'][0],  THR['oil_p']['crit'],  False),
        'Т масла, °C':          (toil,   THR['oil_t']['ok'][1],  THR['oil_t']['crit'],  True),
        'Т выхлопа ср., °C':    (texh,   THR['exhaust']['ok'][1],THR['exhaust']['crit'],True),
        'Давл. картера, кПа':   (crank,  THR['crank']['ok'][1],  THR['crank']['crit'],  True),
    }
    for name,(col,ok,crit,hi_bad) in param_map.items():
        if col and df[col].notna().any():
            v = df[col].dropna()
            cur = v.iloc[-1]
            avg = v.mean()
            _,_,_,slope = linear_forecast(v, 30)
            trend = '↑' if (slope or 0) > 0.01 else '↓' if (slope or 0) < -0.01 else '→'
            if hi_bad:
                risk = '🔴 ВЫСОКИЙ' if cur>crit else '🟡 СРЕДНИЙ' if cur>ok else '🟢 НИЗКИЙ'
            else:
                risk = '🔴 ВЫСОКИЙ' if cur<crit else '🟡 СРЕДНИЙ' if cur<ok else '🟢 НИЗКИЙ'
            risk_rows.append({'Параметр':name, 'Текущее':f'{cur:.1f}',
                              'Среднее':f'{avg:.1f}', 'Тренд':trend, 'Риск':risk})
    if risk_rows:
        st.dataframe(pd.DataFrame(risk_rows), use_container_width=True)


# ── Tab: ECM Fault Codes ──────────────────────────────────────────────────────

def tab_faults(faults):
    st.subheader("⚠️ Коды неисправностей ЭБУ (ECM Fault Codes)")

    if faults.empty:
        st.warning("Нет данных ECM. Убедитесь что архив 'Данные с ECM.7z' распакован "
                   "или папка ECM доступна.")
        return

    # filter
    trucks = sorted(faults['Truck'].dropna().unique().astype(int))
    sel    = st.multiselect("Самосвал:", trucks, default=trucks, key='fc_trucks')
    f = faults[faults['Truck'].isin(sel)] if sel else faults

    col1, col2, col3 = st.columns(3)
    col1.metric("Уникальных кодов", int(f['Fault_Code'].nunique()))
    col2.metric("Активных", int((f['Status']=='Active').sum()))
    col3.metric("Самосвалов с неисправностями", int(f['Truck'].nunique()))

    st.markdown("---")
    c1, c2 = st.columns(2)

    # 1. Heatmap truck × fault code
    with c1:
        pv = (f.groupby(['Truck','Fault_Code'])['Counts']
               .sum().reset_index()
               .pivot(index='Truck', columns='Fault_Code', values='Counts').fillna(0))
        if not pv.empty:
            fig = px.imshow(pv, color_continuous_scale='Reds',
                            title='Тепловая карта: самосвал × код ошибки',
                            labels={'color':'Кол-во'})
            fig.update_layout(height=380, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 2. Pareto fault codes
    with c2:
        top = (f.groupby('Fault_Code')['Counts'].sum()
                .sort_values(ascending=False).head(15).reset_index())
        fig = px.bar(top, x='Counts', y='Fault_Code', orientation='h',
                     color='Counts', color_continuous_scale='Reds',
                     title='Топ-15 кодов неисправностей (по частоте)')
        fig.update_layout(height=380, margin=dict(t=40,b=10),
                          yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)

    # 3. Lamp distribution
    with c3:
        lamp = f['Lamp'].value_counts().reset_index()
        lamp.columns = ['Лампа','Кол-во']
        clr_map = {'Amber':'#ffc107','Red':'#dc3545','Maintenance':'#17a2b8'}
        fig = px.pie(lamp, names='Лампа', values='Кол-во',
                     color='Лампа', color_discrete_map=clr_map,
                     title='Распределение по типу лампы', hole=0.4)
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 4. Count of fault codes per truck
    with c4:
        cnt = f.groupby('Truck')['Fault_Code'].count().reset_index()
        cnt.columns = ['Самосвал','Кодов']
        cnt = cnt.sort_values('Кодов', ascending=False)
        fig = px.bar(cnt, x='Самосвал', y='Кодов',
                     title='Количество кодов ЭБУ по самосвалу',
                     color='Кодов', color_continuous_scale='Reds')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 5. Active faults detail
    st.markdown("---")
    st.markdown("**🔴 Активные неисправности**")
    active = f[f['Status']=='Active'][['Truck','Fault_Code','Lamp','Counts','Description','First','Last']]
    if not active.empty:
        st.dataframe(active.sort_values('Counts',ascending=False),
                     use_container_width=True, height=200)
    else:
        st.success("✅ Активных неисправностей не обнаружено")

    # 6. All faults table with filter
    st.markdown("---")
    st.markdown("**📋 Все коды неисправностей**")
    status_f = st.selectbox("Статус:", ['Все','Active','Inactive','Protection'])
    show_f = f if status_f=='Все' else f[f['Status']==status_f]
    st.dataframe(show_f[['Truck','Fault_Code','Status','Counts','Lamp','Description']],
                 use_container_width=True, height=350)


# ── Tab: Maintenance History ──────────────────────────────────────────────────

def tab_maint(maint, gbc):
    st.subheader("🔧 История технического обслуживания и ремонтов")

    if maint.empty:
        st.warning("Нет данных ТО")
        return

    c1,c2,c3 = st.columns(3)
    c1.metric("Всего событий", len(maint))
    c2.metric("Суммарный простой, ч", f"{maint['Простой_ч'].sum():.0f}")
    c3.metric("Период", f"{maint['Дата'].min().date()} — {maint['Дата'].max().date()}")
    st.markdown("---")

    co1, co2 = st.columns(2)

    # 1. Scatter timeline — downtime per event
    with co1:
        fig = px.scatter(maint.dropna(subset=['Дата','Простой_ч']),
                         x='Дата', y='Гараж_№', color='Категория',
                         size='Простой_ч', size_max=20,
                         hover_data=['Работа','Простой_ч','Мото_часы'],
                         title='Временная шкала событий ТО')
        fig.update_layout(height=380, margin=dict(t=40,b=10),
                          yaxis_title='Самосвал №')
        st.plotly_chart(fig, use_container_width=True)

    # 2. Downtime histogram
    with co2:
        dn = maint['Простой_ч'].dropna()
        fig = px.histogram(dn, nbins=40, title='Распределение продолжительности простоев',
                           labels={'value':'Простой, ч','count':'Кол-во'},
                           color_discrete_sequence=['#1e88e5'])
        fig.update_layout(height=380, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    co3, co4 = st.columns(2)

    # 3. Downtime by truck bar
    with co3:
        dbt = (maint.groupby('Гараж_№')['Простой_ч'].sum()
                    .reset_index().sort_values('Простой_ч', ascending=False))
        fig = px.bar(dbt, x='Гараж_№', y='Простой_ч',
                     title='Суммарный простой по самосвалам, ч',
                     color='Простой_ч', color_continuous_scale='RdYlGn_r')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 4. Stacked bar — category per month
    with co4:
        m2 = maint.copy()
        m2['Месяц'] = m2['Дата'].dt.to_period('M').astype(str)
        mc = m2.groupby(['Месяц','Категория']).size().reset_index(name='Кол-во')
        fig = px.bar(mc, x='Месяц', y='Кол-во', color='Категория',
                     barmode='stack', title='Виды работ по месяцам')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    co5, co6 = st.columns(2)

    # 5. Error codes frequency
    with co5:
        err = maint['Код_ошибки'].dropna().value_counts().reset_index()
        err.columns = ['Код','Кол-во']
        if not err.empty:
            fig = px.bar(err.head(15), x='Код', y='Кол-во',
                         title='Частота кодов ошибок ЭБУ из журнала ТО',
                         color='Кол-во', color_continuous_scale='Reds')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 6. MTBF
    with co6:
        mtbf_data = []
        for tr, grp in maint.groupby('Гараж_№'):
            g = grp.sort_values('Дата')
            gaps = g['Дата'].diff().dt.total_seconds()/3600
            mtbf_data.append({'Самосвал':tr,'MTBF_ч':gaps.mean()})
        mtbf_df = pd.DataFrame(mtbf_data).dropna()
        if not mtbf_df.empty:
            avg = mtbf_df['MTBF_ч'].mean()
            fig = px.bar(mtbf_df.sort_values('MTBF_ч'),
                         x='Самосвал', y='MTBF_ч', title='MTBF по самосвалам, ч',
                         color='MTBF_ч', color_continuous_scale='RdYlGn')
            fig.add_hline(y=avg, line_dash='dash', line_color='red',
                          annotation_text=f'Среднее {avg:.1f}ч')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 7. GBC cylinder heatmap
    if not gbc.empty:
        st.markdown("---")
        st.markdown("**🔩 Тепловая карта ремонтов ГБЦ по цилиндрам**")
        cyl_cols = [c for c in gbc.columns if c.startswith('Цил_')]
        if cyl_cols:
            # build repair count matrix
            heat_data = []
            for _, row in gbc.dropna(subset=['Самосвал']).iterrows():
                truck = row['Самосвал']
                for i, col in enumerate(cyl_cols[:16]):
                    val = str(row.get(col,''))
                    if val not in ('nan','None',''):
                        heat_data.append({'Самосвал':truck, 'Цилиндр':CYL_LABELS[i] if i<16 else col,
                                          'Ремонт':1})
            if heat_data:
                hdf = pd.DataFrame(heat_data)
                pv = hdf.pivot_table(index='Самосвал', columns='Цилиндр',
                                     values='Ремонт', aggfunc='sum').fillna(0)
                fig = px.imshow(pv, color_continuous_scale='Reds',
                                title='Ремонты ГБЦ: самосвал × цилиндр',
                                labels={'color':'Кол-во ремонтов'})
                fig.update_layout(height=max(300, len(pv)*18+80),
                                  margin=dict(t=40,b=20))
                st.plotly_chart(fig, use_container_width=True)

    # 8. Detailed table
    st.markdown("---")
    cat_sel = st.multiselect("Фильтр по категории:",
                              maint['Категория'].unique().tolist(),
                              default=maint['Категория'].unique().tolist(),
                              key='maint_cat')
    show = maint[maint['Категория'].isin(cat_sel)]
    cols_show = ['Дата','Гараж_№','Категория','Работа','Простой_ч','Мото_часы','Механик']
    st.dataframe(show[cols_show].sort_values('Дата', ascending=False),
                 use_container_width=True, height=350)

    if st.button("💾 Экспорт журнала ТО"):
        buf = maint.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Скачать CSV", buf, 'maintenance_log.csv', 'text/csv')


# ── Tab: Payload ──────────────────────────────────────────────────────────────

def tab_payload(payload):
    st.subheader("⚖️ Анализ нагрузки и производительности рейсов")

    if payload.empty:
        st.warning("Нет данных о нагрузке (Весовая_*.xlsx)")
        return

    trucks = sorted(payload['Truck'].dropna().unique().astype(int))
    sel    = st.multiselect("Самосвал:", trucks, default=trucks[:8], key='pay_trucks')
    p = payload[payload['Truck'].isin(sel)] if sel else payload

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Рейсов всего", len(p))
    c2.metric("Ср. нагрузка, %",   f"{p['LoadPercentage'].mean():.1f}")
    c3.metric("Ср. полезная нагр., т", f"{p['FinalPayload'].mean():.0f}")
    c4.metric("Перегрузок (>110%)", int((p['LoadPercentage']>110).sum()))
    st.markdown("---")

    c1, c2 = st.columns(2)

    # 1. Load % histogram
    with c1:
        fig = px.histogram(p, x='LoadPercentage', color='Truck',
                           nbins=30, barmode='overlay', opacity=0.7,
                           title='Распределение нагрузки, %',
                           labels={'LoadPercentage':'Нагрузка %'})
        fig.add_vline(x=100, line_dash='dash', line_color='green',
                      annotation_text='100%')
        fig.add_vline(x=110, line_dash='dot',  line_color='red',
                      annotation_text='Перегруз 110%')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 2. Final payload scatter
    with c2:
        if '_dt' in p.columns:
            fig = px.scatter(p, x='_dt', y='FinalPayload', color='Truck',
                             title='Фактическая нагрузка по рейсам, т',
                             labels={'_dt':'Дата','FinalPayload':'Нагрузка, т'})
        else:
            fig = px.histogram(p, x='FinalPayload', nbins=30,
                               title='Фактическая нагрузка, т')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)

    # 3. TKPH boxplot by axle
    with c3:
        tkph_cols = ['LF_TKPH','RF_TKPH','LR_TKPH','RR_TKPH']
        tkph_labels = ['Лев. передн.','Прав. передн.','Лев. задн.','Прав. задн.']
        avail = [c for c in tkph_cols if c in p.columns and p[c].notna().any()]
        if avail:
            melted = p[avail].melt(var_name='Ось', value_name='TKPH')
            melted['Ось'] = melted['Ось'].map(dict(zip(tkph_cols, tkph_labels)))
            fig = px.box(melted, x='Ось', y='TKPH', color='Ось',
                         title='TKPH по осям (диагностика шин)',
                         labels={'TKPH':'TKPH'})
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 4. Max speed loaded vs empty
    with c4:
        if 'LdMaxSpeed' in p.columns and 'EmMaxSpeed' in p.columns:
            sp = p[['Truck','LdMaxSpeed','EmMaxSpeed']].dropna()
            sp['LdMaxSpeed'] = sp['LdMaxSpeed'] / 1000  # convert to km/h if needed
            sp['EmMaxSpeed'] = sp['EmMaxSpeed'] / 1000
            # If values look too large, they're in something else
            if sp['LdMaxSpeed'].mean() > 200:
                sp['LdMaxSpeed'] /= 1000
                sp['EmMaxSpeed'] /= 1000
            sp_mean = sp.groupby('Truck')[['LdMaxSpeed','EmMaxSpeed']].mean().reset_index()
            fig = go.Figure()
            fig.add_bar(x=sp_mean['Truck'], y=sp_mean['LdMaxSpeed'],
                        name='Гружёный', marker_color='#e53935')
            fig.add_bar(x=sp_mean['Truck'], y=sp_mean['EmMaxSpeed'],
                        name='Порожний',  marker_color='#1e88e5')
            fig.update_layout(title='Макс. скорость гружёный/порожний (ср.)',
                              barmode='group', height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    c5, c6 = st.columns(2)

    # 5. Distance loaded vs empty
    with c5:
        if 'DistanceLoaded' in p.columns and 'DistanceEmpty' in p.columns:
            d = p[['Truck','DistanceLoaded','DistanceEmpty']].dropna()
            dm = d.groupby('Truck').mean().reset_index()
            fig = go.Figure()
            fig.add_bar(x=dm['Truck'], y=dm['DistanceLoaded'],
                        name='Гружёный, м', marker_color='#e53935')
            fig.add_bar(x=dm['Truck'], y=dm['DistanceEmpty'],
                        name='Порожний, м',  marker_color='#1e88e5')
            fig.update_layout(title='Ср. дистанция рейса, м',
                              barmode='group', height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 6. Overload timeline
    with c6:
        overload = p[p['LoadPercentage']>110].copy()
        if not overload.empty and '_dt' in overload.columns:
            ol_by_day = overload.groupby(overload['_dt'].dt.date).size().reset_index()
            ol_by_day.columns = ['Дата','Перегрузок']
            fig = px.bar(ol_by_day, x='Дата', y='Перегрузок',
                         title='Перегрузы самосвалов (>110%) по дням',
                         color_discrete_sequence=['#dc3545'])
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            # Overload by truck
            ol_truck = p.groupby('Truck').apply(
                lambda x: (x['LoadPercentage']>110).sum()).reset_index()
            ol_truck.columns = ['Самосвал','Перегрузок']
            fig = px.bar(ol_truck.sort_values('Перегрузок',ascending=False),
                         x='Самосвал', y='Перегрузок',
                         title='Перегрузы (>110%) по самосвалу',
                         color='Перегрузок', color_continuous_scale='Reds')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 7. Load weights balance
    wt_cols = ['LeftFrontLoadWeight','RightFrontLoadWeight',
               'LeftRearLoadWeight','RightRearLoadWeight']
    wt_avail= [c for c in wt_cols if c in p.columns and p[c].notna().any()]
    if wt_avail:
        st.markdown("**⚖️ Распределение нагрузки по осям**")
        wt_mean = p[wt_avail].mean()
        fig = px.bar(x=['Лев.перед','Прав.перед','Лев.зад','Прав.зад'][:len(wt_avail)],
                     y=wt_mean.values,
                     title='Средняя нагрузка по осям',
                     color=wt_mean.values, color_continuous_scale='RdYlGn_r')
        fig.update_layout(height=280, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 8. Time breakdown pie
    time_cols = {'StopLoadedTime':'Стоянка гружёный','MovingLoadedTime':'Движение гружёный',
                 'MovingEmptyTime':'Движение порожний','StopEmptyTime':'Стоянка порожний',
                 'LoadingTime':'Погрузка','DumpingTime':'Разгрузка'}
    t_avail = {k:v for k,v in time_cols.items() if k in p.columns and p[k].notna().any()}
    if t_avail:
        t_means = {v: p[k].mean() for k,v in t_avail.items()}
        fig = px.pie(names=list(t_means.keys()), values=list(t_means.values()),
                     title='Структура времени рейса (среднее)', hole=0.35,
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)


# ── Tab: GE Events ────────────────────────────────────────────────────────────

def tab_ge(ge):
    st.subheader("⚡ Система электропривода GE — журнал событий")

    if ge.empty:
        st.warning("Нет данных GE (Данные GE.7z не найдены или не распакованы)")
        return

    c1,c2,c3 = st.columns(3)
    c1.metric("Всего событий", len(ge))
    c2.metric("Самосвалов", ge['Truck'].nunique())
    c3.metric("Перегрузок", int((ge['Event_No']==635).sum()))
    st.markdown("---")

    co1, co2 = st.columns(2)

    # 1. Event distribution
    with co1:
        ev_cnt = ge['Event_No'].value_counts().head(15).reset_index()
        ev_cnt.columns = ['Event_No','Кол-во']
        # add name
        name_map = ge.drop_duplicates('Event_No').set_index('Event_No')['Name'].to_dict()
        ev_cnt['Название'] = ev_cnt['Event_No'].map(name_map).fillna(ev_cnt['Event_No'].astype(str))
        fig = px.bar(ev_cnt, x='Кол-во', y='Название', orientation='h',
                     title='Топ-15 типов событий GE',
                     color='Кол-во', color_continuous_scale='Blues')
        fig.update_layout(height=380, margin=dict(t=40,b=10),
                          yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    # 2. Timeline scatter
    with co2:
        ge_plot = ge[ge['Time'].notna()].copy()
        if not ge_plot.empty:
            fig = px.scatter(ge_plot, x='Time', y='Truck', color='Event_No',
                             hover_data=['Name','Restriction','TruckSpd'],
                             title='Временная шкала событий GE',
                             labels={'Event_No':'Код события'})
            fig.update_layout(height=380, margin=dict(t=40,b=10),
                              yaxis_title='Самосвал №')
            st.plotly_chart(fig, use_container_width=True)

    co3, co4 = st.columns(2)

    # 3. Overload events per truck
    with co3:
        overload = ge[ge['Event_No']==635]
        if not overload.empty:
            ov_cnt = overload['Truck'].value_counts().reset_index()
            ov_cnt.columns = ['Самосвал','Перегрузок']
            fig = px.bar(ov_cnt, x='Самосвал', y='Перегрузок',
                         title='Перегрузы самосвала (событие 635) по самосвалу',
                         color='Перегрузок', color_continuous_scale='Reds')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 4. Speed at service brake events
    with co4:
        brake = ge[(ge['Event_No']==647) & ge['TruckSpd'].notna()]
        if not brake.empty:
            fig = px.histogram(brake, x='TruckSpd', nbins=20,
                               title='Скорость при нажатии рабочего тормоза (событие 647)',
                               labels={'TruckSpd':'Скорость км/ч'},
                               color_discrete_sequence=['#e53935'])
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 5. Restricting events
    st.markdown("**🚫 Ограничивающие события**")
    restricting = ge[ge['Restriction'].astype(str).str.contains(
        'ограни|restrict', case=False, na=False) == False]
    if not restricting.empty:
        rcnt = restricting['Name'].value_counts().head(10).reset_index()
        rcnt.columns = ['Событие','Кол-во']
        fig = px.bar(rcnt, x='Кол-во', y='Событие', orientation='h',
                     title='Ограничивающие события (требуют внимания)',
                     color='Кол-во', color_continuous_scale='Oranges')
        fig.update_layout(height=280, margin=dict(t=40,b=10),
                          yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

    # 6. Raw table
    st.markdown("**📋 Журнал событий**")
    truck_f = st.multiselect("Самосвал:", sorted(ge['Truck'].unique()),
                             key='ge_truck_filter')
    show = ge[ge['Truck'].isin(truck_f)] if truck_f else ge
    st.dataframe(show[['Truck','Time','Type','Event_No','Name','Restriction',
                        'TruckSpd','State']].sort_values('Time', ascending=False),
                 use_container_width=True, height=300)


# ── Tab: Productivity ─────────────────────────────────────────────────────────

def tab_prod(prod):
    st.subheader("📈 Производительность парка")

    if prod.empty:
        st.warning("Нет данных о производительности (speed and production.csv)")
        return

    # detect columns
    cols = prod.columns.tolist()
    year_col = next((c for c in cols if 'year' in c.lower() or 'год' in c.lower()), None)
    mon_col  = next((c for c in cols if 'month' in c.lower() or 'мес' in c.lower()), None)
    speed_col= next((c for c in cols if 'speed' in c.lower() or 'скор' in c.lower()), None)
    tons_col = next((c for c in cols if 'ton' in c.lower() or 'тонн' in c.lower()
                     and 'оборот' not in c.lower()), None)
    load_col = next((c for c in cols if 'load' in c.lower() or 'загруз' in c.lower()), None)
    eq_col   = next((c for c in cols if 'equip' in c.lower() or 'модел' in c.lower()
                     or 'model' in c.lower()), None)
    name_col = next((c for c in cols if 'name' in c.lower() or 'имя' in c.lower()
                     or 'номер' in c.lower()), None)
    turn_col = next((c for c in cols if 'turn' in c.lower() or 'оборот' in c.lower()), None)

    if year_col and mon_col:
        def _mk_date(row):
            y = str(row[year_col]).split('.')[0]
            mraw = str(row[mon_col]).strip().lower()
            # numeric month?
            mnum = None
            mm = re.match(r'(\d{1,2})', mraw)
            if mm and int(mm.group(1)) <= 12:
                mnum = mm.group(1).zfill(2)
            else:
                for ru, num in RU_MONTHS.items():
                    if mraw.startswith(ru):
                        mnum = num
                        break
            if mnum:
                return pd.to_datetime(f"{y}-{mnum}-01", errors='coerce')
            return pd.NaT
        prod['_date'] = prod.apply(_mk_date, axis=1)

    co1, co2 = st.columns(2)

    # 1. Monthly tons trend
    with co1:
        if tons_col and '_date' in prod.columns:
            pt = prod.dropna(subset=['_date', tons_col])
            if eq_col:
                pt_grp = pt.groupby(['_date', eq_col])[tons_col].mean().reset_index()
                fig = px.line(pt_grp, x='_date', y=tons_col, color=eq_col,
                              title='Ежемесячная производительность, т/год (на единицу)',
                              labels={tons_col:'т/год','_date':'Дата'})
            else:
                pt_grp = pt.groupby('_date')[tons_col].mean().reset_index()
                fig = px.area(pt_grp, x='_date', y=tons_col,
                              title='Ежемесячная производительность, т/год')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 2. Speed distribution
    with co2:
        if speed_col:
            if eq_col:
                fig = px.box(prod.dropna(subset=[speed_col]),
                             x=eq_col, y=speed_col, color=eq_col,
                             title='Скорость гружёного рейса по модели',
                             labels={speed_col:'Скорость км/ч'})
            else:
                fig = px.histogram(prod.dropna(subset=[speed_col]),
                                   x=speed_col, nbins=20,
                                   title='Распределение скорости гружёного рейса')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    co3, co4 = st.columns(2)

    # 3. Average load by truck
    with co3:
        if load_col and name_col:
            ld = (prod.groupby(name_col)[load_col].mean()
                      .reset_index().sort_values(load_col, ascending=False).head(20))
            fig = px.bar(ld, x=name_col, y=load_col,
                         title='Средняя нагрузка по самосвалу, т',
                         color=load_col, color_continuous_scale='Greens')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 4. Cargo turnover trend
    with co4:
        if turn_col and '_date' in prod.columns:
            tc = prod.dropna(subset=['_date',turn_col])
            if eq_col:
                tc_grp = tc.groupby(['_date',eq_col])[turn_col].mean().reset_index()
                fig = px.area(tc_grp, x='_date', y=turn_col, color=eq_col,
                              title='Грузооборот (накопленный) по модели',
                              labels={turn_col:'т·км','_date':'Дата'})
            else:
                tc_grp = tc.groupby('_date')[turn_col].mean().reset_index()
                fig = px.area(tc_grp, x='_date', y=turn_col,
                              title='Грузооборот по месяцам')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 5. Radar: NTE200 vs 730E comparison
    if eq_col and speed_col and tons_col and load_col:
        st.markdown("**🔄 Сравнение NTE200 vs KOMATSU 730E**")
        params = [speed_col, tons_col, load_col]
        avail  = [p for p in params if p in prod.columns]
        groups = prod.groupby(eq_col)[avail].mean()
        if len(groups) >= 2:
            fig = go.Figure()
            palette_r = ['#1e88e5','#e53935','#43a047']
            for i, (idx, row) in enumerate(groups.iterrows()):
                vals = row.values.tolist() + [row.values[0]]
                cats = avail + [avail[0]]
                fig.add_scatterpolar(r=vals, theta=cats, fill='toself',
                                     name=str(idx), line_color=palette_r[i%3])
            fig.update_layout(title='Радарное сравнение моделей техники',
                              height=350, margin=dict(t=50,b=20))
            st.plotly_chart(fig, use_container_width=True)

    # raw data
    with st.expander("📋 Исходные данные производительности"):
        st.dataframe(prod, use_container_width=True, height=300)


# ── Tab: Oil Consumption ──────────────────────────────────────────────────────

def tab_oil(oil, maint, gbc):
    st.subheader("🛢️ Доливки моторного масла и анализ потребления")

    if oil.empty:
        st.warning("Нет данных о доливках (Доливки Горная Евразия.xlsx)")
        return

    c1,c2,c3 = st.columns(3)
    c1.metric("Записей доливок", len(oil))
    c2.metric("Суммарный объём, л", f"{oil['Объём_л'].sum():.0f}")
    c3.metric("Ср. объём доливки, л", f"{oil['Объём_л'].mean():.1f}")
    st.markdown("---")

    co1, co2 = st.columns(2)

    # 1. Total oil by truck
    with co1:
        ot = oil.groupby('Гараж_№')['Объём_л'].sum().reset_index()
        ot.columns = ['Самосвал','Объём_л']
        ot = ot.sort_values('Объём_л', ascending=False)
        fig = px.bar(ot, x='Самосвал', y='Объём_л',
                     title='Суммарные доливки масла по самосвалу, л',
                     color='Объём_л', color_continuous_scale='Oranges')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 2. Timeline
    with co2:
        od = oil.groupby('Дата')['Объём_л'].sum().reset_index()
        fig = px.area(od, x='Дата', y='Объём_л',
                      title='Динамика доливок масла по дням, л',
                      color_discrete_sequence=['#ff7043'])
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    co3, co4 = st.columns(2)

    # 3. Heatmap: truck × month
    with co3:
        oil2 = oil.copy()
        oil2['Месяц'] = oil2['Дата'].dt.to_period('M').astype(str)
        pv = oil2.pivot_table(index='Гараж_№', columns='Месяц',
                               values='Объём_л', aggfunc='sum').fillna(0)
        if not pv.empty:
            fig = px.imshow(pv, color_continuous_scale='Oranges',
                            title='Тепловая карта доливок: самосвал × месяц',
                            labels={'color':'Объём, л'})
            fig.update_layout(height=320, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 4. Oil vol vs engine hours scatter
    with co4:
        oil_hrs = oil.dropna(subset=['Мото_часы','Объём_л'])
        if not oil_hrs.empty:
            fig = px.scatter(oil_hrs, x='Мото_часы', y='Объём_л',
                             color='Гараж_№', trendline='ols',
                             title='Объём доливки vs моточасы ДВС',
                             labels={'Мото_часы':'Моточасы','Объём_л':'Объём, л'})
            fig.update_layout(height=320, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 5. Oil consumption per motor hour (rate)
    st.markdown("**📊 Удельный расход масла (л / 100 мото-часов)**")
    oil_rate = []
    for truck, grp in oil.groupby('Гараж_№'):
        total_l = grp['Объём_л'].sum()
        hrs_range = grp['Мото_часы'].max() - grp['Мото_часы'].min()
        if hrs_range and hrs_range > 0:
            rate = total_l / hrs_range * 100
            oil_rate.append({'Самосвал': truck, 'л/100мч': rate})
    if oil_rate:
        rate_df = pd.DataFrame(oil_rate).sort_values('л/100мч', ascending=False)
        avg_rate = rate_df['л/100мч'].mean()
        fig = px.bar(rate_df, x='Самосвал', y='л/100мч',
                     title='Удельный расход масла ДВС, л/100 мото-часов',
                     color='л/100мч', color_continuous_scale='Oranges')
        fig.add_hline(y=avg_rate, line_dash='dash', line_color='red',
                      annotation_text=f'Среднее: {avg_rate:.1f}')
        fig.update_layout(height=300, margin=dict(t=40,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # 6. Correlation oil consumption → GBC repairs
    if not gbc.empty and not oil.empty:
        st.markdown("**🔗 Корреляция: доливки масла → ремонты ГБЦ**")
        oil_total = oil.groupby('Гараж_№')['Объём_л'].sum().reset_index()
        oil_total.columns = ['Самосвал','Объём_л']
        gbc_cnt   = gbc.groupby('Самосвал').size().reset_index(name='Ремонтов_ГБЦ')
        merged    = oil_total.merge(gbc_cnt, on='Самосвал', how='inner')
        if len(merged) > 3:
            fig = px.scatter(merged, x='Объём_л', y='Ремонтов_ГБЦ',
                             text='Самосвал', trendline='ols',
                             title='Объём доливок масла vs количество ремонтов ГБЦ',
                             labels={'Объём_л':'Доливки масла, л',
                                     'Ремонтов_ГБЦ':'Ремонтов ГБЦ'})
            fig.update_traces(textposition='top center')
            fig.update_layout(height=300, margin=dict(t=40,b=10))
            st.plotly_chart(fig, use_container_width=True)

    # 7. Detail table
    with st.expander("📋 Журнал доливок"):
        st.dataframe(oil.sort_values('Дата', ascending=False),
                     use_container_width=True, height=300)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    # Header
    st.markdown('<div class="big-title">⚙️ QSK50 Predictive Analytics Dashboard</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Cummins QSK50 | NTE200 & KOMATSU 730E | '
                'Предиктивный анализ параметров двигателя</div>', unsafe_allow_html=True)
    st.markdown("---")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## ⚙️ Настройки")
        folder = st.text_input(
            "📁 Путь к папке с данными",
            value=str(Path(__file__).parent),
            help="Укажите путь к папке с CSV и XLSX файлами данных"
        )
        if not Path(folder).exists():
            st.error("Папка не найдена!")
            folder = str(Path(__file__).parent)

        if st.button("🔄 Обновить данные", type="primary"):
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")
        st.markdown("### 🔍 Фильтры")

        # Load maintenance first for truck list
        maint = load_maintenance(folder)
        all_trucks = (sorted(maint['Гараж_№'].dropna().unique().astype(int).tolist())
                      if not maint.empty else list(range(43, 93)))

        sel_trucks = st.multiselect("🚛 Самосвалы (NTE200):", all_trucks,
                                    default=all_trucks)

        st.markdown("---")
        st.markdown("### 💾 Экспорт")
        st.caption("Используйте кнопки экспорта на каждой вкладке")

        st.markdown("---")
        st.markdown("### ℹ️ О дашборде")
        st.caption("v1.0 | Данные: INSITE 8.x, ECM, GE, Весовая")
        st.caption(f"Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M')}")

    # ── Load all data ─────────────────────────────────────────────────────────
    extract_archives(folder)   # auto-unpack any *.7z in the folder (cached)
    with st.spinner("Загрузка данных..."):
        insite_dict = load_insite_csvs(folder)
        faults      = load_ecm_faults(folder)
        ge_events   = load_ge_events(folder)
        payload_df  = load_payload(folder)
        gbc_df      = load_gbc(folder)
        oil_df      = load_oil(folder)
        prod_df     = load_production(folder)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab_labels = [
        "📊 Обзор парка",
        "🔩 Параметры ДВС",
        "🔮 Предиктивный анализ",
        "⚠️ Коды ЭБУ",
        "🔧 История ТО",
        "⚖️ Нагрузка",
        "⚡ Привод GE",
        "📈 Производительность",
        "🛢️ Масло",
    ]
    tabs = st.tabs(tab_labels)

    with tabs[0]: tab_fleet(folder, maint, sel_trucks, faults, oil_df)
    with tabs[1]: tab_engine(insite_dict)
    with tabs[2]: tab_predictive(insite_dict)
    with tabs[3]: tab_faults(faults)
    with tabs[4]: tab_maint(maint, gbc_df)
    with tabs[5]: tab_payload(payload_df)
    with tabs[6]: tab_ge(ge_events)
    with tabs[7]: tab_prod(prod_df)
    with tabs[8]: tab_oil(oil_df, maint, gbc_df)


if __name__ == "__main__":
    main()
