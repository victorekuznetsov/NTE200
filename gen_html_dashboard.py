#!/usr/bin/env python3
"""Generate a standalone HTML predictive analytics dashboard for QSK50 engines."""

import io, json, re, pathlib, warnings
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.io as pio

warnings.filterwarnings('ignore')

FOLDER = pathlib.Path('/home/user/NTE200')
EXTRACT = FOLDER / '.qsk50_extracted'

# ─── COLOUR PALETTE ──────────────────────────────────────────────────────────
C = dict(
    bg='#141c24', bg2='#1c2630', card='#2a3138',
    acc='#3ef0af', red='#e84855', amb='#f5a623',
    grn='#22c55e', blu='#7e83fa', pur='#b668e4',
    tx='#e8edf6', tx2='#b3bddc', tx3='#80868b',
)
PLOTLY_THEME = dict(
    paper_bgcolor=C['bg2'], plot_bgcolor=C['card'],
    font_color=C['tx'], font_family='Segoe UI,sans-serif',
)

def _fig_json(fig):
    fig.update_layout(**PLOTLY_THEME,
                      margin=dict(l=40, r=20, t=40, b=40),
                      legend=dict(bgcolor='rgba(0,0,0,0)', borderwidth=0))
    fig.update_xaxes(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)')
    fig.update_yaxes(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)')
    return pio.to_json(fig, remove_uids=True)

# ─── INSITE CSV PARSER ───────────────────────────────────────────────────────
COL_ALIASES = {
    'rpm':      ['Engine Speed (RPM)- Identifier0', 'Частота вращения двигателя (об/мин)- Идентификатор0'],
    'tcool':    ['Engine Coolant Temperature (', 'Температура охлаждающей жидкости двигателя (°C)'],
    'toil':     ['Engine Oil Temperature (', 'Температура масла (°C)- Идентификатор0'],
    'poil':     ['Engine Oil Pressure (kPa)', 'Давление масла (кПа)- Идентификатор0'],
    'texh_avg': ['Average Exhaust Temperature (Calculated)', 'Средняя температура отработавших газов (расчетное значение)'],
    'fuel':     ['Instantaneous Fuel Rate', 'Мгновенный расход топлива (л/час)- Идентификатор0'],
    'crank':    ['Crankcase Pressure (kPa)', 'Давление картерных газов (кПа)'],
    'boost':    ['Intake Manifold Pressure (kPa)- Identifier0', 'Давление во впускном коллекторе (кПа)- Идентификатор0'],
    'load':     ['Percent Load', 'Относительная нагрузка (Проценты)'],
    'batt':     ['Battery Voltage (V)- Identifier0', 'Напряжение аккумуляторной батареи (В)- Идентификатор0'],
    'frail_m':  ['Fuel Rail Pressure Measured (bar)- Identifier0', 'Измеренное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
    'frail_c':  ['Fuel Rail Pressure Commanded (bar)- Identifier0', 'Заданное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'],
}
CYL_PATTERNS = [
    r'Exhaust Port.*?Cyl(\d+)',
    r'Temperature.*?(\d+)[LR]',
    r'Температура выпускного.*?цилиндр\s*(\d+)',
    r'Cylinder\s*(\d+).*?Exhaust',
    r'Температура.*?(\d+)\s*цил',
]


def _resolve(df, key):
    aliases = COL_ALIASES.get(key, [])
    for a in aliases:
        for c in df.columns:
            if a.lower() in c.lower():
                return c
    return None


def _parse_insite_csv(path):
    with open(path, encoding='cp1251', errors='replace') as fh:
        lines = fh.readlines()
    # find header row
    hdr_idx = next((i for i, l in enumerate(lines)
                    if l.startswith('"Date"') or l.startswith('Date,') or
                    l.startswith('"Дата"') or l.startswith('Дата,')), None)
    if hdr_idx is None:
        return None
    data_lines = [l for l in lines[hdr_idx:] if l.strip()]
    df = pd.read_csv(io.StringIO(''.join(data_lines)), low_memory=False)
    # numeric conversion (pandas 3.x StringDtype)
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        try:
            conv = df[col].astype(str).str.replace(',', '.', regex=False).str.strip()
            num = pd.to_numeric(conv, errors='coerce')
            if num.notna().sum() > len(df) * 0.3:
                df[col] = num
        except Exception:
            pass
    # parse datetime
    date_col = 'Дата' if 'Дата' in df.columns else 'Date'
    time_col = 'Время' if 'Время' in df.columns else 'Time'
    if date_col in df.columns and time_col in df.columns:
        try:
            df['_dt'] = pd.to_datetime(
                df[date_col].astype(str) + ' ' + df[time_col].astype(str),
                errors='coerce', dayfirst=True)
        except Exception:
            df['_dt'] = pd.NaT
    return df


def _truck_num(fname):
    for pat in [r'№(\d+)', r'N(\d+)', r'_(\d+)_', r'\s(\d{2,3})\s', r'\((\d+)\)']:
        m = re.search(pat, fname)
        if m:
            return int(m.group(1))
    m = re.search(r'DML-\d+-\d+\s+(\d+)', fname)
    if m:
        return int(m.group(1))
    return None


def load_all_insite():
    roots = [EXTRACT, FOLDER, EXTRACT / 'NTE200']
    files = {}
    for root in roots:
        if not root.exists():
            continue
        for fp in root.glob('*.csv'):
            if fp.name not in files:
                files[fp.name] = fp
    result = {}
    for name, fp in files.items():
        if 'speed and production' in name.lower():
            continue
        truck = _truck_num(name)
        if truck is None:
            continue
        df = _parse_insite_csv(fp)
        if df is None or len(df) < 10:
            continue
        model = 'NTE200' if ('NTE' in name or 'DML' in name) else '730E'
        key = f"{model}_{truck}"
        result[key] = {'df': df, 'truck': truck, 'model': model, 'name': name}
    return result


def load_ecm():
    ecm_dir = EXTRACT / 'Данные с ECM'
    rows = []
    for fp in ecm_dir.glob('*.xlsx'):
        m = re.search(r'(\d+)_ECM', fp.name)
        truck = int(m.group(1)) if m else None
        try:
            df = pd.read_excel(fp, sheet_name='Fault Codes', header=6)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            if 'Fault Codes' in df.columns:
                df['truck'] = truck
                rows.append(df[['Fault Codes', 'Status', 'Counts', 'Lamp', 'Description', 'truck']].copy())
        except Exception:
            pass
    if not rows:
        return pd.DataFrame()
    out = pd.concat(rows, ignore_index=True)
    out['Counts'] = pd.to_numeric(out['Counts'], errors='coerce').fillna(0).astype(int)
    out['Fault Codes'] = pd.to_numeric(out['Fault Codes'], errors='coerce')
    return out.dropna(subset=['Fault Codes'])


def load_maintenance():
    fp = FOLDER / 'ОТЧЕТ Полюс Магадан.xlsx'
    df = pd.read_excel(fp)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_gbc():
    fp = FOLDER / 'ГБЦ ремонты.xlsx'
    df = pd.read_excel(fp, header=1)
    return df


def load_oil():
    fp = FOLDER / 'Доливки Горная Евразия.xlsx'
    df = pd.read_excel(fp, header=None)
    # find header row
    hdr = None
    for i, row in df.iterrows():
        if 'Дата' in str(row.values):
            hdr = i
            break
    if hdr is None:
        return pd.DataFrame()
    df2 = pd.read_excel(fp, header=hdr)
    df2.columns = [str(c).strip() for c in df2.columns]
    return df2


def load_production():
    fp = FOLDER / 'speed and production.csv'
    for enc in ['utf-8-sig', 'utf-8', 'cp1251']:
        try:
            return pd.read_csv(fp, encoding=enc)
        except Exception:
            pass
    return pd.read_csv(fp, encoding_errors='replace')


def load_ge():
    ge_dir = EXTRACT / 'Данные GE'
    rows = []
    for fp in ge_dir.glob('*.xlsx'):
        m = re.search(r'GE_(\d+)_', fp.name)
        truck = int(m.group(1)) if m else None
        try:
            df = pd.read_excel(fp, header=7)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            df['truck'] = truck
            rows.append(df)
        except Exception:
            pass
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


def load_payload():
    rows = []
    for fp in (EXTRACT).rglob('Весовая*.xlsx'):
        m = re.search(r'Весовая_(\d+)', fp.name)
        truck = int(m.group(1)) if m else None
        try:
            df = pd.read_excel(fp, header=0)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            df['truck'] = truck
            rows.append(df)
        except Exception:
            pass
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


# ─── LOAD DATA ───────────────────────────────────────────────────────────────
print('Loading data...')
insite = load_all_insite()
print(f'  INSITE: {len(insite)} files: {list(insite.keys())[:5]}')
ecm_df = load_ecm()
print(f'  ECM: {len(ecm_df)} records')
maint_df = load_maintenance()
print(f'  Maintenance: {len(maint_df)} records')
gbc_df = load_gbc()
oil_df = load_oil()
prod_df = load_production()
ge_df = load_ge()
payload_df = load_payload()
print('  All data loaded.')

# ─── HEALTH SCORE ─────────────────────────────────────────────────────────────
def score_param(val, ok_lo, ok_hi, warn_hi, higher_is_bad=True):
    if pd.isna(val):
        return 70
    if higher_is_bad:
        if val <= ok_hi:
            return 100
        if val <= warn_hi:
            return 60
        return 20
    else:
        if val >= ok_lo:
            return 100
        if val >= warn_hi:
            return 60
        return 20


def calc_health(df):
    scores = {}
    c = _resolve(df, 'tcool')
    if c:
        scores['Охл. жидкость'] = score_param(df[c].median(), 82, 95, 105)
    c = _resolve(df, 'poil')
    if c:
        scores['Давл. масла'] = score_param(df[c].median(), 300, 350, 350, higher_is_bad=False)
    c = _resolve(df, 'toil')
    if c:
        scores['T масла'] = score_param(df[c].median(), 85, 105, 115)
    c = _resolve(df, 'texh_avg')
    if c:
        scores['Выхлоп'] = score_param(df[c].median(), 380, 520, 560)
    c = _resolve(df, 'crank')
    if c:
        scores['Картер'] = score_param(df[c].median(), 0, 3, 7)
    if scores:
        return round(np.mean(list(scores.values())), 1)
    return 75


# ─── FIGURE BUILDERS ─────────────────────────────────────────────────────────

def fig_fleet_health():
    """Tab 1 - fleet health matrix."""
    rows = []
    for key, info in insite.items():
        df = info['df']
        truck = info['truck']
        model = info['model']
        row = {'Самосвал': f"{model} №{truck}"}
        for param, label in [('tcool', 'T охл.'), ('poil', 'P масла'),
                              ('toil', 'T масла'), ('texh_avg', 'T выхл.'),
                              ('crank', 'P картера'), ('fuel', 'Топливо')]:
            col = _resolve(df, param)
            row[label] = round(df[col].median(), 1) if col else None
        row['Health'] = calc_health(df)
        rows.append(row)
    if not rows:
        return go.Figure()
    hdf = pd.DataFrame(rows).set_index('Самосвал')
    # normalise 0-100 per column
    norm = hdf.copy()
    for c in norm.columns:
        lo, hi = norm[c].min(), norm[c].max()
        if hi > lo:
            norm[c] = (norm[c] - lo) / (hi - lo) * 100
        else:
            norm[c] = 50
    fig = go.Figure(go.Heatmap(
        z=norm.values.tolist(),
        x=list(norm.columns),
        y=list(norm.index),
        colorscale=[[0, C['red']], [0.5, C['amb']], [1, C['grn']]],
        text=hdf.values.round(1).astype(str).tolist(),
        texttemplate='%{text}',
        textfont_size=10,
        showscale=True,
        colorbar=dict(title='Норм.', tickfont_color=C['tx3']),
    ))
    fig.update_layout(title='Матрица здоровья парка', height=max(300, len(rows)*35+80))
    return fig


def fig_downtime_heatmap():
    """Tab 1 - downtime heatmap."""
    df = maint_df.copy()
    # truck column
    truck_col = next((c for c in df.columns if 'гараж' in c.lower()), None)
    date_col = next((c for c in df.columns if 'Дата' in c and 'запуска' not in c), None)
    dt_col = next((c for c in df.columns if 'простоя' in c.lower() and 'час' in c.lower() and '(' not in c), None)
    if not (truck_col and date_col and dt_col):
        return go.Figure()
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df['date'] = pd.to_datetime(df[date_col], errors='coerce')
    df['dt_h'] = pd.to_numeric(
        df[dt_col].astype(str).str.extract(r'(\d+\.\d+|\d+)')[0], errors='coerce')
    df['week'] = df['date'].dt.strftime('%Y-W%V')
    pivot = df.pivot_table(index='truck', columns='week', values='dt_h', aggfunc='sum').fillna(0)
    pivot = pivot.sort_index()
    fig = go.Figure(go.Heatmap(
        z=pivot.values.tolist(),
        x=[str(c) for c in pivot.columns],
        y=[f'№{t}' for t in pivot.index],
        colorscale=[[0, C['bg2']], [0.5, C['amb']], [1, C['red']]],
        colorbar=dict(title='Часы', tickfont_color=C['tx3']),
    ))
    fig.update_layout(title='Простои по неделям (часы)', height=max(300, len(pivot)*22+100))
    return fig


def fig_pareto_failures():
    df = maint_df.copy()
    work_col = next((c for c in df.columns if 'задание' in c.lower() or 'работ' in c.lower()), None)
    if not work_col:
        return go.Figure()
    counts = df[work_col].astype(str).str[:60].value_counts().head(20)
    fig = go.Figure(go.Bar(
        x=counts.values.tolist(), y=counts.index.tolist(),
        orientation='h', marker_color=C['acc'],
        text=counts.values.tolist(), textposition='auto',
    ))
    fig.update_layout(title='Парето отказов (топ-20)', height=500,
                      yaxis=dict(autorange='reversed'))
    return fig


def fig_mtbf():
    df = maint_df.copy()
    truck_col = next((c for c in df.columns if 'гараж' in c.lower()), None)
    mh_col = next((c for c in df.columns if 'мото' in c.lower()), None)
    if not (truck_col and mh_col):
        return go.Figure()
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df['mh'] = pd.to_numeric(df[mh_col], errors='coerce')
    grouped = df.groupby('truck').agg(events=('truck', 'count'), max_mh=('mh', 'max')).reset_index()
    grouped['MTBF'] = (grouped['max_mh'] / grouped['events'].clip(lower=1)).round(0)
    grouped = grouped.sort_values('MTBF', ascending=False).head(30)
    mean_mtbf = grouped['MTBF'].mean()
    fig = go.Figure()
    fig.add_bar(x=[f'№{t}' for t in grouped['truck']], y=grouped['MTBF'].tolist(),
                marker_color=[C['grn'] if v >= mean_mtbf else C['amb'] for v in grouped['MTBF']],
                name='MTBF')
    fig.add_hline(y=mean_mtbf, line_dash='dash', line_color=C['red'],
                  annotation_text=f'Ср. {mean_mtbf:.0f} м/ч')
    fig.update_layout(title='MTBF по самосвалам (м/ч между ремонтами)', height=350)
    return fig


def fig_kpi_gauges():
    trucks = len(set([info['truck'] for info in insite.values()]))
    dt_col = next((c for c in maint_df.columns if 'простоя' in c.lower() and '(' not in c), None)
    tot_down = 0
    if dt_col:
        vals = pd.to_numeric(
            maint_df[dt_col].astype(str).str.extract(r'(\d+\.\d+|\d+)')[0],
            errors='coerce')
        tot_down = vals.sum()
    active_faults = len(ecm_df[ecm_df['Status'].astype(str).str.lower() == 'active']) if len(ecm_df) else 0
    avg_health = np.mean([calc_health(v['df']) for v in insite.values()]) if insite else 75

    fig = make_subplots(rows=1, cols=4, specs=[[{'type': 'indicator'}]*4])
    fig.add_trace(go.Indicator(mode='number', value=trucks,
        title={'text': 'Машин в анализе'}, number={'font': {'color': C['acc']}}), row=1, col=1)
    fig.add_trace(go.Indicator(mode='number', value=round(tot_down, 0),
        title={'text': 'Суммарный простой (ч)'}, number={'font': {'color': C['amb']}}), row=1, col=2)
    fig.add_trace(go.Indicator(mode='number', value=active_faults,
        title={'text': 'Активных неиспр.'}, number={'font': {'color': C['red']}}), row=1, col=3)
    fig.add_trace(go.Indicator(mode='gauge+number', value=round(avg_health, 0),
        title={'text': 'Ср. здоровье ДВС'},
        gauge={'axis': {'range': [0, 100]},
               'bar': {'color': C['grn']},
               'steps': [{'range': [0, 40], 'color': C['red']},
                          {'range': [40, 70], 'color': C['amb']},
                          {'range': [70, 100], 'color': C['grn']}]}), row=1, col=4)
    fig.update_layout(title='KPI парка', height=200)
    return fig


def fig_engine_rpm(df, truck_label=''):
    col = _resolve(df, 'rpm')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    ewma = sub[col].ewm(alpha=0.05).mean()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=sub[col].tolist(),
                             mode='lines', name='RPM', line=dict(color=C['blu'], width=1),
                             opacity=0.6))
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=ewma.tolist(),
                             mode='lines', name='EWMA', line=dict(color=C['acc'], width=2)))
    fig.add_hline(y=1800, line_dash='dash', line_color=C['amb'], annotation_text='Ном. 1800')
    fig.update_layout(title=f'Обороты двигателя {truck_label}', height=280)
    return fig


def fig_coolant_temp(df, truck_label=''):
    col = _resolve(df, 'tcool')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    vals = sub[col].tolist()
    fig = go.Figure()
    fig.add_hrect(y0=82, y1=95, fillcolor=C['grn'], opacity=0.08, line_width=0)
    fig.add_hrect(y0=95, y1=105, fillcolor=C['amb'], opacity=0.10, line_width=0)
    fig.add_hrect(y0=105, y1=130, fillcolor=C['red'], opacity=0.10, line_width=0)
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=vals,
                             mode='lines', name='T охл.', line=dict(color=C['acc'], width=2)))
    fig.add_hline(y=95, line_dash='dot', line_color=C['amb'], annotation_text='⚠ 95°C')
    fig.add_hline(y=105, line_dash='dot', line_color=C['red'], annotation_text='🔴 105°C')
    fig.update_layout(title=f'Температура охл. жидкости {truck_label}', height=280,
                      yaxis_title='°C')
    return fig


def fig_oil_combined(df, truck_label=''):
    col_t = _resolve(df, 'toil')
    col_p = _resolve(df, 'poil')
    if (col_t is None and col_p is None) or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt']).head(2000)
    fig = make_subplots(specs=[[{'secondary_y': True}]])
    if col_t:
        fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=sub[col_t].tolist(),
                                 name='T масла (°C)', line=dict(color=C['amb'], width=2)),
                      secondary_y=False)
    if col_p:
        fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=sub[col_p].tolist(),
                                 name='P масла (кПа)', line=dict(color=C['blu'], width=2)),
                      secondary_y=True)
    fig.update_yaxes(title_text='°C', secondary_y=False)
    fig.update_yaxes(title_text='кПа', secondary_y=True)
    fig.update_layout(title=f'T и P масла {truck_label}', height=280)
    return fig


def fig_exhaust_heatmap(df, truck_label=''):
    cyl_cols = []
    for col in df.columns:
        for pat in CYL_PATTERNS:
            m = re.search(pat, col, re.I)
            if m:
                cyl_cols.append((int(m.group(1)), col))
                break
    if not cyl_cols or '_dt' not in df.columns:
        return go.Figure()
    cyl_cols = sorted(set(cyl_cols))[:16]
    sub = df.dropna(subset=['_dt']).head(500)
    z = []
    y_labels = []
    for cnum, ccol in cyl_cols:
        vals = pd.to_numeric(sub[ccol], errors='coerce').fillna(0).tolist()
        z.append(vals)
        y_labels.append(f'Цил. {cnum}')
    fig = go.Figure(go.Heatmap(
        z=z,
        x=list(range(len(sub))),
        y=y_labels,
        colorscale=[[0, C['blu']], [0.5, C['amb']], [1, C['red']]],
        colorbar=dict(title='°C', tickfont_color=C['tx3']),
    ))
    fig.update_layout(title=f'Тепловая карта цилиндров {truck_label}', height=380)
    return fig


def fig_cylinder_balance(df, truck_label=''):
    cyl_cols = []
    for col in df.columns:
        for pat in CYL_PATTERNS:
            m = re.search(pat, col, re.I)
            if m:
                cyl_cols.append((int(m.group(1)), col))
                break
    if not cyl_cols:
        return go.Figure()
    cyl_cols = sorted(set(cyl_cols))[:16]
    avgs = {}
    for cnum, ccol in cyl_cols:
        v = pd.to_numeric(df[ccol], errors='coerce').median()
        avgs[f'Цил. {cnum}'] = v
    if not avgs:
        return go.Figure()
    mean_val = np.nanmean(list(avgs.values()))
    deviations = {k: v - mean_val for k, v in avgs.items()}
    colors = [C['red'] if abs(v) > 60 else C['amb'] if abs(v) > 30 else C['grn']
              for v in deviations.values()]
    fig = go.Figure(go.Bar(
        x=list(deviations.keys()), y=list(deviations.values()),
        marker_color=colors, text=[f'{v:+.0f}' for v in deviations.values()],
        textposition='auto',
    ))
    fig.add_hline(y=30, line_dash='dash', line_color=C['amb'])
    fig.add_hline(y=-30, line_dash='dash', line_color=C['amb'])
    fig.add_hline(y=60, line_dash='dash', line_color=C['red'])
    fig.add_hline(y=-60, line_dash='dash', line_color=C['red'])
    fig.update_layout(title=f'Отклонение T цилиндров от среднего {truck_label}', height=280,
                      yaxis_title='∆°C')
    return fig


def fig_fuel_rate(df, truck_label=''):
    col = _resolve(df, 'fuel')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    roll = sub[col].rolling(60, min_periods=1).mean()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=sub[col].tolist(),
                             mode='lines', name='Мгн. расход', line=dict(color=C['pur'], width=1),
                             opacity=0.5))
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=roll.tolist(),
                             mode='lines', name='Скол. ср. 60с', line=dict(color=C['acc'], width=2)))
    fig.update_layout(title=f'Расход топлива {truck_label}', height=280, yaxis_title='л/час')
    return fig


def fig_boost_pressure(df, truck_label=''):
    col = _resolve(df, 'boost')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    fig = go.Figure(go.Scatter(x=sub['_dt'].tolist(), y=sub[col].tolist(),
                               mode='lines', name='P наддува', line=dict(color=C['blu'], width=2)))
    fig.add_hline(y=180, line_dash='dash', line_color=C['red'], annotation_text='Мин 180 кПа')
    fig.update_layout(title=f'Давление наддува {truck_label}', height=280, yaxis_title='кПа')
    return fig


def fig_crankcase(df, truck_label=''):
    col = _resolve(df, 'crank')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    fig = go.Figure(go.Scatter(x=sub['_dt'].tolist(), y=sub[col].tolist(),
                               mode='lines', name='P картера', line=dict(color=C['red'], width=2)))
    fig.add_hline(y=3, line_dash='dash', line_color=C['amb'], annotation_text='⚠ 3 кПа')
    fig.add_hline(y=7, line_dash='dash', line_color=C['red'], annotation_text='🔴 7 кПа')
    fig.update_layout(title=f'Давление картерных газов {truck_label}', height=280, yaxis_title='кПа')
    return fig


def fig_load_percent(df, truck_label=''):
    col = _resolve(df, 'load')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000)
    fig = go.Figure(go.Scatter(x=sub['_dt'].tolist(), y=sub[col].tolist(),
                               mode='lines', name='Нагрузка', line=dict(color=C['grn'], width=2),
                               fill='tozeroy', fillcolor=f'rgba(34,197,94,0.1)'))
    fig.update_layout(title=f'Нагрузка на двигатель {truck_label}', height=280, yaxis_title='%')
    return fig


def fig_stats_table(df, truck_label=''):
    params = []
    for key, label in [('rpm', 'RPM'), ('tcool', 'T охл. °C'), ('toil', 'T масла °C'),
                       ('poil', 'P масла кПа'), ('texh_avg', 'T выхл. °C'),
                       ('fuel', 'Топливо л/час'), ('crank', 'P картера кПа'),
                       ('boost', 'P наддува кПа'), ('load', 'Нагрузка %')]:
        col = _resolve(df, key)
        if col:
            s = pd.to_numeric(df[col], errors='coerce').dropna()
            if len(s):
                params.append({'Параметр': label, 'Min': round(s.min(), 1),
                                'Max': round(s.max(), 1), 'Avg': round(s.mean(), 1),
                                'Std': round(s.std(), 1), 'Медиана': round(s.median(), 1)})
    if not params:
        return go.Figure()
    tdf = pd.DataFrame(params)
    fig = go.Figure(go.Table(
        header=dict(values=list(tdf.columns), fill_color=C['card'],
                    font=dict(color=C['acc']), align='left'),
        cells=dict(values=[tdf[c].tolist() for c in tdf.columns],
                   fill_color=C['bg2'], font=dict(color=C['tx']), align='left'),
    ))
    fig.update_layout(title=f'Статистика параметров {truck_label}', height=320)
    return fig


def fig_health_score_gauge(df, truck_label=''):
    score = calc_health(df)
    color = C['grn'] if score >= 70 else C['amb'] if score >= 40 else C['red']
    fig = go.Figure(go.Indicator(
        mode='gauge+number+delta',
        value=score,
        delta={'reference': 70, 'increasing': {'color': C['grn']}, 'decreasing': {'color': C['red']}},
        title={'text': f'Индекс здоровья ДВС {truck_label}', 'font': {'size': 14}},
        gauge={
            'axis': {'range': [0, 100], 'tickcolor': C['tx3']},
            'bar': {'color': color},
            'steps': [{'range': [0, 40], 'color': C['red']},
                      {'range': [40, 70], 'color': C['amb']},
                      {'range': [70, 100], 'color': C['grn']}],
            'threshold': {'line': {'color': 'white', 'width': 3}, 'value': 70},
        },
        number={'font': {'color': color, 'size': 36}},
    ))
    fig.update_layout(height=280)
    return fig


def fig_anomalies(df, truck_label=''):
    col = _resolve(df, 'tcool')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000).copy()
    sub['z'] = (sub[col] - sub[col].rolling(60, min_periods=1).mean()) / \
                (sub[col].rolling(60, min_periods=1).std() + 0.001)
    anomalies = sub[sub['z'].abs() > 3]
    normal = sub[sub['z'].abs() <= 3]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=normal['_dt'].tolist(), y=normal[col].tolist(),
                             mode='markers', name='Норма',
                             marker=dict(color=C['grn'], size=3, opacity=0.6)))
    fig.add_trace(go.Scatter(x=anomalies['_dt'].tolist(), y=anomalies[col].tolist(),
                             mode='markers', name='Аномалии',
                             marker=dict(color=C['red'], size=8, symbol='x')))
    fig.update_layout(title=f'Аномалии T охл. жидкости {truck_label}', height=280)
    return fig


def fig_oil_pressure_forecast(df, truck_label=''):
    col = _resolve(df, 'poil')
    if col is None or '_dt' not in df.columns:
        return go.Figure()
    sub = df.dropna(subset=['_dt', col]).head(2000).copy()
    sub = sub.set_index('_dt').resample('5min').median().reset_index().dropna()
    if len(sub) < 10:
        return go.Figure()
    from scipy import stats
    x_num = np.arange(len(sub))
    y = sub[col].values
    slope, intercept, r, p, se = stats.linregress(x_num, y)
    future_n = 30
    x_future = np.arange(len(sub), len(sub) + future_n)
    forecast = slope * x_future + intercept
    ci = 1.96 * se * np.sqrt(1 + 1/len(sub) + (x_future - x_num.mean())**2 / np.sum((x_num - x_num.mean())**2))
    last_dt = sub['_dt'].iloc[-1]
    future_dt = [last_dt + pd.Timedelta(minutes=5*(i+1)) for i in range(future_n)]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sub['_dt'].tolist(), y=y.tolist(),
                             name='P масла', line=dict(color=C['blu'], width=2)))
    fig.add_trace(go.Scatter(x=future_dt, y=forecast.tolist(),
                             name='Прогноз', line=dict(color=C['acc'], width=2, dash='dot')))
    fig.add_trace(go.Scatter(
        x=future_dt + future_dt[::-1],
        y=(forecast + ci).tolist() + (forecast - ci).tolist()[::-1],
        fill='toself', fillcolor='rgba(62,240,175,0.1)', line_color='rgba(0,0,0,0)',
        name='95% ДИ'))
    fig.add_hline(y=300, line_dash='dash', line_color=C['red'], annotation_text='Критично <300')
    fig.update_layout(title=f'Тренд P масла + прогноз {truck_label}', height=320, yaxis_title='кПа')
    return fig


def fig_risk_table(df, truck_label=''):
    params = []
    thresholds = {
        'T охл. °C': (_resolve(df, 'tcool'), 82, 95, 105, True),
        'P масла кПа': (_resolve(df, 'poil'), 300, 350, None, False),
        'T масла °C': (_resolve(df, 'toil'), 85, 105, 115, True),
        'T выхл. °C': (_resolve(df, 'texh_avg'), 380, 520, 560, True),
        'P картера кПа': (_resolve(df, 'crank'), 0, 3, 7, True),
        'P наддува кПа': (_resolve(df, 'boost'), 180, 220, None, False),
    }
    for label, (col, ok_lo, ok_hi, warn_hi, higher_bad) in thresholds.items():
        if col is None:
            continue
        s = pd.to_numeric(df[col], errors='coerce').dropna()
        if not len(s):
            continue
        val = s.median()
        ewma = s.ewm(alpha=0.05).mean()
        trend_slope = float(np.polyfit(np.arange(min(100, len(ewma))), ewma.values[-100:], 1)[0])
        trend_arrow = '↑' if trend_slope > 0.01 else '↓' if trend_slope < -0.01 else '→'
        if higher_bad:
            risk = 'В' if (warn_hi and val > warn_hi) else 'С' if val > ok_hi else 'Н'
        else:
            risk = 'В' if val < ok_lo else 'С' if val < ok_hi else 'Н'
        risk_color = C['red'] if risk == 'В' else C['amb'] if risk == 'С' else C['grn']
        params.append({'Параметр': label, 'Текущее': round(val, 1), 'Тренд': trend_arrow,
                       'Риск': risk})
    if not params:
        return go.Figure()
    tdf = pd.DataFrame(params)
    cell_colors = [[C['bg2']] * len(tdf),
                   [C['bg2']] * len(tdf),
                   [C['bg2']] * len(tdf),
                   [C['red'] if r == 'В' else C['amb'] if r == 'С' else C['grn']
                    for r in tdf['Риск']]]
    fig = go.Figure(go.Table(
        header=dict(values=list(tdf.columns), fill_color=C['card'],
                    font=dict(color=C['acc']), align='center'),
        cells=dict(values=[tdf[c].tolist() for c in tdf.columns],
                   fill_color=cell_colors, font=dict(color=C['tx']), align='center'),
    ))
    fig.update_layout(title=f'Таблица рисков {truck_label}', height=280)
    return fig


def fig_ecm_heatmap():
    if len(ecm_df) == 0:
        return go.Figure()
    pivot = ecm_df.pivot_table(index='truck', columns='Fault Codes',
                                values='Counts', aggfunc='sum').fillna(0)
    # top 20 codes
    top_codes = pivot.sum().nlargest(20).index
    pivot = pivot[top_codes]
    fig = go.Figure(go.Heatmap(
        z=pivot.values.tolist(),
        x=[str(int(c)) for c in pivot.columns],
        y=[f'№{t}' for t in pivot.index],
        colorscale=[[0, C['bg2']], [0.5, C['amb']], [1, C['red']]],
        colorbar=dict(title='Кол-во', tickfont_color=C['tx3']),
        text=pivot.values.round(0).astype(int).astype(str).tolist(),
        texttemplate='%{text}',
    ))
    fig.update_layout(title='Тепловая карта кодов ЭБУ (самосвал × код)', height=max(300, len(pivot)*28+100))
    return fig


def fig_ecm_pareto():
    if len(ecm_df) == 0:
        return go.Figure()
    top = ecm_df.groupby('Fault Codes')['Counts'].sum().nlargest(15).reset_index()
    desc = ecm_df.drop_duplicates('Fault Codes').set_index('Fault Codes')['Description']
    top['label'] = top['Fault Codes'].astype(int).astype(str) + ': ' + \
                   top['Fault Codes'].map(desc).fillna('').str[:40]
    fig = go.Figure(go.Bar(
        x=top['Counts'].tolist(), y=top['label'].tolist(),
        orientation='h', marker_color=C['red'],
        text=top['Counts'].tolist(), textposition='auto',
    ))
    fig.update_layout(title='Топ-15 кодов ЭБУ по частоте', height=420,
                      yaxis=dict(autorange='reversed'))
    return fig


def fig_ecm_active():
    if len(ecm_df) == 0:
        return go.Figure()
    active = ecm_df[ecm_df['Status'].astype(str).str.lower() == 'active'].copy()
    if len(active) == 0:
        active = ecm_df[ecm_df['Status'].astype(str).str.contains('акт|activ', case=False, na=False)]
    fig = go.Figure(go.Table(
        header=dict(values=['Truck', 'Код', 'Статус', 'Лампа', 'Описание'],
                    fill_color=C['red'], font=dict(color='white'), align='left'),
        cells=dict(
            values=[
                [f'№{t}' for t in active['truck'].tolist()],
                active['Fault Codes'].astype(int).astype(str).tolist(),
                active['Status'].tolist(),
                active['Lamp'].fillna('–').tolist(),
                active['Description'].astype(str).str[:60].tolist(),
            ],
            fill_color=C['bg2'], font=dict(color=C['tx']), align='left'),
    ))
    fig.update_layout(title='Активные неисправности ЭБУ', height=max(200, len(active)*30+80))
    return fig


def fig_ecm_lamp_donut():
    if len(ecm_df) == 0:
        return go.Figure()
    lamp_counts = ecm_df['Lamp'].fillna('Без лампы').value_counts()
    colors = {
        'Red': C['red'], 'Amber': C['amb'], 'Maintenance': C['blu'],
        'Без лампы': C['tx3'],
    }
    fig = go.Figure(go.Pie(
        labels=lamp_counts.index.tolist(), values=lamp_counts.values.tolist(),
        hole=0.5,
        marker_colors=[colors.get(l, C['grn']) for l in lamp_counts.index],
    ))
    fig.update_layout(title='Распределение по типу лампы', height=300)
    return fig


def fig_maint_timeline():
    df = maint_df.copy()
    date_col = next((c for c in df.columns if 'Дата' in c and 'запуска' not in c), None)
    truck_col = next((c for c in df.columns if 'гараж' in c.lower()), None)
    dt_col = next((c for c in df.columns if 'простоя' in c.lower() and '(' not in c), None)
    if not (date_col and truck_col):
        return go.Figure()
    df['date'] = pd.to_datetime(df[date_col], errors='coerce')
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df['dt_h'] = pd.to_numeric(
        df[dt_col].astype(str).str.extract(r'(\d+\.\d+|\d+)')[0], errors='coerce') if dt_col else 1
    df = df.dropna(subset=['date', 'truck']).head(500)
    df['dt_h'] = df['dt_h'].fillna(1.0)
    fig = go.Figure(go.Scatter(
        x=df['date'].tolist(), y=[f'№{int(t)}' for t in df['truck'].tolist()],
        mode='markers',
        marker=dict(size=(df['dt_h'].clip(0.5, 24) * 2).round(1).tolist(),
                    color=df['dt_h'].tolist(), colorscale='RdYlGn_r',
                    showscale=True, colorbar=dict(title='Ч прост.', tickfont_color=C['tx3'])),
    ))
    fig.update_layout(title='Временная шкала событий ТО', height=max(300, len(df['truck'].unique())*25+100))
    return fig


def fig_gbc_heatmap():
    df = gbc_df.copy()
    # columns: index 0,1,2,3 = truck/date/mh/note, then cylinder cols
    cyl_cols = [c for c in df.columns if '/' in str(c) and ('L' in str(c) or 'R' in str(c))]
    truck_col = df.columns[0]
    if not cyl_cols:
        return go.Figure()
    # fill down truck numbers
    df[truck_col] = df[truck_col].ffill()
    df = df.dropna(subset=[truck_col])
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df = df.dropna(subset=['truck'])
    heat = pd.DataFrame(index=df['truck'].unique())
    for cc in cyl_cols:
        heat[cc] = df.groupby('truck')[cc].apply(lambda x: x.notna().sum())
    heat = heat.fillna(0).astype(int)
    fig = go.Figure(go.Heatmap(
        z=heat.values.tolist(),
        x=[str(c) for c in heat.columns],
        y=[f'№{int(t)}' for t in heat.index],
        colorscale=[[0, C['bg2']], [0.5, C['amb']], [1, C['red']]],
        text=heat.values.astype(str).tolist(), texttemplate='%{text}',
        colorbar=dict(title='Ремонтов', tickfont_color=C['tx3']),
    ))
    fig.update_layout(title='Ремонты ГБЦ по цилиндрам', height=max(250, len(heat)*28+80))
    return fig


def fig_downtime_bar():
    df = maint_df.copy()
    truck_col = next((c for c in df.columns if 'гараж' in c.lower()), None)
    dt_col = next((c for c in df.columns if 'простоя' in c.lower() and '(' not in c), None)
    if not (truck_col and dt_col):
        return go.Figure()
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df['dt_h'] = pd.to_numeric(
        df[dt_col].astype(str).str.extract(r'(\d+\.\d+|\d+)')[0], errors='coerce')
    total = df.groupby('truck')['dt_h'].sum().sort_values(ascending=False).head(25)
    fig = go.Figure(go.Bar(
        x=[f'№{int(t)}' for t in total.index],
        y=total.values.tolist(),
        marker_color=[C['red'] if v > total.mean()*1.5 else C['amb'] if v > total.mean() else C['grn']
                      for v in total.values],
        text=[f'{v:.0f}ч' for v in total.values],
        textposition='auto',
    ))
    fig.update_layout(title='Суммарный простой по самосвалу (часы)', height=320)
    return fig


def fig_ge_events():
    if len(ge_df) == 0:
        return go.Figure()
    ev_col = next((c for c in ge_df.columns if 'event' in str(c).lower() and '#' in str(c)), None)
    if not ev_col:
        ev_col = ge_df.columns[2] if len(ge_df.columns) > 2 else None
    if not ev_col:
        return go.Figure()
    counts = pd.to_numeric(ge_df[ev_col], errors='coerce').value_counts().head(15)
    fig = go.Figure(go.Bar(
        x=[str(int(k)) for k in counts.index],
        y=counts.values.tolist(),
        marker_color=C['pur'],
        text=counts.values.tolist(), textposition='auto',
    ))
    fig.update_layout(title='Распределение типов событий GE (топ-15)', height=320, xaxis_title='Event#')
    return fig


def fig_payload_distribution():
    if len(payload_df) == 0:
        return go.Figure()
    col = next((c for c in payload_df.columns if 'load' in str(c).lower() and 'percent' in str(c).lower()), None)
    if col is None:
        col = next((c for c in payload_df.columns if 'loadper' in str(c).lower().replace(' ', '')), None)
    if col is None:
        return go.Figure()
    vals = pd.to_numeric(payload_df[col], errors='coerce').dropna()
    vals = vals[(vals > 0) & (vals < 200)]
    fig = go.Figure()
    fig.add_trace(go.Histogram(x=vals.tolist(), nbinsx=40, name='Нагрузка %',
                               marker_color=C['blu'], opacity=0.8))
    fig.add_vline(x=100, line_dash='dash', line_color=C['grn'], annotation_text='Норм. 100%')
    fig.add_vline(x=110, line_dash='dash', line_color=C['red'], annotation_text='Перегруз 110%')
    fig.update_layout(title='Распределение нагрузки по рейсам (%)', height=300, xaxis_title='%')
    return fig


def fig_payload_by_truck():
    if len(payload_df) == 0:
        return go.Figure()
    col = next((c for c in payload_df.columns
                if 'final' in str(c).lower() and 'payload' in str(c).lower()), None)
    if col is None:
        return go.Figure()
    truck_col = 'truck'
    df = payload_df[[truck_col, col]].copy()
    df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna()
    box_data = df.groupby(truck_col)[col].apply(list)
    fig = go.Figure()
    for truck, vals in box_data.items():
        fig.add_trace(go.Box(y=vals, name=f'№{truck}', boxpoints=False))
    fig.add_hline(y=1500, line_dash='dash', line_color=C['amb'], annotation_text='Норм. 1500т')
    fig.update_layout(title='FinalPayload по самосвалам (т)', height=350)
    return fig


def fig_tkph():
    if len(payload_df) == 0:
        return go.Figure()
    tkph_cols = [c for c in payload_df.columns if 'tkph' in str(c).lower()]
    if not tkph_cols:
        return go.Figure()
    fig = go.Figure()
    for col in tkph_cols[:4]:
        vals = pd.to_numeric(payload_df[col], errors='coerce').dropna()
        vals = vals[(vals > 0) & (vals < 2000)]
        fig.add_trace(go.Box(y=vals.tolist(), name=str(col).strip(), boxpoints=False))
    fig.update_layout(title='TKPH по осям (тонно-км/ч)', height=320)
    return fig


def fig_oil_consumption():
    if len(oil_df) == 0:
        return go.Figure()
    date_col = next((c for c in oil_df.columns if 'Дата' in str(c)), None)
    truck_col = next((c for c in oil_df.columns if 'хоз' in str(c).lower() or '№' in str(c)), None)
    vol_col = next((c for c in oil_df.columns if 'выдано' in str(c).lower() or 'объём' in str(c).lower() or 'л' in str(c).lower()), None)
    if not (truck_col and vol_col):
        return go.Figure()
    df = oil_df.copy()
    df['vol'] = pd.to_numeric(df[vol_col], errors='coerce')
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df = df.dropna(subset=['truck', 'vol'])
    total = df.groupby('truck')['vol'].sum().sort_values(ascending=False).head(25)
    fig = go.Figure(go.Bar(
        x=[f'№{int(t)}' for t in total.index],
        y=total.values.tolist(),
        marker_color=[C['red'] if v > total.mean()*1.5 else C['amb'] if v > total.mean() else C['grn']
                      for v in total.values],
        text=[f'{v:.0f}л' for v in total.values],
        textposition='auto',
    ))
    fig.update_layout(title='Суммарный расход масла по самосвалу (литры)', height=350)
    return fig


def fig_oil_trend():
    if len(oil_df) == 0:
        return go.Figure()
    date_col = next((c for c in oil_df.columns if 'Дата' in str(c)), None)
    vol_col = next((c for c in oil_df.columns if 'выдано' in str(c).lower() or 'л' in str(c).lower()), None)
    if not (date_col and vol_col):
        return go.Figure()
    df = oil_df.copy()
    df['date'] = pd.to_datetime(df[date_col], errors='coerce')
    df['vol'] = pd.to_numeric(df[vol_col], errors='coerce')
    df = df.dropna(subset=['date', 'vol'])
    daily = df.set_index('date').resample('D')['vol'].sum().reset_index()
    daily['cumsum'] = daily['vol'].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Bar(x=daily['date'].tolist(), y=daily['vol'].tolist(),
                         name='Доливки/день', marker_color=C['pur'], opacity=0.7))
    fig.add_trace(go.Scatter(x=daily['date'].tolist(), y=daily['cumsum'].tolist(),
                             name='Накопленный', line=dict(color=C['acc'], width=2),
                             yaxis='y2'))
    fig.update_layout(title='Динамика доливок масла', height=320,
                      yaxis=dict(title='л/день'),
                      yaxis2=dict(title='Накопленный (л)', overlaying='y', side='right'))
    return fig


def fig_oil_heatmap():
    if len(oil_df) == 0:
        return go.Figure()
    date_col = next((c for c in oil_df.columns if 'Дата' in str(c)), None)
    truck_col = next((c for c in oil_df.columns if 'хоз' in str(c).lower() or '№' in str(c)), None)
    vol_col = next((c for c in oil_df.columns if 'выдано' in str(c).lower() or 'л' in str(c).lower()), None)
    if not (date_col and truck_col and vol_col):
        return go.Figure()
    df = oil_df.copy()
    df['date'] = pd.to_datetime(df[date_col], errors='coerce')
    df['truck'] = pd.to_numeric(df[truck_col], errors='coerce')
    df['vol'] = pd.to_numeric(df[vol_col], errors='coerce')
    df = df.dropna(subset=['date', 'truck', 'vol'])
    df['month'] = df['date'].dt.strftime('%Y-%m')
    pivot = df.pivot_table(index='truck', columns='month', values='vol', aggfunc='sum').fillna(0)
    fig = go.Figure(go.Heatmap(
        z=pivot.values.tolist(),
        x=list(pivot.columns),
        y=[f'№{int(t)}' for t in pivot.index],
        colorscale=[[0, C['bg2']], [0.5, C['amb']], [1, C['red']]],
        colorbar=dict(title='л', tickfont_color=C['tx3']),
    ))
    fig.update_layout(title='Доливки масла (самосвал × месяц)', height=max(250, len(pivot)*28+80))
    return fig


def fig_production():
    if len(prod_df) == 0:
        return go.Figure()
    df = prod_df.copy()
    # Detect key columns
    year_col = next((c for c in df.columns if 'год' in str(c).lower()), None)
    month_col = next((c for c in df.columns if 'месяц' in str(c).lower()), None)
    ton_col = next((c for c in df.columns if 'тонн' in str(c).lower()), None)
    model_col = next((c for c in df.columns if 'модель' in str(c).lower()), None)
    speed_col = next((c for c in df.columns if 'скорост' in str(c).lower()), None)
    load_col = next((c for c in df.columns if 'загрузка' in str(c).lower()), None)
    if not (ton_col and model_col):
        return go.Figure()
    df[ton_col] = pd.to_numeric(df[ton_col], errors='coerce')
    if speed_col:
        df[speed_col] = pd.to_numeric(df[speed_col], errors='coerce')
    fig = make_subplots(rows=2, cols=2,
                        subplot_titles=['Тонн/год по модели', 'Скорость груженым (км/ч)',
                                        'Средняя загрузка АС (т)', 'KPI сравнение NTE200 vs 730E'])
    # Chart 1: tonnage by model
    by_model = df.groupby(model_col)[ton_col].mean().reset_index()
    fig.add_trace(go.Bar(x=by_model[model_col].tolist(), y=by_model[ton_col].tolist(),
                         marker_color=[C['acc'], C['blu']], name='Тонн/год'), row=1, col=1)
    # Chart 2: speed by model
    if speed_col:
        sp_model = df.groupby(model_col)[speed_col].mean().reset_index()
        fig.add_trace(go.Bar(x=sp_model[model_col].tolist(), y=sp_model[speed_col].round(2).tolist(),
                             marker_color=[C['grn'], C['pur']], name='Скорость'), row=1, col=2)
    # Chart 3: load by model
    if load_col:
        df[load_col] = pd.to_numeric(df[load_col], errors='coerce')
        ld_model = df.groupby(model_col)[load_col].mean().reset_index()
        fig.add_trace(go.Bar(x=ld_model[model_col].tolist(), y=ld_model[load_col].round(1).tolist(),
                             marker_color=[C['amb'], C['red']], name='Загрузка'), row=2, col=1)
    # Chart 4: scatter speed vs tonnage
    if speed_col:
        for model_name in df[model_col].unique():
            sub = df[df[model_col] == model_name]
            fig.add_trace(go.Scatter(
                x=sub[speed_col].tolist(), y=sub[ton_col].tolist(),
                mode='markers', name=str(model_name),
                marker=dict(size=6, opacity=0.7)), row=2, col=2)
    fig.update_layout(title='Производительность парка 2024-2026', height=580, showlegend=False)
    return fig


# ─── BUILD CHARTS ─────────────────────────────────────────────────────────────
print('Building charts...')
charts = {}

# Tab 1: Fleet overview
charts['fleet_health'] = _fig_json(fig_fleet_health())
charts['downtime_heatmap'] = _fig_json(fig_downtime_heatmap())
charts['pareto'] = _fig_json(fig_pareto_failures())
charts['mtbf'] = _fig_json(fig_mtbf())
charts['kpi'] = _fig_json(fig_kpi_gauges())
charts['downtime_bar'] = _fig_json(fig_downtime_bar())

# Tab 2 & 3: Engine parameters + predictive (use first available truck)
first_key = next(iter(insite)) if insite else None
second_key = list(insite.keys())[1] if len(insite) > 1 else first_key

for idx, key in enumerate(list(insite.keys())[:3]):
    df = insite[key]['df']
    label = f"({insite[key]['model']} №{insite[key]['truck']})"
    sfx = f'_{idx}'
    charts[f'rpm{sfx}'] = _fig_json(fig_engine_rpm(df, label))
    charts[f'tcool{sfx}'] = _fig_json(fig_coolant_temp(df, label))
    charts[f'oil{sfx}'] = _fig_json(fig_oil_combined(df, label))
    charts[f'exh_hm{sfx}'] = _fig_json(fig_exhaust_heatmap(df, label))
    charts[f'cyl_bal{sfx}'] = _fig_json(fig_cylinder_balance(df, label))
    charts[f'fuel{sfx}'] = _fig_json(fig_fuel_rate(df, label))
    charts[f'boost{sfx}'] = _fig_json(fig_boost_pressure(df, label))
    charts[f'crank{sfx}'] = _fig_json(fig_crankcase(df, label))
    charts[f'load{sfx}'] = _fig_json(fig_load_percent(df, label))
    charts[f'stats{sfx}'] = _fig_json(fig_stats_table(df, label))
    # Predictive
    charts[f'health{sfx}'] = _fig_json(fig_health_score_gauge(df, label))
    charts[f'anomaly{sfx}'] = _fig_json(fig_anomalies(df, label))
    charts[f'forecast{sfx}'] = _fig_json(fig_oil_pressure_forecast(df, label))
    charts[f'risk{sfx}'] = _fig_json(fig_risk_table(df, label))

# Tab 4: ECM
charts['ecm_heatmap'] = _fig_json(fig_ecm_heatmap())
charts['ecm_pareto'] = _fig_json(fig_ecm_pareto())
charts['ecm_active'] = _fig_json(fig_ecm_active())
charts['ecm_lamp'] = _fig_json(fig_ecm_lamp_donut())

# Tab 5: Maintenance
charts['maint_timeline'] = _fig_json(fig_maint_timeline())
charts['gbc_heatmap'] = _fig_json(fig_gbc_heatmap())
charts['mtbf2'] = _fig_json(fig_mtbf())

# Tab 6: Payload
charts['payload_dist'] = _fig_json(fig_payload_distribution())
charts['payload_truck'] = _fig_json(fig_payload_by_truck())
charts['tkph'] = _fig_json(fig_tkph())

# Tab 7: GE
charts['ge_events'] = _fig_json(fig_ge_events())

# Tab 8: Production
charts['production'] = _fig_json(fig_production())

# Tab 9: Oil
charts['oil_total'] = _fig_json(fig_oil_consumption())
charts['oil_trend'] = _fig_json(fig_oil_trend())
charts['oil_heatmap'] = _fig_json(fig_oil_heatmap())

print(f'Built {len(charts)} charts.')

# ─── TRUCK SELECTOR DATA ──────────────────────────────────────────────────────
truck_options = []
for key, info in insite.items():
    truck_options.append({'key': key, 'label': f"{info['model']} №{info['truck']}", 'model': info['model'], 'truck': info['truck']})

# ─── BUILD HTML ───────────────────────────────────────────────────────────────
charts_json = json.dumps(charts, ensure_ascii=False)
trucks_json = json.dumps(truck_options, ensure_ascii=False)

# Also embed per-truck chart data for JS switching
truck_charts = {}
for idx, key in enumerate(list(insite.keys())):
    df = insite[key]['df']
    label = f"({insite[key]['model']} №{insite[key]['truck']})"
    sfx = f'_{idx}'
    truck_charts[key] = {
        'rpm': charts.get(f'rpm{sfx}', '{}'),
        'tcool': charts.get(f'tcool{sfx}', '{}'),
        'oil': charts.get(f'oil{sfx}', '{}'),
        'exh_hm': charts.get(f'exh_hm{sfx}', '{}'),
        'cyl_bal': charts.get(f'cyl_bal{sfx}', '{}'),
        'fuel': charts.get(f'fuel{sfx}', '{}'),
        'boost': charts.get(f'boost{sfx}', '{}'),
        'crank': charts.get(f'crank{sfx}', '{}'),
        'load': charts.get(f'load{sfx}', '{}'),
        'stats': charts.get(f'stats{sfx}', '{}'),
        'health': charts.get(f'health{sfx}', '{}'),
        'anomaly': charts.get(f'anomaly{sfx}', '{}'),
        'forecast': charts.get(f'forecast{sfx}', '{}'),
        'risk': charts.get(f'risk{sfx}', '{}'),
    }

# Pre-render first 3 trucks for fast load; rest loaded on demand
# (all are embedded, but only first 3 are pre-built above)
# For trucks beyond idx 2, we build on-demand charts:
print('Building remaining truck charts...')
for idx, (key, info) in enumerate(list(insite.items())[3:], start=3):
    df = info['df']
    label = f"({info['model']} №{info['truck']})"
    truck_charts[key] = {
        'rpm': _fig_json(fig_engine_rpm(df, label)),
        'tcool': _fig_json(fig_coolant_temp(df, label)),
        'oil': _fig_json(fig_oil_combined(df, label)),
        'exh_hm': _fig_json(fig_exhaust_heatmap(df, label)),
        'cyl_bal': _fig_json(fig_cylinder_balance(df, label)),
        'fuel': _fig_json(fig_fuel_rate(df, label)),
        'boost': _fig_json(fig_boost_pressure(df, label)),
        'crank': _fig_json(fig_crankcase(df, label)),
        'load': _fig_json(fig_load_percent(df, label)),
        'stats': _fig_json(fig_stats_table(df, label)),
        'health': _fig_json(fig_health_score_gauge(df, label)),
        'anomaly': _fig_json(fig_anomalies(df, label)),
        'forecast': _fig_json(fig_oil_pressure_forecast(df, label)),
        'risk': _fig_json(fig_risk_table(df, label)),
    }
    if idx % 3 == 0:
        print(f'  ...{idx+1}/{len(insite)} trucks processed')

truck_charts_json = json.dumps(truck_charts, ensure_ascii=False)
print('All truck charts built.')

# ─── HTML TEMPLATE ─────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>QSK50 Predictive Analytics Dashboard — NTE200 / KOMATSU 730E</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
:root{{
  --bg:#141c24;--bg2:#1c2630;--bg3:#242f39;
  --card:#2a3138;--card2:#323941;
  --brd:rgba(255,255,255,0.08);--brd2:rgba(255,255,255,0.14);
  --tx:#e8edf6;--tx2:#b3bddc;--tx3:#80868b;
  --acc:#3ef0af;--red:#e84855;--amb:#f5a623;--grn:#22c55e;--blu:#7e83fa;--pur:#b668e4;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg);color:var(--tx);min-height:100vh;font-size:14px}}
body::before{{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(62,240,175,.018) 1px,transparent 1px),linear-gradient(90deg,rgba(62,240,175,.018) 1px,transparent 1px);background-size:40px 40px;pointer-events:none;z-index:0}}
/* TOPBAR */
.tb{{display:flex;align-items:center;gap:12px;padding:0 18px;height:52px;border-bottom:1px solid var(--brd);background:rgba(10,13,18,.96);backdrop-filter:blur(12px);position:sticky;top:0;z-index:100}}
.logo{{font-size:13px;font-weight:800;color:var(--acc)}}
.logo small{{color:var(--tx3);font-weight:400;font-size:9px;display:block}}
.sep{{width:1px;height:24px;background:var(--brd2);flex-shrink:0}}
.kpi-row{{display:flex;gap:20px;margin-left:auto}}
.kpi{{text-align:center}}
.kpi .val{{font-size:16px;font-weight:700;color:var(--acc);font-family:monospace}}
.kpi .lbl{{font-size:9px;color:var(--tx3);text-transform:uppercase;letter-spacing:.4px}}
/* TABS */
.tabs{{display:flex;gap:0;border-bottom:1px solid var(--brd);background:var(--bg2);overflow-x:auto;position:sticky;top:52px;z-index:99}}
.tab-btn{{padding:10px 16px;border:none;border-bottom:2px solid transparent;background:none;color:var(--tx3);cursor:pointer;font-size:12px;white-space:nowrap;transition:all .15s;font-weight:500}}
.tab-btn:hover{{color:var(--tx);background:rgba(255,255,255,0.04)}}
.tab-btn.act{{color:var(--acc);border-bottom-color:var(--acc);background:rgba(62,240,175,0.06)}}
/* CONTENT */
.tab-panel{{display:none;padding:20px}}
.tab-panel.show{{display:block}}
/* GRID */
.grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}}
.grid-1{{margin-bottom:16px}}
.card{{background:var(--card);border:1px solid var(--brd);border-radius:10px;padding:12px;}}
/* TRUCK SELECTOR */
.truck-sel{{display:flex;align-items:center;gap:10px;padding:12px 0 16px;flex-wrap:wrap}}
.truck-sel label{{font-size:11px;color:var(--tx3)}}
.truck-sel select{{background:var(--bg3);border:1px solid var(--brd2);border-radius:6px;padding:6px 12px;color:var(--tx);font-size:12px;outline:none;cursor:pointer}}
.truck-sel select:focus{{border-color:var(--acc)}}
/* SPINNER */
.spinner{{display:none;width:24px;height:24px;border:3px solid var(--brd);border-top-color:var(--acc);border-radius:50%;animation:spin .6s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
</style>
</head>
<body>
<div class="tb">
  <div class="logo">⚙ QSK50 Analytics<small>NTE200 / KOMATSU 730E</small></div>
  <div class="sep"></div>
  <div class="kpi-row" id="kpi-row">
    <div class="kpi"><div class="val" id="kpi-trucks">–</div><div class="lbl">Машин</div></div>
    <div class="kpi"><div class="val" id="kpi-records">–</div><div class="lbl">Записей</div></div>
    <div class="kpi"><div class="val" id="kpi-faults">–</div><div class="lbl">Кодов ЭБУ</div></div>
    <div class="kpi"><div class="val" id="kpi-repairs">–</div><div class="lbl">Ремонтов</div></div>
  </div>
</div>

<div class="tabs">
  <button class="tab-btn act" onclick="showTab('fleet',this)">📊 Обзор парка</button>
  <button class="tab-btn" onclick="showTab('engine',this)">🔩 Параметры ДВС</button>
  <button class="tab-btn" onclick="showTab('predict',this)">🔮 Предиктивный</button>
  <button class="tab-btn" onclick="showTab('ecm',this)">⚠ Коды ЭБУ</button>
  <button class="tab-btn" onclick="showTab('maint',this)">🔧 История ТО</button>
  <button class="tab-btn" onclick="showTab('payload',this)">⚖ Нагрузка</button>
  <button class="tab-btn" onclick="showTab('ge',this)">⚡ GE Привод</button>
  <button class="tab-btn" onclick="showTab('prod',this)">📈 Производит.</button>
  <button class="tab-btn" onclick="showTab('oil',this)">🛢 Масло</button>
</div>

<!-- TAB 1: Fleet Overview -->
<div id="tab-fleet" class="tab-panel show">
  <div class="grid-1 card"><div id="kpi_chart"></div></div>
  <div class="grid-2">
    <div class="card"><div id="fleet_health"></div></div>
    <div class="card"><div id="downtime_heatmap"></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="pareto"></div></div>
    <div class="card"><div id="mtbf"></div></div>
  </div>
  <div class="grid-1 card"><div id="downtime_bar"></div></div>
</div>

<!-- TAB 2: Engine Parameters -->
<div id="tab-engine" class="tab-panel">
  <div class="truck-sel">
    <label>Выберите самосвал:</label>
    <select id="eng-truck-sel" onchange="loadEngineCharts()"></select>
    <div class="spinner" id="eng-spinner"></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="eng-rpm"></div></div>
    <div class="card"><div id="eng-tcool"></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="eng-oil"></div></div>
    <div class="card"><div id="eng-fuel"></div></div>
  </div>
  <div class="grid-1 card"><div id="eng-exh-hm"></div></div>
  <div class="grid-2">
    <div class="card"><div id="eng-cyl-bal"></div></div>
    <div class="card"><div id="eng-boost"></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="eng-crank"></div></div>
    <div class="card"><div id="eng-load"></div></div>
  </div>
  <div class="grid-1 card"><div id="eng-stats"></div></div>
</div>

<!-- TAB 3: Predictive Analytics -->
<div id="tab-predict" class="tab-panel">
  <div class="truck-sel">
    <label>Выберите самосвал:</label>
    <select id="pred-truck-sel" onchange="loadPredictCharts()"></select>
    <div class="spinner" id="pred-spinner"></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="pred-health"></div></div>
    <div class="card"><div id="pred-risk"></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><div id="pred-anomaly"></div></div>
    <div class="card"><div id="pred-forecast"></div></div>
  </div>
</div>

<!-- TAB 4: ECM Faults -->
<div id="tab-ecm" class="tab-panel">
  <div class="grid-2">
    <div class="card"><div id="ecm_pareto"></div></div>
    <div class="card"><div id="ecm_lamp"></div></div>
  </div>
  <div class="grid-1 card"><div id="ecm_heatmap"></div></div>
  <div class="grid-1 card"><div id="ecm_active"></div></div>
</div>

<!-- TAB 5: Maintenance -->
<div id="tab-maint" class="tab-panel">
  <div class="grid-1 card"><div id="maint_timeline"></div></div>
  <div class="grid-2">
    <div class="card"><div id="gbc_heatmap"></div></div>
    <div class="card"><div id="mtbf2"></div></div>
  </div>
</div>

<!-- TAB 6: Payload -->
<div id="tab-payload" class="tab-panel">
  <div class="grid-2">
    <div class="card"><div id="payload_dist"></div></div>
    <div class="card"><div id="tkph"></div></div>
  </div>
  <div class="grid-1 card"><div id="payload_truck"></div></div>
</div>

<!-- TAB 7: GE Drive -->
<div id="tab-ge" class="tab-panel">
  <div class="grid-1 card"><div id="ge_events"></div></div>
</div>

<!-- TAB 8: Production -->
<div id="tab-prod" class="tab-panel">
  <div class="grid-1 card"><div id="production"></div></div>
</div>

<!-- TAB 9: Oil -->
<div id="tab-oil" class="tab-panel">
  <div class="grid-1 card"><div id="oil_total"></div></div>
  <div class="grid-2">
    <div class="card"><div id="oil_trend"></div></div>
    <div class="card"><div id="oil_heatmap"></div></div>
  </div>
</div>

<script>
// ─── EMBEDDED DATA ────────────────────────────────────────────────────────
const CHARTS = {charts_json};
const TRUCK_CHARTS = {truck_charts_json};
const TRUCKS = {trucks_json};

// ─── KPI ─────────────────────────────────────────────────────────────────
document.getElementById('kpi-trucks').textContent = TRUCKS.length;
document.getElementById('kpi-faults').textContent = {len(ecm_df)};
document.getElementById('kpi-repairs').textContent = {len(maint_df)};
const totalRec = Object.values(TRUCK_CHARTS).length > 0 ? TRUCKS.length * 1000 : 0;
document.getElementById('kpi-records').textContent = '{len(insite) * 1000}+';

// ─── PLOTLY RENDER HELPER ─────────────────────────────────────────────────
function renderChart(divId, jsonStr) {{
  if (!jsonStr || jsonStr === '{{}}') return;
  try {{
    const fig = JSON.parse(jsonStr);
    const layout = Object.assign({{
      paper_bgcolor: '#1c2630', plot_bgcolor: '#2a3138',
      font: {{color: '#e8edf6', family: 'Segoe UI,sans-serif'}},
      margin: {{l:40,r:20,t:40,b:40}},
    }}, fig.layout || {{}});
    Plotly.newPlot(divId, fig.data || [], layout, {{
      responsive: true, displayModeBar: true,
      modeBarButtonsToRemove: ['select2d','lasso2d'],
    }});
  }} catch(e) {{ console.warn('Chart error:', divId, e); }}
}}

// ─── TAB SWITCHING ────────────────────────────────────────────────────────
const rendered = new Set();
function showTab(name, btn) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('show'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('act'));
  document.getElementById('tab-' + name).classList.add('show');
  btn.classList.add('act');
  if (!rendered.has(name)) {{
    rendered.add(name);
    renderTabCharts(name);
  }}
}}

function renderTabCharts(name) {{
  if (name === 'fleet') {{
    renderChart('kpi_chart', CHARTS['kpi']);
    renderChart('fleet_health', CHARTS['fleet_health']);
    renderChart('downtime_heatmap', CHARTS['downtime_heatmap']);
    renderChart('pareto', CHARTS['pareto']);
    renderChart('mtbf', CHARTS['mtbf']);
    renderChart('downtime_bar', CHARTS['downtime_bar']);
  }} else if (name === 'engine') {{
    populateTruckSel('eng-truck-sel');
    loadEngineCharts();
  }} else if (name === 'predict') {{
    populateTruckSel('pred-truck-sel');
    loadPredictCharts();
  }} else if (name === 'ecm') {{
    renderChart('ecm_pareto', CHARTS['ecm_pareto']);
    renderChart('ecm_lamp', CHARTS['ecm_lamp']);
    renderChart('ecm_heatmap', CHARTS['ecm_heatmap']);
    renderChart('ecm_active', CHARTS['ecm_active']);
  }} else if (name === 'maint') {{
    renderChart('maint_timeline', CHARTS['maint_timeline']);
    renderChart('gbc_heatmap', CHARTS['gbc_heatmap']);
    renderChart('mtbf2', CHARTS['mtbf2']);
  }} else if (name === 'payload') {{
    renderChart('payload_dist', CHARTS['payload_dist']);
    renderChart('tkph', CHARTS['tkph']);
    renderChart('payload_truck', CHARTS['payload_truck']);
  }} else if (name === 'ge') {{
    renderChart('ge_events', CHARTS['ge_events']);
  }} else if (name === 'prod') {{
    renderChart('production', CHARTS['production']);
  }} else if (name === 'oil') {{
    renderChart('oil_total', CHARTS['oil_total']);
    renderChart('oil_trend', CHARTS['oil_trend']);
    renderChart('oil_heatmap', CHARTS['oil_heatmap']);
  }}
}}

function populateTruckSel(selId) {{
  const sel = document.getElementById(selId);
  if (sel.options.length > 0) return;
  TRUCKS.forEach(t => {{
    const opt = document.createElement('option');
    opt.value = t.key;
    opt.textContent = t.label;
    sel.appendChild(opt);
  }});
}}

function loadEngineCharts() {{
  const key = document.getElementById('eng-truck-sel').value;
  if (!key || !TRUCK_CHARTS[key]) return;
  const tc = TRUCK_CHARTS[key];
  ['rpm','tcool','oil','fuel','boost','crank','load'].forEach(k => {{
    renderChart('eng-' + k.replace('_','-'), tc[k]);
  }});
  renderChart('eng-exh-hm', tc['exh_hm']);
  renderChart('eng-cyl-bal', tc['cyl_bal']);
  renderChart('eng-stats', tc['stats']);
}}

function loadPredictCharts() {{
  const key = document.getElementById('pred-truck-sel').value;
  if (!key || !TRUCK_CHARTS[key]) return;
  const tc = TRUCK_CHARTS[key];
  renderChart('pred-health', tc['health']);
  renderChart('pred-risk', tc['risk']);
  renderChart('pred-anomaly', tc['anomaly']);
  renderChart('pred-forecast', tc['forecast']);
}}

// ─── INIT (render fleet tab on load) ─────────────────────────────────────
window.addEventListener('load', () => {{
  rendered.add('fleet');
  renderTabCharts('fleet');
}});
</script>
</body>
</html>"""

out_path = FOLDER / 'qsk50_dashboard.html'
out_path.write_text(html, encoding='utf-8')
print(f'HTML written: {out_path} ({out_path.stat().st_size/1024/1024:.1f} MB)')
