#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add cylinder EGT bank asymmetry slides to DML_Анализ_LogFiles_NTE200.pptx
One slide per unit session (9 units = 9 slides appended after slide 12).
Style mirrors slide 39 of main presentation.
"""

import os
import tempfile
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─── Brand colors ─────────────────────────────────────────────────────────────
C_BG    = '#293136'
C_AXES  = '#1e2529'
C_GREEN = '#3EF0AF'
C_WHITE = '#FFFFFF'
C_RED   = '#FF4444'
C_ORANGE= '#FF8C42'
C_GRAY  = '#667788'
C_DARK2 = '#20272b'

def hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def rgb_color(h):
    r, g, b = hex2rgb(h)
    return RGBColor(r, g, b)

plt.rcParams.update({
    'figure.facecolor':  C_BG,
    'axes.facecolor':    C_AXES,
    'axes.edgecolor':    C_GRAY,
    'axes.labelcolor':   C_WHITE,
    'xtick.color':       C_GRAY,
    'ytick.color':       C_GRAY,
    'text.color':        C_WHITE,
    'grid.color':        C_GRAY,
    'grid.alpha':        0.3,
    'legend.facecolor':  C_AXES,
    'legend.edgecolor':  C_GRAY,
    'font.family':       'DejaVu Sans',
    'font.size':         9,
})

LOG_DIR = '/home/user/NTE200/log_files_nte200'
PPTX_PATH = '/home/user/NTE200/DML_Анализ_LogFiles_NTE200.pptx'

FILE_CFG = {
    '43':  (f'{LOG_DIR}/NTE200 №43 16.05.2026.csv',  'cp1251', 'en', '16.05.2026'),
    '45':  (f'{LOG_DIR}/NTE200 №45 18.05.2026.csv',  'cp1251', 'ru', '18.05.2026'),
    '50':  (f'{LOG_DIR}/NTE200 №50 19.05.2026.csv',  'cp1251', 'ru', '19.05.2026'),
    '52':  (f'{LOG_DIR}/NTE200 №52 18.05.2026.csv',  'cp1251', 'ru', '18.05.2026'),
    '56':  (f'{LOG_DIR}/NTE200 №56 19.05.2026.csv',  'cp1251', 'ru', '19.05.2026'),
    '84a': (f'{LOG_DIR}/NTE200 №84 16.05.2026.csv',  'cp1251', 'en', '16.05.2026'),
    '84b': (f'{LOG_DIR}/NTE200 №84 21.05.2026.csv',  'utf-8',  'ru', '21.05.2026'),
    '85':  (f'{LOG_DIR}/NTE200 №85 19.05.2026.csv',  'cp1251', 'ru', '19.05.2026'),
    '87':  (f'{LOG_DIR}/NTE200 №87 21.05.2026.csv',  'utf-8',  'ru', '21.05.2026'),
}

UNIT_LABELS = {
    '43':  '#43 (клапан)',
    '45':  '#45',
    '50':  '#50',
    '52':  '#52',
    '56':  '#56',
    '84a': '#84 (16.05)',
    '84b': '#84 (21.05)',
    '85':  '#85',
    '87':  '#87',
}

RISK_COLORS = {
    'КРИТИЧНО': C_RED,
    'ВНИМАНИЕ':  C_ORANGE,
    'НОРМА':     C_GREEN,
}

# ─── Data loading ─────────────────────────────────────────────────────────────

def parse_time_seconds(series):
    def to_sec(s):
        try:
            parts = str(s).split(':')
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
        except Exception:
            return np.nan
    t = series.apply(to_sec)
    return (t - t.dropna().iloc[0]).values

def get_col(df, *keywords, exclude=None):
    for col in df.columns:
        cl = col.lower()
        if all(k.lower() in cl for k in keywords):
            if exclude and any(e.lower() in cl for e in exclude):
                continue
            return col
    return None

def to_float_arr(df, col):
    if col is None:
        return None
    s = df[col]
    if s.dtype == object:
        s = s.astype(str).str.replace(',', '.').str.strip()
        s = pd.to_numeric(s, errors='coerce')
    return s.values

def load_unit_cyls(uid):
    """Return (sec_array, cyl_egt_dict, date_str, lang) for a unit."""
    path, enc, lang, date = FILE_CFG[uid]
    df = pd.read_csv(path, skiprows=26, encoding=enc, sep=',',
                     decimal=',', low_memory=False, on_bad_lines='skip')

    time_col = 'Time' if lang == 'en' else 'Время'
    if time_col not in df.columns:
        time_col = df.columns[1]
    sec = parse_time_seconds(df[time_col])

    cyl_egt = {}
    for n in range(1, 17):
        if lang == 'en':
            suffix = 'Identifier144' if n % 2 == 1 else 'Identifier1'
            col = get_col(df, f'Exhaust Temperature Sensor Cylinder {n} (', suffix)
        else:
            suffix = 'Идентификатор144' if n % 2 == 1 else 'Идентификатор1'
            col = get_col(df, f'цилиндра {n} (°C)', suffix)
        if col:
            vals = to_float_arr(df, col)
            if vals is not None:
                vals = np.array(vals, dtype=float)
                vals = np.where(vals < 50, np.nan, vals)  # filter fault readings
                cyl_egt[n] = vals

    return sec, cyl_egt, date

# ─── Chart generation ─────────────────────────────────────────────────────────

def make_bank_chart(uid, sec, cyl_egt, date_str):
    """Create A/B bank asymmetry chart, return temp file path."""
    bank_a = [1, 3, 5, 7, 9, 11, 13, 15]
    bank_b = [2, 4, 6, 8, 10, 12, 14, 16]
    colors_a = plt.cm.YlOrRd(np.linspace(0.35, 0.95, 8))
    colors_b = plt.cm.cool(np.linspace(0.25, 0.85, 8))

    fig, axes = plt.subplots(1, 2, figsize=(16, 8))
    fig.patch.set_facecolor(C_BG)

    unit_label = UNIT_LABELS[uid]
    fig.suptitle(f'Агрегат {unit_label}  |  {date_str}  |  Температура цилиндров по банкам',
                 fontsize=12, color=C_WHITE, fontweight='bold', y=0.99)

    bank_stats = {}  # bank_name -> (avg, max)

    for ax_idx, (bank, colors, bank_name, ecm_id) in enumerate([
        (bank_a, colors_a, 'А-банк (нечётные цилиндры)', 'ECM144'),
        (bank_b, colors_b, 'В-банк (чётные цилиндры)', 'ECM1'),
    ]):
        ax = axes[ax_idx]
        ax.set_facecolor(C_AXES)

        for cyl, color in zip(bank, colors):
            if cyl not in cyl_egt:
                continue
            v = cyl_egt[cyl]
            # smooth with rolling window
            s = pd.Series(v)
            v_smooth = s.rolling(5, center=True, min_periods=1).mean().values
            lw = 2.2 if cyl in [1, 2, 6, 11, 15, 16] else 1.4
            ax.plot(sec, v_smooth, color=color, lw=lw,
                    label=f'Ц{cyl}', alpha=0.9)

        ax.axhline(500, color='#FF4444', lw=1.5, ls='--', alpha=0.85, label='500°C')
        ax.axhline(550, color='#FF0000', lw=1.0, ls='-',  alpha=0.55, label='550°C')
        ax.set_ylim(200, 620)
        ax.set_xlim(0, sec[-1] if len(sec) > 0 else 1200)
        ax.set_title(f'{bank_name}\n{ecm_id}',
                     color=C_WHITE, fontsize=11, fontweight='bold')
        ax.set_ylabel('ЕГТ (°C)', color=C_WHITE, fontsize=9)
        ax.set_xlabel('Время (сек)', color=C_WHITE, fontsize=9)
        ax.legend(fontsize=7.5, ncol=2, loc='upper left',
                  framealpha=0.35, facecolor=C_AXES,
                  edgecolor='#445055', labelcolor=C_WHITE)
        ax.grid(True, alpha=0.3)
        ax.tick_params(colors=C_WHITE, labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor('#445055')

        # Stats box
        avgs, maxs = [], []
        for c in bank:
            if c in cyl_egt:
                vals = cyl_egt[c]
                finite = vals[np.isfinite(vals)]
                if len(finite) > 0:
                    avgs.append(np.mean(finite))
                    maxs.append(np.max(finite))
        if avgs:
            b_avg = np.mean(avgs)
            b_max = np.max(maxs)
        else:
            b_avg = b_max = 0.0

        box_color = '#FF4444' if b_max >= 550 else ('#FF8C42' if b_max >= 500 else '#3EF0AF')
        ax.text(0.98, 0.04,
                f'Avg банка: {b_avg:.0f}°C\nМакс банка: {b_max:.0f}°C',
                transform=ax.transAxes, ha='right', va='bottom',
                color=C_WHITE, fontsize=10, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.5', facecolor=C_AXES,
                          edgecolor=box_color, alpha=0.95, linewidth=1.5))
        bank_stats[bank_name] = (b_avg, b_max)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    fd, tmppath = tempfile.mkstemp(suffix='.png')
    os.close(fd)
    fig.savefig(tmppath, dpi=150, bbox_inches='tight',
                facecolor=C_BG, edgecolor='none')
    plt.close(fig)
    return tmppath, bank_stats

# ─── PPTX helpers ─────────────────────────────────────────────────────────────

def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def set_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb_color(C_BG)

def add_rect(slide, x, y, w, h, fill_hex=None, line_hex=None, lw=0):
    shape = slide.shapes.add_shape(1, Cm(x), Cm(y), Cm(w), Cm(h))
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb_color(fill_hex)
    else:
        shape.fill.background()
    if line_hex and lw > 0:
        shape.line.color.rgb = rgb_color(line_hex)
        shape.line.width = Pt(lw)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h,
             font_size=12, bold=False, color=C_WHITE,
             align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb_color(color)
    return tb

# ─── Conclusion text generator ────────────────────────────────────────────────

def build_conclusion(uid, bank_stats):
    a_name = 'А-банк (нечётные цилиндры)'
    b_name = 'В-банк (чётные цилиндры)'
    a_avg, a_max = bank_stats.get(a_name, (0, 0))
    b_avg, b_max = bank_stats.get(b_name, (0, 0))

    delta = b_avg - a_avg
    hotter = 'В-банк горячее А-банка' if delta > 5 else (
             'А-банк горячее В-банка' if delta < -5 else
             'Банки симметричны')

    parts = [f'ВЫВОД: {hotter}']
    if abs(delta) > 5:
        parts.append(f'(Δ={abs(delta):.0f}°C).')
    else:
        parts.append('.')

    # Find hottest cylinder globally
    all_avgs = {}
    for n in range(1, 17):
        # bank_stats only has per-bank aggregates - we need raw data
        pass  # will be overridden below with raw_hottest

    if a_max >= 550 or b_max >= 550:
        hot_bank = 'В-банк' if b_max >= a_max else 'А-банк'
        hot_max = max(a_max, b_max)
        parts.append(f' {hot_bank}: макс. ЕГТ {hot_max:.0f}°C — КРИТИЧНО (>550°C).')
    elif a_max >= 500 or b_max >= 500:
        hot_bank = 'В-банк' if b_max >= a_max else 'А-банк'
        hot_max = max(a_max, b_max)
        parts.append(f' {hot_bank}: макс. ЕГТ {hot_max:.0f}°C — ВНИМАНИЕ (>500°C).')
    else:
        parts.append(f' Макс. ЕГТ: А-банк {a_max:.0f}°C, В-банк {b_max:.0f}°C — в норме.')

    return ''.join(parts)

def build_conclusion_v2(uid, cyl_egt, bank_stats):
    """Build conclusion with per-cylinder hottest info."""
    a_name = 'А-банк (нечётные цилиндры)'
    b_name = 'В-банк (чётные цилиндры)'
    a_avg, a_max = bank_stats.get(a_name, (0, 0))
    b_avg, b_max = bank_stats.get(b_name, (0, 0))
    delta = b_avg - a_avg

    # Find hottest cylinders
    cyl_maxes = {}
    for n, vals in cyl_egt.items():
        finite = vals[np.isfinite(vals)]
        if len(finite) > 0:
            cyl_maxes[n] = np.max(finite)

    if cyl_maxes:
        hot_cyl = max(cyl_maxes, key=cyl_maxes.get)
        hot_val = cyl_maxes[hot_cyl]
    else:
        hot_cyl, hot_val = 0, 0

    # Asymmetry description
    if abs(delta) > 5:
        hotter = 'В-банк горячее А-банка' if delta > 0 else 'А-банк горячее В-банка'
        asym = f'{hotter} на {abs(delta):.0f}°C в среднем.'
    else:
        asym = 'Банки температурно симметричны.'

    # Status
    if hot_val >= 550:
        status = f' Цил.{hot_cyl} достигает {hot_val:.0f}°C — КРИТИЧНО.'
    elif hot_val >= 500:
        status = f' Цил.{hot_cyl} достигает {hot_val:.0f}°C — ВНИМАНИЕ.'
    else:
        status = f' Максимальная ЕГТ цил.{hot_cyl}: {hot_val:.0f}°C.'

    # Sensor fault note for uid 43 (known valve issue)
    extra = ''
    if uid == '43':
        extra = ' Агрегат с известным дефектом клапана — сравнение с нормой.'

    return f'ВЫВОД: {asym}{status}{extra}'

# ─── Slide builder ────────────────────────────────────────────────────────────

def add_bank_slide(prs, uid, sec, cyl_egt, date_str):
    """Append one bank asymmetry slide to prs."""
    unit_label = UNIT_LABELS[uid]

    # Build chart image
    tmppath, bank_stats = make_bank_chart(uid, sec, cyl_egt, date_str)

    conclusion = build_conclusion_v2(uid, cyl_egt, bank_stats)

    slide = blank_slide(prs)
    set_bg(slide)

    # Header bar
    add_rect(slide, 0, 0, 33.86, 2.0, fill_hex=C_DARK2)
    add_rect(slide, 0, 1.85, 33.86, 0.12, fill_hex=C_GREEN)

    add_text(slide, f'ТЕМПЕРАТУРА ЦИЛИНДРОВ — АГРЕГАТ {unit_label}  |  {date_str}',
             0.5, 0.1, 32, 1.1,
             font_size=16, bold=True, color=C_WHITE)
    add_text(slide, 'А-банк (нечётные, ECM144) — В-банк (чётные, ECM1)  |  Индивидуальные ЕГТ по цилиндрам',
             0.5, 1.2, 32, 0.65,
             font_size=9.5, color=C_GRAY)

    # Chart image (occupies the main area)
    slide.shapes.add_picture(tmppath, Cm(0.3), Cm(2.1), Cm(33.2), Cm(15.5))

    # Bottom conclusion bar
    add_rect(slide, 0, 17.7, 33.86, 1.35, fill_hex=C_DARK2,
             line_hex=C_ORANGE, lw=1.2)
    add_text(slide, conclusion,
             0.5, 17.75, 32.5, 1.25,
             font_size=10, bold=True, color=C_ORANGE)

    os.remove(tmppath)
    print(f"  Slide added: Unit {uid}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"Opening {PPTX_PATH}...")
    prs = Presentation(PPTX_PATH)
    print(f"  Current slides: {len(prs.slides)}")

    # Add section divider slide
    divider = blank_slide(prs)
    set_bg(divider)
    add_rect(divider, 0, 6.5, 33.86, 6.0, fill_hex=C_DARK2)
    add_rect(divider, 0, 6.5, 33.86, 0.3, fill_hex=C_GREEN)
    add_text(divider, 'ТЕМПЕРАТУРА', 7, 7.3, 25, 2.5,
             font_size=44, bold=True, color=C_WHITE)
    add_text(divider, 'ПО ЦИЛИНДРАМ', 7, 9.6, 25, 2.0,
             font_size=44, bold=True, color=C_GREEN)
    add_text(divider, 'А-банк (ECM144) vs В-банк (ECM1)  |  Все 9 агрегатов  |  16–21 мая 2026',
             7, 12.0, 25, 1.0,
             font_size=13, color=C_GRAY)
    add_text(divider, '09', 1.5, 7.5, 5, 3,
             font_size=72, bold=True, color=C_GREEN)

    print("  Section divider slide added.")

    # Add one slide per unit
    for uid in ['43', '45', '50', '52', '56', '84a', '84b', '85', '87']:
        print(f"  Processing unit {uid}...")
        try:
            sec, cyl_egt, date_str = load_unit_cyls(uid)
            if len(cyl_egt) == 0:
                print(f"    WARNING: No cylinder EGT data for unit {uid}")
                continue
            add_bank_slide(prs, uid, sec, cyl_egt, date_str)
        except Exception as e:
            print(f"    ERROR processing unit {uid}: {e}")
            import traceback
            traceback.print_exc()

    print(f"\nTotal slides: {len(prs.slides)}")
    prs.save(PPTX_PATH)
    print(f"Saved: {PPTX_PATH}")


if __name__ == '__main__':
    main()
