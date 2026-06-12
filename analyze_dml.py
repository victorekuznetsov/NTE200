import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyBboxPatch, Patch
from matplotlib.lines import Line2D
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Brand colors
DARK = '#293136'
ACCENT = '#3EF0AF'
LIGHT = '#F5F5F5'
RED = '#FF4444'
ORANGE = '#FF8C42'
YELLOW = '#FFD166'
BLUE = '#4EC9F0'
WHITE = '#FFFFFF'

plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.facecolor'] = '#1e2529'
plt.rcParams['figure.facecolor'] = DARK
plt.rcParams['text.color'] = WHITE
plt.rcParams['axes.labelcolor'] = WHITE
plt.rcParams['axes.edgecolor'] = '#445055'
plt.rcParams['xtick.color'] = '#aaaaaa'
plt.rcParams['ytick.color'] = '#aaaaaa'
plt.rcParams['grid.color'] = '#445055'
plt.rcParams['grid.alpha'] = 0.5

# ─────────────────────────────────────────────
# Load data
# ─────────────────────────────────────────────
df1 = pd.read_csv('DML_82_bez_porody.csv', encoding='cp1251', skiprows=26, decimal=',', on_bad_lines='skip')
df2 = pd.read_csv('DML_82_s_porodoy.csv', encoding='cp1251', skiprows=26, decimal=',', on_bad_lines='skip')

# Build timestamps
def make_time(df):
    ts = pd.to_datetime(df['Дата'] + ' ' + df['Время'], format='%d-%b-%Y %H:%M:%S.%f', errors='coerce')
    t0 = ts.iloc[0]
    return ts, (ts - t0).dt.total_seconds()

ts1, sec1 = make_time(df1)
ts2, sec2 = make_time(df2)

# ─────────────────────────────────────────────
# Column references
# ─────────────────────────────────────────────
RPM     = 'Частота вращения двигателя (об/мин)- Идентификатор0'
LOAD    = 'Относительная нагрузка (Проценты)- Идентификатор0'
AVG_EGT = 'Средняя температура отработавших газов (расчетное значение) (°C)- Идентификатор1'
EGT_L   = 'Средняя температура отработавших газов - левый ряд (расчетное значение) (°C)- Идентификатор144'
EGT_R   = 'Средняя температура отработавших газов - правый ряд (расчетное значение) (°C)- Идентификатор1'
COOLANT = 'Датчик температуры (°C)- Идентификатор0'
OIL_P   = 'Давление масла (кПа)- Идентификатор0'
BOOST   = 'Давление во впускном коллекторе (кПа)- Идентификатор0'
OIL_T   = 'Температура масла (°C)- Идентификатор0'
CRANK   = 'Давление картерных газов (кПа)- Идентификатор0'
FUEL_P  = 'Давление на входе топливного регулятора (Бар)- Идентификатор0'
RAIL_P  = 'Измеренное давление в общем топливопроводе высокого давления (Бар)- Идентификатор0'

# EGT per cylinder (ordered 1-16)
def egt_col(n):
    bank = '144' if n % 2 == 1 else '1'
    return f'Датчик температуры отработавших газов цилиндра {n} (°C)- Идентификатор{bank}'

# ─────────────────────────────────────────────
# FIG 1 — 4-panel overview time series
# ─────────────────────────────────────────────
fig, axes = plt.subplots(4, 1, figsize=(18, 14))
fig.patch.set_facecolor(DARK)
fig.suptitle('Агрегат 82 — Динамический анализ DML\n02.06.2026  |  Тест без породы vs Тест с породой',
             fontsize=16, color=WHITE, fontweight='bold', y=0.98)

ax_labels = ['а', 'б', 'в', 'г']
plots = [
    (RPM,     'Обороты двигателя (об/мин)',    ACCENT,   ORANGE,   (700, 2100), 2100),
    (LOAD,    'Относительная нагрузка (%)',     BLUE,     RED,      (0, 110),    None),
    (AVG_EGT, 'Средняя ЕГТ (°C)',              YELLOW,   RED,      (150, 600),  None),
    (OIL_P,   'Давление масла (кПа)',           '#90EE90', ORANGE,  (200, 600),  None),
]

for idx, (col, ylabel, c1, c2, ylim, hline) in enumerate(plots):
    ax = axes[idx]
    ax.set_facecolor('#1e2529')
    v1 = df1[col].rolling(5, center=True).mean()
    v2 = df2[col].rolling(5, center=True).mean()
    ax.plot(sec1, v1, color=c1, lw=1.5, label='Без породы', alpha=0.9)
    ax.plot(sec2 + sec1.max() + 30, v2, color=c2, lw=1.5, label='С породой', alpha=0.9)
    if hline:
        ax.axhline(hline, color=RED, lw=1, ls='--', alpha=0.7, label=f'Лимит {hline}')
    # separator
    ax.axvline(sec1.max() + 15, color='#555555', lw=1.5, ls=':')
    ax.set_ylabel(ylabel, color=WHITE, fontsize=9)
    ax.set_ylim(ylim)
    ax.grid(True, alpha=0.3)
    ax.tick_params(colors=WHITE, labelsize=8)
    for spine in ax.spines.values():
        spine.set_edgecolor('#445055')
    if idx < 3:
        ax.set_xticklabels([])
    else:
        ax.set_xlabel('Время (сек)', color=WHITE, fontsize=9)
    # legend inside
    ax.legend(loc='upper right', fontsize=8, framealpha=0.3,
              facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
    # label panel
    ax.text(0.01, 0.92, f'({ax_labels[idx]})', transform=ax.transAxes,
            color=ACCENT, fontsize=9, fontweight='bold')

# Add test labels on top
ax0 = axes[0]
ax0.text(sec1.max()/2, ax0.get_ylim()[1]*0.92, 'ТЕСТ 1\n(без породы, 13:36)',
         ha='center', color=ACCENT, fontsize=10, fontweight='bold', alpha=0.8)
ax0.text(sec1.max() + 30 + sec2.max()/2, ax0.get_ylim()[1]*0.92, 'ТЕСТ 2\n(с породой, 23:50)',
         ha='center', color=ORANGE, fontsize=10, fontweight='bold', alpha=0.8)

plt.tight_layout(rect=[0, 0, 1, 0.97])
plt.savefig('dml_overview.png', dpi=150, bbox_inches='tight', facecolor=DARK)
plt.close()
print('Saved: dml_overview.png')

# ─────────────────────────────────────────────
# FIG 2 — EGT per cylinder comparison (bar chart)
# ─────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(18, 8))
fig.patch.set_facecolor(DARK)
fig.suptitle('Агрегат 82 — Температура ЕГТ по цилиндрам\nСравнение: без породы vs с породой',
             fontsize=15, color=WHITE, fontweight='bold', y=0.98)

cyls = list(range(1, 17))
labels = [f'Ц{n}' for n in cyls]

# For loaded: cylinder 11 has sensor fault - filter out zeros
def get_egt_stats(df, cyl):
    col = egt_col(cyl)
    v = df[col].dropna()
    # filter sensor fault values (below 50°C)
    v_clean = v[v > 50]
    if len(v_clean) < 5:
        return np.nan, np.nan, np.nan
    return v_clean.mean(), v_clean.max(), v_clean.min()

means1 = []
maxs1 = []
means2 = []
maxs2 = []
for c in cyls:
    m1, x1, _ = get_egt_stats(df1, c)
    m2, x2, _ = get_egt_stats(df2, c)
    means1.append(m1)
    maxs1.append(x1)
    means2.append(m2)
    maxs2.append(x2)

means1 = np.array(means1)
maxs1 = np.array(maxs1)
means2 = np.array(means2)
maxs2 = np.array(maxs2)

x = np.arange(16)
w = 0.35

for ax_idx, (means, maxs, title, color_mean, color_max) in enumerate([
    (means1, maxs1, 'Тест 1: БЕЗ ПОРОДЫ\n(~21% нагрузка, 1441 об/мин)', ACCENT, YELLOW),
    (means2, maxs2, 'Тест 2: С ПОРОДОЙ\n(~88% нагрузка, 1822 об/мин)', ORANGE, RED),
]):
    ax = axes[ax_idx]
    ax.set_facecolor('#1e2529')

    bars_m = ax.bar(x - w/2, means, w, label='Среднее', color=color_mean, alpha=0.85, zorder=3)
    bars_x = ax.bar(x + w/2, maxs, w, label='Максимум', color=color_max, alpha=0.85, zorder=3)

    # Highlight anomalous cylinders
    # Cyl 11 in test 2 had sensor fault
    if ax_idx == 1:
        # Cyl 11 is index 10
        ax.bar(10 - w/2, means[10] if not np.isnan(means[10]) else 0, w,
               color='gray', alpha=0.5, hatch='////', zorder=4, label='Датчик неиспр.')
        ax.bar(10 + w/2, maxs[10] if not np.isnan(maxs[10]) else 0, w,
               color='gray', alpha=0.5, hatch='////', zorder=4)
        ax.text(10, 30, 'FC\n1531', ha='center', color=RED, fontsize=7, fontweight='bold')

    # Highlight top 3 hottest by max
    if not np.all(np.isnan(maxs)):
        m_clean = np.nan_to_num(maxs, nan=0)
        top3 = np.argsort(m_clean)[-3:]
        for ti in top3:
            ax.bar(ti + w/2, maxs[ti], w, color='#FF2222', alpha=1.0, zorder=5)
            ax.text(ti + w/2, maxs[ti] + 5, f'{maxs[ti]:.0f}', ha='center',
                    color=RED, fontsize=7.5, fontweight='bold')

    # Reference lines
    ax.axhline(500, color=RED, lw=1.5, ls='--', alpha=0.8, label='Критич. 500°C')
    ax.axhline(450, color=ORANGE, lw=1, ls=':', alpha=0.7, label='Внимание 450°C')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9, color=WHITE)
    ax.set_ylabel('Температура ЕГТ (°C)', color=WHITE, fontsize=10)
    ax.set_title(title, color=WHITE, fontsize=11, fontweight='bold', pad=8)
    ax.set_ylim(0, 620)
    ax.grid(True, axis='y', alpha=0.3)
    ax.legend(loc='upper right', fontsize=8, framealpha=0.3,
              facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
    ax.tick_params(colors=WHITE)
    for spine in ax.spines.values():
        spine.set_edgecolor('#445055')

    # Bank annotations
    ax.text(3.5, 595, 'A-БАНК (нечётные, ECM144)', ha='center', color='#aaaaaa', fontsize=8)
    ax.text(11.5, 595, 'B-БАНК (чётные, ECM1)', ha='center', color='#aaaaaa', fontsize=8)
    ax.axvline(7.5, color='#666666', lw=1, ls='-', alpha=0.5)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig('dml_egt_cylinders.png', dpi=150, bbox_inches='tight', facecolor=DARK)
plt.close()
print('Saved: dml_egt_cylinders.png')

# ─────────────────────────────────────────────
# FIG 3 — EGT time series per bank, with individual cylinders
# ─────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(18, 12))
fig.patch.set_facecolor(DARK)
fig.suptitle('Агрегат 82 — Временные ряды ЕГТ по банкам\nТест без породы (лево) | Тест с породой (право)',
             fontsize=15, color=WHITE, fontweight='bold')

# Color per cylinder within bank
bank_a = [1, 3, 5, 7, 9, 11, 13, 15]   # A-bank odd
bank_b = [2, 4, 6, 8, 10, 12, 14, 16]  # B-bank even
colors_a = plt.cm.viridis(np.linspace(0.2, 0.9, 8))
colors_b = plt.cm.plasma(np.linspace(0.2, 0.9, 8))

for row, (bank, colors, bank_name) in enumerate([(bank_a, colors_a, 'A-банк (нечётные, ECM144)'),
                                                   (bank_b, colors_b, 'B-банк (чётные, ECM1)')]):
    for col_idx, (df, sec, test_lbl) in enumerate([(df1, sec1, 'Без породы'), (df2, sec2, 'С породой')]):
        ax = axes[row][col_idx]
        ax.set_facecolor('#1e2529')

        for ci, (cyl, color) in enumerate(zip(bank, colors)):
            col_name = egt_col(cyl)
            v = df[col_name]
            # Filter sensor faults for cyl 11
            if cyl == 11 and col_idx == 1:
                v = v.where(v > 50, other=np.nan)
            v_smooth = v.rolling(5, center=True).mean()
            lw = 2.0 if cyl in [6, 11] else 1.2  # highlight problem cylinders
            ls = '--' if cyl == 11 else '-'
            ax.plot(sec, v_smooth, color=color, lw=lw, ls=ls,
                    label=f'Ц{cyl}', alpha=0.9)

        ax.axhline(500, color=RED, lw=1.5, ls='--', alpha=0.7)
        ax.set_ylim(0, 600)
        ax.set_ylabel('ЕГТ (°C)', color=WHITE, fontsize=9)
        ax.set_xlabel('Время (сек)', color=WHITE, fontsize=9)
        ax.set_title(f'{bank_name} — {test_lbl}', color=WHITE, fontsize=10, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.legend(loc='upper right', fontsize=7, ncol=2, framealpha=0.3,
                  facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
        ax.tick_params(colors=WHITE, labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor('#445055')

        # Annotate max point
        # Only for loaded test
        if col_idx == 1:
            for cyl, color in zip(bank, colors):
                col_name = egt_col(cyl)
                v = df[col_name]
                if cyl == 11:
                    v = v.where(v > 50, other=np.nan)
                if v.max() > 540:
                    idx_max = v.idxmax()
                    ax.annotate(f'Ц{cyl}\n{v.max():.0f}°C',
                               xy=(sec[idx_max], v[idx_max]),
                               xytext=(sec[idx_max]+20, v[idx_max]+15),
                               color=RED, fontsize=7.5, fontweight='bold',
                               arrowprops=dict(arrowstyle='->', color=RED, lw=1))

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig('dml_egt_timeseries.png', dpi=150, bbox_inches='tight', facecolor=DARK)
plt.close()
print('Saved: dml_egt_timeseries.png')

# ─────────────────────────────────────────────
# FIG 4 — Key parameters comparison: 6-panel
# ─────────────────────────────────────────────
fig = plt.figure(figsize=(18, 14))
fig.patch.set_facecolor(DARK)
fig.suptitle('Агрегат 82 — Сравнительный анализ рабочих параметров\n02.06.2026',
             fontsize=16, color=WHITE, fontweight='bold', y=0.99)

gs = gridspec.GridSpec(3, 2, hspace=0.45, wspace=0.3)

param_plots = [
    (RPM,    'Обороты двигателя (об/мин)',          ACCENT,  ORANGE,  (700, 2100),  [(2100, 'Лимит 2100', RED)]),
    (LOAD,   'Нагрузка (%)',                         BLUE,    RED,     (0, 110),     None),
    (AVG_EGT,'Средняя ЕГТ (°C)',                     YELLOW,  RED,     (150, 560),   [(500, '500°C крит.', RED)]),
    (OIL_P,  'Давление масла (кПа)',                 '#90EE90', ORANGE,(200, 600),   [(241, 'Мин.241', RED)]),
    (BOOST,  'Давл. наддува впускн. коллект. (кПа)', '#87CEEB', '#FF7F00', (0, 320), None),
    (OIL_T,  'Температура масла (°C)',               '#FFB6C1', '#FF4500', (60, 130), [(110, 'Лимит 110', RED)]),
]

for idx, (col, ylabel, c1, c2, ylim, hlines) in enumerate(param_plots):
    row, col_idx = idx // 2, idx % 2
    ax = fig.add_subplot(gs[row, col_idx])
    ax.set_facecolor('#1e2529')

    v1 = df1[col].rolling(5, center=True).mean()
    v2 = df2[col].rolling(5, center=True).mean()

    ax.fill_between(sec1, v1, alpha=0.15, color=c1)
    ax.fill_between(sec2, v2, alpha=0.15, color=c2)
    ax.plot(sec1, v1, color=c1, lw=1.8, label=f'Без породы (avg={df1[col].mean():.0f})')
    ax.plot(sec2, v2, color=c2, lw=1.8, label=f'С породой (avg={df2[col].mean():.0f})')

    if hlines:
        for val, lbl, color in hlines:
            ax.axhline(val, color=color, lw=1.5, ls='--', alpha=0.8, label=lbl)

    ax.set_ylabel(ylabel, color=WHITE, fontsize=9)
    ax.set_xlabel('Время (сек)', color=WHITE, fontsize=8)
    ax.set_ylim(ylim)
    ax.grid(True, alpha=0.3)
    ax.tick_params(colors=WHITE, labelsize=8)
    ax.legend(loc='best', fontsize=7.5, framealpha=0.3,
              facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
    for spine in ax.spines.values():
        spine.set_edgecolor('#445055')

    # Add max annotation for loaded test
    if df2[col].max() > ylim[1] * 0.85:
        ax.text(0.98, 0.96, f'MAX: {df2[col].max():.0f}',
                transform=ax.transAxes, ha='right', va='top',
                color=RED, fontsize=9, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#1e2529', edgecolor=RED, alpha=0.8))

plt.savefig('dml_params_compare.png', dpi=150, bbox_inches='tight', facecolor=DARK)
plt.close()
print('Saved: dml_params_compare.png')

# ─────────────────────────────────────────────
# FIG 5 — Cylinder 11 anomaly deep-dive
# ─────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.patch.set_facecolor(DARK)
fig.suptitle('Агрегат 82 — Аномалия Цилиндра 11 (A-банк)\nПредиктивный индикатор отказа клапана',
             fontsize=14, color=WHITE, fontweight='bold')

cyl11 = egt_col(11)
cyl9  = egt_col(9)
cyl13 = egt_col(13)
cyl6  = egt_col(6)

# Panel 1: Cyl 11 vs neighbors in test 1
ax = axes[0][0]
ax.set_facecolor('#1e2529')
for col_n, lbl, color in [(cyl9, 'Цил.9', BLUE), (cyl11, 'Цил.11', RED), (cyl13, 'Цил.13', ACCENT)]:
    v = df1[col_n].rolling(5, center=True).mean()
    ax.plot(sec1, v, color=color, lw=2, label=lbl, alpha=0.9)
ax.set_title('Тест 1 (без породы) — Цил.9,11,13 (A-банк)', color=WHITE, fontsize=10)
ax.set_ylabel('ЕГТ (°C)', color=WHITE, fontsize=9)
ax.set_ylim(0, 520)
ax.axhline(500, color=RED, lw=1, ls='--', alpha=0.6)
ax.legend(fontsize=9, framealpha=0.3, facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
ax.grid(True, alpha=0.3)
ax.tick_params(colors=WHITE, labelsize=8)
for spine in ax.spines.values(): spine.set_edgecolor('#445055')

# Panel 2: Cyl 11 in test 2 (showing sensor faults)
ax = axes[0][1]
ax.set_facecolor('#1e2529')
v11_raw = df2[cyl11]
v9 = df2[cyl9].rolling(5, center=True).mean()
v13 = df2[cyl13].rolling(5, center=True).mean()

# Show raw sensor data for cyl 11
ax.plot(sec2, v11_raw, color=RED, lw=1.2, alpha=0.4, label='Цил.11 (сырые)')
v11_filt = v11_raw.where(v11_raw > 50, other=np.nan)
ax.plot(sec2, v11_filt.rolling(3, center=True).mean(), color=RED, lw=2.5, label='Цил.11 (валидные)')
ax.plot(sec2, v9, color=BLUE, lw=1.5, label='Цил.9', alpha=0.8)
ax.plot(sec2, v13, color=ACCENT, lw=1.5, label='Цил.13', alpha=0.8)
# Highlight fault periods (below 50°C)
fault_mask = v11_raw < 50
for i, is_fault in enumerate(fault_mask):
    if is_fault:
        ax.axvspan(sec2.iloc[i]-0.5, sec2.iloc[i]+0.5, alpha=0.2, color=RED)
ax.set_title('Тест 2 (с породой) — Сбои датчика Цил.11', color=WHITE, fontsize=10)
ax.set_ylabel('ЕГТ (°C)', color=WHITE, fontsize=9)
ax.set_ylim(-10, 580)
ax.axhline(500, color=RED, lw=1, ls='--', alpha=0.6, label='500°C')
ax.legend(fontsize=8, framealpha=0.3, facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
ax.grid(True, alpha=0.3)
ax.tick_params(colors=WHITE, labelsize=8)
for spine in ax.spines.values(): spine.set_edgecolor('#445055')
# Annotate
n_faults = fault_mask.sum()
total_pts = len(v11_raw)
ax.text(0.02, 0.97, f'Сбоев датчика: {n_faults}/{total_pts} точек\n({n_faults/total_pts*100:.1f}% времени)',
        transform=ax.transAxes, color=RED, fontsize=9, fontweight='bold', va='top',
        bbox=dict(boxstyle='round', facecolor='#1e2529', edgecolor=RED, alpha=0.8))

# Panel 3: All A-bank cylinders, test 2, sorted by avg temp
ax = axes[1][0]
ax.set_facecolor('#1e2529')
a_bank_data = {}
for c in bank_a:
    v = df2[egt_col(c)]
    if c == 11:
        v = v.where(v > 50, other=np.nan)
    a_bank_data[c] = v.dropna().mean()

sorted_cyls = sorted(a_bank_data.items(), key=lambda x: x[1] if not np.isnan(x[1]) else 0, reverse=True)
bar_colors = [RED if v > 490 else (ORANGE if v > 460 else (YELLOW if v > 430 else ACCENT))
              for _, v in sorted_cyls]
bars = ax.bar([f'Ц{c}' for c, _ in sorted_cyls], [v for _, v in sorted_cyls],
              color=bar_colors, alpha=0.9, zorder=3)
ax.axhline(500, color=RED, lw=1.5, ls='--', alpha=0.8, label='500°C критич.')
ax.axhline(460, color=ORANGE, lw=1, ls=':', alpha=0.7, label='460°C внимание')
for bar, (cyl, val) in zip(bars, sorted_cyls):
    ax.text(bar.get_x() + bar.get_width()/2, val + 3, f'{val:.0f}',
            ha='center', color=WHITE, fontsize=8.5, fontweight='bold')
ax.set_title('A-банк (нечётные) — Средняя ЕГТ, тест с породой', color=WHITE, fontsize=10)
ax.set_ylabel('Средняя ЕГТ (°C)', color=WHITE, fontsize=9)
ax.set_ylim(0, 580)
ax.legend(fontsize=8, framealpha=0.3, facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
ax.grid(True, axis='y', alpha=0.3)
ax.tick_params(colors=WHITE, labelsize=9)
for spine in ax.spines.values(): spine.set_edgecolor('#445055')

# Panel 4: Cyl 6 and 2 (hottest B-bank)
ax = axes[1][1]
ax.set_facecolor('#1e2529')
for col_n, lbl, color in [(egt_col(2), 'Цил.2', ACCENT), (egt_col(6), 'Цил.6', RED),
                           (egt_col(8), 'Цил.8', BLUE), (egt_col(10), 'Цил.10', YELLOW)]:
    v = df2[col_n].rolling(5, center=True).mean()
    lw = 2.5 if col_n in [egt_col(6), egt_col(2)] else 1.5
    ax.plot(sec2, v, color=color, lw=lw, label=lbl, alpha=0.9)
ax.axhline(500, color=RED, lw=1, ls='--', alpha=0.6, label='500°C')
ax.axhline(560, color='#FF0000', lw=1, ls='-', alpha=0.5, label='560°C')
ax.set_title('B-банк (чётные) — Горячие цилиндры, тест с породой', color=WHITE, fontsize=10)
ax.set_ylabel('ЕГТ (°C)', color=WHITE, fontsize=9)
ax.set_ylim(300, 620)
ax.set_xlabel('Время (сек)', color=WHITE, fontsize=9)
ax.legend(fontsize=9, framealpha=0.3, facecolor=DARK, edgecolor='#445055', labelcolor=WHITE)
ax.grid(True, alpha=0.3)
ax.tick_params(colors=WHITE, labelsize=8)
for spine in ax.spines.values(): spine.set_edgecolor('#445055')
# Annotate max of cyl 6
v6 = df2[egt_col(6)]
max_idx = v6.idxmax()
ax.annotate(f'MAX: {v6.max():.0f}°C', xy=(sec2[max_idx], v6[max_idx]),
           xytext=(sec2[max_idx]-50, v6[max_idx]+20),
           color=RED, fontsize=9, fontweight='bold',
           arrowprops=dict(arrowstyle='->', color=RED, lw=1.5))

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig('dml_cyl11_anomaly.png', dpi=150, bbox_inches='tight', facecolor=DARK)
plt.close()
print('Saved: dml_cyl11_anomaly.png')

# ─────────────────────────────────────────────
# Print summary statistics
# ─────────────────────────────────────────────
print('\n' + '='*60)
print('SUMMARY STATISTICS')
print('='*60)
print(f'\nTest 1 (without ore): {len(df1)} samples, duration {sec1.max():.0f}s = {sec1.max()/60:.1f} min')
print(f'Test 2 (with ore):    {len(df2)} samples, duration {sec2.max():.0f}s = {sec2.max()/60:.1f} min')
print()

for df, label in [(df1, 'Without ore'), (df2, 'With ore')]:
    print(f'\n--- {label} ---')
    print(f'RPM:    avg={df[RPM].mean():.0f}, min={df[RPM].min():.0f}, max={df[RPM].max():.0f}')
    print(f'Load%:  avg={df[LOAD].mean():.1f}, max={df[LOAD].max():.0f}')
    print(f'Avg EGT: avg={df[AVG_EGT].mean():.0f}, max={df[AVG_EGT].max():.0f}°C')
    print(f'EGT L: avg={df[EGT_L].mean():.0f}, max={df[EGT_L].max():.0f}°C')
    print(f'EGT R: avg={df[EGT_R].mean():.0f}, max={df[EGT_R].max():.0f}°C')
    print(f'Coolant: avg={df[COOLANT].mean():.1f}, max={df[COOLANT].max():.1f}°C')
    print(f'Oil pres: avg={df[OIL_P].mean():.0f}, min={df[OIL_P].min():.0f} kPa')
    print(f'Boost: avg={df[BOOST].mean():.0f}, max={df[BOOST].max():.0f} kPa')
    print(f'Oil temp: avg={df[OIL_T].mean():.1f}, max={df[OIL_T].max():.1f}°C')

    # Hottest cylinders
    cyl_stats = {}
    for c in range(1, 17):
        v = df[egt_col(c)]
        if c == 11 and label == 'With ore':
            v = v.where(v > 50, other=np.nan)
        cyl_stats[c] = (v.mean(), v.max())
    top3_max = sorted(cyl_stats.items(), key=lambda x: x[1][1] if not np.isnan(x[1][1]) else 0, reverse=True)[:3]
    top3_avg = sorted(cyl_stats.items(), key=lambda x: x[1][0] if not np.isnan(x[1][0]) else 0, reverse=True)[:3]
    print(f'Top3 max EGT: ' + ', '.join([f'Cyl{c}={v[1]:.0f}' for c,v in top3_max]))
    print(f'Top3 avg EGT: ' + ', '.join([f'Cyl{c}={v[0]:.0f}' for c,v in top3_avg]))

print('\nAll output images saved.')
