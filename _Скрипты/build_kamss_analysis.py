#!/usr/bin/env python3
"""
Build calibration analysis PDF using real INSITE data from Камсс.7z.
Sources: GIF screenshots + ECM Image Analyzer Excel files.
"""
import io, os, shutil
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
from matplotlib import gridspec
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image
import numpy as np

ZH_FONT = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'

def fp(size, bold=False):
    p = fm.FontProperties(fname=ZH_FONT, weight='bold' if bold else 'normal')
    p.set_size(size)
    return p

C_DARK  = '#293136'
C_DARK2 = '#1e2529'
C_ACC   = '#3EF0AF'
C_RED   = '#E05252'
C_ORG   = '#E08C2E'
C_YLW   = '#D4C94A'
C_GRY   = '#707880'
C_WHT   = '#FFFFFF'
C_BLU   = '#4EC9F0'
C_GRN   = '#90EE90'
C_GRN2  = '#2d5a3d'

GIF_DIR = '/home/user/NTE200/kamss_data/'

def page_bg(fig): fig.patch.set_facecolor(C_DARK)

def add_header(ax, title, subtitle=''):
    ax.set_xlim(0,1); ax.set_ylim(0,1); ax.axis('off')
    ax.add_patch(mpatches.Rectangle((0,0.82),1,0.18,facecolor=C_DARK2,
                 transform=ax.transAxes,clip_on=False))
    ax.add_patch(mpatches.Rectangle((0,0.82),0.008,0.18,facecolor=C_ACC,
                 transform=ax.transAxes,clip_on=False))
    ax.text(0.015,0.915,title,color=C_WHT,fontproperties=fp(15,bold=True),
            va='center',transform=ax.transAxes)
    if subtitle:
        ax.text(0.015,0.845,subtitle,color=C_GRY,fontproperties=fp(8.5),
                va='center',transform=ax.transAxes)

def footer(ax, txt='АО Развитие  ·  QSK50 MCRS  ·  Камсс данные INSITE  ·  2026'):
    ax.text(0.5,0.01,txt,color=C_GRY,fontproperties=fp(7),
            ha='center',va='bottom',transform=ax.transAxes)

def embed_gif(ax, fname, title=''):
    path = GIF_DIR + fname
    img = Image.open(path)
    ax.imshow(np.array(img))
    ax.axis('off')
    if title:
        ax.set_title(title, color=C_WHT, fontproperties=fp(9,bold=True), pad=3)


# ──────────────────────────────────────────────────────────────────────────────
def page_cover(pdf):
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax = fig.add_axes([0,0,1,1])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.add_patch(mpatches.Rectangle((0,0),0.015,1,facecolor=C_ACC,
                 transform=ax.transAxes,clip_on=False))
    ax.text(0.04,0.74,'Анализ калибровок ECM\nпо данным INSITE КАМСС',
            color=C_WHT,fontproperties=fp(28,bold=True),va='center',linespacing=1.4)
    ax.text(0.04,0.58,'QSK50 MCRS — сравнение 730E ш18 (AQ60217.30) vs NTE200 ш85 (AQ60809.08)',
            color=C_ACC,fontproperties=fp(14))
    ax.add_patch(mpatches.Rectangle((0.04,0.545),0.5,0.002,facecolor=C_ACC,
                 transform=ax.transAxes,clip_on=False))
    lines = [
        'Источник: INSITE ECM образы + Calterm Calibration Compare Report',
        'Снятие данных: 13–20 мая 2026 г. (Полюс Магадан)',
        '730E ш18 = 730Е№28  |  наработка ≈ 35 127 моточасов (ЭСМ)',
        'NTE200 ш85           |  наработка ≈  4 015 моточасов (ЭСМ)',
    ]
    for i,l in enumerate(lines):
        ax.text(0.04, 0.50-i*0.052, l, color=C_GRY, fontproperties=fp(10))
    ax.add_patch(mpatches.Rectangle((0,0),1,0.06,facecolor=C_DARK2,
                 transform=ax.transAxes,clip_on=False))
    ax.text(0.5,0.03,'АО Развитие  ·  QSK50 MCRS  ·  Полюс Магадан  ·  2026',
            color=C_GRY,fontproperties=fp(8),ha='center',va='center')
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_calterm_compare(pdf):
    """Show the Calterm calibration compare report screenshot."""
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Calterm Calibration Compare Report — ключевые идентификаторы',
               'Прямое сравнение ECM-образов: 730E AQ60217.30 vs NTE200 AQ60809.08')
    ax = fig.add_axes([0.01,0.05,0.98,0.82])
    embed_gif(ax, '730E ш18 DC model  из сравнения калибровок.GIF')
    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_duty_cycles(pdf):
    """Side-by-side duty cycle monitors."""
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Монитор рабочего цикла — сравнение распределения нагрузки',
               'Долгосрочный план маршрута: распределение крутящего момента по оборотам')
    ax1 = fig.add_axes([0.01,0.05,0.48,0.82])
    ax2 = fig.add_axes([0.51,0.05,0.48,0.82])
    embed_gif(ax1,'730E ш18 DC model 33223470 AQ60217.28 монитор рабочего цикла.GIF',
              '730E ш18 (AQ60217.28)  |  35 127 м/ч')
    embed_gif(ax2,'NTE200 ш85 model монитор рабочего цикла.GIF',
              'NTE200 ш85 (AQ60809.08)  |  4 015 м/ч')
    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_fuel_consumption(pdf):
    """Fuel consumption monitors."""
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Монитор расхода топлива',
               '730E: долгосрочный расход 139.9 л/ч  |  NTE200: краткосрочные данные недоступны')
    ax1 = fig.add_axes([0.01,0.05,0.48,0.82])
    ax2 = fig.add_axes([0.51,0.05,0.48,0.82])
    embed_gif(ax1,'730E ш18 DC model 33223470 AQ60217.28 расход топлива.GIF',
              '730E ш18  |  Долгосрочный расход = 139.9 л/ч')
    embed_gif(ax2,'NTE200 ш85 model расход топлива.GIF',
              'NTE200 ш85  |  Краткосрочные данные недоступны')
    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_torque_reduction(pdf, gif_730, gif_nte, title, subtitle=''):
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h, title, subtitle)
    ax1 = fig.add_axes([0.01,0.05,0.48,0.82])
    ax2 = fig.add_axes([0.51,0.05,0.48,0.82])
    embed_gif(ax1, gif_730, '730E ш18')
    embed_gif(ax2, gif_nte, 'NTE200 ш85')
    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_differences_summary(pdf):
    """Bar chart and table of confirmed calibration differences."""
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Подтверждённые различия калибровок (из INSITE / Calterm)',
               'Данные КАМСС 13–20 мая 2026 г. — реальные значения из ECM')

    # Differences table
    diffs = [
        # param, 730E val, NTE200 val, unit, risk, note
        ('Fuel Code\n(карта впрыска)', 'FC0BU52', 'FC0WH95', '—', 'WARN',
         'Разные карты CR-давления, таймингов,\nограничителей дыма'),
        ('ECM-код прошивки', 'AQ60217.30', 'AQ60809.08', '—', 'NOTE',
         'Разные версии прошивки ECM —\nразные базовые настройки'),
        ('DO_Option\n(конфиг. цифр. выходов)', '6765', '61042', '—', 'CRITICAL',
         'Разная конфигурация цифровых выходов\n— часть защитных сигналов отключена'),
        ('SC_Option\n(конфиг. системы)', '60621', '61878', '—', 'HIGH',
         'Разные опции системы управления'),
        ('Производитель ОЕМ', 'KOMATSU', 'NHL', '—', 'NOTE',
         'Машины разных производителей\n— разная инженерная документация'),
        ('Порог FC556\n(давл. картерных газов)', '10', '5', 'кПа', 'HIGH',
         'NTE200 триггерит защиту раньше\n(пороговое давл. 2× меньше)'),
        ('Заброс оборотов\n(защитный порог)', '2250', '2020', 'RPM', 'HIGH',
         'Порог защиты от разноса:\nNTE200 срабатывает на 230 RPM ниже'),
        ('Наличие воды\nв топливе FC4642', 'Нет\nканала', 'Вкл.\n1020с', 'с', 'WARN',
         'NTE200 имеет защиту по воде в топливе\nс задержкой 17 мин до снижения оборотов'),
    ]

    RISK_C = {'CRITICAL': C_RED, 'HIGH': C_ORG, 'WARN': C_YLW, 'NOTE': C_GRY}
    RISK_L = {'CRITICAL': 'КРИТИЧНО', 'HIGH': 'ВЫСОКИЙ', 'WARN': 'ПРЕДУПРЕЖД.', 'NOTE': 'ПРИМЕЧ.'}

    ax = fig.add_axes([0.01,0.05,0.98,0.82])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0,1); ax.set_ylim(0,1)

    COL_X = [0.00, 0.155, 0.285, 0.385, 0.445, 0.478, 0.53]
    COL_W = [0.153, 0.128, 0.098, 0.058, 0.031, 0.05,  0.468]
    HDRS  = ['Параметр', '730E ш18', 'NTE200 ш85', 'Ед.', '', 'Риск', 'Комментарий']

    n = len(diffs)
    ROW_H = 0.875 / (n + 1)
    y0 = 0.97

    def cell(y, cx, cw, text, bg, tc, fs=7.5, bold=False, center=False):
        ax.add_patch(mpatches.FancyBboxPatch(
            (cx+0.002, y-ROW_H+0.006), cw-0.004, ROW_H-0.01,
            boxstyle='round,pad=0.003', linewidth=0, facecolor=bg,
            transform=ax.transAxes, clip_on=False))
        ha = 'center' if center else 'left'
        xp = cx+cw/2 if center else cx+0.005
        ax.text(xp, y-ROW_H/2, text, color=tc,
                fontproperties=fp(fs, bold=bold),
                ha=ha, va='center', transform=ax.transAxes,
                clip_on=False, multialignment='left' if not center else 'center')

    for (cx, cw), h in zip(zip(COL_X, COL_W), HDRS):  # header
        cell(y0, cx, cw, h, C_DARK2, C_ACC, fs=8, bold=True, center=(h in ('Ед.', '')))
    y = y0 - ROW_H

    for idx, (param, v730, vnte, unit, risk, note) in enumerate(diffs):
        bg = '#222930' if idx % 2 == 0 else '#1e2529'
        rc = RISK_C[risk]
        rl = RISK_L[risk]
        nte_c = '#FF9999' if risk in ('CRITICAL', 'HIGH') else (C_GRN if risk == 'WARN' else C_WHT)
        tcs  = [C_WHT, C_GRN, nte_c, C_GRY, C_WHT, C_WHT, '#C0C8D0']
        bgs_row = [bg, bg, bg, bg, bg, rc, bg]
        cens = [False, True, True, True, True, True, False]
        row_vals = [param, v730, vnte, unit, rl, rl, note]
        for (cx, cw), v, tc, bgi, ce in zip(zip(COL_X, COL_W), row_vals, tcs, bgs_row, cens):
            cell(y, cx, cw, v, bgi, tc, fs=7.5 if ce else 7.8, center=ce)
        y -= ROW_H

    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_bar_comparison(pdf):
    """Bar charts of numeric protection thresholds."""
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Числовые пороги защиты: 730E vs NTE200 (по данным INSITE)',
               'Настройки "Снижение крутящего момента" и "Снижение частоты вращения"')

    # Identical thresholds (both machines same)
    same = [
        ('FC146\nТемп. ОЖ\n(мом.↓)', 100, 100, '°C'),
        ('FC235\nУровень ОЖ\n(мом.↓)', 30, 30, 'с задержки'),
        ('FC421\nТемп. масла\n(мом.↓)', 121.1, 121.1, '°C'),
        ('FC0228\nДавл. ОЖ\n(мом.↓)', 0, 0, 'кПа'),
        ('FC0155-0165\nТемп. воздуха\nвп.кол.(мом.↓)', 93.3, 93.3, '°C'),
        ('FC0151\nТемп. ОЖ\n(об/мин↓)', 110, 110, '°C'),
    ]
    # Different thresholds
    diff = [
        ('FC556\nКартерные газы\n(мом.↓)', 10, 5, 'кПа'),
        ('Заброс\nоборотов\n(об/мин↓)', 2250, 2020, 'RPM'),
    ]

    gs = gridspec.GridSpec(2, 4+2, figure=fig,
                           hspace=0.65, wspace=0.45,
                           left=0.04, right=0.98, top=0.85, bottom=0.08)

    # Top row: identical
    for i,(name,v730,vnte,unit) in enumerate(same):
        ax = fig.add_subplot(gs[0, i % 4 if i < 4 else i-4])
        if i == 4: ax = fig.add_subplot(gs[1, 0])
        if i == 5: ax = fig.add_subplot(gs[1, 1])
        ax.set_facecolor('#1e2529')
        top = max(v730,vnte)*1.3 or 1
        bars = ax.bar(['730E','NTE200'],[v730,vnte],color=[C_BLU,'#A0FFC8'],width=0.5,zorder=3)
        for bar,val in zip(bars,[v730,vnte]):
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+top*0.05,
                    str(val), ha='center', va='bottom', color=C_WHT,
                    fontproperties=fp(8,bold=True))
        ax.set_title(name, color=C_GRY, fontproperties=fp(7,bold=True), pad=3)
        ax.set_ylabel(unit, color=C_GRY, fontproperties=fp(6.5))
        ax.set_ylim(0,top)
        ax.tick_params(axis='x',colors=C_WHT,labelsize=7)
        ax.tick_params(axis='y',colors=C_GRY,labelsize=6)
        ax.set_xticklabels(['730E','NTE200'],fontproperties=fp(7))
        ax.grid(axis='y',color=C_GRY,alpha=0.18,zorder=0)
        for sp in ax.spines.values(): sp.set_edgecolor(C_GRY); sp.set_linewidth(0.4)
        # Same = green border
        for sp in ax.spines.values(): sp.set_edgecolor('#2d5a3d')

    # Different ones — right side, both rows
    for i,(name,v730,vnte,unit) in enumerate(diff):
        ax = fig.add_subplot(gs[i, 4:])
        ax.set_facecolor('#2a1a1a')
        top = max(v730,vnte)*1.3
        clr_nte = C_RED if vnte < v730 else C_ORG
        bars = ax.bar(['730E','NTE200'],[v730,vnte],color=[C_BLU,clr_nte],width=0.5,zorder=3)
        for bar,val in zip(bars,[v730,vnte]):
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+top*0.04,
                    str(val), ha='center', va='bottom', color=C_WHT,
                    fontproperties=fp(9,bold=True))
        ax.set_title(f'⚠  {name}  — ОТЛИЧАЕТСЯ', color=C_ORG,
                     fontproperties=fp(9,bold=True), pad=4)
        ax.set_ylabel(unit,color=C_GRY,fontproperties=fp(7.5))
        ax.set_ylim(0,top)
        ax.tick_params(axis='x',colors=C_WHT,labelsize=8)
        ax.tick_params(axis='y',colors=C_GRY,labelsize=7)
        ax.set_xticklabels(['730E','NTE200'],fontproperties=fp(8))
        ax.grid(axis='y',color=C_GRY,alpha=0.2,zorder=0)
        for sp in ax.spines.values(): sp.set_edgecolor(C_ORG); sp.set_linewidth(1.2)

    # Legend
    ax_leg = fig.add_axes([0.04, 0.02, 0.92, 0.04])
    ax_leg.set_facecolor(C_DARK); ax_leg.axis('off')
    ax_leg.add_patch(mpatches.Rectangle((0.0,0.1),0.02,0.8,facecolor='#2d5a3d',
                     transform=ax_leg.transAxes,clip_on=False))
    ax_leg.text(0.03,0.5,'= Одинаково у обеих машин',color=C_GRY,
                fontproperties=fp(8),va='center',transform=ax_leg.transAxes)
    ax_leg.add_patch(mpatches.Rectangle((0.3,0.1),0.02,0.8,facecolor=C_ORG,
                     transform=ax_leg.transAxes,clip_on=False))
    ax_leg.text(0.33,0.5,'= Различается (данные INSITE КАМСС)',color=C_ORG,
                fontproperties=fp(8),va='center',transform=ax_leg.transAxes)

    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_rpm_reduction(pdf):
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Настройки снижения частоты вращения (об/мин)',
               'Реальные данные INSITE — Снижение частоты вращения двигателя')
    ax1 = fig.add_axes([0.01,0.05,0.48,0.82])
    ax2 = fig.add_axes([0.51,0.05,0.48,0.82])
    embed_gif(ax1,'730E ш18 DC model 33223470 AQ60217.28 снижение оборотов 0.GIF',
              '730E ш18 — заброс оборотов: 2250 RPM')
    embed_gif(ax2,'NTE200 ш85 model снижение оборотов 0.GIF',
              'NTE200 ш85 — заброс оборотов: 2020 RPM')
    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


def page_conclusion(pdf):
    fig = plt.figure(figsize=(11.69,8.27)); page_bg(fig)
    ax_h = fig.add_axes([0,0.88,1,0.12])
    add_header(ax_h,'Выводы по данным INSITE КАМСС',
               'Подтверждённые и новые находки относительно предыдущего анализа')

    conclusions = [
        (C_RED,   'ПОДТВЕРЖДЕНО / КРИТИЧНО',
         'Fuel Code: FC0BU52 (730E) vs FC0WH95 (NTE200) — подтверждено из Calterm Compare Report.\n'
         'DO_Option 6765 vs 61042 и SC_Option 60621 vs 61878 подтверждают разную конфигурацию\n'
         'цифровых выходов защиты между 730E и NTE200.'),
        (C_ORG,   'НОВАЯ НАХОДКА / ВЫСОКИЙ РИСК',
         'FC556 (давление картерных газов): порог снижения момента = 10 кПа у 730E vs 5 кПа у NTE200.\n'
         'NTE200 реагирует на более низкое картерное давление — потенциально правильно,\n'
         'но требует анализа реальных событий FC556 в эксплуатации.'),
        (C_ORG,   'НОВАЯ НАХОДКА / ВЫСОКИЙ РИСК',
         'Заброс (overspeed) оборотов двигателя: 2250 RPM у 730E vs 2020 RPM у NTE200.\n'
         'При максимальных рабочих 2100 об/мин у NTE200 зазор до заброса = лишь 80 RPM.\n'
         'У 730E (макс. 1990 об/мин): зазор = 260 RPM — в 3.25× больше.'),
        (C_YLW,   'ТОЛЬКО У NTE200',
         'FC4642 / FC1852 "Защита — наличие воды в топливе": задержка 1020 с (17 мин) до снижения оборотов.\n'
         'На 730E этот канал отсутствует в данном наборе защит. NTE200 продолжает работать\n'
         '17 минут с обнаруженной водой в топливе до включения защиты.'),
        (C_BLU,   'ИДЕНТИЧНО / ОДИНАКОВО',
         'Следующие параметры защиты ОДИНАКОВЫ у обеих машин (730E = NTE200):\n'
         'FC146 (Темп. ОЖ 100°C, 5 с), FC421 (Темп. масла 121.1°C), FC235 (Уровень ОЖ, 30 с),\n'
         'FC0155-0165 (Темп. воздуха впускного 93.3°C), FC0228 (Давл. ОЖ), FC0151 (Темп. ОЖ 110°C).'),
        (C_ACC,   'РЕКОМЕНДАЦИЯ',
         'Привести DO_Option NTE200 к значению 730E (6765), проверить SC_Option.\n'
         'Уточнить логику FC556 (5 кПа): не вызывает ли это ложных срабатываний?\n'
         'Зазор заброса оборотов 80 RPM недостаточен при работе на 2100 RPM — снизить макс. RPM до 1990.'),
    ]

    ax = fig.add_axes([0.01,0.05,0.98,0.82])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0,1); ax.set_ylim(0,1)
    n = len(conclusions)
    h = 0.14; gap = 0.008
    total = n*h+(n-1)*gap
    y0 = (1-total)/2+total
    for i,(color,label,text) in enumerate(conclusions):
        y = y0-i*(h+gap)
        ax.add_patch(mpatches.FancyBboxPatch((0.0,y-h+0.003),0.006,h-0.006,
            boxstyle='square,pad=0',facecolor=color,linewidth=0,
            transform=ax.transAxes,clip_on=False))
        ax.add_patch(mpatches.FancyBboxPatch((0.008,y-h+0.003),0.99,h-0.006,
            boxstyle='round,pad=0.005',facecolor=C_DARK2,linewidth=0,
            transform=ax.transAxes,clip_on=False))
        ax.text(0.018,y-0.022,label,color=color,fontproperties=fp(8.5,bold=True),va='top')
        ax.text(0.018,y-0.058,text,color=C_WHT,fontproperties=fp(7.5),
                va='top',multialignment='left')

    ax_f = fig.add_axes([0,0,1,0.06])
    ax_f.set_facecolor(C_DARK2); ax_f.axis('off')
    footer(ax_f)
    pdf.savefig(fig,bbox_inches='tight'); plt.close(fig)


# ──────────────────────────────────────────────────────────────────────────────
def main():
    out = '/home/user/NTE200/Анализ_ИНСАЙТ_КАМСС_NTE200_vs_730E.pdf'

    with PdfPages(out) as pdf:
        page_cover(pdf);                print('1 cover')
        page_calterm_compare(pdf);      print('2 calterm')
        page_duty_cycles(pdf);          print('3 duty cycles')
        page_fuel_consumption(pdf);     print('4 fuel')
        page_torque_reduction(pdf,
            '730E ш18 DC model 33223470 AQ60217.28 снижение крутящего момента0.GIF',
            'NTE200 ш85 model снижение крутящего момента 0.GIF',
            'Снижение крутящего момента [лист 1/3]',
            'FC146 Темп.ОЖ · FC415 Давл.масла · FC235 Уровень ОЖ · FC421 Темп.масла · FC556 Картерн.газы')
        print('5 torque 0')
        page_torque_reduction(pdf,
            '730E ш18 DC model 33223470 AQ60217.28 снижение крутящего момента 1.GIF',
            'NTE200 ш85 model снижение крутящего момента 1.GIF',
            'Снижение крутящего момента [лист 2/3]',
            'FC228 Давл.ОЖ · FC155-165 Темп.воздуха впуск.кол.')
        print('6 torque 1')
        page_torque_reduction(pdf,
            '730E ш18 DC model 33223470 AQ60217.28 снижение крутящего момента2.GIF',
            'NTE200 ш85 model снижение крутящего момента 2.GIF',
            'Снижение крутящего момента [лист 3/3]',
            'FC5945 Перепад давл.наддува · FC4729/5114 Разрежение · FC7313 Топл.насос')
        print('7 torque 2')
        page_rpm_reduction(pdf);        print('8 rpm')
        page_bar_comparison(pdf);       print('9 bars')
        page_conclusion(pdf);           print('10 conclusion')

        d = pdf.infodict()
        d['Title']  = 'Анализ INSITE КАМСС — NTE200 vs 730E-DC QSK50 MCRS'
        d['Author'] = 'АО Развитие'

    print(f'\nSaved: {out}  ({os.path.getsize(out)//1024} KB)')

if __name__ == '__main__':
    main()
