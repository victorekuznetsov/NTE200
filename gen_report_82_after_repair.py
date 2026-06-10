#!/usr/bin/env python3
"""NTE200 #82 — QSK50 MCRS V16 — сравнительный отчёт ДВС ДО/ПОСЛЕ ремонта."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String

pdfmetrics.registerFont(TTFont('Sans',      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('Sans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

C = dict(
    bg       = colors.HexColor('#1a1a2e'),
    panel    = colors.HexColor('#16213e'),
    panel2   = colors.HexColor('#0f3460'),
    border   = colors.HexColor('#2a3a5c'),
    text     = colors.HexColor('#e8e8f0'),
    sub      = colors.HexColor('#8899bb'),
    red      = colors.HexColor('#e74c3c'),
    redbg    = colors.HexColor('#2a0d0d'),
    orange   = colors.HexColor('#f39c12'),
    orangebg = colors.HexColor('#2a1800'),
    green    = colors.HexColor('#27ae60'),
    greenbg  = colors.HexColor('#0a1f0a'),
    blue     = colors.HexColor('#3498db'),
    bluebg   = colors.HexColor('#051530'),
    altrow   = colors.HexColor('#1e1e3f'),
    altrow2  = colors.HexColor('#252550'),
    hdr      = colors.HexColor('#0d1829'),
    white    = colors.HexColor('#ffffff'),
    grey     = colors.HexColor('#555577'),
    greybg   = colors.HexColor('#1c1c30'),
    cyan     = colors.HexColor('#00d4ff'),
    cyanbg   = colors.HexColor('#001e2a'),
    purple   = colors.HexColor('#9b59b6'),
    purplebg = colors.HexColor('#1a0a2e'),
)

def Pc(col, txt, bold=False, size=9, lead=13, align=TA_LEFT):
    fn = 'Sans-Bold' if bold else 'Sans'
    s = ParagraphStyle('x', fontName=fn, fontSize=size, textColor=col,
                       leading=lead, alignment=align)
    return Paragraph(txt, s)

def HR(color=None):
    return HRFlowable(width='100%', thickness=0.5, color=color or C['border'],
                      spaceAfter=4, spaceBefore=4)

def SP(h=3):
    return Spacer(1, h * mm)

def section_title(txt, n=None):
    label = f'{n}. ' if n else ''
    return [
        SP(3),
        Pc(C['blue'], label + txt, bold=True, size=12, lead=16),
        HRFlowable(width='100%', thickness=1.0, color=C['blue'],
                   spaceAfter=4, spaceBefore=2),
    ]

def tbl(rows, cols, style_extra=None):
    t = Table(rows, colWidths=cols)
    ts = TableStyle([
        ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
        ('LEFTPADDING',  (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING',   (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
        ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
    ])
    if style_extra:
        for cmd in style_extra:
            ts.add(*cmd)
    t.setStyle(ts)
    return t

def hdr_row(cells, widths, size=7):
    row = [Pc(C['sub'], c, bold=True, size=size, lead=10) for c in cells]
    t = Table([row], colWidths=widths)
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), C['hdr']),
        ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
        ('LEFTPADDING',  (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING',   (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 4),
    ]))
    return t

def info_box(text, bg_color, text_color=None, bold=False, size=8, border_color=None):
    tc = text_color or C['text']
    bc = border_color or bg_color
    cell = Pc(tc, text, bold=bold, size=size, lead=12)
    t = Table([[cell]], colWidths=['100%'])
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), bg_color),
        ('BOX',          (0, 0), (-1, -1), 1.0, bc),
        ('LEFTPADDING',  (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING',   (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
    ]))
    return t

W, H = A4

def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C['bg'])
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    canvas.setFillColor(C['panel'])
    canvas.rect(0, H - 9 * mm, W, 9 * mm, fill=1, stroke=0)
    canvas.setFont('Sans-Bold', 8)
    canvas.setFillColor(C['blue'])
    canvas.drawString(15 * mm, H - 5.5 * mm,
                      'NTE200 №82  •  QSK50 MCRS V16  •  ДО / ПОСЛЕ РЕМОНТА  •  10.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  ESN: 33238517  |  До: 02.06.2026 / После: 10.06.2026')
    canvas.restoreState()

OUT = '/home/user/NTE200/NTE200_82_before_after_repair_10062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20 * mm,
    rightMargin=20 * mm,
    topMargin=16 * mm,
    bottomMargin=11 * mm,
    title='NTE200 №82 — ДВС QSK50 MCRS V16 — ДО/ПОСЛЕ РЕМОНТА — 10.06.2026',
)
PW = doc.width
story = []

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — СРАВНИТЕЛЬНЫЙ ОБЗОР
# ═══════════════════════════════════════════════════════════════════════════════
story += [
    SP(2),
    Pc(C['white'], 'СРАВНИТЕЛЬНЫЙ ОТЧЁТ ДВИГАТЕЛЯ  —  ДО / ПОСЛЕ РЕМОНТА',
       bold=True, size=15, lead=19, align=TA_CENTER),
    SP(1),
    Pc(C['blue'], 'NTE200 №82  /  QSK50 MCRS V16  /  ESN: 33238517',
       bold=True, size=12, lead=16, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],
       'DML ДО ремонта: 02.06.2026 (1 410 строк, тест с породой)  |  '
       'DML ПОСЛЕ ремонта: 10.06.2026, 11:47 (1 002 строки, рейс с породой)',
       size=8.5, lead=12, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# Two-column header: before / after
header_data = [
    [
        Pc(C['orange'], 'ДО РЕМОНТА  (02.06.2026)', bold=True, size=10, lead=13, align=TA_CENTER),
        Pc(C['green'],  'ПОСЛЕ РЕМОНТА  (10.06.2026)', bold=True, size=10, lead=13, align=TA_CENTER),
    ],
    [
        Pc(C['sub'], 'М/Ч ДВС: 4 501 ч  |  RPM avg: 1 896', size=8, lead=11, align=TA_CENTER),
        Pc(C['sub'], 'М/Ч ДВС: 4 545 ч  (+44 ч)  |  RPM avg: 1 897', size=8, lead=11, align=TA_CENTER),
    ],
]
hdr_t = Table(header_data, colWidths=[PW / 2, PW / 2])
hdr_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (0, -1), C['orangebg']),
    ('BACKGROUND', (1, 0), (1, -1), C['greenbg']),
    ('BOX',        (0, 0), (0, -1), 1.0, C['orange']),
    ('BOX',        (1, 0), (1, -1), 1.0, C['green']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING',   (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
    ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
]))
story += [hdr_t, SP(4)]

# Key changes summary table
story += [
    Pc(C['sub'], 'КЛЮЧЕВЫЕ ИЗМЕНЕНИЯ', bold=True, size=9, lead=12, align=TA_CENTER),
    SP(2),
]

CHANGES = [
    # (параметр, до_val, до_col, после_val, после_col, итог, итог_col)
    ('EGT цил.11 (датчик)',   '67.6°C avg (НЕИСПРАВЕН)',  C['red'],    '506.3°C avg (НОРМА)',    C['green'],  '✓ Датчик заменён',                        C['green']),
    ('Δ EGT (разброс)',       '93.8°C (Порог 60°C)',      C['red'],    '71.2°C (Порог 60°C)',    C['orange'], '↓ Улучшение на 22.6°C, ещё выше нормы',   C['orange']),
    ('Цил. >560°C (max)',     '2 цил. (цил.2, цил.6)',    C['red'],    '0 цил.',                 C['green'],  '✓ Ни одного критичного',                  C['green']),
    ('Цил. >520°C avg',       '3 цил. (цил.2,6,15)',      C['orange'], '2 цил. (цил.2, цил.6)',  C['orange'], '↓ -1 цил. в зоне предупр.',               C['orange']),
    ('EGT цил.2 avg',         '547.7°C  (max 567°C)',     C['red'],    '526.3°C  (max 542°C)',   C['orange'], '↓ -21.4°C — снижение, зона предупр.',     C['orange']),
    ('EGT цил.6 avg',         '548.9°C  (max 567°C)',     C['red'],    '538.8°C  (max 552°C)',   C['orange'], '↓ -10.1°C — снижение, зона предупр.',     C['orange']),
    ('T масла max',           '107.7°C',                  C['orange'], '106.7°C',                C['orange'], '↓ -1.0°C — незначительно, ≈то же',        C['sub']),
    ('P масла min (нагрузка)','366.3 кПа',                C['green'],  '368.0 кПа',              C['green'],  '→ Без изменений',                         C['sub']),
    ('T охл. avg',            '83.0°C',                   C['green'],  '83.1°C',                 C['green'],  '→ Норма',                                 C['sub']),
    ('Расход топлива avg',    '361.2 л/ч',                C['sub'],    '356.8 л/ч',              C['sub'],    '↓ -1.2% — незначительное снижение',       C['sub']),
]

ch_cols = ['Параметр', 'ДО ремонта', 'ПОСЛЕ ремонта', 'Итог изменения']
ch_w    = [PW*0.24, PW*0.22, PW*0.22, PW*0.32]
story.append(hdr_row(ch_cols, ch_w))

ch_rows = []
extra   = []
for i, (param, v_b, c_b, v_a, c_a, result, c_r) in enumerate(CHANGES):
    row_bg = C['greenbg'] if c_r == C['green'] else (C['orangebg'] if c_r == C['orange'] else C['altrow'])
    if i % 2 == 0 and c_r == C['sub']:
        row_bg = C['altrow2']
    extra.append(('BACKGROUND', (0, i), (-1, i), row_bg))
    ch_rows.append([
        Pc(C['text'], param,   bold=True, size=8, lead=11),
        Pc(c_b,       v_b,     bold=True, size=8, lead=11),
        Pc(c_a,       v_a,     bold=True, size=8, lead=11),
        Pc(c_r,       result,             size=8, lead=11),
    ])
story.append(tbl(ch_rows, ch_w, style_extra=extra))

story += [
    SP(3),
    info_box(
        'РЕМОНТ 02.06–10.06.2026: Замена/ремонт датчика EGT цилиндра 11 (8L) — '
        'датчик восстановлен, показывает корректные 506°C. '
        'Δ EGT улучшился с 93.8°C до 71.2°C (-22.6°C), '
        'однако остаётся выше порога 60°C. Цил.2 и цил.6 — по-прежнему в зоне предупреждения (520–560°C). '
        'Все остальные параметры без существенных изменений.',
        bg_color=C['panel2'], text_color=C['text'], size=9, border_color=C['blue']
    ),
    SP(1),
    info_box(
        'ТРЕБУЕТСЯ: Диагностика форсунок цил.2 (3R→max 542°C) и цил.6 (1R→max 552°C). '
        'После замены форсунок снять контрольный DML — Δ EGT должен быть &lt;60°C.',
        bg_color=C['orangebg'], text_color=C['orange'], bold=True, size=9, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — EGT CHART: ДО vs ПОСЛЕ (двойной горизонтальный бар)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СРАВНЕНИЕ EGT ЦИЛИНДРОВ — ДО vs ПОСЛЕ РЕМОНТА', 1)
story += [
    SP(1),
    Pc(C['sub'],
       'Фильтр: RPM &gt; 1700 об/мин (под нагрузкой)  |  '
       'Оранжевая линия = ДО (02.06.2026)  |  Синяя линия = ПОСЛЕ (10.06.2026)',
       size=8, lead=11),
    SP(2),
]

# EGT data: (cyl, before_avg, after_avg, before_max, after_max)
EGT_DATA = [
    (1,  520.5, 511.2, 535.0, 525.0),
    (2,  547.7, 526.3, 567.0, 542.0),   # was critical
    (3,  455.1, 467.6, 471.0, 480.0),
    (4,  478.0, 485.1, 497.0, 504.0),
    (5,  496.8, 488.6, 514.0, 507.0),
    (6,  548.9, 538.8, 567.0, 552.0),   # was critical
    (7,  479.4, 480.2, 498.0, 496.0),
    (8,  508.4, 508.2, 532.0, 521.0),
    (9,  499.6, 505.4, 525.0, 519.0),
    (10, 506.0, 504.9, 526.0, 517.0),
    (11,  67.6, 506.3, 471.0, 519.0),   # SENSOR FIXED
    (12, 484.8, 504.8, 499.0, 518.0),
    (13, 471.4, 474.0, 491.0, 489.0),
    (14, 479.1, 478.1, 500.0, 494.0),
    (15, 525.0, 518.6, 542.0, 534.0),
    (16, 496.1, 492.7, 515.0, 513.0),
]

chart_w = PW
chart_h = 180
n_cyl   = len(EGT_DATA)
row_h   = chart_h / n_cyl
label_w = 34
val_w_l = 85   # space for "before" value
val_w_r = 85   # space for "after" value
bar_area = chart_w - label_w - val_w_l - val_w_r
max_val  = 600
bar_split = bar_area / 2  # each bar gets half the width

d = Drawing(chart_w, chart_h)
d.add(Rect(0, 0, chart_w, chart_h,
           fillColor=C['panel'], strokeColor=C['border'], strokeWidth=0.3))

# background legend bands
d.add(Rect(label_w + val_w_l, 0, (490/max_val)*bar_area, chart_h,
           fillColor=colors.HexColor('#0d1525'), strokeColor=None))  # normal zone hint

for i, (cyl, v_b, v_a, mx_b, mx_a) in enumerate(EGT_DATA):
    y = chart_h - (i + 1) * row_h
    # special backgrounds
    if cyl == 11:
        row_bg = C['cyanbg']   # sensor fixed
    elif v_b > 540 or v_a > 530:
        row_bg = C['orangebg']
    elif i % 2 == 0:
        row_bg = C['altrow']
    else:
        row_bg = C['panel']
    d.add(Rect(0, y, chart_w, row_h, fillColor=row_bg, strokeColor=None))

    # Label
    lbl = f'Цил.{cyl:02d}'
    special = (cyl == 11 or v_b > 540 or v_a > 530)
    lbl_col = C['cyan'] if cyl == 11 else (C['orange'] if (v_b > 540 or v_a > 530) else C['sub'])
    d.add(String(2, y + row_h * 0.28, lbl,
                 fontName='Sans-Bold' if special else 'Sans', fontSize=6.5,
                 fillColor=lbl_col))

    bar_y    = y + row_h * 0.15
    bar_h_px = row_h * 0.30

    # Before bar (orange) — upper half of row
    def bar_col_b(v):
        if v < 200:  return C['grey']
        if v > 560:  return C['red']
        if v >= 520: return C['orange']
        return colors.HexColor('#c87722')

    def bar_col_a(v):
        if v > 560:  return C['red']
        if v >= 520: return C['orange']
        if cyl == 11: return C['cyan']
        return C['blue']

    if v_b > 50:  # before bar
        bar_len_b = (min(v_b, max_val) / max_val) * bar_area
        d.add(Rect(label_w + val_w_l, y + row_h * 0.50, bar_len_b, bar_h_px * 0.8,
                   fillColor=bar_col_b(v_b), strokeColor=None))

    # After bar (blue/cyan) — lower half of row
    bar_len_a = (min(v_a, max_val) / max_val) * bar_area
    d.add(Rect(label_w + val_w_l, bar_y, bar_len_a, bar_h_px * 0.8,
               fillColor=bar_col_a(v_a), strokeColor=None))

    # Before value text (left side)
    if v_b < 200:
        txt_b = 'НЕИСПР.'
        bc = C['grey']
    else:
        txt_b = f'{v_b:.0f}°'
        bc = bar_col_b(v_b)
    d.add(String(label_w, y + row_h * 0.56, txt_b,
                 fontName='Sans-Bold', fontSize=6, fillColor=bc))

    # After value text (right of after bar)
    txt_a = f'{v_a:.0f}°'
    if cyl == 11:
        txt_a = f'{v_a:.0f}° ✓FIX'
    ac = bar_col_a(v_a)
    d.add(String(label_w + val_w_l + bar_len_a + 2, y + row_h * 0.16, txt_a,
                 fontName='Sans-Bold', fontSize=6, fillColor=ac))

# Grid lines
for tick in [400, 450, 490, 520, 560]:
    tx = label_w + val_w_l + (tick / max_val) * bar_area
    d.add(Line(tx, 0, tx, chart_h,
               strokeColor=C['border'], strokeWidth=0.3, strokeDashArray=[2, 3]))
    d.add(String(tx - 8, 1, str(tick),
                 fontName='Sans', fontSize=5, fillColor=C['sub']))

# 560 threshold
thresh_x = label_w + val_w_l + (560 / max_val) * bar_area
d.add(Line(thresh_x, 0, thresh_x, chart_h,
           strokeColor=C['red'], strokeWidth=1.0, strokeDashArray=[4, 2]))
d.add(String(thresh_x + 2, chart_h - 8, '560°C ПОРОГ',
             fontName='Sans-Bold', fontSize=5.5, fillColor=C['red']))

# Avg lines
avg_b = 499.8
avg_a = 499.4
ax_b = label_w + val_w_l + (avg_b / max_val) * bar_area
ax_a = label_w + val_w_l + (avg_a / max_val) * bar_area
d.add(Line(ax_b, 0, ax_b, chart_h,
           strokeColor=C['orange'], strokeWidth=0.7, strokeDashArray=[3, 2]))
d.add(String(ax_b + 1, chart_h - 16, f'avg ДО={avg_b:.0f}°',
             fontName='Sans', fontSize=5.5, fillColor=C['orange']))
d.add(Line(ax_a, 0, ax_a, chart_h,
           strokeColor=C['blue'], strokeWidth=0.7, strokeDashArray=[3, 2]))
d.add(String(ax_a - 42, chart_h - 8, f'avg ПОСЛЕ={avg_a:.0f}°',
             fontName='Sans', fontSize=5.5, fillColor=C['blue']))

story.append(d)
story.append(SP(1))

# Legend
legend_egt = [
    (C['orange'], '■ ДО ремонта (верхняя планка)'),
    (C['blue'],   '■ ПОСЛЕ ремонта (нижняя планка)'),
    (C['cyan'],   '■ Цил.11 — датчик восстановлен'),
    (C['red'],    '■ Критично &gt;560°C'),
    (C['grey'],   '■ Неисправный датчик'),
]
lc = [Pc(tc, lbl, size=7, lead=10) for tc, lbl in legend_egt]
lt = Table([lc], colWidths=[PW / 5] * 5)
lt.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING',   (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
]))
story += [lt, SP(3)]

# Detailed table
story += section_title('ДЕТАЛЬНОЕ СРАВНЕНИЕ EGT ПО ЦИЛИНДРАМ', None)
story += [SP(1)]

egt_cols_hdr = ['Цил.', 'Позиция', 'ДО avg', 'ДО max', 'ПОСЛЕ avg', 'ПОСЛЕ max', 'Δ avg', 'Статус']
egt_w = [PW*0.07, PW*0.10, PW*0.10, PW*0.10, PW*0.11, PW*0.11, PW*0.09, PW*0.32]
story.append(hdr_row(egt_cols_hdr, egt_w))

# Cylinder positions (QSK50 V16 layout)
CYL_POS = {
    1: '1L', 2: '1R', 3: '2L', 4: '2R', 5: '3L', 6: '3R',
    7: '4L', 8: '4R', 9: '5L', 10: '5R', 11: '6L', 12: '6R',
    13: '7L', 14: '7R', 15: '8L', 16: '8R',
}

egt_rows = []
egt_extra = []
for i, (cyl, v_b, v_a, mx_b, mx_a) in enumerate(EGT_DATA):
    delta_avg = v_a - v_b
    pos = CYL_POS.get(cyl, '?')

    if cyl == 11:
        status = 'ДАТЧИК ЗАМЕНЁН ✓'
        status_col = C['cyan']
        row_bg = C['cyanbg']
    elif v_a > 560 or mx_a > 560:
        status = 'КРИТИЧНО &gt;560°C'
        status_col = C['red']
        row_bg = C['redbg']
    elif v_a >= 520:
        status = 'ПРЕДУПРЕЖДЕНИЕ'
        status_col = C['orange']
        row_bg = C['orangebg']
    elif v_b < 200:
        status = 'НЕИСПРАВНЫЙ ДАТЧИК'
        status_col = C['grey']
        row_bg = C['greybg']
    else:
        status = 'НОРМА'
        status_col = C['green']
        row_bg = C['greenbg'] if i % 2 == 0 else C['altrow']

    def egt_col(v, is_before_broken=False):
        if is_before_broken:
            return C['grey']
        if v > 560: return C['red']
        if v >= 520: return C['orange']
        return C['green']

    v_b_str = 'НЕИСПР.' if v_b < 200 else f'{v_b:.1f}°C'
    mx_b_str = f'{mx_b:.0f}°C' if v_b > 200 else '—'
    broken_b = (v_b < 200)

    delta_str = f'{delta_avg:+.1f}°C'
    delta_col = C['green'] if delta_avg < -5 else (C['orange'] if delta_avg > 5 else C['sub'])
    if cyl == 11:
        delta_str = '+438°C (FIX)'
        delta_col = C['cyan']

    egt_rows.append([
        Pc(status_col, str(cyl), bold=True, size=8, lead=11, align=TA_CENTER),
        Pc(C['sub'],   pos,                  size=8, lead=11, align=TA_CENTER),
        Pc(egt_col(v_b, broken_b), v_b_str, bold=broken_b, size=8, lead=11, align=TA_CENTER),
        Pc(egt_col(mx_b, broken_b), mx_b_str, size=8, lead=11, align=TA_CENTER),
        Pc(egt_col(v_a), f'{v_a:.1f}°C', bold=(v_a >= 520), size=8, lead=11, align=TA_CENTER),
        Pc(egt_col(mx_a), f'{mx_a:.0f}°C', bold=(mx_a >= 520), size=8, lead=11, align=TA_CENTER),
        Pc(delta_col, delta_str, bold=True, size=8, lead=11, align=TA_CENTER),
        Pc(status_col, status, bold=(status != 'НОРМА'), size=7.5, lead=11),
    ])
    egt_extra.append(('BACKGROUND', (0, i), (-1, i), row_bg))

story.append(tbl(egt_rows, egt_w, style_extra=egt_extra))
story += [SP(2)]

story += [
    info_box(
        'Цил.11 (6L): датчик ЕGT заменён — показания восстановлены (67.6°C → 506.3°C). '
        'Цил.2 (1R) и цил.6 (3R): после ремонта снизились, но avg 526°C и 539°C — '
        'в зоне предупреждения (520–560°C). Необходима диагностика форсунок. '
        'Δ EGT улучшился: 93.8°C → 71.2°C, но порог 60°C не достигнут.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8.5, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — ПАРАМЕТРЫ ДВС: ДО vs ПОСЛЕ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ПАРАМЕТРЫ ДВИГАТЕЛЯ: ДО vs ПОСЛЕ (под нагрузкой, RPM &gt; 1700)', 2)
story += [SP(1)]

# Full params comparison table
par_hdr = ['Параметр', 'Норма / порог', 'ДО avg', 'ДО max', 'ПОСЛЕ avg', 'ПОСЛЕ max', 'Изм.avg', 'Оценка']
par_w = [PW*0.24, PW*0.15, PW*0.09, PW*0.09, PW*0.10, PW*0.10, PW*0.09, PW*0.14]
story.append(hdr_row(par_hdr, par_w))

# (param, norm, b_avg, b_max, a_avg, a_max, delta, status, row_bg, status_col)
PAR_DATA = [
    ('RPM (об/мин)',           '1500–1900',        1895.9, 1938.0, 1897.3, 1926.0, +1.4,  'НОРМА',     C['greenbg'],  C['green']),
    ('T охл. (°C)',            '&lt; 95°C',         83.0,   84.3,   83.1,   87.0,   +0.1,  'НОРМА',     C['greenbg'],  C['green']),
    ('T масла (°C)',           '&lt; 105°C (пред.)',103.1, 107.7,  103.9, 106.7,   +0.8,  'НАБЛЮДАТЬ', C['orangebg'], C['orange']),
    ('P масла avg (кПа)',      '&gt; 350',          510.2,  556.5,  510.0, 533.9,   -0.2,  'НОРМА',     C['greenbg'],  C['green']),
    ('P масла min (кПа)',      '&gt; 350',          366.3,  '—',    368.0, '—',     +1.7,  'НОРМА',     C['greenbg'],  C['green']),
    ('P наддув (кПа)',         '220–320',           269.5,  283.2,  274.3, 288.9,   +4.7,  'НОРМА',     C['greenbg'],  C['green']),
    ('T впуска (°C)',          '&lt; 80°C',          65.8,   70.2,   65.2,  69.2,   -0.6,  'НОРМА',     C['greenbg'],  C['green']),
    ('Нагрузка (%)',           '—',                 98.0,  100.0,   97.1, 100.0,   -0.9,  'НОРМА',     C['altrow'],   C['sub']),
    ('Расход топлива (л/ч)',   '—',                361.2,  429.0,  356.8, 420.3,   -4.4,  'OK (-1.2%)',C['altrow2'],  C['sub']),
    ('P картера (кПа)',        '&lt; 3 кПа',         0.2,    0.6,    0.2,   0.5,   -0.0,  'НОРМА',     C['greenbg'],  C['green']),
    ('EGT avg все цил. (°C)', '450–520',           499.8,  548.9,  499.4, 538.8,   -0.4,  'НОРМА avg', C['altrow'],   C['sub']),
    ('Δ EGT (°C)',             '&lt; 60°C',          93.8,   '—',    71.2,  '—',   -22.6,  'ВЫШЕ НОРМЫ',C['orangebg'], C['orange']),
]

par_rows = []
par_extra = []
for i, row in enumerate(PAR_DATA):
    param, norm, b_avg, b_max, a_avg, a_max, delta, status, row_bg, scol = row
    b_avg_s = f'{b_avg:.1f}' if isinstance(b_avg, float) else str(b_avg)
    b_max_s = f'{b_max:.1f}' if isinstance(b_max, float) else str(b_max)
    a_avg_s = f'{a_avg:.1f}' if isinstance(a_avg, float) else str(a_avg)
    a_max_s = f'{a_max:.1f}' if isinstance(a_max, float) else str(a_max)
    delta_s = f'{delta:+.1f}'
    delta_col = C['green'] if delta < -0.5 else (C['orange'] if delta > 1 else C['sub'])

    par_rows.append([
        Pc(C['text'],  param,   size=8, lead=11),
        Pc(C['sub'],   norm,    size=7, lead=10),
        Pc(C['sub'],   b_avg_s, size=8, lead=11, align=TA_CENTER),
        Pc(C['sub'],   b_max_s, size=8, lead=11, align=TA_CENTER),
        Pc(scol,       a_avg_s, bold=True, size=8, lead=11, align=TA_CENTER),
        Pc(scol,       a_max_s, size=8, lead=11, align=TA_CENTER),
        Pc(delta_col,  delta_s, bold=True, size=8, lead=11, align=TA_CENTER),
        Pc(scol,       status,  bold=(scol != C['sub']), size=7.5, lead=10),
    ])
    par_extra.append(('BACKGROUND', (0, i), (-1, i), row_bg))

story.append(tbl(par_rows, par_w, style_extra=par_extra))
story += [SP(2)]

story += [
    info_box(
        'T масла max=106.7°C (было 107.7°C) — снижение на 1°C, по-прежнему выше порога 105°C. '
        'Рекомендуется диагностика масляного теплообменника при следующем плановом ТО. '
        'Все остальные параметры — без существенных изменений после ремонта.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8.5, border_color=C['orange']
    ),
    SP(3),
]

# ── Summary KPI comparison ─────────────────────────────────────────────────────
story += section_title('ИТОГОВАЯ ОЦЕНКА: ДО vs ПОСЛЕ', None)
story += [SP(2)]

BEFORE_AFTER_KPI = [
    # (label, before_val, before_col, before_bg, after_val, after_col, after_bg)
    ('EGT ЦИЛ.11',  'НЕИСПРАВЕН\n67.6°C avg',  C['red'],    C['redbg'],    'ИСПРАВЕН\n506.3°C avg', C['green'],  C['greenbg']),
    ('Δ EGT',        '93.8°C\n&gt;&gt; 60°C',   C['red'],    C['redbg'],    '71.2°C\n&gt; 60°C',     C['orange'], C['orangebg']),
    ('ЦИЛ. &gt;560°C (max)', '2 цил.\nцил.2, цил.6',C['red'], C['redbg'],  '0 цил.\nнет критичных', C['green'],  C['greenbg']),
    ('T МАСЛА MAX',  '107.7°C',               C['orange'], C['orangebg'], '106.7°C',               C['orange'], C['orangebg']),
]

kpi_hdr  = [Pc(C['sub'], f'', bold=True, size=7, lead=10, align=TA_CENTER)] + \
           [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, *_ in BEFORE_AFTER_KPI]

kpi_w    = [PW * 0.12] + [PW * 0.22] * len(BEFORE_AFTER_KPI)

def kpi_val_cell(val, col, bg):
    return Pc(col, val, bold=True, size=11, lead=15, align=TA_CENTER)

row_before_hdr = [Pc(C['orange'], 'ДО', bold=True, size=9, lead=12, align=TA_CENTER)]
row_before_val = [Pc(C['orange'], 'ДО', bold=True, size=9, lead=12, align=TA_CENTER)]
row_after_hdr  = [Pc(C['green'],  'ПОСЛЕ', bold=True, size=9, lead=12, align=TA_CENTER)]
row_after_val  = [Pc(C['green'],  'ПОСЛЕ', bold=True, size=9, lead=12, align=TA_CENTER)]

for lbl, bv, bc, bbg, av, ac, abg in BEFORE_AFTER_KPI:
    row_before_val.append(kpi_val_cell(bv, bc, bbg))
    row_after_val.append(kpi_val_cell(av, ac, abg))

hdr_row2 = [Pc(C['sub'], '', bold=True, size=7, lead=10, align=TA_CENTER)] + \
           [Pc(C['sub'], lbl, bold=True, size=8, lead=11, align=TA_CENTER) for lbl, *_ in BEFORE_AFTER_KPI]

kpi_t = Table([hdr_row2, row_before_val, row_after_val], colWidths=kpi_w)
kpi_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING',   (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ('BACKGROUND',   (0, 0), (-1, 0), C['hdr']),
    ('BACKGROUND',   (0, 1), (0, 1), C['orangebg']),
    ('BACKGROUND',   (0, 2), (0, 2), C['greenbg']),
])
for j, (lbl, bv, bc, bbg, av, ac, abg) in enumerate(BEFORE_AFTER_KPI):
    kpi_ts.add('BACKGROUND', (j+1, 1), (j+1, 1), bbg)
    kpi_ts.add('BACKGROUND', (j+1, 2), (j+1, 2), abg)
kpi_t.setStyle(kpi_ts)
story += [kpi_t, SP(4)]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — СЛЕДУЮЩИЕ ШАГИ И РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СЛЕДУЮЩИЕ ШАГИ И РЕКОМЕНДАЦИИ', 3)
story += [SP(2)]

RECS = [
    (
        '1. ВЫПОЛНЕНО ✓',
        'Замена датчика EGT цил.11 (6L) — УСПЕШНО',
        'Цилиндр 11 (позиция 6L, левая банка): датчик EGT восстановлен. '
        'До ремонта avg=67.6°C (false low), после: avg=506.3°C — корректное значение. '
        'Показания цил.11 в норме (450–520°C).',
        C['green'], C['greenbg'],
    ),
    (
        '2. ТРЕБУЕТСЯ',
        'Диагностика/замена форсунки цил.2 (1R) — avg 526°C, max 542°C',
        'Цилиндр 2 (позиция 1R, правая банка): avg=526.3°C, max=542°C — '
        'зона предупреждения (520–560°C). До ремонта было 547.7°C (max 567°C), '
        'снизился на 21.4°C, но порог 520°C не преодолён. '
        'Проверить форсунку, зазоры клапанов ГБЦ цил.2.',
        C['orange'], C['orangebg'],
    ),
    (
        '3. ТРЕБУЕТСЯ',
        'Диагностика/замена форсунки цил.6 (3R) — avg 539°C, max 552°C',
        'Цилиндр 6 (позиция 3R, правая банка): avg=538.8°C, max=552°C — '
        'зона предупреждения, снизился на 10.1°C от 548.9°C. '
        'Наибольший приоритет: этот цилиндр остаётся самым горячим. '
        'Параллельная диагностика с цил.2.',
        C['orange'], C['orangebg'],
    ),
    (
        '4. ТРЕБУЕТСЯ',
        'Контрольный DML после диагностики форсунок цил.2 и цил.6',
        'После замены/регулировки форсунок цил.2 и цил.6 выполнить контрольный '
        'DML-замер (рейс с породой). Цель: Δ EGT &lt; 60°C и оба цилиндра avg &lt; 520°C. '
        'Текущий Δ EGT = 71.2°C — выше нормы на 11.2°C.',
        C['blue'], C['bluebg'],
    ),
    (
        '5. НАБЛЮДЕНИЕ',
        'T масла max=106.7°C &gt; 105°C — диагностика теплообменника при ТО',
        'Пиковое значение T масла незначительно снизилось (107.7→106.7°C), '
        'но по-прежнему превышает порог 105°C. Среднее 103.9°C — в норме. '
        'При следующем плановом ТО провести диагностику масляного теплообменника.',
        C['sub'], C['greybg'],
    ),
]

recs_w = [PW*0.18, PW*0.33, PW*0.49]
rec_rows = []
for prio, title, body, tc, bg in RECS:
    rec_rows.append([
        Pc(tc,        prio,  bold=True, size=9, lead=13),
        Pc(tc,        title, bold=True, size=8.5, lead=13),
        Pc(C['text'], body,             size=8, lead=12),
    ])

rec_t = Table(rec_rows, colWidths=recs_w)
rec_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 7),
    ('RIGHTPADDING', (0, 0), (-1, -1), 7),
    ('TOPPADDING',   (0, 0), (-1, -1), 8),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 8),
    ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
])
for i, (_, _, _, _, bg) in enumerate(RECS):
    rec_ts.add('BACKGROUND', (0, i), (-1, i), bg)
rec_t.setStyle(rec_ts)
story += [rec_t, SP(4)]

story += [
    HR(C['blue']),
    SP(2),
    Pc(C['blue'], 'ИТОГ РЕМОНТА', bold=True, size=11, lead=14, align=TA_CENTER),
    SP(1),
]

# Final status blocks
final_data = [
    [
        Pc(C['sub'], 'УСТРАНЕНО', bold=True, size=7, lead=10, align=TA_CENTER),
        Pc(C['sub'], 'УЛУЧШИЛОСЬ', bold=True, size=7, lead=10, align=TA_CENTER),
        Pc(C['sub'], 'ОСТАЁТСЯ', bold=True, size=7, lead=10, align=TA_CENTER),
        Pc(C['sub'], 'СЛЕДУЮЩИЙ ШАГ', bold=True, size=7, lead=10, align=TA_CENTER),
    ],
    [
        Pc(C['green'],  'EGT цил.11\nдатчик исправен', bold=True, size=10, lead=15, align=TA_CENTER),
        Pc(C['orange'], 'Δ EGT\n93.8 → 71.2°C', bold=True, size=10, lead=15, align=TA_CENTER),
        Pc(C['orange'], 'Цил.2 и цил.6\n&gt;520°C avg', bold=True, size=10, lead=15, align=TA_CENTER),
        Pc(C['blue'],   'Форсунки цил.2+6\n→ DML контроль', bold=True, size=10, lead=15, align=TA_CENTER),
    ],
]
final_t = Table(final_data, colWidths=[PW / 4] * 4)
final_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING',   (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 10),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ('BACKGROUND',   (0, 0), (-1, 0), C['hdr']),
    ('BACKGROUND',   (0, 1), (0, 1), C['greenbg']),
    ('BACKGROUND',   (1, 1), (1, 1), C['orangebg']),
    ('BACKGROUND',   (2, 1), (2, 1), C['orangebg']),
    ('BACKGROUND',   (3, 1), (3, 1), C['bluebg']),
])
final_t.setStyle(final_ts)
story += [final_t, SP(2)]

story += [
    HR(),
    SP(1),
    Pc(C['sub'],
       'Отчёт подготовлен на основе: '
       'DML ДО ремонта: DML-20260602-171259 82 (тест 3 с породой).csv (1 410 строк, 02.06.2026, М/Ч 4501)  •  '
       'DML ПОСЛЕ ремонта: DML-NTE200#82 after repair 20260610-120416 (1 002 строки, 10.06.2026, М/Ч 4545)  |  '
       'ESN: 33238517  •  Горная Евразия  •  10.06.2026',
       size=7, lead=10),
]

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
