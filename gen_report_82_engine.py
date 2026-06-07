#!/usr/bin/env python3
"""NTE200 #82 — QSK50 MCRS V16 — Engine-only Diagnostic Report (07.06.2026)."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String, Group
from reportlab.graphics import renderPDF

# ── Font registration ──────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont('Sans',      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('Sans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# ── Colour palette ─────────────────────────────────────────────────────────────
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
    orangebg = colors.HexColor('#2a1a05'),
    yellow   = colors.HexColor('#f1c40f'),
    yellowbg = colors.HexColor('#2a2200'),
    green    = colors.HexColor('#27ae60'),
    greenbg  = colors.HexColor('#0a1f10'),
    blue     = colors.HexColor('#3498db'),
    bluebg   = colors.HexColor('#051530'),
    altrow   = colors.HexColor('#1e2d4a'),
    hdr      = colors.HexColor('#0d1829'),
    white    = colors.HexColor('#ffffff'),
    grey     = colors.HexColor('#555577'),
    greybg   = colors.HexColor('#1a1a2e'),
)

# ── Helpers ────────────────────────────────────────────────────────────────────
def P(txt, **kw):
    s = ParagraphStyle('x', fontName='Sans', fontSize=9, textColor=C['text'], leading=13, **kw)
    return Paragraph(txt, s)

def PB(txt, **kw):
    s = ParagraphStyle('x', fontName='Sans-Bold', fontSize=9, textColor=C['text'], leading=13, **kw)
    return Paragraph(txt, s)

def Pc(col, txt, bold=False, size=9, lead=13, align=TA_LEFT):
    fn = 'Sans-Bold' if bold else 'Sans'
    s = ParagraphStyle('x', fontName=fn, fontSize=size, textColor=col, leading=lead, alignment=align)
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
        HRFlowable(width='100%', thickness=1.0, color=C['blue'], spaceAfter=4, spaceBefore=2),
    ]

def tbl(rows, cols, style_extra=None, alt=True, hdr_row_idx=None):
    t = Table(rows, colWidths=cols)
    ts = TableStyle([
        ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
        ('LEFTPADDING',  (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING',   (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
        ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
    ])
    for i in range(len(rows)):
        if hdr_row_idx is not None and i == hdr_row_idx:
            ts.add('BACKGROUND', (0, i), (-1, i), C['hdr'])
        else:
            bg = C['panel'] if i % 2 == 0 else (C['altrow'] if alt else C['panel'])
            ts.add('BACKGROUND', (0, i), (-1, i), bg)
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
        ('LEFTPADDING',  (0, 0), (-1, -1), 7),
        ('RIGHTPADDING', (0, 0), (-1, -1), 7),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
    ]))
    return t

# ── Page layout ────────────────────────────────────────────────────────────────
W, H = A4

def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C['bg'])
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # Top bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, H - 9 * mm, W, 9 * mm, fill=1, stroke=0)
    canvas.setFont('Sans-Bold', 8)
    canvas.setFillColor(C['blue'])
    canvas.drawString(15 * mm, H - 5.5 * mm,
                      'NTE200 №82  •  QSK50 MCRS V16  •  Диагностический отчёт ДВС  •  07.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    # Bottom bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  Новый ДВС: 4 258 м/ч  |  DML замер: 02.06.2026')
    canvas.restoreState()

# ── Document setup ─────────────────────────────────────────────────────────────
OUT = '/home/user/NTE200/NTE200_82_engine_report_07062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20 * mm,
    rightMargin=20 * mm,
    topMargin=18 * mm,
    bottomMargin=13 * mm,
    title='NTE200 №82 — Диагностика ДВС QSK50 MCRS V16 — 07.06.2026',
)
PW = doc.width
story = []

# ═══════════════════════════════════════════════════════════════════════════════
# СТРАНИЦА 1 — ТИТУЛ + KPI КАРТОЧКИ
# ═══════════════════════════════════════════════════════════════════════════════
story += [
    SP(2),
    Pc(C['white'], 'ДИАГНОСТИЧЕСКИЙ ОТЧЁТ ДВИГАТЕЛЯ',
       bold=True, size=16, lead=20, align=TA_CENTER),
    SP(1),
    Pc(C['blue'], 'NTE200 №82  /  QSK50 MCRS V16',
       bold=True, size=13, lead=17, align=TA_CENTER),
    SP(1),
    Pc(C['sub'], 'Дата отчёта: 07.06.2026  |  DML замер: 02.06.2026  |  Тест с породой, 1411 строк',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# Machine info block
info_data = [
    [Pc(C['sub'], 'Гаражный №', bold=True, size=8, lead=11),
     Pc(C['text'], '82', size=8, lead=11),
     Pc(C['sub'], 'Модель', bold=True, size=8, lead=11),
     Pc(C['text'], 'NTE200, QSK50 MCRS V16', size=8, lead=11)],
    [Pc(C['sub'], 'М/Ч нового ДВС', bold=True, size=8, lead=11),
     Pc(C['green'], '4 258 м/ч', bold=True, size=8, lead=11),
     Pc(C['sub'], 'Замена ДВС', bold=True, size=8, lead=11),
     Pc(C['orange'], 'Ноябрь/Декабрь 2025', size=8, lead=11)],
    [Pc(C['sub'], 'М/Ч ДВС по ECM (DML)', bold=True, size=8, lead=11),
     Pc(C['text'], '4 501 ч (004501:04:20)', size=8, lead=11),
     Pc(C['sub'], 'Старый ДВС на момент замены', bold=True, size=8, lead=11),
     Pc(C['sub'], '~36 680 м/ч', size=8, lead=11)],
]
info_t = Table(info_data, colWidths=[PW*0.22, PW*0.28, PW*0.22, PW*0.28])
info_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('GRID', (0, 0), (-1, -1), 0.3, C['border']),
    ('LEFTPADDING', (0, 0), (-1, -1), 7),
    ('RIGHTPADDING', (0, 0), (-1, -1), 7),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('BACKGROUND', (0, 1), (-1, 1), C['altrow']),
]))
story += [info_t, SP(3)]

# KPI карточки — 6 штук в 3+3
kpi_cards = [
    ('М/Ч НОВОГО ДВС',         '4 258',       C['green'],   C['greenbg']),
    ('МАСЛО ДВС',              'НОРМА',        C['green'],   C['greenbg']),
    ('ЗАМЕНА ДВС',             'Дек. 2025',   C['orange'],  C['orangebg']),
    ('EGT ЦИЛ. 11',            'НЕИСПРАВЕН',  C['red'],     C['redbg']),
    ('ДЕЛЬТА EGT (без цил.11)','92 °C',       C['orange'],  C['orangebg']),
    ('T МАСЛА MAX',            '107.7 °C',    C['yellow'],  C['yellowbg']),
]
kpi_sub = [
    '(ДВС заменён нояб./дек. 2025)',
    '17 проб — всё чисто',
    'Старый ДВС 36 680 м/ч',
    '99°C вместо ~480°C',
    'Предупреждение (порог 60°C)',
    'Наблюдать (порог 105°C)',
]
kpi_w = [PW / 3] * 3
# Row 1
kpi_hdr1 = [Pc(C['sub'], kpi_cards[i][0], bold=True, size=7, lead=10, align=TA_CENTER) for i in range(3)]
kpi_val1 = [Pc(kpi_cards[i][2], kpi_cards[i][1], bold=True, size=14, lead=18, align=TA_CENTER) for i in range(3)]
kpi_sub1 = [Pc(C['sub'], kpi_sub[i], size=6, lead=9, align=TA_CENTER) for i in range(3)]
kpi_t1 = Table([kpi_hdr1, kpi_val1, kpi_sub1], colWidths=kpi_w)
kpi_ts1 = TableStyle([
    ('GRID', (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
])
for ci in range(3):
    kpi_ts1.add('BACKGROUND', (ci, 0), (ci, 0), C['hdr'])
    kpi_ts1.add('BACKGROUND', (ci, 1), (ci, 1), kpi_cards[ci][3])
    kpi_ts1.add('BACKGROUND', (ci, 2), (ci, 2), C['panel'])
kpi_t1.setStyle(kpi_ts1)

# Row 2
kpi_hdr2 = [Pc(C['sub'], kpi_cards[i][0], bold=True, size=7, lead=10, align=TA_CENTER) for i in range(3, 6)]
kpi_val2 = [Pc(kpi_cards[i][2], kpi_cards[i][1], bold=True, size=14, lead=18, align=TA_CENTER) for i in range(3, 6)]
kpi_sub2 = [Pc(C['sub'], kpi_sub[i], size=6, lead=9, align=TA_CENTER) for i in range(3, 6)]
kpi_t2 = Table([kpi_hdr2, kpi_val2, kpi_sub2], colWidths=kpi_w)
kpi_ts2 = TableStyle([
    ('GRID', (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
])
for ci in range(3):
    kpi_ts2.add('BACKGROUND', (ci, 0), (ci, 0), C['hdr'])
    kpi_ts2.add('BACKGROUND', (ci, 1), (ci, 1), kpi_cards[ci+3][3])
    kpi_ts2.add('BACKGROUND', (ci, 2), (ci, 2), C['panel'])
kpi_t2.setStyle(kpi_ts2)

story += [kpi_t1, SP(1), kpi_t2, SP(4)]

# Summary status box
story += [
    info_box(
        'СТАТУС ДВИГАТЕЛЯ: Новый ДВС (нояб./дек. 2025) наработал 4 258 м/ч. '
        'Масло в норме по всем 17 пробам. '
        'ТРЕБУЕТСЯ: замена датчика EGT цил.11. Контроль цил.3 (438°C — минимальная температура выхлопа). '
        'Дельта EGT = 92°C превышает норму 60°C.',
        bg_color=C['orangebg'], text_color=C['orange'], bold=False, size=9, border_color=C['orange']
    ),
    SP(2),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# СТРАНИЦА 2 — DML ПАРАМЕТРЫ ДВИГАТЕЛЯ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('DML ПАРАМЕТРЫ ДВИГАТЕЛЯ  (02.06.2026, тест с породой)', 1)
story += [SP(1),
    Pc(C['sub'],
       'Запись: 1411 строк  |  М/Ч ЭБУ: 004501:04:20  |  Тест с породой под нагрузкой',
       size=8, lead=11),
    SP(2),
]

dml_cols = ['Параметр', 'Значение', 'Норма / порог', 'Статус']
dml_w = [PW*0.38, PW*0.22, PW*0.24, PW*0.16]
story.append(hdr_row(dml_cols, dml_w))

dml_data = [
    ('RPM средн.',                  '1 822 об/мин',    '1 500–1 900',       'НОРМА',        C['greenbg'],  C['green']),
    ('RPM макс.',                   '1 929 об/мин',    '< 1 950',           'НОРМА',        C['greenbg'],  C['green']),
    ('RPM мин.',                    '799 об/мин',      '(холостой ход)',     'OK',           C['panel'],    C['sub']),
    ('T масла avg',                 '102.2 °C',        '< 105 °C (предупр.)', 'НОРМА',      C['greenbg'],  C['green']),
    ('T масла max',                 '107.7 °C',        '< 105 °C (предупр.)\n< 115 °C (крит.)', 'НАБЛЮДАТЬ', C['yellowbg'], C['yellow']),
    ('T масла min',                 '87.9 °C',         '—',                 'OK',           C['panel'],    C['sub']),
    ('P масла avg',                 '496 кПа',         '> 350 кПа',         'НОРМА',        C['greenbg'],  C['green']),
    ('P масла min (кратковременно)','267 кПа',         '> 350 кПа',         'НАБЛЮДАТЬ',    C['orangebg'], C['orange']),
    ('P масла max',                 '556 кПа',         '< 620 кПа',         'НОРМА',        C['greenbg'],  C['green']),
    ('P наддува (впускной коллект.)','238 кПа',        '220–320 кПа',       'НОРМА',        C['greenbg'],  C['green']),
    ('T впускного коллектора',      '64–66 °C',        '< 80 °C',           'НОРМА',        C['greenbg'],  C['green']),
    ('P картерных газов',           'данные присутств.','< 3.0 кПа',        'НОРМА',        C['greenbg'],  C['green']),
]

dml_rows = []
for param, val, norm, status_txt, row_bg, status_col in dml_data:
    dml_rows.append([
        Pc(C['text'],   param,      size=8, lead=11),
        Pc(status_col,  val,  bold=True, size=8, lead=11),
        Pc(C['sub'],    norm,       size=7, lead=10),
        Pc(status_col,  status_txt, bold=True, size=8, lead=11),
    ])

dml_extra = [('BACKGROUND', (0, i), (-1, i), dml_data[i][4]) for i in range(len(dml_rows))]
story.append(tbl(dml_rows, dml_w, style_extra=dml_extra, alt=False))
story.append(SP(2))

# Warning boxes for DML
story += [
    info_box(
        'T масла MAX 107.7°C: выше порога предупреждения 105°C. '
        'Критический порог 115°C не достигнут. Среднее значение 102.2°C — норма. Контролировать.',
        bg_color=C['yellowbg'], text_color=C['yellow'], size=8, border_color=C['yellow']
    ),
    SP(1),
    info_box(
        'P масла MIN 267 кПа: ниже порога 350 кПа. Возможно при пуске или снижении нагрузки. '
        'Среднее давление 496 кПа — в норме. Контролировать при следующем DML.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(3),
]

# EGT Section header
story += section_title('ТЕМПЕРАТУРА ВЫХЛОПА ЦИЛИНДРОВ — EGT (16 цил., avg при нагрузке)', 2)
story += [SP(1),
    Pc(C['sub'],
       'Средняя без цил.11: ~485°C  |  Цил.11 НЕИСПРАВЕН (датчик 99°C)  |  Дельта = 92°C (порог 60°C)',
       size=8, lead=11),
    SP(2),
]

# ── EGT bar chart (horizontal, via ReportLab Drawing) ─────────────────────────
EGT = [
    (1,  500, False),
    (2,  525, False),
    (3,  438, False),
    (4,  458, False),
    (5,  475, False),
    (6,  530, False),
    (7,  458, False),
    (8,  490, False),
    (9,  481, False),
    (10, 487, False),
    (11,  99, True),   # неисправен
    (12, 467, False),
    (13, 451, False),
    (14, 457, False),
    (15, 502, False),
    (16, 471, False),
]

AVG_NORMAL = 485  # среднее без цил.11

chart_w = PW
chart_h = 175
row_h = chart_h / 16
label_w = 30
val_w = 38
bar_area = chart_w - label_w - val_w - 4
max_val = 580
min_bar = 0

d = Drawing(chart_w, chart_h)

# Background
d.add(Rect(0, 0, chart_w, chart_h,
           fillColor=C['panel'], strokeColor=C['border'], strokeWidth=0.3))

def egt_bar_color(val, broken):
    if broken:
        return C['grey']
    if val > 560:
        return C['red']
    if val >= 520:
        return C['yellow']
    return C['green']

for i, (cyl, val, broken) in enumerate(EGT):
    y = chart_h - (i + 1) * row_h
    # Alternating row bg
    row_bg_col = C['altrow'] if i % 2 == 0 else C['panel']
    d.add(Rect(0, y, chart_w, row_h,
               fillColor=row_bg_col, strokeColor=None, strokeWidth=0))

    # Cylinder label
    lbl = f'Цил.{cyl:2d}'
    d.add(String(2, y + row_h * 0.28, lbl,
                 fontName='Sans', fontSize=6.5, fillColor=C['sub'].hexval()))

    # Bar
    if broken:
        bar_len = 8  # tiny bar for broken sensor
        bar_col = C['grey']
    else:
        bar_len = (val / max_val) * bar_area
        bar_col = egt_bar_color(val, broken)

    bar_y = y + row_h * 0.15
    bar_h_px = row_h * 0.70
    d.add(Rect(label_w, bar_y, bar_len, bar_h_px,
               fillColor=bar_col, strokeColor=None, strokeWidth=0))

    # Value label
    if broken:
        val_txt = '99°C ДАТЧИК!'
        vc = C['red']
    else:
        val_txt = f'{val}°C'
        vc = egt_bar_color(val, broken)
    d.add(String(label_w + bar_len + 3, y + row_h * 0.28, val_txt,
                 fontName='Sans-Bold' if (val >= 520 or broken) else 'Sans',
                 fontSize=6.5, fillColor=vc.hexval()))

# Average reference line
avg_x = label_w + (AVG_NORMAL / max_val) * bar_area
d.add(Line(avg_x, 0, avg_x, chart_h,
           strokeColor=C['blue'].hexval(), strokeWidth=0.8,
           strokeDashArray=[3, 2]))
d.add(String(avg_x + 2, chart_h - 8, f'avg={AVG_NORMAL}°C',
             fontName='Sans', fontSize=6, fillColor=C['blue'].hexval()))

# Green zone marker (350–520)
green_x1 = label_w + (350 / max_val) * bar_area
green_x2 = label_w + (520 / max_val) * bar_area
# Yellow zone 520-560
yellow_x2 = label_w + (560 / max_val) * bar_area

# Axis ticks at bottom
for tick_val in [350, 400, 450, 500, 520, 560]:
    tx = label_w + (tick_val / max_val) * bar_area
    d.add(Line(tx, 0, tx, chart_h,
               strokeColor=C['border'].hexval(), strokeWidth=0.3,
               strokeDashArray=[2, 3]))
    d.add(String(tx - 6, 1, str(tick_val),
                 fontName='Sans', fontSize=5, fillColor=C['sub'].hexval()))

story.append(d)
story.append(SP(1))

# Legend
legend_data = [
    [Pc(C['green'],  '■ Норма (350–520°C)', size=7, lead=10),
     Pc(C['yellow'], '■ Предупреждение (520–560°C)', size=7, lead=10),
     Pc(C['red'],    '■ Критично (>560°C)', size=7, lead=10),
     Pc(C['grey'],   '■ Датчик неисправен', size=7, lead=10),
     Pc(C['blue'],   '─ ─ Среднее (485°C)', size=7, lead=10)],
]
leg_t = Table(legend_data, colWidths=[PW*0.22, PW*0.28, PW*0.20, PW*0.18, PW*0.12])
leg_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
]))
story += [leg_t, SP(2)]

# EGT analysis boxes
story += [
    info_box(
        'ЦИЛ. 11 — НЕИСПРАВЕН: датчик EGT показывает 99°C (реальная T ~480°C). '
        'Требуется немедленная замена датчика.',
        bg_color=C['redbg'], text_color=C['red'], bold=True, size=8, border_color=C['red']
    ),
    SP(1),
    info_box(
        'ДЕЛЬТА EGT = 92°C (без цил.11): порог предупреждения 60°C превышен. '
        'Максимум: цил.6=530°C, цил.2=525°C. Минимум: цил.3=438°C. '
        'Цил.3 значительно ниже среднего → возможна проблема форсунки или ГБЦ цил.3. '
        'Рекомендуется снятие и проверка форсунки цил.3.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(2),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# СТРАНИЦА 3 — СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА НОВОГО ДВС
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА — НОВЫЙ ДВС (дек.2025 — май 2026)', 3)
story += [SP(1)]

story.append(info_box(
    'Новый ДВС (ноябрь/декабрь 2025) — 17 проб, 4 258 м/ч — все показатели в НОРМЕ.\n'
    'Fe = 0–2 мг/кг (норма < 30). Кремний Si = 0 во всех пробах. '
    'TBN снижается с 124 до 95 (норма при цикле замены 250 м/ч — ожидаемо). '
    'TAN стабильный 14.3–16.0. Сажа: рост с 131 → 158 мг/кг — наблюдать (допустимо).',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

oil_cols = ['Дата', 'М/Ч ДВС', 'Нараб.масла', 'Fe', 'Cu', 'Si', 'TAN', 'TBN', 'Сажа', 'Статус']
oil_w = [PW*0.11, PW*0.09, PW*0.10, PW*0.06, PW*0.06, PW*0.06, PW*0.07, PW*0.07, PW*0.08, PW*0.30]
story.append(hdr_row(oil_cols, oil_w))

oil_data_raw = [
    ('01.12.2025',  249,  200,  0, 0, 0, 15.6, 122.2, 134),
    ('12.12.2025',  494,  245,  0, 0, 0, 15.5, 123.4, 131),
    ('24.12.2025',  769,  275,  0, 0, 0, 15.7, 124.7, 132),
    ('03.01.2026', 1007,  238,  0, 0, 0, 15.4, 121.1, 133),
    ('14.01.2026', 1261,  264,  1, 0, 0, 15.3, 118.4, 135),
    ('24.01.2026', 1501,  240,  0, 0, 0, 15.1, 122.1, 128),
    ('04.02.2026', 1748,  247,  1, 0, 0, 15.5, 123.7, 131),
    ('14.02.2026', 1994,  246,  0, 0, 0, 16.0, 124.7, 136),
    ('25.02.2026', 2255,  261,  0, 0, 0, 15.6, 113.2, 146),
    ('08.03.2026', 2504,  259,  1, 0, 0, 14.4,  98.8, 150),
    ('18.03.2026', 2755,  251,  0, 0, 0, 14.4, 100.3, 148),
    ('29.03.2026', 2997,  242,  0, 0, 0, 14.5, 101.8, 147),
    ('09.04.2026', 3246,  249,  0, 0, 0, 14.3,  98.0, 150),
    ('20.04.2026', 3494,  248,  0, 0, 0, 14.5, 101.1, 148),
    ('01.05.2026', 3751,  257,  2, 0, 0, 14.3,  97.7, 151),
    ('12.05.2026', 4001,  250,  1, 0, 0, 14.5,  95.1, 158),
    ('23.05.2026', 4258,  257,  2, 0, 0, 14.3,  95.5, 154),
]

oil_rows = []
for i, row in enumerate(oil_data_raw):
    date, mch, nm, fe, cu, si, tan, tbn, soot = row
    row_bg = C['panel'] if i % 2 == 0 else C['altrow']
    fe_col = C['orange'] if fe >= 10 else C['green']
    si_col = C['orange'] if si > 5 else C['text']
    soot_col = C['yellow'] if soot >= 150 else C['text']
    oil_rows.append([
        Pc(C['text'],  date,        size=7, lead=10),
        Pc(C['text'],  str(mch),    size=7, lead=10),
        Pc(C['text'],  str(nm),     size=7, lead=10),
        Pc(fe_col,     str(fe), bold=(fe >= 10), size=7, lead=10),
        Pc(C['text'],  str(cu),     size=7, lead=10),
        Pc(si_col,     str(si),     size=7, lead=10),
        Pc(C['text'],  str(tan),    size=7, lead=10),
        Pc(C['text'],  str(tbn),    size=7, lead=10),
        Pc(soot_col,   str(soot),   size=7, lead=10),
        Pc(C['green'], 'Допустимое', bold=True, size=7, lead=10),
    ])

oil_extra = [('BACKGROUND', (0, i), (-1, i),
              C['panel'] if i % 2 == 0 else C['altrow']) for i in range(len(oil_rows))]
story.append(tbl(oil_rows, oil_w, style_extra=oil_extra, alt=False))
story.append(SP(2))

# TBN и Сажа тренд
story += [
    info_box(
        'TBN ТРЕНД: начало 124.7 (дек.2025) → 95.5 (май 2026). '
        'Снижение на ~23% за 4 258 м/ч при цикле смены ~250 м/ч — ожидаемая деградация. '
        'Порог внимания TBN < 50 — запаса достаточно.',
        bg_color=C['bluebg'], text_color=C['blue'], size=8
    ),
    SP(1),
    info_box(
        'САЖА ТРЕНД: 131 (дек.2025) → 158 (май 2026) мг/кг — постепенный рост. '
        'При норм. сгорании рост сажи ожидаем. Порог внимания ~200 мг/кг. '
        'Контрольная проба через 250 м/ч (≈4 500 м/ч ДВС).',
        bg_color=C['panel2'], text_color=C['sub'], size=8
    ),
    SP(2),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# СТРАНИЦА 4 — СТАРЫЙ ДВС (СПРАВОЧНО)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПРАВКА: СТАРЫЙ ДВС — ВЫЯВЛЕННЫЕ ОТКЛОНЕНИЯ', 4)
story += [SP(1),
    info_box(
        'Старый ДВС заменён в ноябре/декабре 2025 из-за выработки ресурса (~36 680 м/ч). '
        'Записи с М/Ч = 36 xxx–38 xxx относятся к СТАРОМУ ДВС (продолжают мониторинг заменённого агрегата). '
        'Приведены только пробы с отклонениями.',
        bg_color=C['bluebg'], text_color=C['blue'], size=8
    ),
    SP(2),
]

old_cols = ['Дата', 'М/Ч ДВС', 'Fe', 'Cu', 'Si', 'Статус', 'Примечание']
old_w = [PW*0.12, PW*0.11, PW*0.07, PW*0.07, PW*0.07, PW*0.16, PW*0.40]
story.append(hdr_row(old_cols, old_w))

old_data = [
    ('01.05.2025', '32 851', '225', '0',  '—',  'ТМ',  'Fe=225 — высокий. Предвестник отказа.',      C['orangebg'], C['orange']),
    ('25.12.2025', '36 680', ' 41', '0',  '—',  'ТМ',  'Fe=41 — на момент замены ДВС.',              C['orangebg'], C['orange']),
    ('23.02.2026', '36 981', '—',  '0', ' 35',  'ТМ',  'Si=35 мг/кг — возможное загрязнение.',       C['orangebg'], C['orange']),
    ('05.03.2026', '36 980', '—',  '0', '128',  'ТМ',  'Si=128 мг/кг — ВЫСОКИЙ! Прорыв фильтра?',   C['redbg'],    C['red']),
    ('16.05.2026', '38 166', '—',  '0', '110',  'ТМ',  'Si=110 мг/кг — продолжение. Анализ причин.', C['redbg'],    C['red']),
]

old_rows = []
for date, mch, fe, cu, si, status, note, row_bg, status_col in old_data:
    si_val = si.strip()
    si_c = C['red'] if si_val.isdigit() and int(si_val) > 50 else C['orange'] if si_val != '—' else C['sub']
    fe_val = fe.strip()
    fe_c = C['orange'] if fe_val.isdigit() and int(fe_val) > 30 else C['sub']
    old_rows.append([
        Pc(C['text'],   date,   size=7, lead=10),
        Pc(C['sub'],    mch,    size=7, lead=10),
        Pc(fe_c,        fe,     size=7, lead=10),
        Pc(C['text'],   cu,     size=7, lead=10),
        Pc(si_c,        si, bold=(si_val.isdigit() and int(si_val) > 100), size=7, lead=10),
        Pc(status_col,  status, bold=True, size=7, lead=10),
        Pc(C['text'],   note,   size=7, lead=10),
    ])

old_extra = [('BACKGROUND', (0, i), (-1, i), old_data[i][7]) for i in range(len(old_rows))]
story.append(tbl(old_rows, old_w, style_extra=old_extra, alt=False))
story.append(SP(2))

story += [
    info_box(
        'КЛЮЧЕВОЕ НАБЛЮДЕНИЕ: Si=128 и Si=110 мг/кг (март–май 2026) у старого ДВС '
        'указывают на возможный прорыв воздушного фильтра или загрязнение входящего воздуха. '
        'Рекомендуется проверка и замена воздушного фильтра, инспекция впускного тракта.',
        bg_color=C['redbg'], text_color=C['red'], size=8, border_color=C['red']
    ),
    SP(1),
    Pc(C['sub'],
       'ТМ = "Требует принятия мер"  |  Новый ДВС (м/ч 249–4 258) — все показатели в норме.',
       size=7, lead=10),
    SP(3),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# СТРАНИЦА 5 — ИСТОРИЯ ТО + РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ИСТОРИЯ ТЕХНИЧЕСКОГО ОБСЛУЖИВАНИЯ', 5)
story += [SP(1)]

to_cols = ['Дата', 'Длит.', 'М/Ч ДВС', 'Событие']
to_w = [PW*0.12, PW*0.09, PW*0.10, PW*0.69]
story.append(hdr_row(to_cols, to_w))

to_data = [
    ('01.03.2026', '0:15', '2 372', 'Ремонт клавиши стеклоподъёмника',                             C['panel'],    C['sub']),
    ('07.03.2026', '5:05', '2 504', 'ТО-2',                                                         C['bluebg'],   C['blue']),
    ('17.03.2026', '2:05', '2 755', 'ТО-1',                                                         C['panel'],    C['text']),
    ('23.03.2026', '1:42', '2 879', 'Диагностика пожарной системы',                                 C['altrow'],   C['sub']),
    ('28.03.2026', '4:45', '2 997', 'ТО-3',                                                         C['bluebg'],   C['blue']),
    ('08.04.2026', '2:21', '3 246', 'ТО-1. Установка гаек на хомут выхлопной трубы',               C['panel'],    C['text']),
    ('19.04.2026', '3:57', '3 494', 'ТО-2',                                                         C['bluebg'],   C['blue']),
    ('30.04.2026', '2:05', '3 751', 'ТО-1. Замена РВД на смазку шкворня',                          C['panel'],    C['text']),
    ('11.05.2026', '2:59', '4 001', 'ТО-4',                                                         C['bluebg'],   C['blue']),
    ('19.05.2026', '—',    '4 177', 'Оператор разбил стекло водительской двери',                   C['altrow'],   C['sub']),
]

to_rows = []
for date, dur, mch, event, row_bg, event_col in to_data:
    bold = ('ТО-' in event) and (event.strip().startswith('ТО-'))
    to_rows.append([
        Pc(event_col, date,  bold=bold, size=8, lead=11),
        Pc(C['sub'],  dur,              size=8, lead=11),
        Pc(C['sub'],  mch,              size=8, lead=11),
        Pc(C['text'], event, bold=bold, size=8, lead=11),
    ])

to_extra = [('BACKGROUND', (0, i), (-1, i), to_data[i][4]) for i in range(len(to_rows))]
story.append(tbl(to_rows, to_w, style_extra=to_extra, alt=False))
story.append(SP(3))

# ── Рекомендации ────────────────────────────────────────────────────────────
story += section_title('РЕКОМЕНДАЦИИ', 6)
story.append(SP(1))

recs = [
    ('1. КРИТИЧНО',
     'Замена датчика EGT цилиндра 11. Датчик показывает 99°C вместо ожидаемых ~480°C. '
     'Без замены невозможен температурный контроль цилиндра 11.',
     C['red'], C['redbg']),

    ('2. ПРЕДУПРЕЖДЕНИЕ',
     'Дельта EGT = 92°C (без цил.11) — превышает норму 60°C. '
     'Цил.3 (438°C) значительно ниже среднего (~485°C). '
     'Рекомендуется снятие и проверка форсунки цил.3. Возможна проблема ГБЦ цил.3. '
     'Цил.6 (530°C) и цил.2 (525°C) — повышенные, мониторинг.',
     C['orange'], C['orangebg']),

    ('3. ПРЕДУПРЕЖДЕНИЕ',
     'T масла max 107.7°C — наблюдать. Среднее 102.2°C в норме, но пик превышает порог 105°C. '
     'Диагностика масляного теплообменника при следующем плановом ТО.',
     C['orange'], C['orangebg']),

    ('4. НАБЛЮДЕНИЕ',
     'P масла кратковременный min 267 кПа (ниже порога 350 кПа). '
     'Вероятно при пуске или снижении нагрузки. Среднее давление 496 кПа — норма. '
     'Контролировать на следующем DML замере.',
     C['yellow'], C['yellowbg']),

    ('5. ИНФО — МАСЛО ДВС',
     'Нарастание сажи в масле: 131 → 158 мг/кг (дек.2025 → май 2026). '
     'Плановый контроль через 250 м/ч (следующая проба ≈4 500 м/ч ДВС). '
     'TBN 95 — запас достаточный, замена масла по графику.',
     C['blue'], C['bluebg']),

    ('6. ИНФО — СТАРЫЙ ДВС',
     'Анализ причин высокого кремния (Si=128–110 мг/кг, март–май 2026). '
     'Проверка и замена воздушного фильтра старого агрегата. Инспекция впускного тракта.',
     C['sub'], C['panel']),
]

for title, body, tc, bg in recs:
    rec_data = [[
        Pc(tc,        title, bold=True, size=9, lead=13),
        Pc(C['text'], body,             size=8, lead=12),
    ]]
    rec_t = Table(rec_data, colWidths=[PW * 0.22, PW * 0.78])
    rec_t.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), bg),
        ('BOX',          (0, 0), (-1, -1), 0.8, tc),
        ('LEFTPADDING',  (0, 0), (-1, -1), 7),
        ('RIGHTPADDING', (0, 0), (-1, -1), 7),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
        ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
    ]))
    story += [rec_t, SP(1.5)]

story.append(SP(2))

# ── Footer note ────────────────────────────────────────────────────────────────
story.append(HR())
story.append(SP(1))
story.append(Pc(C['sub'],
    'Отчёт подготовлен на основе: DML-20260602 (тест с породой, 1411 строк) · '
    'Спектральный анализ масла ДВС (17 проб, дек.2025–май.2026) · '
    'Сводный анализ NTE200 · История ТО  |  Горная Евразия, 07.06.2026',
    size=7, lead=10))

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
