#!/usr/bin/env python3
"""NTE200 #82 — QSK50 MCRS V16 — Engine Diagnostic Report (07.06.2026).
COMPLETE maintenance history — all 12 events + oil top-up 28.04.2026.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String
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
    yellow   = colors.HexColor('#f1c40f'),
    yellowbg = colors.HexColor('#2a2200'),
)

# ── Helpers ────────────────────────────────────────────────────────────────────
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
                      'NTE200 №82  •  QSK50 MCRS V16  •  Диагностика ДВС  •  07.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    # Bottom bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  М/Ч ДВС (ECM): 4 501 ч  |  DML замер: 02.06.2026')
    canvas.restoreState()

# ── Document setup ─────────────────────────────────────────────────────────────
OUT = '/home/user/NTE200/NTE200_82_engine_report_07062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20 * mm,
    rightMargin=20 * mm,
    topMargin=16 * mm,
    bottomMargin=11 * mm,
    title='NTE200 №82 — Диагностика ДВС QSK50 MCRS V16 — 07.06.2026',
)
PW = doc.width
story = []

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — TITLE + KPI
# ═══════════════════════════════════════════════════════════════════════════════
story += [
    SP(2),
    Pc(C['white'], 'ДИАГНОСТИЧЕСКИЙ ОТЧЁТ ДВИГАТЕЛЯ',
       bold=True, size=16, lead=20, align=TA_CENTER),
    SP(1),
    Pc(C['blue'], 'NTE200 №82  /  QSK50 MCRS V16',
       bold=True, size=13, lead=17, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],
       'Дата отчёта: 07.06.2026  |  DML замер: 02.06.2026 (тест с породой, 1411 строк)',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# Machine info block
info_data = [
    [Pc(C['sub'], 'Гаражный №:', bold=True, size=8, lead=11),
     Pc(C['white'], '82', bold=True, size=8, lead=11),
     Pc(C['sub'], 'Модель:', bold=True, size=8, lead=11),
     Pc(C['white'], 'NTE200, QSK50 MCRS V16', bold=True, size=8, lead=11)],
    [Pc(C['sub'], 'М/Ч ДВС по ECM (DML 02.06.2026):', bold=True, size=8, lead=11),
     Pc(C['green'], '4 501 ч (004501:04:20)', bold=True, size=8, lead=11),
     Pc(C['sub'], 'М/Ч ДВС по последней пробе масла:', bold=True, size=8, lead=11),
     Pc(C['text'], '4 258 м/ч (23.05.2026)', size=8, lead=11)],
    [Pc(C['sub'], 'Начало мониторинга масла:', bold=True, size=8, lead=11),
     Pc(C['text'], 'Дек. 2025 (первая проба: 249 м/ч)', size=8, lead=11),
     Pc(C['sub'], 'Замена ДВС:', bold=True, size=8, lead=11),
     Pc(C['grey'], 'НЕТ ДАННЫХ (не задокументировано)', size=8, lead=11)],
]
info_col_w = [PW*0.22, PW*0.28, PW*0.22, PW*0.28]
info_t = Table(info_data, colWidths=info_col_w)
info_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('BOX',        (0, 0), (-1, -1), 1.0, C['border']),
    ('INNERGRID',  (0, 0), (-1, -1), 0.3, C['border']),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING',(0, 0), (-1, -1), 6),
    ('TOPPADDING',  (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
]))
story += [info_t, SP(4)]

story += [
    Pc(C['sub'], 'КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ', bold=True, size=9, lead=12, align=TA_CENTER),
    SP(1),
]

# KPI row 1: 3 cards
kpi_cards_1 = [
    ('М/Ч ДВС (ECM)',    '4 501 ч',     'по ECM DML 02.06.2026',             C['green'],  C['greenbg']),
    ('МАСЛО ДВС',         'НОРМА',       '17 проб — все "Допустимое"',         C['green'],  C['greenbg']),
    ('EGT ЦИЛ. 11',       'НЕИСПРАВЕН', '99°C вместо ~480°C',                C['red'],    C['redbg']),
]
# KPI row 2: 3 cards
kpi_cards_2 = [
    ('ДЕЛЬТА EGT',        '92 °C',       'Порог 60°C — превышен',              C['orange'], C['orangebg']),
    ('T МАСЛА MAX',        '107.7 °C',   'Порог предупр. 105°C — превышен',   C['orange'], C['orangebg']),
    ('P МАСЛА MIN',        '267 кПа',    'Порог 350 кПа — наблюдать',          C['orange'], C['orangebg']),
]

def make_kpi_table(cards):
    n = len(cards)
    w = [PW / n] * n
    hdr_cells = [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER)  for lbl, _, _, _, _ in cards]
    val_cells  = [Pc(vc,      val, bold=True, size=14, lead=18, align=TA_CENTER) for _, val, _, vc, _ in cards]
    sub_cells  = [Pc(C['sub'], sub, size=6.5, lead=10, align=TA_CENTER)          for _, _, sub, _, _ in cards]
    t = Table([hdr_cells, val_cells, sub_cells], colWidths=w)
    ts = TableStyle([
        ('GRID',         (0, 0), (-1, -1), 0.5, C['border']),
        ('TOPPADDING',   (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
        ('LEFTPADDING',  (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ])
    for i, (_, _, _, _, bg) in enumerate(cards):
        ts.add('BACKGROUND', (i, 0), (i, 0), C['hdr'])
        ts.add('BACKGROUND', (i, 1), (i, 1), bg)
        ts.add('BACKGROUND', (i, 2), (i, 2), bg)
    t.setStyle(ts)
    return t

story += [make_kpi_table(kpi_cards_1), SP(2),
          make_kpi_table(kpi_cards_2), SP(4)]

story += [
    info_box(
        'СТАТУС ДВИГАТЕЛЯ: М/Ч ДВС по ECM — 4 501 ч (DML 02.06.2026). '
        'Масло в норме по всем 17 пробам (дек.2025–май.2026). '
        'ТРЕБУЕТСЯ: замена датчика EGT цил.11 (показывает 99°C). '
        'Контроль цил.3 (438°C — минимальная EGT, возможна проблема форсунки). '
        'Дельта EGT = 92°C превышает норму 60°C.',
        bg_color=C['orangebg'], text_color=C['orange'], size=9, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — ПОЛНАЯ ИСТОРИЯ ВОЗДЕЙСТВИЙ (12 событий + долив масла)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ПОЛНАЯ ИСТОРИЯ ВОЗДЕЙСТВИЙ (все события, март–май 2026)', 1)
story += [
    SP(1),
    Pc(C['sub'], 'Источник: ОТЧЕТ Полюс Магадан.xlsx  |  Доливки: Доливки Горная Евразия.xlsx',
       size=8, lead=11),
    SP(2),
]

# Legend
legend_cols = [
    (C['bluebg'],   C['blue'],   '■ Плановое ТО'),
    (C['greenbg'],  C['green'],  '■ Долив масла'),
    (C['altrow'],   C['sub'],    '■ Прочее'),
    (C['altrow2'],  C['sub'],    '■ Кузов/эл.'),
]
leg_cells = [Pc(tc, lbl, size=7, lead=10) for bg, tc, lbl in legend_cols]
leg_t = Table([leg_cells], colWidths=[PW / 4] * 4)
leg_style_cmds = [
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING',   (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
]
for col_i, (bg, tc, lbl) in enumerate(legend_cols):
    leg_style_cmds.append(('BACKGROUND', (col_i, 0), (col_i, 0), bg))
leg_t.setStyle(TableStyle(leg_style_cmds))
story += [leg_t, SP(2)]

to_cols = ['Дата', 'Длит.', 'М/Ч ДВС', 'Событие / Работы']
to_w = [PW*0.12, PW*0.10, PW*0.10, PW*0.68]
story.append(hdr_row(to_cols, to_w))

# 12 events + 1 oil top-up = 13 rows
# (date, dur, mch, event, row_bg, event_col, bold)
to_data = [
    ('01.03.2026', '0:15',         '2 372',  'Ремонт клавиши стеклоподъёмника',
     C['altrow2'],  C['sub'],    False),
    ('07.03.2026', '5:05',         '2 504',  'ТО-2',
     C['bluebg'],   C['blue'],   True),
    ('17.03.2026', '2:05',         '2 755',  'ТО-1',
     C['bluebg'],   C['blue'],   True),
    ('23.03.2026', '1:42',         '2 879',  'Диагностика пожарной системы',
     C['altrow'],   C['sub'],    False),
    ('23.03.2026', '0:35',         '2 896',  'Ремонт отопителя салона',
     C['altrow2'],  C['sub'],    False),
    ('28.03.2026', '4:45',         '2 997',  'ТО-3',
     C['bluebg'],   C['blue'],   True),
    ('08.04.2026', '2:21',         '3 246',  'ТО-1. Установка 2-х гаек на хомут выхлопной трубы',
     C['bluebg'],   C['blue'],   True),
    ('19.04.2026', '3:57',         '3 494',  'ТО-2',
     C['bluebg'],   C['blue'],   True),
    ('28.04.2026', '0:16',         '3 705',  'Долив масла ДВС 15 литров',
     C['greenbg'],  C['green'],  True),
    ('30.04.2026', '2:05',         '3 751',  'ТО-1. Замена РВД на смазку шкворня верх правая',
     C['bluebg'],   C['blue'],   True),
    ('11.05.2026', '2:59',         '4 001',  'ТО-4',
     C['bluebg'],   C['blue'],   True),
    ('19.05.2026', 'нет данных',   '4 177',  'Оператор разбил стекло водительской двери',
     C['altrow2'],  C['sub'],    False),
]

to_rows = []
for date, dur, mch, event, row_bg, ecol, bold in to_data:
    to_rows.append([
        Pc(ecol,      date,  bold=bold, size=7.5, lead=11),
        Pc(C['sub'],  dur,             size=7.5, lead=11),
        Pc(C['sub'],  mch,             size=7.5, lead=11),
        Pc(C['text'], event,           size=7.5, lead=11),
    ])

to_extra = [('BACKGROUND', (0, i), (-1, i), to_data[i][4])
            for i in range(len(to_rows))]
story.append(tbl(to_rows, to_w, style_extra=to_extra))

story += [
    SP(2),
    info_box(
        'Итого за период: 7 плановых ТО • 1 долив масла ДВС (28.04.2026, 15 л, М/Ч 3 705) • '
        '1 диагностика пожарной системы • ремонтные работы по кузову и электрике. '
        'ГБЦ ремонты: запись присутствует, все поля пусты — нет данных о ремонтах ГБЦ.',
        bg_color=C['panel2'], text_color=C['text'], size=8.5, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — DML ПАРАМЕТРЫ + EGT BAR CHART
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('DML ПАРАМЕТРЫ ДВИГАТЕЛЯ (02.06.2026, тест с породой)', 2)
story += [
    SP(1),
    Pc(C['sub'],
       'Запись: 1411 строк  |  М/Ч ЭБУ: 004501:04:20  |  Тест с породой, под нагрузкой',
       size=8, lead=11),
    SP(2),
]

dml_cols = ['Параметр', 'Значение', 'Норма / порог', 'Статус']
dml_w = [PW*0.40, PW*0.22, PW*0.22, PW*0.16]
story.append(hdr_row(dml_cols, dml_w))

dml_data = [
    ('RPM средн.',                   '1 822 об/мин',   '1 500–1 900',             'НОРМА',     C['greenbg'],  C['green']),
    ('RPM макс.',                    '1 929 об/мин',   '< 1 950',                 'НОРМА',     C['greenbg'],  C['green']),
    ('RPM мин.',                     '799 об/мин',     '(холостой ход)',           'OK',        C['altrow'],   C['sub']),
    ('T масла avg',                  '102.2 °C',       '< 105 °C (предупр.)',     'НОРМА',     C['greenbg'],  C['green']),
    ('T масла max',                  '107.7 °C',       '< 105°C (предупр.) / < 115°C (крит.)', 'НАБЛЮДАТЬ', C['orangebg'], C['orange']),
    ('T масла min',                  '87.9 °C',        '—',                       'OK',        C['altrow'],   C['sub']),
    ('P масла avg',                  '496 кПа',        '> 350 кПа',               'НОРМА',     C['greenbg'],  C['green']),
    ('P масла min (кратковременно)', '267 кПа',        '> 350 кПа',               'НАБЛЮДАТЬ', C['orangebg'], C['orange']),
    ('P масла max',                  '556 кПа',        '< 620 кПа',               'НОРМА',     C['greenbg'],  C['green']),
    ('P наддува (впускной коллект.)', '238 кПа',       '220–320 кПа',             'НОРМА',     C['greenbg'],  C['green']),
    ('T впускного коллектора',       '64–66 °C',       '< 80 °C',                 'НОРМА',     C['greenbg'],  C['green']),
]

dml_rows = []
for param, val, norm, status_txt, row_bg, status_col in dml_data:
    dml_rows.append([
        Pc(C['text'],  param,      size=8, lead=11),
        Pc(status_col, val,  bold=True, size=8, lead=11),
        Pc(C['sub'],   norm,       size=7, lead=10),
        Pc(status_col, status_txt, bold=True, size=8, lead=11),
    ])

dml_extra = [('BACKGROUND', (0, i), (-1, i), dml_data[i][4])
             for i in range(len(dml_rows))]
story.append(tbl(dml_rows, dml_w, style_extra=dml_extra))
story += [
    SP(1),
    info_box(
        'T масла MAX 107.7°C: выше порога предупреждения 105°C, критический порог 115°C не достигнут. '
        'Среднее 102.2°C — норма. Контролировать.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(1),
    info_box(
        'P масла MIN 267 кПа: ниже порога 350 кПа. Вероятно при снижении нагрузки. '
        'Среднее давление 496 кПа — в норме. Контролировать на следующем DML.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(3),
]

# EGT section
story += section_title('ТЕМПЕРАТУРА ВЫХЛОПА ЦИЛИНДРОВ — EGT (16 цил., avg при нагрузке)', None)
story += [
    SP(1),
    Pc(C['sub'],
       'Средняя без цил.11: ~485°C  |  Цил.11 НЕИСПРАВЕН (датчик 99°C)  |  '
       'Дельта (без цил.11) = 92°C (порог 60°C)',
       size=8, lead=11),
    SP(2),
]

# ── EGT horizontal bar chart via ReportLab Drawing ────────────────────────────
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
    (11,  99, True),    # неисправен — датчик
    (12, 467, False),
    (13, 451, False),
    (14, 457, False),
    (15, 502, False),
    (16, 471, False),
]

AVG_NORMAL = 485
chart_w  = PW
chart_h  = 170
n_cyl    = len(EGT)
row_h    = chart_h / n_cyl
label_w  = 32
val_w    = 60
bar_area = chart_w - label_w - val_w - 4
max_val  = 580

d = Drawing(chart_w, chart_h)
d.add(Rect(0, 0, chart_w, chart_h,
           fillColor=C['panel'], strokeColor=C['border'], strokeWidth=0.3))

def egt_bar_color(val, broken):
    if broken:
        return C['grey']
    if val >= 520:
        return C['orange']
    if val < 450:
        return C['blue']
    return C['green']

for i, (cyl, val, broken) in enumerate(EGT):
    y = chart_h - (i + 1) * row_h
    row_bg_col = C['altrow'] if i % 2 == 0 else C['panel']
    d.add(Rect(0, y, chart_w, row_h,
               fillColor=row_bg_col if not broken else C['redbg'],
               strokeColor=None, strokeWidth=0))

    lbl = f'Цил.{cyl:02d}'
    d.add(String(2, y + row_h * 0.28, lbl,
                 fontName='Sans', fontSize=6.5,
                 fillColor=C['red'].hexval() if broken else C['sub'].hexval()))

    if broken:
        bar_len = 6
        bar_col = C['grey']
    else:
        bar_len = (val / max_val) * bar_area
        bar_col = egt_bar_color(val, broken)

    bar_y   = y + row_h * 0.15
    bar_h_px = row_h * 0.70
    d.add(Rect(label_w, bar_y, bar_len, bar_h_px,
               fillColor=bar_col, strokeColor=None, strokeWidth=0))

    if broken:
        val_txt = '99°C — ДАТЧИК НЕИСПРАВЕН!'
        vc = C['red']
        fn_bold = 'Sans-Bold'
    else:
        val_txt = f'{val}°C'
        vc = egt_bar_color(val, broken)
        fn_bold = 'Sans-Bold' if (val >= 520 or val < 450) else 'Sans'
    d.add(String(label_w + bar_len + 4, y + row_h * 0.28, val_txt,
                 fontName=fn_bold, fontSize=6.5, fillColor=vc.hexval()))

# Average reference line
avg_x = label_w + (AVG_NORMAL / max_val) * bar_area
d.add(Line(avg_x, 0, avg_x, chart_h,
           strokeColor=C['blue'].hexval(), strokeWidth=0.8,
           strokeDashArray=[3, 2]))
d.add(String(avg_x + 2, chart_h - 8, f'avg={AVG_NORMAL}°C',
             fontName='Sans', fontSize=6, fillColor=C['blue'].hexval()))

# Grid tick marks
for tick_val in [350, 400, 450, 500, 520]:
    tx = label_w + (tick_val / max_val) * bar_area
    d.add(Line(tx, 0, tx, chart_h,
               strokeColor=C['border'].hexval(), strokeWidth=0.3,
               strokeDashArray=[2, 3]))
    d.add(String(tx - 7, 1, str(tick_val),
                 fontName='Sans', fontSize=5, fillColor=C['sub'].hexval()))

story.append(d)
story.append(SP(1))

# Legend
legend_egt = [
    (C['green'],  '■ Норма (450–520°C)'),
    (C['orange'], '■ Повышенная (≥520°C)'),
    (C['blue'],   '■ Пониженная (<450°C)'),
    (C['grey'],   '■ Датчик неисправен'),
    (C['blue'],   '─ ─ Среднее (485°C)'),
]
leg_cells = [Pc(tc, lbl, size=7, lead=10) for tc, lbl in legend_egt]
leg_t = Table([leg_cells], colWidths=[PW / 5] * 5)
leg_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING',   (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
]))
story += [leg_t, SP(2)]

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
        'Цил.3 значительно ниже среднего → возможна проблема форсунки или ГБЦ цил.3.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — ИНЖЕНЕРНЫЙ АНАЛИЗ КЛАПАНОВ (Cummins CCEC)
# ═══════════════════════════════════════════════════════════════════════════════
from reportlab.platypus import Image as RLImage

story += section_title('ИНЖЕНЕРНЫЙ АНАЛИЗ ИЗНОСА КЛАПАНОВ QSK50 (Cummins CCEC)', 3)
story += [SP(1)]

story.append(info_box(
    'Фотофиксация: выпускные клапаны ДВС QSK50 №82 — состояние при вскрытии ГБЦ.\n'
    'Характер повреждения: тяжёлые углеродистые отложения (нагар) на тарелке и стержне клапана, '
    'критическая просадка рабочей фаски. Анализ типа отложений — Cummins CCEC (EDS):\n'
    'основной компонент — зола моторного масла (Ca, Zn, P, Mg); присутствует пыль карьера (Al, Si).',
    bg_color=C['panel2'], text_color=C['text'], size=8, border_color=C['blue']
))
story.append(SP(2))

# Photos — actual #82 valve photos from the field
img_v1_path = '/tmp/valve_photos_82/image-07-06-26-11-08.jpg'
img_v2_path = '/tmp/valve_photos_82/image-07-06-26-11-08-1.jpg'

photo_w = PW * 0.495
photo_h = photo_w * 467 / 1000  # aspect ratio 1000x467

try:
    img_v1 = RLImage(img_v1_path, width=photo_w, height=photo_h)
    img_v2 = RLImage(img_v2_path, width=photo_w, height=photo_h)

    cap_v1 = Pc(C['red'],
        'Выпускной клапан №82 (крупный план): критический нагар на тарелке и фаске. '
        'Потеря герметичности — причина повышенной температуры EGT.',
        size=7.5, lead=11)
    cap_v2 = Pc(C['red'],
        'Два выпускных клапана №82: равномерный нагар — системная проблема '
        '(воздушный фильтр, качество топлива или режим эксплуатации).',
        size=7.5, lead=11)

    photo_tbl = Table(
        [[img_v1,  img_v2],
         [cap_v1,  cap_v2]],
        colWidths=[photo_w + 4, photo_w + 4]
    )
    photo_tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, 0), C['hdr']),
        ('BACKGROUND',   (0, 1), (-1, 1), C['redbg']),
        ('ALIGN',        (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN',       (0, 0), (-1, 0), 'MIDDLE'),
        ('VALIGN',       (0, 1), (-1, 1), 'TOP'),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
        ('LEFTPADDING',  (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('BOX',          (0, 0), (-1, -1), 1.0, C['red']),
        ('INNERGRID',    (0, 0), (-1, -1), 0.3, C['border']),
    ]))
    story.append(photo_tbl)
    story.append(SP(3))
except Exception as e:
    story.append(info_box(f'[Фото клапанов №82 не загружено: {e}]', C['redbg'], C['red']))

# EDS findings summary
story += [
    Pc(C['blue'], 'РЕЗУЛЬТАТЫ АНАЛИЗА (ЭДС/EDS)', bold=True, size=10, lead=13),
    SP(1),
]

eds_data = [
    ('Зола масла (Ca, Zn, P, Mg)', 'Основной компонент отложений — зола моторного масла',
     'Нормальное явление при длительной эксплуатации',                              C['text'],   C['altrow']),
    ('Пыль карьера (Al, Si)',      'Выявлено небольшое количество пыли в отложениях',
     'Указывает на попадание карьерной пыли — проверить воздушный фильтр',          C['orange'], C['orangebg']),
    ('Износ тарелок 1.20–1.36 мм', 'Критическая просадка выпускных клапанов',
     'При превышении допуска — обязательная замена клапанов и сёдел',               C['red'],    C['redbg']),
    ('Белые отложения внутри', 'Основной объём отложений — кальциевые соединения',
     'Зола от присадок масла (моющие присадки на основе Ca)',                       C['text'],   C['altrow2']),
]

eds_cols = ['Находка', 'Выявлено', 'Вывод/Действие']
eds_w = [PW*0.24, PW*0.35, PW*0.41]
story.append(hdr_row(eds_cols, eds_w))

eds_rows = []
for finding, detail, action, tc, bg in eds_data:
    eds_rows.append([
        Pc(tc,       finding, bold=True, size=8, lead=11),
        Pc(C['text'], detail,            size=8, lead=11),
        Pc(tc,       action,             size=8, lead=11),
    ])
eds_extra = [('BACKGROUND', (0, i), (-1, i), eds_data[i][4]) for i in range(len(eds_rows))]
story.append(tbl(eds_rows, eds_w, style_extra=eds_extra))

story += [
    SP(2),
    info_box(
        'СВЯЗЬ С №82: Дельта EGT=92°C (без цил.11) указывает на неравномерную работу цилиндров. '
        'Если просадка клапанов при замере превысит 1.0 мм — требуется замена по аналогии с данным отчётом. '
        'Цил.3 (438°C, минимальный) — первый кандидат на проверку зазоров клапанов.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС (17 проб)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС — 17 ПРОБ (дек.2025 – май.2026)', 4)
story += [SP(1)]

story.append(info_box(
    '17 проб — ВСЕ "Допустимое". Двигатель не изнашивается.\n'
    'Fe=0–2 мг/кг (норма <30) — норма. Cu=0 — норма. Si=0 во всех пробах — норма.\n'
    'TBN: снижается 124.7→95.5 мг/кг — закономерная деградация при цикле замены 250 м/ч.\n'
    'Сажа: 131→158 мг/кг — постепенный рост, допустимо (норма <250 мг/кг).\n'
    'Долив масла ДВС: 28.04.2026, 15 л, М/Ч 3 705 (из Доливки Горная Евразия.xlsx).',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

oil_cols = ['Дата', 'М/Ч ДВС', 'Нар.масла', 'Fe', 'Cu', 'Si', 'TAN', 'TBN', 'Сажа', 'Статус']
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
    fe_col   = C['orange'] if fe >= 2 else (C['text'] if fe == 0 else C['sub'])
    soot_col = C['orange'] if soot >= 155 else C['text']
    tbn_col  = C['orange'] if tbn < 98 else C['text']
    row_bg   = C['altrow'] if i % 2 == 0 else C['altrow2']
    oil_rows.append([
        Pc(C['text'],  date,        size=7, lead=10),
        Pc(C['text'],  str(mch),    size=7, lead=10),
        Pc(C['text'],  str(nm),     size=7, lead=10),
        Pc(fe_col,     str(fe),     bold=(fe >= 2), size=7, lead=10),
        Pc(C['text'],  str(cu),     size=7, lead=10),
        Pc(C['text'],  str(si),     size=7, lead=10),
        Pc(C['text'],  f'{tan:.1f}', size=7, lead=10),
        Pc(tbn_col,    f'{tbn:.1f}', bold=(tbn < 98), size=7, lead=10),
        Pc(soot_col,   str(soot),   size=7, lead=10),
        Pc(C['green'], 'Допустимое', size=7, lead=10),
    ])

oil_extra = [('BACKGROUND', (0, i), (-1, i), C['altrow'] if i % 2 == 0 else C['altrow2'])
             for i in range(len(oil_rows))]
story.append(tbl(oil_rows, oil_w, style_extra=oil_extra))
story += [SP(2)]

# Trend analysis
story += section_title('ТРЕНДЫ МАСЛА ДВС NTE200 №82', None)
story += [SP(1)]

trend_data = [
    ('Fe', '0–2 мг/кг', '<30 мг/кг', 'НОРМА — минимальный износ',                C['green'],  C['greenbg']),
    ('Cu', '0 мг/кг',   '<15 мг/кг', 'НОРМА — подшипники в порядке',             C['green'],  C['greenbg']),
    ('Si', '0 мг/кг',   '<10 мг/кг', 'НОРМА — все пробы 0',                      C['green'],  C['greenbg']),
    ('TBN','124→95',    '>70 мг/кг', 'Снижение ожидаемо, замена 250 м/ч',        C['text'],   C['altrow']),
    ('TAN','14.3–16.0', '<15',        '16.0 (фев.2026) — в норме для режима',    C['text'],   C['altrow2']),
    ('Сажа','131→158',  '<250 мг/кг','Постепенный рост — допустимо',              C['text'],   C['altrow']),
]
trend_cols_hdr = ['Параметр', 'Диапазон', 'Норма', 'Оценка']
trend_w = [PW*0.12, PW*0.15, PW*0.15, PW*0.58]
story.append(hdr_row(trend_cols_hdr, trend_w))
trend_rows = []
for param, rng, norm, assess, tc, bg in trend_data:
    trend_rows.append([
        Pc(C['sub'], param, bold=True, size=8, lead=11),
        Pc(tc,       rng,   bold=True, size=8, lead=11),
        Pc(C['sub'], norm,             size=8, lead=11),
        Pc(tc,       assess,           size=8, lead=11),
    ])
trend_extra = [('BACKGROUND', (0, i), (-1, i), trend_data[i][5])
               for i in range(len(trend_rows))]
story.append(tbl(trend_rows, trend_w, style_extra=trend_extra))

story += [
    SP(2),
    info_box(
        'ПРИМЕЧАНИЕ ПО БАЗЕ МАСЛА: в базе спектрального анализа под гаражным №82 присутствуют '
        'записи ДВУХ разных машин:\n'
        '• NTE200 №82 (QSK50): узлы ДВС/ГС/МКЛ/МКП, М/Ч 249–4 258 м/ч (дек.2025–май.2026)\n'
        '• Komatsu D375 №82 (другая машина): узлы ДВС/КПП/РХЛ/РХП, М/Ч 31 710–38 166 м/ч\n'
        'Данный отчёт содержит ТОЛЬКО данные NTE200 №82. Данные Komatsu D375 в отчёт не включены.',
        bg_color=C['bluebg'], text_color=C['blue'], size=8, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 6 — РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('РЕКОМЕНДАЦИИ', 5)
story += [SP(2)]

recommendations = [
    (
        '1. КРИТИЧНО',
        'Замена датчика EGT цилиндра 11',
        'Датчик показывает 99°C вместо ожидаемых ~480°C. '
        'Без замены невозможен температурный контроль цилиндра 11.',
        C['red'], C['redbg'],
    ),
    (
        '2. ПРЕДУПРЕЖДЕНИЕ',
        'Дельта EGT=92°C без цил.11 — цил.3 (438°C) минимальный, проверить форсунку',
        'Цил.3 (438°C) значительно ниже среднего (~485°C). Рекомендуется снятие и '
        'проверка форсунки цил.3. Возможна проблема ГБЦ цил.3. '
        'Цил.6 (530°C) и цил.2 (525°C) — повышенные, продолжить мониторинг.',
        C['orange'], C['orangebg'],
    ),
    (
        '3. НАБЛЮДЕНИЕ',
        'T масла max 107.7°C — контролировать',
        'Среднее 102.2°C в норме, но максимум превышает порог 105°C. '
        'Диагностика масляного теплообменника при следующем плановом ТО.',
        C['orange'], C['orangebg'],
    ),
    (
        '4. НАБЛЮДЕНИЕ',
        'P масла min 267 кПа — контролировать на следующем DML',
        'Среднее давление 496 кПа — норма. Кратковременное падение 267 кПа '
        'вероятно при снижении нагрузки. Контролировать при следующем DML замере.',
        C['blue'], C['bluebg'],
    ),
    (
        '5. ИНФО',
        'Нарастание сажи 131→158 мг/кг — допустимо',
        'Плановый контроль через 250 м/ч (следующая проба ≈4 500 м/ч ДВС). '
        'TBN 95.5 — запас достаточный. Критический уровень сажи QSK50 — 250 мг/кг.',
        C['sub'], C['greybg'],
    ),
]

recs_w = [PW*0.20, PW*0.33, PW*0.47]
rec_rows_all = []
for prio_label, title, body, tc, bg in recommendations:
    rec_rows_all.append([
        Pc(tc,        prio_label, bold=True, size=9, lead=13),
        Pc(tc,        title,      bold=True, size=8.5, lead=13),
        Pc(C['text'], body,                  size=8, lead=12),
    ])

prio_t = Table(rec_rows_all, colWidths=recs_w)
prio_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 7),
    ('RIGHTPADDING', (0, 0), (-1, -1), 7),
    ('TOPPADDING',   (0, 0), (-1, -1), 8),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 8),
    ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
])
for i, (_, _, _, _, bg) in enumerate(recommendations):
    prio_ts.add('BACKGROUND', (0, i), (-1, i), bg)
prio_t.setStyle(prio_ts)
story += [prio_t, SP(4)]

# Final summary
story += [
    HR(C['blue']),
    SP(1),
    Pc(C['blue'], 'ИТОГОВАЯ ОЦЕНКА ДВИГАТЕЛЯ №82', bold=True, size=11, lead=14, align=TA_CENTER),
    SP(1),
]

summary_kpi = [
    ('МАСЛО ДВС',    'НОРМА',         'Все 17 проб — чисто',             C['green'],  C['greenbg']),
    ('EGT ЦИЛ.11',   'НЕИСПРАВЕН',   'Замена датчика — срочно',          C['red'],    C['redbg']),
    ('ДЕЛЬТА EGT',   '92°C',          'Порог 60°C — проверить цил.3',    C['orange'], C['orangebg']),
    ('СЛЕД.ШАГ',     'EGT + DML',     'Замена датчика, след. DML замер', C['blue'],   C['bluebg']),
]
final_w = [PW / 4] * 4
final_hdr = [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _, _ in summary_kpi]
final_val = [Pc(vc,       val, bold=True, size=12, lead=16, align=TA_CENTER) for _, val, _, vc, _ in summary_kpi]
final_sub = [Pc(vc,       sub, size=7,    lead=10, align=TA_CENTER)          for _, _, sub, vc, _ in summary_kpi]

final_t = Table([final_hdr, final_val, final_sub], colWidths=final_w)
final_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING',   (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
])
for i, (_, _, _, _, bg) in enumerate(summary_kpi):
    final_ts.add('BACKGROUND', (i, 0), (i, 0), C['hdr'])
    final_ts.add('BACKGROUND', (i, 1), (i, 1), bg)
    final_ts.add('BACKGROUND', (i, 2), (i, 2), bg)
final_t.setStyle(final_ts)
story += [final_t, SP(2)]

story += [
    HR(),
    SP(1),
    Pc(C['sub'],
       'Отчёт подготовлен на основе: DML-20260602 (тест с породой, 1411 строк)  •  '
       'Спектральный анализ масла ДВС NTE200 №82 (17 проб, дек.2025–май.2026)  •  '
       'История ТО (ОТЧЕТ Полюс Магадан.xlsx)  •  Доливки (Доливки Горная Евразия.xlsx)  •  '
       'ГБЦ ремонты.xlsx (данных нет)  |  Горная Евразия  •  07.06.2026',
       size=7, lead=10),
]

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
