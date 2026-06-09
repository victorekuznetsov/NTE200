#!/usr/bin/env python3
"""NTE200 #88 — QSK50 MCRS V16 — Engine Diagnostic Report (08.06.2026)."""

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
    yellow   = colors.HexColor('#f1c40f'),
    yellowbg = colors.HexColor('#2a2200'),
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
                      'NTE200 №88  •  QSK50 MCRS V16  •  Диагностика ДВС  •  08.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  М/Ч ДВС (ECM): 3 872 ч  |  DML замер: 08.06.2026')
    canvas.restoreState()

OUT = '/home/user/NTE200/NTE200_88_engine_report_08062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20 * mm,
    rightMargin=20 * mm,
    topMargin=16 * mm,
    bottomMargin=11 * mm,
    title='NTE200 №88 — Диагностика ДВС QSK50 MCRS V16 — 08.06.2026',
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
    Pc(C['blue'], 'NTE200 №88  /  QSK50 MCRS V16',
       bold=True, size=13, lead=17, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],
       'Дата отчёта: 08.06.2026  |  DML замер: 08.06.2026, 09:15 (1183 строки, 217 параметров)',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

info_data = [
    [Pc(C['sub'], 'Гаражный №:', bold=True, size=8, lead=11),
     Pc(C['white'], '88', bold=True, size=8, lead=11),
     Pc(C['sub'], 'Модель:', bold=True, size=8, lead=11),
     Pc(C['white'], 'NTE200, QSK50 MCRS V16', bold=True, size=8, lead=11)],
    [Pc(C['sub'], 'ESN:', bold=True, size=8, lead=11),
     Pc(C['text'], '33238554', size=8, lead=11),
     Pc(C['sub'], 'ECM код:', bold=True, size=8, lead=11),
     Pc(C['text'], 'AQ60809.08 / AR60193.16 / AR60194.15', size=8, lead=11)],
    [Pc(C['sub'], 'М/Ч ДВС по ECM (DML 08.06.2026):', bold=True, size=8, lead=11),
     Pc(C['green'], '3 872 ч (ECM1/ECM144: 003872)', bold=True, size=8, lead=11),
     Pc(C['sub'], 'DML дата / длит.:', bold=True, size=8, lead=11),
     Pc(C['text'], '08.06.2026, 09:15  |  20 мин', size=8, lead=11)],
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

kpi_cards_1 = [
    ('М/Ч ДВС (ECM1)',  '3 872 ч',       'по ECM DML 08.06.2026',                            C['green'],  C['greenbg']),
    ('МАСЛО ДВС',        'НОРМА',         '14 проб — все \"Допустимое\"',          C['green'],  C['greenbg']),
    ('EGT ЦИЛ.15',      '549/565°C',     'КРИТИЧНО: max 565°C &gt; 560°C',         C['red'],    C['redbg']),
]
kpi_cards_2 = [
    ('Δ EGT',           '60°C',           'Порог 60°C — на границе',             C['orange'], C['orangebg']),
    ('T МАСЛА MAX',     '107.7°C',        'Порог предупр. 105°C — превышен',     C['orange'], C['orangebg']),
    ('ECM КОД 611',     '10 раз',         'Горячий останов — неактивен',         C['orange'], C['orangebg']),
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
        'СТАТУС ДВИГАТЕЛЯ: М/Ч ДВС по ECM — 3 872 ч (DML 08.06.2026). '
        'Масло в норме по всем 14 пробам (янв.–май 2026). '
        'ТРЕБУЕТСЯ: диагностика форсунки/ГБЦ цил.15 (max 565°C) и цил.6 (max 561°C) — '
        'оба превысили порог 560°C. Δ EGT=60°C — на границе порога. '
        'T масла max=107.7°C превышает порог 105°C — контроль теплообменника.',
        bg_color=C['redbg'], text_color=C['red'], size=9, border_color=C['red']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — ИСТОРИЯ ВОЗДЕЙСТВИЙ (14 событий, март–май 2026)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ИСТОРИЯ ВОЗДЕЙСТВИЙ (все события, март–май 2026)', 1)
story += [
    SP(1),
    Pc(C['sub'], 'Источник: ОТЧЕТ Полюс Магадан.хлсх  |  14 событий за период',
       size=8, lead=11),
    SP(2),
]

legend_cols = [
    (C['bluebg'],    C['blue'],    '■ Плановое ТО'),
    (C['orangebg'],  C['orange'],  '■ EGT / проводка'),
    (C['greybg'],    C['sub'],     '■ Прочее'),
]
leg_cells = [Pc(tc, lbl, size=7, lead=10) for bg, tc, lbl in legend_cols]
leg_t = Table([leg_cells], colWidths=[PW / 3] * 3)
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

to_data = [
    ('02.03.2026', '4:30',  '~1 608', 'Сварочные работы топливного бака',
     C['greybg'],    C['sub'],    False),
    ('09.03.2026', '2:59',  '1 748',  'ТО-1',
     C['bluebg'],    C['blue'],   True),
    ('13.03.2026', '0:20',  '~1 848', 'Протяжка селектора хода. Диагностика ДВС',
     C['greybg'],    C['sub'],    False),
    ('19.03.2026', '4:50',  '1 998',  'ТО-4',
     C['bluebg'],    C['blue'],   True),
    ('24.03.2026', '0:35',  '~2 101', 'Ремонт эл. проводки датчика отработавших газов (EGT)',
     C['orangebg'],  C['orange'], True),
    ('30.03.2026', '1:52',  '2 250',  'ТО-1',
     C['bluebg'],    C['blue'],   True),
    ('08.04.2026', '0:08',  '~2 460', 'Подтекание антифриза — протяжка хомутов системы охлаждения',
     C['greybg'],    C['sub'],    False),
    ('10.04.2026', '3:39',  '2 498',  'ТО-2',
     C['bluebg'],    C['blue'],   True),
    ('15.04.2026', '0:12',  '~2 622', 'Забор проб из РМК (редуктора)',
     C['greybg'],    C['sub'],    False),
    ('18.04.2026', '0:29',  '~2 696', 'Ошибка ECM 33/1 — залипание контакта RP реверса, диагностика GE, сброс ошибок',
     C['orangebg'],  C['orange'], False),
    ('20.04.2026', '2:36',  '~2 730', 'Диагностика ПГП (пневмогидропривод подвески)',
     C['greybg'],    C['sub'],    False),
    ('21.04.2026', '3:11',  '2 744',  'ТО-1',
     C['bluebg'],    C['blue'],   True),
    ('01.05.2026', '2:19',  '2 995',  'ТО-3',
     C['bluebg'],    C['blue'],   True),
    ('12.05.2026', '2:45',  '3 249',  'ТО-1. Замена заправочного клапана передней левой стойки ПГП, прокачка стоек',
     C['bluebg'],    C['blue'],   True),
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
        'Итого за период: 6 плановых ТО (ТО-1×4, ТО-2×1, ТО-3×1, ТО-4×1) • '
        '1 ремонт проводки EGT (24.03.2026) • 1 ошибка ECM 33/1 (18.04.2026) • '
        '2 прочих (подтекание антифриза, диагностика ПГП) • 1 сварка бака. '
        'Доливки масла: нет данных. ГБЦ ремонты: нет данных.',
        bg_color=C['panel2'], text_color=C['text'], size=8.5, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — DML ПАРАМЕТРЫ + EGT BAR CHART
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('DML ПАРАМЕТРЫ ДВИГАТЕЛЯ (08.06.2026, под нагрузкой)', 2)
story += [
    SP(1),
    Pc(C['sub'],
       'Запись: 1183 строки  |  217 параметров  |  М/Ч ЭБУ ECM1: 003872  |  RPM avg=1854 (под нагрузкой)',
       size=8, lead=11),
    SP(2),
]

dml_cols = ['Параметр', 'Значение', 'Норма / порог', 'Статус']
dml_w = [PW*0.40, PW*0.22, PW*0.22, PW*0.16]
story.append(hdr_row(dml_cols, dml_w))

dml_data = [
    ('RPM средн. (под нагрузкой)',         '1 854 об/мин',   '1 500–1 900',              'НОРМА',     C['greenbg'],  C['green']),
    ('Относит. нагрузка avg',              '91.8%',           '—',                        'НОРМА',     C['greenbg'],  C['green']),
    ('Относит. нагрузка max',              '100%',            '—',                        'НОРМА',     C['greenbg'],  C['green']),
    ('Расход топлива avg (управляемый)',    '401.9 л/ч',       '—',                        'OK',        C['altrow'],   C['sub']),
    ('Расход топлива max',                 '481.2 л/ч',       '—',                        'OK',        C['altrow2'],  C['sub']),
    ('T охл. (Датчик температуры) avg',    '83.9°C',          '&lt; 95°C',                   'НОРМА',     C['greenbg'],  C['green']),
    ('T охл. max',                         '85.3°C',          '&lt; 95°C',                   'НОРМА',     C['greenbg'],  C['green']),
    ('T масла avg',                        '102.4°C',         '&lt; 105°C (предупр.)',        'НОРМА',     C['greenbg'],  C['green']),
    ('T масла max',                        '107.7°C',         '&lt; 105°C (предупр.) / &lt; 115°C (крит.)', 'НАБЛЮДАТь', C['orangebg'], C['orange']),
    ('P масла avg',                        '508 кПа',         '&gt; 350 кПа',                'НОРМА',     C['greenbg'],  C['green']),
    ('P масла min',                        '380 кПа',         '&gt; 350 кПа',                'НОРМА',     C['greenbg'],  C['green']),
    ('P наддув (впускной коллект.) avg',   '240 кПа',         '220–320 кПа',              'НОРМА',     C['greenbg'],  C['green']),
    ('P наддув max',                       '273 кПа',         '220–320 кПа',              'НОРМА',     C['greenbg'],  C['green']),
    ('T впускного коллектора avg',         '68.8°C',          '&lt; 80°C',                   'НОРМА',     C['greenbg'],  C['green']),
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
        'Среднее 102.4°C — норма. Рекомендуется диагностика масляного теплообменника.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(3),
]

story += section_title('ТЕМПЕРАТУРА ВЫХЛОПА ЦИЛИНДРОВ — EGT (16 цил., avg при RPM&gt;1700, 1066 строк)', None)
story += [
    SP(1),
    Pc(C['sub'],
       'Средняя EGT: 518°C  |  Δ EGT: 60°C (цил.15 549°C − цил.7 489°C)  |  '
       'Цил.15 max=565°C, Цил.6 max=561°C — КРИТИЧНО (&gt;560°C)',
       size=8, lead=11),
    SP(2),
]

EGT = [
    (1,  531, False),
    (2,  521, False),
    (3,  490, False),
    (4,  502, False),
    (5,  500, False),
    (6,  543, True),
    (7,  489, False),
    (8,  530, False),
    (9,  534, False),
    (10, 517, False),
    (11, 533, False),
    (12, 519, False),
    (13, 493, False),
    (14, 503, False),
    (15, 549, True),
    (16, 536, False),
]
EGT_MAX = {6: 561, 15: 565}

AVG_EGT  = 518
chart_w  = PW
chart_h  = 170
n_cyl    = len(EGT)
row_h    = chart_h / n_cyl
label_w  = 34
val_w    = 80
bar_area = chart_w - label_w - val_w - 4
max_val  = 600

d = Drawing(chart_w, chart_h)
d.add(Rect(0, 0, chart_w, chart_h,
           fillColor=C['panel'], strokeColor=C['border'], strokeWidth=0.3))

def egt_bar_color(val, critical):
    if val > 560:
        return C['red']
    if val >= 520:
        return C['orange']
    if val < 450:
        return C['blue']
    return C['green']

for i, (cyl, val, critical) in enumerate(EGT):
    y = chart_h - (i + 1) * row_h
    row_bg_col = C['redbg'] if critical else (C['altrow'] if i % 2 == 0 else C['panel'])
    d.add(Rect(0, y, chart_w, row_h,
               fillColor=row_bg_col, strokeColor=None, strokeWidth=0))

    lbl = f'Цил.{cyl:02d}'
    lbl_color = C['red'] if critical else C['sub']
    d.add(String(2, y + row_h * 0.28, lbl,
                 fontName='Sans-Bold' if critical else 'Sans', fontSize=6.5,
                 fillColor=lbl_color))

    bar_len = (val / max_val) * bar_area
    bar_col = egt_bar_color(val, critical)
    bar_y   = y + row_h * 0.15
    bar_h_px = row_h * 0.70
    d.add(Rect(label_w, bar_y, bar_len, bar_h_px,
               fillColor=bar_col, strokeColor=None, strokeWidth=0))

    fn_bold = 'Sans-Bold' if (critical or val >= 520) else 'Sans'
    if critical:
        cyl_max = EGT_MAX.get(cyl, val + 12)
        val_txt = f'{val}°C  (max {cyl_max}°C!)'
        vc = C['red']
    else:
        val_txt = f'{val}°C'
        vc = egt_bar_color(val, critical)
    d.add(String(label_w + bar_len + 4, y + row_h * 0.28, val_txt,
                 fontName=fn_bold, fontSize=6.5, fillColor=vc))

avg_x = label_w + (AVG_EGT / max_val) * bar_area
d.add(Line(avg_x, 0, avg_x, chart_h,
           strokeColor=C['blue'], strokeWidth=0.8,
           strokeDashArray=[3, 2]))
d.add(String(avg_x + 2, chart_h - 8, f'avg={AVG_EGT}°C',
             fontName='Sans', fontSize=6, fillColor=C['blue']))

for tick_val in [400, 450, 490, 520, 560]:
    tx = label_w + (tick_val / max_val) * bar_area
    d.add(Line(tx, 0, tx, chart_h,
               strokeColor=C['border'], strokeWidth=0.3,
               strokeDashArray=[2, 3]))
    d.add(String(tx - 7, 1, str(tick_val),
                 fontName='Sans', fontSize=5, fillColor=C['sub']))

thresh_x = label_w + (560 / max_val) * bar_area
d.add(Line(thresh_x, 0, thresh_x, chart_h,
           strokeColor=C['red'], strokeWidth=0.8,
           strokeDashArray=[4, 2]))
d.add(String(thresh_x + 2, chart_h - 8, '560°C',
             fontName='Sans-Bold', fontSize=6, fillColor=C['red']))

story.append(d)
story.append(SP(1))

legend_egt = [
    (C['green'],  '■ Норма (450–520°C)'),
    (C['orange'], '■ Повышенная (520–560°C)'),
    (C['red'],    '■ КРИТИЧНО (&gt;560°C max)'),
    (C['blue'],   '■ Пониженная (&lt;450°C)'),
    (C['blue'],   '─ ─ Среднее (518°C)'),
]
leg_cells2 = [Pc(tc, lbl, size=7, lead=10) for tc, lbl in legend_egt]
leg_t2 = Table([leg_cells2], colWidths=[PW / 5] * 5)
leg_t2.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), C['panel']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING',   (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 3),
]))
story += [leg_t2, SP(2)]

story += [
    info_box(
        'ЦИЛ.15 КРИТИЧНО: avg=549°C, max=565°C &gt; порог 560°C. Позиция 8R (правая банка). '
        'Требуется проверка форсунки и ГБЦ цил.15.',
        bg_color=C['redbg'], text_color=C['red'], bold=True, size=8, border_color=C['red']
    ),
    SP(1),
    info_box(
        'ЦИЛ.6 ПРЕДУПРЕЖДЕНИЕ: avg=543°C, max=561°C &gt; порог 560°C. Позиция 3R (правая банка). '
        'Требуется проверка форсунки и ГБЦ цил.6.',
        bg_color=C['orangebg'], text_color=C['orange'], bold=True, size=8, border_color=C['orange']
    ),
    SP(1),
    info_box(
        'ДЕЛЬТА EGT=60°C (цил.15: 549°C − цил.7: 489°C): '
        'значение на уровне порога предупреждения 60°C. '
        'Рекомендуется снять DML после диагностики и замены форсунок цил.6, цил.15.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС (14 проб)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС — 14 ПРОБ (янв.–май 2026)', 3)
story += [SP(1)]

story.append(info_box(
    '14 проб — ВСЕ \"Допустимое\". Двигатель по металлам — в норме.\n'
    'Fe=0–2 мг/кг (норма &lt;30) — норма. Cu=0–2 мг/кг — норма. Si=0 — норма.\n'
    'TBN: снижается до мин. 93.0 мг/кг (13.05.2026) — запас достаточный (норма &gt;70).\n'
    'Сажа: рост до 164 мг/кг (13.05.2026, &gt;150!), последняя проба 155 мг/кг — допустимо (&lt;250).\n'
    'Смена масла: 10.03.2026 (М/Ч 1748) — G-Profi заменён на SUPER HPD PLUS 15W-40.',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

oil_cols = ['Дата', 'М/Ч ДВС', 'Нар.масла', 'Fe', 'Cu', 'Si', 'TAN', 'TBN', 'Сажа', 'Статус']
oil_w = [PW*0.11, PW*0.09, PW*0.10, PW*0.06, PW*0.06, PW*0.06, PW*0.07, PW*0.07, PW*0.08, PW*0.30]
story.append(hdr_row(oil_cols, oil_w))

oil_data_raw = [
    ('05.01.2026',  256,  242,  0, 0, 0, 15.8, 120.9, 138),
    ('15.01.2026',  498,  242,  0, 0, 0, 15.4, 118.2, 136),
    ('26.01.2026',  747,  249,  0, 0, 0, 15.6, 120.1, 137),
    ('05.02.2026',  999,  252,  2, 6, 0, 15.5, 118.0, 138),
    ('16.02.2026', 1251,  252,  0, 0, 0, 15.4, 122.1, 132),
    ('26.02.2026', 1495,  244,  1, 0, 0, 15.6, 117.0, 141),
    ('10.03.2026', 1748,  253,  0, 0, 0, 14.2,  97.5, 149),
    ('20.03.2026', 1998,  250,  0, 0, 0, 14.2, 104.3, 139),
    ('31.03.2026', 2250,  252,  0, 0, 0, 14.3,  98.5, 149),
    ('10.04.2026', 2498,  248,  0, 0, 0, 14.1,  97.4, 148),
    ('21.04.2026', 2744,  246,  0, 0, 0, 14.4, 112.0, 131),
    ('02.05.2026', 2995,  251,  1, 0, 0, 14.3,  97.1, 152),
    ('13.05.2026', 3249,  254,  0, 0, 0, 14.6,  93.0, 164),
    ('24.05.2026', 3504,  255,  1, 0, 0, 14.8,  99.4, 155),
]

CHANGE_ROW = 6
SOOT_ROW   = 12

oil_rows = []
for i, row in enumerate(oil_data_raw):
    date, mch, nm, fe, cu, si, tan, tbn, soot = row
    fe_col   = C['orange'] if fe >= 2 else C['text']
    cu_col   = C['orange'] if cu >= 5 else C['text']
    soot_col = C['orange'] if soot >= 150 else C['text']
    tbn_col  = C['orange'] if tbn < 95 else C['text']
    if i == CHANGE_ROW:
        row_bg = C['bluebg']
    elif i == SOOT_ROW:
        row_bg = C['orangebg']
    else:
        row_bg = C['altrow'] if i % 2 == 0 else C['altrow2']

    note = 'Допустимое'
    note_col = C['green']
    if i == CHANGE_ROW:
        note = 'Допустимое  ← Смена масла'
        note_col = C['blue']
    elif i == SOOT_ROW:
        note = 'Допустимое  ← Сажа &gt;150!'
        note_col = C['orange']

    oil_rows.append([
        Pc(C['text'],  date,         size=7, lead=10),
        Pc(C['text'],  str(mch),     size=7, lead=10),
        Pc(C['text'],  str(nm),      size=7, lead=10),
        Pc(fe_col,     str(fe),      bold=(fe >= 2), size=7, lead=10),
        Pc(cu_col,     str(cu),      bold=(cu >= 5), size=7, lead=10),
        Pc(C['text'],  str(si),      size=7, lead=10),
        Pc(C['text'],  f'{tan:.1f}', size=7, lead=10),
        Pc(tbn_col,    f'{tbn:.1f}', bold=(tbn < 95), size=7, lead=10),
        Pc(soot_col,   str(soot),    bold=(soot >= 150), size=7, lead=10),
        Pc(note_col,   note,         size=7, lead=10),
    ])

oil_extra = [('BACKGROUND', (0, i), (-1, i),
              C['bluebg'] if i == CHANGE_ROW else
              (C['orangebg'] if i == SOOT_ROW else
               (C['altrow'] if i % 2 == 0 else C['altrow2'])))
             for i in range(len(oil_rows))]
story.append(tbl(oil_rows, oil_w, style_extra=oil_extra))
story += [SP(2)]

story += section_title('ТРЕНДЫ МАСЛА ДВС NTE200 №88', None)
story += [SP(1)]

trend_data = [
    ('Fe',   '0–2 мг/кг',     '&lt;30 мг/кг',  'НОРМА — минимальный износ',                              C['green'],  C['greenbg']),
    ('Cu',   '0–6 мг/кг',     '&lt;15 мг/кг',  'Разовый Cu=6 (05.02.2026) — норма', C['text'],   C['altrow']),
    ('Si',   '0 мг/кг',       '&lt;10 мг/кг',  'НОРМА — все пробы 0',                                    C['green'],  C['greenbg']),
    ('TBN',  '93.0–122.1',    '&gt;70 мг/кг',  'Мин. 93.0 (13.05.2026) — допустимо, строго по графику', C['text'],   C['altrow2']),
    ('TAN',  '14.1–15.8',     '&lt;15',         '15.8 (05.01.2026) — в норме для режима',                C['text'],   C['altrow']),
    ('Сажа', '131–164 мг/кг', '&lt;250 мг/кг', 'Рост до 164 мг/кг — допустимо, контролировать',         C['orange'], C['orangebg']),
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
        'СМЕНА МАСЛА 10.03.2026 (М/Ч 1748): G-Profi заменён на SUPER HPD PLUS 15W-40. '
        'После смены TBN стабилизировался на 97–112 мг/кг. '
        'Сажа 164 мг/кг (13.05) — допустимо, следующая проба 155 мг/кг. '
        'Критический уровень сажи QSK50 — 250 мг/кг. Строго придерживаться графика ТО масла 250 м/ч.',
        bg_color=C['bluebg'], text_color=C['blue'], size=8, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('РЕКОМЕНДАЦИИ', 4)
story += [SP(2)]

recommendations = [
    (
        '1. КРИТИЧНО',
        'Цил.15 max=565°C &gt;560°C — проверить форсунку/ГБЦ цил.15 (8R)',
        'Цилиндр 15 (позиция 8R, правая банка): avg=549°C, max=565°C превышает критический порог 560°C. '
        'Необходима замена форсунки и/или диагностика ГБЦ цил.15. '
        'Риск: разрушение выпускного клапана при дальнейшей эксплуатации.',
        C['red'], C['redbg'],
    ),
    (
        '2. ПРЕДУПРЕЖДЕНИЕ',
        'Цил.6 max=561°C &gt;560°C — проверить форсунку/ГБЦ цил.6 (3R)',
        'Цилиндр 6 (позиция 3R, правая банка): avg=543°C, max=561°C превышает порог 560°C. '
        'Проверка форсунки цил.6 совместно с цил.15. '
        'После диагностики снять повторный DML для подтверждения.',
        C['orange'], C['orangebg'],
    ),
    (
        '3. ПРЕДУПРЕЖДЕНИЕ',
        'Δ EGT=60°C = порог — снять DML после диагностики форсунок',
        'Разброс EGT = 60°C (цил.15 549°C − цил.7 489°C) достиг порога предупреждения 60°C. '
        'После замены/проверки форсунок цил.15 и цил.6 выполнить повторный DML-замер '
        'для оценки равномерности работы цилиндров.',
        C['orange'], C['orangebg'],
    ),
    (
        '4. НАБЛЮДЕНИЕ',
        'T масла max=107.7°C &gt;105°C — диагностика масляного теплообменника',
        'Среднее 102.4°C — в норме. Максимум 107.7°C выше порога предупреждения 105°C. '
        'Диагностика масляного теплообменника при следующем плановом ТО. '
        'Критический порог 115°C не достигнут.',
        C['blue'], C['bluebg'],
    ),
    (
        '5. НАБЛЮДЕНИЕ',
        'TBN=93 мин. (13.05.2026), сажа 164 мг/кг — строго по графику ТО масла 250 м/ч',
        'TBN 93 мг/кг (13.05.2026) — минимальное значение за период; на следующей пробе '
        '(≈ 3 750 м/ч) проверить динамику. '
        'Сажа 164 мг/кг выше уровня наблюдения 150 мг/кг. '
        'Критический уровень для QSK50 — 250 мг/кг. '
        'Строго соблюдать интервал замены масла 250 м/ч.',
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

story += section_title('ECM КОДЫ НЕИСПРАВНОСТЕЙ (12.05.2026, ESN 33238554)', None)
story += [SP(1)]

fault_cols = ['Код', 'Статус', 'Кол-во', 'Описание', 'Лампа']
fault_w = [PW*0.08, PW*0.12, PW*0.08, PW*0.58, PW*0.14]
story.append(hdr_row(fault_cols, fault_w))

fault_data = [
    ('611',  'Inactive', '10', 'Остановка горячего двигателя — состояние сохраняется',
     'No lamp',  C['orangebg'], C['orange']),
    ('234',  'Inactive', '1',  'ЧВ/положение коленвала &gt;норма',
     'Red lamp', C['redbg'],   C['red']),
    ('1518', 'Inactive', '1',  'Один/несколько модулей — неисправности среднего уровня',
     '—',        C['altrow'],  C['sub']),
]

fault_rows = []
for code, status, cnt, descr, lamp, row_bg, tc in fault_data:
    fault_rows.append([
        Pc(tc,        code,   bold=True, size=8, lead=11),
        Pc(C['sub'],  status,            size=8, lead=11),
        Pc(tc,        cnt,    bold=True, size=8, lead=11),
        Pc(C['text'], descr,             size=8, lead=11),
        Pc(tc,        lamp,   bold=True, size=8, lead=11),
    ])
fault_extra = [('BACKGROUND', (0, i), (-1, i), fault_data[i][5])
               for i in range(len(fault_rows))]
story.append(tbl(fault_rows, fault_w, style_extra=fault_extra))

story += [
    SP(2),
    info_box(
        'Все коды — Inactive (неактивны на 12.05.2026). '
        'Код 611 (горячий останов): 10 раз — указывает на периодические остановки горячего двигателя. '
        'Код 234 (Red lamp): 1 раз — требует наблюдения. '
        'Рекомендуется сброс кодов после диагностики форсунок цил.15 и цил.6.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(3),
]

story += [
    HR(C['blue']),
    SP(1),
    Pc(C['blue'], 'ИТОГОВАЯ ОЦЕНКА ДВИГАТЕЛЯ №88', bold=True, size=11, lead=14, align=TA_CENTER),
    SP(1),
]

summary_kpi = [
    ('МАСЛО ДВС',    'НОРМА',         'Все 14 проб — чисто',              C['green'],  C['greenbg']),
    ('EGT ЦИЛ.15',  'КРИТИЧНО',      'max=565°C — форсунка/ГБЦ',   C['red'],    C['redbg']),
    ('Δ EGT',        '60°C',          'Порог достигнут — DML после ТО',   C['orange'], C['orangebg']),
    ('СЛЕД.ШАГ',     'Форсунки 15+6', 'Диагностика, потом DML-замер',     C['blue'],   C['bluebg']),
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
       'Отчёт подготовлен на основе: DML-20260608 (1183 строки, 217 параметров, 20 мин)  •  '
       'Спектральный анализ масла ДВС NTE200 №88 (14 проб, янв.–май 2026)  •  '
       'История ТО (ОТЧЕТ Полюс Магадан.хлсх)  •  '
       'ECM Fault Codes ESN 33238554 (12.05.2026)  |  Горная Евразия  •  08.06.2026',
       size=7, lead=10),
]

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
