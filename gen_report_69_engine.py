#!/usr/bin/env python3
"""NTE200 #69 — QSK50 MCRS V16 — Engine Diagnostic Report (07.06.2026).
COMPLETE maintenance history — all 20 events from ОТЧЕТ Полюс Магадан.xlsx.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether,
                                Flowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String, PolyLine
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
                      'NTE200 №69  •  QSK50 MCRS V16  •  Диагностика ДВС  •  07.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    # Bottom bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  Наработка ДВС: ~10 603 м/ч  •  07.06.2026')
    canvas.restoreState()

# ── Document setup ─────────────────────────────────────────────────────────────
OUT = '/home/user/NTE200/NTE200_69_engine_report_07062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=20 * mm,
    rightMargin=20 * mm,
    topMargin=16 * mm,
    bottomMargin=11 * mm,
    title='NTE200 №69 — Диагностика ДВС QSK50 MCRS V16 — 07.06.2026',
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
    Pc(C['blue'], 'NTE200 №69  /  QSK50 MCRS V16',
       bold=True, size=13, lead=17, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],
       'Дата отчёта: 07.06.2026  |  Наработка ДВС: ~10 603 м/ч (по состоянию на 15.05.2026)',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# Machine info block
info_data = [
    [Pc(C['sub'], 'Гаражный №:', bold=True, size=8, lead=11),
     Pc(C['white'], '69', bold=True, size=8, lead=11),
     Pc(C['sub'], 'Модель:', bold=True, size=8, lead=11),
     Pc(C['white'], 'NTE200  QSK50 MCRS V16', bold=True, size=8, lead=11)],
    [Pc(C['sub'], 'Текущие М/Ч:', bold=True, size=8, lead=11),
     Pc(C['green'], '~10 603 м/ч', bold=True, size=8, lead=11),
     Pc(C['sub'], 'DML файлы:', bold=True, size=8, lead=11),
     Pc(C['grey'], 'ОТСУТСТВУЮТ', bold=True, size=8, lead=11)],
    [Pc(C['sub'], 'Тех. отчёт:', bold=True, size=8, lead=11),
     Pc(C['blue'], 'Замена клапанов 25.05.2026', size=8, lead=11),
     Pc(C['sub'], 'Масло ДВС:', bold=True, size=8, lead=11),
     Pc(C['green'], '34 пробы — ВСЕ "Допустимое"', size=8, lead=11)],
]
info_col_w = [PW * 0.17, PW * 0.33, PW * 0.17, PW * 0.33]
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

# KPI section header
story += [
    Pc(C['sub'], 'КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ', bold=True, size=9, lead=12, align=TA_CENTER),
    SP(1),
]

kpi_cards_1 = [
    ('М/Ч ДВС',        '~10 603',        'м/ч (15.05.2026)',                         C['green'],  C['greenbg']),
    ('МАСЛО ДВС',      'ОТЛИЧНО',         '34 пробы — все Fe=0-3 мг/кг',             C['green'],  C['greenbg']),
    ('РЕМОНТ ДВС',     '16.04.2026',      '197 ч простоя (8 суток)',                  C['orange'], C['orangebg']),
]
kpi_cards_2 = [
    ('СОСТАВ РЕМОНТА', 'Цил. 3L,4L,3R,4R','клапаны/поршни/ГБЦ',                     C['red'],    C['redbg']),
    ('ОШИБКИ EGT',     'Цил.3 и 16',      'повторяющиеся после ремонта',              C['orange'], C['orangebg']),
    ('DML ДАННЫЕ',     'ОТСУТСТВУЮТ',     'не предоставлены для №69',                 C['grey'],   C['greybg']),
]

def make_kpi_table(cards):
    n = len(cards)
    w = [PW / n] * n
    hdr_cells = [Pc(C['sub'],  lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _, _ in cards]
    val_cells  = [Pc(vc,       val, bold=True, size=13, lead=17, align=TA_CENTER) for _, val, _, vc, _ in cards]
    sub_cells  = [Pc(C['sub'], sub, size=7, lead=10, align=TA_CENTER) for _, _, sub, _, _ in cards]
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
        'ОБЩАЯ ОЦЕНКА ДВС: Двигатель работоспособен. Масло в отличном состоянии (апр.2025–май.2026, '
        '34 пробы — все "Допустимое"). Основные риски — последствия ремонта ДВС (16.04.2026, 197 ч): '
        'замена клапанов/поршней/ГБЦ цил.3L,4L,3R,4R и повторяющиеся ошибки EGT цил.3 и 16. '
        'DML-данные отсутствуют.',
        bg_color=C['panel2'], text_color=C['text'], size=9, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — ПОЛНАЯ ИСТОРИЯ ВОЗДЕЙСТВИЙ (все 20 событий)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ПОЛНАЯ ИСТОРИЯ ВОЗДЕЙСТВИЙ (все события, март–май 2026)', 1)
story += [
    SP(1),
    Pc(C['sub'],
       'Источник: ОТЧЕТ Полюс Магадан.xlsx + Технический отчёт 25.05.2026',
       size=8, lead=11),
    SP(2),
]

# Legend
legend_cols = [
    (C['bluebg'],   C['blue'],   '■ Плановое ТО'),
    (C['redbg'],    C['red'],    '■ Ремонт ДВС'),
    (C['orangebg'], C['orange'], '■ Ошибки EGT'),
    (C['greenbg'],  C['green'],  '■ Тех. отчёт'),
    (C['altrow'],   C['sub'],    '■ Прочее'),
]
leg_cells = [Pc(tc, lbl, size=7, lead=10) for bg, tc, lbl in legend_cols]
leg_t = Table([leg_cells], colWidths=[PW / 5] * 5)
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

# Full history — 20 events
# Columns: Дата | Длит. | М/Ч | Событие
hist_cols = ['Дата', 'Длит.', 'М/Ч', 'Событие / Работы']
hist_w = [PW*0.12, PW*0.10, PW*0.10, PW*0.68]
story.append(hdr_row(hist_cols, hist_w))

# (date, dur, mh, event, row_bg, event_col, bold)
hist_data = [
    ('01.03.2026', '0:05',  '9 115',   'Осмотр проводки бокового освещения',
     C['altrow'],   C['sub'],    False),
    ('02.03.2026', '0:12',  '9 119',   'Диагностика эл.оборудования (вышли из строя задние фонари)',
     C['altrow2'],  C['sub'],    False),
    ('02.03.2026', '0:34',  '9 127',   'Замена отвальных фонарей',
     C['altrow'],   C['sub'],    False),
    ('07.03.2026', '2:56',  '9 251',   'ТО-1',
     C['bluebg'],   C['blue'],   True),
    ('18.03.2026', '2:55',  '9 511',   'ТО-2',
     C['bluebg'],   C['blue'],   True),
    ('25.03.2026', '17:13', '9 656',   'Диагностика и ремонт проводки ДВС',
     C['orangebg'], C['orange'], True),
    ('30.03.2026', '4:45',  '9 756',   'ТО-1',
     C['bluebg'],   C['blue'],   True),
    ('09.04.2026', '21:26', '10 001',  'ТО-6 — Замена гофры на обдув шкафов, сварочные работы, '
                                       'восстановление кронштейна, трассы гидр. труб, замена впускных труб',
     C['orangebg'], C['orange'], False),
    ('13.04.2026', '0:15',  '10 067',  'Диагностика блока динамического замедления',
     C['altrow'],   C['sub'],    False),
    ('16.04.2026', '197 ч', '10 130',  'РЕМОНТ ДВС — Диагностика и ремонт ДВС (8 суток простоя)',
     C['redbg'],    C['red'],    True),
    ('26.04.2026', '0:15',  '10 169',  'Протяжка хомутов на интеркулере',
     C['altrow2'],  C['sub'],    False),
    ('27.04.2026', '0:28',  '10 186',  'Ошибка EGT цилиндр 3 — поджатие пинов, осмотр эл. проводки',
     C['orangebg'], C['orange'], False),
    ('27.04.2026', '0:20',  '10 197',  'Ошибка EGT цилиндр 16 — поджатие пинов, осмотр эл. проводки',
     C['orangebg'], C['orange'], False),
    ('28.04.2026', '0:33',  '10 209',  'Ошибка EGT цилиндр 3 (ПОВТОР) — поджатие пинов',
     C['redbg'],    C['red'],    True),
    ('30.04.2026', '2:59',  '10 248',  'ТО-1. Сварочные работы бочки РМК',
     C['bluebg'],   C['blue'],   True),
    ('04.05.2026', '0:20',  '10 348',  'Ошибка EGT цилиндр 16 (ПОВТОР) — поджатие пинов',
     C['redbg'],    C['red'],    True),
    ('11.05.2026', '2:19',  '10 500',  'ТО-2. Регулировка стоек ПГП',
     C['bluebg'],   C['blue'],   True),
    ('12.05.2026', '0:25',  '10 522',  'Диагностика ошибок GE, осмотр плат',
     C['altrow'],   C['sub'],    False),
    ('15.05.2026', '0:02',  '10 603',  'Замена щётки стеклоочистителя',
     C['altrow2'],  C['sub'],    False),
    ('25.05.2026', 'нет данных', '~10 650',
                               'Замена клапанов (технический отчёт 25.05.2026)',
     C['greenbg'],  C['green'],  True),
]

hist_rows = []
for date, dur, mh, event, row_bg, ecol, bold in hist_data:
    hist_rows.append([
        Pc(ecol,    date,  bold=bold, size=7, lead=10),
        Pc(C['sub'], dur,            size=7, lead=10),
        Pc(C['sub'], mh,             size=7, lead=10),
        Pc(C['text'], event,         size=7, lead=10),
    ])

hist_extra = [('BACKGROUND', (0, i), (-1, i), hist_data[i][4])
              for i in range(len(hist_rows))]
story.append(tbl(hist_rows, hist_w, style_extra=hist_extra))

story += [
    SP(2),
    info_box(
        'Итого за период: 7 плановых ТО • 1 крупный ремонт ДВС (197 ч) • '
        '4 ошибки EGT (цил.3 и 16, повторяющиеся) • 1 длительная диагностика проводки ДВС (17:13 ч) • '
        'Доливки масла ДВС: данных НЕТ.',
        bg_color=C['panel2'], text_color=C['text'], size=8.5, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — СОСТАВ РЕМОНТА ДВС 16.04.2026 + ОБЪЯСНЕНИЕ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СОСТАВ РЕМОНТА ДВС 16.04.2026 (197 ч)', 2)
story += [
    SP(1),
    Pc(C['sub'], 'Источник: ГБЦ ремонты.xlsx  |  Наработка на момент ремонта: 10 130 м/ч',
       size=8, lead=11),
    SP(2),
]

# Summary repair KPIs
repair_kpi = [
    ('ДАТА НАЧАЛА',        '16.04.2026',   C['orange'], C['orangebg']),
    ('ДЛИТЕЛЬНОСТЬ',       '197 часов',    C['red'],    C['redbg']),
    ('НАРАБОТКА ДВС',      '10 130 м/ч',   C['text'],   C['panel']),
    ('ЦИЛИНДРОВ ЗАТРОНУТО', '4',           C['red'],    C['redbg']),
]
kpi_w4 = [PW / 4] * 4
kpi_hdr4 = [Pc(C['sub'],  lbl,  bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _ in repair_kpi]
kpi_val4 = [Pc(vc,        val,  bold=True, size=14, lead=18, align=TA_CENTER) for _, val, vc, _ in repair_kpi]
kpi_t4   = Table([kpi_hdr4, kpi_val4], colWidths=kpi_w4)
kpi_ts4  = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING',   (0, 0), (-1, -1), 6),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 6),
    ('ALIGN',        (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
])
for i, (_, _, _, bg) in enumerate(repair_kpi):
    kpi_ts4.add('BACKGROUND', (i, 0), (i, 0), C['hdr'])
    kpi_ts4.add('BACKGROUND', (i, 1), (i, 1), bg)
kpi_t4.setStyle(kpi_ts4)
story += [kpi_t4, SP(3)]

# Repair composition table
story += [
    Pc(C['red'], 'СОСТАВ РЕМОНТА ДВС (из ГБЦ ремонты.xlsx)', bold=True, size=10, lead=14),
    HRFlowable(width='100%', thickness=0.8, color=C['red'], spaceAfter=4, spaceBefore=2),
    SP(1),
]

rep_cols = ['Позиция', 'Цилиндр', 'Расположение', 'Выполненные работы']
rep_w = [PW*0.08, PW*0.15, PW*0.25, PW*0.52]
story.append(hdr_row(rep_cols, rep_w))

rep_data = [
    ('1', '3L (цил. №5)', 'Левый ряд, 3-я позиция', 'Замена 2 выпускных клапанов',                  C['orangebg'], C['orange']),
    ('2', '4L (цил. №7)', 'Левый ряд, 4-я позиция', 'Замена 2 выпускных клапанов',                  C['orangebg'], C['orange']),
    ('3', '3R (цил. №6)', 'Правый ряд, 3-я позиция','Замена ГБЦ + Замена поршня',                   C['redbg'],    C['red']),
    ('4', '4R (цил. №8)', 'Правый ряд, 4-я позиция','Замена 2 выпускных клапанов + Замена поршня',  C['redbg'],    C['red']),
]

rep_rows = []
for num, cyl, pos, work, row_bg, ecol in rep_data:
    rep_rows.append([
        Pc(C['sub'],  num,  size=8, lead=11),
        Pc(ecol,      cyl,  bold=True, size=8, lead=11),
        Pc(C['sub'],  pos,  size=8, lead=11),
        Pc(C['text'], work, size=8, lead=11),
    ])

rep_extra = [('BACKGROUND', (0, i), (-1, i), rep_data[i][4])
             for i in range(len(rep_rows))]
story.append(tbl(rep_rows, rep_w, style_extra=rep_extra))

story += [
    SP(3),
    Pc(C['text'],
       '197 часов (8 суток) — ремонт выявил износ поршней и клапанов цилиндров 3-4 правого '
       'и левого рядов. После ремонта — ошибки датчиков EGT на цил.3 и 16, что может быть '
       'связано с нарушением контактов при демонтаже/монтаже ДВС.',
       size=9, lead=14),
    SP(3),
]

# EGT errors after repair
story += [
    Pc(C['orange'], 'ОШИБКИ EGT ПОСЛЕ РЕМОНТА ДВС', bold=True, size=10, lead=14),
    HRFlowable(width='100%', thickness=0.8, color=C['orange'], spaceAfter=4, spaceBefore=2),
    SP(1),
]

egt_cols = ['Дата', 'М/Ч', 'Цилиндр', 'Тип события', 'Принятые меры']
egt_w = [PW*0.12, PW*0.10, PW*0.10, PW*0.22, PW*0.46]
story.append(hdr_row(egt_cols, egt_w))

egt_data = [
    ('27.04.2026', '10 186', 'Цил.3',  '1-я ошибка EGT',  'Поджатие пинов, осмотр эл. проводки',  C['orangebg'], C['orange'], False),
    ('27.04.2026', '10 197', 'Цил.16', '1-я ошибка EGT',  'Поджатие пинов, осмотр эл. проводки',  C['orangebg'], C['orange'], False),
    ('28.04.2026', '10 209', 'Цил.3',  'ПОВТОР ошибки',   'Повторное поджатие пинов',               C['redbg'],    C['red'],    True),
    ('04.05.2026', '10 348', 'Цил.16', 'ПОВТОР ошибки',   'Поджатие пинов (третий случай)',          C['redbg'],    C['red'],    True),
]

egt_rows = []
for dt, mh, cyl, ev, action, row_bg, ecol, bold in egt_data:
    egt_rows.append([
        Pc(ecol,      dt,     bold=bold, size=8, lead=11),
        Pc(C['sub'],  mh,                size=8, lead=11),
        Pc(ecol,      cyl,    bold=True, size=8, lead=11),
        Pc(ecol,      ev,     bold=bold, size=8, lead=11),
        Pc(C['text'], action,            size=8, lead=11),
    ])

egt_extra = [('BACKGROUND', (0, i), (-1, i), egt_data[i][5])
             for i in range(len(egt_rows))]
story.append(tbl(egt_rows, egt_w, style_extra=egt_extra))

story += [
    SP(2),
    info_box(
        'Важно: все 4 ошибки EGT появились ПОСЛЕ 197-часового ремонта ДВС. '
        'Поджатие пинов — временная мера. Повторение ошибок указывает на повреждение '
        'термопар или жгута при демонтаже/монтаже ДВС. Требуется замена датчиков EGT цил.3 и 16.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8.5, border_color=C['orange']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС (все 34 пробы + тренды)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС — 34 ПРОБЫ (апр.2025 – май.2026)', 3)
story += [SP(1)]

story.append(info_box(
    'ВСЕ 34 пробы — "Допустимое". Двигатель не изнашивается.\n'
    'Fe=0–3 мг/кг (норма <30) — исключительно чисто. Cu=0–6 мг/кг — норма.\n'
    'Si: единичный выброс 27 мг/кг (30.04.2026) — вероятно загрязнение при отборе. '
    'TBN снижается 120→92 мг/кг — закономерное истощение, замена масла каждые 250 м/ч. '
    'Сажа: умеренный рост 119→153 мг/кг — допустимо.\n'
    'Доливки масла ДВС для №69: данных НЕТ (из файла Доливки Горная Евразия.xlsx).',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

oil_cols = ['Дата', 'М/Ч ДВС', 'Нар.масла', 'Fe', 'Cu', 'Si', 'TAN', 'TBN', 'Сажа', 'Статус']
oil_w = [PW*0.11, PW*0.10, PW*0.10, PW*0.06, PW*0.06, PW*0.06, PW*0.07, PW*0.07, PW*0.07, PW*0.30]
story.append(hdr_row(oil_cols, oil_w))

oil_data_raw = [
    ('12.04.2025',  1761,  250,  0, 0,  0, 14.1, 117.8, 119),
    ('23.04.2025',  2019,  250,  0, 0,  0, 15.1, 123.8, 126),
    ('04.05.2025',  2252,  250,  0, 0,  0, 14.7, 115.2, 131),
    ('14.05.2025',  2499,  250,  0, 0,  0, 14.8, 115.3, 132),
    ('29.05.2025',  2750,  250,  0, 0,  0, 14.9, 120.3, 127),
    ('09.06.2025',  3000,  250,  0, 0,  0, 14.9, 119.3, 128),
    ('21.06.2025',  3256,  250,  0, 0,  0, 14.9, 113.3, 136),
    ('01.07.2025',  3503,  250,  0, 0,  0, 14.9, 116.2, 132),
    ('12.07.2025',  3762,  250,  0, 0,  0, 15.5, 125.0, 130),
    ('22.07.2025',  4000,  250,  0, 0,  0, 16.8, 123.9, 147),
    ('02.08.2025',  4250,  250,  0, 0,  0, 15.6, 116.4, 141),
    ('13.08.2025',  4500,  250,  0, 0,  0, 15.5, 116.1, 140),
    ('24.08.2025',  4753,  250,  0, 0,  0, 15.3, 120.6, 132),
    ('05.09.2025',  4985,  235,  0, 0,  0, 15.4, 114.0, 142),
    ('17.09.2025',  5247,  262,  0, 0,  0, 15.5, 119.3, 136),
    ('27.09.2025',  5500,  253,  0, 2,  2, 15.7, 118.4, 140),
    ('08.10.2025',  5752,  252,  0, 2,  0, 15.5, 119.6, 136),
    ('19.10.2025',  6002,  250,  2, 1,  0, 15.5, 118.1, 138),
    ('30.10.2025',  6256,  254,  2, 0,  0, 15.5, 118.6, 137),
    ('10.11.2025',  6517,  261,  0, 0,  0, 15.3, 118.3, 135),
    ('12.12.2025',  1255,  252,  3, 6,  0, 15.8, 121.8, 137),  # аномальный М/Ч
    ('23.12.2025',  7493,  238,  0, 0,  0, 15.6, 122.5, 134),
    ('03.01.2026',  7753,  260,  0, 0,  0, 15.8, 124.1, 134),
    ('14.01.2026',  8007,  254,  0, 0,  0, 15.3, 119.0, 134),
    ('24.01.2026',  8248,  241,  0, 0,  0, 15.0, 121.2, 127),
    ('04.02.2026',  8500,  252,  1, 0,  0, 15.0, 128.9, 119),
    ('14.02.2026',  8747,  247,  0, 0,  0, 16.0, 121.5, 140),
    ('26.02.2026',  9004,  251,  1, 0,  0, 15.7, 111.2, 150),
    ('08.03.2026',  9251,  247,  0, 0,  0, 14.2,  97.1, 150),
    ('19.03.2026',  9511,  262,  0, 0,  0, 14.0,  96.5, 148),
    ('31.03.2026',  9756,  245,  0, 0,  0, 14.3, 100.1, 147),
    ('10.04.2026', 10001,  245,  0, 0,  0, 14.2,  99.2, 147),
    ('30.04.2026', 10248,  249,  0, 0, 27, 14.4,  97.1, 153),
    ('11.05.2026', 10500,  252,  0, 0,  0, 13.0,  91.9, 140),
]

oil_rows = []
for i, row in enumerate(oil_data_raw):
    date, mh, oil_h, fe, cu, si, tan, tbn, soot = row
    is_anomaly_mh = (mh == 1255)
    is_si_spike   = (si == 27)
    is_tbn_low    = (tbn < 100)
    fe_col  = C['orange'] if fe > 0 else C['text']
    cu_col  = C['orange'] if cu > 3 else C['text']
    si_col  = C['red']    if is_si_spike else (C['orange'] if si > 0 else C['text'])
    tbn_col = C['orange'] if is_tbn_low else C['text']
    mh_str  = f'{mh}*' if is_anomaly_mh else str(mh)
    oil_rows.append([
        Pc(C['orange'] if is_anomaly_mh else C['text'], date, size=6.5, lead=9),
        Pc(C['orange'] if is_anomaly_mh else C['text'], mh_str, bold=is_anomaly_mh, size=6.5, lead=9),
        Pc(C['text'],  str(oil_h),  size=6.5, lead=9),
        Pc(fe_col,     str(fe),     bold=(fe > 0), size=6.5, lead=9),
        Pc(cu_col,     str(cu),     bold=(cu > 3), size=6.5, lead=9),
        Pc(si_col,     str(si),     bold=is_si_spike, size=6.5, lead=9),
        Pc(C['text'],  f'{tan:.1f}', size=6.5, lead=9),
        Pc(tbn_col,    f'{tbn:.1f}', bold=is_tbn_low, size=6.5, lead=9),
        Pc(C['text'],  str(soot),   size=6.5, lead=9),
        Pc(C['green'], 'Допустимое', size=6.5, lead=9),
    ])

oil_extra = []
for i, row in enumerate(oil_data_raw):
    _, mh, _, _, _, si, _, _, _ = row
    if mh == 1255:
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), C['orangebg']))
    elif si == 27:
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), C['redbg']))
    else:
        bg = C['altrow'] if i % 2 == 0 else C['altrow2']
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), bg))

story.append(tbl(oil_rows, oil_w, style_extra=oil_extra))
story += [
    SP(2),
    info_box(
        '* М/Ч=1255 (12.12.2025) — аномальная запись: возможно ошибка в базе или другой агрегат. '
        'Все параметры масла в норме (Fe=3, Cu=6), статус "Допустимое".',
        bg_color=C['orangebg'], text_color=C['orange'], size=7.5
    ),
    SP(1),
    info_box(
        '30.04.2026 — Si=27 мг/кг: единичный выброс, следующая проба 11.05 Si=0. '
        'Вероятное загрязнение при отборе пробы. Не признак износа.',
        bg_color=C['redbg'], text_color=C['red'], size=7.5
    ),
    SP(2),
]

# Trend analysis mini-table
story += section_title('ТРЕНДЫ МАСЛА ДВС', None)
story += [SP(1)]

trend_data = [
    ('Fe', '0–3 мг/кг', '<30 мг/кг', 'ОТЛИЧНО — минимальный износ', C['green'],  C['greenbg']),
    ('Cu', '0–6 мг/кг', '<15 мг/кг', 'НОРМА — подшипники в порядке', C['green'],  C['greenbg']),
    ('Si', '0; пик 27', '<10 мг/кг', 'Единичный выброс — загрязнение пробы',     C['orange'], C['orangebg']),
    ('TBN','119→92',    '>70 мг/кг', 'Снижение закономерно, замена 250 м/ч',     C['text'],   C['altrow']),
    ('TAN','13–16',     '<15',        '16.8 (июл) и 16.0 (фев) — в норме для режима', C['text'], C['altrow2']),
    ('Сажа','119→153', '<250 мг/кг','Умеренный рост, режим эксплуатации',        C['text'],   C['altrow']),
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

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — РЕКОМЕНДАЦИИ
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('РЕКОМЕНДАЦИИ', 4)
story += [SP(2)]

recommendations = [
    (
        '1. КРИТИЧНО',
        'Проверить EGT датчики цил.3 и 16',
        'Ошибки EGT возникли ПОСЛЕ ремонта ДВС. Если ошибки сохраняются — заменить термопары. '
        'Эксплуатация без температурного контроля двух цилиндров недопустима.',
        C['red'], C['redbg'],
    ),
    (
        '2. ВАЖНО',
        'Запросить DML-данные для оценки EGT по всем 16 цилиндрам после ремонта',
        'Снять DML-запись под нагрузкой (гружёный рейс) — оценить температуры EGT '
        'по всем 16 цилиндрам, давление наддува, температуру масла и ОЖ. '
        'Особое внимание — цилиндры 3 и 16 (повторные ошибки EGT).',
        C['orange'], C['orangebg'],
    ),
    (
        '3. ВАЖНО',
        'Убедиться, что ТО по маслу продолжается каждые 250 м/ч',
        'TBN=91.9 (11.05.2026) — минимальное значение за весь период мониторинга. '
        'Замена масла обязательна по графику 250 м/ч. '
        'При TBN <80 мг/кг — досрочная замена.',
        C['orange'], C['orangebg'],
    ),
    (
        '4. ИНФО',
        'Si=27 (30.04.2026) — единичный, не тревожно',
        'Следующая проба 11.05.2026 Si=0. Подтверждает загрязнение при отборе, '
        'а не реальный источник кремния. При следующем Si>10 — проверить воздушный фильтр ДВС.',
        C['sub'], C['greybg'],
    ),
]

recs_w = [PW*0.18, PW*0.30, PW*0.52]
rec_rows_all = []
for prio_label, title, body, tc, bg in recommendations:
    rec_rows_all.append([
        Pc(tc,        prio_label, bold=True, size=9, lead=13),
        Pc(tc,        title,      bold=True, size=9, lead=13),
        Pc(C['text'], body,                  size=8.5, lead=12),
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
    Pc(C['blue'], 'ИТОГОВАЯ ОЦЕНКА ДВИГАТЕЛЯ №69', bold=True, size=11, lead=14, align=TA_CENTER),
    SP(1),
]

summary_kpi = [
    ('МАСЛО ДВС',   'ОТЛИЧНО',      'Все 34 пробы — Fe 0-3',              C['green'],  C['greenbg']),
    ('СОСТОЯНИЕ',   'РАБОЧЕЕ',      'После ремонта апр.2026',              C['orange'], C['orangebg']),
    ('EGT РИСК',    'Цил.3 + 16',   'Повторные ошибки — нужна проверка',   C['red'],    C['redbg']),
    ('СЛЕД.ШАГ',    'DML + EGT',    'Запросить DML, проверить термопары',  C['blue'],   C['bluebg']),
]
final_w = [PW / 4] * 4
final_hdr = [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _, _ in summary_kpi]
final_val = [Pc(vc,       val, bold=True, size=12, lead=16, align=TA_CENTER) for _, val, _, vc, _ in summary_kpi]
final_sub = [Pc(vc,       sub, size=7,    lead=10, align=TA_CENTER) for _, _, sub, vc, _ in summary_kpi]

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
       'Отчёт подготовлен на основе: Спектральный анализ масла ДВС (34 пробы, апр.2025–май.2026)  •  '
       'История ТО и ремонтов №69 (ОТЧЕТ Полюс Магадан.xlsx)  •  '
       'ГБЦ ремонты.xlsx (состав ремонта 16.04.2026)  •  '
       'Тех. отчёт 25.05.2026 (замена клапанов)  •  DML: ОТСУТСТВУЮТ  |  Горная Евразия  •  07.06.2026',
       size=7, lead=10),
]

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
