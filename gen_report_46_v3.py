#!/usr/bin/env python3
"""NTE200 #46 — QSK50 MCRS V16 — Comprehensive Diagnostic Report v3 (06.06.2026)."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Font registration ──────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont('Sans',      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('Sans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# ── Colour palette ─────────────────────────────────────────────────────────────
C = dict(
    bg      = colors.HexColor('#1a1a2e'),   # page background
    panel   = colors.HexColor('#16213e'),   # section header / card bg
    panel2  = colors.HexColor('#0f3460'),   # alternate panel
    border  = colors.HexColor('#2a3a5c'),   # grid lines
    text    = colors.HexColor('#e8e8f0'),   # body text
    sub     = colors.HexColor('#8899bb'),   # secondary / table headers
    red     = colors.HexColor('#e94560'),   # critical
    redbg   = colors.HexColor('#2a0d18'),   # critical row bg
    orange  = colors.HexColor('#f5a623'),   # warning
    orangebg= colors.HexColor('#2a1a05'),   # warning row bg
    green   = colors.HexColor('#4caf50'),   # ok
    greenbg = colors.HexColor('#0a1f0a'),   # ok row bg
    blue    = colors.HexColor('#2196f3'),   # info
    bluebg  = colors.HexColor('#051530'),   # info row bg
    altrow  = colors.HexColor('#1e2d4a'),   # alternate table row
    hdr     = colors.HexColor('#0d1829'),   # table header bg
    white   = colors.HexColor('#ffffff'),
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
    """Build a styled dark table."""
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
    """Single-cell table used as a coloured note box."""
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
    # Full-page dark background
    canvas.setFillColor(C['bg'])
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # Top bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, H - 9 * mm, W, 9 * mm, fill=1, stroke=0)
    canvas.setFont('Sans-Bold', 8)
    canvas.setFillColor(C['blue'])
    canvas.drawString(15 * mm, H - 5.5 * mm,
                      'NTE200 №46  •  QSK50 MCRS V16  •  Диагностический отчёт  •  06.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C['sub'])
    canvas.drawRightString(W - 15 * mm, H - 5.5 * mm, f'Стр. {doc.page}')
    # Bottom bar
    canvas.setFillColor(C['panel'])
    canvas.rect(0, 0, W, 7 * mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C['sub'])
    canvas.drawString(15 * mm, 2.5 * mm,
                      'Конфиденциально  •  Горная Евразия  •  Наработка: 18 827 м/ч')
    canvas.restoreState()

# ── Document setup ─────────────────────────────────────────────────────────────
OUT = '/home/user/NTE200/NTE200_46_report_v3_06062026.pdf'
doc = SimpleDocTemplate(
    OUT,
    pagesize=A4,
    leftMargin=13 * mm,
    rightMargin=13 * mm,
    topMargin=16 * mm,
    bottomMargin=11 * mm,
    title='NTE200 №46 — Диагностика QSK50 MCRS V16 — 06.06.2026',
)
PW = doc.width   # usable width
story = []

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — TITLE + KPI CARDS
# ═══════════════════════════════════════════════════════════════════════════════
story += [
    SP(2),
    Pc(C['white'], 'ДИАГНОСТИЧЕСКИЙ ОТЧЁТ NTE200 #46 — QSK50 MCRS V16',
       bold=True, size=14, lead=18, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],
       'Дата съёмки: 06.06.2026  |  Наработка машины: ~18 827 м/ч  |  ЭБУ: 11 016 м/ч (после замены)',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# KPI cards — 4 cards side by side
kpi_cards = [
    ('МАШИНО-ЧАСЫ', '18 827 м/ч', C['blue'],   C['bluebg']),
    ('HEALTH SCORE', '42 / 100',  C['red'],    C['redbg']),
    ('ФОРСУНКИ ОК',  '3 из 16',   C['red'],    C['redbg']),
    ('МКЛ СТАТУС',   'ЗАМЕНА\nВЫПОЛНЕНА', C['orange'], C['orangebg']),
]

kpi_w = [PW / 4] * 4
kpi_header_row = [Pc(C['sub'], label, bold=True, size=8, lead=11, align=TA_CENTER)
                  for label, _, _, _ in kpi_cards]
kpi_value_row  = [Pc(val_c, val, bold=True, size=15, lead=19, align=TA_CENTER)
                  for _, val, val_c, _ in kpi_cards]

kpi_t = Table([kpi_header_row, kpi_value_row], colWidths=kpi_w)
kpi_ts = TableStyle([
    ('GRID',        (0, 0), (-1, -1), 0.5, C['border']),
    ('TOPPADDING',  (0, 0), (-1, -1), 7),
    ('BOTTOMPADDING',(0,0), (-1, -1), 7),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING',(0, 0), (-1, -1), 4),
    ('ALIGN',       (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
])
for col_i, (_, _, _, bg) in enumerate(kpi_cards):
    kpi_ts.add('BACKGROUND', (col_i, 0), (col_i, 0), C['hdr'])
    kpi_ts.add('BACKGROUND', (col_i, 1), (col_i, 1), bg)
kpi_t.setStyle(kpi_ts)
story += [kpi_t, SP(4)]

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — OIL ANALYSIS ДВС
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС — 40 ПРОБ (апр 2025 — май 2026)', 2)
story += [SP(1)]

story.append(info_box(
    'Общий вывод по ДВС: Масло в ДОПУСТИМОМ состоянии во всех пробах.\n'
    'Fe: 0-10 мг/кг (норма) — износ поверхностей трения минимален.\n'
    'TBN снижение: с 11.5 (апр 2025) до 9.6-10.5 (к. 2025) — нормальное истощение щелочного резерва.\n'
    'TAN: 0.7 → 3.7 (пик ноябрь 2025) → 2.2-3.0 — кислотное число в норме.\n'
    'Вязкость V100: 14.4-17.2 сСт — соответствует спецификации.',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

dvs_cols = ['Дата', 'М/ч', 'Н.масла', 'Fe', 'Cu', 'Cr', 'Al', 'V100', 'TAN', 'TBN', 'Сажа', 'Статус']
dvs_w = [PW*0.10, PW*0.08, PW*0.08, PW*0.05, PW*0.05, PW*0.05, PW*0.05,
         PW*0.07, PW*0.06, PW*0.06, PW*0.07, PW*0.14]
story.append(hdr_row(dvs_cols, dvs_w))

dvs_data = [
    ('2026-05-20', '18379', '132',  '0',  '0', '0', '0', '14.6', '2.7', '9.9',  '0.3', 'Допустимое', C['greenbg']),
    ('2026-05-14', '18247', '254',  '4',  '0', '0', '0', '14.4', '3.0', '10.3', '0.3', 'Допустимое', C['greenbg']),
    ('2026-05-03', '17993', '247',  '0',  '0', '0', '0', '14.5', '2.4', '11.5', '0.4', 'Допустимое', C['greenbg']),
    ('2026-04-10', '17746', '243',  '0',  '0', '0', '0', '14.7', '2.2', '9.7',  '0.3', 'Допустимое', C['greenbg']),
    ('2026-03-30', '17493', '238',  '0',  '0', '0', '0', '14.6', '3.0', '11.5', '0.2', 'Допустимое', C['greenbg']),
    ('2026-03-25', '17378', '123',  '0',  '0', '0', '0', '14.6', '2.2', '10.5', '0.3', 'Допустимое', C['greenbg']),
    ('2026-02-04', '16249', '254',  '4',  '0', '6', '5', '15.9', '2.1', '9.6',  '0.3', 'Допустимое', C['greenbg']),
    ('2026-01-24', '16003', '249',  '3',  '0', '2', '2', '16.1', '3.1', '9.4',  '0.3', 'Допустимое', C['greenbg']),
]

dvs_rows = []
for row in dvs_data:
    *vals, status_txt, row_bg = row
    cells = []
    for i, v in enumerate(vals):
        col = C['green'] if i == 11 else C['text']
        bold = (i == 11)
        cells.append(Pc(col, v, bold=bold, size=7, lead=10))
    dvs_rows.append(cells)

dvs_extra = [('BACKGROUND', (0, i), (-1, i), dvs_data[i][12]) for i in range(len(dvs_rows))]
story.append(tbl(dvs_rows, dvs_w, style_extra=dvs_extra, alt=False))
story.append(SP(3))

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — МКЛ OIL ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА МКЛ (ЛЕВЫЙ) — ХРОНИКА ОТКАЗА', 3)
story.append(SP(1))

story.append(info_box(
    'КАТАСТРОФИЧЕСКИЙ ОТКАЗ МКЛ: 10.04.2026 при 17 746 м/ч\n'
    'Fe=593 мг/кг, Cu=29 мг/кг, Al=86 мг/кг, Si=476 мг/кг\n'
    '11.04.2026 — ЗАМЕНА МКЛ (с а/с №48). Новый МКЛ работает нормально.',
    bg_color=C['redbg'], text_color=C['red'], bold=True, size=8, border_color=C['red']
))
story.append(SP(2))

mkl_cols = ['Дата', 'М/ч', 'Fe', 'Cu', 'Cr', 'Al', 'Si', 'Статус / Примечание']
mkl_w = [PW*0.10, PW*0.08, PW*0.07, PW*0.06, PW*0.06, PW*0.06, PW*0.07, PW*0.50]
story.append(hdr_row(mkl_cols, mkl_w))

mkl_data = [
    ('2025-05-01', '10007',  '50',  '0', '0',  '0',  '5',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2025-06-17', '11002', '401',  '6', '8',  '5', '64',   'Допустимое [высокое Fe, смена масла]', C['panel'], C['green']),
    ('2025-07-09', '11489', '177',  '0', '0',  '0', '15',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2025-07-31', '12000', '247',  '0', '0',  '0', '22',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2025-09-12', '13001', '239',  '0', '0',  '0', '13',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2025-10-04', '13497', '435',  '2', '3',  '9', '25',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2025-11-18', '14505', '390',  '2', '0', '19', '54',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2026-01-03', '15506', '345',  '0', '0', '15', '77',   'Требует мер',                   C['orangebg'], C['orange']),
    ('2026-01-24', '16003', '493',  '6', '9', '29', '96',   'Требует мер — КРИТИЧНО!',       C['redbg'],    C['red']),
    ('2026-02-23', '16685', '103',  '0', '0',  '4', '17',   'Допустимое [после замены масла]', C['greenbg'], C['green']),
    ('2026-03-06', '16939', '116',  '0', '0',  '2', '18',   'Допустимое',                    C['greenbg'], C['green']),
    ('2026-03-30', '17505', '439',  '6', '0', '25','108',   'Требует мер — РОСТ!',           C['redbg'],    C['red']),
    ('2026-04-10', '17746', '593', '29', '0', '86','476',   'КРИТИЧНО — ЗАМЕНА МКЛ',         C['redbg'],    C['red']),
    ('2026-05-03', '17993',   '4',  '0', '0',  '0',  '1',   'Допустимое [новый МКЛ]',        C['greenbg'], C['green']),
    ('2026-05-14', '18247',   '8',  '0', '0',  '0',  '3',   'Допустимое [норма]',            C['greenbg'], C['green']),
]

mkl_rows = []
for row in mkl_data:
    date, mh, fe, cu, cr, al, si, status_txt, row_bg, status_col = row
    mkl_rows.append([
        Pc(C['text'],  date,       size=7, lead=10),
        Pc(C['text'],  mh,         size=7, lead=10),
        Pc(status_col, fe, bold=(int(fe) > 300), size=7, lead=10),
        Pc(status_col if int(cu) > 5 else C['text'], cu, size=7, lead=10),
        Pc(C['text'],  cr,         size=7, lead=10),
        Pc(status_col if int(al) > 20 else C['text'], al, size=7, lead=10),
        Pc(status_col if int(si) > 50 else C['text'], si, bold=(int(si) > 100), size=7, lead=10),
        Pc(status_col, status_txt, bold=True, size=7, lead=10),
    ])

mkl_extra = [('BACKGROUND', (0, i), (-1, i), mkl_data[i][8]) for i in range(len(mkl_rows))]
story.append(tbl(mkl_rows, mkl_w, style_extra=mkl_extra, alt=False))

story += [SP(1),
    info_box(
        'Si=476 мг/кг — массовый выход из строя уплотнений (Si-содержащий герметик/уплотнение).\n'
        'Cu=29 мг/кг — износ подшипников (медный сплав).\n'
        'Al=86 мг/кг — эрозия корпуса (алюминий).',
        bg_color=C['bluebg'], text_color=C['blue'], size=8
    ),
    SP(3),
]

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — МКП STATUS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('МКЛ ПРАВЫЙ (МКП) — МОНИТОРИНГ', 4)
story.append(SP(1))

mkp_cols = ['Дата', 'М/ч', 'Fe', 'Cu', 'Si', 'Статус']
mkp_w = [PW*0.13, PW*0.10, PW*0.10, PW*0.10, PW*0.10, PW*0.47]
story.append(hdr_row(mkp_cols, mkp_w))

mkp_data = [
    ('2025-05-01', '10007',  '27', '0',  '5', 'Допустимое',       C['panel'],    C['green']),
    ('2025-07-09', '11489',  '59', '0',  '5', 'Требует мер',      C['orangebg'], C['orange']),
    ('2025-10-04', '13497', '141', '0', '13', 'Допустимое',       C['panel'],    C['green']),
    ('2026-01-24', '16003', '185', '0', '61', 'Требует мер',      C['orangebg'], C['orange']),
    ('2026-03-08', '16996',  '91', '0', '32', 'Допустимое',       C['panel'],    C['green']),
    ('2026-03-30', '17505', '178', '0', '51', 'Требует мер',      C['orangebg'], C['orange']),
    ('2026-05-03', '17993', '151', '0', '40', 'Требует мер — ПОСЛЕДНЯЯ ПРОБА', C['orangebg'], C['orange']),
]

mkp_rows = []
for date, mh, fe, cu, si, status_txt, row_bg, status_col in mkp_data:
    mkp_rows.append([
        Pc(C['text'],  date,       size=8, lead=11),
        Pc(C['text'],  mh,         size=8, lead=11),
        Pc(status_col, fe, bold=(int(fe) > 150), size=8, lead=11),
        Pc(C['text'],  cu,         size=8, lead=11),
        Pc(status_col if int(si) > 30 else C['text'], si, size=8, lead=11),
        Pc(status_col, status_txt, bold=True, size=8, lead=11),
    ])

mkp_extra = [('BACKGROUND', (0, i), (-1, i), mkp_data[i][6]) for i in range(len(mkp_rows))]
story.append(tbl(mkp_rows, mkp_w, style_extra=mkp_extra, alt=False))

story += [SP(1),
    info_box(
        'МКП — последняя проба 03.05.2026: Fe=151, Si=40 — "Требует принятия мер".\n'
        'Рекомендуется отбор проб и контроль. При Fe>200 — рассмотреть замену.',
        bg_color=C['orangebg'], text_color=C['orange'], bold=True, size=8, border_color=C['orange']
    ),
    SP(3),
]

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — INJECTOR BENCH TEST
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ТЕСТ ФОРСУНОК НА СТЕНДЕ ПОТОК CR-2 (05.06.2026)', 5)
story += [SP(1),
    Pc(C['sub'],
       'Деталь: 2867147 CUMMINS MCRS  |  Дата замены ЭБУ/двигателя: ~март 2026',
       size=8, lead=11),
    SP(1),
]

# Summary line
inj_summary_cols = [PW * 0.33] * 3
inj_summary_data = [
    [Pc(C['green'],  '3  НОРМА',      bold=True, size=13, lead=17, align=TA_CENTER),
     Pc(C['orange'], '8  НЕЗНАЧИТ.',  bold=True, size=13, lead=17, align=TA_CENTER),
     Pc(C['red'],    '5  КРИТИЧНО',   bold=True, size=13, lead=17, align=TA_CENTER)],
]
inj_sum_t = Table(inj_summary_data, colWidths=inj_summary_cols)
inj_sum_t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (0, 0), C['greenbg']),
    ('BACKGROUND', (1, 0), (1, 0), C['orangebg']),
    ('BACKGROUND', (2, 0), (2, 0), C['redbg']),
    ('GRID', (0, 0), (-1, -1), 0.5, C['border']),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('TOPPADDING', (0, 0), (-1, -1), 8),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
]))
story += [inj_sum_t, SP(2)]

# Injector table
inj_cols = ['Цил.', 'P01 (441-496)', 'P02 (220-260)', 'P03 (16-30)', 'P04 (41-51)', 'Оценка']
inj_w = [PW*0.07, PW*0.20, PW*0.20, PW*0.20, PW*0.20, PW*0.13]
story.append(hdr_row(inj_cols, inj_w, size=7))

# (cyl, P01_val, P01_note, P02_val, P02_note, P03_val, P03_note, P04_val, P04_note, result, result_col, row_bg)
inj_data = [
    ('1L',  '501.4', '+1.2%↑', '270.5', '+4.0%↑', '35.2', '+17%↑',  '48.1',  'OK',     'КРИТИЧНО', C['red'],    C['redbg']),
    ('1R',  '470.7', 'OK',     '248.8', 'OK',      '20.0', 'OK',      '41.8',  'OK',     'НОРМА',    C['green'],  C['greenbg']),
    ('2L',  '491.4', 'OK',     '261.1', '+0.4%↑',  '25.1', 'OK',      '34.9',  '-15%↓',  'ОТКЛ.',    C['orange'], C['orangebg']),
    ('2R',  '392.8', '-11%↓',  '214.8', '-2.4%↓',  '23.2', 'OK',      '41.1',  'OK',     'КРИТ.',    C['red'],    C['redbg']),
    ('3L',  '451.8', 'OK',     '189.9', '-14%↓',   '19.0', 'OK',      '26.9',  '-34%↓',  'КРИТ.',    C['red'],    C['redbg']),
    ('3R',  '470.9', 'OK',     '245.6', 'OK',       '21.6', 'OK',      '43.9',  'OK',     'НОРМА',    C['green'],  C['greenbg']),
    ('4L',  '472.7', 'OK',     '245.9', 'OK',       '18.2', 'OK',      '42.1',  'OK',     'НОРМА',    C['green'],  C['greenbg']),
    ('4R',  '521.0', '+5.1%↑', '271.4', '+4.4%↑',  '41.3', '+38%↑',  '51.3',  '+0.6%↑', 'КРИТ.',    C['red'],    C['redbg']),
    ('5L',  '492.8', 'OK',     '262.8', '+1.1%↑',  '40.2', '+34%↑',  '52.6',  '+3.1%↑', 'КРИТ.',    C['red'],    C['redbg']),
    ('5R',  '495.9', '+0.1%↑', '248.9', 'OK',       '28.0', 'OK',      '40.5',  '-1.2%↓', 'ОТКЛ.',    C['orange'], C['orangebg']),
    ('6L',  '497.9', '+0.5%↑', '224.9', 'OK',       '22.8', 'OK',      '34.6',  '-16%↓',  'ОТКЛ.',    C['orange'], C['orangebg']),
    ('6R',  '503.8', '+1.7%↑', '255.3', 'OK',       '29.5', 'OK',      '41.4',  'OK',     'ОТКЛ.',    C['orange'], C['orangebg']),
    ('7L',  '497.6', '+0.4%↑', '272.6', '+4.8%↑',  '34.7', '+16%↑',  '52.6',  '+3.1%↑', 'КРИТ.',    C['red'],    C['redbg']),
    ('7R',  '416.7', '-5.6%↓', '243.3', 'OK',       '39.2', '+31%↑',  '52.7',  '+3.3%↑', 'ОТКЛ.',    C['orange'], C['orangebg']),
    ('8L',  '500.3', '+1.0%↑', '256.9', 'OK',       '38.4', '+28%↑',  '50.5',  'OK',     'ОТКЛ.',    C['orange'], C['orangebg']),
    ('8R',  '497.1', '+0.3%↑', '247.1', 'OK',       '32.0', '+7%↑',   '42.4',  'OK',     'ОТКЛ.',    C['orange'], C['orangebg']),
]

def inj_val_cell(val, note):
    note_col = C['red'] if '↑' in note or '↓' in note else C['sub']
    note_bold = note not in ('OK', '')
    txt = val
    if note != 'OK' and note:
        txt = f'{val}  <font color="#{note_col.hexval()[2:]}" size="6">{note}</font>'
    return Pc(C['text'] if note == 'OK' else C['orange'] if '↓' in note or '↑' in note else C['text'],
              txt, size=7, lead=10)

inj_rows = []
for row in inj_data:
    cyl, p01, p01n, p02, p02n, p03, p03n, p04, p04n, result, result_col, row_bg = row
    p01_col = C['red'] if '%' in p01n else C['text']
    p02_col = C['red'] if '%' in p02n else C['text']
    p03_col = C['red'] if '%' in p03n else C['text']
    p04_col = C['red'] if '%' in p04n else C['text']

    def fmt(val, note):
        if note == 'OK':
            return Pc(C['text'], val, size=7, lead=10)
        arrow = ' ↑' if '↑' in note else ' ↓' if '↓' in note else ''
        nc = C['red'] if ('↑' in note and 'P03' not in str(val)) or '↓' in note else C['orange']
        return Pc(C['orange'], f'{val} ({note})', size=7, lead=10)

    inj_rows.append([
        Pc(C['text'], cyl, bold=True, size=7, lead=10),
        fmt(p01, p01n),
        fmt(p02, p02n),
        fmt(p03, p03n),
        fmt(p04, p04n),
        Pc(result_col, result, bold=True, size=7, lead=10),
    ])

inj_extra = [('BACKGROUND', (0, i), (-1, i), inj_data[i][11]) for i in range(len(inj_rows))]
story.append(tbl(inj_rows, inj_w, style_extra=inj_extra, alt=False))

story += [SP(1),
    info_box(
        'P01 высокий (8 цил.): увеличение распылительных отверстий — износ сопла.\n'
        'P01 низкий (2 цил. — 2R, 7R): залипание иглы / закоксованность — угроза перегрева ГБЦ.\n'
        'P03 высокий (7 цил.): износ управляющего клапана — перетечки топлива.\n'
        'P04 низкий (4 цил.): потеря давления открытия — деградация возвратного клапана.',
        bg_color=C['bluebg'], text_color=C['blue'], size=8
    ),
    SP(3),
]

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — ENGINE PARAMETERS FROM DML
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ПАРАМЕТРЫ ДВС — DML 04.06.2026 и 06.06.2026', 6)
story += [SP(1),
    Pc(C['sub'],
       '04.06.2026 15:27-15:34 — стоянка/диагностика (442 строки, ~7 мин)  |  '
       '06.06.2026 — гружёный рейс 18 мин', size=8, lead=11),
    SP(1),
]

dml_cols = ['Параметр', 'Значение', 'Норма', 'Статус']
dml_w = [PW*0.38, PW*0.22, PW*0.22, PW*0.18]
story.append(hdr_row(dml_cols, dml_w))

dml_data = [
    ('RPM средн.',              '1 640 об/мин',  '1500-1800',    'OK',         C['greenbg'], C['green']),
    ('T охл. жидк. макс',      '99.2 °C',        '<95 °C',       'ПРЕДУПР.',   C['orangebg'],C['orange']),
    ('T масла макс',            '110.3 °C',       '<105 °C',      'ПРЕВЫШЕНИЕ', C['orangebg'],C['orange']),
    ('P масла мин.',            '380 кПа',        '>350 кПа',     'OK',         C['greenbg'], C['green']),
    ('T выхл. левый банк',      '495 °C ср.',     '400-520 °C',   'OK',         C['greenbg'], C['green']),
    ('T выхл. правый банк',     '487 °C ср.',     '400-520 °C',   'OK',         C['greenbg'], C['green']),
    ('Цил. 3 EGT',              '600 °C → 0 °C',  '—',            'ОБРЫВ ДАТЧ.',C['redbg'],   C['red']),
    ('P наддува',               '285 кПа',        '220-320 кПа',  'OK',         C['greenbg'], C['green']),
    ('Расход топлива',          '142 л/ч',        '120-160 л/ч',  'OK',         C['greenbg'], C['green']),
    ('P картера',               '2.1 кПа',        '<3.0 кПа',     'OK',         C['greenbg'], C['green']),
]

dml_rows = []
for param, val, norm, status_txt, row_bg, status_col in dml_data:
    dml_rows.append([
        Pc(C['text'],   param,      size=8, lead=11),
        Pc(status_col,  val,  bold=True, size=8, lead=11),
        Pc(C['sub'],    norm,       size=8, lead=11),
        Pc(status_col,  status_txt, bold=True, size=8, lead=11),
    ])

dml_extra = [('BACKGROUND', (0, i), (-1, i), dml_data[i][4]) for i in range(len(dml_rows))]
story.append(tbl(dml_rows, dml_w, style_extra=dml_extra, alt=False))

story += [SP(1),
    info_box(
        'Цилиндр 3 (2L): датчик EGT — сигнатура обрыва термопары '
        '(насыщение 600 °C → спад до нуля). Цилиндр без температурного контроля.',
        bg_color=C['redbg'], text_color=C['red'], size=8, border_color=C['red']
    ),
    SP(3),
]

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — FAULT / MAINTENANCE HISTORY
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ХРОНОЛОГИЯ НЕИСПРАВНОСТЕЙ И ОБСЛУЖИВАНИЯ', 7)
story.append(SP(1))

hist_cols = ['Дата', 'М/ч', 'Узел', 'Событие']
hist_w = [PW*0.12, PW*0.09, PW*0.13, PW*0.66]
story.append(hdr_row(hist_cols, hist_w))

hist_data = [
    ('02.08.2025', '12039',  'ДВС',       'Регулировка клапанов (ТО-3+)',                                                              C['panel'],    C['text']),
    ('05.08.2025', '12112',  'МКЛ',       'Замена масла МКЛ (плохие анализы)',                                                         C['orangebg'], C['orange']),
    ('04.10.2025', '13497',  'ДВС',       'Диагностика и ремонт ДВС — шумы, замена осей коромысла',                                    C['orangebg'], C['orange']),
    ('11.12.2025', '15030',  'ДВС',       'Замена 2 форсунок',                                                                        C['orangebg'], C['orange']),
    ('15.12.2025', '15103',  'ДВС',       'Сизый дым, замена клапанов 2R и коромысла 5L',                                             C['redbg'],    C['red']),
    ('24.01.2026', '16003',  'МКЛ',       'Fe=493 — критически высокий',                                                              C['redbg'],    C['red']),
    ('11.02.2026', '16400',  'МКЛ',       'Отбор масла МКЛ лев. (плохие анализы), долив ДВС 30 л',                                   C['orangebg'], C['orange']),
    ('23.02.2026', '16676',  'МКЛ',       'Замена масла МКЛ левый',                                                                   C['orangebg'], C['orange']),
    ('25.03.2026', '17378',  'ДВС',       'Отбор проб масла ДВС',                                                                    C['panel'],    C['text']),
    ('30.03.2026', '17493',  'ДВС/МКЛ',   'Отбор проб масла (МКЛ: Fe=439, Si=108 — рост!)',                                          C['orangebg'], C['orange']),
    ('10.04.2026', '17746',  'МКЛ',       'МКЛ лев: Fe=593, Cu=29, Si=476 — КАТАСТРОФА',                                             C['redbg'],    C['red']),
    ('11.04.2026', '17767',  'МКЛ',       'Замена МКЛ левый (с а/с №48)',                                                             C['greenbg'],  C['green']),
    ('27.04.2026', '17850',  'ДВС',       'Замена датчика давления маслоподкачки',                                                    C['panel'],    C['text']),
    ('03.05.2026', '17993',  'МКП',       'Fe=151, Si=40 — требует внимания',                                                        C['orangebg'], C['orange']),
    ('04.06.2026', '18827',  'ДВС',       'Диагностика CAMSS. DML-запись (15:27-15:34)',                                              C['bluebg'],   C['blue']),
    ('05.06.2026', '~18827', 'ДВС',       'Тест 16 форсунок на стенде Поток CR-2',                                                   C['bluebg'],   C['blue']),
    ('06.06.2026', '18700',  'ДВС',       'DML гружёный рейс — мониторинг',                                                          C['bluebg'],   C['blue']),
]

hist_rows = []
for date, mh, unit, event, row_bg, event_col in hist_data:
    hist_rows.append([
        Pc(event_col, date,  bold=True, size=7, lead=10),
        Pc(C['sub'],  mh,              size=7, lead=10),
        Pc(event_col, unit,  bold=True, size=7, lead=10),
        Pc(C['text'],  event,           size=7, lead=10),
    ])

hist_extra = [('BACKGROUND', (0, i), (-1, i), hist_data[i][4]) for i in range(len(hist_rows))]
story.append(tbl(hist_rows, hist_w, style_extra=hist_extra, alt=False))
story.append(SP(3))

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — CONCLUSIONS AND RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ВЫВОДЫ И РЕКОМЕНДАЦИИ', 8)
story.append(SP(1))

# RED findings
red_findings = [
    ('1. ФОРСУНКИ (13 из 16 не в норме)',
     'Требуется замена всего комплекта (16 шт., арт. 2867147). '
     'Цилиндры 4R, 7L — 4 параметра вне нормы. '
     'Закоксованность 2R, 7R — угроза перегрева ГБЦ.'),
    ('2. МКП ПРАВЫЙ',
     'Последняя проба 03.05.2026 Fe=151, Si=40 — "Требует принятия мер". '
     'Контроль, возможная замена при Fe>200.'),
]
for title, body in red_findings:
    story.append(info_box(
        f'{title}\n{body}',
        bg_color=C['redbg'], text_color=C['red'], bold=False, size=8, border_color=C['red']
    ))
    story.append(SP(1))

# ORANGE findings
orange_findings = [
    ('3. ДАТЧИК EGT ЦИЛ. 3 (2L)',
     'Обрыв термопары — замена датчика.'),
    ('4. ТЕМПЕРАТУРА МАСЛА',
     'Пик 110°C при нагрузке — диагностика масляного теплообменника.'),
    ('5. МАСЛО ДВС (TAN)',
     'Кислотное число достигало 3.7 — контроль при следующей смене.'),
]
for title, body in orange_findings:
    story.append(info_box(
        f'{title}\n{body}',
        bg_color=C['orangebg'], text_color=C['orange'], bold=False, size=8, border_color=C['orange']
    ))
    story.append(SP(1))

# GREEN findings
green_findings = [
    ('6. МКЛ ЛЕВЫЙ — ЗАМЕНА ВЫПОЛНЕНА',
     'Замена выполнена 11.04.2026. Новый МКЛ в норме (Fe=4-8).'),
    ('7. МАСЛО ДВС',
     'Все 40 проб — "Допустимое". Продолжить штатный мониторинг.'),
]
for title, body in green_findings:
    story.append(info_box(
        f'{title}\n{body}',
        bg_color=C['greenbg'], text_color=C['green'], bold=False, size=8, border_color=C['green']
    ))
    story.append(SP(1))

story.append(SP(1))

# Priority action table
priority_data = [
    ('ПРИОРИТЕТ 1\n(немедленно)',
     'Замена 16 форсунок арт. 2867147. Снять ДВС из эксплуатации при первом плановом ТО.',
     C['red'], C['redbg']),
    ('ПРИОРИТЕТ 2\n(ближайшие 250 м/ч)',
     'Контрольная проба масла МКП, при Fe>200 — замена.',
     C['orange'], C['orangebg']),
    ('ПРИОРИТЕТ 3',
     'Замена датчика EGT цил. 3 (2L).',
     C['orange'], C['orangebg']),
    ('ПРИОРИТЕТ 4',
     'Диагностика масляного теплообменника.',
     C['blue'], C['bluebg']),
]

prio_w = [PW * 0.22, PW * 0.78]
prio_rows = []
for prio_label, prio_text, prio_col, prio_bg in priority_data:
    prio_rows.append([
        Pc(prio_col, prio_label, bold=True, size=9, lead=13),
        Pc(C['text'], prio_text, size=9, lead=13),
    ])
prio_extra = [
    ('BACKGROUND', (0, i), (-1, i), priority_data[i][3])
    for i in range(len(priority_data))
]
story.append(tbl(prio_rows, prio_w, style_extra=prio_extra, alt=False))
story.append(SP(3))

# ── Footer ─────────────────────────────────────────────────────────────────────
story.append(HR())
story.append(SP(1))
story.append(Pc(C['sub'],
    'Отчёт подготовлен на основе: DML-20260604 / DML-20260606 (INSITE Professional) · '
    'Стенд Поток CR-2 (16 форсунок, 05.06.2026) · '
    'Спектральный анализ масла ДВС (40 проб) · '
    'Спектральный анализ МКЛ/МКП · '
    'Сводный анализ NTE200 · База ГБЦ ремонтов  |  Горная Евразия, 06.06.2026',
    size=7, lead=10))

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
