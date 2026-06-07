#!/usr/bin/env python3
"""Generate NTE200 #46 trip analysis PDF report."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os, sys

# ── Fonts ──────────────────────────────────────────────────────────────────
FONT_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
]
for p in FONT_PATHS:
    if not os.path.exists(p):
        sys.exit(f'Font not found: {p}')

pdfmetrics.registerFont(TTFont('Sans', FONT_PATHS[0]))
pdfmetrics.registerFont(TTFont('Sans-Bold', FONT_PATHS[1]))

# ── Colours ────────────────────────────────────────────────────────────────
C_DARK    = colors.HexColor('#0d1117')
C_PANEL   = colors.HexColor('#161b22')
C_BORDER  = colors.HexColor('#30363d')
C_ACCENT  = colors.HexColor('#3EF0AF')
C_RED     = colors.HexColor('#ff4060')
C_ORANGE  = colors.HexColor('#FBAE40')
C_GREEN   = colors.HexColor('#3EF0AF')
C_BLUE    = colors.HexColor('#38bdf8')
C_TEXT    = colors.HexColor('#e6edf3')
C_SUBTEXT = colors.HexColor('#8b949e')
C_HEADER  = colors.HexColor('#1f2937')
C_ROW_ALT = colors.HexColor('#1a2332')

W, H = A4

def style(name='Normal', **kw):
    base = ParagraphStyle(name,
        fontName='Sans', fontSize=9, textColor=C_TEXT,
        leading=13, **kw)
    return base

S_TITLE    = ParagraphStyle('title',  fontName='Sans-Bold', fontSize=20, textColor=C_ACCENT,
                            leading=26, alignment=TA_LEFT)
S_SUBTITLE = ParagraphStyle('sub',    fontName='Sans', fontSize=11, textColor=C_SUBTEXT,
                            leading=16, alignment=TA_LEFT)
S_H2       = ParagraphStyle('h2',     fontName='Sans-Bold', fontSize=13, textColor=C_TEXT,
                            leading=18, spaceAfter=4)
S_H3       = ParagraphStyle('h3',     fontName='Sans-Bold', fontSize=10, textColor=C_ACCENT,
                            leading=14, spaceBefore=6, spaceAfter=2)
S_BODY     = ParagraphStyle('body',   fontName='Sans', fontSize=9, textColor=C_TEXT,
                            leading=14, spaceAfter=3)
S_CAPTION  = ParagraphStyle('cap',    fontName='Sans', fontSize=8, textColor=C_SUBTEXT,
                            leading=11)
S_ALERT    = ParagraphStyle('alert',  fontName='Sans-Bold', fontSize=9, textColor=C_RED,
                            leading=13)
S_WARN     = ParagraphStyle('warn',   fontName='Sans-Bold', fontSize=9, textColor=C_ORANGE,
                            leading=13)
S_OK       = ParagraphStyle('ok',     fontName='Sans-Bold', fontSize=9, textColor=C_GREEN,
                            leading=13)

def hr():
    return HRFlowable(width='100%', thickness=0.5, color=C_BORDER, spaceAfter=6, spaceBefore=6)

def sp(h=4):
    return Spacer(1, h*mm)

def hdr_table(cells, col_widths, bg=C_HEADER):
    data = [[Paragraph(f'<font name="Sans-Bold" color="{C_SUBTEXT.hexval()}">{c}</font>', S_CAPTION)
             for c in cells]]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [bg]),
        ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    return t

def data_table(rows, col_widths, alt=True):
    tdata = []
    for i, row in enumerate(rows):
        cells = []
        for cell in row:
            if isinstance(cell, str):
                cells.append(Paragraph(cell, S_BODY))
            else:
                cells.append(cell)
        tdata.append(cells)
    bg_even = C_PANEL
    bg_odd  = C_ROW_ALT if alt else C_PANEL
    t = Table(tdata, colWidths=col_widths)
    ts = TableStyle([
        ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ])
    for i in range(len(rows)):
        ts.add('BACKGROUND', (0,i), (-1,i), bg_even if i%2==0 else bg_odd)
    t.setStyle(ts)
    return t

def pill(text, bg, fg=C_DARK):
    return Paragraph(
        f'<font name="Sans-Bold" color="{fg.hexval() if hasattr(fg,"hexval") else fg}">{text}</font>',
        ParagraphStyle('pill', fontName='Sans-Bold', fontSize=8,
                       backColor=bg, textColor=fg if not hasattr(fg,'hexval') else fg,
                       leading=10, borderPadding=2))

# ── BUILD PDF ──────────────────────────────────────────────────────────────
OUT = '/home/user/NTE200/NTE200_46_report_06062026.pdf'
doc = SimpleDocTemplate(OUT, pagesize=A4,
    leftMargin=15*mm, rightMargin=15*mm,
    topMargin=15*mm, bottomMargin=15*mm,
    title='NTE200 №46 — Анализ рейса 06.06.2026')

story = []
PW = doc.width  # usable width ≈ 180mm

# ── COVER BLOCK ───────────────────────────────────────────────────────────
story.append(Paragraph('ТЕХНИЧЕСКИЙ ОТЧЁТ', S_SUBTITLE))
story.append(sp(1))
story.append(Paragraph('NTE200 №46 — Анализ рейса с породой', S_TITLE))
story.append(sp(2))
story.append(Paragraph('06 июня 2026 г.  •  15:15 – 15:33  •  18 минут  •  Наработка ДВС: 11 016 ч', S_SUBTITLE))
story.append(sp(1))
story.append(Paragraph('Двигатель Cummins QSK50 MCRS V16  •  Самосвал NTE200  •  Полная нагрузка', S_CAPTION))
story.append(hr())
story.append(sp(2))

# ── SECTION 1: KPI CARDS ─────────────────────────────────────────────────
story.append(Paragraph('1. Ключевые показатели рейса', S_H2))
story.append(sp(1))

kpi_cols = [PW*0.16]*6
kpi_headers = ['Нагрузка ДВС', 'Обороты', 'Расход топлива', 'T охл.жидк.', 'T масла (макс)', 'P масла']
kpi_values  = ['94 %', '1 876 об/мин', '339 л/ч', '85.7 °C', '110.0 °C', '516 кПа']
kpi_status  = [C_GREEN, C_GREEN, C_BLUE, C_GREEN, C_ORANGE, C_GREEN]

hrow = [Paragraph(f'<font name="Sans-Bold" color="{C_SUBTEXT.hexval()}">{h}</font>', S_CAPTION)
        for h in kpi_headers]
vrow = [Paragraph(f'<font name="Sans-Bold" color="{s.hexval()}">{v}</font>',
                  ParagraphStyle('kv', fontName='Sans-Bold', fontSize=12,
                                 textColor=s, leading=16, alignment=TA_CENTER))
        for v, s in zip(kpi_values, kpi_status)]

kpi_t = Table([hrow, vrow], colWidths=kpi_cols)
kpi_t.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), C_PANEL),
    ('GRID', (0,0), (-1,-1), 0.5, C_BORDER),
    ('TOPPADDING',    (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('LEFTPADDING',   (0,0), (-1,-1), 4),
    ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('BACKGROUND', (4,1), (4,1), colors.HexColor('#2a1f0a')),
]))
story.append(kpi_t)
story.append(sp(3))

# ── SECTION 2: PARAMETER STATUS ───────────────────────────────────────────
story.append(Paragraph('2. Состояние параметров двигателя', S_H2))
story.append(sp(1))

def status_cell(icon, txt, s):
    return Paragraph(f'<font name="Sans-Bold" color="{s.hexval()}">{icon} {txt}</font>',
                     ParagraphStyle('sc', fontName='Sans-Bold', fontSize=9, textColor=s, leading=12))

param_rows = [
    ['Температура охлаждающей жидкости', '84.1 °C avg / 85.7 °C макс',
     'Норма 82–95 °C', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Температура масла ДВС', '104.0 °C avg / 110.0 °C макс',
     'Порог предупр. 105 °C', status_cell('!', 'ВНИМАНИЕ', C_ORANGE)],
    ['Давление масла', '516 кПа avg / 384 кПа мин',
     'Норма 350–650 кПа', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Давление картерных газов', '0.69 кПа avg / 1.10 кПа макс',
     'Норма < 3 кПа', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Давление наддува', '268 кПа avg / 295 кПа макс',
     'Норма 220–320 кПа', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Напряжение АКБ', '27.8 В avg', 'Норма 24–30 В', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Жёлтая / Красная лампа', 'Выкл. (0 сек)', 'Активных кодов нет', status_cell('✓', 'НОРМА', C_GREEN)],
    ['Давление топливной рейки', '1 522 бар измеренное', 'Задание: 1 523 бар, Δ avg 12 бар',
     status_cell('✓', 'НОРМА', C_GREEN)],
]
pcols = [PW*0.30, PW*0.25, PW*0.25, PW*0.20]
story.append(hdr_table(['Параметр', 'Значение', 'Диапазон', 'Статус'], pcols))
story.append(data_table(param_rows, pcols))
story.append(sp(1))
story.append(Paragraph(
    '⚠ Температура масла превышала 105 °C в течение 643 секунд из 1 087 (59% рейса). '
    'Динамика: старт 87.8 °C → середина 105.3 °C → финиш 109.3 °C и рост продолжается.',
    S_WARN))
story.append(sp(3))

# ── SECTION 3: CYLINDERS ─────────────────────────────────────────────────
story.append(Paragraph('3. Температуры выхлопа по цилиндрам', S_H2))
story.append(sp(1))
story.append(Paragraph(
    'Среднее по парку (без Ц3): 501.6 °C  •  Разброс: ±11.8 °C  '
    '(отличный результат — ЭБУ эффективно балансирует подачу)', S_BODY))
story.append(sp(1))

def cyl_status(n, avg, dev):
    if n == 3:
        return Paragraph('<font name="Sans-Bold" color="#ff4060">✗ ДАТЧИК ОТКАЗАЛ</font>', S_BODY)
    if avg > 525:
        return Paragraph('<font name="Sans-Bold" color="#FBAE40">! Повышенная</font>', S_BODY)
    if abs(dev) > 25:
        return Paragraph('<font name="Sans-Bold" color="#FBAE40">! Отклонение</font>', S_BODY)
    return Paragraph('<font name="Sans-Bold" color="#3EF0AF">✓ Норма</font>', S_BODY)

# (cyl_num, bank, avg, dev, maxT, minT)
cyls = [
    (1,'L','499.6','+13.8','524','244','✓'),
    (2,'R','513.7','+12.1','540','266','✓'),
    (3,'L','250.0','—','600','-22','✗'),
    (4,'R','495.8', '+9.9','526','227','✓'),
    (5,'L','493.3', '+7.4','517','244','✓'),
    (6,'R','530.4','+28.8','557','276','!'),
    (7,'L','492.1', '+6.3','519','248','✓'),
    (8,'R','488.9', '+3.1','528','242','✓'),
    (9,'L','497.8','+12.0','536','255','✓'),
    (10,'R','497.8','+12.0','527','242','✓'),
    (11,'L','513.5','+12.0','538','267','!'),
    (12,'R','501.7','+15.9','527','236','✓'),
    (13,'L','502.3','+16.5','526','272','✓'),
    (14,'R','479.4', '-6.4','504','246','✓'),
    (15,'L','508.5','+22.6','531','264','✓'),
    (16,'R','508.8','+22.9','534','235','✓'),
]

half = len(cyls)//2
def mk_cyl_half(cyl_list):
    rows = []
    for n, bank, avg, dev, mx, mn, st in cyl_list:
        if n == 3:
            s_col = Paragraph('<font name="Sans-Bold" color="#ff4060">✗ ДАТЧИК ОТКАЗАЛ</font>', S_CAPTION)
        elif st == '!':
            s_col = Paragraph(f'<font name="Sans-Bold" color="#FBAE40">! {avg} °C</font>', S_CAPTION)
        else:
            s_col = Paragraph(f'<font name="Sans" color="#3EF0AF">{avg} °C</font>', S_CAPTION)
        rows.append([
            Paragraph(f'<font name="Sans-Bold" color="#e6edf3">Ц{n} ({bank}-банк)</font>', S_CAPTION),
            s_col,
            Paragraph(f'<font color="#8b949e">{dev}</font>', S_CAPTION),
            Paragraph(f'<font color="#8b949e">{mx} / {mn}</font>', S_CAPTION),
        ])
    return rows

cw2 = [PW*0.14, PW*0.155, PW*0.075, PW*0.115]*2  # two halves side by side

# Build side-by-side table
left_rows  = mk_cyl_half(cyls[:8])
right_rows = mk_cyl_half(cyls[8:])

hdr_cyl = [Paragraph(f'<font name="Sans-Bold" color="{C_SUBTEXT.hexval()}">{h}</font>', S_CAPTION)
           for h in ['Цилиндр','Ср. T выхл.','Откл.','Макс/Мин']]

combined = []
for i in range(8):
    combined.append(left_rows[i] + right_rows[i])

cyl_col_w = [PW*0.13, PW*0.135, PW*0.07, PW*0.115,
             PW*0.13, PW*0.135, PW*0.07, PW*0.115]

hdr_row = hdr_cyl + hdr_cyl
story.append(hdr_table(
    ['Цилиндр','Ср. T выхл.','Откл.','Макс/Мин','Цилиндр','Ср. T выхл.','Откл.','Макс/Мин'],
    cyl_col_w))

ct = Table(combined, colWidths=cyl_col_w)
ts = TableStyle([
    ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
    ('LEFTPADDING',   (0,0),(-1,-1), 5),
    ('RIGHTPADDING',  (0,0),(-1,-1), 5),
    ('TOPPADDING',    (0,0),(-1,-1), 4),
    ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ('LINEBEFORE', (4,0), (4,-1), 1.0, C_BORDER),
])
for i in range(8):
    bg = C_PANEL if i%2==0 else C_ROW_ALT
    ts.add('BACKGROUND', (0,i), (-1,i), bg)
# highlight cyl 3 (row index 2) and cyl 6 (row 5)
ts.add('BACKGROUND', (0,2), (3,2), colors.HexColor('#2a0a0f'))  # red tint for cyl3
ts.add('BACKGROUND', (4,4), (7,4), colors.HexColor('#2a1f0a'))  # orange tint for cyl11
ct.setStyle(ts)
story.append(ct)
story.append(sp(2))

story.append(Paragraph(
    '🔴 Цилиндр 3 (2L): первые 3 секунды датчик показывал 600 °C (верхний предел), '
    'затем рухнул до −22 °C — классический обрыв термопары EGT. '
    'Реальная температура выхлопа и состояние форсунки Ц3 недоступны.', S_ALERT))
story.append(sp(1))
story.append(Paragraph(
    '⚠ Цилиндр 6 (3R): средняя EGT 530 °C — на +29 °C выше нормы по парку. '
    'Форсунка прошла тест, вероятна проблема с клапаном или калибровкой.', S_WARN))
story.append(sp(3))

# ── SECTION 4: INJECTORS ─────────────────────────────────────────────────
story.append(Paragraph('4. Состояние форсунок (по тест-отчётам)', S_H2))
story.append(sp(1))
story.append(Paragraph(
    'Тест-отчёты для всех 16 форсунок имеются в репозитории. '
    'Результаты кодируются как (+) — годен, (−) — отказ по данному режиму.',
    S_BODY))
story.append(sp(1))

inj_data = [
    (1,'1L','(−)','ОТКАЗ', C_RED),
    (2,'1R','(+)','Годен', C_GREEN),
    (3,'2L','(+−+−)','Частичный', C_ORANGE),
    (4,'2R','(−−++)','Частичный', C_ORANGE),
    (5,'3L','(+−+−)','Частичный', C_ORANGE),
    (6,'3R','(+)','Годен', C_GREEN),
    (7,'4L','(+)','Годен', C_GREEN),
    (8,'4R','(−)','ОТКАЗ', C_RED),
    (9,'5L','(+−−−)','Преим. отказ', C_RED),
    (10,'5R','(−++−)','Частичный', C_ORANGE),
    (11,'6L','(−++−)','Частичный', C_ORANGE),
    (12,'6R','(−+++)','Частичный', C_ORANGE),
    (13,'7L','(−)','ОТКАЗ', C_RED),
    (14,'7R','(−+−−)','Преим. отказ', C_RED),
    (15,'8L','(−+−+)','Частичный', C_ORANGE),
    (16,'8R','(−+−+)','Частичный', C_ORANGE),
]

left_inj  = inj_data[:8]
right_inj = inj_data[8:]

def inj_row(n, pos, code, res, col):
    return [
        Paragraph(f'<font name="Sans-Bold" color="#e6edf3">Ц{n} ({pos})</font>', S_CAPTION),
        Paragraph(f'<font name="Sans" color="#8b949e">{code}</font>', S_CAPTION),
        Paragraph(f'<font name="Sans-Bold" color="{col.hexval()}">{res}</font>', S_CAPTION),
    ]

inj_rows = []
for i in range(8):
    l = left_inj[i]; r = right_inj[i]
    inj_rows.append(
        inj_row(*l) + inj_row(*r)
    )

icols = [PW*0.115, PW*0.095, PW*0.12] * 2
story.append(hdr_table(
    ['Форсунка','Тест','Итог','Форсунка','Тест','Итог'], icols))
it = Table(inj_rows, colWidths=icols)
its = TableStyle([
    ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
    ('LEFTPADDING',   (0,0),(-1,-1), 5),
    ('RIGHTPADDING',  (0,0),(-1,-1), 5),
    ('TOPPADDING',    (0,0),(-1,-1), 4),
    ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ('LINEBEFORE', (3,0), (3,-1), 1.0, C_BORDER),
])
for i in range(8):
    bg = C_PANEL if i%2==0 else C_ROW_ALT
    its.add('BACKGROUND', (0,i), (-1,i), bg)
# red rows for full fails
for i, (n,pos,code,res,col) in enumerate(left_inj):
    if res in ('ОТКАЗ','Преим. отказ'):
        its.add('BACKGROUND', (0,i), (2,i), colors.HexColor('#2a0a0f'))
for i, (n,pos,code,res,col) in enumerate(right_inj):
    if res in ('ОТКАЗ','Преим. отказ'):
        its.add('BACKGROUND', (3,i), (5,i), colors.HexColor('#2a0a0f'))
it.setStyle(its)
story.append(it)
story.append(sp(2))

# Summary pills
summary_txt = (
    '<font name="Sans-Bold" color="#ff4060">Полный отказ: Ц1, Ц8, Ц13</font>  '
    '  <font name="Sans-Bold" color="#FBAE40">Преим. отказ: Ц9, Ц14</font>  '
    '  <font name="Sans-Bold" color="#FBAE40">Частичный: Ц3, Ц4, Ц5, Ц10, Ц11, Ц12, Ц15, Ц16</font>  '
    '  <font name="Sans-Bold" color="#3EF0AF">Годен: Ц2, Ц6, Ц7</font>'
)
story.append(Paragraph(summary_txt, S_BODY))
story.append(sp(1))
story.append(Paragraph(
    'ЭБУ компенсирует деградацию форсунок корректировкой подачи — поэтому EGT по '
    'цилиндрам остаётся относительно ровным (±12 °C). '
    'Однако при 11 000 ч ресурс компенсации исчерпывается. '
    'ТНВД пока справляется: давление рейки 1 522 / 1 523 бар, '
    'рабочий цикл насоса 15.7% (норма).', S_BODY))
story.append(sp(3))

# ── SECTION 5: MAINTENANCE ────────────────────────────────────────────────
story.append(Paragraph('5. История обслуживания (март – июнь 2026)', S_H2))
story.append(sp(1))

maint = [
    ('17.12.2025','15 103 ч','Замена 2 выпускных клапана (цил. 2R/Ц4)'),
    ('06.03.2026','16 939 ч','Отбор проб масла, долив масла ДВС 30 л'),
    ('08.03.2026','16 996 ч','ТО-3'),
    ('16.03.2026','17 184 ч','Долив масла ДВС 20 л'),
    ('19.03.2026','17 255 ч','ТО-1'),
    ('30.03.2026','17 503 ч','ТО-2'),
    ('08.04.2026','17 688 ч','Долив масла ДВС 20 л, РМК правый 10 л'),
    ('10.04.2026','17 746 ч','ТО-1, замена фланца карданного вала'),
    ('11.04.2026','17 767 ч','Диагностика РМК лев. → простой 12 сут. 7 ч'),
    ('27.04.2026','17 850 ч','Замена датчика давления маслоподкачки, выравнивание ПГП'),
    ('29.04.2026','17 908 ч','Долив масла ДВС 20 л'),
    ('03.05.2026','17 993 ч','ТО-4'),
    ('11.05.2026','18 180 ч','Долив масла ДВС 30 л'),
    ('13.05.2026','18 229 ч','Ошибка перегрев гидравлики, ремонт фишки электромагнита'),
    ('14.05.2026','18 247 ч','ТО-1'),
    ('20.05.2026','~18 300 ч','Долив масла ДВС 20 л'),
    ('06.06.2026','~18 400 ч','Снятие данных INSITE под нагрузкой (рейс с породой)'),
]

mcols = [PW*0.16, PW*0.16, PW*0.68]
story.append(hdr_table(['Дата','М/часы','Событие'], mcols))
m_rows = []
for date, mh, work in maint:
    is_red = 'простой' in work or 'замена' in work.lower()
    is_oil = 'Долив' in work
    wc = C_RED if is_red else (C_ORANGE if is_oil else C_TEXT)
    m_rows.append([
        Paragraph(date, S_CAPTION),
        Paragraph(mh, S_CAPTION),
        Paragraph(f'<font color="{wc.hexval()}">{work}</font>', S_CAPTION),
    ])
mt = Table(m_rows, colWidths=mcols)
mts = TableStyle([
    ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
    ('LEFTPADDING',   (0,0),(-1,-1), 5),
    ('RIGHTPADDING',  (0,0),(-1,-1), 5),
    ('TOPPADDING',    (0,0),(-1,-1), 3),
    ('BOTTOMPADDING', (0,0),(-1,-1), 3),
])
for i in range(len(maint)):
    bg = C_PANEL if i%2==0 else C_ROW_ALT
    mts.add('BACKGROUND', (0,i), (-1,i), bg)
# highlight big downtime
mts.add('BACKGROUND', (0,8), (-1,8), colors.HexColor('#2a0a0f'))
# highlight last row
mts.add('BACKGROUND', (0,16), (-1,16), colors.HexColor('#0a1f0a'))
mt.setStyle(mts)
story.append(mt)
story.append(sp(1))
story.append(Paragraph(
    'Расход масла: за март–май добавлено не менее 140 литров (≈ 70 л/мес, ~7 л/100 м/ч) — '
    'критически высокое потребление. Указывает на угар через поршневые кольца или '
    'уплотнения турбокомпрессора.', S_WARN))
story.append(sp(1))
story.append(Paragraph(
    'Расхождение наработки: ЭБУ фиксирует 11 016 ч (двигатель/ECM), '
    'журнал ТО показывает 17 000–18 000 машиночасов. Разница ≈ 6 000–7 000 ч — '
    'вероятна замена двигателя или ЭБУ. Необходимо уточнить историю агрегатов.', S_CAPTION))
story.append(sp(3))

# ── SECTION 6: ACTIONS ────────────────────────────────────────────────────
story.append(Paragraph('6. Рекомендуемые мероприятия', S_H2))
story.append(sp(1))

actions = [
    ('🔴 СРОЧНО','Замена термопары EGT цил. 3',
     'Датчик в обрыве — цилиндр без температурного контроля. '
     'Риск: невозможно диагностировать отказ форсунки или клапана Ц3.'),
    ('🔴 СРОЧНО','Замена всех 16 форсунок\n(приоритет: Ц1, Ц8, Ц9, Ц13, Ц14)',
     '5 форсунок в отказе или почти в отказе; остальные с нарушениями. '
     'При 11 000 ч плановый ресурс MCRS-форсунок исчерпан.'),
    ('🟡 БЛИЖАЙШЕЕ','Проверка масляного охладителя',
     'Масло 109 °C и растёт. Проверить маслоохладитель на засорение, '
     'термостатический клапан масла, уровень и качество охлаждающей жидкости.'),
    ('🟡 БЛИЖАЙШЕЕ','Диагностика угара масла',
     '140 л за 2 месяца. Компрессионный тест, анализ картерных газов, '
     'осмотр турбокомпрессора, проверка уплотнений масляных форсунок поршней.'),
    ('🔵 МОНИТОРИНГ','Цилиндр 6 — повышенная EGT',
     'EGT +29 °C выше нормы при "годной" форсунке. '
     'После замены всего комплекта форсунок проверить повторно. '
     'При сохранении — осмотр клапана и ГБЦ Ц6.'),
    ('🔵 УТОЧНИТЬ','Расхождение наработки ЭБУ / машины',
     'Поднять историю замены агрегатов, '
     'уточнить реальную наработку двигателя для планирования КР.'),
]

acols = [PW*0.14, PW*0.28, PW*0.58]
story.append(hdr_table(['Приоритет','Мероприятие','Обоснование'], acols))
arows = []
pcolmap = {'🔴': C_RED, '🟡': C_ORANGE, '🔵': C_BLUE}
for prio, action, reason in actions:
    pc = pcolmap.get(prio[:2], C_TEXT)
    arows.append([
        Paragraph(f'<font name="Sans-Bold" color="{pc.hexval()}">{prio}</font>', S_CAPTION),
        Paragraph(f'<font name="Sans-Bold" color="#e6edf3">{action}</font>', S_CAPTION),
        Paragraph(reason, S_CAPTION),
    ])
at = Table(arows, colWidths=acols)
ats = TableStyle([
    ('GRID', (0,0), (-1,-1), 0.3, C_BORDER),
    ('LEFTPADDING',   (0,0),(-1,-1), 5),
    ('RIGHTPADDING',  (0,0),(-1,-1), 5),
    ('TOPPADDING',    (0,0),(-1,-1), 5),
    ('BOTTOMPADDING', (0,0),(-1,-1), 5),
    ('VALIGN', (0,0), (-1,-1), 'TOP'),
])
for i in range(len(actions)):
    bg = C_PANEL if i%2==0 else C_ROW_ALT
    ats.add('BACKGROUND', (0,i), (-1,i), bg)
ats.add('BACKGROUND', (0,0), (-1,1), colors.HexColor('#2a0a0f'))
at.setStyle(ats)
story.append(at)
story.append(sp(3))
story.append(hr())

# ── FOOTER ────────────────────────────────────────────────────────────────
story.append(Paragraph(
    'Отчёт сформирован автоматически на основе данных INSITE Professional (DML-20260606-153327) '
    'и внутренней документации. Данные: NTE200 №46, Cummins QSK50 MCRS V16, '
    'серийный номер ДВС 33230485.',
    S_CAPTION))

# ── BUILD ──────────────────────────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_DARK)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # top bar
    canvas.setFillColor(C_PANEL)
    canvas.rect(0, H-8*mm, W, 8*mm, fill=1, stroke=0)
    canvas.setFont('Sans-Bold', 8)
    canvas.setFillColor(C_ACCENT)
    canvas.drawString(15*mm, H-5.5*mm, 'NTE200 №46  •  Анализ рейса 06.06.2026')
    canvas.setFont('Sans', 8)
    canvas.setFillColor(C_SUBTEXT)
    canvas.drawRightString(W-15*mm, H-5.5*mm, f'Стр. {doc.page}')
    # bottom bar
    canvas.setFillColor(C_PANEL)
    canvas.rect(0, 0, W, 7*mm, fill=1, stroke=0)
    canvas.setFont('Sans', 7)
    canvas.setFillColor(C_SUBTEXT)
    canvas.drawString(15*mm, 2.5*mm,
        'Конфиденциально  •  Cummins QSK50 MCRS  •  Горная Евразия')
    canvas.restoreState()

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'PDF: {OUT}')
