#!/usr/bin/env python3
"""NTE200 #46 comprehensive engine diagnostic report — v2, updated with fleet history."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Sans',      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('Sans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

C = dict(
    dark    = colors.HexColor('#0d1117'),
    panel   = colors.HexColor('#161b22'),
    border  = colors.HexColor('#30363d'),
    accent  = colors.HexColor('#3EF0AF'),
    red     = colors.HexColor('#ff4060'),
    orange  = colors.HexColor('#FBAE40'),
    green   = colors.HexColor('#3EF0AF'),
    blue    = colors.HexColor('#38bdf8'),
    text    = colors.HexColor('#e6edf3'),
    sub     = colors.HexColor('#8b949e'),
    hdr     = colors.HexColor('#1f2937'),
    alt     = colors.HexColor('#1a2332'),
    redbg   = colors.HexColor('#2a0a0f'),
    orangebg= colors.HexColor('#2a1f0a'),
    greenbg = colors.HexColor('#0a1f0a'),
)

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

def HR(): return HRFlowable(width='100%', thickness=0.5, color=C['border'], spaceAfter=5, spaceBefore=5)
def SP(h=3): return Spacer(1, h*mm)

W, H = A4

def table(rows, cols, style_extra=None, alt=True):
    t = Table(rows, colWidths=cols)
    ts = TableStyle([
        ('GRID',         (0,0),(-1,-1), 0.3, C['border']),
        ('LEFTPADDING',  (0,0),(-1,-1), 5),
        ('RIGHTPADDING', (0,0),(-1,-1), 5),
        ('TOPPADDING',   (0,0),(-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('VALIGN',       (0,0),(-1,-1), 'TOP'),
    ])
    for i in range(len(rows)):
        bg = C['panel'] if i%2==0 else (C['alt'] if alt else C['panel'])
        ts.add('BACKGROUND', (0,i), (-1,i), bg)
    if style_extra:
        for cmd in style_extra:
            ts.add(*cmd)
    t.setStyle(ts)
    return t

def hdr_row(cells, widths):
    row = [Pc(C['sub'], c, bold=True, size=8, lead=11) for c in cells]
    t = Table([row], colWidths=widths)
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0),(-1,-1), C['hdr']),
        ('GRID',         (0,0),(-1,-1), 0.3, C['border']),
        ('LEFTPADDING',  (0,0),(-1,-1), 5),
        ('RIGHTPADDING', (0,0),(-1,-1), 5),
        ('TOPPADDING',   (0,0),(-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
    ]))
    return t

def status(icon, txt, col):
    return Pc(col, f'{icon} {txt}', bold=True, size=8)

# ── PAGE BACKGROUND / HEADERS ─────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C['dark']); canvas.rect(0,0,W,H,fill=1,stroke=0)
    canvas.setFillColor(C['panel']); canvas.rect(0,H-9*mm,W,9*mm,fill=1,stroke=0)
    canvas.setFont('Sans-Bold',8); canvas.setFillColor(C['accent'])
    canvas.drawString(15*mm, H-6*mm, 'NTE200 №46  •  Диагностический отчёт ДВС  •  06.06.2026')
    canvas.setFont('Sans',8); canvas.setFillColor(C['sub'])
    canvas.drawRightString(W-15*mm, H-6*mm, f'Стр. {doc.page}')
    canvas.setFillColor(C['panel']); canvas.rect(0,0,W,7*mm,fill=1,stroke=0)
    canvas.setFont('Sans',7); canvas.setFillColor(C['sub'])
    canvas.drawString(15*mm,2.5*mm,'Конфиденциально  •  QSK50 MCRS V16 s/n 33230485  •  Горная Евразия')
    canvas.restoreState()

OUT = '/home/user/NTE200/NTE200_46_report_v2_06062026.pdf'
doc = SimpleDocTemplate(OUT, pagesize=A4,
    leftMargin=15*mm, rightMargin=15*mm, topMargin=18*mm, bottomMargin=12*mm,
    title='NTE200 №46 — Диагностика ДВС 06.06.2026')
PW = doc.width

story = []

# ═══════════════════════════════════════════════════════════
# ОБЛОЖКА
# ═══════════════════════════════════════════════════════════
story += [
    Pc(C['sub'], 'ТЕХНИЧЕСКИЙ ДИАГНОСТИЧЕСКИЙ ОТЧЁТ', size=10),
    SP(1),
    Pc(C['accent'], 'NTE200 №46 — Анализ состояния двигателя', bold=True, size=19, lead=25),
    SP(2),
    Pc(C['sub'], '06 июня 2026 г.  •  Рейс с породой  •  Наработка по журналу: ~18 743 м/ч', size=10),
    SP(1),
    Pc(C['sub'], 'Cummins QSK50 MCRS V16  •  S/N 33230485  •  ECM: 11 016 моточасов', size=9),
    SP(1),
    Pc(C['orange'], '⚠  Диагностика CAMSS проведена 04.06.2026 — за 2 дня до захвата данных', bold=True, size=9),
    HR(), SP(2),
]

# ═══════════════════════════════════════════════════════════
# 1. KPI КАРТОЧКИ
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '1. Ключевые параметры рейса', bold=True, size=13))
story.append(SP(1))

kpi_h = ['Нагрузка ДВС','Обороты','Расход топлива','T охл.жидк.','T масла (макс)','P масла']
kpi_v = ['94 %','1 876 об/мин','339 л/ч','85.7 °C','110.0 °C','516 кПа']
kpi_c = [C['green'],C['green'],C['blue'],C['green'],C['orange'],C['green']]

hrow = [Pc(C['sub'],h,size=8,lead=11,align=TA_CENTER) for h in kpi_h]
vrow = [Pc(c,v,bold=True,size=13,lead=17,align=TA_CENTER) for v,c in zip(kpi_v,kpi_c)]
cw6 = [PW/6]*6
kt = Table([hrow,vrow], colWidths=cw6)
kts = TableStyle([
    ('BACKGROUND',(0,0),(-1,-1),C['panel']),
    ('GRID',(0,0),(-1,-1),0.5,C['border']),
    ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
    ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),
    ('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ('BACKGROUND',(4,1),(4,1),C['orangebg']),
])
kt.setStyle(kts); story += [kt, SP(3)]

# ═══════════════════════════════════════════════════════════
# 2. ПАРАМЕТРЫ ДВС
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '2. Параметры двигателя', bold=True, size=13))
story.append(SP(1))

pw = [PW*0.30, PW*0.25, PW*0.25, PW*0.20]
story.append(hdr_row(['Параметр','Значение','Диапазон','Статус'], pw))

params = [
    ('Температура охлаждающей жидкости','84.1 °C avg / 85.7 °C макс','Норма 82–95 °C', status('✓','НОРМА',C['green'])),
    ('Температура масла ДВС','104.0 °C avg / 110.0 °C макс','⚠ Порог предупр. 105 °C', status('!','ВНИМАНИЕ',C['orange'])),
    ('Давление масла','516 кПа avg / 384 кПа мин','Норма 350–650 кПа', status('✓','НОРМА',C['green'])),
    ('Давление картерных газов','0.69 кПа avg / 1.10 кПа макс','Норма < 3 кПа', status('✓','НОРМА',C['green'])),
    ('Давление наддува','268 кПа avg / 295 кПа макс','Норма 220–320 кПа', status('✓','НОРМА',C['green'])),
    ('Давление топл. рейки (заданное/факт)','1 523 / 1 522 бар  •  Δ avg 12 бар','Расхожд. <30 бар норма', status('✓','НОРМА',C['green'])),
    ('Рабочий цикл ТНВД','15.7 % avg / 24 % макс','Норма < 50 %', status('✓','НОРМА',C['green'])),
    ('Напряжение АКБ','27.8 В avg','Норма 24–30 В', status('✓','НОРМА',C['green'])),
    ('Жёлтая/Красная лампа (AWL/RSL)','Выкл. всё время рейса','Активных кодов нет', status('✓','НОРМА',C['green'])),
]

prows = [[Pc(C['text'],a,size=8,lead=12), Pc(C['text'],b,size=8,lead=12),
          Pc(C['sub'],c,size=8,lead=12), d] for a,b,c,d in params]

story.append(table(prows, pw, style_extra=[
    ('BACKGROUND',(0,1),(-1,1),C['orangebg']),
]))
story += [SP(1),
    Pc(C['orange'],
       '⚠ Масло 105–110 °C в течение 643 с из 1 087 (59% рейса). '
       'Динамика нарастающая: старт 87.8 °C → финиш 109.3 °C и продолжает расти. '
       'При длинных рейсах или жаркой погоде неизбежен выход в красную зону (>115 °C).',
       bold=True, size=8),
    SP(3),
]

# ═══════════════════════════════════════════════════════════
# 3. ТЕМПЕРАТУРЫ ВЫХЛОПА ЦИЛИНДРОВ
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '3. Температуры выхлопа по цилиндрам', bold=True, size=13))
story.append(SP(1))
story.append(Pc(C['sub'],
    'Среднее по 15 исправным цилиндрам: 501.6 °C  •  Разброс: ±11.8 °C  (отличный результат — ЭБУ корректирует подачу)',
    size=8))
story.append(SP(1))

cyls = [
    (1,'L','499.6','+13.8','✓',C['green']),
    (2,'R','513.7','+12.1','✓',C['green']),
    (3,'L','250.0 *','—','✗ ДАТЧИК',C['red']),
    (4,'R','495.8','+9.9','✓',C['green']),
    (5,'L','493.3','+7.4','✓',C['green']),
    (6,'R','530.4','+28.8','! Повыш.',C['orange']),
    (7,'L','492.1','+6.3','✓',C['green']),
    (8,'R','488.9','+3.1','✓',C['green']),
    (9,'L','497.8','+12.0','✓',C['green']),
    (10,'R','497.8','+12.0','✓',C['green']),
    (11,'L','513.5','+12.0','! Слег.',C['orange']),
    (12,'R','501.7','+15.9','✓',C['green']),
    (13,'L','502.3','+16.5','✓',C['green']),
    (14,'R','479.4','-6.4','✓',C['green']),
    (15,'L','508.5','+22.6','✓',C['green']),
    (16,'R','508.8','+22.9','✓',C['green']),
]

def cyl_row(n, bank, avg, dev, st, sc):
    return [
        Pc(C['text'], f'Ц{n} ({bank}-банк)', bold=True, size=8, lead=11),
        Pc(sc, avg+' °C', bold=(sc!=C['green']), size=8, lead=11),
        Pc(C['sub'], dev, size=8, lead=11),
        Pc(sc, st, bold=True, size=8, lead=11),
    ]

left_rows  = [cyl_row(*c) for c in cyls[:8]]
right_rows = [cyl_row(*c) for c in cyls[8:]]
combined   = [l+r for l,r in zip(left_rows, right_rows)]

cw8 = [PW*0.12, PW*0.13, PW*0.07, PW*0.08]*2
story.append(hdr_row(
    ['Цилиндр','EGT сред.','Откл.','Статус']*2, cw8))

ct = Table(combined, colWidths=cw8)
cts = TableStyle([
    ('GRID',(0,0),(-1,-1),0.3,C['border']),
    ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
    ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
    ('LINEBEFORE',(4,0),(4,-1),1.0,C['border']),
])
for i in range(8):
    cts.add('BACKGROUND',(0,i),(-1,i),C['panel'] if i%2==0 else C['alt'])
cts.add('BACKGROUND',(0,2),(3,2),C['redbg'])   # Цил3 left half
cts.add('BACKGROUND',(4,4),(7,4),C['orangebg'])  # Цил11
ct.setStyle(cts); story.append(ct)

story += [SP(1),
    Pc(C['red'],
       '🔴 Цилиндр 3 (2L): датчик EGT в обрыве — первые 3 с показывал 600 °C (предел), затем рухнул до −22 °C. '
       'Цилиндр без температурного контроля. Примечательно: по истории Dec-2025 именно Ц2R и Ц5L '
       'ремонтировались из-за заниженной EGT (обрыв/износ клапанов). '
       'Неисправность датчика Ц3 требует немедленной замены для контроля состояния.',
       bold=True, size=8),
    SP(1),
    Pc(C['orange'],
       '⚠ Цилиндр 6 (3R): EGT +29 °C выше нормы несмотря на "годный" результат теста форсунки. '
       'Вероятна проблема клапана или незначительная утечка охлаждения.',
       bold=True, size=8),
    SP(3),
]

# ═══════════════════════════════════════════════════════════
# 4. ФОРСУНКИ
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '4. Состояние форсунок (тест-отчёты)', bold=True, size=13))
story.append(SP(1))
story.append(Pc(C['sub'],
    'Тест проведён на стенде до установки. Коды: (+) — годен по данному режиму, (−) — превышение норм утечки/подачи.',
    size=8))
story.append(SP(1))

inj = [
    (1,'1L','(−)','ОТКАЗ',C['red'],C['redbg']),
    (2,'1R','(+)','Годен',C['green'],C['panel']),
    (3,'2L','(+−+−)','Частичный',C['orange'],C['orangebg']),
    (4,'2R','(−−++)','Частичный',C['orange'],C['orangebg']),
    (5,'3L','(+−+−)','Частичный',C['orange'],C['orangebg']),
    (6,'3R','(+)','Годен',C['green'],C['panel']),
    (7,'4L','(+)','Годен',C['green'],C['panel']),
    (8,'4R','(−)','ОТКАЗ',C['red'],C['redbg']),
    (9,'5L','(+−−−)','Преим. отказ',C['red'],C['redbg']),
    (10,'5R','(−++−)','Частичный',C['orange'],C['orangebg']),
    (11,'6L','(−++−)','Частичный',C['orange'],C['orangebg']),
    (12,'6R','(−+++)','Частичный',C['orange'],C['orangebg']),
    (13,'7L','(−)','ОТКАЗ',C['red'],C['redbg']),
    (14,'7R','(−+−−)','Преим. отказ',C['red'],C['redbg']),
    (15,'8L','(−+−+)','Частичный',C['orange'],C['orangebg']),
    (16,'8R','(−+−+)','Частичный',C['orange'],C['orangebg']),
]

def inj_cell(n, pos, code, res, col):
    return [
        Pc(C['text'],f'Ц{n} ({pos})',bold=True,size=8,lead=11),
        Pc(C['sub'],code,size=8,lead=11),
        Pc(col,res,bold=True,size=8,lead=11),
    ]

li = inj[:8]; ri = inj[8:]
irows = [inj_cell(*l[:5]) + inj_cell(*r[:5]) for l,r in zip(li,ri)]
icols = [PW*0.115,PW*0.095,PW*0.115]*2
story.append(hdr_row(['Форсунка','Тест','Итог']*2, icols))

it = Table(irows, colWidths=icols)
its = TableStyle([
    ('GRID',(0,0),(-1,-1),0.3,C['border']),
    ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
    ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
    ('LINEBEFORE',(3,0),(3,-1),1.0,C['border']),
])
for i,(l,r) in enumerate(zip(li,ri)):
    bg = C['panel'] if i%2==0 else C['alt']
    its.add('BACKGROUND',(0,i),(2,i), l[5] if l[5]!=C['panel'] else bg)
    its.add('BACKGROUND',(3,i),(5,i), r[5] if r[5]!=C['panel'] else bg)
it.setStyle(its); story.append(it)

story += [SP(1),
    Pc(C['red'],
       'Итог: 3 форсунки в полном отказе (Ц1, Ц8, Ц13), 2 преимущественно отказали (Ц9, Ц14), '
       '9 частично нарушены, лишь 3 прошли тест без замечаний (Ц2, Ц6, Ц7). '
       'ЭБУ удерживает баланс EGT через коррекцию подачи — это скрывает деградацию, '
       'но не предотвращает рост давления рейки и нагрузки на ТНВД.',
       bold=True, size=8),
    SP(3),
]

# ═══════════════════════════════════════════════════════════
# 5. АНАЛИЗ МАСЛА И ПОТРЕБЛЕНИЕ
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '5. Расход моторного масла', bold=True, size=13))
story.append(SP(1))

oil_events = [
    ('06.03.2026','16 939','30','—'),
    ('16.03.2026','17 184','20','—'),
    ('24.03.2026','17 364','20','—'),
    ('08.04.2026','17 688','20','—'),
    ('29.04.2026','17 908','20','—'),
    ('11.05.2026','18 180','30','—'),
    ('20.05.2026','18 379','20','—'),
    ('30.05.2026','18 627','30','—'),
    ('02.06.2026','18 700','15','—'),
]
oil_rows = [[Pc(C['text'],d,size=8,lead=11), Pc(C['text'],m+' ч',size=8,lead=11),
             Pc(C['orange'],v+' л',bold=True,size=8,lead=11), Pc(C['sub'],n,size=8,lead=11)]
            for d,m,v,n in oil_events]

oil_w = [PW*0.22, PW*0.22, PW*0.22, PW*0.34]
story.append(hdr_row(['Дата','Мото-часы','Объём доливки','Примечание'], oil_w))
story.append(table(oil_rows, oil_w))

story += [SP(1),
    Pc(C['red'],
       '🔴 КРИТИЧНО: за 1 761 моточас (март–июнь 2026) добавлено 205 литров моторного масла. '
       'Удельный расход: 11.6 л/100 ч — в 2.3 раза выше допустимого (норма QSK50: ≤ 5 л/100 ч). '
       'Масло уходит через: выпускные клапанные сёдла (история сизого дыма Dec-2025), '
       'уплотнения турбокомпрессора (визуальная проверка не проводилась), '
       'возможно — изношенные поршневые кольца.',
       bold=True, size=8),
    SP(2),
    Pc(C['orange'],
       'История декабря 2025: при 15 103 м/ч зафиксирован повышенный СИЗЫЙ ДЫМ (масло). '
       'Тогда же: заниженная EGT на 2R и 5L, замена клапанов/коромысел. '
       'Нынешняя картина (11.6 л/100ч) означает, что угар масла продолжается и нарастает.',
       bold=True, size=8),
    SP(3),
]

# ═══════════════════════════════════════════════════════════
# 6. ИСТОРИЯ ОТКАЗОВ ДВС (полная)
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '6. История неисправностей ДВС (из базы Сводного анализа)', bold=True, size=13))
story.append(SP(1))
story.append(Pc(C['sub'],
    'Данные из сводной базы отказов по парку NTE200. Сортировка по хронологии.',
    size=8))
story.append(SP(1))

fault_hist = [
    ('10.04.2025','13 497','ДВС — ГРМ',
     'ТО: посторонние шумы ДВС. Замена осей коромысел впускных и выпускных, крейцкопфа, прокладок клапанных крышек',
     C['orange'], C['orangebg']),
    ('02.08.2025','12 039','ДВС',
     'Регулировка клапанов',
     C['sub'], C['panel']),
    ('12.12.2025','15 030','Топливная система',
     'Замена 2 форсунок (цилиндры не уточнены)',
     C['orange'], C['orangebg']),
    ('15.12.2025','15 103','ДВС — ГБЦ',
     'Повышенная дымность (СИЗЫЙ ДЫМ). Заниженная EGT на 2R и шум с 5L. '
     'Замена 2R: 2 выпускных клапана + 2 седла + коромысло + ось + штифты + траверса. '
     'Замена 5L: коромысло + ось + штифты + траверса',
     C['red'], C['redbg']),
    ('11.04.2026','17 767','Гидравлика — РМК',
     'Диагностика РМК левый (течь масла, превышение браковочных показателей). Замена РМК. Простой 12 суток 7 часов.',
     C['orange'], C['orangebg']),
    ('04.06.2026','~18 743','ДВС — диагностика',
     'Диагностика ДВС в системе CAMSS. Результаты послужили основой для захвата данных 06.06.2026',
     C['blue'], C['panel']),
]

fw = [PW*0.14, PW*0.12, PW*0.18, PW*0.56]
story.append(hdr_row(['Дата','М/часы','Система','Описание событий'], fw))
frows = []
for date, mh, sys_, desc, col, bg in fault_hist:
    frows.append([
        Pc(col, date, bold=True, size=8, lead=11),
        Pc(C['sub'], mh+' ч', size=8, lead=11),
        Pc(col, sys_, bold=True, size=8, lead=11),
        Pc(C['text'], desc, size=8, lead=11),
    ])
ft = Table(frows, colWidths=fw)
fts = TableStyle([
    ('GRID',(0,0),(-1,-1),0.3,C['border']),
    ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
    ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
    ('VALIGN',(0,0),(-1,-1),'TOP'),
])
for i,(_, _, _, _, _, bg) in enumerate(fault_hist):
    fts.add('BACKGROUND',(0,i),(-1,i),bg)
ft.setStyle(fts); story.append(ft)
story.append(SP(1))
story.append(Pc(C['orange'],
    '⚠ Расхождение наработки: Сводный анализ фиксирует наработку 12 039 ч (авг.2025) и 15 103 ч (дек.2025), '
    'тогда как ECM в рейсе 06.06.2026 показывает 11 016 ч. '
    'ЭБУ или двигатель был заменён в период дек.2025 — март.2026 (счётчик был обнулён). '
    'Реальная суммарная наработка двигателя: ~18 743 ч (по журналу ТО).',
    bold=True, size=8))
story.append(SP(3))

# ═══════════════════════════════════════════════════════════
# 7. КОНТЕКСТ ФЛОТА — схожие отказы у других NTE200
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '7. Паттерн флота — схожие отказы на других NTE200', bold=True, size=13))
story.append(SP(1))
story.append(Pc(C['sub'],
    'Анализ Сводного отчёта показывает системный характер: форсунки и клапана — главные виды отказов по всему парку.',
    size=8))
story.append(SP(1))

fleet = [
    ('№43','2 075 ч','2026-01','Замена 4 форсунок + выпускные клапаны'),
    ('№43','11 943 ч','2025-07','K/Na превышение → замена ГБЦ + охлаждение'),
    ('№45','17 866 ч','2026-03','Сизый дым, замена клапанов 2R'),
    ('№47','15 522 ч','2025-12','Замена 4 форсунок (2L,3L,4L,5R) + клапана'),
    ('№47','15 663 ч','2025-12','Замена форсунки 5R'),
    ('№48','15 147 ч','2025-12','ОЖ в масле → замена ГБЦ + гильза + поршень + клапана'),
    ('№53','16 770 ч','2026-03','Рост свинца в масле → диагностика, лобовая крышка'),
    ('№57','14 808 ч','2026-02','Диагностика всех цил., демонтаж 5 ГБЦ (2R,3R,2L,3L,1R)'),
    ('№60','16 096 ч','2026-04','Замена форсунки 1L (высокая T), диагностика 8,9'),
]
fw2 = [PW*0.08, PW*0.12, PW*0.12, PW*0.68]
story.append(hdr_row(['Машина','М/часы','Дата','Описание'], fw2))
fleet_rows = [[Pc(C['accent'],m,bold=True,size=8,lead=11), Pc(C['sub'],h+' ч',size=8,lead=11),
               Pc(C['sub'],d,size=8,lead=11), Pc(C['text'],desc,size=8,lead=11)]
              for m,h,d,desc in fleet]
story.append(table(fleet_rows, fw2))
story += [SP(1),
    Pc(C['orange'],
       'Флотовый паттерн: форсунки MCRS выходят из строя при 10 000–17 000 м/ч, '
       'что совпадает с ресурсом Cummins (рек. замена 10 000–12 000 ч). '
       'На №46 плановая замена форсунок не проведена вовремя — компенсация ЭБУ работает '
       'на пределе, что подтверждается высоким расходом топлива (339 л/ч при норме ~260–280 л/ч на аналогичной нагрузке).',
       bold=True, size=8),
    SP(3),
]

# ═══════════════════════════════════════════════════════════
# 8. МЕРОПРИЯТИЯ
# ═══════════════════════════════════════════════════════════
story.append(Pc(C['text'], '8. Рекомендуемые мероприятия', bold=True, size=13))
story.append(SP(1))

actions = [
    ('🔴','НЕМЕДЛЕННО','Замена термопары EGT цилиндра 3 (2L)',
     'Датчик в обрыве. Цилиндр 3 — вне контроля. После замены: верифицировать EGT Ц3 '
     'под нагрузкой (по данным форсунка частично в отказе — возможна аномалия).', C['red']),

    ('🔴','НЕМЕДЛЕННО','Замена всех 16 форсунок',
     'Приоритет: Ц1, Ц8, Ц13 (полный отказ), Ц9, Ц14 (преимущественный отказ). '
     'При 11 000 ч ECM (~18 700 ч по журналу) плановый ресурс QSK50 MCRS-форсунок полностью исчерпан. '
     'После замены: захват данных INSITE под нагрузкой для верификации, проверить расход топлива '
     '(ожидается снижение до 270–290 л/ч при эквивалентной нагрузке).', C['red']),

    ('🔴','НЕМЕДЛЕННО','Диагностика угара масла (11.6 л/100 ч)',
     'Провести: компрессионный тест всех 16 цилиндров, осмотр турбокомпрессора (синий дым на выхлопе при '
     'остановке), анализ картерных газов, осмотр свечей накаливания на нагар. '
     'Возможные причины: изношенные поршневые кольца (история сизого дыма с Dec-2025), '
     'повреждённые сёдла клапанов после ремонта Dec-2025, уплотнения ТРК. '
     'Лабораторный анализ масла: железо (Fe), свинец (Pb), медь (Cu), вязкость, ТЩЧ.', C['red']),

    ('🟡','БЛИЖАЙШЕЕ ТО','Проверка системы охлаждения масла',
     'Масло достигает 109°C за 18 минут рейса под полной нагрузкой. '
     'Проверить: маслоохладитель (промывка или замена), термостатический клапан масла, '
     'уровень и качество ОЖ, работу масляной помпы. '
     'Параллельно: заменить масло (текущее работает при 105–110°C — ускоренная деградация).', C['orange']),

    ('🟡','ПОСЛЕ ЗАМЕНЫ ФОРСУНОК','Контроль цилиндра 6 (3R)',
     'EGT Ц6 = 530°C (+29°C выше нормы) несмотря на "годную" форсунку. '
     'После установки нового комплекта форсунок: повторная запись INSITE под нагрузкой. '
     'Если Ц6 остаётся высоким — осмотр ГБЦ Ц6: клапан, седло, прокладка ГБЦ.', C['orange']),

    ('🔵','УТОЧНИТЬ','История замены ЭБУ/двигателя',
     'ECM фиксирует 11 016 м/ч, журнал ТО — 18 700+ м/ч. '
     'Разница ~7 700 ч указывает на замену ЭБУ или двигателя в период дек.2025 – март.2026. '
     'Для планирования следующего КР необходимо знать фактическую суммарную наработку.', C['blue']),
]

aw = [PW*0.06, PW*0.16, PW*0.32, PW*0.46]
story.append(hdr_row(['','Срок','Мероприятие','Обоснование и детали'], aw))
arows = []
aextra = []
for i,(ico,when,what,why,col) in enumerate(actions):
    arows.append([
        Pc(col, ico, bold=True, size=10, lead=13, align=TA_CENTER),
        Pc(col, when, bold=True, size=8, lead=11),
        Pc(C['text'], what, bold=True, size=8, lead=11),
        Pc(C['text'], why, size=8, lead=11),
    ])
    bg = C['redbg'] if col==C['red'] else (C['orangebg'] if col==C['orange'] else C['panel'])
    aextra.append(('BACKGROUND',(0,i),(-1,i),bg))

story.append(table(arows, aw, style_extra=aextra, alt=False))
story.append(SP(3))
story.append(HR())
story.append(SP(1))
story.append(Pc(C['sub'],
    'Отчёт подготовлен на основе: DML-20260606-153327 (INSITE Professional, 1 087 записей) · '
    'Тест-отчёты форсунок (16 PDF) · ОТЧЕТ Полюс Магадан (обновл. 05.06.2026, 31 запись по №46) · '
    'Доливки Горная Евразия · Сводный анализ неисправностей NTE200 · База ГБЦ ремонтов',
    size=7))

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
