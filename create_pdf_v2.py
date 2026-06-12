"""Генерация PDF-отчёта v2 — с таблицей ECM как первопричиной."""
import fitz
import io, os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, Image, PageBreak, HRFlowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

BASE = "/home/user/NTE200/"

pdfmetrics.registerFont(TTFont('DV',  '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DVB', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
OBL = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf'
pdfmetrics.registerFont(TTFont('DVI', OBL if os.path.exists(OBL) else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFontFamily('DV', normal='DV', bold='DVB', italic='DVI', boldItalic='DVB')

W, H = A4
C_DARK  = colors.HexColor('#1F497D')
C_MED   = colors.HexColor('#2E74B5')
C_LIGHT = colors.HexColor('#EBF3FB')
C_YELL  = colors.HexColor('#FFE699')
C_GREEN = colors.HexColor('#D9EAD3')
C_RED_L = colors.HexColor('#FFD7D7')
C_ORG   = colors.HexColor('#FCE5CD')
C_GREY  = colors.HexColor('#606060')
C_RED   = colors.HexColor('#C00000')
C_GRN   = colors.HexColor('#007000')
C_WARN  = colors.HexColor('#BF5F00')

def s(name='n', font='DV', sz=10, col=colors.black,
      al=TA_LEFT, bold=False, italic=False, lead=None, sb=2, sa=2):
    fn = 'DVB' if bold else ('DVI' if italic else font)
    return ParagraphStyle(name, fontName=fn, fontSize=sz, textColor=col,
                          alignment=al, leading=lead or sz*1.35,
                          spaceBefore=sb, spaceAfter=sa)

S_BIG   = s('big',  sz=20, col=C_DARK,  bold=True,  al=TA_CENTER, lead=26)
S_MED   = s('med',  sz=13, bold=True,   al=TA_CENTER)
S_SUB   = s('sub',  sz=10,              al=TA_CENTER)
S_H1    = s('h1',   sz=14, col=C_DARK,  bold=True,  sb=10, sa=4)
S_H2    = s('h2',   sz=12, col=C_MED,   bold=True,  sb=7,  sa=3)
S_H3    = s('h3',   sz=10, col=C_DARK,  bold=True,  sb=4,  sa=2)
S_BODY  = s('body', sz=10, sb=2, sa=2)
S_BUL   = s('bul',  sz=9,  sb=1, sa=1)
S_CAP   = s('cap',  sz=8,  col=C_GREY, al=TA_CENTER, italic=True)
S_SMALL = s('sml',  sz=8,  col=C_GREY, italic=True)
S_BOLD  = s('bld',  sz=10, bold=True)
S_RED   = s('red',  sz=10, col=C_RED,  bold=True)
S_WARN  = s('wrn',  sz=11, col=C_RED,  bold=True)

def p(text, style=None):
    return Paragraph(text, style or S_BODY)

def hdr_style(bg):
    return [
        ('BACKGROUND',   (0,0),(-1,0),  bg),
        ('TEXTCOLOR',    (0,0),(-1,0),  colors.white),
        ('FONTNAME',     (0,0),(-1,0),  'DVB'),
        ('FONTSIZE',     (0,0),(-1,-1), 9),
        ('FONTNAME',     (0,1),(-1,-1), 'DV'),
        ('GRID',         (0,0),(-1,-1), 0.4, colors.grey),
        ('VALIGN',       (0,0),(-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, C_LIGHT]),
        ('LEFTPADDING',  (0,0),(-1,-1), 4),
        ('RIGHTPADDING', (0,0),(-1,-1), 4),
        ('TOPPADDING',   (0,0),(-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
    ]

def c(text, bold=False, sz=9, col=colors.black, al=TA_LEFT):
    return Paragraph(text, s('c', sz=sz, bold=bold, col=col, al=al))

def box(text, label='', bg=C_YELL, w=17*cm):
    t = Table([[p(f'<b>{label}</b> {text}', s('bx', sz=9))]], colWidths=[w])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), bg),
        ('BOX',        (0,0),(-1,-1), 1, colors.HexColor('#AAAAAA')),
        ('LEFTPADDING',(0,0),(-1,-1), 6), ('RIGHTPADDING',(0,0),(-1,-1),6),
        ('TOPPADDING', (0,0),(-1,-1), 5), ('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    return t

def pdf_img(pdf_path, pg, w_cm=16, dpi=100):
    doc = fitz.open(pdf_path)
    page = doc[pg]
    mat = fitz.Matrix(dpi/72, dpi/72)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
    img = pix.tobytes("jpeg", jpg_quality=72)
    doc.close()
    wp = w_cm*cm
    asp = pix.height/pix.width
    hp = wp*asp
    mx = 18*cm
    if hp > mx: hp=mx; wp=hp/asp
    return Image(io.BytesIO(img), width=wp, height=hp)

def sp(n=6): return Spacer(1, n)
def hr(): return HRFlowable(width='100%', thickness=1, color=C_DARK)

# =========================================================
story = []

# ── ТИТУЛ ─────────────────────────────────────────────────
story += [sp(40),
    p('ТЕХНИЧЕСКИЙ ОТЧЁТ', S_BIG), sp(4),
    p('Анализ причин массового износа клапанного механизма', S_MED), sp(2),
    p('ДВС Cummins QSK50 / Автосамосвалы NTE200', S_SUB),
    p('АО «Полюс Магадан»', S_SUB), sp(20)]

info = [
    [c('Заказчик:', bold=True),          c('АО «Полюс Магадан»')],
    [c('Оборудование:', bold=True),      c('Автосамосвал NTE200, ДВС Cummins QSK50')],
    [c('Дата анализа:', bold=True),      c('20.05.2026')],
    [c('Случаев отказов:', bold=True),   c('7 случаев, 5 единиц техники')],
    [c('Сравнение с:', bold=True),       c('Komatsu 730E-DC (тот же объект, те же условия — без отказов клапанов)')],
]
t = Table(info, colWidths=[5*cm, 12*cm])
t.setStyle(TableStyle([
    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
    ('BACKGROUND',(0,0),(0,-1),C_LIGHT),
    ('FONTNAME',(0,0),(-1,-1),'DV'), ('FONTSIZE',(0,0),(-1,-1),10),
    ('LEFTPADDING',(0,0),(-1,-1),5), ('TOPPADDING',(0,0),(-1,-1),4),
    ('BOTTOMPADDING',(0,0),(-1,-1),4),
]))
story += [t, PageBreak()]

# ── ГЛАВНЫЙ ВЫВОД (сразу после титула) ────────────────────
story += [p('ГЛАВНЫЙ ВЫВОД', S_H1), hr(), sp(4)]
story.append(box(
    'На том же объекте (АО «Полюс Магадан»), в тех же климатических и горнодобывающих условиях '
    'эксплуатируются самосвалы Komatsu 730E-DC с ДВС QSK50. '
    'Отказов клапанного механизма на 730E-DC НЕ ЗАФИКСИРОВАНО. '
    'Единственное принципиальное отличие — настройки ECM (электронного управления двигателем). '
    'На NTE200 отключены или ослаблены ключевые функции защиты двигателя, '
    'что приводит к систематическому перегреву клапанного механизма.',
    label='ПЕРВОПРИЧИНА:', bg=C_RED_L, w=17*cm))
story.append(sp(6))
story.append(box(
    'Зольные отложения масла (Ca, Zn, P) и карьерная пыль (Si, Al) подтверждены лабораторно (NHL MS&T2026033) '
    'как продукт износа и усугубляющий фактор, но НЕ являются первопричиной. '
    'Отложения накапливаются быстрее именно вследствие повышенных температур '
    'от некорректных настроек ECM.',
    label='УСУГУБЛЯЮЩИЙ ФАКТОР:', bg=C_YELL, w=17*cm))
story.append(PageBreak())

# ── ЧАСТЬ 1. ECM — СРАВНЕНИЕ ──────────────────────────────
story += [p('ЧАСТЬ 1. КРИТИЧЕСКИЕ ОТЛИЧИЯ ECM — ЗАЩИТА ДВИГАТЕЛЯ', S_H1), hr(), sp(4)]
story.append(p(
    'Сравнение параметров ECM Cummins QSK50 на Komatsu 730E-DC и NTE200 '
    '(тот же двигатель, тот же объект эксплуатации). '
    'Красным выделены параметры, создающие риск перегрева и ускоренного износа клапанов.', S_BODY))
story.append(sp(6))

# Полная таблица ECM
ecm_rows = [
    # заголовок
    [c('Параметр ECM', bold=True, al=TA_CENTER),
     c('730E-DC', bold=True, al=TA_CENTER),
     c('NTE200', bold=True, al=TA_CENTER),
     c('Ед.', bold=True, al=TA_CENTER),
     c('Влияние на клапаны', bold=True, al=TA_CENTER)],
    # строки
    [c('Fuel Code (карта впрыска)'),
     c('FC0BU52'), c('FC0WH95'), c('—'),
     c('Разные карты Rail Pressure, Timing, Smoke Limiter, Boost — разный режим горения')],

    [c('Дерейт RPM по T масла\nC_EPD_OT_RPM_Drt_En', sz=8),
     c('ВКЛ (1)', col=C_GRN, bold=True),
     c('ВЫКЛ (0)', col=C_RED, bold=True),
     c('флаг'),
     c('КРИТИЧНО: NTE200 не снижает обороты при перегреве масла; 730E — снижает автоматически',
       col=C_RED)],

    [c('Обороты дерейта по T масла\nC_EPD_OT_Spd_Derate', sz=8),
     c('2200'), c('1675'), c('RPM'),
     c('Флаг выключен на NTE200 — значение не используется; на 730E работает')],

    [c('Время нарастания дерейта (масло)\nC_EPD_CP_TB_Time_To_Max_Trq_Drt', sz=8),
     c('0 с (мгн.)', col=C_GRN),
     c('51 сек', col=C_RED, bold=True),
     c('с'),
     c('NTE200: 51 секунда на полной нагрузке при критически низком давлении масла', col=C_RED)],

    [c('Время нарастания дерейта (ОЖ)\nC_EPD_CT_TB_Time_To_Max_Trq_Drt', sz=8),
     c('0 с (мгн.)', col=C_GRN),
     c('43 сек', col=C_RED, bold=True),
     c('с'),
     c('NTE200: 43 секунды при перегреве ОЖ без снижения нагрузки', col=C_RED)],

    [c('Макс. рабочие обороты\nC_ABS_Max_Acctr_Speed', sz=8),
     c('1990'), c('2100', col=C_WARN, bold=True), c('RPM'),
     c('+110 RPM → температура выхлопных газов выше на ~25–35°C, частота ударов клапана выше',
       col=C_WARN)],

    [c('Макс. Droop регулятора\nC_ABS_DroopMax', sz=8),
     c('0% (изохрон)', col=C_GRN),
     c('20%', col=C_WARN, bold=True),
     c('%'),
     c('730E держит жёсткие постоянные обороты; NTE200 допускает просадку/всплеск +20%')],

    [c('Порог горячего останова (момент)\nC_HSST_Net_Torque_High_Thd', sz=8),
     c('5509'), c('3650'), c('Н·м'),
     c('NTE200 срабатывает на 1859 Н·м раньше — возможно агрессивная нагрузка до останова')],

    [c('Нижний порог горячего останова\nC_HSST_Net_Torque_Low_Thd', sz=8),
     c('4722'), c('1850', col=C_WARN, bold=True), c('Н·м'),
     c('В 2,5 раза меньше — нестабильный гистерезис останова', col=C_WARN)],

    [c('Лимит коррекции топлива TIB\nC_TIB_Fuel_Adjust_Upper_Limit', sz=8),
     c('100%'), c('200%', col=C_RED, bold=True), c('%'),
     c('NTE200 может подавать в 2× больше топлива по коррекции — риск переобогащения, нагар',
       col=C_RED)],

    [c('Порог включения вентилятора (наддув)\nC_FCC_Charge_Temp_Y', sz=8),
     c('50°C'), c('46°C'), c('°C'),
     c('NTE200 агрессивнее охлаждает наддув (на 4°C раньше) — компенсация, не защита')],

    [c('Триггер выхода Charge Air Temp HIGH\nC_DO_01_A_High_Threshold', sz=8),
     c('54°C', col=C_GRN),
     c('ВЫКЛ', col=C_RED, bold=True),
     c('°C'),
     c('Только 730E сигнализирует о перегреве наддувочного воздуха', col=C_RED)],

    [c('Триггер выхода Coolant Temp HIGH\nC_DO_07_A_High_Threshold', sz=8),
     c('85°C', col=C_GRN),
     c('ВЫКЛ', col=C_RED, bold=True),
     c('°C'),
     c('Только 730E сигнализирует о перегреве ОЖ на внешний контроллер', col=C_RED)],

    [c('Триггер выхода Fuel Temp HIGH\nC_DO_08_A_High_Threshold', sz=8),
     c('68°C', col=C_GRN),
     c('ВЫКЛ', col=C_RED, bold=True),
     c('°C'),
     c('Только 730E сигнализирует о перегреве топлива', col=C_RED)],

    [c('Макс. RPM позиция 2/3\nT_ABS_MaxRef_2/3', sz=8),
     c('1990'), c('2225.8', col=C_WARN, bold=True), c('RPM'),
     c('При определённых условиях NTE200 может разгоняться до 2225 RPM', col=C_WARN)],
]

col_w = [4.2*cm, 2.0*cm, 2.0*cm, 1.3*cm, 7.5*cm]
t = Table(ecm_rows, colWidths=col_w, repeatRows=1)
ts = hdr_style(C_DARK)
# Подсветить критические строки
crit_rows = [2, 4, 5, 10]  # индексы строк данных (0=заголовок)
for r in crit_rows:
    ts.append(('BACKGROUND', (0,r),(-1,r), colors.HexColor('#FFF0F0')))
t.setStyle(TableStyle(ts))
story += [t, sp(8)]

# Итоговый блок по ECM
story.append(box(
    '1. Дерейт по температуре масла ОТКЛЮЧЁН (флаг=0) — при перегреве масла NTE200 продолжает '
    'работать на полных оборотах. 730E автоматически снижается до 1675 RPM.\n'
    '2. Задержка дерейта по давлению масла 51 сек и по температуре ОЖ 43 сек — '
    'двигатель работает в критическом режиме почти минуту.\n'
    '3. Максимальные обороты +110 RPM (2100 vs 1990) → температура выхлопных газов '
    'выше на ~25–35°C → ускоренный износ клапанов.\n'
    '4. Лимит коррекции топлива TIB 200% vs 100% → возможное переобогащение → '
    'нагар → зольные отложения на клапанах.\n'
    '5. Все три выходных триггера температуры (наддув, ОЖ, топливо) ОТКЛЮЧЕНЫ — '
    'система управления карьером не получает сигналов о перегреве.',
    label='КЛЮЧЕВЫЕ ПРОБЛЕМЫ ECM NTE200:', bg=C_RED_L, w=17*cm))
story.append(PageBreak())

# ── ЧАСТЬ 2. МЕХАНИЗМЫ ИЗНОСА (кратко) ───────────────────
story += [p('ЧАСТЬ 2. МЕХАНИЗМЫ ИЗНОСА И РОЛЬ ECM', S_H1), hr(), sp(4)]

story.append(p('2.1 Цепочка отказа — как настройки ECM приводят к износу клапанов', S_H2))
chain = [
    ('<b>ECM: отключена защита по T масла и завышены обороты</b>',
     'Двигатель работает при перегреве масла на полных оборотах (2100 RPM). '
     'Температура масла растёт, вязкость падает — теплоотвод от клапана через седло снижается.'),
    ('<b>Повышенная температура клапана</b>',
     'Exhaust valve на NTE200 — Inconel 751 без наплавки Stellite (подтверждено лаб. NHL). '
     'При T выхлопа выше на 25–35°C и без автозащиты — тарелка клапана перегревается, '
     'начинается пластическая деформация (заусенец на кромке).'),
    ('<b>Нарушение посадки клапана</b>',
     'Деформированная кромка → неплотная посадка → прорыв горячих газов → '
     'локальный перегрев → эрозия обеих поверхностей.'),
    ('<b>Зольные отложения как усугубление</b>',
     'На горячих поверхностях масло (Ca, Zn, P) и пыль карьера (Si, Al) '
     'образуют твёрдые отложения — они ускоряют абразивный износ, '
     'но не являются первопричиной. Аналогичные отложения на 730E '
     'не приводят к отказу — потому что температура там под контролем ECM.'),
    ('<b>Системный характер отказов</b>',
     'Одновременно 10 клапанов (ед. №47) и 17 клапанов (ед. №48) — '
     'это не случайные дефекты, это системный перегрев от общей причины: ECM.'),
]
for i, (title, desc) in enumerate(chain, 1):
    story.append(p(f'<b>Шаг {i}.</b> {title}', S_BODY))
    story.append(p(desc, s('cd', sz=9, sb=0, sa=1)))
    if i < len(chain):
        story.append(p('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;↓',
                        s('ar', sz=12, col=C_RED, sb=0, sa=0)))
    story.append(sp(2))

story.append(sp(6))
story.append(p('2.2 Почему 730E не имеет проблем с клапанами', S_H2))
cmp = [
    [c('Условие', bold=True, al=TA_CENTER), c('730E-DC', bold=True, al=TA_CENTER), c('NTE200', bold=True, al=TA_CENTER)],
    [c('Перегрев масла'), c('RPM → 1675 (авто)', col=C_GRN), c('RPM не снижается', col=C_RED)],
    [c('Критическое давление масла'), c('Мгновенный дерейт', col=C_GRN), c('51 сек на полной мощности', col=C_RED)],
    [c('Перегрев ОЖ'), c('Мгновенный дерейт', col=C_GRN), c('43 сек на полной мощности', col=C_RED)],
    [c('Макс. обороты'), c('1990 RPM', col=C_GRN), c('2100 RPM (+5.5%)', col=C_WARN)],
    [c('Сигнал оператору о перегреве'), c('Есть (3 выхода)', col=C_GRN), c('Отсутствует (ВЫКЛ)', col=C_RED)],
    [c('Коррекция топлива (макс.)'), c('100%', col=C_GRN), c('200%', col=C_WARN)],
    [c('Итог'), c('Клапаны в норме', col=C_GRN, bold=True), c('Массовый износ', col=C_RED, bold=True)],
]
t = Table(cmp, colWidths=[5.5*cm, 5.5*cm, 6*cm])
t.setStyle(TableStyle(hdr_style(C_DARK)))
story += [t, PageBreak()]

# ── ЧАСТЬ 3. ЛАБОРАТОРНЫЙ АНАЛИЗ ─────────────────────────
story += [p('ЧАСТЬ 3. ЛАБОРАТОРНЫЙ АНАЛИЗ NHL (MS&T2026033)', S_H1), hr(), sp(4)]
story.append(p('Инженер: Tao Lang  |  Утверждён: Shi Shi Liang  |  Дата: 10.03.2026  |  '
               'ДВС: 33232926 (ед. №48)  |  Наработка: 15 126 м/ч', S_SMALL))
story.append(sp(6))

lab = [
    [c('Параметр',bold=True,al=TA_CENTER), c('Факт',bold=True,al=TA_CENTER),
     c('Требование',bold=True,al=TA_CENTER), c('Соответствие',bold=True,al=TA_CENTER)],
    [c('Материал клапана'), c('Inconel 751 (CES51005)'), c('CES51005'), c('✓ ДА', col=C_GRN)],
    [c('Твёрдость клапана'), c('40,0–44,1 HRC'), c('≤46 HRC'), c('✓ ДА', col=C_GRN)],
    [c('Твёрдость седла'), c('56 HRC'), c('52–62 HRC'), c('✓ ДА', col=C_GRN)],
    [c('Наплавка тарелки клапана'), c('ОТСУТСТВУЕТ', col=C_RED, bold=True),
     c('По чертежу — отсутствует'), c('⚠ КОНСТРУКТИВНЫЙ ФАКТОР', col=C_WARN)],
    [c('Состав отложений'), c('Ca, Zn, P (зола масла)\n+ Si, Al (пыль)'),
     c('—'), c('Усугубляющий фактор')],
    [c('Характер износа'), c('Пластическая деформация\n(заусенец) — признак перегрева'),
     c('—'), c('Следствие T от ECM', col=C_RED)],
]
t = Table(lab, colWidths=[4.5*cm, 4.5*cm, 4.5*cm, 3.5*cm])
t.setStyle(TableStyle(hdr_style(C_DARK)))
story += [t, sp(8)]
story.append(box(
    '«Просадка (износ) рабочей поверхности клапана произошла в результате изнашивания под действием '
    'большого количества отложений масляной золы в условиях ПОВЫШЕННОЙ ТЕМПЕРАТУРЫ». '
    'Повышенная температура = прямое следствие некорректных настроек ECM NTE200.',
    label='ВЫВОД ЛАБОРАТОРИИ (переосмысление):', bg=C_YELL))
story.append(sp(8))

story.append(p('3.2 Фотоматериалы лабораторного анализа', S_H2))
lab_pdf = BASE + "气门盘部磨损分析报告-1.docx RUS.pdf"
for pg, cap in [
    (1,  'Рис. 1 — Макроскопический вид изношенных клапанов (ДВС 33232926, 15 126 м/ч)'),
    (2,  'Рис. 2 — Морфология отложений: чёрный слой + белая зольная основа (SEM)'),
    (4,  'Рис. 4 — EDS-анализ: Ca, Zn, P = зола масла; Si, Al = пыль карьера'),
    (7,  'Рис. 7 — Кратеры и царапины на тарелке клапана'),
    (9,  'Рис. 9 — ЗАУСЕНЕЦ на тарелке = пластическая деформация от перегрева'),
    (10, 'Рис. 10 — Твёрдость тарелки 40,0–44,1 HRC (без наплавки Stellite)'),
]:
    try:
        story += [sp(4), pdf_img(lab_pdf, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# ── ЧАСТЬ 4. АНАЛИЗ СЛУЧАЕВ ──────────────────────────────
story += [p('ЧАСТЬ 4. АНАЛИЗ КОНКРЕТНЫХ СЛУЧАЕВ', S_H1), hr(), sp(4)]
story.append(p('Сводная таблица отказов', S_H2))
cases = [
    [c('Ед.',bold=True,al=TA_CENTER), c('ДВС S/N',bold=True,al=TA_CENTER),
     c('Дата',bold=True,al=TA_CENTER), c('Наработка',bold=True,al=TA_CENTER),
     c('Тип отказа',bold=True,al=TA_CENTER), c('Цил.',bold=True,al=TA_CENTER)],
    [c('58'),c('33229431'),c('10.08.25'),c('10 702 м/ч'),c('Клапан/седло + толкатель'),c('7L')],
    [c('51'),c('33231742'),c('17.07.25'),c('10 783 м/ч'),c('Трещина ГБЦ → ОЖ в масло'),c('6R')],
    [c('47'),c('33232872'),c('07.12.25'),c('12 890 м/ч'),c('10 клапанов, 7 ГБЦ'),c('10 цил.')],
    [c('48'),c('33232926'),c('23.11.25'),c('14 997 м/ч'),c('17 клапанов, 11 ГБЦ + турбо'),c('11 цил.')],
    [c('48'),c('33232926'),c('31.03.26'),c('15 126 м/ч'),c('ОЖ в масло, гидроудар'),c('6L')],
    [c('83',col=C_RED,bold=True),c('33238542'),c('21.03.26'),
     c('2 776 м/ч',col=C_RED,bold=True),c('ОЖ в масло, разрушение клапана'),c('6L')],
    [c('58'),c('33229431'),c('16.03.26'),c('~15 572 км'),c('Износ выпускного клапана/седла'),c('4R')],
]
t = Table(cases, colWidths=[1.2*cm, 2.8*cm, 2.2*cm, 2.5*cm, 6.3*cm, 2*cm])
t.setStyle(TableStyle(hdr_style(C_DARK)))
story += [t, sp(10)]

# 4.1 — случай 47
story.append(p('4.1 Единица №47 — ДВС 33232872 (07.12.2025, 12 890 м/ч)', S_H2))
for item in ['Жалобы: сильное дымление, потеря мощности',
             'INSITE: неисправные форсунки 2L, 4L, 5R — неполное сгорание → дополнительный нагрев клапанов',
             'Дефектовка: 10 дефектных выпускных клапанов на 7 ГБЦ одновременно',
             '4 дефектные форсунки заменены']:
    story.append(p(f'• {item}', S_BUL))
story.append(sp(4))
story.append(box('Некорректные настройки ECM (нет дерейта по T масла, завышены обороты) → '
                 'базовая температура клапанов выше. Неисправные форсунки — '
                 'дополнительный нагрев. Суммарный эффект: одновременный износ 10 клапанов на 7 ГБЦ.',
                 label='ВЫВОД:', bg=C_GREEN))

pdf47 = BASE + "Тех отчет  47.pdf"
for pg, cap in [
    (1,'Ед. №47 — Дефектные выпускные клапаны с отложениями после демонтажа'),
    (2,'Ед. №47 — Замеры тепловых зазоров (занижены — признак просадки тарелки)'),
    (3,'Ед. №47 — Новые форсунки и клапаны после замены'),
]:
    try: story += [sp(4), pdf_img(pdf47, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# 4.2 — случай 48 (первый)
story.append(p('4.2 Единица №48 — ДВС 33232926 (23.11.2025, 14 997 м/ч)', S_H2))
for item in ['Синее дымление, потеря мощности, расход масла',
             'Код 0556 — топливо в картере (лимит коррекции TIB 200% → возможное переобогащение)',
             'Уровень масла +30 мм выше нормы — следствие дилюции форсунками',
             '17 дефектных выпускных клапанов на 11 ГБЦ',
             'Разрушен ротор турбокомпрессора ВД правого ряда']:
    story.append(p(f'• {item}', S_BUL))
story.append(sp(4))
story.append(box('Комбинация: некорректный ECM (нет защиты по T масла, обороты 2100) + '
                 'разрушенный турбокомпрессор (снижение наддува → дальнейший рост T) + '
                 'форсунки (код 0556) → системный износ 17 клапанов на 11 ГБЦ.',
                 label='ВЫВОД:', bg=C_GREEN))

pdf48 = BASE + "Тех.отчет 48.pdf"
for pg, cap in [
    (2,'Ед. №48 — Разрушенные выпускные клапаны ГБЦ 6L'),
    (3,'Ед. №48 — Дефектные клапаны нескольких ГБЦ'),
    (4,'Ед. №48 — Разрушённый ротор турбокомпрессора ВД'),
    (6,'Ед. №48 — Новые ГБЦ и клапаны после замены'),
]:
    try: story += [sp(4), pdf_img(pdf48, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# 4.3 — случай 48 (гидроудар)
story.append(p('4.3 Единица №48 — ДВС 33232926 (31.03.2026, 15 126 м/ч) — гидроудар', S_H2))
for item in ['Аварийная остановка, ОЖ в поддоне картера',
             'Трещина гильзы цилиндра 6L → гидроудар',
             'Разрушения: гильза 6L, шатунный подшипник, шейка коленвала, штанга толкателя',
             'Рекомендация завода: полная замена ДВС']:
    story.append(p(f'• {item}', S_BUL))
story.append(sp(4))
story.append(box('Гидроудар — СЛЕДСТВИЕ предыдущего отказа. '
                 'Трещина гильзы 6L от тепловой усталости после продолжительного перегрева '
                 '(ECM не обеспечил своевременный дерейт).',
                 label='ВЫВОД:', bg=C_ORG))

pdf48b = BASE + "Тех оточет 48 от 31.03.2026.pdf"
for pg, cap in [
    (0,'Ед. №48 — Разрушения в поддоне картера после гидроудара'),
    (1,'Ед. №48 — Разрушение гильзы 6L, состояние клапанов'),
    (2,'Ед. №48 — Шатунный подшипник и шейка коленвала 6L'),
]:
    try: story += [sp(4), pdf_img(pdf48b, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# 4.4 — случай 83
story.append(p('4.4 Единица №83 — ДВС 33238542 (21.03.2026, 2 776 м/ч)', S_H2))
story.append(p('КРИТИЧЕСКИ МАЛАЯ НАРАБОТКА — 2 776 м/ч при катастрофическом отказе', S_WARN))
story.append(sp(4))
for item in ['ГБЦ 6L: разрушенный выпускной клапан, приварившийся впускной',
             'Разрушения поршня, юбки, пальца, пружин, штанг толкателей',
             'INSITE: цил. 3 → 329°C, цил. 6 → 323°C (аномально высокие)',
             'Повторяемость: цилиндр 6L — отказ также у ед. №48']:
    story.append(p(f'• {item}', S_BUL))
story.append(sp(4))
story.append(box('Аномально высокие температуры цил. 3 и 6 при 2 776 м/ч — '
                 'прямое следствие отключённой защиты ECM. '
                 'Отсутствие дерейта по T масла привело к раннему разрушению.',
                 label='ВЫВОД:', bg=C_RED_L))

pdf83 = BASE + "Тех отчет 83 от 30.03.2026.pdf"
for pg, cap in [
    (1,'Ед. №83 — Разрушенный клапан ГБЦ 6L, приварившийся впускной'),
    (2,'Ед. №83 — Разрушения поршня и юбки цилиндра 6L'),
    (3,'Ед. №83 — Разрушения пальца поршня, штанги толкателей'),
    (4,'Ед. №83 — INSITE: T выхлопа цил. 3 (329°C) и цил. 6 (323°C)'),
]:
    try: story += [sp(4), pdf_img(pdf83, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# 4.5 — случай 58
story.append(p('4.5 Единица №58 — ДВС 33229431 (два эпизода)', S_H2))
story.append(p('Эпизод 1: 10.08.2025, 10 702 м/ч — цилиндр 7L', S_H3))
story.append(p('Посторонний стук → дефекты клапанов, сёдел, рычага и штанги толкателя 7L.', S_BODY))
story.append(p('Эпизод 2: 16.03.2026 — цилиндр 4R (повторный отказ, тот же двигатель)', S_H3))
story.append(p('Синее дымление. Износ выпускного клапана 4R (арт. 3035110) и седла (арт. 3086192).', S_BODY))
story.append(p('Заключение сервиса: «низкое качество материала» — ОПРОВЕРГНУТО лабораторией NHL.',
               s('it', sz=9, col=C_GREY, italic=True)))
story.append(sp(4))
story.append(box('Повторный отказ на том же двигателе за 8 месяцев. '
                 'Причина та же — некорректный ECM. '
                 'Замена клапанов без перепрограммирования ECM — гарантия нового отказа.',
                 label='ВЫВОД:', bg=C_GREEN))

pdf58 = BASE + "58 клапан гбц 16.03.26.pdf"
for pg, cap in [
    (1,'Ед. №58 — Изношенный выпускной клапан и седло цилиндра 4R'),
    (2,'Ед. №58 — ГБЦ после ремонта'),
]:
    try: story += [sp(4), pdf_img(pdf58, pg), p(cap, S_CAP)]
    except: pass

story.append(PageBreak())

# ── ЧАСТЬ 5. ВЫВОДЫ ───────────────────────────────────────
story += [p('ЧАСТЬ 5. ИТОГОВЫЕ ВЫВОДЫ И РЕКОМЕНДАЦИИ', S_H1), hr(), sp(4)]

story.append(p('5.1 Матрица выводов', S_H2))
matrix = [
    [c('Версия',bold=True,al=TA_CENTER), c('Статус',bold=True,al=TA_CENTER),
     c('Обоснование',bold=True,al=TA_CENTER)],
    [c('ECM: отключена защита по T масла'), c('✓ ПЕРВОПРИЧИНА', col=C_RED, bold=True),
     c('730E (те же условия) — защита включена, отказов нет')],
    [c('ECM: завышены обороты (+110 RPM)'), c('✓ ПЕРВОПРИЧИНА', col=C_RED, bold=True),
     c('+25–35°C к T выхлопа, выше частота ударов клапана')],
    [c('ECM: задержка дерейта 43–51 сек'), c('✓ ПЕРВОПРИЧИНА', col=C_RED, bold=True),
     c('До 51 сек на полной нагрузке при критических условиях')],
    [c('Зольные отложения масла и пыли'), c('⚠ УСУГУБЛЕНИЕ', col=C_WARN, bold=True),
     c('Лаб. NHL подтверждено, но на 730E не приводит к отказу')],
    [c('Отсутствие наплавки Stellite'), c('⚠ КОНСТРУКТИВНЫЙ ФАКТОР', col=C_WARN, bold=True),
     c('Снижает порог износа, усугубляет при перегреве')],
    [c('Низкое качество материала клапана'), c('✗ ОПРОВЕРГНУТО', col=C_GRN, bold=True),
     c('NHL MS&T2026033: состав и твёрдость соответствуют чертежу')],
]
t = Table(matrix, colWidths=[5*cm, 3.5*cm, 8.5*cm])
t.setStyle(TableStyle(hdr_style(C_DARK)))
story += [t, sp(10)]

story.append(p('5.2 Рекомендации', S_H2))
recs = [
    ('НЕМЕДЛЕННО — перепрограммирование ECM NTE200',
     'Включить дерейт по температуре масла (C_EPD_OT_RPM_Drt_En = 1). '
     'Снизить максимальные обороты до уровня 730E (1990 RPM). '
     'Установить мгновенный дерейт по давлению масла и температуре ОЖ (0 сек). '
     'Ограничить лимит коррекции TIB до 100%. '
     'Включить выходные триггеры температуры (наддув, ОЖ, топливо).'),
    ('НЕМЕДЛЕННО — диагностика парка',
     'Проверка форсунок через INSITE на всех единицах. '
     'Замена дефектных форсунок до любого клапанного ремонта. '
     'Контроль уровня масла — при уровне выше нормы диагностика на дилюцию.'),
    ('СРЕДНЕСРОЧНО — конструктивные меры',
     'Запрос в NHL на выпускные клапаны с наплавкой Stellite. '
     'Переход на масло с низким SAPS (CES 20081). '
     'Расследование трещин ГБЦ позиции 6 (ед. №48, №51, №83).'),
    ('ДОЛГОСРОЧНО — мониторинг',
     'Систематическая запись температур по цилиндрам через INSITE. '
     'Отклонение >30°C от среднего — немедленная диагностика. '
     'Проверка тепловых зазоров каждые 2 000 м/ч (вместо 4 000 м/ч).'),
]
for i, (title, desc) in enumerate(recs, 1):
    story.append(p(f'<b>{i}. {title}:</b>', S_BODY))
    story.append(p(desc, s('rd', sz=9, sb=0, sa=4)))

story.append(sp(10))
story.append(HRFlowable(width='100%', thickness=0.5, color=colors.grey))
story.append(p(
    'Анализ выполнен на основании фактических данных: отчёты №47, №48, №51, №58, №83, '
    'лабораторный анализ NHL MS&T2026033, данные INSITE N43/N84, сравнение ECM 730E-DC/NTE200. '
    'Дата: 20.05.2026',
    s('ft', sz=8, col=C_GREY, al=TA_CENTER, italic=True)))

# ── СБОРКА ────────────────────────────────────────────────
out = BASE + "Анализ_износа_клапанов_QSK50_NTE200_v2.pdf"
doc_pdf = SimpleDocTemplate(
    out, pagesize=A4,
    leftMargin=2*cm, rightMargin=1.5*cm,
    topMargin=2*cm, bottomMargin=2*cm,
    title="Анализ износа клапанов QSK50 NTE200 v2",
    author="АО Полюс Магадан",
)
doc_pdf.build(story)
sz = os.path.getsize(out)/1024/1024
print(f"ГОТОВО: {out} ({sz:.1f} МБ)")
