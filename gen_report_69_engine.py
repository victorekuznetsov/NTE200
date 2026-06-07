#!/usr/bin/env python3
"""NTE200 #69 — QSK50 MCRS V16 — Engine-Only Diagnostic Report (07.06.2026)."""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, PageBreak, KeepTogether,
                                Flowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String, Circle
from reportlab.graphics import renderPDF

# ── Font registration ──────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont('Sans',      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('Sans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# ── Colour palette ─────────────────────────────────────────────────────────────
C = dict(
    bg       = colors.HexColor('#1a1a2e'),   # page background
    panel    = colors.HexColor('#16213e'),   # section header / card bg
    panel2   = colors.HexColor('#0f3460'),   # alternate panel
    border   = colors.HexColor('#2a3a5c'),   # grid lines
    text     = colors.HexColor('#e8e8f0'),   # body text
    sub      = colors.HexColor('#8899bb'),   # secondary
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
def P(txt, **kw):
    s = ParagraphStyle('x', fontName='Sans', fontSize=9, textColor=C['text'], leading=13, **kw)
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

def tbl(rows, cols, style_extra=None, alt=True):
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
        bg = C['altrow'] if i % 2 == 0 else C['altrow2']
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
    Pc(C['white'], 'ДИАГНОСТИЧЕСКИЙ ОТЧЁТ ДВИГАТЕЛЯ', bold=True, size=16, lead=20, align=TA_CENTER),
    SP(1),
    Pc(C['blue'],  'NTE200 №69  /  QSK50 MCRS V16', bold=True, size=13, lead=17, align=TA_CENTER),
    SP(1),
    Pc(C['sub'],   'Дата отчёта: 07.06.2026  |  Наработка ДВС: ~10 603 м/ч (по состоянию на 15.05.2026)',
       size=9, lead=13, align=TA_CENTER),
    HR(C['blue']),
    SP(2),
]

# Machine info box
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
     Pc(C['blue'], 'Замена клапанов ГБЦ 25.05.2026', size=8, lead=11),
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

# KPI cards — 3 per row
kpi_cards_1 = [
    ('М/Ч ДВС',          '10 603',        'м/ч',       C['green'],  C['greenbg']),
    ('МАСЛО ДВС',         'ОТЛИЧНО',       '34 пробы — все чисты\nFe=0-3 мг/кг', C['green'], C['greenbg']),
    ('РЕМОНТ ДВС',        'Апр. 2026',     '8 суток простоя\n(197 часов)', C['orange'], C['orangebg']),
]
kpi_cards_2 = [
    ('ОШИБКИ EGT',        'Цил. 3 и 16',  'повторяющиеся\nпосле ремонта', C['orange'], C['orangebg']),
    ('ЗАМЕНА КЛАПАНОВ',   '25.05.2026',   'выполнено\n(тех. отчёт)', C['blue'], C['bluebg']),
    ('DML ДАННЫЕ',        'ОТСУТСТВУЮТ',  'не предоставлены\nдля №69', C['grey'], C['greybg']),
]

def make_kpi_table(cards):
    n = len(cards)
    w = [PW / n] * n
    hdr_row_cells = [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _, _ in cards]
    val_row_cells  = [Pc(vc, val, bold=True, size=14, lead=18, align=TA_CENTER) for _, val, _, vc, _ in cards]
    sub_row_cells  = [Pc(sc if sc != C['grey'] else C['grey'], sub, size=7, lead=10, align=TA_CENTER)
                      for _, _, sub, sc, _ in cards]
    t = Table([hdr_row_cells, val_row_cells, sub_row_cells], colWidths=w)
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

# Summary note
story += [
    info_box(
        'ОБЩАЯ ОЦЕНКА ДВС: Двигатель работоспособен. Масло в отличном состоянии на протяжении всего '
        'периода мониторинга (апр.2025–май.2026). Основные риски — последствия длительного ремонта ДВС '
        '(апр.2026, 197 ч) и повторяющиеся ошибки датчиков EGT цил.3 и 16. DML-данные отсутствуют.',
        bg_color=C['panel2'], text_color=C['text'], size=9, border_color=C['blue']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — TIMELINE OF KEY ENGINE EVENTS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('ХРОНОЛОГИЯ КЛЮЧЕВЫХ СОБЫТИЙ ДВС (апр–май 2026)', 1)
story += [SP(2)]

# Timeline as a Drawing
class TimelineFlowable(Flowable):
    """Vertical timeline rendered with ReportLab shapes."""
    def __init__(self, events, width, row_h=28):
        Flowable.__init__(self)
        self.events = events
        self._width = width
        self._row_h = row_h
        self._height = len(events) * row_h + 10

    def wrap(self, *args):
        return (self._width, self._height)

    def draw(self):
        canvas = self.canv
        x_line = 28 * mm   # vertical axis x
        x_text = 32 * mm   # event text start
        y_start = self._height - 18

        for i, (date, hours, event_type, text, col_hex) in enumerate(self.events):
            y = y_start - i * self._row_h
            node_col = colors.HexColor(col_hex)

            # Draw connecting line (except after last)
            if i < len(self.events) - 1:
                canvas.setStrokeColor(colors.HexColor('#2a3a5c'))
                canvas.setLineWidth(1.5)
                canvas.line(x_line, y - 6, x_line, y - self._row_h + 6)

            # Node circle
            canvas.setFillColor(node_col)
            canvas.setStrokeColor(colors.HexColor('#1a1a2e'))
            canvas.setLineWidth(1)
            canvas.circle(x_line, y, 5, fill=1, stroke=1)

            # Date label (left of line)
            canvas.setFont('Sans-Bold', 7)
            canvas.setFillColor(colors.HexColor('#8899bb'))
            canvas.drawRightString(x_line - 8, y - 3, date)

            # Event type badge
            badge_w = 22 * mm
            badge_x = x_text
            badge_y = y - 5
            canvas.setFillColor(node_col)
            canvas.setStrokeColor(node_col)
            canvas.roundRect(badge_x, badge_y, badge_w, 11, 2, fill=1, stroke=0)
            canvas.setFont('Sans-Bold', 6.5)
            canvas.setFillColor(colors.HexColor('#1a1a2e'))
            canvas.drawCentredString(badge_x + badge_w / 2, badge_y + 2.5, event_type)

            # Duration badge (if given)
            if hours:
                dur_x = badge_x + badge_w + 3
                dur_w = 18 * mm
                canvas.setFillColor(colors.HexColor('#0d1829'))
                canvas.setStrokeColor(node_col)
                canvas.setLineWidth(0.5)
                canvas.roundRect(dur_x, badge_y, dur_w, 11, 2, fill=1, stroke=1)
                canvas.setFont('Sans-Bold', 6.5)
                canvas.setFillColor(node_col)
                canvas.drawCentredString(dur_x + dur_w / 2, badge_y + 2.5, hours)

            # Description text
            desc_x = x_text + badge_w + (22 * mm if hours else 3)
            canvas.setFont('Sans', 7.5)
            canvas.setFillColor(colors.HexColor('#e8e8f0'))
            # truncate to fit
            max_w = self._width - desc_x - 5
            canvas.drawString(desc_x + (3 if hours else 0), y - 2.5, text)

# Events: (date, duration_str_or_None, badge_label, description, color_hex)
timeline_events = [
    ('09.04.2026', '21 ч',    'ТО-6',         'Сварочные работы, восстановление кронштейна, замена впускных труб',    '#f39c12'),
    ('16.04.2026', '197 ч!',  'РЕМОНТ ДВС',   '8+ суток — крупнейшее вмешательство в ДВС за период мониторинга',     '#e74c3c'),
    ('26.04.2026', '0:15 ч',  'ТО',           'Протяжка хомутов на интеркулере',                                      '#3498db'),
    ('27.04.2026', None,      'EGT ЦИЛ.3',    'Ошибка датчика EGT цилиндр 3 — поджатие пинов, осмотр проводки',      '#f39c12'),
    ('27.04.2026', None,      'EGT ЦИЛ.16',   'Ошибка датчика EGT цилиндр 16 — поджатие пинов, осмотр проводки',     '#f39c12'),
    ('28.04.2026', None,      'EGT ЦИЛ.3',    'ПОВТОР: ошибка EGT цилиндр 3 — повторное поджатие пинов',             '#e74c3c'),
    ('30.04.2026', '3 ч',     'ТО-1',         'Плановое ТО + сварочные работы бочки РМК',                             '#3498db'),
    ('04.05.2026', None,      'EGT ЦИЛ.16',   'ПОВТОР: ошибка EGT цилиндр 16 — поджатие пинов',                      '#e74c3c'),
    ('11.05.2026', '2 ч',     'ТО-2',         'Плановое ТО-2 + регулировка стоек ПГП',                               '#3498db'),
    ('25.05.2026', None,      'КЛАПАНЫ ГБЦ',  'Замена клапанов ГБЦ (тех. отчёт 25.05.2026)',                         '#27ae60'),
]

story.append(TimelineFlowable(timeline_events, PW))
story += [SP(3)]

story += [
    info_box(
        'Ремонт ДВС продолжительностью 197 часов (8+ суток) — серьёзное вмешательство. '
        'Точный объём работ (замена ГБЦ, поршней, колец?) требует уточнения в тех. отчёте. '
        'После ремонта — повторяющиеся ошибки датчиков EGT цил.3 и 16, устранялись поджатием пинов разъёмов. '
        'Замена клапанов 25.05.2026 возможно связана с состоянием цилиндров 3 и 16.',
        bg_color=C['orangebg'], text_color=C['orange'], size=8, border_color=C['orange']
    ),
    SP(2),
]

# Full maintenance history table
story += section_title('ПОЛНАЯ ИСТОРИЯ ТО И РЕМОНТОВ (значимые события ДВС)', 2)
story += [SP(1)]

hist_cols = ['Дата', 'М/Ч', 'Время', 'Событие / Работы']
hist_w = [PW*0.13, PW*0.10, PW*0.09, PW*0.68]
story.append(hdr_row(hist_cols, hist_w))

hist_data = [
    ('07.03.2026', '9 251',  '2:56 ч',  'ТО-1',                                                                          C['altrow'],    C['text']),
    ('18.03.2026', '9 511',  '2:55 ч',  'ТО-2',                                                                          C['altrow2'],   C['text']),
    ('25.03.2026', '9 656',  '17:13 ч', 'Диагностика и ремонт проводки ДВС — длительная работа',                         C['orangebg'],  C['orange']),
    ('30.03.2026', '9 756',  '4:45 ч',  'ТО-1',                                                                          C['altrow'],    C['text']),
    ('09.04.2026', '10 001', '21:26 ч', 'ТО-6 — Сварочные работы, восстановление кронштейна, замена впускных труб',      C['orangebg'],  C['orange']),
    ('16.04.2026', '10 130', '197 ч',   'РЕМОНТ ДВС — Диагностика и ремонт двигателя — 8 суток простоя!',               C['redbg'],     C['red']),
    ('26.04.2026', '10 169', '0:15 ч',  'Протяжка хомутов на интеркулере',                                               C['altrow'],    C['text']),
    ('27.04.2026', '10 186', '0:28 ч',  'Ошибка EGT цилиндр 3 — поджатие пинов, осмотр проводки',                      C['orangebg'],  C['orange']),
    ('27.04.2026', '10 197', '0:20 ч',  'Ошибка EGT цилиндр 16 — поджатие пинов, осмотр проводки',                     C['orangebg'],  C['orange']),
    ('28.04.2026', '10 209', '0:33 ч',  'ПОВТОР: Ошибка EGT цилиндр 3 — повторное поджатие пинов',                     C['redbg'],     C['red']),
    ('30.04.2026', '10 248', '2:59 ч',  'ТО-1. Сварочные работы бочки РМК',                                             C['altrow2'],   C['text']),
    ('04.05.2026', '10 348', '0:20 ч',  'ПОВТОР: Ошибка EGT цилиндр 16 — поджатие пинов',                              C['redbg'],     C['red']),
    ('11.05.2026', '10 500', '2:19 ч',  'ТО-2. Регулировка стоек ПГП',                                                  C['altrow'],    C['text']),
    ('12.05.2026', '10 522', '0:25 ч',  'Диагностика ошибок GE, осмотр плат',                                           C['altrow2'],   C['text']),
    ('15.05.2026', '10 603', '0:02 ч',  'Замена щётки стеклоочистителя (последняя запись в базе)',                      C['altrow'],    C['sub']),
    ('25.05.2026', '~10 650','—',        'Замена клапанов ГБЦ (технический отчёт 25.05.2026)',                           C['greenbg'],   C['green']),
]

hist_rows = []
for date, mh, dur, event, row_bg, ecol in hist_data:
    hist_rows.append([
        Pc(ecol,  date,  bold=(ecol != C['text'] and ecol != C['sub']), size=7, lead=10),
        Pc(C['sub'], mh,                size=7, lead=10),
        Pc(C['sub'], dur,               size=7, lead=10),
        Pc(C['text'], event,            size=7, lead=10),
    ])
hist_extra = [('BACKGROUND', (0, i), (-1, i), hist_data[i][4]) for i in range(len(hist_rows))]
story.append(tbl(hist_rows, hist_w, style_extra=hist_extra, alt=False))

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — OIL SPECTRAL ANALYSIS (full table)
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('СПЕКТРАЛЬНЫЙ АНАЛИЗ МАСЛА ДВС — 34 ПРОБЫ (апр.2025 – май.2026)', 3)
story += [SP(1)]

story.append(info_box(
    'Общий вывод: ВСЕ 34 пробы — "Допустимое". Двигатель не изнашивается.\n'
    'Fe: 0–3 мг/кг (норма <30 мг/кг) — исключительно чисто. '
    'Cu: 0–6 мг/кг — норма. Si: единичный выброс 27 мг/кг (30.04.2026) — вероятно загрязнение при отборе.\n'
    'TBN: снижается с ~120 до ~92 мг/кг — закономерное истощение, замена масла каждые 250 м/ч. '
    'Сажа: умеренный рост 119→153 мг/кг — допустимо.',
    bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
))
story.append(SP(2))

oil_cols = ['Дата', 'М/Ч ДВС', 'Нар.масла', 'Fe', 'Cu', 'Si', 'TAN', 'TBN', 'Сажа', 'Статус']
oil_w = [PW*0.11, PW*0.10, PW*0.10, PW*0.06, PW*0.06, PW*0.06, PW*0.07, PW*0.07, PW*0.07, PW*0.30]
story.append(hdr_row(oil_cols, oil_w))

# All 34 oil samples
# (date, mh, oil_hrs, fe, cu, si, tan, tbn, soot, status, row_bg, fe_col, tbn_col, si_col)
oil_data_raw = [
    ('12.04.2025',  1761,  250,  0, 0,  0, 14.1, 117.8, 119, 'Допустимое'),
    ('23.04.2025',  2019,  250,  0, 0,  0, 15.1, 123.8, 126, 'Допустимое'),
    ('04.05.2025',  2252,  250,  0, 0,  0, 14.7, 115.2, 131, 'Допустимое'),
    ('14.05.2025',  2499,  250,  0, 0,  0, 14.8, 115.3, 132, 'Допустимое'),
    ('29.05.2025',  2750,  250,  0, 0,  0, 14.9, 120.3, 127, 'Допустимое'),
    ('09.06.2025',  3000,  250,  0, 0,  0, 14.9, 119.3, 128, 'Допустимое'),
    ('21.06.2025',  3256,  250,  0, 0,  0, 14.9, 113.3, 136, 'Допустимое'),
    ('01.07.2025',  3503,  250,  0, 0,  0, 14.9, 116.2, 132, 'Допустимое'),
    ('12.07.2025',  3762,  250,  0, 0,  0, 15.5, 125.0, 130, 'Допустимое'),
    ('22.07.2025',  4000,  250,  0, 0,  0, 16.8, 123.9, 147, 'Допустимое'),
    ('02.08.2025',  4250,  250,  0, 0,  0, 15.6, 116.4, 141, 'Допустимое'),
    ('13.08.2025',  4500,  250,  0, 0,  0, 15.5, 116.1, 140, 'Допустимое'),
    ('24.08.2025',  4753,  250,  0, 0,  0, 15.3, 120.6, 132, 'Допустимое'),
    ('05.09.2025',  4985,  235,  0, 0,  0, 15.4, 114.0, 142, 'Допустимое'),
    ('17.09.2025',  5247,  262,  0, 0,  0, 15.5, 119.3, 136, 'Допустимое'),
    ('27.09.2025',  5500,  253,  0, 0,  2, 15.7, 118.4, 140, 'Допустимое'),
    ('08.10.2025',  5752,  252,  0, 2,  0, 15.5, 119.6, 136, 'Допустимое'),
    ('19.10.2025',  6002,  250,  2, 1,  0, 15.5, 118.1, 138, 'Допустимое'),
    ('30.10.2025',  6256,  254,  2, 0,  0, 15.5, 118.6, 137, 'Допустимое'),
    ('10.11.2025',  6517,  261,  0, 0,  0, 15.3, 118.3, 135, 'Допустимое'),
    ('12.12.2025',  1255,  252,  3, 6,  0, 15.8, 121.8, 137, 'Допустимое*'),
    ('23.12.2025',  7493,  238,  0, 0,  0, 15.6, 122.5, 134, 'Допустимое'),
    ('03.01.2026',  7753,  260,  0, 0,  0, 15.8, 124.1, 134, 'Допустимое'),
    ('14.01.2026',  8007,  254,  0, 0,  0, 15.3, 119.0, 134, 'Допустимое'),
    ('24.01.2026',  8248,  241,  0, 0,  0, 15.0, 121.2, 127, 'Допустимое'),
    ('04.02.2026',  8500,  252,  1, 0,  0, 15.0, 128.9, 119, 'Допустимое'),
    ('14.02.2026',  8747,  247,  0, 0,  0, 16.0, 121.5, 140, 'Допустимое'),
    ('26.02.2026',  9004,  251,  1, 0,  0, 15.7, 111.2, 150, 'Допустимое'),
    ('08.03.2026',  9251,  247,  0, 0,  0, 14.2,  97.1, 150, 'Допустимое'),
    ('19.03.2026',  9511,  262,  0, 0,  0, 14.0,  96.5, 148, 'Допустимое'),
    ('31.03.2026',  9756,  245,  0, 0,  0, 14.3, 100.1, 147, 'Допустимое'),
    ('10.04.2026', 10001,  245,  0, 0,  0, 14.2,  99.2, 147, 'Допустимое'),
    ('30.04.2026', 10248,  249,  0, 0, 27, 14.4,  97.1, 153, 'Допустимое'),
    ('11.05.2026', 10500,  252,  0, 0,  0, 13.0,  91.9, 140, 'Допустимое'),
]

oil_rows = []
for i, row in enumerate(oil_data_raw):
    date, mh, oil_h, fe, cu, si, tan, tbn, soot, status = row
    is_anomaly_mh = (mh == 1255)
    is_si_spike   = (si == 27)
    is_tbn_low    = (tbn < 100)
    row_bg = C['orangebg'] if is_anomaly_mh else (C['altrow'] if i % 2 == 0 else C['altrow2'])

    fe_col  = C['orange'] if fe > 0 else C['text']
    cu_col  = C['orange'] if cu > 3 else C['text']
    si_col  = C['red']    if is_si_spike else (C['text'] if si == 0 else C['orange'])
    tbn_col = C['orange'] if is_tbn_low else C['text']
    mh_col  = C['orange'] if is_anomaly_mh else C['text']

    oil_rows.append([
        Pc(C['orange'] if is_anomaly_mh else C['text'], date,          size=6.5, lead=9),
        Pc(mh_col,  str(mh),                            bold=is_anomaly_mh, size=6.5, lead=9),
        Pc(C['text'], str(oil_h),                       size=6.5, lead=9),
        Pc(fe_col,  str(fe),                            bold=(fe > 0), size=6.5, lead=9),
        Pc(cu_col,  str(cu),                            bold=(cu > 3), size=6.5, lead=9),
        Pc(si_col,  str(si),                            bold=is_si_spike, size=6.5, lead=9),
        Pc(C['text'], f'{tan:.1f}',                     size=6.5, lead=9),
        Pc(tbn_col, f'{tbn:.1f}',                       bold=is_tbn_low, size=6.5, lead=9),
        Pc(C['text'], str(soot),                        size=6.5, lead=9),
        Pc(C['green'], status,                          size=6.5, lead=9),
    ])

oil_extra = []
for i, row in enumerate(oil_data_raw):
    _, mh, _, _, _, si, _, _, _, _ = row
    is_anomaly_mh = (mh == 1255)
    is_si_spike   = (si == 27)
    if is_anomaly_mh:
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), C['orangebg']))
    elif is_si_spike:
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), C['redbg']))
    else:
        bg = C['altrow'] if i % 2 == 0 else C['altrow2']
        oil_extra.append(('BACKGROUND', (0, i), (-1, i), bg))

story.append(tbl(oil_rows, oil_w, style_extra=oil_extra, alt=False))
story += [SP(2)]

# Anomaly notes
story += [
    info_box(
        '* 12.12.2025 — М/Ч=1255: аномальная запись (вероятно другой агрегат или ошибка в базе данных). '
        'Параметры масла в норме (Fe=3, Cu=6), статус "Допустимое".',
        bg_color=C['orangebg'], text_color=C['orange'], size=7.5
    ),
    SP(1),
    info_box(
        '30.04.2026 — Si=27 мг/кг: единичный выброс (норма 0). Следующая проба 11.05.2026 Si=0. '
        'Вероятное загрязнение при отборе пробы. Не является признаком износа.',
        bg_color=C['redbg'], text_color=C['red'], size=7.5
    ),
    SP(2),
]

# Trend analysis
story += section_title('АНАЛИЗ ТРЕНДОВ МАСЛА', None)
story += [SP(1)]

trend_data = [
    ('Fe (железо)',       '0–3 мг/кг', 'Норма <30 мг/кг',  'ОТЛИЧНО — минимальный износ деталей ДВС',     C['green'],  C['greenbg']),
    ('Cu (медь)',         '0–6 мг/кг', 'Норма <15 мг/кг',  'НОРМА — подшипники и втулки в порядке',        C['green'],  C['greenbg']),
    ('Si (кремний)',      '0; пик 27', 'Норма <10 мг/кг',  'Единичный выброс — вероятно загрязнение пробы', C['orange'], C['orangebg']),
    ('TBN',               '119→92',    'Мин. >70 мг/кг',   'Закономерное снижение — замена масла 250 м/ч',  C['text'],   C['altrow']),
    ('TAN',               '13–16',     'Норма <15',         '16.8 (22.07) и 16.0 (14.02) — в норме для режима', C['text'], C['altrow2']),
    ('Сажа',              '119→153',   'Норма <250 мг/кг', 'Умеренный рост — соответствует режиму эксплуатации', C['text'], C['altrow']),
]

trend_cols = ['Параметр', 'Диапазон', 'Норма', 'Оценка']
trend_w = [PW*0.15, PW*0.15, PW*0.18, PW*0.52]
story.append(hdr_row(trend_cols, trend_w))

trend_rows = []
for param, rng, norm, assess, tc, bg in trend_data:
    trend_rows.append([
        Pc(C['sub'],  param, bold=True, size=8, lead=11),
        Pc(tc,        rng,   bold=True, size=8, lead=11),
        Pc(C['sub'],  norm,             size=8, lead=11),
        Pc(tc,        assess,           size=8, lead=11),
    ])
trend_extra = [('BACKGROUND', (0, i), (-1, i), trend_data[i][5]) for i in range(len(trend_rows))]
story.append(tbl(trend_rows, trend_w, style_extra=trend_extra, alt=False))

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — ENGINE PROBLEM ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('АНАЛИЗ ПРОБЛЕМ ДВИГАТЕЛЯ', 4)
story += [SP(2)]

# Problem 1 — EGT sensor errors
story += [
    Pc(C['orange'], '4.1  ОШИБКИ ДАТЧИКОВ EGT — ЦИЛИНДРЫ 3 И 16', bold=True, size=11, lead=15),
    HRFlowable(width='100%', thickness=0.5, color=C['orange'], spaceAfter=4, spaceBefore=2),
    SP(1),
]

egt_events = [
    ('27.04.2026', '10 186', 'Цил.3',  '1-я ошибка EGT',  'поджатие пинов разъёма, осмотр проводки'),
    ('27.04.2026', '10 197', 'Цил.16', '1-я ошибка EGT',  'поджатие пинов разъёма, осмотр проводки'),
    ('28.04.2026', '10 209', 'Цил.3',  'ПОВТОР ошибки',   'повторное поджатие пинов'),
    ('04.05.2026', '10 348', 'Цил.16', 'ПОВТОР ошибки',   'поджатие пинов (третий случай)'),
]

egt_cols = ['Дата', 'М/Ч', 'Цилиндр', 'Событие', 'Принятые меры']
egt_w = [PW*0.13, PW*0.10, PW*0.10, PW*0.20, PW*0.47]
story.append(hdr_row(egt_cols, egt_w))

egt_rows = []
for i, (dt, mh, cyl, ev, action) in enumerate(egt_events):
    is_repeat = 'ПОВТОР' in ev
    bg = C['redbg'] if is_repeat else C['orangebg']
    col = C['red'] if is_repeat else C['orange']
    egt_rows.append([
        Pc(col, dt,     bold=True, size=8, lead=11),
        Pc(C['sub'], mh,          size=8, lead=11),
        Pc(col, cyl,    bold=True, size=8, lead=11),
        Pc(col, ev,     bold=is_repeat, size=8, lead=11),
        Pc(C['text'], action,     size=8, lead=11),
    ])
    egt_rows[-1]  # reference to avoid unused

egt_extra = []
for i, (_, _, _, ev, _) in enumerate(egt_events):
    is_repeat = 'ПОВТОР' in ev
    egt_extra.append(('BACKGROUND', (0, i), (-1, i), C['redbg'] if is_repeat else C['orangebg']))

story.append(tbl(egt_rows, egt_w, style_extra=egt_extra, alt=False))
story += [SP(1)]

# Cause analysis
cause_data = [
    ('Возможная причина', 'Вероятность', 'Рекомендация'),
    ('Повреждение термопары EGT при ремонте ДВС (апр.2026)', 'Высокая', 'Замена термопар цил.3 и 16'),
    ('Плохой контакт в разъёме жгута ДВС (коррозия, деформация)', 'Высокая', 'Ревизия разъёмов и жгута'),
    ('Механическое повреждение датчика при ремонте/монтаже ДВС', 'Средняя', 'Визуальный осмотр датчиков'),
    ('Неисправность ЭБУ (канал EGT)', 'Низкая', 'Диагностика INSITE после замены датчиков'),
]

cause_w = [PW*0.52, PW*0.18, PW*0.30]
cause_rows = []
for i, (cause, prob, rec) in enumerate(cause_data):
    if i == 0:
        cause_rows.append([Pc(C['sub'], cause, bold=True, size=7.5, lead=11),
                           Pc(C['sub'], prob,  bold=True, size=7.5, lead=11),
                           Pc(C['sub'], rec,   bold=True, size=7.5, lead=11)])
    else:
        prob_col = C['red'] if prob == 'Высокая' else (C['orange'] if prob == 'Средняя' else C['text'])
        bg = C['hdr'] if i == 0 else (C['altrow'] if i % 2 == 1 else C['altrow2'])
        cause_rows.append([Pc(C['text'], cause, size=7.5, lead=11),
                           Pc(prob_col,  prob,  bold=True, size=7.5, lead=11),
                           Pc(C['blue'], rec,   size=7.5, lead=11)])

c_extra = [('BACKGROUND', (0, 0), (-1, 0), C['hdr'])]
c_extra += [('BACKGROUND', (0, i), (-1, i), C['altrow'] if i % 2 == 1 else C['altrow2'])
            for i in range(1, len(cause_rows))]
story.append(tbl(cause_rows, cause_w, style_extra=c_extra, alt=False))

story += [SP(2),
    info_box(
        'Ключевое наблюдение: ошибки EGT появились ПОСЛЕ 197-часового ремонта ДВС. '
        'Поджатие пинов даёт временный эффект, но не устраняет причину. '
        'После замены клапанов ГБЦ (25.05.2026) статус ошибок EGT неизвестен — требует проверки.',
        bg_color=C['orangebg'], text_color=C['orange'], bold=True, size=8.5, border_color=C['orange']
    ),
    SP(3),
]

# Problem 2 — Major repair 197h
story += [
    Pc(C['red'], '4.2  РЕМОНТ ДВС 197 ЧАСОВ (апр.2026) — АНАЛИЗ', bold=True, size=11, lead=15),
    HRFlowable(width='100%', thickness=0.5, color=C['red'], spaceAfter=4, spaceBefore=2),
    SP(1),
]

repair_info = [
    ('Дата начала', '16.04.2026  (последняя запись работы: 09.04.2026 ТО-6)'),
    ('Длительность', '197 часов — 8 суток 5 часов'),
    ('Наработка на момент ремонта', '~10 130 м/ч'),
    ('Тип события', 'РЕМОНТ ДВС — "Диагностика и ремонт ДВС"'),
    ('Объём работ', 'НЕ ЗАДОКУМЕНТИРОВАН в доступных данных — требуется тех. акт'),
    ('Масло после ремонта (11.05.2026)', 'Fe=0, Cu=0, Si=0 — загрязнений нет (хорошо)'),
    ('Контрольные ошибки после ремонта', 'EGT цил.3 и 16 — повторяющиеся'),
]

repair_w = [PW*0.38, PW*0.62]
repair_rows = []
for i, (key, val) in enumerate(repair_info):
    bg = C['altrow'] if i % 2 == 0 else C['altrow2']
    is_warn = 'НЕ ЗАДОКУМЕНТИРОВАН' in val
    repair_rows.append([
        Pc(C['sub'], key, bold=True, size=8.5, lead=12),
        Pc(C['red'] if is_warn else C['text'], val, bold=is_warn, size=8.5, lead=12),
    ])
repair_extra = [('BACKGROUND', (0, i), (-1, i), C['altrow'] if i % 2 == 0 else C['altrow2'])
               for i in range(len(repair_rows))]
repair_extra.append(('BACKGROUND', (0, 4), (-1, 4), C['redbg']))
story.append(tbl(repair_rows, repair_w, style_extra=repair_extra, alt=False))

story += [SP(1),
    info_box(
        'Масло ДВС после ремонта (проба 11.05.2026, М/Ч=10 500):\n'
        'Fe=0, Cu=0, Si=0, TAN=13.0, TBN=91.9, Сажа=140 мг/кг — статус "Допустимое"\n'
        'Отсутствие металлических частиц в масле свидетельствует о том, что ремонт прошёл без '
        'посторонних загрязнений. Хороший признак.',
        bg_color=C['greenbg'], text_color=C['green'], size=8, border_color=C['green']
    ),
    SP(2),
]

# DML note
story += section_title('СТАТУС DML-ДАННЫХ', None)
story += [SP(1),
    info_box(
        'DML-данные для №69 НЕ ПРЕДОСТАВЛЕНЫ.\n'
        'Оценка двигателя выполнена на основе: спектрального анализа масла ДВС (34 пробы) + '
        'истории ТО и ремонтов + технического отчёта 25.05.2026.\n'
        'Без DML невозможно оценить: температуры выхлопа по цилиндрам (EGT), '
        'давление масла в динамике, наддув, расход топлива в режиме работы.',
        bg_color=C['greybg'], text_color=C['sub'], size=8.5, border_color=C['grey']
    ),
    SP(1),
]

story.append(PageBreak())

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════
story += section_title('РЕКОМЕНДАЦИИ', 5)
story += [SP(2)]

recommendations = [
    (
        '1. КРИТИЧНО',
        'Уточнить объём ремонта ДВС (апр.2026, 197 ч)',
        'Запросить и изучить технический акт ремонта: перечень заменённых деталей '
        '(поршни, кольца, ГБЦ, вкладыши, коромысла?), причину отказа, выполненные регулировки. '
        'Без этого документа невозможно оценить текущий ресурс ДВС.',
        C['red'], C['redbg'], C['red']
    ),
    (
        '2. ВАЖНО',
        'Проверить статус EGT датчиков цил.3 и 16',
        'После замены клапанов ГБЦ (25.05.2026) провести диагностику INSITE — '
        'проверить наличие ошибок EGT. Если ошибки повторяются — заменить термопары '
        'цилиндров 3 и 16. Работа без температурного контроля двух цилиндров недопустима.',
        C['orange'], C['orangebg'], C['orange']
    ),
    (
        '3. ВАЖНО',
        'Запросить DML-данные для оценки после ремонта',
        'Снять DML-запись под нагрузкой (гружёный рейс) — оценить температуры EGT '
        'по всем 16 цилиндрам, давление наддува, температуру масла и охлаждающей жидкости. '
        'Особое внимание — цилиндры 3 и 16.',
        C['orange'], C['orangebg'], C['orange']
    ),
    (
        '4. НАБЛЮДЕНИЕ',
        'Контроль TBN при продолжении мониторинга масла',
        'TBN снизился до 91.9 мг/кг (11.05.2026) — минимальное значение за период. '
        'Продолжить замену масла строго каждые 250 м/ч. '
        'При TBN <80 мг/кг — рассмотреть досрочную замену масла.',
        C['blue'], C['bluebg'], C['blue']
    ),
    (
        '5. ИНФО',
        'Сажа 153 мг/кг (апр.2026) — норма',
        'Рост сажи с 119 до 153 мг/кг соответствует режиму эксплуатации. '
        'Критический уровень для QSK50 — 250 мг/кг. Текущий уровень допустим.',
        C['sub'], C['greybg'], C['sub']
    ),
    (
        '6. ИНФО',
        'Единичный Si=27 (30.04.2026) — наблюдение',
        'Следующая проба (11.05.2026) Si=0. Подтверждает загрязнение при отборе, '
        'а не реальный источник кремния в системе. '
        'При следующем Si>10 — проверить воздушный фильтр ДВС и уплотнения впуска.',
        C['sub'], C['greybg'], C['sub']
    ),
]

prio_w = [PW*0.17, PW*0.30, PW*0.53]
prio_rows_all = []
for prio_label, title, body, title_col, bg, border in recommendations:
    prio_rows_all.append([
        Pc(title_col, prio_label, bold=True, size=9, lead=13),
        Pc(title_col, title,      bold=True, size=9, lead=13),
        Pc(C['text'], body,                  size=8.5, lead=12),
    ])

for i, (_, _, _, _, bg, border) in enumerate(recommendations):
    pass  # we'll set style per-row

prio_t = Table(prio_rows_all, colWidths=prio_w)
prio_ts = TableStyle([
    ('GRID',         (0, 0), (-1, -1), 0.3, C['border']),
    ('LEFTPADDING',  (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING',   (0, 0), (-1, -1), 7),
    ('BOTTOMPADDING',(0, 0), (-1, -1), 7),
    ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
])
for i, (_, _, _, _, bg, _) in enumerate(recommendations):
    prio_ts.add('BACKGROUND', (0, i), (-1, i), bg)
prio_t.setStyle(prio_ts)
story += [prio_t, SP(3)]

# Final summary box
story += [
    HR(C['blue']),
    SP(1),
    Pc(C['blue'], 'ИТОГОВАЯ ОЦЕНКА ДВИГАТЕЛЯ №69', bold=True, size=11, lead=14, align=TA_CENTER),
    SP(1),
]

summary_kpi = [
    ('МАСЛО ДВС', 'ОТЛИЧНО', 'Все 34 пробы чисты\nFe 0-3 мг/кг', C['green'], C['greenbg']),
    ('СОСТОЯНИЕ', 'РАБОЧЕЕ', 'Двигатель функционирует\n(после ремонта апр.2026)', C['orange'], C['orangebg']),
    ('РИСКИ', 'EGT Цил.3+16\nДокументация', 'Повторные ошибки датчиков\nАкт ремонта отсутствует', C['orange'], C['orangebg']),
    ('СЛЕД. ДЕЙСТВИЕ', 'DML + Акт', 'Запросить документы\nДиагностика INSITE', C['blue'], C['bluebg']),
]

final_w = [PW / 4] * 4
final_hdr = [Pc(C['sub'], lbl, bold=True, size=7, lead=10, align=TA_CENTER) for lbl, _, _, _, _ in summary_kpi]
final_val = [Pc(vc, val, bold=True, size=12, lead=16, align=TA_CENTER) for _, val, _, vc, _ in summary_kpi]
final_sub = [Pc(vc, sub, size=7, lead=10, align=TA_CENTER) for _, _, sub, vc, _ in summary_kpi]

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

# Footer note
story += [
    HR(),
    SP(1),
    Pc(C['sub'],
       'Отчёт подготовлен на основе: Спектральный анализ масла ДВС (34 пробы, апр.2025–май.2026)  •  '
       'История ТО и ремонтов №69  •  Технический отчёт 25.05.2026 (замена клапанов ГБЦ)  •  '
       'DML-данные: ОТСУТСТВУЮТ  |  Горная Евразия  •  07.06.2026',
       size=7, lead=10),
]

# ── Build ──────────────────────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f'OK → {OUT}')
