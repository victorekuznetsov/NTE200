#!/usr/bin/env python3
"""
Presentation: Analysis of exhaust valve failures on QSK50 engines, NTE200 dump trucks
АО «Полюс Магадан» / Омчак / Тенькинский район
"""
import os
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches, Pt
import fitz  # PyMuPDF

WORK_DIR = "/home/user/NTE200"

# ── Colors ──────────────────────────────────────────────────────────────────
C_DARK    = RGBColor(0x1A, 0x3A, 0x5C)   # dark navy
C_MID     = RGBColor(0x2E, 0x6D, 0xA8)   # mid blue
C_ACCENT  = RGBColor(0xE8, 0x4C, 0x2C)   # red-orange (alerts)
C_GOLD    = RGBColor(0xF5, 0xA6, 0x23)   # amber (warnings)
C_GREEN   = RGBColor(0x27, 0xAE, 0x60)   # green (ok)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT   = RGBColor(0xE8, 0xF4, 0xFF)   # pale blue
C_GRAY    = RGBColor(0x55, 0x55, 0x55)
C_LGRAY   = RGBColor(0xF2, 0xF2, 0xF2)
C_BLACK   = RGBColor(0x00, 0x00, 0x00)

FONT = "DejaVu Sans"

# ── Helpers ─────────────────────────────────────────────────────────────────
def add_bg(slide, color):
    from pptx.util import Emu
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def txf(shape, bold=False, size=18, color=C_BLACK, align=PP_ALIGN.LEFT, wrap=True):
    tf = shape.text_frame
    tf.word_wrap = wrap
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.name = FONT
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color
    return tf

def add_title_box(slide, text, left, top, width, height,
                  size=28, bold=True, color=C_WHITE, bg=C_DARK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    fill = box.fill
    fill.solid()
    fill.fore_color.rgb = bg
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box

def add_text_box(slide, text, left, top, width, height,
                 size=14, bold=False, color=C_BLACK, bg=None, align=PP_ALIGN.LEFT,
                 italic=False):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if bg:
        fill = box.fill
        fill.solid()
        fill.fore_color.rgb = bg
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box

def add_multiline_box(slide, lines, left, top, width, height,
                      size=13, bold=False, color=C_BLACK, bg=None,
                      align=PP_ALIGN.LEFT, line_spacing=1.0, indent=None):
    """lines: list of (text, bold, color, size) or plain strings"""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if bg:
        fill = box.fill
        fill.solid()
        fill.fore_color.rgb = bg
    tf = box.text_frame
    tf.word_wrap = True
    from pptx.util import Pt as PT
    from pptx.oxml.ns import qn
    import lxml.etree as etree
    first = True
    for item in lines:
        if isinstance(item, str):
            txt, b, c, s = item, bold, color, size
        else:
            if len(item) == 4:
                txt, b, c, s = item
            elif len(item) == 3:
                txt, b, c = item; s = size
            else:
                txt, b = item; c = color; s = size
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if indent is not None:
            p.level = indent
        run = p.add_run()
        run.text = txt
        run.font.name = FONT
        run.font.size = PT(s)
        run.font.bold = b
        run.font.color.rgb = c
    return box

def add_rect(slide, left, top, width, height, fill_color=C_DARK, line_color=None):
    from pptx.util import Pt as PT
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    return shape

def add_table(slide, rows_data, col_widths, left, top, width, height,
              header_bg=C_DARK, header_color=C_WHITE,
              row_bg=C_WHITE, alt_bg=C_LIGHT,
              cell_size=11, header_size=12,
              header_bold=True):
    """rows_data[0] = header row; subsequent = data rows"""
    from pptx.util import Pt as PT
    from pptx.oxml.ns import qn
    nrows = len(rows_data)
    ncols = len(rows_data[0])
    tbl = slide.shapes.add_table(nrows, ncols,
                                  Inches(left), Inches(top),
                                  Inches(width), Inches(height)).table
    # Column widths
    total = sum(col_widths)
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = Inches(width * w / total)
    # Rows
    for r, row in enumerate(rows_data):
        for c, cell_text in enumerate(row):
            cell = tbl.cell(r, c)
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_bg
            elif r % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = row_bg
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = alt_bg
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(cell_text)
            run.font.name = FONT
            run.font.size = PT(header_size if r == 0 else cell_size)
            run.font.bold = header_bold if r == 0 else False
            run.font.color.rgb = header_color if r == 0 else C_BLACK
    return tbl

def extract_page_image(pdf_path, page_num=0, dpi=100):
    """Extract a page from PDF as PNG bytes"""
    try:
        doc = fitz.open(pdf_path)
        if page_num >= len(doc):
            page_num = 0
        page = doc[page_num]
        mat = fitz.Matrix(dpi/72, dpi/72)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception as e:
        print(f"  [WARN] Cannot extract image from {pdf_path}: {e}")
        return None

def add_pdf_image(slide, pdf_path, page_num=0, left=0.1, top=1.2,
                  width=4.5, height=4.5, dpi=100):
    img_bytes = extract_page_image(pdf_path, page_num, dpi)
    if img_bytes:
        img_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(img_stream, Inches(left), Inches(top),
                                  Inches(width), Inches(height))
        return True
    return False

def slide_header(slide, title, subtitle=None):
    """Standard slide header: colored bar at top"""
    add_rect(slide, 0, 0, 13.33, 1.1, fill_color=C_DARK)
    add_title_box(slide, title,
                  0.2, 0.05, 12.5, 0.9,
                  size=26, bold=True, color=C_WHITE, bg=C_DARK)
    if subtitle:
        add_text_box(slide, subtitle, 0.2, 0.9, 12.5, 0.4,
                     size=13, bold=False, color=RGBColor(0xB0, 0xC8, 0xE8),
                     align=PP_ALIGN.LEFT)
    add_rect(slide, 0, 1.1, 13.33, 0.04, fill_color=C_GOLD)


# ═══════════════════════════════════════════════════════════════════════════
#  Build Presentation
# ═══════════════════════════════════════════════════════════════════════════
def build_presentation():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]  # blank layout

    print("Creating presentation slides...")

    # ── SLIDE 1: TITLE ──────────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_DARK)
    # Gradient effect via rect
    add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0x0D, 0x1F, 0x33))
    add_rect(slide, 0, 5.8, 13.33, 1.7, fill_color=RGBColor(0x12, 0x2B, 0x45))
    add_rect(slide, 0, 5.78, 13.33, 0.06, fill_color=C_GOLD)
    # Title
    add_title_box(slide, "АНАЛИЗ ПРИЧИН ИЗНОСА ВЫПУСКНЫХ КЛАПАНОВ",
                  0.8, 0.7, 11.7, 1.2, size=34, bold=True,
                  color=C_WHITE, bg=RGBColor(0x0D, 0x1F, 0x33))
    add_title_box(slide, "Двигатели Cummins QSK50 | Самосвалы NHL NTE200",
                  0.8, 1.9, 11.7, 0.7, size=22, bold=False,
                  color=RGBColor(0xB0, 0xD4, 0xF0), bg=RGBColor(0x0D, 0x1F, 0x33))
    # Separator
    add_rect(slide, 0.8, 2.75, 11.7, 0.05, fill_color=C_GOLD)
    # Info blocks
    info = [
        ("Объект:", "Золотодобывающее предприятие АО «Полюс Магадан»"),
        ("Район:", "Омчак, Тенькинский район, Магаданская область"),
        ("Двигатель:", "QSK50 MCRS (50L, 16 цил., 1492 кВт) — серийный номер 33232776 / 33238516"),
        ("Масштаб:", "20+ единиц, 30+ событий ремонта, 2 776 — 18 505 м/ч наработки"),
        ("Период:", "Август 2025 — Май 2026"),
    ]
    y = 2.9
    for label, val in info:
        add_text_box(slide, label, 0.8, y, 2.5, 0.38,
                     size=14, bold=True, color=C_GOLD,
                     bg=RGBColor(0x0D, 0x1F, 0x33))
        add_text_box(slide, val, 3.3, y, 9.3, 0.38,
                     size=14, bold=False, color=C_WHITE,
                     bg=RGBColor(0x0D, 0x1F, 0x33))
        y += 0.5
    add_text_box(slide,
                 "Конфиденциально — только для внутреннего использования",
                 0.8, 6.95, 11.7, 0.4,
                 size=11, bold=False, color=RGBColor(0x70, 0x90, 0xB0),
                 bg=RGBColor(0x0D, 0x1F, 0x33), align=PP_ALIGN.CENTER)
    print("  Slide 1: Title done")

    # ── SLIDE 2: МАСШТАБ ПРОБЛЕМЫ ────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "МАСШТАБ ПРОБЛЕМЫ",
                 "Статистика отказов выпускных клапанов QSK50 | Парк NTE200")
    # KPI boxes
    kpis = [
        ("20+", "единиц\nзатронуто", C_ACCENT),
        ("30+", "событий\nремонта", C_MID),
        ("2 776", "м/ч минимум\nдо отказа", C_GOLD),
        ("18 505", "м/ч максимум\nдо отказа", C_GREEN),
    ]
    for i, (val, label, col) in enumerate(kpis):
        x = 0.3 + i * 3.25
        add_rect(slide, x, 1.3, 3.0, 1.8, fill_color=col)
        add_title_box(slide, val,
                      x+0.1, 1.35, 2.8, 0.9,
                      size=38, bold=True, color=C_WHITE, bg=col, align=PP_ALIGN.CENTER)
        add_text_box(slide, label,
                     x+0.1, 2.25, 2.8, 0.7,
                     size=14, bold=False, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Summary table
    summary = [
        ["Единица", "Наработка до 1-го отказа, м/ч", "Событий ремонта", "Масштаб поражения"],
        ["№43", "16 033",  "1", "5 позиций (3L,4L,1R,2R,3R)"],
        ["№44", "14 751",  "1", "ВСЕ 14 позиций"],
        ["№45", "17 863",  "1", "1 позиция (2R)"],
        ["№47", "12 890",  "2", "7 позиций + замена ГБЦ"],
        ["№48", "14 997",  "2", "17 клапанов + турбина + гидроудар"],
        ["№51", "10 783",  "1", "Замена ГБЦ (трещина, 6R)"],
        ["№53", "17 550",  "1", "ВСЕ 16 позиций (2 выпуск + 2 впуск)"],
        ["№55", "17 316",  "1", "ВСЕ 16 позиций + 4 форсунки"],
        ["№57", "14 808",  "1", "6 позиций"],
        ["№58", "10 702",  "2", "Клапан + замена толкателя"],
        ["№59", "15 143",  "2", "Многократные отказы"],
        ["№62", "13 917",  "1", "Клапаны"],
        ["№69", "10 130",  "2", "Засорение возд. фильтров + поршни"],
        ["№72", "10 000",  "2", "8 позиций (2 эпизода)"],
        ["№73", "9 998",   "1", "Клапаны + форсунки"],
        ["№74", "9 869",   "1", "Клапаны"],
        ["№76", "9 754",   "2", "Клапаны (2 эпизода)"],
        ["№77", "10 016",  "1", "Клапаны"],
        ["№78", "9 323",   "1", "3R/4R/6R/6L"],
        ["№81", "4 696",   "1", "ВСЕ позиции + замена ГБЦ"],
        ["№83", "2 776",   "1", "КАТАСТРОФИЧЕСКИЙ (замена шатуна/поршня/ГБЦ)"],
    ]
    add_table(slide, summary, [0.8, 2.0, 1.4, 3.8], 0.15, 3.2, 13.0, 4.0,
              header_bg=C_DARK, cell_size=9, header_size=10)
    print("  Slide 2: Scale done")

    # ── SLIDE 3: ХРОНОЛОГИЯ ОТКАЗОВ ──────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "ХРОНОЛОГИЯ И ДИНАМИКА ОТКАЗОВ",
                 "Временная шкала ремонтных событий | Август 2025 — Май 2026")

    timeline = [
        ("Авг 2025",  "№47", "Замена ГБЦ (ранний эпизод) | 12 890 м/ч"),
        ("Авг 2025",  "№51", "Замена ГБЦ, трещина | 10 783 м/ч"),
        ("Авг 2025",  "№58", "Первый эпизод — замена толкателя | 10 702 м/ч"),
        ("Янв 2026",  "№43", "Замена выпускных клапанов + форсунки | 16 033 м/ч"),
        ("Фев 2026",  "№48", "17 клапанов + турбина | 14 997 м/ч"),
        ("Мар 2026",  "№58", "Второй эпизод — клапан + ГБЦ | 15 572 м/ч"),
        ("Мар 2026",  "№83", "КАТАСТРОФА — шатун/поршень/ГБЦ | 2 776 м/ч"),
        ("Мар 2026",  "№47", "Второй эпизод — 10 клапанов + форсунки | 12 890 м/ч"),
        ("Апр 2026",  "№48", "Гидроудар, водомасляная эмульсия | после ремонта"),
        ("Май 2026",  "№55", "ВСЕ 16 позиций + 4 форсунки | 17 316 м/ч"),
        ("Май 2026",  "№69", "4L/5L/6L/7L/7R + засорение фильтров | 10 721 м/ч"),
        ("Май 2026",  "№78", "3R/4R/6R/6L | 9 323 м/ч"),
        ("Май 2026",  "№72", "Инжектор (форсунка) | 10 283 м/ч"),
        ("Май 2026",  "№73", "Форсунки (тест L2+R1 аномалии) | ~9 998 м/ч"),
    ]
    row_h = 0.38
    y0 = 1.3
    add_rect(slide, 0.15, y0, 13.0, 0.35, fill_color=C_DARK)
    for col, label in zip([0.15, 1.8, 3.2, 6.5],
                          ["Дата", "Ед.", "Двигатель / событие", "Наработка"]):
        add_text_box(slide, label, col+0.1, y0+0.02, 1.5, 0.32,
                     size=11, bold=True, color=C_WHITE)

    for i, (dt, unit, event) in enumerate(timeline):
        y = y0 + 0.35 + i * row_h
        bg = C_LGRAY if i % 2 == 0 else C_WHITE
        add_rect(slide, 0.15, y, 13.0, row_h, fill_color=bg)
        # Highlight catastrophic
        if "КАТАСТРОФ" in event:
            add_rect(slide, 0.15, y, 13.0, row_h, fill_color=RGBColor(0xFF, 0xE8, 0xE0))
        add_text_box(slide, dt, 0.25, y+0.03, 1.5, row_h-0.05,
                     size=10, bold=False, color=C_GRAY)
        col_c = C_ACCENT if "83" in unit else C_MID
        add_text_box(slide, unit, 1.9, y+0.03, 1.3, row_h-0.05,
                     size=10, bold=True, color=col_c)
        add_text_box(slide, event, 3.3, y+0.03, 9.7, row_h-0.05,
                     size=10, bold=False, color=C_BLACK)
    print("  Slide 3: Timeline done")

    # ── SLIDE 4: ФИЗИЧЕСКИЕ ПРИЗНАКИ ─────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "ФИЗИЧЕСКИЕ ПРИЗНАКИ ПОВРЕЖДЕНИЙ",
                 "Данные разборки и дефектовки | Сводка по всем агрегатам")
    findings = [
        ("ПРОСАДКА ФАСКИ КЛАПАНА (РЕЦЕССИЯ)",
         C_ACCENT, "Главный признак. Посадочная поверхность выпускного клапана «утопает» в седле.\n"
         "Твёрдость фаски 40-44 HRC (базовая) против 55-65 HRC (Stellite). "
         "Обнаружена на всех исследованных двигателях."),
        ("ОТЛОЖЕНИЯ ЗОЛЫ НА ТАРЕЛКЕ",
         C_GOLD, "EDS-анализ NHL (MS&T2026033): Ca, Zn, P — зола масляных присадок ZDDP.\n"
         "Si, Al — силикатная пыль через воздушный тракт. "
         "Роль: ОТЯГЧАЮЩИЙ фактор (абразив), не первопричина."),
        ("ПЛАСТИЧЕСКАЯ ДЕФОРМАЦИЯ («ЗАУСЕНЕЦ»)",
         C_ACCENT, "Бурт пластической деформации в зоне фаски = признак перегрева тарелки.\n"
         "Температура клапана вышла за допустимый предел. "
         "Зафиксировано в лабораторном заключении NHL."),
        ("ЧАСТИЧНОЕ РАЗРУШЕНИЕ / ПРОБОЙ ГБЦ",
         C_MID, "Агрегаты №47, №48, №53, №55, №81: замена ГБЦ целиком.\n"
         "№83 (2 776 м/ч): разрушение шатуна и поршня — катастрофический отказ.\n"
         "№48: гидроудар (водомасляная эмульсия) после первого ремонта."),
        ("КОРРЕЛЯЦИЯ С ЦИЛИНДРАМИ",
         C_GREEN, "Наиболее часто поражены: 3R, 4R, 6R, 6L, 7L, 7R.\n"
         "Потенциально связано с охлаждением (крайние ряды, расположение термостатов)."),
        ("ФОРСУНКИ",
         C_MID, "Форсунки заменены вместе с клапанами: №43, №47, №55, №73.\n"
         "Протечка форсунки = дополнительное тепловыделение → перегрев клапана."),
    ]
    y = 1.25
    for title, col, text in findings:
        add_rect(slide, 0.15, y, 0.08, 0.7, fill_color=col)
        add_text_box(slide, title, 0.35, y, 4.2, 0.28,
                     size=12, bold=True, color=col)
        add_text_box(slide, text, 0.35, y+0.28, 12.8, 0.55,
                     size=10, bold=False, color=C_GRAY)
        y += 0.95
    print("  Slide 4: Physical findings done")

    # ── SLIDE 5: АНАЛИЗ NHL LABORATORY ───────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "АНАЛИЗ ЗАВОДСКОЙ ЛАБОРАТОРИИ NHL",
                 "Отчёт MS&T2026033 | Инженер Тао Лан | Декабрь 2025")
    # Image from NHL lab PDF
    nhl_pdf = os.path.join(WORK_DIR, "气门盘部磨损分析报告-1.docx RUS.pdf")
    img_added = add_pdf_image(slide, nhl_pdf, page_num=2,
                               left=0.2, top=1.3, width=5.0, height=4.5)
    img2_added = add_pdf_image(slide, nhl_pdf, page_num=3,
                                left=5.2, top=1.3, width=5.0, height=4.5)
    # Right side findings
    x = 10.3
    findings_nhl = [
        ("МЕТОДЫ АНАЛИЗА:", C_DARK, True),
        ("• Сканирующая электронная микроскопия (СЭМ)", C_BLACK, False),
        ("• EDS-спектроскопия (элементный анализ)", C_BLACK, False),
        ("• Измерение твёрдости фаски", C_BLACK, False),
        ("• Профилометрия поверхности", C_BLACK, False),
        ("", C_BLACK, False),
        ("КЛЮЧЕВЫЕ РЕЗУЛЬТАТЫ:", C_DARK, True),
        ("Ca / Zn / P = зола ZDDP присадок", C_ACCENT, False),
        ("Si / Al = силикатная пыль (воздух)", C_ACCENT, False),
        ("Пластическая деформация фаски", C_ACCENT, False),
        ("→ превышение температуры", C_GRAY, False),
        ("", C_BLACK, False),
        ("ВЫВОД NHL:", C_DARK, True),
        ("Комплексная причина:", C_BLACK, False),
        ("1. Термическое повреждение", C_BLACK, False),
        ("2. Абразивный износ (зола+пыль)", C_BLACK, False),
        ("3. Адгезионный износ", C_BLACK, False),
    ]
    y = 1.35
    for txt, col, bold in findings_nhl:
        add_text_box(slide, txt, x, y, 2.9, 0.3,
                     size=10 if not bold else 11, bold=bold, color=col)
        y += 0.3
    print("  Slide 5: NHL lab analysis done")

    # ── SLIDE 6: EDS ANALYSIS DETAIL ─────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "EDS-АНАЛИЗ ОТЛОЖЕНИЙ | ЛАБОРАТОРИЯ NHL",
                 "Элементный состав отложений на фаске клапана — MS&T2026033")
    # Images from pages 6-7 of NHL report
    add_pdf_image(slide, nhl_pdf, page_num=5, left=0.2, top=1.3, width=6.3, height=4.2)
    add_pdf_image(slide, nhl_pdf, page_num=6, left=6.7, top=1.3, width=6.2, height=4.2)
    # Bottom text
    add_rect(slide, 0.15, 5.6, 13.0, 1.65, fill_color=RGBColor(0xFF, 0xF8, 0xE8))
    add_text_box(slide,
                 "ИНТЕРПРЕТАЦИЯ: Ca (кальций) + Zn (цинк) + P (фосфор) = зола масляной присадки ZDDP "
                 "(цинкдиалкилдитиофосфат). Si (кремний) + Al (алюминий) = силикатная пыль, "
                 "попавшая через воздушный тракт. Отложения создают абразивную среду в зоне контакта "
                 "клапан-седло, ускоряя износ. Однако зола является ОТЯГЧАЮЩИМ фактором, а не первопричиной "
                 "повреждений.",
                 0.3, 5.65, 12.8, 1.5,
                 size=11, color=RGBColor(0x44, 0x33, 0x00))
    print("  Slide 6: EDS analysis done")

    # ── SLIDE 7: UNIT 83 — CATASTROPHIC FAILURE ──────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    add_rect(slide, 0, 0, 13.33, 1.1, fill_color=RGBColor(0x7B, 0x1A, 0x0A))
    add_title_box(slide, "КАТАСТРОФИЧЕСКИЙ ОТКАЗ | АГРЕГАТ №83",
                  0.2, 0.05, 12.5, 0.9, size=26, bold=True,
                  color=C_WHITE, bg=RGBColor(0x7B, 0x1A, 0x0A))
    add_rect(slide, 0, 1.1, 13.33, 0.04, fill_color=C_GOLD)
    # Photos
    pdf83 = os.path.join(WORK_DIR, "Тех отчет 83 от 30.03.2026.pdf")
    add_pdf_image(slide, pdf83, page_num=0, left=0.2, top=1.25, width=4.2, height=3.2)
    add_pdf_image(slide, pdf83, page_num=2, left=4.5, top=1.25, width=4.2, height=3.2)
    add_pdf_image(slide, pdf83, page_num=4, left=8.9, top=1.25, width=4.2, height=3.2)
    # Key data
    data83 = [
        ("Гаражный номер:", "№83"),
        ("Наработка до отказа:", "2 776 м/ч (!!! — норм. ресурс >12 000 м/ч)"),
        ("Серийный номер:", "двигатель QSK50"),
        ("Зафиксировано:", "Разрушение шатуна и поршня"),
        ("Замена:", "ГБЦ + шатун + поршень + уплотнения"),
        ("Значимость:", "Самый ранний отказ в парке — катастрофический"),
        ("Вероятная причина:", "Перегрев → потеря плотности → гидроудар/поломка"),
    ]
    add_rect(slide, 0.15, 4.6, 13.0, 2.7, fill_color=RGBColor(0xFF, 0xF0, 0xED))
    y = 4.65
    for lbl, val in data83:
        add_text_box(slide, lbl, 0.3, y, 3.2, 0.33,
                     size=11, bold=True, color=RGBColor(0x7B, 0x1A, 0x0A))
        add_text_box(slide, val, 3.5, y, 9.7, 0.33,
                     size=11, bold=False, color=C_BLACK)
        y += 0.36
    print("  Slide 7: Unit 83 catastrophic failure done")

    # ── SLIDE 8: UNIT 48 DETAIL ───────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "АГРЕГАТ №48 — ДВОЙНОЙ ОТКАЗ",
                 "14 997 м/ч: 17 клапанов + турбина | Повторный отказ: гидроудар")
    pdf48 = os.path.join(WORK_DIR, "Тех.отчет 48.pdf")
    pdf48b = os.path.join(WORK_DIR, "Тех оточет 48 от 31.03.2026.pdf")
    add_pdf_image(slide, pdf48, page_num=0, left=0.2, top=1.3, width=4.0, height=3.0)
    add_pdf_image(slide, pdf48, page_num=2, left=4.3, top=1.3, width=4.0, height=3.0)
    add_pdf_image(slide, pdf48b, page_num=1, left=8.6, top=1.3, width=4.5, height=3.0)
    # Two-column summary
    add_rect(slide, 0.15, 4.45, 6.4, 2.8, fill_color=RGBColor(0xE8, 0xF4, 0xFF))
    add_text_box(slide, "ПЕРВЫЙ РЕМОНТ (Фев 2026)", 0.3, 4.5, 6.0, 0.35,
                 size=12, bold=True, color=C_DARK)
    ep1 = [
        "Наработка: 14 997 м/ч",
        "17 выпускных клапанов (почти все позиции)",
        "Замена турбокомпрессора",
        "Замена ГБЦ — позиция 3R",
        "Визуально: нагар, просадка, абразивный износ",
    ]
    y = 4.9
    for s in ep1:
        add_text_box(slide, "• " + s, 0.4, y, 5.9, 0.32, size=11, color=C_BLACK)
        y += 0.34

    add_rect(slide, 6.75, 4.45, 6.4, 2.8, fill_color=RGBColor(0xFF, 0xF0, 0xED))
    add_text_box(slide, "ВТОРОЙ ОТКАЗ (Мар 2026 — после ремонта)", 6.9, 4.5, 6.2, 0.35,
                 size=12, bold=True, color=RGBColor(0x7B, 0x1A, 0x0A))
    ep2 = [
        "Обнаружена водомасляная эмульсия",
        "Гидроудар — потенциальное разрушение ЦПГ",
        "Источник воды: нарушение уплотнения ГБЦ",
        "После клапанного ремонта — через ~3-4 недели",
        "Повторная разборка и дефектовка ГБЦ",
    ]
    y = 4.9
    for s in ep2:
        add_text_box(slide, "• " + s, 7.0, y, 6.0, 0.32, size=11, color=C_BLACK)
        y += 0.34
    print("  Slide 8: Unit 48 done")

    # ── SLIDE 9: UNIT 55 & 69 DETAIL ─────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "АГРЕГАТЫ №55 И №69 — ПОКАЗАТЕЛЬНЫЕ СЛУЧАИ",
                 "№55: ВСЕ 16 позиций | №69: Засорение воздушных фильтров")
    # Unit 55
    add_rect(slide, 0.15, 1.25, 6.5, 6.1, fill_color=C_LGRAY)
    add_text_box(slide, "АГРЕГАТ №55 (17 316 м/ч)",
                 0.25, 1.3, 6.2, 0.45, size=14, bold=True, color=C_DARK)
    doc55_pdf = os.path.join(WORK_DIR, "Тех отчет 55.pdf")
    img_ok = add_pdf_image(slide, doc55_pdf, page_num=0,
                            left=0.25, top=1.8, width=4.0, height=2.5)
    u55_data = [
        "Наработка: 17 316 м/ч",
        "Заменены ВСЕ 16 позиций клапанов",
        "Заменены 4 форсунки (1L, 2L, 3L, 6L)",
        "Все клапаны: признаки перегрева",
        "Самая большая наработка при полном поражении",
        "→ Длительная работа с отклонениями",
    ]
    y = 1.8
    for s in u55_data:
        add_text_box(slide, "• " + s, 4.4, y, 2.0, 0.38, size=10, color=C_BLACK)
        y += 0.38
    # Unit 69
    add_rect(slide, 6.75, 1.25, 6.45, 6.1, fill_color=RGBColor(0xFF, 0xF8, 0xE8))
    add_text_box(slide, "АГРЕГАТ №69 (10 721 м/ч)",
                 6.85, 1.3, 6.2, 0.45, size=14, bold=True, color=C_GOLD)
    doc69 = os.path.join(WORK_DIR, "69 ТЕХНИЧЕСКИЙ ОТЧЁТ NTE 200 клапанов 25.05.2026.docx")
    # Try to get a photo from unit 69 PDF if exists
    pdf69 = os.path.join(WORK_DIR, "Тех.отчет 48.pdf")  # fallback
    add_pdf_image(slide, pdf69, page_num=3, left=6.85, top=1.8, width=3.5, height=2.5)
    u69_data = [
        "Наработка: 10 721 м/ч",
        "Поражены: 4L, 5L, 6L, 7L, 7R",
        "ЗАСОРЕНЫ ВОЗДУШНЫЕ ФИЛЬТРЫ (!)",
        "Это единственный агрегат с задокументированным",
        " засорением фильтров в момент отказа",
        "Замена поршней 3R и 4R",
        "Второй эпизод (10 130 м/ч) — те же цилиндры",
    ]
    y = 1.8
    for s in u69_data:
        col = C_ACCENT if "ЗАСОРЕННЫЕ" in s or "ЗАСОРЕН" in s else C_BLACK
        bold = "ЗАСОРЕН" in s
        add_text_box(slide, "• " + s if not s.startswith(" ") else s,
                     10.45, y, 2.6, 0.38, size=10, bold=bold, color=col)
        y += 0.38
    # Note
    add_rect(slide, 0.15, 5.5, 13.0, 1.8, fill_color=RGBColor(0xFF, 0xF0, 0xED))
    add_text_box(slide, "ВАЖНО: Засорение воздушных фильтров → снижение массового расхода воздуха → "
                 "обогащение смеси → повышение температуры выпускных газов → ускоренный износ клапанов. "
                 "Однако это отмечено только у №69. У других агрегатов данных о состоянии фильтров нет.",
                 0.3, 5.55, 12.8, 1.5,
                 size=11, color=RGBColor(0x44, 0x1A, 0x00))
    print("  Slide 9: Units 55 & 69 done")

    # ── SLIDE 10: ОБЗОР ГИПОТЕЗ ────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "ОБЗОР ГИПОТЕЗ О ПРИЧИНАХ ОТКАЗОВ",
                 "8 версий — от конструктивных особенностей до условий эксплуатации")
    hypotheses = [
        ("H1", "Настройки ECM двигателя",
         "Отключённая защита по дератингу, задержки 51/43 с, обороты 2100 vs 1990, топлив. коррекция 200%",
         C_ACCENT, "ВЫСОКАЯ"),
        ("H2", "Отсутствие наплавки Stellite",
         "NHL чертёж: без Stellite (40-44 HRC). Komatsu 730E AC: со Stellite (55-65 HRC). Разница твёрдости >50%",
         C_ACCENT, "ВЫСОКАЯ"),
        ("H3", "Ошибка регулировки зазоров клапанов",
         "Малый тепловой зазор → клапан не закрывается полностью → потеря охлаждения → перегрев",
         C_GOLD, "СРЕДНЯЯ"),
        ("H4", "Неисправность форсунок",
         "Протечка/нарушение распыла → дожигание в камере → перегрев клапана. №43, №47, №55, №73",
         C_GOLD, "СРЕДНЯЯ"),
        ("H5", "Засорение воздушного фильтра",
         "Богатая смесь → высокая температура ОГ. Задокументировано только для №69.",
         C_MID, "НИЗКАЯ"),
        ("H6", "Зола масла и абразивная пыль",
         "EDS NHL: Ca/Zn/P + Si/Al. Отягчающий фактор — не первопричина.",
         C_MID, "НИЗКАЯ (ОТЯГЧАЮЩИЙ)"),
        ("H7", "Рабочий цикл / условия нагружения",
         "Тяжёлые условия (горный рудник, высота, жара/мороз). Необходим сравнительный анализ.",
         C_GRAY, "НЕОПРЕДЕЛЕНА"),
        ("H8", "Конструктив охлаждения ГБЦ",
         "Неравномерное охлаждение по цилиндрам (3R, 4R, 6R, 6L чаще поражаются).",
         C_GRAY, "ТРЕБУЕТ ПРОВЕРКИ"),
    ]
    y = 1.3
    for hnum, htitle, hdesc, col, level in hypotheses:
        add_rect(slide, 0.15, y, 1.0, 0.7, fill_color=col)
        add_title_box(slide, hnum, 0.15, y, 1.0, 0.7,
                      size=18, bold=True, color=C_WHITE, bg=col, align=PP_ALIGN.CENTER)
        add_text_box(slide, htitle, 1.3, y+0.01, 4.8, 0.32,
                     size=12, bold=True, color=C_DARK)
        add_text_box(slide, hdesc, 1.3, y+0.32, 8.5, 0.38,
                     size=10, bold=False, color=C_GRAY)
        lev_col = C_ACCENT if level == "ВЫСОКАЯ" else (C_GOLD if level == "СРЕДНЯЯ" else C_MID)
        add_rect(slide, 9.9, y+0.1, 3.3, 0.5, fill_color=lev_col)
        add_text_box(slide, level, 9.95, y+0.15, 3.2, 0.4,
                     size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        y += 0.77
    print("  Slide 10: Hypotheses overview done")

    # ── SLIDE 11: H1 — ECM SETTINGS ───────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H1: НАСТРОЙКИ ECM — СРАВНЕНИЕ NTE200 vs 730E AC",
                 "Один двигатель QSK50 — один объект — разные параметры защиты")
    add_rect(slide, 0.15, 1.25, 13.0, 0.45, fill_color=RGBColor(0xFF, 0xF0, 0xED))
    add_text_box(slide,
                 "КЛЮЧЕВОЙ ФАКТ: Komatsu 730E AC с аналогичными двигателями QSK50 работает на ЭТОМ ЖЕ "
                 "объекте. Случаев замены клапанов на 730E AC НЕ ЗАФИКСИРОВАНО.",
                 0.25, 1.28, 12.8, 0.38, size=12, bold=True, color=C_ACCENT)
    # ECM comparison table
    ecm_table = [
        ["Параметр ECM", "Komatsu 730E AC", "NHL NTE200", "Риск"],
        ["C_EPD_OT_RPM_Drt_En\n(Флаг дератинга)", "ВКЛЮЧЁН (1)", "ОТКЛЮЧЁН (0)", "КРИТИЧЕСКИЙ"],
        ["Задержка дератинга по давл. масла", "Немедленно", "51 секунда (!)", "КРИТИЧЕСКИЙ"],
        ["Задержка дератинга по темп. ОЖ", "Немедленно", "43 секунды (!)", "ВЫСОКИЙ"],
        ["Максимальные обороты", "1 990 об/мин", "2 100 об/мин", "ВЫСОКИЙ"],
        ["Топливная коррекция TIB", "100%", "200% (!)", "ВЫСОКИЙ"],
        ["Триггеры по температуре ОГ", "Активны", "ВСЕ ОТКЛЮЧЕНЫ", "КРИТИЧЕСКИЙ"],
        ["Дератинг при высокой темп.", "Активен", "Не работает", "КРИТИЧЕСКИЙ"],
        ["Обороты ХХ", "700 об/мин", "700 об/мин", "Норма"],
    ]
    add_table(slide, ecm_table,
              [4.5, 2.8, 2.8, 2.5], 0.15, 1.8, 13.0, 5.4,
              header_bg=C_DARK, cell_size=10, header_size=11,
              row_bg=C_WHITE, alt_bg=C_LIGHT)
    # Color critical cells manually via workaround: add colored text boxes over
    # (We can't easily color individual cells post-creation in pptx, so we note it in the header)
    add_rect(slide, 0.15, 7.1, 13.0, 0.32, fill_color=RGBColor(0xFF, 0xE8, 0xE0))
    add_text_box(slide, "«КРИТИЧЕСКИЙ» = потенциально ведёт к перегреву двигателя без срабатывания защиты",
                 0.25, 7.12, 12.8, 0.28, size=11, bold=True, color=C_ACCENT)
    print("  Slide 11: ECM comparison done")

    # ── SLIDE 12: H1 — ECM МЕХАНИЗМ ───────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H1: КАК НАСТРОЙКИ ECM ВЕДУТ К ОТКАЗУ КЛАПАНОВ",
                 "Цепочка причинно-следственных связей")
    # Flowchart via boxes and arrows
    steps = [
        ("1. ECM: Дератинг ОТКЛЮЧЁН\n+ задержки 43-51 сек\n+ обороты 2100\n+ TIB коррекция 200%",
         C_ACCENT, 0.3, 1.5, 3.5, 1.5),
        ("2. Двигатель работает\nс ПЕРЕГРУЗКОЙ без\nснижения мощности\nпри перегреве",
         C_GOLD, 4.1, 1.5, 3.3, 1.5),
        ("3. Температура\nвыпускных газов\nвыше нормы\nбез сигнала оператору",
         C_GOLD, 7.6, 1.5, 3.3, 1.5),
        ("4. ПЕРЕГРЕВ\nвыпускного клапана\n→ пластическая\nдеформация фаски",
         C_ACCENT, 11.1, 1.5, 2.1, 1.5),
    ]
    for txt, col, x, y, w, h in steps:
        add_rect(slide, x, y, w, h, fill_color=col)
        add_title_box(slide, txt, x+0.05, y+0.05, w-0.1, h-0.1,
                      size=12, bold=False, color=C_WHITE, bg=col, align=PP_ALIGN.CENTER)
    # Arrow-like connectors (just rectangles)
    for ax in [3.8, 7.35, 10.85]:
        add_rect(slide, ax, 2.1, 0.3, 0.3, fill_color=C_DARK)
    # Step 5 - consequence
    add_rect(slide, 0.3, 3.3, 12.9, 1.4, fill_color=RGBColor(0xFF, 0xE8, 0xE0))
    add_title_box(slide,
                  "5. РЕЗУЛЬТАТ: Просадка (рецессия) фаски → потеря компрессии → "
                  "прорыв газов → ускоренный износ → выход из строя целого ряда/блока клапанов",
                  0.4, 3.35, 12.7, 1.3,
                  size=13, bold=True, color=C_ACCENT, bg=RGBColor(0xFF, 0xE8, 0xE0))
    # Evidence box
    add_rect(slide, 0.3, 4.9, 12.9, 2.4, fill_color=C_LGRAY)
    add_text_box(slide, "КОСВЕННЫЕ ДОКАЗАТЕЛЬСТВА:", 0.4, 4.95, 12.5, 0.35,
                 size=13, bold=True, color=C_DARK)
    evidence = [
        "• NTE200 инструктируются операторам: «не глушить двигатель, давать остывать на ХХ» — "
        "что указывает на известные перегревы",
        "• В INSITE CSV (агрегат N43): зафиксированы отклонения по температуре во время работы",
        "• 730E AC на том же объекте, с теми же двигателями — клапанных отказов НЕТ",
        "• Самые ранние отказы (№83 за 2776 м/ч, №81 за 4696 м/ч) не объяснимы обычным износом",
    ]
    y = 5.35
    for e in evidence:
        add_text_box(slide, e, 0.4, y, 12.7, 0.38, size=11, color=C_BLACK)
        y += 0.42
    print("  Slide 12: ECM mechanism done")

    # ── SLIDE 13: H2 — STELLITE ───────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H2: ОТСУТСТВИЕ НАПЛАВКИ STELLITE НА КЛАПАНАХ NTE200",
                 "Конструктивное различие: твёрдость фаски 40-44 HRC vs 55-65 HRC")
    # Two-column comparison
    add_rect(slide, 0.15, 1.25, 6.5, 5.5, fill_color=C_LGRAY)
    add_text_box(slide, "NTE200 | QSK50 (CES51005)",
                 0.25, 1.3, 6.2, 0.4, size=14, bold=True, color=C_ACCENT)
    nte_data = [
        ("Материал клапана:", "Inconel 751"),
        ("Наплавка фаски:", "ОТСУТСТВУЕТ (по чертежу NHL)"),
        ("Твёрдость фаски:", "40-44 HRC"),
        ("Защита от износа:", "только базовый материал"),
        ("Состояние при дефектовке:", "Рецессия на всех единицах"),
        ("Ресурс клапана:", "10 000 — 18 000 м/ч"),
    ]
    y = 1.8
    for lbl, val in nte_data:
        add_text_box(slide, lbl, 0.3, y, 2.5, 0.35, size=11, bold=True, color=C_DARK)
        c = C_ACCENT if "ОТСУТСТВУЕТ" in val else C_BLACK
        add_text_box(slide, val, 2.85, y, 3.7, 0.35, size=11, bold=False, color=c)
        y += 0.5

    add_rect(slide, 6.75, 1.25, 6.45, 5.5, fill_color=RGBColor(0xE8, 0xFF, 0xEC))
    add_text_box(slide, "Komatsu 730E AC | QSK50",
                 6.85, 1.3, 6.2, 0.4, size=14, bold=True, color=C_GREEN)
    kat_data = [
        ("Материал клапана:", "Inconel 751 или аналог"),
        ("Наплавка фаски:", "STELLITE (Co-Cr сплав)"),
        ("Твёрдость фаски:", "55-65 HRC"),
        ("Защита от износа:", "твёрдая наплавка + термостойкость"),
        ("Состояние при дефектовке:", "нет данных об отказах"),
        ("Ресурс клапана:", ">18 000 м/ч (нет отказов)"),
    ]
    y = 1.8
    for lbl, val in kat_data:
        add_text_box(slide, lbl, 6.9, y, 2.5, 0.35, size=11, bold=True, color=C_DARK)
        c = C_GREEN if "STELLITE" in val or "нет отказов" in val.lower() else C_BLACK
        add_text_box(slide, val, 9.45, y, 3.6, 0.35, size=11, bold=False, color=c)
        y += 0.5
    # Bottom note
    add_rect(slide, 0.15, 6.8, 13.0, 0.55, fill_color=RGBColor(0xFF, 0xE8, 0xE0))
    add_text_box(slide,
                 "ВЫВОД H2: Stellite снижает износ в 3-5 раз при высоких температурах (Co-Cr сплав сохраняет "
                 "твёрдость до 800°C). Применение Stellite клапанов в NTE200 — конкретная рекомендация по улучшению.",
                 0.25, 6.83, 12.8, 0.48, size=12, bold=True, color=RGBColor(0x7B, 0x1A, 0x0A))
    print("  Slide 13: H2 Stellite done")

    # ── SLIDE 14: H3 — VALVE CLEARANCE ────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H3: ОШИБКА РЕГУЛИРОВКИ ЗАЗОРОВ КЛАПАНОВ",
                 "Гипотеза из анализа повторных отказов | Агрегаты №59 и №62")
    # Left: mechanism
    add_rect(slide, 0.15, 1.25, 6.5, 5.5, fill_color=C_LGRAY)
    add_text_box(slide, "МЕХАНИЗМ НАРУШЕНИЯ:", 0.25, 1.3, 6.2, 0.4,
                 size=14, bold=True, color=C_DARK)
    mechanism = [
        ("Тепловой зазор клапана (QSK50):", ""),
        ("Впускной: 0.36 мм ± 0.05 мм", ""),
        ("Выпускной: 0.69 мм ± 0.05 мм", ""),
        ("", ""),
        ("ЕСЛИ зазор < нормы:", ""),
        ("→ Клапан не закрывается полностью", C_ACCENT),
        ("→ Площадь контакта с седлом ↓", C_ACCENT),
        ("→ Теплоотвод от тарелки ↓", C_ACCENT),
        ("→ Локальный перегрев фаски", C_ACCENT),
        ("→ Мягкий металл «прилипает» к золе", C_ACCENT),
        ("→ Ускоренная рецессия", C_ACCENT),
        ("", ""),
        ("Если зазор > нормы:", ""),
        ("→ Удар при посадке → механический износ", C_GOLD),
    ]
    y = 1.75
    for txt, col in mechanism if isinstance(mechanism[0], tuple) else [(m, C_BLACK) for m in mechanism]:
        c = col if col else C_BLACK
        b = "ЕСЛИ" in txt or txt.startswith("→")
        add_text_box(slide, txt, 0.3, y, 6.1, 0.33, size=11, bold=b, color=c)
        y += 0.34
    # Right: unit 59 case
    add_rect(slide, 6.75, 1.25, 6.45, 5.5, fill_color=RGBColor(0xFF, 0xF8, 0xE8))
    add_text_box(slide, "АГРЕГАТ №59 — МНОГОКРАТНЫЕ ОТКАЗЫ:", 6.85, 1.3, 6.2, 0.4,
                 size=14, bold=True, color=C_GOLD)
    case59 = [
        "Первый отказ: 15 143 м/ч",
        "Второй отказ: 16 219 м/ч (быстро после ремонта)",
        "Ремонт № 2 прошёл без измерения зазоров?",
        "",
        "Тех. отчет NTE (ед. 59, 62) ставит вопрос:",
        "«Не является ли систематическая ошибка",
        " регулировки первопричиной отказов?»",
        "",
        "Также зафиксированы быстрые повторные",
        "отказы у: №47, №48, №69, №72, №76",
        "",
        "ТРЕБУЕТ ПРОВЕРКИ:",
        "• Документирование зазоров ДО и ПОСЛЕ ремонта",
        "• Сравнение с нормативами Cummins",
        "• Анализ квалификации механиков",
    ]
    y = 1.75
    for s in case59:
        c = C_GOLD if s.startswith("«") else (C_ACCENT if "ТРЕБУЕТ" in s else C_BLACK)
        b = "ТРЕБУЕТ" in s or s.startswith("Тех.")
        add_text_box(slide, s, 6.9, y, 6.1, 0.33, size=11, bold=b, color=c)
        y += 0.34
    # Bottom note
    add_rect(slide, 0.15, 6.85, 13.0, 0.55, fill_color=RGBColor(0xFF, 0xF8, 0xE0))
    add_text_box(slide,
                 "СТАТУС H3: Промежуточная приоритетность. Быстрые повторные отказы могут объясняться "
                 "неправильной регулировкой при ремонте. Требует инструментальной проверки.",
                 0.25, 6.88, 12.8, 0.48, size=12, bold=False, color=RGBColor(0x44, 0x33, 0x00))
    print("  Slide 14: H3 valve clearance done")

    # ── SLIDE 15: H4+H5 ──────────────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H4: ФОРСУНКИ | H5: ВОЗДУШНЫЕ ФИЛЬТРЫ",
                 "Дополнительные факторы теплового повреждения клапанов")
    # H4 top half
    add_rect(slide, 0.15, 1.25, 13.0, 2.75, fill_color=RGBColor(0xE8, 0xF4, 0xFF))
    add_text_box(slide, "H4: НЕИСПРАВНОСТЬ ФОРСУНОК", 0.25, 1.3, 12.5, 0.4,
                 size=14, bold=True, color=C_MID)
    add_rect(slide, 0.15, 1.7, 0.06, 2.3, fill_color=C_MID)
    h4_data = [
        "Протечка или нарушение распылителя форсунки → неполное сгорание → «дожигание» у клапана",
        "Заменены совместно с клапанами: №43 (1L,2L,3L,4L,5L,6L,1R,2R,3R — 9 форсунок!), "
        "№47, №55 (1L,2L,3L,6L), №73",
        "Агрегат №85 (гар №12): 8 002 м/ч — отказ форсунки (отдельный отчёт)",
        "Агрегат №72 (гар №924): 3 700 м/ч — отказ форсунки цил. 5",
        "Тест L2/R1 для №73: аномальные показатели возврата топлива и давления",
        "СТАТУС: Форсунки — вероятный СОПУТСТВУЮЩИЙ фактор, ускоряющий износ клапанов",
    ]
    y = 1.75
    for s in h4_data:
        c = C_ACCENT if "9 форсунок" in s or "СТАТУС" in s else C_BLACK
        b = "СТАТУС" in s
        add_text_box(slide, "• " + s, 0.3, y, 12.8, 0.38, size=11, bold=b, color=c)
        y += 0.42
    # H5 bottom half
    add_rect(slide, 0.15, 4.1, 13.0, 3.15, fill_color=RGBColor(0xFF, 0xF8, 0xE8))
    add_text_box(slide, "H5: ЗАСОРЕНИЕ ВОЗДУШНЫХ ФИЛЬТРОВ", 0.25, 4.15, 12.5, 0.4,
                 size=14, bold=True, color=C_GOLD)
    add_rect(slide, 0.15, 4.55, 0.06, 2.7, fill_color=C_GOLD)
    h5_data = [
        "Задокументировано только у ОДНОГО агрегата: №69 (25.05.2026, 10 721 м/ч)",
        "Засорённый фильтр → снижение расхода воздуха → богатая смесь → T-выпуска ↑↑",
        "У №69 поражены 4L, 5L, 6L, 7L, 7R — цилиндры левого ряда и крайние правого",
        "У других агрегатов состояние фильтров при отказе НЕ задокументировано",
        "РЕКОМЕНДАЦИЯ: Регулярная проверка и документирование состояния фильтров при каждом ТО",
        "Эксплуатация в условиях горнорудного карьера → повышенное пылевыделение",
        "СТАТУС H5: Подтверждённый фактор для №69; для парка в целом — требует проверки",
    ]
    y = 4.6
    for s in h5_data:
        c = C_GOLD if "РЕКОМЕНДАЦИЯ" in s or "СТАТУС" in s else C_BLACK
        b = "ЗАДОКУМЕНТИРОВАНО ТОЛЬКО" in s or "СТАТУС" in s or "РЕКОМЕНДАЦИЯ" in s
        add_text_box(slide, "• " + s, 0.3, y, 12.8, 0.38, size=11, bold=b, color=c)
        y += 0.42
    print("  Slide 15: H4+H5 done")

    # ── SLIDE 16: OIL CONSUMPTION / ASH ───────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "H6: ПОТРЕБЛЕНИЕ МАСЛА И ЗОЛЬНЫЕ ОТЛОЖЕНИЯ",
                 "Отягчающий фактор, а не первопричина | Данные «Горной Евразии»")
    # Image
    pdf_repnhl = os.path.join(WORK_DIR, "气门盘部磨损分析报告-1.docx RUS.pdf")
    add_pdf_image(slide, pdf_repnhl, page_num=8, left=0.15, top=1.3, width=4.5, height=4.0)
    # Right data
    add_rect(slide, 4.8, 1.25, 8.4, 6.15, fill_color=C_LGRAY)
    add_text_box(slide, "ДОЛИВКИ МАСЛА (апрель-май 2026):", 4.9, 1.3, 8.2, 0.4,
                 size=14, bold=True, color=C_DARK)
    oil_data = [
        "Доливки: 10-30 литров на агрегат раз в 2-7 дней",
        "Данные охватывают: 20+ единиц NTE200",
        "Норм. расход QSK50: <1 л/час ≈ 10 л/смена (10 ч)",
        "Фактический расход: значительно выше нормы",
        "",
        "ЗОЛЬНЫЕ ОТЛОЖЕНИЯ (EDS NHL):",
        "Ca + Zn + P = зола ZDDP (масляная присадка)",
        "ZDDP разлагается при T>180°C → образует золу",
        "Зола абразивна — ускоряет трение фаски",
        "",
        "МЕХАНИЗМ ОТЯГЧЕНИЯ:",
        "Перегрев → сгорание масляной плёнки →",
        "интенсивное золообразование → абразив →",
        "цикл ускоренного износа",
        "",
        "НО: Это НЕ первопричина!",
        "Зола появляется КАК СЛЕДСТВИЕ перегрева,",
        "а не как его причина.",
        "",
        "Высокое потребление масла само по себе",
        "указывает на проблему ЦПГ или уплотнений.",
    ]
    y = 1.75
    for s in oil_data:
        c = C_ACCENT if "ПЕРЕГРЕВ" in s.upper() or "НО:" in s else \
            (C_DARK if s.endswith(":") else C_BLACK)
        b = s.endswith(":") or "НО:" in s
        add_text_box(slide, s, 4.9, y, 8.1, 0.33, size=11, bold=b, color=c)
        y += 0.33
    print("  Slide 16: H6 oil/ash done")

    # ── SLIDE 17: HYPOTHESIS MATRIX ───────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "МАТРИЦА ОЦЕНКИ ГИПОТЕЗ",
                 "Доказательная база и приоритет проверки")
    matrix = [
        ["Гипотеза", "Доказательств\nпрямых", "Косвенные\nдоказательства", "Охватывает\nвесь парк?",
         "Объясняет\nранние отказы", "Приоритет"],
        ["H1: ECM (дератинг, RPM, TIB)", "3 (сравн. 730E)", "730E = 0 отказов", "ДА", "ДА", "★★★★★"],
        ["H2: Нет Stellite (конструкция)", "2 (NHL чертёж)", "Твёрдость <HRC", "ДА", "ДА", "★★★★★"],
        ["H3: Зазоры клапанов (ТО)", "1 (доп. анализ)", "Повторные отказы", "ЧАСТИЧНО", "НЕТ", "★★★☆☆"],
        ["H4: Форсунки", "2 (тест №73)", "Сопутств. замены", "ЧАСТИЧНО", "НЕТ", "★★★☆☆"],
        ["H5: Возд. фильтр", "1 (только №69)", "Только 1 агрегат", "НЕТ", "НЕТ", "★★☆☆☆"],
        ["H6: Зола масла (ОТЯГЧ.)", "3 (EDS NHL)", "Все агрегаты", "ДА", "ЧАСТИЧНО", "ОТЯГЧ."],
        ["H7: Режим эксплуатации", "0", "Общий контекст", "ДА", "ЧАСТИЧНО", "★★☆☆☆"],
        ["H8: Охлаждение ГБЦ", "0", "Топогр. карта", "ЧАСТИЧНО", "НЕТ", "★★☆☆☆"],
    ]
    add_table(slide, matrix,
              [3.8, 1.7, 2.0, 1.6, 1.8, 1.6], 0.15, 1.25, 13.0, 5.9,
              header_bg=C_DARK, cell_size=10, header_size=11)
    add_rect(slide, 0.15, 7.2, 13.0, 0.22, fill_color=RGBColor(0xFF, 0xE8, 0xE0))
    add_text_box(slide,
                 "H1 + H2 = СИСТЕМНЫЕ причины | H3 + H4 = СОПУТСТВУЮЩИЕ | H6 = ОТЯГЧАЮЩИЙ фактор",
                 0.25, 7.22, 12.8, 0.18, size=11, bold=True, color=C_ACCENT)
    print("  Slide 17: Hypothesis matrix done")

    # ── SLIDE 18: FLEET REPAIR MAP ─────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "КАРТА ОТКАЗОВ ПО ПАРКУ | ГБЦ РЕМОНТЫ.XLSX",
                 "463 строки ремонтных событий | 43 единицы | авг 2025 — май 2026")
    # Unit grid
    units_with_failures = {
        43: "16033", 44: "14751", 45: "17863", 46: "15103",
        47: "12890", 48: "14997", 50: "15433", 51: "10783",
        52: "18505", 53: "17550", 55: "17316", 57: "14808",
        58: "10702", 59: "15143", 62: "13917", 69: "10130",
        72: "10000", 73: "9998", 74: "9869", 76: "9754",
        77: "10016", 78: "9323", 81: "4696", 83: "2776",
    }
    all_units = list(range(43, 93))
    cols = 10
    cell_w, cell_h = 1.27, 0.85
    y0 = 1.3
    add_text_box(slide, "Каждая ячейка = 1 агрегат NTE200. Красный = отказ клапанов зафиксирован",
                 0.2, 1.3, 12.9, 0.32, size=11, bold=False, color=C_GRAY, italic=True)
    y0 = 1.65
    for i, unit in enumerate(all_units):
        row = i // cols
        col = i % cols
        x = 0.15 + col * cell_w
        y = y0 + row * cell_h
        if unit in units_with_failures:
            # Critical
            c83 = unit == 83 or unit == 81
            bg = RGBColor(0xC0, 0x20, 0x10) if c83 else C_ACCENT
        else:
            bg = C_GREEN
        add_rect(slide, x, y, cell_w-0.06, cell_h-0.08, fill_color=bg)
        label = f"№{unit}"
        add_title_box(slide, label, x+0.02, y+0.02, cell_w-0.1, 0.35,
                      size=13, bold=True, color=C_WHITE, bg=bg, align=PP_ALIGN.CENTER)
        if unit in units_with_failures:
            mh = units_with_failures[unit]
            add_text_box(slide, f"{mh} м/ч", x+0.02, y+0.38, cell_w-0.1, 0.32,
                         size=9, bold=False, color=C_WHITE, align=PP_ALIGN.CENTER)
    # Legend
    add_rect(slide, 0.15, 6.8, 13.0, 0.55, fill_color=C_LGRAY)
    for lx, col, label in [(0.3, C_ACCENT, "Клапанный отказ"),
                            (3.5, RGBColor(0xC0, 0x20, 0x10), "Катастрофический (81/83)"),
                            (7.2, C_GREEN, "Отказов не зафиксировано")]:
        add_rect(slide, lx, 6.88, 0.4, 0.3, fill_color=col)
        add_text_box(slide, label, lx+0.45, 6.88, 3.0, 0.3, size=11, color=C_DARK)
    print("  Slide 18: Fleet map done")

    # ── SLIDE 19: VERIFICATION PLAN ───────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "ПЛАН ВЕРИФИКАЦИИ ГИПОТЕЗ",
                 "Конкретные действия для подтверждения/опровержения каждой гипотезы")
    vplan = [
        ("H1", "ECM НАСТРОЙКИ — ВЕРИФИКАЦИЯ", C_ACCENT, [
            "Запросить полный файл ECM с 730E AC (Komatsu) с объекта",
            "Провести сравнительный дамп ECM с 2-3 исправных NTE200",
            "Инициировать изменение параметра C_EPD_OT_RPM_Drt_En на 1",
            "Установить задержки по маслу/ОЖ на ≤10 с (по требованию Cummins)",
            "Снизить max RPM до 1990 и TIB коррекцию до 100%",
            "Наблюдение: отсутствие новых отказов в течение 6 мес.",
        ]),
        ("H2", "STELLITE КЛАПАНЫ — ВЕРИФИКАЦИЯ", C_MID, [
            "Заказать 16 выпускных клапанов со Stellite-наплавкой (CES51005-S)",
            "Установить на 2-3 агрегата в следующем плановом ремонте",
            "Зафиксировать наработку до следующего осмотра клапанов",
            "Сравнить состояние фаски при осмотре (через 6000 м/ч)",
            "Измерить твёрдость фаски по Роквеллу при демонтаже",
        ]),
        ("H3", "ЗАЗОРЫ КЛАПАНОВ — ВЕРИФИКАЦИЯ", C_GOLD, [
            "Измерить и задокументировать зазоры при КАЖДОМ ремонте (щупами)",
            "Сравнить с нормативами QSK50: вып. 0.69±0.05 мм, впуск. 0.36±0.05 мм",
            "Провести аудит квалификации механиков (calibrated training)",
            "При повторном отказе — первым делом проверить зазоры",
        ]),
        ("H4", "ФОРСУНКИ — ВЕРИФИКАЦИЯ", C_GREEN, [
            "Провести стендовые тесты форсунок всего парка (приоритет: №43, 47, 48, 55, 73)",
            "Анализ данных INSITE: параметры впрыска, balance rates",
            "При замене клапанов — обязательная проверка форсунок того же цилиндра",
        ]),
    ]
    y = 1.25
    for hnum, htitle, col, actions in vplan:
        add_rect(slide, 0.15, y, 0.7, len(actions)*0.3 + 0.45, fill_color=col)
        add_text_box(slide, hnum, 0.15, y+0.05, 0.7, 0.35,
                     size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_text_box(slide, htitle, 0.95, y+0.02, 12.2, 0.38,
                     size=12, bold=True, color=col)
        ay = y + 0.42
        for a in actions:
            add_text_box(slide, "→ " + a, 0.95, ay, 12.2, 0.3, size=10, color=C_BLACK)
            ay += 0.32
        y = ay + 0.2
    print("  Slide 19: Verification plan done")

    # ── SLIDE 20: RECOMMENDATIONS ─────────────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_WHITE)
    slide_header(slide, "РЕКОМЕНДАЦИИ",
                 "Приоритетные действия по устранению причин отказов клапанов QSK50")
    recs = [
        ("СРОЧНО (до следующего ремонта)", C_ACCENT, [
            "Перепрограммировать ECM: включить C_EPD_OT_RPM_Drt_En=1, снизить задержки защит до ≤10 с",
            "Снизить max RPM до 1990 об/мин, установить TIB коррекцию 100%",
            "Проверить и активировать все триггеры по температуре ОГ",
            "Провести аудит состояния воздушных фильтров всего парка",
            "Обязательная проверка форсунок при каждом клапанном ремонте",
        ]),
        ("КРАТКОСРОЧНО (в течение 1 месяца)", C_GOLD, [
            "Заказать выпускные клапаны со Stellite-наплавкой (CES51005-S) для пилотной установки",
            "Разработать и внедрить форму контроля зазоров клапанов",
            "Провести обучение механиков по регулировке клапанов QSK50",
            "Установить режим ведения журнала INSITE на всех агрегатах",
            "Запросить у NHL обновлённое техническое руководство по клапанам",
        ]),
        ("СРЕДНЕСРОЧНО (в течение 3 месяцев)", C_MID, [
            "Проанализировать рабочий цикл (нагрузка, температуры) через INSITE на 5+ агрегатах",
            "Провести сравнение с параметрами 730E AC — идентичный двигатель, тот же объект",
            "По результатам пилота Stellite — принять решение о серийном переходе",
            "Создать базу данных ТО с отслеживанием зазоров и состояния форсунок",
        ]),
        ("СИСТЕМНЫЕ МЕРЫ", C_DARK, [
            "Ввести регулярные ТО по часам (не только по пробегу) с контролем клапанов",
            "Разработать протокол замены масла с учётом высокого потребления",
            "Предложить NHL внесение изменений в конструкцию (Stellite) как обязательного стандарта",
            "Провести анализ всех случаев с NHL на уровне конструкторов",
        ]),
    ]
    y = 1.25
    for title, col, items in recs:
        n = len(items)
        h = n * 0.33 + 0.52
        add_rect(slide, 0.15, y, 13.0, h, fill_color=RGBColor(0xF5, 0xF8, 0xFF))
        add_rect(slide, 0.15, y, 3.2, 0.42, fill_color=col)
        add_text_box(slide, title, 0.2, y+0.03, 3.1, 0.36,
                     size=11, bold=True, color=C_WHITE)
        ay = y + 0.46
        for item in items:
            add_text_box(slide, "✓ " + item, 0.3, ay, 12.8, 0.32, size=10, color=C_BLACK)
            ay += 0.33
        y = ay + 0.1
    print("  Slide 20: Recommendations done")

    # ── SLIDE 21: SUMMARY / CONCLUSIONS ──────────────────────────────────────
    slide = prs.slides.add_slide(blank)
    add_bg(slide, C_DARK)
    add_rect(slide, 0, 0, 13.33, 7.5, fill_color=RGBColor(0x0D, 0x1F, 0x33))
    add_rect(slide, 0, 1.1, 13.33, 0.05, fill_color=C_GOLD)
    add_title_box(slide, "ВЫВОДЫ И СЛЕДУЮЩИЕ ШАГИ",
                  0.3, 0.15, 12.7, 0.95,
                  size=30, bold=True, color=C_WHITE,
                  bg=RGBColor(0x0D, 0x1F, 0x33))
    conclusions = [
        ("1", "Масштаб проблемы СИСТЕМНЫЙ",
         "20+ единиц, 30+ ремонтных событий, от 2776 до 18505 м/ч. Не единичный дефект."),
        ("2", "ECM настройки — главная рабочая версия",
         "Отключённый дератинг + задержки защит + повышенные RPM и TIB. 730E AC (те же двигатели, тот же объект) — клапанных отказов нет."),
        ("3", "Отсутствие Stellite — конструктивная уязвимость",
         "40-44 HRC против 55-65 HRC. Stellite сохраняет твёрдость при 800°C — критично для выпускного клапана."),
        ("4", "Зола масла — отягчающий, но не первичный фактор",
         "EDS анализ NHL подтверждает: Ca/Zn/P + Si/Al. Зола ускоряет износ, но появляется как следствие перегрева."),
        ("5", "Верификация ECM — приоритет №1",
         "Перепрограммирование ECM — быстрый, обратимый тест. Если за 6 мес. нет новых отказов → H1 подтверждена."),
    ]
    y = 1.3
    for num, title, desc in conclusions:
        add_rect(slide, 0.2, y, 0.65, 0.8, fill_color=C_GOLD)
        add_title_box(slide, num, 0.2, y, 0.65, 0.8,
                      size=22, bold=True, color=C_DARK, bg=C_GOLD, align=PP_ALIGN.CENTER)
        add_text_box(slide, title, 0.95, y+0.02, 12.2, 0.32,
                     size=14, bold=True, color=C_GOLD,
                     bg=RGBColor(0x0D, 0x1F, 0x33))
        add_text_box(slide, desc, 0.95, y+0.38, 12.2, 0.42,
                     size=11, bold=False, color=RGBColor(0xCC, 0xDD, 0xEE),
                     bg=RGBColor(0x0D, 0x1F, 0x33))
        y += 0.95
    add_rect(slide, 0.2, 6.1, 12.9, 0.05, fill_color=C_GOLD)
    add_text_box(slide,
                 "Следующий шаг: Согласование с Cummins и NHL → Перепрограммирование ECM → Пилотная установка Stellite клапанов",
                 0.2, 6.2, 12.9, 0.42,
                 size=13, bold=True, color=C_WHITE,
                 align=PP_ALIGN.CENTER,
                 bg=RGBColor(0x0D, 0x1F, 0x33))
    add_text_box(slide,
                 "Конфиденциально | АО «Полюс Магадан» | Анализ неисправностей QSK50 / NTE200 | Май 2026",
                 0.2, 6.8, 12.9, 0.5,
                 size=10, bold=False,
                 color=RGBColor(0x60, 0x80, 0xA0),
                 align=PP_ALIGN.CENTER,
                 bg=RGBColor(0x0D, 0x1F, 0x33))
    print("  Slide 21: Conclusions done")

    # ── SAVE ──────────────────────────────────────────────────────────────────
    out_path = os.path.join(WORK_DIR, "Презентация_анализ_клапанов_QSK50_NTE200.pptx")
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    import os as _os
    size = _os.path.getsize(out_path)
    print(f"File size: {size/1024/1024:.1f} MB")
    return out_path


if __name__ == "__main__":
    build_presentation()
