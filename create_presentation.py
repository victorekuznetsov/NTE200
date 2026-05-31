#!/usr/bin/env python3
"""
Presentation: Analysis of exhaust valve failures on QSK50 engines, NTE200 dump trucks
АО «Развитие» / АО «Полюс Магадан» brand style
Dark theme: #141c24 bg, #3ef0af accent, #e84855 red, #f5a623 amber
"""
import os
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

WORK_DIR = "/home/user/NTE200"

# ── Brand Palette (АО Развитие / dashboard_final.html) ──────────────────────
C_BG      = RGBColor(0x14, 0x1C, 0x24)   # --bg  main background
C_BG2     = RGBColor(0x1C, 0x26, 0x30)   # --bg2 sidebar
C_BG3     = RGBColor(0x24, 0x2F, 0x39)   # --bg3 section bg
C_CARD    = RGBColor(0x2A, 0x31, 0x38)   # --card
C_CARD2   = RGBColor(0x32, 0x39, 0x41)   # --card2
C_ACC     = RGBColor(0x3E, 0xF0, 0xAF)   # --acc teal/green accent
C_RED     = RGBColor(0xE8, 0x48, 0x55)   # --red alerts
C_AMB     = RGBColor(0xF5, 0xA6, 0x23)   # --amb amber/warnings
C_GRN     = RGBColor(0x22, 0xC5, 0x5E)   # --grn green ok
C_BLU     = RGBColor(0x7E, 0x83, 0xFA)   # --blu blue
C_PUR     = RGBColor(0xB6, 0x68, 0xE4)   # --pur purple
C_TX      = RGBColor(0xE8, 0xED, 0xF6)   # --tx primary text
C_TX2     = RGBColor(0xB3, 0xBD, 0xDC)   # --tx2 secondary text
C_TX3     = RGBColor(0x80, 0x86, 0x8B)   # --tx3 muted text
C_BRD     = RGBColor(0x30, 0x3A, 0x44)   # border-like separator
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

FONT      = "DejaVu Sans"
FONT_MONO = "DejaVu Sans Mono"

import fitz  # PyMuPDF

# ── Low-level helpers ────────────────────────────────────────────────────────
def add_bg(slide, color=C_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill=C_CARD, line=None, line_w=1):
    shape = slide.shapes.add_shape(
        1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)
    return shape

def add_label(slide, text, left, top, width, height,
              size=12, bold=False, color=C_TX, align=PP_ALIGN.LEFT,
              bg=None, mono=False, italic=False):
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    if bg:
        box.fill.solid()
        box.fill.fore_color.rgb = bg
    else:
        box.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT_MONO if mono else FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box

def slide_chrome(slide, title, subtitle=None, accent_bar=True):
    """Dark header bar with teal accent line"""
    add_rect(slide, 0, 0, 13.33, 0.95, fill=C_BG2)
    add_label(slide, title, 0.25, 0.06, 12.6, 0.58,
              size=22, bold=True, color=C_TX)
    if subtitle:
        add_label(slide, subtitle, 0.25, 0.62, 12.6, 0.3,
                  size=11, bold=False, color=C_TX3)
    if accent_bar:
        add_rect(slide, 0, 0.95, 13.33, 0.035, fill=C_ACC)

def kpi_card(slide, value, label, left, top, w=2.9, h=1.5,
             val_color=C_ACC, bg=C_CARD):
    add_rect(slide, left, top, w, h, fill=bg, line=C_BRD, line_w=1)
    add_rect(slide, left, top, w, 0.04, fill=val_color)
    add_label(slide, value, left+0.12, top+0.1, w-0.24, 0.72,
              size=34, bold=True, color=val_color, align=PP_ALIGN.CENTER, mono=True)
    add_label(slide, label, left+0.1, top+0.82, w-0.2, 0.6,
              size=11, bold=False, color=C_TX2, align=PP_ALIGN.CENTER)

def section_card(slide, title, items, left, top, w, h,
                 title_color=C_ACC, bg=C_CARD, bullet="▸"):
    add_rect(slide, left, top, w, h, fill=bg, line=C_BRD, line_w=1)
    add_rect(slide, left, top, 0.04, h, fill=title_color)
    add_label(slide, title, left+0.12, top+0.05, w-0.18, 0.32,
              size=12, bold=True, color=title_color)
    y = top + 0.38
    step = (h - 0.42) / max(len(items), 1)
    for item in items:
        if isinstance(item, tuple):
            txt, col = item
        else:
            txt, col = item, C_TX2
        add_label(slide, f"{bullet} {txt}", left+0.14, y, w-0.2, step+0.05,
                  size=10, color=col)
        y += step

def add_table_dark(slide, rows, col_widths, left, top, w, h,
                   hdr_bg=C_BG2, row_bg=C_CARD, alt_bg=C_BG3,
                   hdr_size=10, cell_size=9, hdr_color=C_ACC):
    from pptx.util import Pt as PT
    nr, nc = len(rows), len(rows[0])
    tbl = slide.shapes.add_table(nr, nc,
                                  Inches(left), Inches(top),
                                  Inches(w), Inches(h)).table
    total = sum(col_widths)
    for i, cw in enumerate(col_widths):
        tbl.columns[i].width = Inches(w * cw / total)
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = hdr_bg
            elif r % 2 == 1:
                cell.fill.fore_color.rgb = row_bg
            else:
                cell.fill.fore_color.rgb = alt_bg
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(cell_text)
            run.font.name = FONT
            run.font.size = PT(hdr_size if r == 0 else cell_size)
            run.font.bold = (r == 0)
            run.font.color.rgb = hdr_color if r == 0 else C_TX2

def extract_pdf_image(pdf_path, page=0, dpi=100):
    try:
        doc = fitz.open(pdf_path)
        if page >= len(doc): page = 0
        mat = fitz.Matrix(dpi/72, dpi/72)
        pix = doc[page].get_pixmap(matrix=mat)
        data = pix.tobytes("png")
        doc.close()
        return data
    except Exception as e:
        print(f"  [WARN] {pdf_path} p{page}: {e}")
        return None

def add_pdf_image(slide, pdf_path, page=0, left=0.1, top=1.1,
                  width=5.0, height=4.0, dpi=100):
    data = extract_pdf_image(pdf_path, page, dpi)
    if data:
        slide.shapes.add_picture(io.BytesIO(data),
                                  Inches(left), Inches(top),
                                  Inches(width), Inches(height))
        return True
    return False

# ═══════════════════════════════════════════════════════════════════════════════
def build():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    print("Building АО Развитие styled presentation...")

    # ── SLIDE 1: TITLE ─────────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    # Subtle grid lines — mimics dashboard bg pattern
    for i in range(0, 35):
        add_rect(s, 0, i*0.22, 13.33, 0.005, fill=RGBColor(0x1E, 0x2A, 0x34))
    for i in range(0, 60):
        add_rect(s, i*0.23, 0, 0.004, 7.5, fill=RGBColor(0x1E, 0x2A, 0x34))

    # Left accent stripe
    add_rect(s, 0, 0, 0.06, 7.5, fill=C_ACC)

    # Logo area
    add_rect(s, 0.15, 0.18, 3.2, 0.52, fill=C_BG3)
    add_label(s, "АО РАЗВИТИЕ", 0.22, 0.2, 3.0, 0.34,
              size=13, bold=True, color=C_ACC, mono=True)
    add_label(s, "Технический анализ", 0.22, 0.5, 3.0, 0.18,
              size=9, color=C_TX3)

    # Main title
    add_rect(s, 0.15, 1.1, 12.9, 0.04, fill=C_ACC)
    add_label(s, "АНАЛИЗ ПРИЧИН ИЗНОСА", 0.22, 1.2, 12.6, 1.05,
              size=46, bold=True, color=C_TX)
    add_label(s, "ВЫПУСКНЫХ КЛАПАНОВ QSK50", 0.22, 2.22, 12.6, 1.0,
              size=46, bold=True, color=C_ACC)
    add_rect(s, 0.15, 3.2, 12.9, 0.04, fill=C_BRD)

    add_label(s,
              "Самосвалы NHL NTE200  ·  АО «Полюс Магадан»  ·  Омчак, Магаданская область",
              0.22, 3.3, 12.6, 0.42, size=16, color=C_TX2)

    # Stat pills
    stats = [
        ("20+",   "ЕДИНИЦ ПАРКА",         C_RED),
        ("30+",   "РЕМОНТНЫХ СОБЫТИЙ",     C_AMB),
        ("2 776", "М/Ч — РАННИЙ ОТКАЗ",   C_ACC),
        ("18 505","М/Ч — ПОЗДНИЙ ОТКАЗ",  C_GRN),
    ]
    for i, (val, lbl, col) in enumerate(stats):
        x = 0.15 + i * 3.25
        add_rect(s, x, 3.9, 3.1, 0.9, fill=C_BG3, line=col, line_w=1)
        add_label(s, val, x+0.1, 3.93, 2.9, 0.5,
                  size=26, bold=True, color=col, align=PP_ALIGN.CENTER, mono=True)
        add_label(s, lbl, x+0.1, 4.41, 2.9, 0.32,
                  size=9, color=C_TX3, align=PP_ALIGN.CENTER)

    add_label(s, "Конфиденциально  ·  Только для внутреннего использования  ·  Май 2026",
              0.15, 7.1, 12.9, 0.3, size=10, color=C_TX3, align=PP_ALIGN.CENTER)
    print("  01 Title")

    # ── SLIDE 2: МАСШТАБ ───────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "МАСШТАБ ПРОБЛЕМЫ",
                 "Статистика отказов выпускных клапанов QSK50 | Весь парк NTE200")

    kpis = [
        ("20+",   "агрегатов\nзатронуто",       C_RED,  C_BG3),
        ("30+",   "ремонтных\nсобытий",          C_AMB,  C_BG3),
        ("2 776", "м/ч\nмин. наработка",         C_ACC,  C_BG3),
        ("100%",  "парка №43-83\nс отказами",    C_PUR,  C_BG3),
    ]
    for i, (v, l, vc, bg) in enumerate(kpis):
        kpi_card(s, v, l, 0.2 + i*3.28, 1.1, w=3.1, h=1.55,
                 val_color=vc, bg=bg)

    tbl = [
        ["Ед.", "Наработка до отказа", "Событий", "Поражение"],
        ["№43", "16 033 м/ч", "1", "5 позиций (3L,4L,1R,2R,3R)"],
        ["№44", "14 751 м/ч", "1", "ВСЕ 14 позиций"],
        ["№47", "12 890 м/ч", "2", "7 поз. + замена ГБЦ"],
        ["№48", "14 997 м/ч", "2", "17 клапанов + турбина + гидроудар"],
        ["№51", "10 783 м/ч", "1", "Замена ГБЦ (трещина, 6R)"],
        ["№53", "17 550 м/ч", "1", "ВСЕ 16 поз. (2 выпуск + 2 впуск)"],
        ["№55", "17 316 м/ч", "1", "ВСЕ 16 поз. + 4 форсунки"],
        ["№57", "14 808 м/ч", "1", "6 позиций"],
        ["№58", "10 702 м/ч", "2", "Клапан + толкатель"],
        ["№69", "10 130 м/ч", "2", "Засорение воздушных фильтров + поршни 3R/4R"],
        ["№72", "10 000 м/ч", "2", "8 позиций (2 эпизода)"],
        ["№76", " 9 754 м/ч", "2", "Клапаны (2 эпизода)"],
        ["№78", " 9 323 м/ч", "1", "3R/4R/6R/6L"],
        ["№81", " 4 696 м/ч", "1", "ВСЕ позиции + замена ГБЦ"],
        ["№83", " 2 776 м/ч", "1", "КАТАСТРОФИЧЕСКИЙ — шатун/поршень/ГБЦ"],
    ]
    add_table_dark(s, tbl, [0.8, 1.8, 1.0, 4.4], 0.2, 2.75, 13.0, 4.6)
    print("  02 Scale")

    # ── SLIDE 3: ХРОНОЛОГИЯ ────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "ХРОНОЛОГИЯ ОТКАЗОВ",
                 "Авг 2025 — Май 2026 | 14 ключевых событий")

    events = [
        ("Авг 2025", "№47", "Замена ГБЦ — ранний эпизод",            "12 890", C_AMB),
        ("Авг 2025", "№51", "Замена ГБЦ, трещина головки",            "10 783", C_AMB),
        ("Авг 2025", "№58", "Замена толкателя клапана",                "10 702", C_AMB),
        ("Янв 2026", "№43", "Замена клапанов + 9 форсунок",           "16 033", C_BLU),
        ("Фев 2026", "№48", "17 клапанов + турбина",                   "14 997", C_RED),
        ("Мар 2026", "№47", "Повт. отказ — 10 клапанов + форсунки",  "12 890", C_RED),
        ("Мар 2026", "№58", "Повт. отказ — клапан + ГБЦ",            "15 572", C_AMB),
        ("Мар 2026", "№83", "КАТАСТРОФА — шатун/поршень/ГБЦ",        " 2 776", C_RED),
        ("Апр 2026", "№48", "Гидроудар, водомасляная эмульсия",       "после рем.", C_RED),
        ("Май 2026", "№55", "ВСЕ 16 позиций + 4 форсунки",           "17 316", C_BLU),
        ("Май 2026", "№69", "Клапаны + засорение фильтров + поршни", "10 721", C_AMB),
        ("Май 2026", "№78", "3R/4R/6R/6L",                            " 9 323", C_AMB),
        ("Май 2026", "№72", "Отказ форсунки цил. 5",                  "10 283", C_TX3),
        ("Май 2026", "№73", "Форсунки — аномалии теста L2/R1",        "~9 998", C_TX3),
    ]
    add_rect(s, 0.2, 1.1, 13.0, 0.34, fill=C_BG2)
    for xoff, lbl in [(0.1,"ДАТА"), (1.25,"ЕД."), (2.5,"СОБЫТИЕ"), (10.1,"НАРАБОТКА")]:
        add_label(s, lbl, 0.2+xoff, 1.12, 1.8, 0.28,
                  size=9, bold=True, color=C_TX3)
    row_h = 0.38
    for i, (dt, unit, evt, mh, col) in enumerate(events):
        y = 1.44 + i * row_h
        bg = C_CARD if i % 2 == 0 else C_BG3
        add_rect(s, 0.2, y, 13.0, row_h-0.02, fill=bg)
        add_rect(s, 0.2, y, 0.04, row_h-0.02, fill=col)
        add_label(s, dt,   0.32, y+0.04, 0.9,  0.3, size=10, color=C_TX3)
        add_label(s, unit, 1.45, y+0.04, 1.0,  0.3, size=11, bold=True, color=col)
        add_label(s, evt,  2.55, y+0.04, 7.5,  0.3, size=10, color=C_TX)
        add_label(s, mh,   10.1, y+0.04, 3.0,  0.3, size=10, color=C_ACC,
                  align=PP_ALIGN.RIGHT, mono=True)
    print("  03 Timeline")

    # ── SLIDE 4: ФИЗИЧЕСКИЕ ПРИЗНАКИ ──────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "ФИЗИЧЕСКИЕ ПРИЗНАКИ ПОВРЕЖДЕНИЙ",
                 "Данные разборки и дефектовки | Сводка по всем агрегатам")

    findings = [
        ("ПРОСАДКА ФАСКИ (РЕЦЕССИЯ КЛАПАНА)",
         "Главный признак. Посадочная поверхность выпускного клапана «утопает» в седле. "
         "Обнаружена на всех исследованных двигателях. Твёрдость фаски 40-44 HRC (базовая).", C_RED),
        ("ПЛАСТИЧЕСКАЯ ДЕФОРМАЦИЯ — «ЗАУСЕНЕЦ» НА ФАСКЕ",
         "Бурт пластической деформации = признак превышения допустимой температуры тарелки. "
         "Зафиксировано в отчёте лаборатории NHL (MS&T2026033).", C_RED),
        ("ЗОЛЬНЫЕ ОТЛОЖЕНИЯ (Ca/Zn/P + Si/Al)",
         "EDS-анализ NHL: Ca, Zn, P — зола ZDDP-присадок. Si, Al — силикатная пыль. "
         "Роль: ОТЯГЧАЮЩИЙ фактор (абразив). Не первопричина — следствие перегрева.", C_AMB),
        ("ПОВРЕЖДЕНИЕ СМЕЖНЫХ ЭЛЕМЕНТОВ",
         "№47/48/53/55/81: замена ГБЦ целиком. №83 (2 776 м/ч): разрушение шатуна и поршня. "
         "№48: гидроудар (водомасляная эмульсия) после первого ремонта.", C_AMB),
        ("ФОРСУНКИ",
         "Заменены совместно с клапанами: №43 (9 шт.), №47, №55 (4 шт.), №73. "
         "Протечка форсунки → дополнительный нагрев → перегрев клапана.", C_BLU),
        ("ТОПОГРАФИЯ ОТКАЗОВ",
         "Наиболее часто поражены: 3R, 4R, 6R, 6L, 7L, 7R. "
         "Потенциально указывает на неравномерность охлаждения ГБЦ.", C_ACC),
    ]
    y = 1.12
    for title, desc, col in findings:
        add_rect(s, 0.2, y, 13.0, 0.9, fill=C_CARD, line=C_BRD, line_w=1)
        add_rect(s, 0.2, y, 0.05, 0.9, fill=col)
        add_label(s, title, 0.38, y+0.04, 12.7, 0.32, size=12, bold=True, color=col)
        add_label(s, desc,  0.38, y+0.38, 12.7, 0.48, size=10, color=C_TX2)
        y += 0.94
    print("  04 Physical findings")

    # ── SLIDE 5: NHL LAB ───────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "АНАЛИЗ ЛАБОРАТОРИИ NHL  |  MS&T2026033",
                 "Инженер Тао Лан | СЭМ + EDS спектроскопия + профилометрия")
    nhl = os.path.join(WORK_DIR, "气门盘部磨损分析报告-1.docx RUS.pdf")
    add_pdf_image(s, nhl, page=2, left=0.2,  top=1.12, width=5.8, height=3.3)
    add_pdf_image(s, nhl, page=3, left=6.1,  top=1.12, width=5.8, height=3.3)
    add_rect(s, 0.2, 4.55, 12.9, 2.75, fill=C_BG3, line=C_BRD, line_w=1)
    add_rect(s, 0.2, 4.55, 12.9, 0.04, fill=C_ACC)
    cols_nhl = [
        ("МЕТОДЫ",        ["СЭМ (электронная микроскопия)", "EDS-спектроскопия", "Измерение твёрдости HRC", "Профилометрия"], C_ACC),
        ("EDS РЕЗУЛЬТАТ", ["Ca + Zn + P = зола ZDDP", "Si + Al = силикатная пыль", "Слои толщиной 5-50 мкм", "Абразивная среда в зоне контакта"], C_AMB),
        ("МОРФОЛОГИЯ",    ["Пластич. деформация фаски", "«Заусенец» = признак T° > лимита", "Адгезионный + абразивный износ", "Термическое повреждение"], C_RED),
        ("ВЫВОД NHL",     ["Комплексная причина:", "1. Термическое повреждение", "2. Абразивный износ (зола+пыль)", "3. Адгезионный износ"], C_BLU),
    ]
    for i, (ttl, items, col) in enumerate(cols_nhl):
        x = 0.3 + i * 3.2
        add_label(s, ttl, x, 4.62, 3.1, 0.28, size=10, bold=True, color=col)
        for j, item in enumerate(items):
            add_label(s, item, x, 4.95 + j*0.52, 3.1, 0.48, size=10, color=C_TX2)
    print("  05 NHL lab")

    # ── SLIDE 6: EDS ───────────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "EDS-АНАЛИЗ ОТЛОЖЕНИЙ НА КЛАПАНЕ",
                 "Элементный состав | Ca/Zn/P = зола ZDDP | Si/Al = силикатная пыль")
    add_pdf_image(s, nhl, page=5, left=0.2, top=1.12, width=6.2, height=4.1)
    add_pdf_image(s, nhl, page=6, left=6.6, top=1.12, width=6.4, height=4.1)
    add_rect(s, 0.2, 5.32, 12.9, 1.98, fill=C_BG3, line=C_BRD, line_w=1)
    add_rect(s, 0.2, 5.32, 12.9, 0.04, fill=C_AMB)
    add_label(s,
              "ИНТЕРПРЕТАЦИЯ: Ca + Zn + P = зола ZDDP (разлагается при T > 180°C). "
              "Si + Al = силикатная пыль через воздушный тракт. "
              "Создают абразивную среду в зоне клапан-седло. "
              "Зола — ОТЯГЧАЮЩИЙ фактор: появляется КАК СЛЕДСТВИЕ перегрева, а не его причина.",
              0.32, 5.38, 12.6, 1.82, size=11, color=C_TX2)
    print("  06 EDS")

    # ── SLIDE 7: UNIT 83 ───────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    add_rect(s, 0, 0, 13.33, 0.95, fill=C_BG2)
    add_label(s, "КАТАСТРОФИЧЕСКИЙ ОТКАЗ  |  АГРЕГАТ №83",
              0.25, 0.06, 12.6, 0.58, size=22, bold=True, color=C_RED)
    add_label(s, "2 776 м/ч наработки — самый ранний отказ в парке",
              0.25, 0.62, 12.6, 0.28, size=11, color=C_TX3)
    add_rect(s, 0, 0.95, 13.33, 0.035, fill=C_RED)
    pdf83 = os.path.join(WORK_DIR, "Тех отчет 83 от 30.03.2026.pdf")
    add_pdf_image(s, pdf83, page=0, left=0.2, top=1.12, width=4.1, height=3.1)
    add_pdf_image(s, pdf83, page=2, left=4.4, top=1.12, width=4.1, height=3.1)
    add_pdf_image(s, pdf83, page=4, left=8.7, top=1.12, width=4.4, height=3.1)
    add_rect(s, 0.2, 4.35, 12.9, 2.95, fill=C_BG3, line=C_RED, line_w=1)
    data83 = [
        ("Гаражный №:", "83"),
        ("Наработка:", "2 776 м/ч  (!!! норм. ресурс >12 000 м/ч)"),
        ("Зафиксировано:", "Разрушение шатуна и поршня"),
        ("Замена:", "ГБЦ + шатун + поршень + уплотнения"),
        ("Вер. причина:", "Перегрев → потеря плотности клапана → прорыв газов → разрушение ЦПГ"),
        ("Значение:", "Катастрофический отказ при минимальной наработке — системная проблема"),
    ]
    y = 4.42
    for lbl, val in data83:
        add_label(s, lbl, 0.35, y, 2.4, 0.4, size=11, bold=True, color=C_AMB)
        add_label(s, val, 2.85, y, 10.1, 0.4, size=11, color=C_TX)
        y += 0.44
    print("  07 Unit 83")

    # ── SLIDE 8: UNIT 48 ───────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "АГРЕГАТ №48 — ДВОЙНОЙ ОТКАЗ",
                 "14 997 м/ч: 17 клапанов + турбина | Повторный: гидроудар")
    pdf48  = os.path.join(WORK_DIR, "Тех.отчет 48.pdf")
    pdf48b = os.path.join(WORK_DIR, "Тех оточет 48 от 31.03.2026.pdf")
    add_pdf_image(s, pdf48,  page=0, left=0.2, top=1.12, width=3.9, height=2.9)
    add_pdf_image(s, pdf48,  page=2, left=4.2, top=1.12, width=3.9, height=2.9)
    add_pdf_image(s, pdf48b, page=1, left=8.3, top=1.12, width=4.8, height=2.9)
    section_card(s, "ПЕРВЫЙ РЕМОНТ (Фев 2026)",
                 ["Наработка: 14 997 м/ч",
                  "17 выпускных клапанов (почти все позиции)",
                  "Замена турбокомпрессора",
                  "Замена ГБЦ — позиция 3R",
                  "Визуально: нагар, просадка, абразивный износ"],
                 0.2, 4.15, 6.35, 2.95, title_color=C_AMB, bg=C_CARD)
    section_card(s, "ВТОРОЙ ОТКАЗ (Мар 2026 — после ремонта)",
                 [("Обнаружена водомасляная эмульсия", C_RED),
                  ("Гидроудар — угроза разрушения ЦПГ", C_RED),
                  "Источник воды: нарушение уплотнения ГБЦ",
                  "Через ~3-4 недели после первого ремонта",
                  "Повторная разборка и дефектовка ГБЦ"],
                 6.7, 4.15, 6.45, 2.95, title_color=C_RED, bg=C_CARD)
    print("  08 Unit 48")

    # ── SLIDE 9: UNITS 55 & 69 ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "АГРЕГАТЫ №55 И №69 — ПОКАЗАТЕЛЬНЫЕ СЛУЧАИ",
                 "№55: ВСЕ 16 позиций | №69: Засорение воздушных фильтров")
    pdf55 = os.path.join(WORK_DIR, "Тех отчет 55.pdf")
    add_pdf_image(s, pdf55, page=0, left=0.2, top=1.12, width=3.8, height=2.7)
    section_card(s, "№55 — 17 316 М/Ч",
                 ["ВСЕ 16 позиций клапанов заменены",
                  "Заменены 4 форсунки (1L, 2L, 3L, 6L)",
                  "Самая большая наработка при полном поражении",
                  ("Все клапаны: признаки перегрева", C_RED),
                  "→ Длительная работа с отклонениями"],
                 4.15, 1.12, 4.1, 2.7, title_color=C_BLU, bg=C_CARD)
    pdf48_fb = os.path.join(WORK_DIR, "Тех.отчет 48.pdf")
    add_pdf_image(s, pdf48_fb, page=3, left=8.4, top=1.12, width=3.5, height=2.7)
    section_card(s, "№69 — 10 721 М/Ч",
                 [("ЗАСОРЕНЫ ВОЗДУШНЫЕ ФИЛЬТРЫ (!)", C_RED),
                  "Поражены: 4L, 5L, 6L, 7L, 7R",
                  "Замена поршней 3R и 4R",
                  "2 эпизода: 10 130 и 10 721 м/ч",
                  "Единственный агрегат с документ. засорением"],
                 8.3, 3.95, 4.85, 2.85, title_color=C_AMB, bg=C_CARD)
    section_card(s, "№55 — СВЯЗЬ С ECM",
                 ["ECM без дератинга → двигатель перегружен",
                  "Сбой форсунки = двойное тепловыделение",
                  "Итог: замена ВСЕХ клапанов + ГБЦ",
                  "Длительная работа без сигналов тревоги"],
                 0.2, 3.95, 8.0, 2.85, title_color=C_ACC, bg=C_CARD)
    add_rect(s, 0.2, 6.9, 12.9, 0.48, fill=C_BG3)
    add_label(s,
              "№69: засорение фильтров → богатая смесь → T° выхлопа ↑↑ → ускоренный износ. "
              "У остальных агрегатов состояние фильтров при отказе не задокументировано.",
              0.32, 6.92, 12.7, 0.44, size=10, color=C_TX3)
    print("  09 Units 55 & 69")

    # ── SLIDE 10: HYPOTHESES OVERVIEW ─────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "ОБЗОР ГИПОТЕЗ",
                 "8 версий — от конструктивных особенностей до условий эксплуатации")

    hyps = [
        ("H1", "Настройки ECM двигателя",
         "Отключённый дератинг, задержки 51/43 с, обороты 2100 vs 1990, TIB 200% vs 100%", C_RED, "ВЫСОКИЙ"),
        ("H2", "Отсутствие наплавки Stellite",
         "Фаска 40-44 HRC (NTE200) vs 55-65 HRC (730E AC). Разница твёрдости >50%", C_RED, "ВЫСОКИЙ"),
        ("H3", "Ошибка регулировки зазоров",
         "Малый зазор → клапан не закрывается → потеря охлаждения → перегрев", C_AMB, "СРЕДНИЙ"),
        ("H4", "Неисправность форсунок",
         "Протечка/нарушение распыла → дожигание → перегрев. №43/47/55/73 — замена форсунок", C_AMB, "СРЕДНИЙ"),
        ("H5", "Засорение воздушного фильтра",
         "Богатая смесь → T° ОГ ↑. Задокументировано только для №69.", C_BLU, "НИЗКИЙ"),
        ("H6", "Зола масла и абразивная пыль",
         "EDS NHL: Ca/Zn/P + Si/Al. ОТЯГЧАЮЩИЙ фактор — не первопричина.", C_BLU, "ОТЯГЧ."),
        ("H7", "Рабочий цикл / нагружение",
         "Горный рудник, высота, температурные перепады. Требует сравнит. анализа.", C_TX3, "НЕОПР."),
        ("H8", "Конструктив охлаждения ГБЦ",
         "Топография: 3R/4R/6R/6L чаще поражаются. Неравномерный теплоотвод?", C_TX3, "НЕОПР."),
    ]
    y = 1.12
    for hnum, htitle, hdesc, col, level in hyps:
        add_rect(s, 0.2, y, 1.05, 0.72, fill=col)
        add_label(s, hnum, 0.2, y+0.08, 1.05, 0.56,
                  size=22, bold=True, color=C_BG, align=PP_ALIGN.CENTER, mono=True)
        add_rect(s, 1.35, y, 9.2, 0.72, fill=C_CARD, line=C_BRD, line_w=1)
        add_label(s, htitle, 1.45, y+0.03, 9.0, 0.3, size=12, bold=True, color=col)
        add_label(s, hdesc,  1.45, y+0.34, 9.0, 0.36, size=10, color=C_TX2)
        lev_c = C_RED if level == "ВЫСОКИЙ" else (C_AMB if level == "СРЕДНИЙ" else C_BLU if level in ("НИЗКИЙ","ОТЯГЧ.") else C_TX3)
        add_rect(s, 10.65, y, 2.55, 0.72, fill=lev_c)
        add_label(s, level, 10.65, y+0.18, 2.55, 0.36,
                  size=14, bold=True, color=C_BG, align=PP_ALIGN.CENTER)
        y += 0.77
    print("  10 Hypotheses overview")

    # ── SLIDE 11: H1 ECM TABLE ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    add_rect(s, 0, 0, 13.33, 0.95, fill=C_BG2)
    add_label(s, "H1: НАСТРОЙКИ ECM — NTE200 vs KOMATSU 730E AC",
              0.25, 0.06, 12.6, 0.58, size=22, bold=True, color=C_RED)
    add_label(s, "Один двигатель QSK50 — один объект — разные параметры защиты",
              0.25, 0.62, 12.6, 0.28, size=11, color=C_TX3)
    add_rect(s, 0, 0.95, 13.33, 0.035, fill=C_RED)
    add_rect(s, 0.2, 1.1, 12.9, 0.44, fill=RGBColor(0x2D, 0x10, 0x10))
    add_label(s,
              "КЛЮЧЕВОЙ ФАКТ: Komatsu 730E AC с аналогичными QSK50 работает НА ТОМ ЖЕ объекте. "
              "Случаев замены клапанов на 730E AC НЕ ЗАФИКСИРОВАНО.",
              0.3, 1.12, 12.7, 0.38, size=12, bold=True, color=C_RED)
    ecm = [
        ["Параметр ECM", "Komatsu 730E AC", "NHL NTE200", "Риск"],
        ["C_EPD_OT_RPM_Drt_En\n(флаг дератинга)", "ВКЛЮЧЁН (1)", "ОТКЛЮЧЁН (0)", "КРИТИЧЕСКИЙ"],
        ["Задержка по давл. масла", "Немедленно", "51 секунда (!)", "КРИТИЧЕСКИЙ"],
        ["Задержка по темп. ОЖ", "Немедленно", "43 секунды (!)", "ВЫСОКИЙ"],
        ["Максимальные обороты", "1 990 об/мин", "2 100 об/мин", "ВЫСОКИЙ"],
        ["Топливная коррекция TIB", "100%", "200% (!)", "ВЫСОКИЙ"],
        ["Триггеры по темп. ОГ", "Активны (все)", "ВСЕ ОТКЛЮЧЕНЫ", "КРИТИЧЕСКИЙ"],
        ["Дератинг при перегреве", "Активен", "Не работает", "КРИТИЧЕСКИЙ"],
        ["Обороты ХХ", "700 об/мин", "700 об/мин", "Норма"],
    ]
    add_table_dark(s, ecm, [4.2, 2.8, 2.8, 2.2], 0.2, 1.62, 13.0, 5.55,
                   hdr_bg=C_BG2, hdr_color=C_RED, hdr_size=11, cell_size=10)
    add_rect(s, 0.2, 7.22, 12.9, 0.24, fill=C_BG3)
    add_label(s, "«КРИТИЧЕСКИЙ» = ведёт к перегреву без снижения мощности и без сигнала оператору",
              0.3, 7.23, 12.7, 0.2, size=10, bold=True, color=C_RED)
    print("  11 ECM table")

    # ── SLIDE 12: H1 MECHANISM ────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "H1: ЦЕПОЧКА ПОВРЕЖДЕНИЙ ЧЕРЕЗ ECM",
                 "Как отключённый дератинг ведёт к отказу клапанов")

    steps = [
        ("ECM: Дератинг ОТКЛ\n+ задержки 43-51 с\n+ RPM 2100\n+ TIB 200%", C_RED),
        ("Двигатель работает\nбез снижения мощности\nпри перегреве\nи аварии", C_AMB),
        ("T° выхлопных газов\nвыше нормы\nбез сигнала\nоператору", C_AMB),
        ("ПЕРЕГРЕВ тарелки\nклапана → пластич.\nдеформация фаски\n→ рецессия", C_RED),
    ]
    for i, (txt, col) in enumerate(steps):
        x = 0.2 + i * 3.3
        add_rect(s, x, 1.18, 3.1, 1.8, fill=C_CARD, line=col, line_w=2)
        add_rect(s, x, 1.18, 3.1, 0.06, fill=col)
        add_label(s, txt, x+0.1, 1.3, 2.9, 1.6,
                  size=13, color=col, align=PP_ALIGN.CENTER)
        if i < 3:
            add_label(s, "→", x+3.1, 1.9, 0.2, 0.4,
                      size=22, bold=True, color=C_TX3, align=PP_ALIGN.CENTER)

    add_rect(s, 0.2, 3.18, 12.9, 0.76, fill=RGBColor(0x2D, 0x10, 0x10), line=C_RED, line_w=1)
    add_label(s,
              "РЕЗУЛЬТАТ: Рецессия фаски → потеря компрессии → прорыв горячих газов → "
              "ускоренный абразивный износ → последовательный выход из строя клапанов",
              0.35, 3.22, 12.6, 0.68, size=13, bold=True, color=C_RED)

    add_rect(s, 0.2, 4.08, 12.9, 3.22, fill=C_BG3, line=C_BRD, line_w=1)
    add_label(s, "КОСВЕННЫЕ ДОКАЗАТЕЛЬСТВА:", 0.35, 4.14, 12.5, 0.32,
              size=12, bold=True, color=C_ACC)
    evidence = [
        "▸  NTE200: операторов инструктируют «не глушить, давать остывать на ХХ» — косвенное признание перегревов",
        "▸  INSITE CSV (агрегат N43, 18 523 м/ч): зафиксированы температурные отклонения в рабочем цикле",
        "▸  730E AC на том же объекте, с теми же двигателями QSK50 — клапанных отказов НЕТ",
        "▸  Самые ранние отказы (№83: 2 776 м/ч | №81: 4 696 м/ч) не объяснимы обычным усталостным износом",
        "▸  Диапазон наработок (2 776 — 18 505 м/ч) указывает на вариабельный триггер, а не дефект партии",
    ]
    for i, e in enumerate(evidence):
        add_label(s, e, 0.35, 4.52 + i*0.52, 12.6, 0.48, size=11, color=C_TX2)
    print("  12 ECM mechanism")

    # ── SLIDE 13: H2 STELLITE ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "H2: ОТСУТСТВИЕ НАПЛАВКИ STELLITE",
                 "Конструктивное различие фаски клапана NTE200 vs 730E AC")

    add_rect(s, 0.2, 1.12, 6.4, 5.8, fill=C_CARD, line=C_RED, line_w=1)
    add_rect(s, 0.2, 1.12, 6.4, 0.06, fill=C_RED)
    add_label(s, "NTE200 / QSK50 (CES51005)",
              0.3, 1.2, 6.1, 0.38, size=14, bold=True, color=C_RED)
    nte = [
        ("Материал клапана:", "Inconel 751"),
        ("Наплавка фаски:", "ОТСУТСТВУЕТ (по чертежу NHL)"),
        ("Твёрдость фаски:", "40–44 HRC"),
        ("Защита от износа:", "только базовый материал"),
        ("Поведение при T°:", "размягчается при 700-800°C"),
        ("Статус в парке:", "рецессия у ВСЕХ исследованных агрегатов"),
    ]
    y = 1.65
    for lbl, val in nte:
        add_label(s, lbl, 0.3, y, 2.4, 0.38, size=11, bold=True, color=C_TX3)
        c = C_RED if "ОТСУТСТВУЕТ" in val or "ВСЕХ" in val else C_TX
        add_label(s, val, 2.75, y, 3.75, 0.38, size=11, color=c)
        y += 0.52

    add_rect(s, 6.75, 1.12, 6.4, 5.8, fill=C_CARD, line=C_ACC, line_w=1)
    add_rect(s, 6.75, 1.12, 6.4, 0.06, fill=C_ACC)
    add_label(s, "Komatsu 730E AC / QSK50",
              6.85, 1.2, 6.1, 0.38, size=14, bold=True, color=C_ACC)
    kat = [
        ("Материал клапана:", "Inconel 751 / аналог"),
        ("Наплавка фаски:", "STELLITE Co-Cr сплав"),
        ("Твёрдость фаски:", "55–65 HRC"),
        ("Защита от износа:", "твёрдая + термостойкая наплавка"),
        ("Поведение при T°:", "сохраняет твёрдость до 800-900°C"),
        ("Статус в парке:", "клапанных отказов не зафиксировано"),
    ]
    y = 1.65
    for lbl, val in kat:
        add_label(s, lbl, 6.85, y, 2.4, 0.38, size=11, bold=True, color=C_TX3)
        c = C_ACC if "STELLITE" in val or "не зафиксировано" in val.lower() else C_TX
        add_label(s, val, 9.3, y, 3.75, 0.38, size=11, color=c)
        y += 0.52

    add_rect(s, 0.2, 7.0, 12.9, 0.4, fill=C_BG3)
    add_label(s,
              "Stellite снижает износ в 3-5× при высоких температурах. "
              "Применение Stellite-клапанов в NTE200 — конкретная конструктивная рекомендация.",
              0.3, 7.02, 12.7, 0.34, size=11, bold=True, color=C_ACC)
    print("  13 H2 Stellite")

    # ── SLIDE 14: H3 + H4 ─────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "H3: ЗАЗОРЫ КЛАПАНОВ  |  H4: ФОРСУНКИ",
                 "Сопутствующие факторы теплового повреждения")

    add_rect(s, 0.2, 1.12, 6.4, 5.85, fill=C_CARD, line=C_AMB, line_w=1)
    add_rect(s, 0.2, 1.12, 6.4, 0.06, fill=C_AMB)
    add_label(s, "H3: ЗАЗОРЫ КЛАПАНОВ",
              0.3, 1.2, 6.1, 0.34, size=13, bold=True, color=C_AMB)
    h3_lines = [
        ("Нормативы QSK50:", False, C_TX3),
        ("  Выпускной:  0.69 ± 0.05 мм", False, C_ACC),
        ("  Впускной:   0.36 ± 0.05 мм", False, C_ACC),
        ("", False, C_TX3),
        ("ЕСЛИ зазор < нормы:", True, C_AMB),
        ("→ Клапан не закрывается полностью", False, C_RED),
        ("→ Площадь контакта с седлом ↓", False, C_RED),
        ("→ Теплоотвод от тарелки ↓", False, C_RED),
        ("→ Локальный перегрев фаски", False, C_RED),
        ("→ Ускоренная рецессия", False, C_RED),
        ("", False, C_TX3),
        ("Агрегат №59 — многократные отказы:", True, C_AMB),
        ("  15 143 → ремонт → 16 219 → повтор", False, C_TX2),
        ("  Регулировались ли зазоры при ремонте?", False, C_TX3),
        ("", False, C_TX3),
        ("Повторные отказы также у:", True, C_AMB),
        ("  №47, №48, №69, №72, №76", False, C_TX2),
    ]
    y = 1.6
    for txt, b, c in h3_lines:
        add_label(s, txt, 0.3, y, 6.1, 0.33, size=10, bold=b, color=c)
        y += 0.34

    add_rect(s, 6.75, 1.12, 6.4, 5.85, fill=C_CARD, line=C_BLU, line_w=1)
    add_rect(s, 6.75, 1.12, 6.4, 0.06, fill=C_BLU)
    add_label(s, "H4: ФОРСУНКИ",
              6.85, 1.2, 6.1, 0.34, size=13, bold=True, color=C_BLU)
    h4_lines = [
        ("Механизм:", True, C_TX3),
        ("  Протечка/нарушение распылителя", False, C_TX2),
        ("  → неполное сгорание", False, C_RED),
        ("  → дожигание у клапана", False, C_RED),
        ("  → перегрев тарелки", False, C_RED),
        ("", False, C_TX3),
        ("Задокументированные случаи:", True, C_TX3),
        ("  №43 — 9 форсунок (!)", False, C_AMB),
        ("  №47 — форсунки + клапаны", False, C_TX2),
        ("  №55 — 4 форсунки (1L,2L,3L,6L)", False, C_TX2),
        ("  №73 — тест L2+R1: аномалии", False, C_TX2),
        ("  №85 — 8 002 м/ч (отдельный отчёт)", False, C_TX2),
        ("  №72 — 3 700 м/ч, цил. 5", False, C_TX2),
        ("", False, C_TX3),
        ("INSITE — проверка:", True, C_TX3),
        ("  Fuel balance rates по цилиндрам", False, C_ACC),
        ("  Отклонение >±3% = подозрит. форс.", False, C_ACC),
    ]
    y = 1.6
    for txt, b, c in h4_lines:
        add_label(s, txt, 6.85, y, 6.1, 0.33, size=10, bold=b, color=c)
        y += 0.34
    print("  14 H3+H4")

    # ── SLIDE 15: H6 OIL/ASH ──────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "H6: ЗОЛА МАСЛА — ОТЯГЧАЮЩИЙ ФАКТОР",
                 "Данные EDS NHL + доливки «Горной Евразии» | Апр-Май 2026")
    add_pdf_image(s, nhl, page=8, left=0.2, top=1.12, width=4.5, height=4.0)
    add_rect(s, 4.85, 1.12, 8.3, 6.2, fill=C_BG3, line=C_BRD, line_w=1)
    add_label(s, "ДАННЫЕ О ДОЛИВКАХ МАСЛА (Апр-Май 2026):",
              4.95, 1.18, 8.0, 0.34, size=12, bold=True, color=C_AMB)
    oil = [
        ("Доливки:", "10-30 л на агрегат раз в 2-7 дней"),
        ("Охват:", "20+ единиц NTE200"),
        ("Норм. расход QSK50:", "<1 л/ч ≈ 10 л/смену"),
        ("Факт. расход:", "значительно выше нормы"),
    ]
    y = 1.6
    for lbl, val in oil:
        add_label(s, lbl, 4.95, y, 2.8, 0.34, size=11, bold=True, color=C_TX3)
        add_label(s, val, 7.8,  y, 5.2, 0.34, size=11, color=C_TX)
        y += 0.4
    add_rect(s, 4.85, 2.78, 8.3, 0.04, fill=C_BRD)
    add_label(s, "МЕХАНИЗМ ЗОЛООБРАЗОВАНИЯ:", 4.95, 2.85, 8.0, 0.34,
              size=12, bold=True, color=C_AMB)
    mech = [
        "Перегрев клапана → сгорание масляной плёнки",
        "ZDDP разлагается при T > 180°C → Ca/Zn/P зола",
        "Зола абразивна → ускоряет трение фаски",
        "Цикл: перегрев → зола → больше трения → перегрев",
    ]
    y = 3.26
    for m in mech:
        c = C_ACC if "Цикл" in m else C_TX2
        add_label(s, "▸ " + m, 4.95, y, 8.0, 0.36, size=11, color=c)
        y += 0.38
    add_rect(s, 4.85, 5.12, 8.3, 0.04, fill=C_BRD)
    add_rect(s, 4.85, 5.2, 8.3, 1.12, fill=RGBColor(0x1C, 0x2A, 0x20))
    add_label(s, "ВАЖНО: Зола — СЛЕДСТВИЕ перегрева, не его причина.",
              4.95, 5.24, 8.0, 0.36, size=12, bold=True, color=C_GRN)
    add_label(s,
              "Высокое потребление масла указывает на проблему ЦПГ или уплотнений. "
              "Необходим анализ масла (металлы износа, вискозиметрия).",
              4.95, 5.62, 8.0, 0.65, size=11, color=C_TX2)
    print("  15 H6 oil/ash")

    # ── SLIDE 16: HYPOTHESIS MATRIX ───────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "МАТРИЦА ОЦЕНКИ ГИПОТЕЗ",
                 "Доказательная база и приоритет верификации")
    matrix = [
        ["Гипотеза", "Прямые\nдоказательства", "Косвенные\nдоказательства",
         "Весь\nпарк?", "Ранние\nотказы", "Приоритет"],
        ["H1: ECM (дератинг, RPM, TIB)", "3 (сравн. 730E)", "730E = 0 отказов", "ДА", "ДА", "★★★★★"],
        ["H2: Нет Stellite (конструкция)", "2 (чертёж NHL)", "HRC < нормы", "ДА", "ДА", "★★★★★"],
        ["H3: Зазоры клапанов (ТО)", "1 (доп. анализ)", "Повторные отказы", "ЧАСТИЧНО", "НЕТ", "★★★☆☆"],
        ["H4: Форсунки", "2 (тест №73)", "Сопутств. замены", "ЧАСТИЧНО", "НЕТ", "★★★☆☆"],
        ["H5: Возд. фильтр", "1 (только №69)", "1 агрегат", "НЕТ", "НЕТ", "★★☆☆☆"],
        ["H6: Зола масла (ОТЯГЧАЮЩИЙ)", "3 (EDS NHL)", "Все агрегаты", "ДА", "ЧАСТИЧНО", "ОТЯГЧ."],
        ["H7: Режим эксплуатации", "0", "Контекст", "ДА", "ЧАСТИЧНО", "★★☆☆☆"],
        ["H8: Охлаждение ГБЦ", "0", "Топография", "ЧАСТИЧНО", "НЕТ", "★★☆☆☆"],
    ]
    add_table_dark(s, matrix,
                   [3.8, 1.9, 2.0, 1.3, 1.5, 1.5], 0.2, 1.12, 12.9, 5.9,
                   hdr_bg=C_BG2, hdr_color=C_ACC, hdr_size=11, cell_size=10)
    add_rect(s, 0.2, 7.1, 12.9, 0.3, fill=C_BG3)
    add_label(s,
              "H1 + H2 = СИСТЕМНЫЕ причины  ·  H3 + H4 = СОПУТСТВУЮЩИЕ  ·  H6 = ОТЯГЧАЮЩИЙ фактор",
              0.3, 7.12, 12.7, 0.24, size=11, bold=True, color=C_ACC)
    print("  16 Matrix")

    # ── SLIDE 17: FLEET MAP ───────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "КАРТА ОТКАЗОВ ПО ПАРКУ",
                 "База данных ГБЦ Ремонты.xlsx | 463 строки | Агрегаты №43-92")

    failures = {
        43:"16033", 44:"14751", 45:"17863", 46:"15103",
        47:"12890", 48:"14997", 50:"15433", 51:"10783",
        52:"18505", 53:"17550", 55:"17316", 57:"14808",
        58:"10702", 59:"15143", 62:"13917", 69:"10130",
        72:"10000", 73:"9998",  74:"9869",  76:"9754",
        77:"10016", 78:"9323",  81:"4696",  83:"2776",
    }
    add_label(s,
              "Красный = отказ клапанов зафиксирован  ·  Тёмно-красный = катастрофический  ·  Зелёный = нет данных об отказе",
              0.2, 1.04, 12.9, 0.22, size=10, color=C_TX3)
    cols_n = 10
    cw, ch = 1.28, 0.86
    x0, y0 = 0.2, 1.3
    for i, unit in enumerate(range(43, 93)):
        row = i // cols_n
        col = i % cols_n
        x = x0 + col * cw
        y = y0 + row * ch
        if unit in failures:
            bg = RGBColor(0x6B, 0x0F, 0x0F) if unit in (81, 83) else \
                 RGBColor(0x38, 0x10, 0x10)
            border = C_RED
            vc = C_RED
        else:
            bg = C_CARD
            border = C_GRN
            vc = C_GRN
        add_rect(s, x, y, cw-0.06, ch-0.08, fill=bg, line=border, line_w=1)
        add_label(s, f"#{unit}", x+0.05, y+0.04, cw-0.14, 0.32,
                  size=13, bold=True, color=vc, align=PP_ALIGN.CENTER, mono=True)
        if unit in failures:
            add_label(s, failures[unit], x+0.04, y+0.38, cw-0.1, 0.28,
                      size=8, color=C_TX3, align=PP_ALIGN.CENTER, mono=True)
    print("  17 Fleet map")

    # ── SLIDE 18: VERIFICATION PLAN ──────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "ПЛАН ВЕРИФИКАЦИИ ГИПОТЕЗ",
                 "Конкретные действия для подтверждения/опровержения каждой версии")

    vplan = [
        ("H1", "ECM — ПРИОРИТЕТ №1", C_RED, [
            "Считать полный ECM с 2-3 NTE200 + 1 Komatsu 730E AC через INSITE",
            "Включить C_EPD_OT_RPM_Drt_En=1 | Задержки по маслу/ОЖ ≤10 с | RPM 1990 | TIB 100%",
            "Активировать все триггеры температуры ОГ",
            "Наблюдение 6 мес: 0 отказов на перепрограмм. агрегатах = подтверждено",
        ]),
        ("H2", "STELLITE КЛАПАНЫ", C_AMB, [
            "Запросить у NHL клапаны CES51005-S со Stellite-наплавкой",
            "Пилотная установка на 2-3 агрегата при следующем плановом ремонте",
            "Осмотр через 6 000 м/ч: измерить твёрдость (HRC) и рецессию фаски",
            "Сравнить с базовыми клапанами той же наработки — разница ≥50% = подтверждено",
        ]),
        ("H3", "ЗАЗОРЫ КЛАПАНОВ", C_BLU, [
            "Ввести протокол: измерение зазоров ДО и ПОСЛЕ каждого ремонта с записью",
            "При ближайшем ТО проверить зазоры на 5 агрегатах без клапанного ремонта",
            "Обучение механиков: калиброванный щуп 0.05 мм шаг | допуск ±0.05 мм",
        ]),
        ("H4", "ФОРСУНКИ", C_ACC, [
            "Стендовые тесты форсунок: №43, №47, №55, №73 — давление, возврат, распыл",
            "INSITE: fuel balance rates → отклонение >±3% = проблемная форсунка",
            "Правило: при любом клапанном ремонте — обязательная проверка форсунки цилиндра",
        ]),
    ]
    y = 1.12
    for hnum, htitle, col, actions in vplan:
        nh = len(actions)
        h = nh * 0.4 + 0.5
        add_rect(s, 0.2, y, 12.9, h, fill=C_CARD, line=C_BRD, line_w=1)
        add_rect(s, 0.2, y, 0.06, h, fill=col)
        add_rect(s, 0.26, y, 1.4, h, fill=C_BG3)
        add_label(s, hnum, 0.27, y+0.06, 1.35, 0.38,
                  size=18, bold=True, color=col, align=PP_ALIGN.CENTER, mono=True)
        add_label(s, htitle, 1.78, y+0.06, 11.2, 0.34,
                  size=12, bold=True, color=col)
        for j, action in enumerate(actions):
            add_label(s, "→  " + action,
                      1.78, y + 0.46 + j*0.38, 11.2, 0.35, size=10, color=C_TX2)
        y += h + 0.08
    print("  18 Verification")

    # ── SLIDE 19: RECOMMENDATIONS ─────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    slide_chrome(s, "РЕКОМЕНДАЦИИ",
                 "Приоритетные действия по устранению причин отказов клапанов QSK50")

    recs = [
        ("СРОЧНО",       "до следующего ремонта",   C_RED, [
            "Перепрограммировать ECM: включить дератинг, задержки ≤10 с, RPM 1990, TIB 100%",
            "Активировать все триггеры температуры выхлопных газов",
            "Проверить состояние воздушных фильтров всего парка",
            "Обязательная проверка форсунки при каждом клапанном ремонте",
        ]),
        ("КРАТКОСРОЧНО", "в течение 1 месяца",      C_AMB, [
            "Заказать Stellite-клапаны CES51005-S — пилотная партия на 3 агрегата",
            "Разработать и внедрить форму контроля зазоров клапанов",
            "Стендовые тесты форсунок приоритетных агрегатов",
            "Включить логирование INSITE на всех агрегатах",
        ]),
        ("СРЕДНЕСРОЧНО", "в течение 3 месяцев",     C_BLU, [
            "Анализ рабочего цикла по INSITE на 5+ агрегатах",
            "Сравнительный анализ с 730E AC на том же объекте",
            "По результатам пилота Stellite: решение о серийном переходе",
        ]),
        ("СИСТЕМНО",     "постоянные меры",          C_ACC, [
            "Предложить NHL внесение Stellite как обязательного стандарта конструкции",
            "База данных ТО: зазоры, форсунки, T° — на каждый агрегат",
            "Протокол замены масла с учётом высокого потребления",
        ]),
    ]
    y = 1.12
    for urgency, sub, col, items in recs:
        h = len(items) * 0.42 + 0.56
        add_rect(s, 0.2, y, 12.9, h, fill=C_CARD, line=C_BRD, line_w=1)
        add_rect(s, 0.2, y, 0.06, h, fill=col)
        add_rect(s, 0.26, y, 2.0, h, fill=C_BG3)
        add_label(s, urgency, 0.28, y+0.1, 1.9, 0.36,
                  size=13, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_label(s, sub, 0.28, y+0.46, 1.9, 0.24,
                  size=9, color=C_TX3, align=PP_ALIGN.CENTER)
        for j, item in enumerate(items):
            add_label(s, "✓  " + item, 2.38, y + 0.1 + j*0.42, 10.6, 0.38,
                      size=11, color=C_TX2)
        y += h + 0.08
    print("  19 Recommendations")

    # ── SLIDE 20: CONCLUSIONS ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    add_bg(s, C_BG)
    # Top accent
    add_rect(s, 0, 0, 13.33, 0.06, fill=C_ACC)
    add_rect(s, 0, 0.06, 13.33, 0.9, fill=C_BG2)
    add_label(s, "ВЫВОДЫ И СЛЕДУЮЩИЕ ШАГИ",
              0.25, 0.1, 12.6, 0.58, size=22, bold=True, color=C_TX)
    add_label(s, "Анализ причин износа выпускных клапанов QSK50 | NTE200 | АО «Полюс Магадан»",
              0.25, 0.66, 12.6, 0.28, size=11, color=C_TX3)
    add_rect(s, 0, 0.96, 13.33, 0.035, fill=C_ACC)

    concs = [
        ("1", "МАСШТАБ СИСТЕМНЫЙ",
         "20+ единиц, 30+ событий, 2 776 — 18 505 м/ч. Не единичный дефект партии.", C_RED),
        ("2", "ECM — ГЛАВНАЯ РАБОЧАЯ ВЕРСИЯ",
         "Отключённый дератинг + задержки + RPM 2100 + TIB 200%. 730E AC (те же двигатели, тот же объект) — 0 отказов.", C_RED),
        ("3", "STELLITE — КОНСТРУКТИВНАЯ УЯЗВИМОСТЬ",
         "40-44 HRC против 55-65 HRC. Stellite сохраняет твёрдость при 800°C — критично для выпускного клапана.", C_AMB),
        ("4", "ЗОЛА — ОТЯГЧАЮЩИЙ ФАКТОР",
         "EDS NHL подтверждает Ca/Zn/P + Si/Al. Ускоряет износ, но является следствием перегрева.", C_BLU),
        ("5", "СЛЕДУЮЩИЙ ШАГ — ВЕРИФИКАЦИЯ",
         "Перепрограммировать ECM (быстрый тест) + пилотные Stellite-клапаны на 2-3 агрегата.", C_ACC),
    ]
    y = 1.1
    for num, title, desc, col in concs:
        add_rect(s, 0.2, y, 12.9, 1.1, fill=C_CARD, line=C_BRD, line_w=1)
        add_rect(s, 0.2, y, 0.06, 1.1, fill=col)
        add_rect(s, 0.26, y, 0.95, 1.1, fill=C_BG3)
        add_label(s, num, 0.27, y+0.24, 0.93, 0.6,
                  size=26, bold=True, color=col, align=PP_ALIGN.CENTER, mono=True)
        add_label(s, title, 1.32, y+0.08, 11.5, 0.36, size=13, bold=True, color=col)
        add_label(s, desc,  1.32, y+0.52, 11.5, 0.5,  size=11, color=C_TX2)
        y += 1.14

    add_rect(s, 0.2, 6.88, 12.9, 0.55, fill=C_BG3)
    add_label(s,
              "АО «Развитие»  ·  Технический анализ  ·  Конфиденциально  ·  Май 2026",
              0.2, 6.9, 12.9, 0.26, size=11, color=C_TX3, align=PP_ALIGN.CENTER)
    add_label(s,
              "Cummins QSK50  ·  NHL NTE200  ·  АО «Полюс Магадан»",
              0.2, 7.18, 12.9, 0.22, size=10, color=C_ACC, align=PP_ALIGN.CENTER)
    print("  20 Conclusions")

    # ── SAVE ──────────────────────────────────────────────────────────────────
    out = os.path.join(WORK_DIR, "Презентация_анализ_клапанов_QSK50_NTE200.pptx")
    prs.save(out)
    sz = os.path.getsize(out)
    print(f"\nSaved: {out}")
    print(f"Size:  {sz/1024/1024:.1f} MB  |  {len(prs.slides)} slides")
    return out


if __name__ == "__main__":
    build()
