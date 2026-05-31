#!/usr/bin/env python3
"""
Presentation: Analysis of exhaust valve failures on QSK50 engines, NTE200 dump trucks
КТГ БЕЛАЗ brand style — two-tone (dark title/section + light content slides)
"""
import os, io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import fitz

WORK_DIR = "/home/user/NTE200"

# ── КТГ БЕЛАЗ Brand Palette ──────────────────────────────────────────────────
C_DARK   = RGBColor(0x2A, 0x31, 0x38)  # dark card / title bg
C_DARK2  = RGBColor(0x1F, 0x26, 0x2D)  # deeper dark
C_LIGHT  = RGBColor(0xF5, 0xF5, 0xF5)  # content slide bg
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_ACC    = RGBColor(0x34, 0xD3, 0x99)  # bright emerald green (primary accent)
C_GRN2   = RGBColor(0x10, 0xB9, 0x81)  # dark emerald
C_AMB    = RGBColor(0xF5, 0x9E, 0x0B)  # amber
C_RED    = RGBColor(0xEF, 0x44, 0x44)  # red
C_BLU    = RGBColor(0x3B, 0x82, 0xF6)  # blue
C_TX1    = RGBColor(0x1F, 0x29, 0x37)  # dark text on light bg
C_TX2    = RGBColor(0x6B, 0x72, 0x80)  # medium gray text
C_TX3    = RGBColor(0x9C, 0xA3, 0xAF)  # muted text
C_LGRAY  = RGBColor(0xD1, 0xD5, 0xDB)  # light border / divider
C_CARD_L = RGBColor(0xF9, 0xFA, 0xFB)  # light card bg
C_CARD_D = RGBColor(0x37, 0x41, 0x51)  # dark card on dark slide

FONT = "Calibri"

# ── Helpers ──────────────────────────────────────────────────────────────────
def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def rect(slide, l, t, w, h, fill=C_DARK, line=None, lw=1):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(lw)
    return s

def txt(slide, text, l, t, w, h,
        size=12, bold=False, color=C_TX1,
        align=PP_ALIGN.LEFT, bg=None, italic=False):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    if bg: box.fill.solid(); box.fill.fore_color.rgb = bg
    else: box.fill.background()
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return box

def multiline(slide, lines, l, t, w, h,
              def_size=11, def_bold=False, def_color=C_TX1,
              def_align=PP_ALIGN.LEFT, bg=None):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    if bg: box.fill.solid(); box.fill.fore_color.rgb = bg
    else: box.fill.background()
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):   tx, b, c, sz, al = item, def_bold, def_color, def_size, def_align
        elif len(item)==5:           tx, b, c, sz, al = item
        elif len(item)==4:           tx, b, c, sz = item; al = def_align
        elif len(item)==3:           tx, b, c = item; sz = def_size; al = def_align
        else:                        tx, b = item; c=def_color; sz=def_size; al=def_align
        if first: p = tf.paragraphs[0]; first = False
        else:     p = tf.add_paragraph()
        p.alignment = al
        r = p.add_run()
        r.text = tx; r.font.name = FONT
        r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c
    return box

def chrome_light(slide, title, subtitle=None):
    """Light slide chrome: white header area"""
    set_bg(slide, C_LIGHT)
    rect(slide, 0, 0, 13.33, 1.05, fill=C_WHITE)
    rect(slide, 0, 1.05, 13.33, 0.04, fill=C_ACC)
    txt(slide, title, 0.35, 0.08, 12.6, 0.58,
        size=22, bold=True, color=C_TX1)
    if subtitle:
        txt(slide, subtitle, 0.35, 0.65, 12.6, 0.35,
            size=11, color=C_TX3)

def chrome_dark(slide, title, subtitle=None):
    """Dark slide chrome"""
    set_bg(slide, C_DARK)
    rect(slide, 0, 0, 13.33, 1.0, fill=C_DARK2)
    rect(slide, 0, 1.0, 13.33, 0.04, fill=C_ACC)
    txt(slide, title, 0.35, 0.08, 12.6, 0.56,
        size=22, bold=True, color=C_WHITE)
    if subtitle:
        txt(slide, subtitle, 0.35, 0.66, 12.6, 0.3,
            size=11, color=C_ACC)

def kpi_light(slide, value, label1, label2, l, t, w=2.95, h=1.45,
              val_color=C_ACC, bg=C_WHITE):
    """KPI card on light slide"""
    rect(slide, l, t, w, h, fill=bg, line=C_LGRAY, lw=1)
    rect(slide, l, t, w, 0.04, fill=val_color)
    txt(slide, value, l+0.1, t+0.12, w-0.2, 0.62,
        size=32, bold=True, color=val_color, align=PP_ALIGN.CENTER)
    txt(slide, label1, l+0.1, t+0.76, w-0.2, 0.3,
        size=10, bold=False, color=C_TX1, align=PP_ALIGN.CENTER)
    if label2:
        txt(slide, label2, l+0.1, t+1.05, w-0.2, 0.3,
            size=10, color=C_TX3, align=PP_ALIGN.CENTER)

def kpi_dark(slide, value, label1, label2, l, t, w=2.95, h=1.45,
             val_color=C_ACC):
    """KPI card on dark slide"""
    rect(slide, l, t, w, h, fill=C_CARD_D)
    rect(slide, l, t, w, 0.04, fill=val_color)
    txt(slide, value, l+0.1, t+0.12, w-0.2, 0.62,
        size=32, bold=True, color=val_color, align=PP_ALIGN.CENTER)
    txt(slide, label1, l+0.1, t+0.76, w-0.2, 0.3,
        size=10, color=C_WHITE, align=PP_ALIGN.CENTER)
    if label2:
        txt(slide, label2, l+0.1, t+1.05, w-0.2, 0.3,
            size=10, color=C_TX3, align=PP_ALIGN.CENTER)

def section_divider(prs, num, title, desc):
    """Dark section divider slide — КТГ БЕЛАЗ style"""
    blank = prs.slide_layouts[6]
    s = prs.slides.add_slide(blank)
    set_bg(s, C_DARK)
    # Accent bar on left
    rect(s, 0, 0, 0.08, 7.5, fill=C_ACC)
    # Large section number
    txt(s, num, 0.3, 1.8, 3.0, 1.4,
        size=80, bold=True, color=C_WHITE)
    # Thin divider
    rect(s, 0.3, 3.4, 5.0, 0.05, fill=C_ACC)
    # Section title
    txt(s, title, 0.3, 3.55, 9.5, 1.0,
        size=38, bold=True, color=C_WHITE)
    # Description
    if desc:
        txt(s, desc, 0.3, 4.6, 9.5, 0.5,
            size=14, color=C_ACC)
    return s

def card_light(slide, title, items, l, t, w, h,
               title_color=C_ACC, bg=C_WHITE, bullet="▸"):
    rect(slide, l, t, w, h, fill=bg, line=C_LGRAY, lw=1)
    rect(slide, l, t, w, 0.04, fill=title_color)
    txt(slide, title, l+0.14, t+0.08, w-0.2, 0.3,
        size=12, bold=True, color=title_color)
    y = t + 0.42
    step = (h - 0.48) / max(len(items), 1)
    for item in items:
        c = item[1] if isinstance(item, tuple) else C_TX2
        s = item[0] if isinstance(item, tuple) else item
        txt(slide, f"{bullet} {s}", l+0.14, y, w-0.2, step+0.05,
            size=10, color=c)
        y += step

def table_ktg(slide, rows, col_w, l, t, w, h,
              hdr_bg=C_DARK, row_bg=C_WHITE, alt_bg=C_CARD_L,
              hdr_sz=10, cell_sz=9):
    from pptx.util import Pt as PT
    nr, nc = len(rows), len(rows[0])
    tbl = slide.shapes.add_table(nr, nc, Inches(l), Inches(t), Inches(w), Inches(h)).table
    total = sum(col_w)
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(w * cw / total)
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            if r == 0:    cell.fill.fore_color.rgb = hdr_bg
            elif r%2==1:  cell.fill.fore_color.rgb = row_bg
            else:         cell.fill.fore_color.rgb = alt_bg
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(cell_text)
            run.font.name = FONT
            run.font.size = PT(hdr_sz if r == 0 else cell_sz)
            run.font.bold = (r == 0)
            run.font.color.rgb = C_ACC if r == 0 else C_TX1

def pdf_img(slide, pdf, page=0, l=0.2, t=1.12, w=5.0, h=4.0, dpi=100):
    try:
        doc = fitz.open(pdf)
        if page >= len(doc): page = 0
        mat = fitz.Matrix(dpi/72, dpi/72)
        pix = doc[page].get_pixmap(matrix=mat)
        data = pix.tobytes("png")
        doc.close()
        slide.shapes.add_picture(io.BytesIO(data),
                                  Inches(l), Inches(t), Inches(w), Inches(h))
        return True
    except Exception as e:
        print(f"  [WARN] {pdf} p{page}: {e}")
        return False

# ═══════════════════════════════════════════════════════════════════════════════
def build():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    print("Building КТГ БЕЛАЗ styled presentation...")

    # ── SLIDE 1: TITLE ────────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    set_bg(s, C_DARK)
    rect(s, 0, 0, 0.08, 7.5, fill=C_ACC)
    # Decorative diagonal stripe
    rect(s, 9.5, 0, 0.06, 7.5, fill=C_CARD_D)
    rect(s, 11.5, 0, 0.04, 7.5, fill=C_CARD_D)
    # Logo area
    rect(s, 0.25, 0.25, 3.5, 0.55, fill=C_CARD_D)
    txt(s, "КТГ  •  АНАЛИЗ НЕИСПРАВНОСТЕЙ",
        0.35, 0.28, 3.3, 0.46,
        size=12, bold=True, color=C_ACC, align=PP_ALIGN.LEFT)
    # Title
    rect(s, 0.25, 1.1, 9.0, 0.06, fill=C_ACC)
    txt(s, "АНАЛИЗ ПРИЧИН ИЗНОСА",
        0.25, 1.25, 9.5, 1.1,
        size=46, bold=True, color=C_WHITE)
    txt(s, "ВЫПУСКНЫХ КЛАПАНОВ QSK50",
        0.25, 2.3, 9.5, 1.1,
        size=46, bold=True, color=C_ACC)
    rect(s, 0.25, 3.38, 9.0, 0.05, fill=C_CARD_D)
    txt(s, "Самосвалы NHL NTE200  •  АО «Полюс Магадан»  •  Омчак, Магаданская область",
        0.25, 3.5, 9.5, 0.5,
        size=14, color=C_TX3)
    # Stats
    stats = [
        ("20+",   "единиц парка",          C_RED),
        ("30+",   "ремонтных событий",      C_AMB),
        ("2 776", "м/ч — ранний отказ",    C_ACC),
        ("18 505","м/ч — поздний отказ",   C_GRN2),
    ]
    for i, (v, l2, col) in enumerate(stats):
        x = 0.25 + i * 2.24
        rect(s, x, 4.15, 2.1, 1.3, fill=C_CARD_D)
        rect(s, x, 4.15, 2.1, 0.05, fill=col)
        txt(s, v, x+0.08, 4.22, 1.94, 0.62,
            size=28, bold=True, color=col, align=PP_ALIGN.CENTER)
        txt(s, l2, x+0.06, 4.84, 1.98, 0.55,
            size=10, color=C_TX3, align=PP_ALIGN.CENTER)
    txt(s, "Конфиденциально  •  Только для внутреннего использования  •  Май 2026",
        0.25, 7.1, 12.9, 0.3,
        size=10, color=C_TX3, align=PP_ALIGN.CENTER)
    print("  01 Title")

    # ── SECTION 01 ────────────────────────────────────────────────────────────
    section_divider(prs, "01", "МАСШТАБ ПРОБЛЕМЫ",
                    "Статистика отказов выпускных клапанов | Парк NTE200")
    print("  02 Section 01")

    # ── SLIDE 3: SCALE ────────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "МАСШТАБ ПРОБЛЕМЫ",
                 "Статистика отказов выпускных клапанов QSK50 | Весь парк NTE200")
    kpis = [
        ("20+",   "агрегатов затронуто",   "",             C_RED),
        ("30+",   "ремонтных событий",      "",             C_AMB),
        ("2 776", "м/ч мин. наработка",    "до отказа",    C_ACC),
        ("100%",  "парка №43-83",           "с отказами",   C_BLU),
    ]
    for i, (v, l1, l2, col) in enumerate(kpis):
        kpi_light(s, v, l1, l2, 0.2 + i*3.28, 1.2, val_color=col)
    tbl = [
        ["Ед.", "Наработка до отказа", "Событий", "Поражение"],
        ["№43", "16 033 м/ч", "1", "5 позиций (3L, 4L, 1R, 2R, 3R)"],
        ["№44", "14 751 м/ч", "1", "ВСЕ 14 позиций"],
        ["№47", "12 890 м/ч", "2", "7 позиций + замена ГБЦ"],
        ["№48", "14 997 м/ч", "2", "17 клапанов + турбина + гидроудар"],
        ["№53", "17 550 м/ч", "1", "ВСЕ 16 позиций (2 выпуск + 2 впуск)"],
        ["№55", "17 316 м/ч", "1", "ВСЕ 16 позиций + 4 форсунки"],
        ["№57", "14 808 м/ч", "1", "6 позиций"],
        ["№69", "10 130 м/ч", "2", "Засорение воздушных фильтров + поршни 3R/4R"],
        ["№72", "10 000 м/ч", "2", "8 позиций (2 эпизода)"],
        ["№78", " 9 323 м/ч", "1", "3R / 4R / 6R / 6L"],
        ["№81", " 4 696 м/ч", "1", "ВСЕ позиции + замена ГБЦ"],
        ["№83", " 2 776 м/ч", "1", "КАТАСТРОФИЧЕСКИЙ — шатун / поршень / ГБЦ"],
    ]
    table_ktg(s, tbl, [0.8, 1.8, 1.0, 4.4], 0.2, 2.75, 13.0, 4.6)
    print("  03 Scale")

    # ── SLIDE 4: TIMELINE ─────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "ХРОНОЛОГИЯ ОТКАЗОВ",
                 "Авг 2025 — Май 2026 | 14 ключевых событий")
    events = [
        ("Авг 2025","№47","Замена ГБЦ — ранний эпизод",           "12 890", C_AMB),
        ("Авг 2025","№51","Замена ГБЦ, трещина головки",           "10 783", C_AMB),
        ("Авг 2025","№58","Замена толкателя клапана",               "10 702", C_AMB),
        ("Янв 2026","№43","Замена клапанов + 9 форсунок",          "16 033", C_BLU),
        ("Фев 2026","№48","17 клапанов + турбина",                  "14 997", C_RED),
        ("Мар 2026","№47","Повт. отказ — 10 клапанов + форсунки", "12 890", C_RED),
        ("Мар 2026","№58","Повт. отказ — клапан + ГБЦ",           "15 572", C_AMB),
        ("Мар 2026","№83","КАТАСТРОФА — шатун / поршень / ГБЦ",   " 2 776", C_RED),
        ("Апр 2026","№48","Гидроудар, водомасляная эмульсия",      "после рем.", C_RED),
        ("Май 2026","№55","ВСЕ 16 позиций + 4 форсунки",          "17 316", C_BLU),
        ("Май 2026","№69","Клапаны + засорение фильтров + поршни","10 721", C_AMB),
        ("Май 2026","№78","3R / 4R / 6R / 6L",                    " 9 323", C_AMB),
        ("Май 2026","№72","Отказ форсунки цил. 5",                 "10 283", C_TX3),
        ("Май 2026","№73","Форсунки — аномалии теста L2/R1",       "~9 998", C_TX3),
    ]
    rect(s, 0.2, 1.15, 13.0, 0.35, fill=C_DARK)
    for xoff, label in [(0.1,"ДАТА"),(1.2,"ЕД."),(2.4,"СОБЫТИЕ"),(10.1,"НАРАБОТКА")]:
        txt(s, label, 0.2+xoff, 1.17, 1.8, 0.28,
            size=9, bold=True, color=C_TX3)
    rh = 0.38
    for i, (dt, unit, ev, mh, col) in enumerate(events):
        y = 1.5 + i*rh
        bg = C_WHITE if i%2==0 else C_CARD_L
        rect(s, 0.2, y, 13.0, rh-0.02, fill=bg)
        rect(s, 0.2, y, 0.04, rh-0.02, fill=col)
        txt(s, dt,   0.32, y+0.04, 0.9, 0.28, size=10, color=C_TX3)
        txt(s, unit, 1.42, y+0.04, 1.0, 0.28, size=11, bold=True, color=col)
        txt(s, ev,   2.52, y+0.04, 7.5, 0.28, size=10, color=C_TX1)
        txt(s, mh,   10.1, y+0.04, 3.0, 0.28, size=10, bold=True, color=col,
            align=PP_ALIGN.RIGHT)
    print("  04 Timeline")

    # ── SECTION 02 ────────────────────────────────────────────────────────────
    section_divider(prs, "02", "РЕЗУЛЬТАТЫ РАЗБОРКИ",
                    "Физические признаки и лабораторный анализ NHL")
    print("  05 Section 02")

    # ── SLIDE 6: PHYSICAL FINDINGS ────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "ФИЗИЧЕСКИЕ ПРИЗНАКИ ПОВРЕЖДЕНИЙ",
                 "Данные разборки и дефектовки | Сводка по всем агрегатам")
    findings = [
        ("ПРОСАДКА ФАСКИ (РЕЦЕССИЯ КЛАПАНА)",
         "Главный признак. Посадочная поверхность выпускного клапана «утопает» в седле. "
         "Твёрдость фаски 40-44 HRC (базовая). Обнаружена на всех исследованных двигателях.", C_RED),
        ("ПЛАСТИЧЕСКАЯ ДЕФОРМАЦИЯ — «ЗАУСЕНЕЦ» НА ФАСКЕ",
         "Бурт пластической деформации = признак превышения допустимой T° тарелки. "
         "Зафиксировано в отчёте лаборатории NHL (MS&T2026033).", C_RED),
        ("ЗОЛЬНЫЕ ОТЛОЖЕНИЯ  Ca / Zn / P + Si / Al",
         "EDS-анализ NHL: Ca, Zn, P — зола ZDDP-присадок. Si, Al — силикатная пыль. "
         "ОТЯГЧАЮЩИЙ фактор (абразив). Не первопричина — следствие перегрева.", C_AMB),
        ("ПОВРЕЖДЕНИЕ СМЕЖНЫХ ЭЛЕМЕНТОВ",
         "№47/48/53/55/81: замена ГБЦ целиком. №83 (2 776 м/ч): разрушение шатуна и поршня. "
         "№48: гидроудар (водомасляная эмульсия) после первого ремонта.", C_AMB),
        ("ФОРСУНКИ",
         "Заменены совместно с клапанами: №43 (9 шт.), №47, №55 (4 шт.), №73. "
         "Протечка форсунки → дополнительный нагрев → перегрев клапана.", C_BLU),
        ("ТОПОГРАФИЯ ОТКАЗОВ",
         "Наиболее часто поражены: 3R, 4R, 6R, 6L, 7L, 7R. "
         "Потенциально — неравномерность охлаждения ГБЦ.", C_GRN2),
    ]
    y = 1.2
    for title, desc, col in findings:
        rect(s, 0.2, y, 13.0, 0.88, fill=C_WHITE, line=C_LGRAY, lw=1)
        rect(s, 0.2, y, 0.05, 0.88, fill=col)
        txt(s, title, 0.38, y+0.04, 12.7, 0.3, size=12, bold=True, color=col)
        txt(s, desc,  0.38, y+0.38, 12.7, 0.46, size=10, color=C_TX2)
        y += 0.93
    print("  06 Physical findings")

    # ── SLIDE 7: NHL LAB ──────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "АНАЛИЗ ЛАБОРАТОРИИ NHL  •  MS&T2026033",
                 "Инженер Тао Лан | СЭМ + EDS спектроскопия + профилометрия")
    nhl = os.path.join(WORK_DIR, "气门盘部磨损分析报告-1.docx RUS.pdf")
    pdf_img(s, nhl, page=2, l=0.2,  t=1.12, w=5.8, h=3.3)
    pdf_img(s, nhl, page=3, l=6.1,  t=1.12, w=5.8, h=3.3)
    rect(s, 0.2, 4.55, 12.9, 2.75, fill=C_DARK, line=None)
    rect(s, 0.2, 4.55, 12.9, 0.04, fill=C_ACC)
    cols_nhl = [
        ("МЕТОДЫ",        ["СЭМ (электронная микроскопия)", "EDS-спектроскопия", "Измерение твёрдости HRC", "Профилометрия"], C_ACC),
        ("EDS РЕЗУЛЬТАТ", ["Ca + Zn + P = зола ZDDP", "Si + Al = силикатная пыль", "Слои 5-50 мкм", "Абразивная среда в зоне контакта"], C_AMB),
        ("МОРФОЛОГИЯ",    ["Пластич. деформация фаски", "«Заусенец» = T° > лимита", "Адгезионный + абразивный износ", "Термическое повреждение"], C_RED),
        ("ВЫВОД NHL",     ["Комплексная причина:", "1. Термическое повреждение", "2. Абразивный износ (зола+пыль)", "3. Адгезионный износ"], C_BLU),
    ]
    for i, (ttl, items, col) in enumerate(cols_nhl):
        x = 0.32 + i * 3.2
        txt(s, ttl, x, 4.62, 3.0, 0.28, size=10, bold=True, color=col)
        for j, item in enumerate(items):
            txt(s, item, x, 4.95+j*0.52, 3.0, 0.48, size=10, color=C_TX3)
    print("  07 NHL lab")

    # ── SECTION 03 ────────────────────────────────────────────────────────────
    section_divider(prs, "03", "ГИПОТЕЗЫ О ПРИЧИНАХ",
                    "8 версий — от конструктивных особенностей до условий эксплуатации")
    print("  08 Section 03")

    # ── SLIDE 9: HYPOTHESES OVERVIEW ─────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "ОБЗОР ГИПОТЕЗ",
                 "8 версий — оценка доказательной базы и приоритет верификации")
    hyps = [
        ("H1", "Настройки ECM двигателя",
         "Отключённый дератинг, задержки 51/43 с, обороты 2100 vs 1990, TIB 200% vs 100%", C_RED, "ВЫСОКИЙ"),
        ("H2", "Отсутствие наплавки Stellite",
         "Фаска 40-44 HRC (NTE200) vs 55-65 HRC (Komatsu 730E AC). Разница твёрдости >50%", C_RED, "ВЫСОКИЙ"),
        ("H3", "Ошибка регулировки зазоров",
         "Малый зазор → клапан не закрывается → потеря охлаждения → перегрев фаски", C_AMB, "СРЕДНИЙ"),
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
    y = 1.2
    for hnum, htitle, hdesc, col, level in hyps:
        rect(s, 0.2, y, 1.05, 0.72, fill=col)
        txt(s, hnum, 0.2, y+0.1, 1.05, 0.54,
            size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        rect(s, 1.35, y, 9.3, 0.72, fill=C_WHITE, line=C_LGRAY, lw=1)
        txt(s, htitle, 1.48, y+0.04, 9.1, 0.3, size=12, bold=True, color=col)
        txt(s, hdesc,  1.48, y+0.36, 9.1, 0.32, size=10, color=C_TX2)
        lev_c = C_RED if level == "ВЫСОКИЙ" else (C_AMB if level == "СРЕДНИЙ" else C_BLU if level in ("НИЗКИЙ","ОТЯГЧ.") else C_TX3)
        rect(s, 10.75, y, 2.4, 0.72, fill=lev_c)
        txt(s, level, 10.75, y+0.2, 2.4, 0.36,
            size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        y += 0.77
    print("  09 Hypotheses overview")

    # ── SECTION 04 ────────────────────────────────────────────────────────────
    section_divider(prs, "04", "H1: НАСТРОЙКИ ECM",
                    "Сравнение NTE200 vs Komatsu 730E AC | Один объект — один двигатель")
    print("  10 Section 04")

    # ── SLIDE 11: ECM TABLE ───────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "H1: НАСТРОЙКИ ECM — NTE200 vs KOMATSU 730E AC",
                 "Один двигатель QSK50 — один объект — разные параметры защиты")
    rect(s, 0.2, 1.12, 12.9, 0.44, fill=RGBColor(0xFE, 0xE2, 0xE2))
    rect(s, 0.2, 1.12, 0.05, 0.44, fill=C_RED)
    txt(s,
        "КЛЮЧЕВОЙ ФАКТ: Komatsu 730E AC с аналогичными QSK50 работает НА ТОМ ЖЕ объекте. "
        "Случаев замены клапанов на 730E AC НЕ ЗАФИКСИРОВАНО.",
        0.35, 1.14, 12.6, 0.38, size=12, bold=True, color=C_RED)
    ecm = [
        ["Параметр ECM", "Komatsu 730E AC", "NHL NTE200", "Риск"],
        ["C_EPD_OT_RPM_Drt_En (флаг дератинга)", "ВКЛЮЧЁН (1)", "ОТКЛЮЧЁН (0)", "КРИТИЧЕСКИЙ"],
        ["Задержка по давлению масла", "Немедленно", "51 секунда (!)", "КРИТИЧЕСКИЙ"],
        ["Задержка по температуре ОЖ", "Немедленно", "43 секунды (!)", "ВЫСОКИЙ"],
        ["Максимальные обороты", "1 990 об/мин", "2 100 об/мин", "ВЫСОКИЙ"],
        ["Топливная коррекция TIB", "100%", "200% (!)", "ВЫСОКИЙ"],
        ["Триггеры по температуре ОГ", "Активны (все)", "ВСЕ ОТКЛЮЧЕНЫ", "КРИТИЧЕСКИЙ"],
        ["Дератинг при перегреве", "Активен", "Не работает", "КРИТИЧЕСКИЙ"],
        ["Обороты холостого хода", "700 об/мин", "700 об/мин", "Норма"],
    ]
    table_ktg(s, ecm, [4.2, 2.8, 2.8, 2.2], 0.2, 1.64, 13.0, 5.5)
    rect(s, 0.2, 7.22, 12.9, 0.22, fill=C_CARD_L)
    txt(s, "«КРИТИЧЕСКИЙ» — ведёт к перегреву без снижения мощности и без сигнала оператору",
        0.3, 7.23, 12.7, 0.18, size=10, bold=True, color=C_RED)
    print("  11 ECM table")

    # ── SLIDE 12: ECM MECHANISM ───────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "H1: ЦЕПОЧКА ПОВРЕЖДЕНИЙ ЧЕРЕЗ ECM",
                 "Как отключённый дератинг ведёт к отказу клапанов")
    steps = [
        ("ECM: Дератинг ОТКЛ\n+ задержки 43-51 с\n+ RPM 2100\n+ TIB 200%", C_RED),
        ("Двигатель работает\nбез снижения мощности\nпри перегреве\nи аварии", C_AMB),
        ("T° выхлопных газов\nвыше нормы\nбез сигнала\nоператору", C_AMB),
        ("ПЕРЕГРЕВ тарелки\nклапана → деформация\nфаски → рецессия", C_RED),
    ]
    for i, (text, col) in enumerate(steps):
        x = 0.2 + i*3.3
        rect(s, x, 1.18, 3.1, 1.8, fill=C_WHITE, line=col, lw=2)
        rect(s, x, 1.18, 3.1, 0.05, fill=col)
        txt(s, text, x+0.1, 1.3, 2.9, 1.6,
            size=13, color=col, align=PP_ALIGN.CENTER)
        if i < 3:
            txt(s, "→", x+3.1, 1.92, 0.2, 0.38,
                size=22, bold=True, color=C_TX3, align=PP_ALIGN.CENTER)
    rect(s, 0.2, 3.16, 12.9, 0.72, fill=RGBColor(0xFE, 0xE2, 0xE2), line=C_RED, lw=1)
    txt(s,
        "РЕЗУЛЬТАТ: Рецессия фаски → потеря компрессии → прорыв горячих газов → "
        "ускоренный абразивный износ → последовательный выход из строя клапанов",
        0.35, 3.2, 12.6, 0.64, size=13, bold=True, color=C_RED)
    rect(s, 0.2, 4.02, 12.9, 3.28, fill=C_WHITE, line=C_LGRAY, lw=1)
    txt(s, "КОСВЕННЫЕ ДОКАЗАТЕЛЬСТВА:", 0.35, 4.08, 12.5, 0.3,
        size=12, bold=True, color=C_ACC)
    evidence = [
        "▸  NTE200: операторов инструктируют «не глушить, давать остывать на ХХ» — косвенное признание перегревов",
        "▸  INSITE CSV (агрегат N43, 18 523 м/ч): зафиксированы температурные отклонения в рабочем цикле",
        "▸  730E AC на том же объекте, с теми же двигателями QSK50 — клапанных отказов НЕТ",
        "▸  Самые ранние отказы (№83: 2 776 м/ч | №81: 4 696 м/ч) не объяснимы обычным усталостным износом",
        "▸  Диапазон наработок (2 776 — 18 505 м/ч) указывает на вариабельный триггер, а не дефект партии",
    ]
    for i, e in enumerate(evidence):
        txt(s, e, 0.35, 4.45+i*0.52, 12.6, 0.48, size=11, color=C_TX2)
    print("  12 ECM mechanism")

    # ── SLIDE 13: H2 STELLITE ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "H2: ОТСУТСТВИЕ НАПЛАВКИ STELLITE",
                 "Конструктивное различие фаски клапана NTE200 vs Komatsu 730E AC")
    rect(s, 0.2, 1.12, 6.4, 5.8, fill=C_WHITE, line=C_RED, lw=1)
    rect(s, 0.2, 1.12, 6.4, 0.05, fill=C_RED)
    txt(s, "NTE200 / QSK50 (CES51005)", 0.3, 1.2, 6.1, 0.38, size=14, bold=True, color=C_RED)
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
        txt(s, lbl, 0.3, y, 2.4, 0.38, size=11, bold=True, color=C_TX2)
        c = C_RED if "ОТСУТСТВУЕТ" in val or "ВСЕХ" in val else C_TX1
        txt(s, val, 2.75, y, 3.75, 0.38, size=11, color=c)
        y += 0.52
    rect(s, 6.75, 1.12, 6.4, 5.8, fill=C_WHITE, line=C_ACC, lw=1)
    rect(s, 6.75, 1.12, 6.4, 0.05, fill=C_ACC)
    txt(s, "Komatsu 730E AC / QSK50", 6.85, 1.2, 6.1, 0.38, size=14, bold=True, color=C_GRN2)
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
        txt(s, lbl, 6.85, y, 2.4, 0.38, size=11, bold=True, color=C_TX2)
        c = C_GRN2 if "STELLITE" in val or "не зафиксировано" in val.lower() else C_TX1
        txt(s, val, 9.3, y, 3.75, 0.38, size=11, color=c)
        y += 0.52
    rect(s, 0.2, 7.0, 12.9, 0.4, fill=C_CARD_L)
    txt(s,
        "Stellite снижает износ в 3-5× при высоких температурах. "
        "Применение Stellite-клапанов в NTE200 — конкретная конструктивная рекомендация.",
        0.3, 7.02, 12.7, 0.34, size=11, bold=True, color=C_GRN2)
    print("  13 H2 Stellite")

    # ── SECTION 05 ────────────────────────────────────────────────────────────
    section_divider(prs, "05", "КЛЮЧЕВЫЕ СЛУЧАИ",
                    "Агрегаты №83, №48, №55, №69")
    print("  14 Section 05")

    # ── SLIDE 15: UNIT 83 ─────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    set_bg(s, C_DARK)
    add_rect = rect  # alias
    rect(s, 0, 0, 0.08, 7.5, fill=C_RED)
    rect(s, 0, 0, 13.33, 0.95, fill=C_DARK2)
    rect(s, 0, 0.95, 13.33, 0.04, fill=C_RED)
    txt(s, "КАТАСТРОФИЧЕСКИЙ ОТКАЗ  •  АГРЕГАТ №83",
        0.25, 0.08, 12.6, 0.56, size=22, bold=True, color=C_WHITE)
    txt(s, "2 776 м/ч наработки — самый ранний отказ в парке",
        0.25, 0.64, 12.6, 0.28, size=11, color=C_TX3)
    pdf83 = os.path.join(WORK_DIR, "Тех отчет 83 от 30.03.2026.pdf")
    pdf_img(s, pdf83, page=0, l=0.2, t=1.1, w=4.1, h=3.1)
    pdf_img(s, pdf83, page=2, l=4.4, t=1.1, w=4.1, h=3.1)
    pdf_img(s, pdf83, page=4, l=8.7, t=1.1, w=4.4, h=3.1)
    rect(s, 0.2, 4.35, 12.9, 2.95, fill=C_CARD_D)
    rect(s, 0.2, 4.35, 0.05, 2.95, fill=C_RED)
    data83 = [
        ("Гаражный №:", "83"),
        ("Наработка:", "2 776 м/ч  (!!! нормальный ресурс >12 000 м/ч)"),
        ("Зафиксировано:", "Разрушение шатуна и поршня"),
        ("Замена:", "ГБЦ + шатун + поршень + уплотнения"),
        ("Вер. причина:", "Перегрев → потеря плотности клапана → прорыв газов → разрушение ЦПГ"),
        ("Значение:", "Катастрофический отказ при минимальной наработке — системная проблема"),
    ]
    y = 4.42
    for lbl, val in data83:
        txt(s, lbl, 0.38, y, 2.4, 0.4, size=11, bold=True, color=C_AMB)
        txt(s, val, 2.88, y, 10.1, 0.4, size=11, color=C_WHITE)
        y += 0.44
    print("  15 Unit 83")

    # ── SLIDE 16: MATRIX ──────────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "МАТРИЦА ОЦЕНКИ ГИПОТЕЗ",
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
    table_ktg(s, matrix, [3.8, 1.9, 2.0, 1.3, 1.5, 1.5],
              0.2, 1.12, 12.9, 5.9,
              hdr_bg=C_DARK, row_bg=C_WHITE, alt_bg=C_CARD_L,
              hdr_sz=11, cell_sz=10)
    rect(s, 0.2, 7.1, 12.9, 0.3, fill=C_CARD_L)
    txt(s, "H1 + H2 = СИСТЕМНЫЕ причины  •  H3 + H4 = СОПУТСТВУЮЩИЕ  •  H6 = ОТЯГЧАЮЩИЙ фактор",
        0.3, 7.12, 12.7, 0.24, size=11, bold=True, color=C_ACC)
    print("  16 Matrix")

    # ── SECTION 06 ────────────────────────────────────────────────────────────
    section_divider(prs, "06", "ПЛАН ВЕРИФИКАЦИИ",
                    "Конкретные действия для подтверждения каждой гипотезы")
    print("  17 Section 06")

    # ── SLIDE 18: VERIFICATION PLAN ───────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "ПЛАН ВЕРИФИКАЦИИ ГИПОТЕЗ",
                 "Конкретные действия для подтверждения / опровержения каждой версии")
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
            "Разница ≥50% vs базовый клапан той же наработки = подтверждено",
        ]),
        ("H3", "ЗАЗОРЫ КЛАПАНОВ", C_BLU, [
            "Ввести протокол: измерение зазоров ДО и ПОСЛЕ каждого ремонта с записью",
            "При ближайшем ТО проверить зазоры на 5 агрегатах без клапанного ремонта",
            "Калиброванный щуп 0.05 мм шаг | допуск ±0.05 мм | обучение механиков",
        ]),
        ("H4", "ФОРСУНКИ", C_GRN2, [
            "Стендовые тесты: №43, №47, №55, №73 — давление, возврат, распыл",
            "INSITE: fuel balance rates → отклонение >±3% = проблемная форсунка",
            "Правило: при клапанном ремонте — обязательная проверка форсунки цилиндра",
        ]),
    ]
    y = 1.12
    for hnum, htitle, col, actions in vplan:
        h = len(actions) * 0.4 + 0.5
        rect(s, 0.2, y, 12.9, h, fill=C_WHITE, line=C_LGRAY, lw=1)
        rect(s, 0.2, y, 0.05, h, fill=col)
        rect(s, 0.25, y, 1.45, h, fill=C_CARD_L)
        txt(s, hnum, 0.27, y+0.06, 1.38, 0.36,
            size=18, bold=True, color=col, align=PP_ALIGN.CENTER)
        txt(s, htitle, 1.82, y+0.06, 11.2, 0.32, size=12, bold=True, color=col)
        for j, action in enumerate(actions):
            txt(s, "→  " + action, 1.82, y+0.46+j*0.38, 11.2, 0.34,
                size=10, color=C_TX2)
        y += h + 0.08
    print("  18 Verification plan")

    # ── SLIDE 19: RECOMMENDATIONS ─────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    chrome_light(s, "РЕКОМЕНДАЦИИ",
                 "Приоритетные действия по устранению причин отказов клапанов QSK50")
    recs = [
        ("СРОЧНО",       "до следующего ремонта",  C_RED, [
            "Перепрограммировать ECM: включить дератинг, задержки ≤10 с, RPM 1990, TIB 100%",
            "Активировать все триггеры температуры выхлопных газов",
            "Проверить состояние воздушных фильтров всего парка",
            "Обязательная проверка форсунки при каждом клапанном ремонте",
        ]),
        ("КРАТКОСРОЧНО", "в течение 1 месяца",     C_AMB, [
            "Заказать Stellite-клапаны CES51005-S — пилотная партия на 3 агрегата",
            "Разработать и внедрить форму контроля зазоров клапанов",
            "Стендовые тесты форсунок приоритетных агрегатов",
            "Включить логирование INSITE на всех агрегатах",
        ]),
        ("СРЕДНЕСРОЧНО", "в течение 3 месяцев",    C_BLU, [
            "Анализ рабочего цикла по INSITE на 5+ агрегатах",
            "Сравнительный анализ с 730E AC на том же объекте",
            "По результатам пилота Stellite: решение о серийном переходе",
        ]),
        ("СИСТЕМНО",     "постоянные меры",         C_ACC, [
            "Предложить NHL внесение Stellite как обязательного стандарта",
            "База данных ТО: зазоры + форсунки + T° — на каждый агрегат",
        ]),
    ]
    y = 1.12
    for urgency, sub, col, items in recs:
        h = len(items) * 0.42 + 0.56
        rect(s, 0.2, y, 12.9, h, fill=C_WHITE, line=C_LGRAY, lw=1)
        rect(s, 0.2, y, 0.05, h, fill=col)
        rect(s, 0.25, y, 2.0, h, fill=C_CARD_L)
        txt(s, urgency, 0.28, y+0.1, 1.9, 0.36,
            size=13, bold=True, color=col, align=PP_ALIGN.CENTER)
        txt(s, sub, 0.28, y+0.46, 1.9, 0.24, size=9, color=C_TX3, align=PP_ALIGN.CENTER)
        for j, item in enumerate(items):
            txt(s, "✓  " + item, 2.38, y+0.1+j*0.42, 10.6, 0.36, size=11, color=C_TX2)
        y += h + 0.08
    print("  19 Recommendations")

    # ── SLIDE 20: CONCLUSIONS ─────────────────────────────────────────────────
    s = prs.slides.add_slide(blank)
    set_bg(s, C_DARK)
    rect(s, 0, 0, 0.08, 7.5, fill=C_ACC)
    rect(s, 0, 0, 13.33, 0.95, fill=C_DARK2)
    rect(s, 0, 0.95, 13.33, 0.04, fill=C_ACC)
    txt(s, "ВЫВОДЫ И СЛЕДУЮЩИЕ ШАГИ",
        0.25, 0.08, 12.6, 0.56, size=22, bold=True, color=C_WHITE)
    txt(s, "Анализ причин износа выпускных клапанов QSK50 | NTE200 | АО «Полюс Магадан»",
        0.25, 0.64, 12.6, 0.28, size=11, color=C_TX3)
    concs = [
        ("1", "МАСШТАБ СИСТЕМНЫЙ",
         "20+ единиц, 30+ событий, 2 776 — 18 505 м/ч. Не единичный дефект партии.", C_RED),
        ("2", "ECM — ГЛАВНАЯ РАБОЧАЯ ВЕРСИЯ",
         "Отключённый дератинг + задержки + RPM 2100 + TIB 200%. 730E AC (тот же объект) — 0 отказов.", C_RED),
        ("3", "STELLITE — КОНСТРУКТИВНАЯ УЯЗВИМОСТЬ",
         "40-44 HRC против 55-65 HRC. Stellite сохраняет твёрдость при 800°C — критично для клапана.", C_AMB),
        ("4", "ЗОЛА — ОТЯГЧАЮЩИЙ ФАКТОР",
         "EDS NHL: Ca/Zn/P + Si/Al. Ускоряет износ, но является следствием перегрева.", C_BLU),
        ("5", "СЛЕДУЮЩИЙ ШАГ — ВЕРИФИКАЦИЯ",
         "Перепрограммировать ECM (быстрый тест) + пилотные Stellite-клапаны на 2-3 агрегата.", C_ACC),
    ]
    y = 1.08
    for num, title, desc, col in concs:
        rect(s, 0.2, y, 12.9, 1.08, fill=C_CARD_D)
        rect(s, 0.2, y, 0.05, 1.08, fill=col)
        rect(s, 0.25, y, 0.95, 1.08, fill=C_DARK)
        txt(s, num, 0.27, y+0.22, 0.9, 0.62,
            size=26, bold=True, color=col, align=PP_ALIGN.CENTER)
        txt(s, title, 1.32, y+0.08, 11.5, 0.36, size=13, bold=True, color=col)
        txt(s, desc,  1.32, y+0.5, 11.5, 0.5, size=11, color=C_TX3)
        y += 1.12
    rect(s, 0.2, 6.85, 12.9, 0.55, fill=C_DARK2)
    txt(s, "КТГ — Технический анализ  •  АО «Полюс Магадан»  •  Конфиденциально  •  Май 2026",
        0.2, 6.88, 12.9, 0.26, size=11, color=C_TX3, align=PP_ALIGN.CENTER)
    txt(s, "Cummins QSK50  •  NHL NTE200",
        0.2, 7.18, 12.9, 0.22, size=10, color=C_ACC, align=PP_ALIGN.CENTER)
    print("  20 Conclusions")

    # ── SAVE ─────────────────────────────────────────────────────────────────
    out = os.path.join(WORK_DIR, "Презентация_анализ_клапанов_QSK50_NTE200.pptx")
    prs.save(out)
    sz = os.path.getsize(out)
    print(f"\nSaved: {out}")
    print(f"Size:  {sz/1024/1024:.1f} MB  |  {len(prs.slides)} slides")
    return out

if __name__ == "__main__":
    build()
