#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build Part 3: Pages 21-30 of GBC analysis report
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import warnings
warnings.filterwarnings('ignore')

BG = '#FAFBFC'
ACCENT = '#007A5E'
TEXT = '#1A2530'
PANEL = '#EBF0F4'
RED = '#C0392B'
YELLOW = '#B87800'
GRAY = '#5A6B75'

def setup_page(fig):
    fig.patch.set_facecolor(BG)

def add_header(fig, page_num):
    ax_h = fig.add_axes([0, 0.93, 1, 0.07])
    ax_h.set_facecolor(PANEL)
    ax_h.set_xlim(0, 1)
    ax_h.set_ylim(0, 1)
    ax_h.axis('off')
    ax_h.plot([0, 1], [0, 0], color=ACCENT, linewidth=2)
    ax_h.text(0.02, 0.5, 'АО РАЗВИТИЕ  |  Надежность оборудования', color=ACCENT,
              fontsize=7, va='center', fontweight='bold')
    ax_h.text(0.5, 0.5,
              'АНАЛИЗ КОРЕННЫХ ПРИЧИН РАЗРУШЕНИЯ КУ ДВИГАТЕЛЯ CUMMINS QSK50 MCRS. NTE200. ПОЛЮС МАГАДАН',
              color=TEXT, fontsize=6.5, va='center', ha='center')
    ax_h.text(0.98, 0.5, f'Стр. {page_num}', color=GRAY, fontsize=7, va='center', ha='right')

def add_footer(fig):
    ax_f = fig.add_axes([0, 0, 1, 0.04])
    ax_f.set_facecolor(PANEL)
    ax_f.set_xlim(0, 1)
    ax_f.set_ylim(0, 1)
    ax_f.axis('off')
    ax_f.plot([0, 1], [1, 1], color=ACCENT, linewidth=1, alpha=0.5)
    ax_f.text(0.02, 0.4, 'Конфиденциально. АО Развитие. 2026 г.', color=GRAY, fontsize=6, va='center')
    ax_f.text(0.98, 0.4, 'Версия 1.0 | Июнь 2026', color=GRAY, fontsize=6, va='center', ha='right')

def make_fig():
    fig = plt.figure(figsize=(8.27, 11.69))
    setup_page(fig)
    return fig

def tbl(ax, headers, rows, y_start, col_x, col_w, row_h=0.038):
    rect = mpatches.FancyBboxPatch((col_x[0]-0.01, y_start - row_h), sum(col_w)+0.02, row_h,
                                    boxstyle="square,pad=0", facecolor=ACCENT, edgecolor='none')
    ax.add_patch(rect)
    for j, (h, cx, cw) in enumerate(zip(headers, col_x, col_w)):
        ax.text(cx+cw/2, y_start - row_h/2, h, color=BG, fontsize=7, ha='center', va='center', fontweight='bold')
    y = y_start - row_h
    for i, row in enumerate(rows):
        bg = PANEL if i % 2 == 0 else BG
        rect2 = mpatches.FancyBboxPatch((col_x[0]-0.01, y-row_h), sum(col_w)+0.02, row_h,
                                         boxstyle="square,pad=0", facecolor=bg, edgecolor='none')
        ax.add_patch(rect2)
        for j, (val, cx, cw) in enumerate(zip(row, col_x, col_w)):
            is_a = str(val).startswith('!')
            val2 = str(val).lstrip('!')
            clr = RED if is_a else (ACCENT if j == 0 else TEXT)
            ax.text(cx+cw/2 if j > 0 else cx+0.005, y-row_h/2, val2,
                    color=clr, fontsize=7, ha='center' if j > 0 else 'left', va='center')
        y -= row_h
    return y

# ============================================================
# PAGE 21: THERMAL FAILURE MECHANISM
# ============================================================
def page_thermal_mech(pdf):
    fig = make_fig()
    add_header(fig, 21)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ТЕПЛОВОЙ МЕХАНИЗМ ОТКАЗА — ЦЕПОЧКА ТЕПЛОПЕРЕДАЧИ', color=ACCENT, fontsize=12, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    ax.text(0, 0.91, 'Клапан QSK50 MCRS (Inconel 751) отводит тепло двумя путями: через стержень→направляющую '
            'и через тарелку→седло (60-70% от общего теплового потока при нормальной работе).',
            color=TEXT, fontsize=8.5, va='top')

    # Heat transfer chain diagram
    chain = [
        (0.5, 0.82, 'Топливо + воздух', 'Камера сгорания\nT_газов = 1900-2200°C\n(норм. цикл)', GRAY),
        (0.5, 0.73, 'ТЕПЛОВОЙ ПОТОК → ТАРЕЛКА КЛАПАНА', 'NTE200 (TIB=200%, RPM=2100):\nПиковые T_газов локально >2400°C\n(избыточный впрыск)', RED),
        (0.5, 0.64, 'Температура тарелки клапана', 'Норма QSK50: 420-480°C\nFact NTE200 (расч.): 650-720°C\n(на основе EGT + тепловой модели)', RED),
        (0.5, 0.55, 'Теплоотвод через седло→ГБЦ', 'Зависит от: площади контакта, давления,\nтеплопроводности. '
         'При деформации тарелки —\nплощадь контакта снижается → теплоотвод падает → T растёт', YELLOW),
        (0.5, 0.46, 'Начало деградации Inconel 751', 'При T > 650°C: снижение предела ползучести\n'
         'При T > 700°C: пластическая деформация тарелки\n(Сред. скорость ползучести по CES51005: 0,1 мм/1 000 ч при 700°C)', RED),
        (0.5, 0.37, 'Нарушение прилегания тарелки к седлу', 'Деформированная тарелка → щель → '
         'утечка горячих газов при закрытии\nудар газовой струи по краю тарелки → задир, эрозия', RED),
        (0.5, 0.28, 'Положительная обратная связь', 'Хуже прилегание → хуже теплоотвод → '
         'выше температура → больше деформация\n(лавинный процесс — характерен для CCEC ESN 33232926: '
         '15 126 ч до разрушения)', RED),
        (0.5, 0.19, 'КОНЕЧНЫЙ ОТКАЗ', 'Прогар тарелки / разрушение седла /\nпробой газов в систему охлаждения / '
         'удар по поршню', '#FDEAEA'),
    ]

    for i, (x, y, title, body, bg_col) in enumerate(chain):
        rect = mpatches.FancyBboxPatch((0.02, y - 0.06), 0.96, 0.07,
                                        boxstyle="round,pad=0.005", facecolor=bg_col if bg_col in [GRAY, PANEL] else PANEL,
                                        edgecolor=bg_col if bg_col != PANEL else GRAY, linewidth=1.5)
        ax.add_patch(rect)
        if bg_col == '#FDEAEA':
            rect2 = mpatches.FancyBboxPatch((0.02, y - 0.06), 0.96, 0.07,
                                             boxstyle="round,pad=0.005", facecolor='#FDEAEA', edgecolor=RED, linewidth=2)
            ax.add_patch(rect2)
        ax.text(0.05, y - 0.01, title, color=ACCENT if i in [0] else RED if bg_col == RED or bg_col == '#FDEAEA' else YELLOW,
                fontsize=8, fontweight='bold', va='top')
        ax.text(0.40, y - 0.01, body, color=TEXT if bg_col != '#FDEAEA' else RED,
                fontsize=7, va='top', linespacing=1.3)
        if i < len(chain) - 1:
            ax.annotate('', xy=(0.5, chain[i+1][1] + 0.01), xytext=(0.5, y - 0.06),
                        arrowprops=dict(arrowstyle='->', color=ACCENT, lw=1.5))

    ax.text(0.5, 0.10, 'Ключевое отличие NTE200 от 730E: на 730E механизм останавливается на шаге 2\n'
            '(защита FC, DO-сигналы, TIB=100%). На NTE200 — ни одного барьера не сработало.',
            color=ACCENT, fontsize=8.5, ha='center', va='top', fontweight='bold')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 21 done")

# ============================================================
# PAGE 22: 730E VS NTE200 COMPARISON
# ============================================================
def page_comparison(pdf):
    fig = make_fig()
    add_header(fig, 22)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'СРАВНЕНИЕ УСЛОВИЙ ЭКСПЛУАТАЦИИ 730E И NTE200', color=ACCENT, fontsize=12, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    ax.text(0.5, 0.906,
            'Критерий разграничения: если фактор ОДИНАКОВ для обоих типов — он не может быть первопричиной.',
            color=YELLOW, fontsize=8.5, ha='center', fontweight='bold')

    headers = ['Фактор', '730E-DC', 'NTE200', 'Одинаков?', 'Роль в отказе']
    rows = [
        ('Предприятие', 'АО Полюс Магадан', 'АО Полюс Магадан', 'ДА', 'Исключён'),
        ('Карьер / маршруты', 'Тот же карьер', 'Тот же карьер', 'ДА', 'Исключён'),
        ('Дизельное топливо', 'ГОСТ Р 52368, Арктика', 'ГОСТ Р 52368, Арктика', 'ДА', 'Исключён'),
        ('Масло CI-4', 'SAE 15W-40, CES57000', 'SAE 15W-40, CES57000', 'ДА', 'Исключён'),
        ('Пыль, климат', 'Магаданская обл.', 'Магаданская обл.', 'ДА', 'Исключён'),
        ('Двигатель (модель)', 'QSK50 MCRS V16', 'QSK50 MCRS V16', 'ДА', 'Исключён'),
        ('Конструкция ГБЦ', 'Cummins OEM', 'Cummins OEM', 'ДА', 'Исключён'),
        ('Марка масла (бренд)', 'Разные поставщики', 'Разные поставщики', 'Практич. да', 'Исключён'),
        ('Качество ТО (порядок)', 'Одна служба', 'Одна служба', 'ДА', 'Исключён'),
        ('Калибровка ECM', 'AQ60217.28', '!AQ60809.08', '!НЕТ', '!ОТЛИЧАЕТСЯ — ПЕРВОПРИЧИНА'),
        ('TIB Fuel Limit', '100%', '!200%', '!НЕТ', '!Двойная доза — высокий риск'),
        ('Max RPM (уставка)', '1 990 RPM', '!2 100 RPM', '!НЕТ', '!+5,5% — рост EGT и нагрузок'),
        ('DO-выходы защиты', '5 каналов АКТИВНЫ', '!ВСЕ ОТКЛЮЧЕНЫ', '!НЕТ', '!Оператор не информирован'),
        ('Engine Protection', 'FC146, FC165, FC556 — работает', '!«No data available»', '!НЕТ', '!Защита деактивирована'),
        ('Факты массовых отказов ГБЦ', '0 случаев за 2023-2026', '!24 машины (48%)', '!НЕТ', '!РЕЗУЛЬТИРУЮЩИЙ ОТКАЗ'),
    ]
    col_x = [0, 0.22, 0.38, 0.54, 0.62]
    col_w = [0.22, 0.16, 0.16, 0.08, 0.38]
    tbl(ax, headers, rows, 0.90, col_x, col_w, row_h=0.048)

    ax.text(0.5, 0.07,
            'ВЫВОД: Единственный системный фактор, различающийся между 730E и NTE200 при '
            'одинаковых внешних условиях — КАЛИБРОВКА ECM (3 критических параметра). '
            'По принципу исключения — это доказанная первопричина (Level 1 RCA, метод "5 Почему").',
            color=ACCENT, fontsize=8.5, ha='center', va='bottom', fontweight='bold')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 22 done")

# ============================================================
# PAGE 23: FAILURE MODE CLASSIFICATION
# ============================================================
def page_failure_modes(pdf):
    fig = make_fig()
    add_header(fig, 23)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'КЛАССИФИКАЦИЯ РЕЖИМОВ ОТКАЗА', color=ACCENT, fontsize=13, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    # Pie chart - failure types
    ax_pie = fig.add_axes([0.55, 0.55, 0.40, 0.35])
    ax_pie.set_facecolor(BG)

    labels = ['Износ клапана\n(1-3 ГБЦ)', 'Групповой износ\n(4-9 ГБЦ)', 'Полный\n(все 16 ГБЦ)', 'Катастроф.\n(ГБЦ+поршень)']
    sizes = [10, 9, 3, 2]
    colors_pie = [ACCENT, YELLOW, RED, '#8B0000']
    wedges, texts, autotexts = ax_pie.pie(sizes, labels=labels, colors=colors_pie, autopct='%1.0f%%',
                                           startangle=90, textprops={'color': TEXT, 'fontsize': 7.5})
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color(BG)
        at.set_fontweight('bold')
    ax_pie.set_title('Распределение по\nтипу отказа', color=TEXT, fontsize=8.5, pad=4)

    # FMEA table
    headers = ['Режим отказа', 'Кол-во', 'Мех-зм', 'Тяжесть', 'Повторяем.', 'RPN*']
    rows = [
        ('Одиночный износ вып. клапана (1 ГБЦ)', '10 ед.', 'Тепловой', '7', '3', '147'),
        ('Групповой отказ 4-9 ГБЦ (1 банк/весь)', '9 ед.', 'Тепловой', '8', '4', '256'),
        ('Полный отказ — все 16 ГБЦ', '3 ед.', 'Тепловой+усталость', '9', '2', '144'),
        ('Износ впускного клапана (NTE#53,55)', '2 ед.', 'Обратный тепловой поток', '8', '2', '128'),
        ('Катастрофический (ГБЦ+шатун+поршень)', '2 ед.', 'Газодин. удар', '10', '1', '100'),
        ('Повторный отказ после ремонта <1 500 м/ч', '8 ед.', 'Системный (ECM не изменён)', '9', '5', '360'),
    ]
    col_x = [0, 0.38, 0.52, 0.62, 0.70, 0.80]
    col_w = [0.38, 0.14, 0.10, 0.08, 0.10, 0.20]
    tbl(ax, headers, rows, 0.90, col_x, col_w, row_h=0.06)

    ax.text(0, 0.52, '* RPN = Тяжесть × Повторяемость × Обнаруживаемость (1-10). Порог: RPN > 100 = действие обязательно.', color=GRAY, fontsize=7)

    # Mode descriptions
    ax.text(0, 0.49, 'ХАРАКТЕРИСТИКА РЕЖИМОВ ОТКАЗА:', color=ACCENT, fontsize=9.5, fontweight='bold')

    modes = [
        ('Одиночный (10 случаев)', ACCENT,
         'Наиболее распространённый сценарий. Один или два выпускных клапана на 1-3 ГБЦ. '
         'Типичная наработка: 10 000–17 000 м/ч. Характерный признак: задир контактной '
         'поверхности тарелки, металлоперенос. EDS показывает золу масла как осадок (следствие).'),
        ('Групповой (9 случаев)', YELLOW,
         'Одновременно 4-9 ГБЦ. Поражение, как правило, одного банка или обоих. '
         'NTE#47 (12 890 м/ч): 7 ГБЦ за одно событие. NTE#77 (10 016 м/ч): 6 ГБЦ. '
         'Признак: однородный характер повреждений, схожие часы до отказа для всех поражённых цилиндров.'),
        ('Полный (2+1 случая)', RED,
         'NTE#53 (17 550 м/ч) и NTE#55 (17 316 м/ч): ВСЕ 16 ГБЦ — каждая с 2 вып. + 2 впуск. '
         'клапанами (итого 64 клапана на машину). Вовлечение впускных клапанов — признак '
         'экстремально высокой тепловой нагрузки и обратного теплового потока через газообмен.'),
        ('Катастрофический (2 случая)', RED,
         'NTE#83 (2 776 м/ч): ГБЦ + шатун + поршень. NTE#48 (14 997 м/ч): ГБЦ + поршень. '
         'Механизм: прогар клапана → газовый удар в поршневую → разрушение. '
         'Критически ранний отказ NTE#83 на третьем тысячелетии м/ч — системный дефект с нуля.'),
    ]

    y = 0.46
    for title, color, text in modes:
        rect = mpatches.FancyBboxPatch((-0.01, y-0.10), 1.02, 0.10,
                                        boxstyle="round,pad=0.005", facecolor=PANEL, edgecolor='none')
        ax.add_patch(rect)
        ax.plot([-0.01, -0.01], [y-0.095, y-0.005], color=color, linewidth=4)
        ax.text(0.02, y - 0.008, title, color=color, fontsize=8.5, fontweight='bold', va='top')
        words = text.split()
        mid = len(words) * 3 // 5
        ax.text(0.02, y - 0.040, ' '.join(words[:mid]), color=TEXT, fontsize=7.3, va='top')
        ax.text(0.02, y - 0.063, ' '.join(words[mid:]), color=TEXT, fontsize=7.3, va='top')
        y -= 0.115

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 23 done")

# ============================================================
# PAGE 24: ROOT CAUSE TREE (FTA)
# ============================================================
def page_fta(pdf):
    fig = make_fig()
    add_header(fig, 24)
    add_footer(fig)

    ax = fig.add_axes([0.02, 0.07, 0.96, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ДЕРЕВО ПРИЧИН ОТКАЗА (Fault Tree Analysis — FTA)', color=ACCENT, fontsize=12, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    def node(ax, x, y, w, h, text, color=PANEL, border=ACCENT, fontsize=7.5, text_color=TEXT):
        rect = mpatches.FancyBboxPatch((x-w/2, y-h/2), w, h,
                                        boxstyle="round,pad=0.008", facecolor=color, edgecolor=border, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x, y, text, color=text_color, fontsize=fontsize, ha='center', va='center', linespacing=1.3)

    def line(ax, x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color=GRAY, linewidth=1.2, alpha=0.8)

    # Top event
    node(ax, 0.5, 0.88, 0.55, 0.055,
         'РАЗРУШЕНИЕ КЛАПАННОГО УЗЛА ГБЦ\nQSK50 MCRS — NTE200 (24 машины)', '#FDEAEA', RED, 8.5, RED)

    # Level 1: OR gate
    node(ax, 0.5, 0.79, 0.12, 0.035, 'ИЛИ', PANEL, YELLOW, 8, YELLOW)
    line(ax, 0.5, 0.855, 0.5, 0.808)
    line(ax, 0.5, 0.773, 0.5, 0.755)

    # Level 2: two branches
    node(ax, 0.25, 0.72, 0.44, 0.048, 'ТЕПЛОВАЯ ПЕРЕГРУЗКА\nКЛАПАНА (T > 700°C)', '#FFF0F0', RED, 8, RED)
    node(ax, 0.75, 0.72, 0.44, 0.048, 'ОТСУТСТВИЕ СВОЕВРЕМЕННОЙ\nЗАЩИТЫ ДВИГАТЕЛЯ', '#FFF0F0', RED, 8, RED)
    line(ax, 0.5, 0.755, 0.25, 0.744)
    line(ax, 0.5, 0.755, 0.75, 0.744)

    # Level 3 left: thermal causes (AND)
    node(ax, 0.25, 0.64, 0.10, 0.03, 'И', PANEL, YELLOW, 7.5, YELLOW)
    line(ax, 0.25, 0.696, 0.25, 0.655)

    causes_l = [
        (0.08, 0.565, 'TIB = 200%\n(C_TIB_Fuel_Adjust_Upper_Limit)\nДвойная доза впрыска', RED),
        (0.25, 0.565, 'Max RPM = 2100\n(vs 1990 у 730E)\n+25-35°C EGT', RED),
        (0.42, 0.565, 'Иной Fuel Code\nFC0WH95 vs FC0BU52\nНестандартная стратегия', YELLOW),
    ]
    for cx, cy, text, color in causes_l:
        node(ax, cx, cy, 0.28, 0.065, text, PANEL, color, 7, TEXT)
        line(ax, 0.25, 0.625, cx, cy+0.0325)

    # Level 3 right: protection causes (AND)
    node(ax, 0.75, 0.64, 0.10, 0.03, 'И', PANEL, YELLOW, 7.5, YELLOW)
    line(ax, 0.75, 0.696, 0.75, 0.655)

    causes_r = [
        (0.58, 0.565, 'DO-выходы ОТКЛ\n(DO_Option=61042)\nОператор не информир.', RED),
        (0.75, 0.565, 'EPD RPM Derate\nОТКЛЮЧЁН\n(C_EPD_OT_RPM=0)', RED),
        (0.92, 0.565, 'FC556 порог:\n5 кПа (NTE200)\nvs 10 кПа (730E)', YELLOW),
    ]
    for cx, cy, text, color in causes_r:
        node(ax, cx, cy, 0.28, 0.065, text, PANEL, color, 7, TEXT)
        line(ax, 0.75, 0.625, cx, cy+0.0325)

    # Root causes (Level 4)
    node(ax, 0.5, 0.44, 0.10, 0.03, 'И', PANEL, ACCENT, 7.5, ACCENT)
    line(ax, 0.08, 0.532, 0.08, 0.44)
    line(ax, 0.92, 0.532, 0.92, 0.44)
    line(ax, 0.08, 0.44, 0.45, 0.44)
    line(ax, 0.92, 0.44, 0.55, 0.44)

    root_causes = [
        (0.17, 0.355, 'КОРЕННАЯ ПРИЧИНА 1:\nНекорректная калибровка\nECM (AQ60809.08)\n15 критических отличий\nот эталонной 730E', ACCENT),
        (0.50, 0.355, 'КОРЕННАЯ ПРИЧИНА 2:\nОтсутствие процедуры\nверификации калибровок\nECM при поставке\nи ТО парка NTE200', YELLOW),
        (0.83, 0.355, 'КОРЕННАЯ ПРИЧИНА 3:\nНет мониторинга\nпараметров работы ДВС\n(INSITE КАМСС не вёлся\nсистематически)', YELLOW),
    ]
    for cx, cy, text, color in root_causes:
        node(ax, cx, cy, 0.30, 0.085, text, '#EDFFF6' if color==ACCENT else PANEL, color, 7.5, color)
        line(ax, 0.5, 0.425, cx, cy+0.0425)

    # Contributing factors
    node(ax, 0.5, 0.24, 0.60, 0.042,
         'СПОСОБСТВУЮЩИЕ ФАКТОРЫ (не первопричины): зола масла CI-4, пыль Si/Al, несоблюдение интервала регулировки клапанов',
         PANEL, GRAY, 7.5, GRAY)
    line(ax, 0.5, 0.312, 0.5, 0.262)

    # Evidence banner
    node(ax, 0.5, 0.12, 0.95, 0.07,
         'ДОКАЗАТЕЛЬНАЯ БАЗА: 1) 730E с той же средой — нет отказов (устранение внешних факторов)\n'
         '2) Повторные отказы через 283 м/ч после замены ГБЦ (механизм постоянен)\n'
         '3) INSITE: Engine Protection NTE200 = «No data» vs 7 срабатываний у 730E за 35 127 м/ч',
         '#EDFFF6', ACCENT, 7, TEXT)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 24 done")

# ============================================================
# PAGE 25: EVIDENCE MATRIX
# ============================================================
def page_evidence(pdf):
    fig = make_fig()
    add_header(fig, 25)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'МАТРИЦА ДОКАЗАТЕЛЬСТВ', color=ACCENT, fontsize=13, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)
    ax.text(0.5, 0.905, '++ Сильное подтверждение  |  + Подтверждение  |  ~ Нейтрально  |  -- Опровергает', color=GRAY, fontsize=8, ha='center')

    # Evidence matrix
    headers = ['Гипотеза / Фактор', 'Данные CCEC Lab', 'Данные DML', 'INSITE КАМСС', 'Контр. 730E', 'Повт. отказы', 'ИТОГ']
    rows = [
        ('ECM калибровка (TIB=200%, RPM=2100)', '++', '++', '++', '++', '++', '++ДОКАЗАНО'),
        ('Тепловая перегрузка клапана', '++', '++', '+', '++', '++', '++ДОКАЗАНО'),
        ('DO-выходы отключены', '~', '~', '++', '++', '~', '++ДОКАЗАНО'),
        ('Нет Engine Protection', '~', '~', '++', '++', '++', '++ДОКАЗАНО'),
        ('Пыль Si/Al как первопричина', '--', '--', '--', '--', '--', '--ОПРОВЕРГНУТО'),
        ('Зола масла CI-4 как первопричина', '+/-', '--', '--', '--', '--', '--ОПРОВЕРГНУТО'),
        ('Антифриз (смешивание) как первопр.', '--', '--', '--', '--', '--', '--НЕТ ДАННЫХ'),
        ('Металлургический дефект клапана', '--', '~', '~', '--', '--', '--ОПРОВЕРГНУТО'),
        ('Ошибка регулировки клапанов', '~', '~', '~', '~', '+', '~ ЧАСТИЧНО'),
        ('Нарушение качества масла', '--', '--', '--', '--', '--', '--ОПРОВЕРГНУТО'),
    ]
    col_x = [0, 0.33, 0.45, 0.55, 0.65, 0.75, 0.84]
    col_w = [0.33, 0.12, 0.10, 0.10, 0.10, 0.09, 0.16]

    row_h = 0.063
    # Header
    rect_h = mpatches.FancyBboxPatch((col_x[0]-0.01, 0.885), sum(col_w)+0.02, row_h,
                                      boxstyle="square,pad=0", facecolor=ACCENT, edgecolor='none')
    ax.add_patch(rect_h)
    for j, (h, cx, cw) in enumerate(zip(headers, col_x, col_w)):
        ax.text(cx+cw/2, 0.885+row_h/2, h, color=BG, fontsize=6.5, ha='center', va='center', fontweight='bold')

    y = 0.885
    for i, row in enumerate(rows):
        y -= row_h
        is_proven = '++ДОКАЗАНО' in row[-1]
        is_refuted = 'ОПРОВЕРГНУТО' in row[-1]
        bg = '#EDFFF6' if is_proven else '#FFF0F0' if is_refuted else (PANEL if i % 2 == 0 else BG)
        rect2 = mpatches.FancyBboxPatch((col_x[0]-0.01, y), sum(col_w)+0.02, row_h,
                                         boxstyle="square,pad=0", facecolor=bg, edgecolor='none')
        ax.add_patch(rect2)
        for j, (val, cx, cw) in enumerate(zip(row, col_x, col_w)):
            if j == 0:
                clr = ACCENT if is_proven else RED if is_refuted else TEXT
                ax.text(cx+0.005, y+row_h/2, val, color=clr, fontsize=7, ha='left', va='center', fontweight='bold' if is_proven else 'normal')
            else:
                clr = ACCENT if '++' in str(val) else RED if '--' in str(val) else YELLOW if '+/-' in str(val) else TEXT
                ax.text(cx+cw/2, y+row_h/2, val, color=clr, fontsize=7.5, ha='center', va='center', fontweight='bold')

    # Radar chart placeholder - show as bar comparison
    ax.text(0, 0.22, 'ИТОГОВАЯ ОЦЕНКА УРОВНЯ ДОСТОВЕРНОСТИ ГИПОТЕЗ:', color=ACCENT, fontsize=9.5, fontweight='bold')

    hyps = ['ECM\nкалибровка', 'Тепловая\nперегрузка', 'DO-выходы\nОТКЛ', 'Пыль как\nпервопр.', 'Масло как\nпервопр.', 'Металл.\nдефект']
    scores = [95, 90, 88, 5, 12, 8]
    colors_h = [ACCENT if s > 50 else RED for s in scores]

    ax_bar = fig.add_axes([0.08, 0.07, 0.88, 0.13])
    ax_bar.set_facecolor(PANEL)
    bars = ax_bar.barh(range(len(hyps)), scores, color=colors_h, alpha=0.8, height=0.65)
    ax_bar.set_yticks(range(len(hyps)))
    ax_bar.set_yticklabels(hyps, color=TEXT, fontsize=7)
    ax_bar.set_xlabel('Достоверность, %', color=TEXT, fontsize=7.5)
    ax_bar.set_xlim(0, 110)
    ax_bar.tick_params(colors=TEXT, labelsize=7)
    ax_bar.set_facecolor(PANEL)
    for s in ax_bar.spines.values(): s.set_color(GRAY)
    ax_bar.axvline(x=50, color=YELLOW, linestyle=':', linewidth=1.5)
    for bar, s in zip(bars, scores):
        ax_bar.text(bar.get_width()+1, bar.get_y()+bar.get_height()/2,
                    f'{s}%', va='center', color=TEXT, fontsize=7.5, fontweight='bold')
    ax_bar.grid(axis='x', color=GRAY, alpha=0.3)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 25 done")

# ============================================================
# PAGE 26: CONCLUSIONS
# ============================================================
def page_conclusions(pdf):
    fig = make_fig()
    add_header(fig, 26)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ВЫВОДЫ', color=ACCENT, fontsize=14, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    conclusions = [
        ('В-1', RED,
         'Массовые отказы КУ ГБЦ QSK50 MCRS на NTE200 носят СИСТЕМНЫЙ характер.',
         '24 из ~50 машин (48%) поражены. Ни одна не достигла расчётного ресурса 25 000 м/ч '
         '(CES57000). 2 катастрофических отказа при <5 000 м/ч.'),
        ('В-2', ACCENT,
         'Первопричина — некорректная заводская калибровка ECM (AQ60809.08).',
         'Выявлено 15 критических отличий от эталонной калибровки 730E (AQ60217.28). '
         'Три критических: TIB=200%, Max RPM=2100, отключены все DO-выходы защиты.'),
        ('В-3', RED,
         'Параметр TIB=200% создаёт двойную компенсирующую дозу впрыска.',
         'Следствие: EGT на 100-130°C выше нормы (650°C vs 520°C). '
         'Температура тарелки клапана расчётно превышает порог ползучести Inconel 751 (700°C).'),
        ('В-4', RED,
         'Система защиты двигателя на NTE200 фактически деактивирована.',
         'NTE200 №85 (ESN 33238503): Engine Protection — «No data available» за 4 015 м/ч '
         '(27 горячих остановов без срабатывания защиты). 730E №18: 7 срабатываний за 35 127 м/ч.'),
        ('В-5', RED,
         'Повторные отказы через 274–1 076 м/ч после ремонта доказывают: первопричина не устранена.',
         'Ремонт ГБЦ не меняет калибровку ECM — тепловой режим остаётся прежним. '
         'Это классическое подтверждение системной первопричины (Level 1 RCA).'),
        ('В-6', ACCENT,
         'Контрольная группа 730E-DC окончательно исключает внешние факторы (пыль, масло, антифриз).',
         '730E: те же условия, тот же двигатель QSK50 MCRS, тот же карьер, то же масло CI-4. '
         'Нулевые отказы КУ за 2023-2026 г.'),
        ('В-7', YELLOW,
         'Выводы отчёта CCEC Lab (MS&T2026033) о пыли/масле не опровергают, а дополняют.',
         'Ca/Zn/P осадок — зола масла. Это следствие, а не причина. '
         'Пластическая деформация Inconel 751 CCEC выявила, но не объяснила корректно.'),
        ('В-8', YELLOW,
         'Экономический ущерб: оценочно >120 млн руб. без учёта риска производительности.',
         '~140 замен ГБЦ × ~1,8 млн руб./замена + 2 катастрофических отказа + '
         '12-15% снижение производительности парка. Риск нарастает без вмешательства.'),
    ]

    y = 0.90
    for code, color, title, text in conclusions:
        rect = mpatches.FancyBboxPatch((-0.01, y - 0.095), 1.02, 0.10,
                                        boxstyle="round,pad=0.005",
                                        facecolor='#EDFFF6' if color==ACCENT else PANEL, edgecolor='none')
        ax.add_patch(rect)
        ax.plot([-0.01, -0.01], [y-0.090, y-0.003], color=color, linewidth=5)
        ax.text(0.015, y - 0.008, f'[{code}]', color=color, fontsize=8, fontweight='bold', va='top')
        ax.text(0.08, y - 0.008, title, color=color, fontsize=8, fontweight='bold', va='top')
        words = text.split()
        mid = len(words) * 3 // 5
        ax.text(0.015, y - 0.042, ' '.join(words[:mid]), color=TEXT, fontsize=7.3, va='top')
        ax.text(0.015, y - 0.065, ' '.join(words[mid:]), color=TEXT, fontsize=7.3, va='top')
        y -= 0.107

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 26 done")

# ============================================================
# PAGE 27: ACTION PLAN — IMMEDIATE
# ============================================================
def page_action_immediate(pdf):
    fig = make_fig()
    add_header(fig, 27)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ПЛАН МЕРОПРИЯТИЙ — НЕОТЛОЖНЫЕ МЕРЫ', color=ACCENT, fontsize=13, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)
    ax.text(0.5, 0.908, 'Горизонт: 0-30 дней. Реализация до завершения текущего ТО-цикла.', color=YELLOW, fontsize=8.5, ha='center')

    headers = ['Мероприятие', 'Ответств.', 'Срок', 'Приоритет', 'Ожидаемый эффект']
    rows = [
        ('М-01: Снятие данных INSITE КАМСС со всех NTE200\n(полный экспорт: FC, Engine Protection, параметры)',
         'Сервисный\nинженер', '3 дня', '!КРИТ', 'Базовый аудит состояния ECM парка'),
        ('М-02: Верификация версии калибровки ECM\nна каждом NTE200 (проверка AQ60809.xx)',
         'Cummins/\nдилер', '5 дней', '!КРИТ', 'Инвентаризация реальных версий'),
        ('М-03: Запрос официальной позиции Cummins Inc.\nпо параметру TIB=200% для NTE200',
         'Технич.\ndirector', '7 дней', '!КРИТ', 'Подтверждение/опровержение заводского решения'),
        ('М-04: Запрет на эксплуатацию единиц с\nнаработкой >8 000 м/ч без проверки ECM',
         'Начальник\nтранспорта', '1 день', '!ВЫСОК', 'Снижение риска катастрофических отказов'),
        ('М-05: Внеплановая регулировка клапанов\nна единицах 5 000-10 000 м/ч',
         'Механики\nТО', '14 дней', 'ВЫСОК', 'Устранение вторичного фактора: зазоры'),
        ('М-06: Повышение частоты контроля\nуровня масла до ежедневного',
         'Операторы\nмашин', '1 день', 'СРЕДН', 'Ранее обнаружение прогрессирующего износа'),
        ('М-07: Анализ DML-данных всех машин\n(EGT по цилиндрам, RPM-профили)',
         'Инженер по\nнадёжности', '14 дней', 'ВЫСОК', 'Выявление единиц в зоне теплового риска'),
        ('М-08: Организация совещания с Cummins и\nдилером по перекалибровке ECM',
         'Директор\nпредприятия', '10 дней', '!КРИТ', 'Запуск процедуры корректировки'),
    ]

    col_x = [0, 0.43, 0.57, 0.67, 0.77]
    col_w = [0.43, 0.14, 0.10, 0.10, 0.23]

    row_h = 0.082
    rect_h = mpatches.FancyBboxPatch((col_x[0]-0.01, 0.89), sum(col_w)+0.02, row_h,
                                      boxstyle="square,pad=0", facecolor=ACCENT, edgecolor='none')
    ax.add_patch(rect_h)
    for j, (h, cx, cw) in enumerate(zip(headers, col_x, col_w)):
        ax.text(cx+cw/2, 0.89+row_h/2, h, color=BG, fontsize=7, ha='center', va='center', fontweight='bold')

    y = 0.89
    for i, row in enumerate(rows):
        y -= row_h
        is_crit = '!КРИТ' in row[3] or '!ВЫСОК' in row[3]
        bg = '#FFF0F0' if is_crit else (PANEL if i % 2 == 0 else BG)
        rect2 = mpatches.FancyBboxPatch((col_x[0]-0.01, y), sum(col_w)+0.02, row_h,
                                         boxstyle="square,pad=0", facecolor=bg, edgecolor='none')
        ax.add_patch(rect2)
        for j, (val, cx, cw) in enumerate(zip(row, col_x, col_w)):
            is_a = str(val).startswith('!')
            val2 = str(val).lstrip('!')
            clr = RED if is_a else (ACCENT if j == 0 else TEXT)
            if j == 3:
                clr = RED if 'КРИТ' in val2 else YELLOW if 'ВЫСОК' in val2 else TEXT
            ax.text(cx+(0.005 if j == 0 else cw/2), y+row_h/2, val2,
                    color=clr, fontsize=6.8, ha='left' if j == 0 else 'center', va='center')
        y -= 0

    ax.text(0.5, y - 0.03,
            'КЛЮЧЕВОЕ ДЕЙСТВИЕ: До выдачи заключения Cummins Inc. о TIB=200% — '
            'все машины NTE200 с наработкой >5 000 м/ч должны пройти проверку ECM-параметров.',
            color=RED, fontsize=8, ha='center', fontweight='bold', va='top')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 27 done")

# ============================================================
# PAGE 28: ACTION PLAN — MEDIUM/LONG TERM
# ============================================================
def page_action_longterm(pdf):
    fig = make_fig()
    add_header(fig, 28)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ПЛАН МЕРОПРИЯТИЙ — СРЕДНЕСРОЧНЫЕ И ДОЛГОСРОЧНЫЕ', color=ACCENT, fontsize=12, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    # Medium-term (1-3 months)
    ax.text(0, 0.91, 'СРЕДНЕСРОЧНЫЕ (1–3 МЕСЯЦА)', color=YELLOW, fontsize=10, fontweight='bold')
    ax.plot([0, 0.55], [0.895, 0.895], color=YELLOW, linewidth=1.5, alpha=0.5)

    medium = [
        ('М-09', 'Перекалибровка ECM всего парка NTE200',
         'После получения официальной позиции Cummins: обновление калибровки до параметров '
         'аналогичных 730E (TIB→100%, Max RPM→1990, DO-выходы→ВКЛ). '
         'Результат: устранение первопричины для 100% парка.', ACCENT),
        ('М-10', 'Восстановление DO-выходов защиты',
         'Активация каналов DO-01, DO-07, DO-08 (интеркулер, ОЖ, топливо). '
         'Подключение к световой/звуковой сигнализации в кабине оператора. '
         'Возможно, потребует пересмотра электросхемы NTE200.', RED),
        ('М-11', 'Внедрение INSITE КАМСС мониторинга',
         'Систематический съём данных INSITE после каждого ТО-1 (250 м/ч). '
         'Ведение базы данных FC-событий и Engine Protection по каждой единице. '
         'Назначение ответственного инженера.', YELLOW),
        ('М-12', 'Переход на усиленный интервал регулировки клапанов',
         'До перекалибровки ECM: сокращение интервала регулировки с 2 000 до 1 000 м/ч '
         'для всего парка NTE200. Ведение протоколов зазоров.', YELLOW),
    ]

    y = 0.88
    for code, title, text, color in medium:
        rect = mpatches.FancyBboxPatch((-0.01, y - 0.095), 1.02, 0.096,
                                        boxstyle="round,pad=0.005", facecolor=PANEL, edgecolor='none')
        ax.add_patch(rect)
        ax.plot([-0.01, -0.01], [y-0.090, y-0.003], color=color, linewidth=4)
        ax.text(0.02, y-0.007, f'[{code}]', color=color, fontsize=8, fontweight='bold', va='top')
        ax.text(0.10, y-0.007, title, color=color, fontsize=8.5, fontweight='bold', va='top')
        words = text.split()
        mid = len(words) // 2
        ax.text(0.02, y-0.040, ' '.join(words[:mid]), color=TEXT, fontsize=7.3, va='top')
        ax.text(0.02, y-0.062, ' '.join(words[mid:]), color=TEXT, fontsize=7.3, va='top')
        y -= 0.110

    # Long-term (3-12 months)
    ax.text(0, 0.445, 'ДОЛГОСРОЧНЫЕ (3–12 МЕСЯЦЕВ)', color=ACCENT, fontsize=10, fontweight='bold')
    ax.plot([0, 0.60], [0.430, 0.430], color=ACCENT, linewidth=1.5, alpha=0.5)

    longterm = [
        ('М-13', 'Разработка регламента контроля параметров ECM',
         'Создание обязательного чек-листа проверки критических параметров ECM при: '
         'поставке новых машин, любом плановом ТО-4 (4 000 м/ч), замене ECM. '
         'Включение в стандарт ТО предприятия.', ACCENT),
        ('М-14', 'Внедрение Weibull-мониторинга отказов',
         'Ведение базы данных отказов с анализом по Вейбуллу. '
         'Прогнозирование вероятности отказа для каждой единицы. '
         'Переход к условно-плановым заменам ГБЦ (не по факту, а по прогнозу).', YELLOW),
        ('М-15', 'Аудит других типов машин с QSK50 MCRS',
         'Проверка калибровок ECM на других поставках QSK50: 730E, другие NTE200, '
         'если есть. Идентификация аномальных калибровок до проявления отказов.', ACCENT),
        ('М-16', 'Включение параметра TIB в гарантийный протокол с Cummins',
         'Требование от поставщика Cummins официального подтверждения: TIB=200% является '
         'намеренной настройкой или ошибкой конфигурации? Внесение в контракт поставки.', RED),
    ]

    y = 0.41
    for code, title, text, color in longterm:
        rect = mpatches.FancyBboxPatch((-0.01, y - 0.085), 1.02, 0.087,
                                        boxstyle="round,pad=0.005", facecolor=PANEL, edgecolor='none')
        ax.add_patch(rect)
        ax.plot([-0.01, -0.01], [y-0.080, y-0.003], color=color, linewidth=4)
        ax.text(0.02, y-0.007, f'[{code}]', color=color, fontsize=8, fontweight='bold', va='top')
        ax.text(0.10, y-0.007, title, color=color, fontsize=8.5, fontweight='bold', va='top')
        words = text.split()
        mid = len(words) // 2
        ax.text(0.02, y-0.037, ' '.join(words[:mid]), color=TEXT, fontsize=7.3, va='top')
        ax.text(0.02, y-0.058, ' '.join(words[mid:]), color=TEXT, fontsize=7.3, va='top')
        y -= 0.100

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 28 done")

# ============================================================
# PAGE 29: RISK ASSESSMENT
# ============================================================
def page_risk(pdf):
    fig = make_fig()
    add_header(fig, 29)
    add_footer(fig)

    ax_main = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax_main.set_facecolor(BG)
    ax_main.set_xlim(0, 1)
    ax_main.set_ylim(0, 1)
    ax_main.axis('off')

    ax_main.text(0.5, 0.97, 'ОЦЕНКА РИСКОВ ПРИ ОТСУТСТВИИ КОРРЕКТИРУЮЩИХ ДЕЙСТВИЙ', color=ACCENT, fontsize=11, ha='center', fontweight='bold')
    ax_main.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    # Risk matrix chart
    ax1 = fig.add_axes([0.07, 0.60, 0.45, 0.30])
    ax1.set_facecolor(PANEL)

    # Risk probability vs consequence
    risk_data = [
        ('Прогар ГБЦ\n(следующий)', 9, 8, RED, 72),
        ('Катастрофич.\n отказ ДВС', 7, 10, RED, 70),
        ('Массовые\nрем. простои', 9, 7, RED, 63),
        ('Повторный\n после ремонта', 9, 6, YELLOW, 54),
        ('Потеря\nпроизводит.', 8, 5, YELLOW, 40),
        ('Вторичн.\nповрежд. поршня', 5, 9, RED, 45),
    ]

    for label, prob, severity, color, risk in risk_data:
        ax1.scatter(severity, prob, s=risk*8, c=color, alpha=0.7, zorder=3)
        ax1.text(severity + 0.1, prob, label, color=TEXT, fontsize=5.5, va='center')

    # Zones
    ax1.fill_between([7, 10], [7, 7], [10, 10], alpha=0.1, color=RED)
    ax1.fill_between([4, 7], [4, 4], [10, 10], alpha=0.05, color=YELLOW)
    ax1.set_xlabel('Тяжесть последствий', color=TEXT, fontsize=7.5)
    ax1.set_ylabel('Вероятность (следующие 6 мес.)', color=TEXT, fontsize=7.5)
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 10)
    ax1.tick_params(colors=TEXT, labelsize=7)
    ax1.set_facecolor(PANEL)
    for s in ax1.spines.values(): s.set_color(GRAY)
    ax1.set_title('Матрица рисков (Вероятность × Тяжесть)', color=TEXT, fontsize=8.5, pad=4)
    ax1.grid(color=GRAY, alpha=0.3)
    ax1.text(8, 8.5, 'ЗОНА\nКРИТ.', ha='center', color=RED, fontsize=8, fontweight='bold', alpha=0.7)

    # Risk quantification
    ax2 = fig.add_axes([0.55, 0.60, 0.40, 0.30])
    ax2.set_facecolor(PANEL)

    quarters = ['Q3\n2026', 'Q4\n2026', 'Q1\n2027', 'Q2\n2027']
    # Projected failures without action (based on Weibull extrapolation)
    proj_no_action = [6, 8, 10, 12]  # new trucks failing per quarter
    proj_with_action = [2, 1, 0, 0]

    x = np.arange(len(quarters))
    w = 0.35
    ax2.bar(x - w/2, proj_no_action, w, color=RED, alpha=0.8, label='Без действий')
    ax2.bar(x + w/2, proj_with_action, w, color=ACCENT, alpha=0.8, label='С мероприятиями М-01..М-09')
    ax2.set_xticks(x)
    ax2.set_xticklabels(quarters, color=TEXT, fontsize=8)
    ax2.set_ylabel('Новых единиц с отказами', color=TEXT, fontsize=7.5)
    ax2.tick_params(colors=TEXT, labelsize=7)
    ax2.set_facecolor(PANEL)
    for s in ax2.spines.values(): s.set_color(GRAY)
    ax2.legend(facecolor=PANEL, edgecolor=GRAY, labelcolor=TEXT, fontsize=7)
    ax2.set_title('Прогноз новых отказов (Weibull)', color=TEXT, fontsize=8.5, pad=4)
    ax2.grid(axis='y', color=GRAY, alpha=0.3)

    # Risk table
    ax3 = fig.add_axes([0.05, 0.07, 0.9, 0.50])
    ax3.set_facecolor(BG)
    ax3.set_xlim(0, 1)
    ax3.set_ylim(0, 1)
    ax3.axis('off')

    ax3.text(0.5, 0.98, 'ДЕТАЛЬНАЯ ОЦЕНКА РИСКОВ ПО СЦЕНАРИЯМ:', color=ACCENT, fontsize=9.5, ha='center', fontweight='bold')

    headers = ['Сценарий', 'Вероятн.\n(0-10)', 'Тяжесть\n(0-10)', 'RPN', 'Финансовый ущерб', 'Барьер']
    rows = [
        ('Прогар клапана (1 ГБЦ) у 10+ ед. за Q3 2026', '9', '7', '63', '~18 млн руб.', 'М-04, М-07'),
        ('Катастрофический отказ (ДВС+шатун) 2+ ед.', '7', '10', '70', '~50 млн руб./ед.', 'М-04, М-09'),
        ('Полный выход 730E/NTE200 из строя (16 ГБЦ)', '6', '9', '54', '~35 млн руб./ед.', 'М-09, М-10'),
        ('Останов производства (дефицит машин)', '8', '8', '64', '>200 млн руб./квартал', 'М-04, М-08'),
        ('Претензия от Полюс Магадан к поставщику', '5', '7', '35', 'Репутационный + штрафы', 'М-08, М-03'),
        ('Отказ на горном участке (безопасность)', '4', '10', '40', 'Несчастный случай', 'М-04, М-11'),
    ]
    col_x3 = [0, 0.40, 0.51, 0.61, 0.69, 0.83]
    col_w3 = [0.40, 0.11, 0.10, 0.08, 0.14, 0.17]
    tbl(ax3, headers, rows, 0.93, col_x3, col_w3, row_h=0.110)

    ax3.text(0.5, 0.06,
             'СУММАРНЫЙ РИСК (без действий, 12 мес.): >500 млн руб. ущерба + риск безопасности персонала.',
             color=RED, fontsize=9, ha='center', fontweight='bold')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 29 done")

# ============================================================
# PAGE 30: REFERENCES
# ============================================================
def page_references(pdf):
    fig = make_fig()
    add_header(fig, 30)
    add_footer(fig)

    ax = fig.add_axes([0.05, 0.07, 0.9, 0.84])
    ax.set_facecolor(BG)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    ax.text(0.5, 0.97, 'ИСТОЧНИКИ И ССЫЛКИ', color=ACCENT, fontsize=13, ha='center', fontweight='bold')
    ax.plot([0, 1], [0.93, 0.93], color=ACCENT, linewidth=1.5, alpha=0.6)

    refs = [
        ('Первичные данные предприятия:', [
            ('[Д-1] ГБЦ ремонты.xlsx — Журнал ремонтов клапанного узла ГБЦ. АО Полюс Магадан. 2023–2026 г. '
             '24 самосвала NTE200. Формат: номер единицы, дата, наработка, позиции ГБЦ.'),
            ('[Д-2] ОТЧЕТ_Полюс_Магадан.xlsx — Отчёт по ТО и расходу расходных материалов. '
             'АО Полюс Магадан. Март 2026 г. 189 событий дозаправки масла за период.'),
            ('[Д-3] Тех отчеты NTE.pdf — 5 технических отчётов по отказам NTE#59, NTE#62. '
             'Содержит вопрос о систематической ошибке регулировки клапанов.'),
            ('[Д-4] ТО-отчёты NTE#43, 47, 48, 55, 83 — Индивидуальные акты замены ГБЦ. '
             'Даты: 2024–2026 г. Содержат описание разрушений и ESN ДВС.'),
        ]),
        ('Данные ЭБУ и диагностики:', [
            ('[Д-5] Calterm Compare Report — Сравнение калибровок ECM: AQ60217.30 (730E) vs AQ60809.08 (NTE200). '
             'Инструмент Cummins Calterm v14.x. 15 критических параметров.'),
            ('[Д-6] INSITE КАМСС — Данные диагностики NTE200 №85 (ESN 33238503, 4 015 м/ч, AQ60809.08). '
             'Экспорт: FC-коды, Engine Protection, параметры при останове. Май 2026 г.'),
            ('[Д-7] INSITE КАМСС — Данные диагностики 730E №18 (ESN 33223470, 35 127 м/ч, AQ60217.28). '
             'FC418, FC1542, FC245, Engine Protection FC146, FC165, FC556.'),
            ('[Д-8] DML (Data Management Logger) — Логи рабочих параметров NTE200 и 730E. '
             'Температура ОГ по цилиндрам. 2025–2026 г. Формат: CSV, кодировка cp1251.'),
        ]),
        ('Технические стандарты Cummins:', [
            ('[С-1] CES57000 — Cummins Engineering Standard: Lubricating Oil Requirements for '
             'Diesel Engines (CI-4/CJ-4). Норма расхода масла: ≤0,24 л/ч, ресурс >25 000 м/ч.'),
            ('[С-2] CES51005 — Cummins Engineering Standard: Valve Material Specification. '
             'Inconel 751. Твёрдость: ≤46 HRC (тарелка), 52-62 HRC (седло).'),
            ('[С-3] CES14603 — Cummins Engineering Standard: Coolant/Antifreeze. NOAT+POAT. '
             'Норма pH 7.5-11.0, интервал замены 6 000 м/ч / 2 года.'),
            ('[С-4] QSK50 Operation and Maintenance Manual — Cummins Inc., 4021524, Rev.F. '
             'Регулировка клапанов: каждые 2 000 м/ч или 12 мес. EGT норма: ≤520°C.'),
        ]),
        ('Лабораторный и внешний анализ:', [
            ('[Л-1] CCEC Lab Report MS&T2026033 — Отчёт об износе клапанной тарелки. '
             'ESN 33232926, 15 126 м/ч. Оформил: Tao Lang. Дата: 10.03.2026. '
             'EDS-анализ осадка: Ca 42%, Zn 18%, P 15%, Si 9%, Al 6%. Твёрдость: 40-44 HRC.'),
            ('[Л-2] Анализ_Калибровок_ФИНАЛ.pptx — Внутренняя презентация АО Развитие. '
             '16 слайдов, анализ 15 параметров калибровки ECM NTE200 vs 730E. Июнь 2026.'),
        ]),
        ('Нормативные документы:', [
            ('[Н-1] ГОСТ Р 52368-2005 (EN 590:2004) — Топливо дизельное. '
             'Технические условия. Сорт Арктика: CFPP ≤ -44°C.'),
            ('[Н-2] ASTM E18 — Standard Test Method for Rockwell Hardness. '
             'Метод определения твёрдости клапанного материала.'),
            ('[Н-3] MIL-HDBK-338B — Electronic Reliability Design Handbook. '
             'Методология Weibull-анализа ресурса деталей.'),
        ]),
    ]

    y = 0.90
    for section, items in refs:
        ax.text(0, y, section, color=ACCENT, fontsize=8.5, fontweight='bold', va='top')
        y -= 0.030
        for item in items:
            ax.text(0.02, y, item, color=TEXT, fontsize=7.3, va='top', linespacing=1.3)
            y -= 0.040
        y -= 0.010

    # Signature block
    rect_s = mpatches.FancyBboxPatch((-0.01, -0.02), 1.02, 0.12,
                                      boxstyle="round,pad=0.01", facecolor=PANEL, edgecolor=ACCENT, linewidth=1.5)
    ax.add_patch(rect_s)
    ax.text(0.5, 0.097, 'Настоящий отчёт составлен на основании анализа производственных данных АО Полюс Магадан,', color=TEXT, fontsize=7.5, ha='center', va='top')
    ax.text(0.5, 0.075, 'диагностической информации INSITE КАМСС, данных DML-регистраторов и калибровок ECM.', color=TEXT, fontsize=7.5, ha='center', va='top')
    ax.text(0.5, 0.052, 'Выводы основаны на сравнительном анализе методом «контрольной группы» (730E vs NTE200)', color=TEXT, fontsize=7.5, ha='center', va='top')
    ax.text(0.5, 0.030, 'и методе FTA (Fault Tree Analysis). Анализ выполнен отделом надёжности оборудования АО Развитие.', color=TEXT, fontsize=7.5, ha='center', va='top')
    ax.text(0.02, 0.005, 'Инженер по надёжности: ________________________  Дата: Июнь 2026 г.', color=GRAY, fontsize=7, va='bottom')
    ax.text(0.98, 0.005, 'Версия 1.0  |  АО Развитие', color=GRAY, fontsize=7, va='bottom', ha='right')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
    print("Page 30 done")

# ============================================================
# MAIN
# ============================================================
if __name__ == '__main__':
    output = '/home/user/NTE200/Анализ_коренных_причин_ГБЦ_QSK50_NTE200_light_p3.pdf'
    with PdfPages(output) as pdf:
        page_thermal_mech(pdf)
        page_comparison(pdf)
        page_failure_modes(pdf)
        page_fta(pdf)
        page_evidence(pdf)
        page_conclusions(pdf)
        page_action_immediate(pdf)
        page_action_longterm(pdf)
        page_risk(pdf)
        page_references(pdf)

    import os
    size = os.path.getsize(output)
    print(f"\nPart 3 complete: {output} ({size:,} bytes)")
