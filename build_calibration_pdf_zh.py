#!/usr/bin/env python3
"""
Build a Chinese-language PDF report comparing ECM calibration parameters
between 730E-DC and NTE200 (QSK50 MCRS).
"""
import io
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
from matplotlib import gridspec
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np

# ─── Fonts ───────────────────────────────────────────────────────────────────
ZH_FONT  = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'
FP       = fm.FontProperties(fname=ZH_FONT)
FP_BOLD  = fm.FontProperties(fname=ZH_FONT, weight='bold')

def fp(size, bold=False):
    p = fm.FontProperties(fname=ZH_FONT, weight='bold' if bold else 'normal')
    p.set_size(size)
    return p

# ─── Palette ─────────────────────────────────────────────────────────────────
C_DARK  = '#293136'
C_DARK2 = '#1e2529'
C_ACC   = '#3EF0AF'
C_RED   = '#E05252'
C_ORG   = '#E08C2E'
C_YLW   = '#D4C94A'
C_GRY   = '#707880'
C_WHT   = '#FFFFFF'
C_BLU   = '#4EC9F0'
C_GRN   = '#90EE90'

RISK_COLOR = {'CRITICAL': C_RED, 'HIGH': C_ORG, 'WARN': C_YLW, 'NOTE': C_GRY}
RISK_LABEL_ZH = {
    'CRITICAL': '严重',
    'HIGH':     '高风险',
    'WARN':     '警告',
    'NOTE':     '备注',
}

# ─── Calibration data (Chinese) ───────────────────────────────────────────────
PARAMS = [
    {
        'param':   '燃油代码\n(喷射图谱)',
        'code':    'Fuel Code',
        'val_730': 'FC0BU52',
        'val_nte': 'FC0WH95',
        'unit':    '—',
        'risk':    'WARN',
        'effect':  '轨压、正时、限烟器、增压图谱不同\n——燃烧模式完全不同',
    },
    {
        'param':   '机油温度降额\n转速标志',
        'code':    'C_EPD_OT_RPM_Drt_En',
        'val_730': '√ 启用 (1)',
        'val_nte': '× 禁用 (0)',
        'unit':    '标志',
        'risk':    'CRITICAL',
        'effect':  'NTE200：机油过热时转速不自动降低\n730E：自动进入转速降额保护',
    },
    {
        'param':   '机油过热\n降额转速',
        'code':    'C_EPD_OT_Spd_Derate',
        'val_730': '2200',
        'val_nte': '1675',
        'unit':    'RPM',
        'risk':    'NOTE',
        'effect':  'NTE200上该标志已禁用\n——此参数值实际不起作用',
    },
    {
        'param':   '升至最大扭矩\n时间（机油）',
        'code':    'C_EPD_CP_TB_Time_To_Max_Trq_Brt',
        'val_730': '0秒（即时）',
        'val_nte': '51秒',
        'unit':    '秒',
        'risk':    'HIGH',
        'effect':  'NTE200：机油压力严重偏低时\n仍以全功率运行51秒',
    },
    {
        'param':   '升至最大扭矩\n时间（冷却液）',
        'code':    'C_EPD_CT_TB_Time_To_Max_Trq_Drt',
        'val_730': '0秒（即时）',
        'val_nte': '43秒',
        'unit':    '秒',
        'risk':    'HIGH',
        'effect':  'NTE200：冷却液过热时43秒内\n不降低负荷',
    },
    {
        'param':   '最高工作转速',
        'code':    'C_ABS_Max_Acctr_Speed',
        'val_730': '1990',
        'val_nte': '2100',
        'unit':    'RPM',
        'risk':    'HIGH',
        'effect':  '+110 RPM = 排气温度高约25–35°C\n气门冲击频率增加约5.5%',
    },
    {
        'param':   '调速器最大\n下垂率',
        'code':    'C_ABS_DroopMax',
        'val_730': '0%（等速）',
        'val_nte': '20%',
        'unit':    '%',
        'risk':    'WARN',
        'effect':  '730E保持严格恒定转速\nNTE200允许±20%转速波动',
    },
    {
        'param':   '热停机阈值\n（最大扭矩）',
        'code':    'C_HSST_Net_Torque_High_Thd',
        'val_730': '5509',
        'val_nte': '3650',
        'unit':    'N·m',
        'risk':    'HIGH',
        'effect':  'NTE200提前1859 N·m触发\n——停机前存在过载风险',
    },
    {
        'param':   '热停机阈值\n（最小扭矩）',
        'code':    'C_HSST_Net_Torque_Low_Thd',
        'val_730': '4722',
        'val_nte': '1850',
        'unit':    'N·m',
        'risk':    'CRITICAL',
        'effect':  '迟滞低2.5倍\n——热停机控制极不稳定',
    },
    {
        'param':   'TIB燃油修正\n上限',
        'code':    'C_TIB_Fuel_Adjust_Upper_Limit',
        'val_730': '100%',
        'val_nte': '200%',
        'unit':    '%',
        'risk':    'CRITICAL',
        'effect':  'NTE200可将喷油量提高2倍\n——混合气过浓，排气温度飙升',
    },
    {
        'param':   '增压冷却风扇\n启动阈值',
        'code':    'C_FCC_Charge_Temp_Y',
        'val_730': '50°C',
        'val_nte': '46°C',
        'unit':    '°C',
        'risk':    'NOTE',
        'effect':  'NTE200提前4°C冷却增压空气\n——属补偿措施，非主动保护',
    },
    {
        'param':   '数字输出：增压\n空气温度高阈值',
        'code':    'C_DO_01_A_High_Threshold',
        'val_730': '54°C',
        'val_nte': '× 禁用',
        'unit':    '°C',
        'risk':    'CRITICAL',
        'effect':  '仅730E向外部控制器\n发出增压空气过热报警',
    },
    {
        'param':   '数字输出：冷却液\n温度高阈值',
        'code':    'C_DO_07_A_High_Threshold',
        'val_730': '85°C',
        'val_nte': '× 禁用',
        'unit':    '°C',
        'risk':    'CRITICAL',
        'effect':  '仅730E向外部控制器\n发送冷却液过热信号',
    },
    {
        'param':   '数字输出：燃油\n温度高阈值',
        'code':    'C_DO_08_A_High_Threshold',
        'val_730': '68°C',
        'val_nte': '× 禁用',
        'unit':    '°C',
        'risk':    'CRITICAL',
        'effect':  '仅730E发出燃油过热报警',
    },
    {
        'param':   '最大转速参考\n(位置2/3)',
        'code':    'T_ABS_MaxRef_2/3',
        'val_730': '1990',
        'val_nte': '2225.8',
        'unit':    'RPM',
        'risk':    'HIGH',
        'effect':  '特定工况下NTE200可加速\n至2225 RPM',
    },
]

# ─── Helpers ─────────────────────────────────────────────────────────────────
def add_header(ax, title, subtitle=''):
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    # Dark bar
    ax.add_patch(mpatches.Rectangle((0, 0.82), 1, 0.18,
                  facecolor=C_DARK2, transform=ax.transAxes, clip_on=False))
    ax.add_patch(mpatches.Rectangle((0, 0.82), 0.008, 0.18,
                  facecolor=C_ACC, transform=ax.transAxes, clip_on=False))
    ax.text(0.015, 0.915, title, color=C_WHT, fontproperties=fp(16, bold=True),
            va='center', transform=ax.transAxes)
    if subtitle:
        ax.text(0.015, 0.845, subtitle, color=C_GRY, fontproperties=fp(8.5),
                va='center', transform=ax.transAxes)

def add_footer_text(ax, text):
    ax.text(0.5, 0.01, text, color=C_GRY, fontproperties=fp(7),
            ha='center', va='bottom', transform=ax.transAxes)

def page_bg(fig):
    fig.patch.set_facecolor(C_DARK)


# ─── Page builders ────────────────────────────────────────────────────────────

def page_cover(pdf):
    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    # Left accent bar
    ax.add_patch(mpatches.Rectangle((0, 0), 0.015, 1,
                  facecolor=C_ACC, transform=ax.transAxes, clip_on=False))

    ax.text(0.04, 0.72, 'ECM标定参数对比分析',
            color=C_WHT, fontproperties=fp(32, bold=True), va='center')
    ax.text(0.04, 0.60, 'QSK50 MCRS — NTE200 vs 730E-DC',
            color=C_ACC, fontproperties=fp(20), va='center')

    ax.add_patch(mpatches.Rectangle((0.04, 0.555), 0.35, 0.002,
                  facecolor=C_ACC, transform=ax.transAxes, clip_on=False))

    lines = [
        '数据来源：INSITE ECM快照存档《在负载状态下采集的数据》',
        '采集日期：2026年5月19–20日',
        '对比范围：发动机保护、燃油喷射、转速管理',
        '分析机构：АО Развитие  ·  Полюс Магадан',
    ]
    for i, ln in enumerate(lines):
        ax.text(0.04, 0.50 - i * 0.055, ln,
                color=C_GRY, fontproperties=fp(10), va='center')

    # Bottom bar
    ax.add_patch(mpatches.Rectangle((0, 0), 1, 0.06,
                  facecolor=C_DARK2, transform=ax.transAxes, clip_on=False))
    ax.text(0.5, 0.03,
            'АО Развитие  ·  QSK50 MCRS 气门磨损分析  ·  Полюс Магадан  ·  2026年  ·  机密文件',
            color=C_GRY, fontproperties=fp(8), ha='center', va='center')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def page_table(pdf, params, title, subtitle='', page_num=''):
    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)

    # header axes
    ax_h = fig.add_axes([0, 0.88, 1, 0.12])
    add_header(ax_h, title, subtitle)

    ax = fig.add_axes([0.01, 0.04, 0.98, 0.83])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    # Column definitions (x, width)
    cols = [
        (0.00, 0.155),   # 参数名称
        (0.157, 0.095),  # 代码
        (0.254, 0.095),  # 730E
        (0.351, 0.095),  # NTE200
        (0.448, 0.048),  # 单位
        (0.498, 0.032),  # 风险
        (0.532, 0.468),  # 影响
    ]
    hdrs = ['参数名称', '参数代码', '730E-DC', 'NTE200', '单位', '', '对发动机的影响']

    n = len(params)
    row_h = 0.875 / (n + 1)
    y0 = 0.97

    def draw_cell(y, cx, cw, text, bg, tc, fs=8, bold=False, center=False):
        ax.add_patch(mpatches.FancyBboxPatch(
            (cx + 0.002, y - row_h + 0.008), cw - 0.004, row_h - 0.012,
            boxstyle='round,pad=0.003', linewidth=0, facecolor=bg,
            transform=ax.transAxes, clip_on=False))
        ha = 'center' if center else 'left'
        xp = cx + cw / 2 if center else cx + 0.006
        ax.text(xp, y - row_h / 2, text, color=tc,
                fontproperties=fp(fs, bold=bold),
                ha=ha, va='center', transform=ax.transAxes,
                clip_on=False, multialignment='left' if not center else 'center')

    # Header row
    for (cx, cw), hdr in zip(cols, hdrs):
        draw_cell(y0, cx, cw, hdr, C_DARK2, C_ACC, fs=8.5, bold=True,
                  center=(hdr in ('单位', '')))
    y = y0 - row_h

    for idx, p in enumerate(params):
        rc = RISK_COLOR[p['risk']]
        rl = RISK_LABEL_ZH[p['risk']]
        bg = '#222930' if idx % 2 == 0 else '#1e2529'
        nte_tc = '#FF9999' if p['risk'] in ('CRITICAL', 'HIGH') else C_WHT

        values = [p['param'], p['code'], p['val_730'], p['val_nte'],
                  p['unit'], rl, p['effect']]
        text_colors = [C_WHT, C_GRY, C_GRN, nte_tc, C_GRY, C_WHT, '#C0C8D0']
        centers = [False, False, True, True, True, True, False]
        sizes   = [8, 6.5, 8, 8, 7.5, 7.5, 7.5]
        bolds   = [True, False, True, True, False, True, False]
        bgs     = [bg, bg, bg, bg, bg, rc, bg]

        for (cx, cw), val, tc, ce, fs, bld, bg_c in zip(
                cols, values, text_colors, centers, sizes, bolds, bgs):
            draw_cell(y, cx, cw, val, bg_c, tc, fs=fs, bold=bld, center=ce)

        y -= row_h

    # Footer
    add_footer_text(ax,
        'АО Развитие  ·  QSK50 MCRS 气门磨损分析  ·  Полюс Магадан  ·  2026  ' + page_num)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def page_bar_chart(pdf):
    numeric = [
        ('最高工作转速',     1990, 2100, 'RPM'),
        ('机油过热降额转速', 2200, 1675, 'RPM'),
        ('升至最大扭矩\n(机油压力, 秒)',  0, 51, '秒'),
        ('升至最大扭矩\n(冷却液, 秒)',    0, 43, '秒'),
        ('TIB燃油修正上限',  100, 200, '%'),
        ('HSST扭矩\n高阈值', 5509, 3650, 'N·m'),
        ('HSST扭矩\n低阈值', 4722, 1850, 'N·m'),
        ('调速器下垂率',     0, 20, '%'),
    ]

    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)

    ax_h = fig.add_axes([0, 0.88, 1, 0.12])
    add_header(ax_h, '数值参数对比：730E-DC vs NTE200',
               'QSK50 MCRS ECM关键设定值对比')

    n = len(numeric)
    cols_n = 4
    rows_n = 2
    gs = gridspec.GridSpec(rows_n, cols_n, figure=fig,
                           hspace=0.55, wspace=0.38,
                           left=0.05, right=0.97, top=0.86, bottom=0.07)

    for i, (name, v730, vnte, unit) in enumerate(numeric):
        ax = fig.add_subplot(gs[i // cols_n, i % cols_n])
        ax.set_facecolor(C_DARK2)
        ax.tick_params(colors=C_GRY, labelsize=7)
        for sp in ax.spines.values():
            sp.set_edgecolor(C_GRY); sp.set_linewidth(0.4)

        clr_nte = C_RED if vnte > v730 * 1.05 else (C_GRN if vnte < v730 * 0.95 else C_GRY)
        bars = ax.bar(['730E', 'NTE200'], [v730, vnte],
                      color=[C_BLU, clr_nte], width=0.5, zorder=3)
        top = max(v730, vnte)
        for bar, val in zip(bars, [v730, vnte]):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + top * 0.05,
                    str(val), ha='center', va='bottom',
                    color=C_WHT, fontproperties=fp(8, bold=True))

        ax.set_title(name, color=C_WHT, fontproperties=fp(8, bold=True), pad=3)
        ax.set_ylabel(unit, color=C_GRY, fontproperties=fp(7))
        ax.set_ylim(0, top * 1.28)
        ax.tick_params(axis='x', colors=C_WHT, labelsize=8)
        ax.tick_params(axis='y', colors=C_GRY, labelsize=6)
        ax.set_xticklabels(['730E', 'NTE200'],
                           fontproperties=fp(7.5))
        ax.grid(axis='y', color=C_GRY, alpha=0.18, zorder=0)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def page_disabled_protections(pdf):
    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)

    ax_h = fig.add_axes([0, 0.88, 1, 0.12])
    add_header(ax_h, 'NTE200上禁用的保护功能',
               '在730E-DC上有效、在NTE200上已停用的参数')

    ax = fig.add_axes([0.01, 0.04, 0.98, 0.83])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    disabled = [
        ('C_EPD_OT_RPM_Drt_En',     '机油过热时自动降低转速',  '启用 → 禁用',  C_RED),
        ('C_DO_01_A_High_Threshold', '增压空气过热报警输出',    '54°C → 禁用', C_RED),
        ('C_DO_07_A_High_Threshold', '冷却液过热报警输出',      '85°C → 禁用', C_RED),
        ('C_DO_08_A_High_Threshold', '燃油过热报警输出',        '68°C → 禁用', C_RED),
    ]
    elevated = [
        ('C_TIB_Fuel_Adjust_Upper_Limit',       'TIB燃油修正上限',     '100% → 200%', C_ORG),
        ('C_ABS_Max_Acctr_Speed',               '最高工作转速',         '1990 → 2100 RPM', C_ORG),
        ('C_ABS_DroopMax',                      '调速器最大下垂率',     '0% → 20%',    C_ORG),
        ('C_EPD_CT_TB_Time_To_Max_Trq_Drt',     '冷却液过热降额延迟',  '0秒 → 43秒',  C_ORG),
        ('C_EPD_CP_TB_Time_To_Max_Trq_Brt',     '机油压力降额延迟',    '0秒 → 51秒',  C_ORG),
    ]

    # Left panel header
    ax.text(0.25, 0.97, '×  在NTE200上已禁用',
            color=C_RED, fontproperties=fp(13, bold=True),
            ha='center', va='top', transform=ax.transAxes)
    ax.text(0.25, 0.90, '(730E-DC上正常启用)',
            color=C_GRY, fontproperties=fp(9), ha='center', va='top')

    for i, (code, name, change, col) in enumerate(disabled):
        y_top = 0.83 - i * 0.175
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.01, y_top - 0.145), 0.47, 0.14,
            boxstyle='round,pad=0.01', facecolor='#3a1a1a',
            edgecolor=col, linewidth=1.2,
            transform=ax.transAxes, clip_on=False))
        ax.text(0.03, y_top - 0.03, name,
                color=C_WHT, fontproperties=fp(9.5, bold=True), va='top')
        ax.text(0.03, y_top - 0.075, code,
                color=C_GRY, fontproperties=fp(7), va='top', style='italic')
        ax.text(0.46, y_top - 0.075, change,
                color=col, fontproperties=fp(9, bold=True),
                ha='right', va='center')

    # Divider
    ax.axvline(0.505, color=C_GRY, lw=0.8, alpha=0.4, ymin=0.02, ymax=0.97)

    # Right panel header
    ax.text(0.75, 0.97, '△  NTE200上的高风险设定',
            color=C_ORG, fontproperties=fp(13, bold=True),
            ha='center', va='top', transform=ax.transAxes)
    ax.text(0.75, 0.90, '(比730E-DC危险的参数值)',
            color=C_GRY, fontproperties=fp(9), ha='center', va='top')

    for i, (code, name, change, col) in enumerate(elevated):
        y_top = 0.83 - i * 0.145
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.52, y_top - 0.125), 0.47, 0.122,
            boxstyle='round,pad=0.01', facecolor='#2e2000',
            edgecolor=col, linewidth=1.2,
            transform=ax.transAxes, clip_on=False))
        ax.text(0.535, y_top - 0.025, name,
                color=C_WHT, fontproperties=fp(9, bold=True), va='top')
        ax.text(0.535, y_top - 0.075, code,
                color=C_GRY, fontproperties=fp(6.5), va='top', style='italic')
        ax.text(0.97, y_top - 0.065, change,
                color=col, fontproperties=fp(9, bold=True),
                ha='right', va='center')

    add_footer_text(ax,
        'АО Развитие  ·  QSK50 MCRS 气门磨损分析  ·  Полюс Магадан  ·  2026')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def page_fuel_code(pdf):
    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)

    ax_h = fig.add_axes([0, 0.88, 1, 0.12])
    add_header(ax_h, '燃油代码对比：FC0BU52 (730E)  vs  FC0WH95 (NTE200)',
               '喷射图谱差异及其对气门的影响')

    ax = fig.add_axes([0.01, 0.04, 0.98, 0.83])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    items = [
        ('共轨压力图谱',   '标准矿山工况优化',   '修改版——引导/主喷射模式不同'),
        ('喷射正时',       '针对1990 RPM优化',   '针对2100 RPM调整\n（缸盖温度更高）'),
        ('限烟器',         '有效，限制混合气量', '减弱——TIB=200%时允许更多燃油'),
        ('增压图谱',       '对应1990 RPM + 高度', '对应2100 RPM——进气温度更高'),
        ('排气温度',       '额定工况430–480°C',  '相同工况下470–560°C'),
    ]

    ax.text(0.26, 0.97, 'FC0BU52  —  730E-DC',
            color=C_BLU, fontproperties=fp(13, bold=True),
            ha='center', va='top')
    ax.text(0.76, 0.97, 'FC0WH95  —  NTE200',
            color=C_RED, fontproperties=fp(13, bold=True),
            ha='center', va='top')
    ax.axvline(0.505, color=C_GRY, lw=0.8, alpha=0.4, ymin=0.05, ymax=0.94)

    for i, (lbl, v730, vnte) in enumerate(items):
        y = 0.87 - i * 0.158

        # 730E card
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.01, y - 0.125), 0.46, 0.12,
            boxstyle='round,pad=0.01', facecolor='#1a2d3a',
            edgecolor=C_BLU, linewidth=0.8,
            transform=ax.transAxes, clip_on=False))
        # NTE200 card
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.53, y - 0.125), 0.46, 0.12,
            boxstyle='round,pad=0.01', facecolor='#2d1a1a',
            edgecolor=C_RED, linewidth=0.8,
            transform=ax.transAxes, clip_on=False))

        # Label in centre
        ax.text(0.505, y - 0.06, lbl,
                color=C_ACC, fontproperties=fp(9, bold=True),
                ha='center', va='center')

        ax.text(0.47, y - 0.06, v730,
                color=C_WHT, fontproperties=fp(8),
                ha='right', va='center', multialignment='right')
        ax.text(0.545, y - 0.06, vnte,
                color='#FF9999', fontproperties=fp(8),
                ha='left', va='center', multialignment='left')

    add_footer_text(ax,
        'АО Развитие  ·  QSK50 MCRS 气门磨损分析  ·  Полюс Магадан  ·  2026')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


def page_conclusion(pdf):
    fig = plt.figure(figsize=(11.69, 8.27))
    page_bg(fig)

    ax_h = fig.add_axes([0, 0.88, 1, 0.12])
    add_header(ax_h, '结论：NTE200与730E的ECM配置差异',
               '关键差异及其与气门故障的关联')

    ax = fig.add_axes([0.01, 0.04, 0.98, 0.83])
    ax.set_facecolor(C_DARK); ax.axis('off')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    conclusions = [
        (C_RED,  '严重',
         'NTE200禁用了全部4个数字输出保护（DO01/07/08、EPD_OT_RPM_Drt）。\n'
         '机器在无任何外部过热报警且无自动降额的情况下持续运行。'),
        (C_ORG,  '高风险',
         'TIB=200% → 平衡喷油修正量加倍。\n'
         '最高转速2100 vs 730E的1990 RPM → 气门冲击负荷多5.5%，排气温度高25–35°C。'),
        (C_ORG,  '高风险',
         '降额延迟51秒（机油）和43秒（冷却液）——\n'
         '在临界条件下发动机仍以全功率继续运行。'),
        (C_YLW,  '警告',
         '燃油代码FC0WH95与FC0BU52的共轨压力图谱和喷射正时不同——\n'
         '可能导致膨胀行程末期热量过高，加速气门面退缩。'),
        (C_YLW,  '警告',
         '调速器下垂率20% vs 0%——转速不稳定，\n'
         '气门机构承受周期性温度冲击（排气温度波动）。'),
        (C_ACC,  '建议',
         '将NTE200的ECM配置调整为与730E一致：\n'
         '启用DO01/07/08和EPD_OT_RPM_Drt_En，最高转速降至1990 RPM，\n'
         'TIB上限设为100%，降额延迟改为0秒。'),
    ]

    n = len(conclusions)
    h = 0.13
    gap = 0.01
    total = n * h + (n - 1) * gap
    y0 = (1 - total) / 2 + total

    for i, (color, label, text) in enumerate(conclusions):
        y = y0 - i * (h + gap)
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.0, y - h + 0.003), 0.006, h - 0.006,
            boxstyle='square,pad=0', facecolor=color, linewidth=0,
            transform=ax.transAxes, clip_on=False))
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.008, y - h + 0.003), 0.99, h - 0.006,
            boxstyle='round,pad=0.005', facecolor=C_DARK2, linewidth=0,
            transform=ax.transAxes, clip_on=False))
        ax.text(0.018, y - 0.022, label,
                color=color, fontproperties=fp(9, bold=True), va='top')
        ax.text(0.018, y - 0.058, text,
                color=C_WHT, fontproperties=fp(8),
                va='top', multialignment='left')

    add_footer_text(ax,
        'АО Развитие  ·  QSK50 MCRS 气门磨损分析  ·  Полюс Магадан  ·  2026')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    out = '/home/user/NTE200/Калибровки_ECM_NTE200_vs_730E_ZH.pdf'

    with PdfPages(out) as pdf:
        # Cover
        page_cover(pdf)
        print('Page 1: cover done')

        # Table p1
        page_table(pdf, PARAMS[:8],
                   'ECM保护参数对比：730E-DC vs NTE200  [1/2]',
                   '负载管理 · 燃油代码 · 温度保护',
                   page_num='[2/7]')
        print('Page 2: table 1 done')

        # Table p2
        page_table(pdf, PARAMS[8:],
                   'ECM保护参数对比：730E-DC vs NTE200  [2/2]',
                   '热保护 · 数字输出 · 转速参考',
                   page_num='[3/7]')
        print('Page 3: table 2 done')

        # Bar chart
        page_bar_chart(pdf)
        print('Page 4: bar chart done')

        # Disabled protections
        page_disabled_protections(pdf)
        print('Page 5: disabled protections done')

        # Fuel code
        page_fuel_code(pdf)
        print('Page 6: fuel code done')

        # Conclusion
        page_conclusion(pdf)
        print('Page 7: conclusion done')

        # PDF metadata
        d = pdf.infodict()
        d['Title']   = 'ECM标定参数对比：NTE200 vs 730E-DC'
        d['Author']  = 'АО Развитие — Технический анализ'
        d['Subject'] = 'QSK50 MCRS 发动机保护参数 — Полюс Магадан 2026'

    print(f'\nSaved: {out}')
    sz = os.path.getsize(out)
    print(f'Size: {sz/1024:.0f} KB')


if __name__ == '__main__':
    main()
