#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
8D 问题解决报告生成器
=====================
接收产品名、缺陷描述、客户名等参数，输出 8D 报告 .xlsx（单 Sheet）和 .docx 文件。

依赖：
    - openpyxl（生成 Excel）
    - python-docx（生成 Word）

未安装时自动 pip install。

用法：
    python3 generate_8d.py \\
        --product "前保险杠总成（涂装件）" \\
        --defect "漆面颗粒/杂质" \\
        --customer "XX汽车有限公司" \\
        --defect-rate "11.5%" \\
        --batch-size "500" \\
        --template paint-defect \\
        --output-dir ~/Desktop
"""

import argparse
import json
import os
import subprocess
import sys
import datetime
from pathlib import Path

# ============================================================
# 依赖检查与自动安装
# ============================================================

REQUIRED_PACKAGES = {
    "openpyxl": "openpyxl",
    "docx": "python-docx",  # 注意：导入名是 docx，包名是 python-docx
}


def ensure_packages():
    """检查并自动安装所需的 Python 包。"""
    import importlib

    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            print(f"[INFO] {pip_name} 未安装，正在自动安装...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pip_name, "--quiet"]
                )
                print(f"[INFO] {pip_name} 安装完成。")
            except subprocess.CalledProcessError as e:
                print(f"[ERROR] {pip_name} 安装失败：{e}")
                print(f"        请手动执行：pip install {pip_name}")
                sys.exit(1)


ensure_packages()

# 重新导入确保可用
import openpyxl
from openpyxl.styles import (
    Alignment,
    Border,
    Side,
    PatternFill,
    Font,
)
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement


# ============================================================
# 常量
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent  # 8d-report/
TEMPLATES_DIR = SKILL_ROOT / "templates"

# 颜色
HEADER_FILL = "003366"  # 深蓝色表头
ALT_ROW_FILL = "D6E4F0"  # 交替行浅蓝
ROOT_CAUSE_FILL = "FFF2CC"  # 根本原因行黄色高亮
SUBHEADER_FILL = "4472C4"  # 次级表头
YELLOW_SECTION_FILL = "FFF8E1"  # 章节标题淡黄色底

# 字体
FONT_NAME = "微软雅黑"
FONT_NAME_EN = "Microsoft YaHei"


# ============================================================
# 工具函数
# ============================================================

def replace_placeholders(text, context):
    """替换文本中的占位符。"""
    if not isinstance(text, str):
        return text
    return (
        text.replace("{defect_type}", context.get("defect", ""))
        .replace("{product_name}", context.get("product", ""))
        .replace("{customer}", context.get("customer", ""))
    )


def replace_placeholders_deep(obj, context):
    """递归替换字典/列表中的占位符。"""
    if isinstance(obj, str):
        return replace_placeholders(obj, context)
    if isinstance(obj, list):
        return [replace_placeholders_deep(item, context) for item in obj]
    if isinstance(obj, dict):
        return {k: replace_placeholders_deep(v, context) for k, v in obj.items()}
    return obj


def generate_8d_number():
    """生成 8D 编号：8D-YYYYMMDD-HHMMSS"""
    now = datetime.datetime.now()
    return now.strftime("8D-%Y%m%d-%H%M%S")


def safe_filename(name):
    """将字符串转换为安全的文件名（去除非法字符）。"""
    illegal_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '\n', '\r', '\t']
    for ch in illegal_chars:
        name = name.replace(ch, '_')
    # 截断过长的文件名
    if len(name) > 50:
        name = name[:50]
    return name


def load_template(template_slug, context):
    """加载模板 JSON 并替换占位符。"""
    template_path = TEMPLATES_DIR / template_slug / "template.json"
    if not template_path.exists():
        print(f"[WARN] 模板 {template_slug} 不存在，回退到 generic-defect")
        template_path = TEMPLATES_DIR / "generic-defect" / "template.json"

    with open(template_path, "r", encoding="utf-8") as f:
        template = json.load(f)

    # 递归替换占位符
    template = replace_placeholders_deep(template, context)
    return template


# ============================================================
# Excel 生成
# ============================================================

def get_thin_border():
    """获取细边框。"""
    side = Side(border_style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)


def get_header_font():
    return Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")


def get_body_font():
    return Font(name=FONT_NAME, size=10, color="000000")


def get_header_fill():
    return PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")


def get_alt_fill():
    return PatternFill(start_color=ALT_ROW_FILL, end_color=ALT_ROW_FILL, fill_type="solid")


def get_root_cause_fill():
    return PatternFill(start_color=ROOT_CAUSE_FILL, end_color=ROOT_CAUSE_FILL, fill_type="solid")


def get_subheader_fill():
    return PatternFill(start_color=SUBHEADER_FILL, end_color=SUBHEADER_FILL, fill_type="solid")


def get_subheader_font():
    return Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")


def get_yellow_section_fill():
    """获取淡黄色章节标题填充。"""
    return PatternFill(start_color=YELLOW_SECTION_FILL, end_color=YELLOW_SECTION_FILL, fill_type="solid")


def get_yellow_section_font():
    """获取黄色章节标题字体（深蓝字）。"""
    return Font(name=FONT_NAME, size=12, bold=True, color=HEADER_FILL)


def apply_yellow_section_style(cell):
    """应用黄色章节标题样式。"""
    cell.font = get_yellow_section_font()
    cell.fill = get_yellow_section_fill()
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = get_thin_border()


def apply_header_style(cell):
    """应用表头样式。"""
    cell.font = get_header_font()
    cell.fill = get_header_fill()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = get_thin_border()


def apply_body_style(cell, row_idx_in_data, is_root_cause=False):
    """应用正文样式（交替行 / 根因高亮）。"""
    cell.font = get_body_font()
    if is_root_cause:
        cell.fill = get_root_cause_fill()
    elif row_idx_in_data % 2 == 1:
        cell.fill = get_alt_fill()
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = get_thin_border()


def apply_subheader_style(cell):
    """应用次级表头样式。"""
    cell.font = get_subheader_font()
    cell.fill = get_subheader_fill()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = get_thin_border()


def set_column_widths(ws, widths):
    """设置列宽。"""
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_table(ws, start_row, headers, rows, col_widths, root_cause_row_indices=None):
    """
    在指定起始行写入一个表格（含表头 + 数据行）。
    root_cause_row_indices: 数据行中需要黄色高亮的索引列表（0-based 数据行索引）
    返回：下一可用行号
    """
    if root_cause_row_indices is None:
        root_cause_row_indices = []

    # 表头
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        apply_header_style(cell)
    ws.row_dimensions[start_row].height = 28

    # 数据行
    for i, row_data in enumerate(rows):
        current_row = start_row + 1 + i
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            is_rc = i in root_cause_row_indices
            apply_body_style(cell, i, is_root_cause=is_rc)
        # 自动行高（根据内容长度估算）
        max_len = max((len(str(v)) for v in row_data), default=10)
        ws.row_dimensions[current_row].height = max(20, min(80, 15 + max_len // 4))

    # 设置列宽
    set_column_widths(ws, col_widths)

    # 冻结首行
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    return start_row + 1 + len(rows)


def write_section_title(ws, row, title, span_cols=6):
    """写入章节标题（深蓝底白字，合并单元格）。"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span_cols)
    cell = ws.cell(row=row, column=1, value=title)
    apply_header_style(cell)
    ws.row_dimensions[row].height = 30
    return row + 1


def write_yellow_section_title(ws, row, title, span_cols=6):
    """写入黄色章节标题（淡黄底深蓝字，用于区分章节目录层级）。"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span_cols)
    cell = ws.cell(row=row, column=1, value=title)
    apply_yellow_section_style(cell)
    ws.row_dimensions[row].height = 30
    return row + 1


def write_kv_block(ws, start_row, kv_pairs, label_col_width=22, value_col_width=40):
    """
    写入键值对块（左标签 + 右值），2 列布局。
    kv_pairs: [(label, value), ...]
    """
    for i, (label, value) in enumerate(kv_pairs):
        current_row = start_row + i
        # 标签列
        label_cell = ws.cell(row=current_row, column=1, value=label)
        label_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        label_cell.fill = get_subheader_fill()
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        label_cell.border = get_thin_border()
        # 值列
        value_cell = ws.cell(row=current_row, column=2, value=value)
        value_cell.font = get_body_font()
        value_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        value_cell.border = get_thin_border()
        ws.row_dimensions[current_row].height = max(20, min(60, 15 + len(str(value)) // 6))

    set_column_widths(ws, [label_col_width, value_col_width])
    return start_row + len(kv_pairs)


# ============================================================
# Excel 生成（单 Sheet 版本：D0-D8 合并为一个表格）
# ============================================================

def generate_excel(context, template, output_path, report_number):
    """生成 Excel 文件（单 Sheet：8D报告）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "8D报告"
    ws.sheet_properties.tabColor = "003366"

    row = 1
    d0_d2 = template.get("d0_d2", {})

    # ==================== D0-D2 ====================
    row = write_yellow_section_title(ws, row, "D0-D2  问题基本信息与团队组建", span_cols=4)
    row += 1

    # 报告元信息
    row = write_kv_block(ws, row, [
        ("8D 报告编号", report_number),
        ("报告生成日期", datetime.datetime.now().strftime("%Y-%m-%d")),
        ("报告发起人", "____"),
        ("报告状态", "进行中"),
    ])
    row += 1

    # 客户信息
    row = write_section_title(ws, row, "一、客户信息", span_cols=2)
    row = write_kv_block(ws, row, [
        ("客户名称", context.get("customer", "____")),
        ("客户联系人", "____"),
        ("客户投诉日期", "____"),
        ("客户投诉单号", "____"),
        ("客户反馈渠道", "____（邮件 / 电话 / 8D通知单 / 现场）"),
    ])
    row += 1

    # 产品信息
    row = write_section_title(ws, row, "二、产品信息", span_cols=2)
    row = write_kv_block(ws, row, [
        ("产品名称", context.get("product", "____")),
        ("产品编号 / 零件号", "____"),
        ("批次号", "____"),
        ("批量数量", context.get("batch_size", "____")),
        ("不良数量", "____"),
        ("不良率", context.get("defect_rate", "____")),
        ("缺陷等级", d0_d2.get("defect_level_hint", "____")),
        ("发现位置", d0_d2.get("discovery_location_hint", "____")),
        ("影响", d0_d2.get("impact_hint", "____")),
    ])
    row += 1

    # D2 5W2H
    row = write_section_title(ws, row, "三、D2 问题描述（5W2H）", span_cols=2)
    row = write_kv_block(ws, row, [
        ("What（什么问题）", context.get("defect", "____")),
        ("When（何时发现）", "____（日期 / 班次）"),
        ("Where（何处发现）", "____（工序 / 客户环节 / 产品位置）"),
        ("Who（谁发现的）", "____"),
        ("Why（为什么是问题）", "____（违反的标准 / 规格要求）"),
        ("How（如何发现）", "____（检验方法 / 测试方法）"),
        ("How many（不良数量/率）", f"{context.get('defect_rate', '____')}，批量 {context.get('batch_size', '____')}"),
    ])
    row += 1

    # 问题陈述
    row = write_section_title(ws, row, "四、问题陈述", span_cols=2)
    problem_stmt = (
        f"产品「{context.get('product', '____')}」在 {d0_d2.get('discovery_location_hint', '____')} "
        f"发现「{context.get('defect', '____')}」问题，"
        f"不良率 {context.get('defect_rate', '____')}，"
        f"涉及批次数量 {context.get('batch_size', '____')}。"
        f"该问题已导致 {d0_d2.get('impact_hint', '____')}，需要立即启动 8D 问题解决流程。"
    )
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    cell = ws.cell(row=row, column=1, value=problem_stmt)
    cell.font = get_body_font()
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = get_thin_border()
    ws.row_dimensions[row].height = 60
    row += 1

    # D1 团队
    row += 1
    row = write_section_title(ws, row, "五、D1 团队组建表", span_cols=4)
    row = write_table(ws, row,
        ["角色", "姓名", "部门", "联系方式"],
        [
            ["团队领导（质量工程师）", "____", "质量部", "____"],
            ["工艺工程师", "____", "工艺部", "____"],
            ["设备工程师", "____", "设备部", "____"],
            ["生产主管", "____", "生产部", "____"],
            ["设计工程师", "____", "研发部", "____"],
            ["SQE（供应商质量工程师）", "____", "质量部", "____"],
            ["质量经理（审核）", "____", "质量部", "____"],
        ],
        col_widths=[28, 18, 14, 22],
    )

    # ==================== D3 ====================
    row += 2
    row = write_yellow_section_title(ws, row, "D3  临时遏制措施（Interim Containment Actions）", span_cols=6)
    row += 1
    row = write_section_title(ws, row, "一、遏制措施清单", span_cols=6)
    d3 = template.get("d3_template", {})
    actions = d3.get("containment_actions", [])
    while len(actions) < 5:
        actions.append("____（请补充遏制措施）")
    row = write_table(ws, row,
        ["序号", "遏制措施", "责任人", "完成时间", "验证方法", "状态"],
        [[i + 1, actions[i], "____", "____", "____（如100%全检记录/不良率对比）", "待执行"] for i in range(5)],
        col_widths=[6, 50, 12, 14, 28, 10],
    )
    row += 1
    row = write_section_title(ws, row, "二、遏制有效性验证", span_cols=2)
    row = write_kv_block(ws, row, [
        ("遏制前不良率", context.get("defect_rate", "____")),
        ("遏制开始日期", "____"),
        ("遏制后不良率（24h内）", "____"),
        ("遏制后不良率（72h内）", "____"),
        ("遏制结论", "____（达标 / 未达标，是否需要继续遏制）"),
        ("遏制措施截止日期", "____"),
        ("客户是否认可", "____"),
    ])

    # ==================== D4 ====================
    row += 2
    row = write_yellow_section_title(ws, row, "D4  根本原因分析（Root Cause Analysis）", span_cols=4)
    row += 1
    row = write_section_title(ws, row, "一、5Why 分析", span_cols=4)
    d4 = template.get("d4_template", {})
    five_why = d4.get("5why_path", {})
    steps = five_why.get("steps", [])
    why_rows = []
    root_cause_indices = []

    if steps:
        for idx, step in enumerate(steps):
            level = step.get("level", "____")
            question = step.get("question", "____")
            answer = step.get("answer", "____")
            evidence = step.get("evidence", "____")
            why_rows.append([level, question, answer, evidence])
            if "根因" in str(level):
                root_cause_indices.append(idx)
    else:
        problem = five_why.get("problem", "____")
        why_rows.append(["问题", problem, "—", "—"])
        why1 = five_why.get("why1", "____")
        why_rows.append(["Why 1", "为什么出现该问题？", why1, "____"])
        why2_hints = five_why.get("why2_hints", [])
        why_rows.append(["Why 2", f"为什么：{why1[:30]}...", "\n".join(f"• {h}" for h in why2_hints) if why2_hints else "____", "____"])
        why3_hints = five_why.get("why3_hints", [])
        why_rows.append(["Why 3", "为什么会出现上述原因？", "\n".join(f"• {h}" for h in why3_hints) if why3_hints else "____", "____"])
        why4_hints = five_why.get("why4_hints", [])
        why_rows.append(["Why 4", "为什么管理/流程未能预防？", "\n".join(f"• {h}" for h in why4_hints) if why4_hints else "____", "____"])
        why5_root = five_why.get("why5_root", "____")
        why_rows.append(["Why 5（根因）", "为什么流程/规范存在缺陷？", why5_root, "____"])
        root_cause_indices = [5]

    row = write_table(ws, row,
        ["层级", "问题", "答案", "证据"],
        why_rows,
        col_widths=[14, 38, 50, 22],
        root_cause_row_indices=root_cause_indices,
    )

    row += 1
    row = write_section_title(ws, row, "二、鱼骨图 6M 排查", span_cols=4)
    six_m = d4.get("6m_analysis", {})

    if isinstance(six_m, list):
        m_rows = []
        m_rc_indices = []
        for idx, item in enumerate(six_m):
            m_name = item.get("m", "____")
            finding = item.get("finding", "____")
            judgment = item.get("judgment", "____")
            m_rows.append([m_name, finding, judgment])
            if "根本原因" in str(judgment):
                m_rc_indices.append(idx)
        row = write_table(ws, row,
            ["6M 维度", "排查结果", "判定"],
            m_rows,
            col_widths=[18, 55, 18],
            root_cause_row_indices=m_rc_indices,
        )
    else:
        row = write_table(ws, row,
            ["6M 维度", "候选原因", "证据", "是否根因"],
            [
                ["Man（人）", six_m.get("man", "____"), "____", "____（是/否）"],
                ["Machine（机）", six_m.get("machine", "____"), "____", "____（是/否）"],
                ["Material（料）", six_m.get("material", "____"), "____", "____（是/否）"],
                ["Method（法）", six_m.get("method", "____"), "____", "____（是/否）"],
                ["Measurement（测）", six_m.get("measurement", "____"), "____", "____（是/否）"],
                ["Environment（环）", six_m.get("environment", "____"), "____", "____（是/否）"],
            ],
            col_widths=[16, 50, 24, 14],
        )

    row += 1
    row = write_section_title(ws, row, "三、根本原因总结", span_cols=2)
    rc_summary = d4.get("root_cause_summary", [])
    if rc_summary:
        rc_kv = [(f"{rc.get('id', '____')}（{rc.get('type', '____')}）", rc.get("description", "____")) for rc in rc_summary]
        row = write_kv_block(ws, row, rc_kv, label_col_width=22, value_col_width=70)
    else:
        row = write_kv_block(ws, row, [
            ("RC1（直接原因）", "____（参考 5Why 第 1-2 层结论）"),
            ("RC2（管理原因）", "____（参考 5Why 第 3-4 层结论）"),
            ("RC3（系统原因）", "____（参考 5Why 第 5 层结论）"),
        ], label_col_width=22, value_col_width=70)

    row += 1
    row = write_section_title(ws, row, "四、根本原因验证结论", span_cols=2)
    verify_text = d4.get("verification", "")
    if verify_text:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        cell = ws.cell(row=row, column=1, value=verify_text)
        cell.font = get_body_font()
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = get_thin_border()
        ws.row_dimensions[row].height = max(30, min(120, 15 + len(verify_text) // 4))
    else:
        row = write_kv_block(ws, row, [
            ("验证方法", "____（现场观察 / 数据收集 / 重现试验）"),
            ("验证数据", "____"),
            ("验证结论", "____（已确认 / 待进一步验证）"),
            ("验证人 / 日期", "____"),
        ], label_col_width=22, value_col_width=70)

    # ==================== D5-D6 ====================
    row += 2
    row = write_yellow_section_title(ws, row, "D5-D6  永久纠正措施（Permanent Corrective Actions）", span_cols=7)
    row += 1
    row = write_section_title(ws, row, "一、D5 CA 方案评估矩阵", span_cols=6)
    d5_d6 = template.get("d5_d6_template", {})
    permanent_actions = d5_d6.get("permanent_actions", [])

    ca_rows = []
    for i, action in enumerate(permanent_actions, start=1):
        if isinstance(action, dict):
            ca_rows.append([i, action.get("action", "____"), action.get("target", "____"), "____（高/中/低）", "____（是否引入新风险）", "____（采纳/否决）"])
        else:
            ca_rows.append([i, str(action), "____", "____（高/中/低）", "____（是否引入新风险）", "____（采纳/否决）"])
    while len(ca_rows) < 3:
        ca_rows.append([len(ca_rows) + 1, "____（请补充 CA 方案）", "____", "____", "____", "____"])

    row = write_table(ws, row,
        ["序号", "CA 方案", "针对根因", "可行性", "风险评估", "决策"],
        ca_rows,
        col_widths=[6, 50, 14, 12, 30, 12],
    )

    row += 1
    row = write_section_title(ws, row, "二、D6 实施计划跟踪表", span_cols=7)
    d6_rows = []
    for i, action in enumerate(permanent_actions, start=1):
        if isinstance(action, dict):
            d6_rows.append([i, action.get("action", "____"), action.get("target", "____"), action.get("responsible", "____"), action.get("due_date", "____"), "____（未开始/进行中/已完成）", "____"])
        else:
            d6_rows.append([i, str(action), "____", "____", "____", "____（未开始/进行中/已完成）", "____"])
    while len(d6_rows) < 3:
        d6_rows.append([len(d6_rows) + 1, "____", "____", "____", "____", "____", "____"])

    row = write_table(ws, row,
        ["序号", "实施措施", "目标根因", "责任人", "完成时间", "状态", "验证结果"],
        d6_rows,
        col_widths=[6, 40, 12, 14, 14, 18, 22],
    )

    # ==================== D7 ====================
    row += 2
    row = write_yellow_section_title(ws, row, "D7  预防再发生 — 横向展开（Prevent Recurrence — Yokoten）", span_cols=6)
    row += 1
    row = write_section_title(ws, row, "一、横向展开措施清单", span_cols=5)
    d7 = template.get("d7_template", {})
    yokoten = d7.get("yokoten", [])
    y_rows = [[i + 1, item, "____（同类产线/产品/客户）", "____", "____", "____"] for i, item in enumerate(yokoten)]
    while len(y_rows) < 4:
        y_rows.append([len(y_rows) + 1, "____（请补充 Yokoten 措施）", "____", "____", "____", "____"])
    row = write_table(ws, row,
        ["序号", "横向展开措施", "推广范围", "责任人", "完成时间", "状态"],
        y_rows,
        col_widths=[6, 50, 22, 14, 14, 14],
    )

    row += 1
    row = write_section_title(ws, row, "二、PFMEA 更新", span_cols=2)
    row = write_kv_block(ws, row, [
        ("本次失效模式", "____"),
        ("原 RPN", "____"),
        ("更新后 RPN", "____"),
        ("PFMEA 更新人 / 日期", "____"),
    ], label_col_width=24, value_col_width=60)

    # ==================== D8 ====================
    row += 2
    row = write_yellow_section_title(ws, row, "D8  团队认可与关闭（Team Recognition & Closure）", span_cols=4)
    row += 1
    row = write_section_title(ws, row, "一、关闭确认", span_cols=2)
    row = write_kv_block(ws, row, [
        ("所有 CA 是否实施完成", "____（是/否）"),
        ("所有 CA 是否验证有效", "____（是/否）"),
        ("客户是否认可", "____（是/否，附客户确认邮件）"),
        ("所有文件是否更新（SOP/PFMEA/培训教材）", "____（是/否）"),
        ("横向展开是否完成", "____（是/否）"),
        ("关闭日期", "____"),
        ("关闭结论", "____（同意关闭 / 暂不关闭，原因：____）"),
    ], label_col_width=34, value_col_width=60)

    row += 1
    row = write_section_title(ws, row, "二、经验教训", span_cols=2)
    row = write_kv_block(ws, row, [
        ("本次问题处理经验", "____"),
        ("可改进之处", "____"),
        ("对其他产品的启示", "____"),
        ("建议改进的管理流程", "____"),
    ], label_col_width=24, value_col_width=60)

    row += 1
    row = write_section_title(ws, row, "三、签名栏", span_cols=4)
    row = write_table(ws, row,
        ["角色", "姓名", "签名", "日期"],
        [
            ["编制（质量工程师）", "____", "____", "____"],
            ["审核（质量经理）", "____", "____", "____"],
            ["批准（质量总监）", "____", "____", "____"],
            ["客户确认（如需）", "____", "____", "____"],
        ],
        col_widths=[24, 20, 25, 18],
    )

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    cell = ws.cell(row=row, column=1, value=f"8D 报告编号：{report_number}")
    cell.font = Font(name=FONT_NAME, size=10, italic=True, color="666666")
    cell.alignment = Alignment(horizontal="right", vertical="center")

    wb.save(output_path)
    print(f"[OK] Excel 已生成：{output_path}")


# ============================================================
# Word 生成
# ============================================================

def set_cell_bg(cell, color_hex):
    """设置 Word 表格单元格背景色。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def set_cell_borders(cell):
    """设置 Word 表格单元格四边框。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        tc_borders.append(border)
    tc_pr.append(tc_borders)


def set_run_font(run, font_name=FONT_NAME, size=10.5, bold=False, color=None):
    """设置 run 的字体（中文 + 英文）。"""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    # 中文字体
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)


def add_paragraph(doc, text, bold=False, size=10.5, color=None, alignment=None, indent=False):
    """添加段落。"""
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_heading(doc, text, level=1):
    """添加章节标题（自定义样式，避免依赖 Word 默认 Heading 样式字体问题）。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, size=14, bold=True, color=HEADER_FILL)
    elif level == 2:
        set_run_font(run, size=12, bold=True, color=SUBHEADER_FILL)
    else:
        set_run_font(run, size=11, bold=True, color="333333")
    return p


def add_table(doc, headers, rows, col_widths_cm=None, root_cause_row_indices=None):
    """
    添加带格式的表格。
    headers: [str]
    rows: [[str]]
    col_widths_cm: [float]
    root_cause_row_indices: list of 0-based row indices that should be yellow-highlighted
    """
    if root_cause_row_indices is None:
        root_cause_row_indices = []

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # 设置列宽
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)

    # 表头
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = hdr_cells[i]
        cell.text = ""
        set_cell_bg(cell, HEADER_FILL)
        set_cell_borders(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        set_run_font(run, size=10.5, bold=True, color="FFFFFF")

    # 数据行
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        is_rc = r_idx in root_cause_row_indices
        for c_idx, value in enumerate(row_data):
            cell = row_cells[c_idx]
            cell.text = ""
            if is_rc:
                set_cell_bg(cell, ROOT_CAUSE_FILL)
            elif r_idx % 2 == 1:
                set_cell_bg(cell, ALT_ROW_FILL)
            set_cell_borders(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(value))
            set_run_font(run, size=10, bold=is_rc)

    return table


def add_kv_table(doc, kv_pairs, label_width_cm=4.5, value_width_cm=12):
    """添加键值对表格（2 列）。"""
    table = doc.add_table(rows=len(kv_pairs), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for i, (label, value) in enumerate(kv_pairs):
        # 标签列
        label_cell = table.rows[i].cells[0]
        label_cell.width = Cm(label_width_cm)
        label_cell.text = ""
        set_cell_bg(label_cell, SUBHEADER_FILL)
        set_cell_borders(label_cell)
        label_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = label_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(label)
        set_run_font(run, size=10, bold=True, color="FFFFFF")

        # 值列
        value_cell = table.rows[i].cells[1]
        value_cell.width = Cm(value_width_cm)
        value_cell.text = ""
        if i % 2 == 1:
            set_cell_bg(value_cell, ALT_ROW_FILL)
        set_cell_borders(value_cell)
        value_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = value_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(value))
        set_run_font(run, size=10)

    return table


def set_doc_default_font(doc):
    """设置文档默认字体为微软雅黑。"""
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(10.5)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), FONT_NAME)
    rFonts.set(qn('w:ascii'), FONT_NAME)
    rFonts.set(qn('w:hAnsi'), FONT_NAME)


def set_page_header(doc, report_number):
    """设置页眉，右侧标注 8D-编号。"""
    section = doc.sections[0]
    header = section.header
    # 清空默认段落
    for p in header.paragraphs:
        p.text = ""
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # report_number 已包含 8D- 前缀，无需重复添加
    run = p.add_run(report_number)
    set_run_font(run, size=9, color="666666")


def generate_word(context, template, output_path, report_number):
    """生成 Word 文档。"""
    doc = Document()

    # 设置默认字体
    set_doc_default_font(doc)

    # 页面设置：A4
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(20)
    section.right_margin = Mm(20)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)

    # 页眉
    set_page_header(doc, report_number)

    # 主标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(6)
    title_p.paragraph_format.space_after = Pt(12)
    title_run = title_p.add_run("8D 问题解决报告")
    set_run_font(title_run, size=20, bold=True, color=HEADER_FILL)

    # 报告编号
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(6)
    sub_run = sub_p.add_run(f"报告编号：{report_number}    生成日期：{datetime.datetime.now().strftime('%Y-%m-%d')}")
    set_run_font(sub_run, size=10, color="666666")

    # ============ 一、D0-D2 问题基本信息 ============
    add_heading(doc, "一、D0-D2 问题基本信息与团队组建", level=1)

    d0_d2 = template.get("d0_d2", {})

    add_heading(doc, "1.1 报告基本信息", level=2)
    add_kv_table(doc, [
        ("8D 报告编号", report_number),
        ("报告生成日期", datetime.datetime.now().strftime("%Y-%m-%d")),
        ("报告发起人", "____"),
        ("报告状态", "进行中"),
    ])

    add_heading(doc, "1.2 客户信息", level=2)
    add_kv_table(doc, [
        ("客户名称", context.get("customer", "____")),
        ("客户联系人", "____"),
        ("客户投诉日期", "____"),
        ("客户投诉单号", "____"),
        ("客户反馈渠道", "____（邮件 / 电话 / 8D 通知单 / 现场）"),
    ])

    add_heading(doc, "1.3 产品信息", level=2)
    add_kv_table(doc, [
        ("产品名称", context.get("product", "____")),
        ("产品编号 / 零件号", "____"),
        ("批次号", "____"),
        ("批量数量", context.get("batch_size", "____")),
        ("不良数量", "____"),
        ("不良率", context.get("defect_rate", "____")),
        ("缺陷等级", d0_d2.get("defect_level_hint", "____")),
        ("发现位置", d0_d2.get("discovery_location_hint", "____")),
        ("影响", d0_d2.get("impact_hint", "____")),
    ])

    add_heading(doc, "1.4 D2 问题描述（5W2H）", level=2)
    add_kv_table(doc, [
        ("What（什么问题）", context.get("defect", "____")),
        ("When（何时发现）", "____（日期 / 班次）"),
        ("Where（何处发现）", "____（工序 / 客户环节 / 产品位置）"),
        ("Who（谁发现的）", "____"),
        ("Why（为什么是问题）", "____（违反的标准 / 规格要求）"),
        ("How（如何发现）", "____（检验方法 / 测试方法）"),
        ("How many（不良数量/率）", f"{context.get('defect_rate', '____')}，批量 {context.get('batch_size', '____')}"),
    ])

    add_heading(doc, "1.5 问题陈述", level=2)
    problem_stmt = (
        f"产品「{context.get('product', '____')}」在 {d0_d2.get('discovery_location_hint', '____')} "
        f"发现「{context.get('defect', '____')}」问题，"
        f"不良率 {context.get('defect_rate', '____')}，"
        f"涉及批次数量 {context.get('batch_size', '____')}。"
        f"该问题已导致 {d0_d2.get('impact_hint', '____')}，需要立即启动 8D 问题解决流程。"
    )
    add_paragraph(doc, problem_stmt)

    add_heading(doc, "1.6 D1 团队组建表", level=2)
    add_table(
        doc,
        headers=["角色", "姓名", "部门", "联系方式"],
        rows=[
            ["团队领导（质量工程师）", "____", "质量部", "____"],
            ["工艺工程师", "____", "工艺部", "____"],
            ["设备工程师", "____", "设备部", "____"],
            ["生产主管", "____", "生产部", "____"],
            ["设计工程师", "____", "研发部", "____"],
            ["SQE（供应商质量工程师）", "____", "质量部", "____"],
            ["质量经理（审核）", "____", "质量部", "____"],
        ],
        col_widths_cm=[5.5, 3.5, 2.5, 4.5],
    )

    # ============ 二、D3 临时遏制措施 ============
    add_heading(doc, "二、D3 临时遏制措施", level=1)

    add_heading(doc, "2.1 遏制措施清单", level=2)
    d3 = template.get("d3_template", {})
    actions = d3.get("containment_actions", [])
    while len(actions) < 5:
        actions.append("____（请补充遏制措施）")

    d3_rows = []
    for i, action in enumerate(actions[:5], start=1):
        d3_rows.append([str(i), action, "____", "____", "____（如100%全检记录/不良率对比）", "待执行"])

    add_table(
        doc,
        headers=["序号", "遏制措施", "责任人", "完成时间", "验证方法", "状态"],
        rows=d3_rows,
        col_widths_cm=[1.2, 6.5, 1.8, 2.0, 4.0, 1.5],
    )

    add_heading(doc, "2.2 遏制有效性验证", level=2)
    add_kv_table(doc, [
        ("遏制前不良率", context.get("defect_rate", "____")),
        ("遏制开始日期", "____"),
        ("遏制后不良率（24h内）", "____"),
        ("遏制后不良率（72h内）", "____"),
        ("遏制结论", "____（达标 / 未达标，是否需要继续遏制）"),
        ("遏制措施截止日期", "____"),
        ("客户是否认可", "____"),
    ])

    # ============ 三、D4 根本原因分析 ============
    add_heading(doc, "三、D4 根本原因分析", level=1)

    add_heading(doc, "3.1 5Why 分析", level=2)
    d4 = template.get("d4_template", {})
    five_why = d4.get("5why_path", {})

    # Use new steps array format, fall back to old format
    steps = five_why.get("steps", [])
    why_rows = []
    root_cause_indices = []

    if steps:
        for idx, step in enumerate(steps):
            level = step.get("level", "____")
            question = step.get("question", "____")
            answer = step.get("answer", "____")
            evidence = step.get("evidence", "____")
            why_rows.append([level, question, answer, evidence])
            if "根因" in str(level):
                root_cause_indices.append(idx)
    else:
        why1 = five_why.get("why1", "____")
        why2_hints = five_why.get("why2_hints", [])
        why3_hints = five_why.get("why3_hints", [])
        why4_hints = five_why.get("why4_hints", [])
        why5_root = five_why.get("why5_root", "____")
        why_rows = [
            ["问题", five_why.get("problem", "____"), "—", "—"],
            ["Why 1", "为什么出现该问题？", why1, "____（数据/观察/文件）"],
            ["Why 2", f"为什么：{why1[:30]}...", "\n".join(f"• {h}" for h in why2_hints) if why2_hints else "____", "____"],
            ["Why 3", "为什么会出现上述原因？", "\n".join(f"• {h}" for h in why3_hints) if why3_hints else "____", "____"],
            ["Why 4", "为什么管理/流程未能预防？", "\n".join(f"• {h}" for h in why4_hints) if why4_hints else "____", "____"],
            ["Why 5（根因）", "为什么流程/规范存在缺陷？", why5_root, "____（流程文件审查）"],
        ]
        root_cause_indices = [5]

    add_table(
        doc,
        headers=["层级", "问题", "答案", "证据"],
        rows=why_rows,
        col_widths_cm=[2.5, 4.5, 6.0, 3.0],
        root_cause_row_indices=root_cause_indices,
    )

    add_heading(doc, "3.2 鱼骨图 6M 排查", level=2)
    six_m = d4.get("6m_analysis", {})

    if isinstance(six_m, list):
        m_rows = []
        for item in six_m:
            m_name = item.get("m", "____")
            finding = item.get("finding", "____")
            judgment = item.get("judgment", "____")
            m_rows.append([m_name, finding, judgment])
        add_table(
            doc,
            headers=["6M 维度", "排查结果", "判定"],
            rows=m_rows,
            col_widths_cm=[3.0, 8.5, 3.5],
        )
    else:
        m_rows = [
            ["Man（人）", six_m.get("man", "____"), "____", "____（是/否）"],
            ["Machine（机）", six_m.get("machine", "____"), "____", "____（是/否）"],
            ["Material（料）", six_m.get("material", "____"), "____", "____（是/否）"],
            ["Method（法）", six_m.get("method", "____"), "____", "____（是/否）"],
            ["Measurement（测）", six_m.get("measurement", "____"), "____", "____（是/否）"],
            ["Environment（环）", six_m.get("environment", "____"), "____", "____（是/否）"],
        ]
        add_table(
            doc,
            headers=["6M 维度", "候选原因", "证据", "是否根因"],
            rows=m_rows,
            col_widths_cm=[2.8, 7.0, 3.5, 2.2],
        )

    add_heading(doc, "3.3 根本原因总结", level=2)
    rc_summary = d4.get("root_cause_summary", [])
    if rc_summary:
        rc_kv = []
        for rc in rc_summary:
            rc_id = rc.get("id", "____")
            rc_desc = rc.get("description", "____")
            rc_type = rc.get("type", "____")
            rc_kv.append((f"{rc_id}（{rc_type}）", rc_desc))
        add_kv_table(doc, rc_kv, label_width_cm=4.5, value_width_cm=12)
    else:
        add_kv_table(doc, [
            ("RC1（直接原因）", "____（参考 5Why 第 1-2 层结论）"),
            ("RC2（管理原因）", "____（参考 5Why 第 3-4 层结论）"),
            ("RC3（系统原因）", "____（参考 5Why 第 5 层结论）"),
        ], label_width_cm=4.5, value_width_cm=12)

    add_heading(doc, "3.4 根本原因验证结论", level=2)
    verify_text = d4.get("verification", "")
    if verify_text:
        add_paragraph(doc, verify_text)
    else:
        add_kv_table(doc, [
            ("验证方法", "____（现场观察 / 数据收集 / 重现试验）"),
            ("验证数据", "____"),
            ("验证结论", "____（已确认 / 待进一步验证）"),
            ("验证人 / 日期", "____"),
        ], label_width_cm=4.5, value_width_cm=12)

    # ============ 四、D5-D6 永久纠正措施 ============
    add_heading(doc, "四、D5-D6 永久纠正措施", level=1)

    add_heading(doc, "4.1 D5 CA 方案评估矩阵", level=2)
    d5_d6 = template.get("d5_d6_template", {})
    permanent_actions = d5_d6.get("permanent_actions", [])

    ca_rows = []
    for i, action in enumerate(permanent_actions, start=1):
        if isinstance(action, dict):
            action_text = action.get("action", "____")
            target = action.get("target", "____")
        else:
            action_text = str(action)
            target = "____"
        ca_rows.append([str(i), action_text, target, "____（高/中/低）", "____（是否引入新风险）", "____（采纳/否决）"])

    while len(ca_rows) < 3:
        ca_rows.append([str(len(ca_rows) + 1), "____（请补充 CA 方案）", "____", "____", "____", "____"])

    add_table(
        doc,
        headers=["序号", "CA 方案", "针对根因", "可行性", "风险评估", "决策"],
        rows=ca_rows,
        col_widths_cm=[1.0, 6.5, 1.8, 1.8, 4.0, 1.8],
    )

    add_heading(doc, "4.2 D6 实施计划跟踪表", level=2)
    d6_rows = []
    for i, action in enumerate(permanent_actions, start=1):
        if isinstance(action, dict):
            action_text = action.get("action", "____")
            target = action.get("target", "____")
            responsible = action.get("responsible", "____")
            due = action.get("due_date", "____")
        else:
            action_text = str(action)
            target = "____"
            responsible = "____"
            due = "____"
        d6_rows.append([str(i), action_text, target, responsible, due, "____", "____"])

    while len(d6_rows) < 3:
        d6_rows.append([str(len(d6_rows) + 1), "____", "____", "____", "____", "____", "____"])

    add_table(
        doc,
        headers=["序号", "实施措施", "目标根因", "责任人", "完成时间", "状态", "验证结果"],
        rows=d6_rows,
        col_widths_cm=[1.0, 5.5, 1.8, 1.8, 1.8, 2.5, 2.6],
    )

    # ============ 五、D7 预防再发生 ============
    add_heading(doc, "五、D7 预防再发生", level=1)

    add_heading(doc, "5.1 横向展开（Yokoten）措施清单", level=2)
    d7 = template.get("d7_template", {})
    yokoten = d7.get("yokoten", [])
    yokoten_rows = []
    for i, item in enumerate(yokoten, start=1):
        yokoten_rows.append([str(i), item, "____（同类产线/产品/客户）", "____", "____", "____"])

    while len(yokoten_rows) < 4:
        yokoten_rows.append([str(len(yokoten_rows) + 1), "____（请补充 Yokoten 措施）", "____", "____", "____", "____"])

    add_table(
        doc,
        headers=["序号", "横向展开措施", "推广范围", "责任人", "完成时间", "状态"],
        rows=yokoten_rows,
        col_widths_cm=[1.0, 6.5, 3.0, 1.8, 1.8, 1.8],
    )

    add_heading(doc, "5.2 PFMEA 更新", level=2)
    add_kv_table(doc, [
        ("本次失效模式", "____"),
        ("原 RPN", "____"),
        ("更新后 RPN", "____"),
        ("PFMEA 更新人 / 日期", "____"),
    ], label_width_cm=4.5, value_width_cm=12)

    # ============ 六、D8 关闭与签名 ============
    add_heading(doc, "六、D8 关闭与签名", level=1)

    add_heading(doc, "6.1 关闭确认", level=2)
    add_kv_table(doc, [
        ("所有 CA 是否实施完成", "____（是/否）"),
        ("所有 CA 是否验证有效", "____（是/否）"),
        ("客户是否认可", "____（是/否，附客户确认邮件）"),
        ("所有文件是否更新", "____（是/否，SOP/PFMEA/培训教材）"),
        ("横向展开是否完成", "____（是/否）"),
        ("关闭日期", "____"),
        ("关闭结论", "____（同意关闭 / 暂不关闭，原因：____）"),
    ], label_width_cm=6.0, value_width_cm=10.5)

    add_heading(doc, "6.2 经验教训", level=2)
    add_kv_table(doc, [
        ("本次问题处理经验", "____"),
        ("可改进之处", "____"),
        ("对其他产品的启示", "____"),
        ("建议改进的管理流程", "____"),
    ], label_width_cm=4.5, value_width_cm=12)

    add_heading(doc, "6.3 签名栏", level=2)
    add_table(
        doc,
        headers=["角色", "姓名", "签名", "日期"],
        rows=[
            ["编制（质量工程师）", "____", "____", "____"],
            ["审核（质量经理）", "____", "____", "____"],
            ["批准（质量总监）", "____", "____", "____"],
            ["客户确认（如需）", "____", "____", "____"],
        ],
        col_widths_cm=[5.0, 4.0, 4.5, 3.0],
    )

    # 末尾报告编号
    end_p = doc.add_paragraph()
    end_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    end_p.paragraph_format.space_before = Pt(12)
    end_run = end_p.add_run(f"8D 报告编号：{report_number}")
    set_run_font(end_run, size=9, color="666666")

    doc.save(output_path)
    print(f"[OK] Word 已生成：{output_path}")


# ============================================================
# 主入口
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="8D 问题解决报告生成器（Excel + Word）"
    )
    parser.add_argument("--product", required=True, help="产品名称")
    parser.add_argument("--defect", required=True, help="缺陷描述")
    parser.add_argument("--customer", required=False, default="____", help="客户名称")
    parser.add_argument("--defect-rate", required=False, default="____", help="不良率")
    parser.add_argument("--batch-size", required=False, default="____", help="批量数量")
    parser.add_argument(
        "--template",
        required=False,
        default="generic-defect",
        choices=["paint-defect", "assembly-defect", "welding-defect", "dimensional-defect", "generic-defect"],
        help="选用模板 slug",
    )
    parser.add_argument(
        "--output-dir",
        required=False,
        default=os.path.expanduser("~/Desktop"),
        help="输出目录（默认 ~/Desktop）",
    )
    return parser.parse_args()


def main():
    # Windows 终端 UTF-8 编码修复（避免 print 中文乱码导致 Agent 解析失败）
    import io
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

    args = parse_args()

    # 构造 context
    context = {
        "product": args.product,
        "defect": args.defect,
        "customer": args.customer,
        "defect_rate": args.defect_rate,
        "batch_size": args.batch_size,
    }

    # 生成 8D 编号
    report_number = generate_8d_number()
    print(f"[INFO] 8D 报告编号：{report_number}")

    # 加载模板
    template = load_template(args.template, context)
    print(f"[INFO] 使用模板：{template.get('slug', args.template)} - {template.get('name', '')}")

    # 确保输出目录存在
    output_dir = os.path.expanduser(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)
    print(f"[INFO] 输出目录：{output_dir}")

    # 生成文件名
    safe_product = safe_filename(args.product)
    safe_defect = safe_filename(args.defect)

    excel_filename = f"8D报告_{safe_product}_{safe_defect}_模板.xlsx"
    word_filename = f"8D报告_{safe_product}_{safe_defect}.docx"

    excel_path = os.path.join(output_dir, excel_filename)
    word_path = os.path.join(output_dir, word_filename)

    # 生成 Excel
    generate_excel(context, template, excel_path, report_number)

    # 生成 Word
    generate_word(context, template, word_path, report_number)

    # 输出 JSON 结果供调用方解析
    result = {
        "status": "success",
        "report_number": report_number,
        "template_slug": template.get("slug", args.template),
        "template_name": template.get("name", ""),
        "excel_path": excel_path,
        "word_path": word_path,
        "output_dir": output_dir,
    }
    print("\n[RESULT_JSON]")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
