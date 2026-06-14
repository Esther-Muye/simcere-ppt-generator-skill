#!/usr/bin/env python3
"""
先声药业集团PPT生成脚本

Usage:
    python generate.py --template <template.pptx> --input <content.json> --output <output.pptx>

content.json 格式:
{
  "title": "PPT主标题",
  "subtitle": "副标题",
  "slides": [
    {
      "type": "cover",
      "title": "封面标题",
      "subtitle": "封面副标题",
      "date": "2026年6月",
      "presenter": "汇报人"
    },
    {
      "type": "toc",
      "title": "目录",
      "items": ["第一章", "第二章", "第三章"]
    },
    {
      "type": "section",
      "title": "第一章标题",
      "subtitle": "章节说明"
    },
    {
      "type": "content",
      "title": "页面标题",
      "subtitle": "页面副标题",
      "body": [
        {"text": "要点一", "level": 0},
        {"text": "子要点", "level": 1}
      ],
      "table": {...},
      "chart": {...}
    },
    {
      "type": "title_only",
      "title": "仅标题页"
    },
    {
      "type": "blank",
      "shapes": [...]
    },
    {
      "type": "ending",
      "title": "谢谢",
      "subtitle": "Thank You"
    }
  ]
}
"""

import argparse
import json
import sys
import os

# Ensure python-pptx is available
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu, Cm
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.chart.data import ChartData
    from pptx.oxml.ns import nsmap, qn
    from pptx.oxml import parse_xml
except ImportError:
    print("Error: python-pptx is required. Install with: pip install python-pptx", file=sys.stderr)
    sys.exit(1)


# === 先声药业主题色 ===
SIMCERE_COLORS = {
    'dk1': RGBColor(0x33, 0x33, 0x33),
    'lt1': RGBColor(0xFF, 0xFF, 0xFF),
    'dk2': RGBColor(0x00, 0x00, 0x00),
    'lt2': RGBColor(0xC6, 0xC7, 0xC7),
    'accent1': RGBColor(0x00, 0xB0, 0x52),   # 主绿色
    'accent2': RGBColor(0x8F, 0xD4, 0x00),   # 亮绿
    'accent3': RGBColor(0x00, 0x66, 0x47),   # 深绿
    'accent4': RGBColor(0x00, 0xB5, 0xBD),   # 青色
    'accent5': RGBColor(0xFC, 0xC9, 0x17),   # 黄色
    'accent6': RGBColor(0xF5, 0x66, 0x00),   # 橙色
    'orange': RGBColor(0xE4, 0x80, 0x30),   # 评分橙色（新增）
    'hlk_green': RGBColor(0xC8, 0xEB, 0xD6), # 高亮浅绿（新增）
    'hk_green': RGBColor(0xE8, 0xF5, 0xEE), # 浅绿背景（新增）
    'hlk_orange': RGBColor(0xFF, 0xED, 0xE1), # 浅橙背景（新增）
    'hlink': RGBColor(0xE2, 0x22, 0x41),
    'folHlink': RGBColor(0xBF, 0xBF, 0xBF),
}

# 版式名称映射
LAYOUT_NAMES = {
    'cover': '封面',
    'title_only': '仅标题页',
    'section': '节标题',
    'content': '标题和内容',
    'blank': '空白',
    'logo_only': '仅logo页',
    'logo_only_2': '1_仅logo页',
    'ending': '末尾幻灯片',
}


def find_layout(prs, layout_name):
    """按名称查找版式"""
    for layout in prs.slide_layouts:
        if layout.name == layout_name:
            return layout
    return None


def get_layout_by_type(prs, slide_type):
    """根据slide类型获取版式"""
    layout_map = {
        'cover': '封面',
        'toc': '1_仅logo页',
        'section': '节标题',
        'content': '标题和内容',
        'title_only': '仅标题页',
        'blank': '空白',
        'logo_only': '仅logo页',
        'ending': '末尾幻灯片',
    }
    name = layout_map.get(slide_type, '标题和内容')
    layout = find_layout(prs, name)
    if layout is None:
        # fallback
        for l in prs.slide_layouts:
            if name in l.name or l.name in name:
                return l
        return prs.slide_layouts[0] if len(prs.slide_layouts) > 0 else None
    return layout


def set_text_frame_text(shape, text, font_name=None, font_size=None, bold=None, color=None):
    """设置文本框内容，保留原有格式"""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    
    # 清除现有文本，保留格式
    first_para = tf.paragraphs[0]
    if first_para.runs:
        first_run = first_para.runs[0]
        first_run.text = text
    else:
        first_para.text = text
    
    # 应用额外格式（如果指定）
    for para in tf.paragraphs:
        for run in para.runs:
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = Pt(font_size)
            if bold is not None:
                run.font.bold = bold
            if color:
                run.font.color.rgb = color


def remove_shape(shape):
    """从 slide 中彻底删除一个 shape（包括占位符）"""
    sp = shape.element
    sp.getparent().remove(sp)


def remove_body_placeholders(slide):
    """删除 slide 上所有 BODY 占位符，使用先收集再删除模式避免跳过"""
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)


def add_bullet_points(shape, items, font_name='微软雅黑', font_size=14):
    """向文本框添加项目符号列表"""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    tf.clear()

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        level = item.get('level', 0) if isinstance(item, dict) else 0
        text = item.get('text', '') if isinstance(item, dict) else str(item)

        # Use run-based text setting for reliability
        p.clear()
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(font_size)
        p.level = level
        p.space_after = Pt(8)

        # 一级项目加粗
        if level == 0:
            run.font.bold = True


def add_bullet_points_to_text_frame(tf, items, font_name='微软雅黑', font_size=14):
    """直接向 TextFrame 对象添加项目符号列表"""
    tf.clear()

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        level = item.get('level', 0) if isinstance(item, dict) else 0
        text = item.get('text', '') if isinstance(item, dict) else str(item)

        p.clear()
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(font_size)
        p.level = level
        p.space_after = Pt(8)

        # 一级项目加粗
        if level == 0:
            run.font.bold = True


def add_table_to_slide(slide, data, left, top, width, height):
    """向幻灯片添加表格"""
    if not data or not data.get('headers') or not data.get('rows'):
        return None
    
    headers = data['headers']
    rows = data['rows']
    num_rows = len(rows) + 1
    num_cols = len(headers)
    
    table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table
    
    # 设置表头
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = str(header)
        # 表头格式
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.name = '微软雅黑'
                run.font.color.rgb = SIMCERE_COLORS['lt1']
        # 表头背景色（绿色）
        cell.fill.solid()
        cell.fill.fore_color.rgb = SIMCERE_COLORS['accent1']
    
    # 设置数据行
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            if col_idx >= num_cols:
                break
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(cell_text)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(11)
                    run.font.name = '微软雅黑'
                    run.font.color.rgb = SIMCERE_COLORS['dk1']
            # 隔行背景色
            if row_idx % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF0, 0xF7, 0xF4)
    
    # 设置列宽
    col_width = int(width / num_cols)
    for col in table.columns:
        col.width = col_width
    
    return table


def add_chart_to_slide(slide, chart_spec, left, top, width, height):
    """向幻灯片添加图表"""
    if not chart_spec:
        return None
    
    chart_type_str = chart_spec.get('chart_type', 'column')
    categories = chart_spec.get('categories', [])
    series_list = chart_spec.get('series', [])
    title = chart_spec.get('title', '')
    
    if not categories or not series_list:
        return None
    
    # 映射图表类型
    chart_type_map = {
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'line': XL_CHART_TYPE.LINE_MARKERS,
        'pie': XL_CHART_TYPE.PIE,
        'area': XL_CHART_TYPE.AREA,
    }
    chart_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)
    
    # 创建图表数据
    chart_data = ChartData()
    chart_data.categories = categories
    
    # 主题色序列
    theme_colors = [
        SIMCERE_COLORS['accent1'],
        SIMCERE_COLORS['accent2'],
        SIMCERE_COLORS['accent3'],
        SIMCERE_COLORS['accent4'],
        SIMCERE_COLORS['accent5'],
        SIMCERE_COLORS['accent6'],
    ]
    
    for idx, series in enumerate(series_list):
        series_name = series.get('name', f'系列{idx+1}')
        values = series.get('values', [])
        color = series.get('color')
        if color is None and idx < len(theme_colors):
            color = theme_colors[idx]
        chart_data.add_series(series_name, values)
    
    # 添加图表
    graphic_frame = slide.shapes.add_chart(
        chart_type, left, top, width, height, chart_data
    )
    chart = graphic_frame.chart
    
    # 设置图表标题
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
        for para in chart.chart_title.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.name = '微软雅黑'
                run.font.bold = True
                run.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 设置字体
    from pptx.enum.chart import XL_LEGEND_POSITION
    
    # 图例位置
    if chart.has_legend:
        chart.legend.include_in_layout = False
        # Note: python-pptx Legend does not expose text_frame for font customization
    
    # 类别轴（饼图没有category_axis）
    if chart_type_str != 'pie':
        if hasattr(chart, 'category_axis') and chart.category_axis:
            chart.category_axis.tick_labels.font.size = Pt(10)
            chart.category_axis.tick_labels.font.name = '微软雅黑'
            chart.category_axis.tick_labels.font.color.rgb = SIMCERE_COLORS['dk1']
        
        # 数值轴
        if hasattr(chart, 'value_axis') and chart.value_axis:
            chart.value_axis.tick_labels.font.size = Pt(10)
            chart.value_axis.tick_labels.font.name = '微软雅黑'
            chart.value_axis.tick_labels.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 设置系列颜色
    if chart_type_str != 'pie':
        for idx, series in enumerate(chart.series):
            if idx < len(theme_colors):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = theme_colors[idx]
    else:
        # 饼图：为每个点设置颜色
        for series in chart.series:
            for idx, point in enumerate(series.points):
                if idx < len(theme_colors):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = theme_colors[idx]
    
    return chart


def set_placeholder_text(slide, placeholder_type, text):
    """设置指定类型的占位符文本"""
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph = shape.placeholder_format
            ph_type_name = str(ph.type).split('.')[-1] if ph.type else ''
            if placeholder_type.lower() in ph_type_name.lower():
                set_text_frame_text(shape, text)
                return True
    return False


def find_placeholder(slide, placeholder_type):
    """查找指定类型的占位符"""
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph = shape.placeholder_format
            ph_type_name = str(ph.type).split('.')[-1] if ph.type else ''
            if placeholder_type.lower() in ph_type_name.lower():
                return shape
    return None


def add_text_box(slide, left, top, width, height, text, font_size=14, bold=False, color=None):
    """添加文本框"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = '微软雅黑'
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    return txBox


# ============================================================
# 新增辅助函数 — 形状、卡片、箭头
# ============================================================

def add_shape_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    """添加矩形形状（用于卡片背景、色块）"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width if line_width else Pt(1)
    else:
        shape.line.fill.background()
    
    return shape


def add_data_card(slide, left, top, width, height, title, value, description=None,
                  title_font_size=11, value_font_size=28, desc_font_size=10,
                  bg_color=None, accent_color=None):
    """添加数据卡片（矩形背景 + 标题 + 大数值 + 说明）"""
    # 背景矩形
    if bg_color:
        bg = add_shape_rect(slide, left, top, width, height, fill_color=bg_color)
    else:
        bg = add_shape_rect(slide, left, top, width, height, fill_color=RGBColor(0xF0, 0xF7, 0xF4))
    
    # 左侧装饰条
    if accent_color:
        add_shape_rect(slide, left, top, Inches(0.06), height, fill_color=accent_color)
    
    padding = Inches(0.15)
    text_left = left + padding
    text_top = top + padding
    text_width = width - padding * 2
    
    # 标题
    title_box = slide.shapes.add_textbox(text_left, text_top, text_width, Inches(0.25))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(title_font_size)
    p.font.name = '微软雅黑'
    p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 大数值
    value_box = slide.shapes.add_textbox(text_left, text_top + Inches(0.28), text_width, Inches(0.45))
    tf = value_box.text_frame
    p = tf.paragraphs[0]
    p.text = value
    p.font.size = Pt(value_font_size)
    p.font.name = '微软雅黑'
    p.font.bold = True
    p.font.color.rgb = SIMCERE_COLORS['accent1']
    
    # 说明文字
    if description:
        desc_box = slide.shapes.add_textbox(text_left, text_top + Inches(0.72), text_width, Inches(0.35))
        tf = desc_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = description
        p.font.size = Pt(desc_font_size)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk1']


def add_green_card(slide, left, top, width, height, text, font_size=11):
    """添加绿色背景卡片（用于甘特图单元格、时间线节点等）"""
    add_shape_rect(slide, left, top, width, height, fill_color=RGBColor(0xE8, 0xF5, 0xEE))
    txBox = slide.shapes.add_textbox(left + Inches(0.08), top + Inches(0.06), width - Inches(0.16), height - Inches(0.12))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = '微软雅黑'
    p.font.color.rgb = SIMCERE_COLORS['dk1']


# ============================================================
# 新增页面类型构建函数
# ============================================================

def build_gantt_slide(prs, slide, slide_spec):
    """构建矩阵甘特图页面
    
    参数:
        rows: 行标题列表 (如 ["宣传", "覆盖", "陈列", "培训", "复购"])
        months: 月份/阶段列表 (如 ["6月", "7月", "8月", "9月", "10月", "11月", "12月"])
        data: 二维数据 [{"row": "宣传", "month": "7月", "text": "社媒传播", "highlight": true}, ...]
        milestones: 底部里程碑 [{"month": "10月", "text": "达100%"}, ...]
    """
    title = slide_spec.get('title', '')
    subtitle = slide_spec.get('subtitle', '')
    rows = slide_spec.get('rows', [])
    months = slide_spec.get('months', [])
    data = slide_spec.get('data', [])
    milestones = slide_spec.get('milestones', [])
    conclusion = slide_spec.get('conclusion', '')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    # 布局计算
    margin_left = Inches(0.6)
    margin_top = Inches(1.35)
    row_header_width = Inches(1.4)
    cell_height = Inches(0.85)
    col_width = Inches(1.5)
    
    num_rows = len(rows)
    num_cols = len(months)
    
    if num_rows == 0 or num_cols == 0:
        return
    
    # 绘制表头背景（月份行）
    header_bg = add_shape_rect(slide, margin_left + row_header_width, margin_top,
                                col_width * num_cols, Inches(0.35),
                                fill_color=SIMCERE_COLORS['accent1'])
    
    # 月份表头文字
    for i, month in enumerate(months):
        txBox = slide.shapes.add_textbox(
            margin_left + row_header_width + i * col_width, margin_top,
            col_width, Inches(0.35)
        )
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = month
        p.font.size = Pt(12)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['lt1']
        p.alignment = PP_ALIGN.CENTER
    
    # 绘制行标题 + 数据单元格
    for r_idx, row_name in enumerate(rows):
        row_y = margin_top + Inches(0.35) + r_idx * cell_height
        
        # 行标题背景
        add_shape_rect(slide, margin_left, row_y, row_header_width, cell_height,
                       fill_color=RGBColor(0xE8, 0xF5, 0xEE))
        
        # 行标题文字
        txBox = slide.shapes.add_textbox(
            margin_left + Inches(0.08), row_y + Inches(0.08),
            row_header_width - Inches(0.16), cell_height - Inches(0.16)
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = row_name
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['accent3']
        
        # 数据单元格（默认背景）
        for c_idx in range(num_cols):
            cell_x = margin_left + row_header_width + c_idx * col_width
            add_shape_rect(slide, cell_x, row_y, col_width, cell_height,
                           fill_color=RGBColor(0xFA, 0xFA, 0xFA))
    
    # 填充数据
    for item in data:
        row_name = item.get('row', '')
        month = item.get('month', '')
        text = item.get('text', '')
        highlight = item.get('highlight', False)
        
        if row_name in rows and month in months:
            r_idx = rows.index(row_name)
            c_idx = months.index(month)
            row_y = margin_top + Inches(0.35) + r_idx * cell_height
            cell_x = margin_left + row_header_width + c_idx * col_width
            
            # 高亮色块
            if highlight:
                bg_color = RGBColor(0xC8, 0xEB, 0xD6)
            else:
                bg_color = RGBColor(0xFA, 0xFA, 0xFA)
            
            add_shape_rect(slide, cell_x, row_y, col_width, cell_height, fill_color=bg_color)
            
            # 文字
            if text:
                txBox = slide.shapes.add_textbox(
                    cell_x + Inches(0.06), row_y + Inches(0.06),
                    col_width - Inches(0.12), cell_height - Inches(0.12)
                )
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = text
                p.font.size = Pt(9)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 底部里程碑
    if milestones:
        milestone_y = margin_top + Inches(0.35) + num_rows * cell_height + Inches(0.15)
        for m in milestones:
            month = m.get('month', '')
            text = m.get('text', '')
            if month in months:
                c_idx = months.index(month)
                cell_x = margin_left + row_header_width + c_idx * col_width
                
                # 绿色标签背景
                add_shape_rect(slide, cell_x + Inches(0.1), milestone_y,
                               col_width - Inches(0.2), Inches(0.3),
                               fill_color=SIMCERE_COLORS['accent1'])
                
                txBox = slide.shapes.add_textbox(
                    cell_x + Inches(0.1), milestone_y,
                    col_width - Inches(0.2), Inches(0.3)
                )
                tf = txBox.text_frame
                p = tf.paragraphs[0]
                p.text = text
                p.font.size = Pt(9)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['lt1']
                p.alignment = PP_ALIGN.CENTER
    
    # 底部结论
    if conclusion:
        conclusion_y = margin_top + Inches(0.35) + num_rows * cell_height + Inches(0.6)
        txBox = slide.shapes.add_textbox(margin_left, conclusion_y,
                                         Inches(12.0), Inches(0.4))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = conclusion
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.italic = True
        p.font.color.rgb = SIMCERE_COLORS['accent3']


def build_timeline_horizontal_slide(prs, slide, slide_spec):
    """构建横向时间线页面（Campaign规划风格）
    
    参数:
        items: 列表，每个 = {"time": "7月", "title": "夏日出行", "audience": "职场高压族",
                              "actions": ["地铁专列", "UGC征集", "H5测试"]}
        subtitle_text: 顶部副标题/逻辑主线说明
    """
    title = slide_spec.get('title', '')
    subtitle_text = slide_spec.get('subtitle', '')
    items = slide_spec.get('items', [])

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    # 顶部逻辑主线说明
    if subtitle_text:
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(1.05), Inches(12.0), Inches(0.35))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = subtitle_text
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.italic = True
        p.font.color.rgb = SIMCERE_COLORS['accent3']
    
    if not items:
        return
    
    # 横向排列卡片
    margin_left = Inches(0.5)
    margin_top = Inches(1.5)
    card_width = Inches(1.75)
    card_height = Inches(4.5)
    gap = Inches(0.08)
    
    # 计算总宽度和起始位置（居中）
    total_width = len(items) * card_width + (len(items) - 1) * gap
    start_x = max(margin_left, (prs.slide_width - total_width) / 2)
    
    for i, item in enumerate(items):
        time_label = item.get('time', '')
        card_title = item.get('title', '')
        audience = item.get('audience', '')
        actions = item.get('actions', [])
        
        card_x = start_x + i * (card_width + gap)
        
        # 卡片背景（浅绿色边框 + 白色背景）
        add_shape_rect(slide, card_x, margin_top, card_width, card_height,
                       fill_color=RGBColor(0xFF, 0xFF, 0xFF),
                       line_color=RGBColor(0xC8, 0xEB, 0xD6), line_width=1)
        
        # 月份标签背景
        add_shape_rect(slide, card_x, margin_top, card_width, Inches(0.35),
                       fill_color=SIMCERE_COLORS['accent1'])
        
        # 月份文字
        txBox = slide.shapes.add_textbox(card_x, margin_top, card_width, Inches(0.35))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = time_label
        p.font.size = Pt(12)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['lt1']
        p.alignment = PP_ALIGN.CENTER
        
        # 主题名称
        title_box = slide.shapes.add_textbox(
            card_x + Inches(0.08), margin_top + Inches(0.42),
            card_width - Inches(0.16), Inches(0.6)
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = card_title
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['dk1']
        
        # 目标人群标签
        if audience:
            aud_box = slide.shapes.add_textbox(
                card_x + Inches(0.08), margin_top + Inches(1.05),
                card_width - Inches(0.16), Inches(0.3)
            )
            tf = aud_box.text_frame
            p = tf.paragraphs[0]
            p.text = f'▌ {audience}'
            p.font.size = Pt(9)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['accent1']
        
        # 动作列表
        if actions:
            action_box = slide.shapes.add_textbox(
                card_x + Inches(0.08), margin_top + Inches(1.4),
                card_width - Inches(0.16), Inches(2.5)
            )
            tf = action_box.text_frame
            tf.word_wrap = True
            for a_idx, action in enumerate(actions):
                if a_idx == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = action
                p.font.size = Pt(8)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['dk1']
                p.space_after = Pt(4)


def build_big_number_slide(prs, slide, slide_spec):
    """构建大数字展示页面（如"5,000 场科普活动"）
    
    参数:
        title: 页面标题
        big_number: 大数字文本（如"5,000"）
        unit: 单位（如"场科普活动"）
        description: 说明文字（如计算公式）
        cards: 底部卡片 [{"title": "讲课费", "calculation": "2,000元/场×5,000场", "result": "1,000W"}, ...]
        side_cards: 右侧卡片（如覆盖率、转化率等）[{"title": "140w患者", "description": "..."}, ...]
    """
    title = slide_spec.get('title', '')
    big_number = slide_spec.get('big_number', '')
    unit = slide_spec.get('unit', '')
    description = slide_spec.get('description', '')
    cards = slide_spec.get('cards', [])
    side_cards = slide_spec.get('side_cards', [])

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    # 大数字区域（左侧）
    big_num_left = Inches(0.6)
    big_num_top = Inches(1.5)
    big_num_width = Inches(5.0)
    big_num_height = Inches(2.0)
    
    # 背景色块
    add_shape_rect(slide, big_num_left, big_num_top, big_num_width, big_num_height,
                   fill_color=RGBColor(0xE8, 0xF5, 0xEE))
    
    # 大数字
    num_box = slide.shapes.add_textbox(big_num_left + Inches(0.2), big_num_top + Inches(0.15),
                                       big_num_width - Inches(0.4), Inches(0.9))
    tf = num_box.text_frame
    p = tf.paragraphs[0]
    p.text = big_number
    p.font.size = Pt(54)
    p.font.name = '微软雅黑'
    p.font.bold = True
    p.font.color.rgb = SIMCERE_COLORS['accent1']
    
    # 单位
    unit_box = slide.shapes.add_textbox(big_num_left + Inches(0.2), big_num_top + Inches(1.0),
                                        big_num_width - Inches(0.4), Inches(0.5))
    tf = unit_box.text_frame
    p = tf.paragraphs[0]
    p.text = unit
    p.font.size = Pt(16)
    p.font.name = '微软雅黑'
    p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 说明
    if description:
        desc_box = slide.shapes.add_textbox(big_num_left + Inches(0.2), big_num_top + Inches(1.5),
                                            big_num_width - Inches(0.4), Inches(0.5))
        tf = desc_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = description
        p.font.size = Pt(10)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 右侧卡片（如转化率、患者数）
    if side_cards:
        side_left = big_num_left + big_num_width + Inches(0.3)
        side_top = big_num_top
        side_card_width = Inches(4.0)
        side_card_height = Inches(0.95)
        
        for s_idx, sc in enumerate(side_cards):
            sc_y = side_top + s_idx * (side_card_height + Inches(0.1))
            add_shape_rect(slide, side_left, sc_y, side_card_width, side_card_height,
                           fill_color=RGBColor(0xE8, 0xF5, 0xEE))
            
            sc_title = sc.get('title', '')
            sc_desc = sc.get('description', '')
            
            t_box = slide.shapes.add_textbox(side_left + Inches(0.15), sc_y + Inches(0.1),
                                             side_card_width - Inches(0.3), Inches(0.4))
            tf = t_box.text_frame
            p = tf.paragraphs[0]
            p.text = sc_title
            p.font.size = Pt(16)
            p.font.name = '微软雅黑'
            p.font.bold = True
            p.font.color.rgb = SIMCERE_COLORS['accent1']
            
            if sc_desc:
                d_box = slide.shapes.add_textbox(side_left + Inches(0.15), sc_y + Inches(0.45),
                                                 side_card_width - Inches(0.3), Inches(0.4))
                tf = d_box.text_frame
                p = tf.paragraphs[0]
                p.text = sc_desc
                p.font.size = Pt(9)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 底部成本卡片
    if cards:
        card_top = big_num_top + big_num_height + Inches(0.3)
        card_width = Inches(3.8)
        card_height = Inches(1.8)
        gap = Inches(0.15)
        total_cards_width = len(cards) * card_width + (len(cards) - 1) * gap
        start_x = max(Inches(0.6), (prs.slide_width - total_cards_width) / 2)
        
        for c_idx, card in enumerate(cards):
            card_x = start_x + c_idx * (card_width + gap)
            
            # 卡片背景
            add_shape_rect(slide, card_x, card_top, card_width, card_height,
                           fill_color=RGBColor(0xF5, 0xF5, 0xF5))
            # 左侧装饰条
            add_shape_rect(slide, card_x, card_top, Inches(0.06), card_height,
                           fill_color=SIMCERE_COLORS['accent1'])
            
            c_title = card.get('title', '')
            c_calc = card.get('calculation', '')
            c_result = card.get('result', '')
            
            # 标题
            t_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(0.12),
                                             card_width - Inches(0.3), Inches(0.3))
            tf = t_box.text_frame
            p = tf.paragraphs[0]
            p.text = c_title
            p.font.size = Pt(11)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['dk1']
            
            # 计算过程
            if c_calc:
                calc_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(0.45),
                                                    card_width - Inches(0.3), Inches(0.3))
                tf = calc_box.text_frame
                p = tf.paragraphs[0]
                p.text = c_calc
                p.font.size = Pt(10)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['dk1']
            
            # 结果（大数字）
            if c_result:
                r_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(0.8),
                                                 card_width - Inches(0.3), Inches(0.5))
                tf = r_box.text_frame
                p = tf.paragraphs[0]
                p.text = c_result
                p.font.size = Pt(20)
                p.font.name = '微软雅黑'
                p.font.bold = True
                p.font.color.rgb = SIMCERE_COLORS['accent1']


def build_comparison_slide(prs, slide, slide_spec):
    """构建数据对比页面（如复购率对比）
    
    参数:
        intro: 顶部引言文字
        items: 对比项列表 [{"label": "电商整体", "value": "30%", "description": "...", "image": "..."}, ...]
        target: 目标说明文字
    """
    title = slide_spec.get('title', '')
    intro = slide_spec.get('intro', '')
    items = slide_spec.get('items', [])
    target = slide_spec.get('target', '')
    note = slide_spec.get('note', '')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    # 引言
    if intro:
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(1.15), Inches(12.0), Inches(0.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = intro
        p.font.size = Pt(12)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk1']

    if not items:
        return

    # 对比卡片
    card_top = Inches(1.8)
    card_height = Inches(2.3)
    card_width = Inches(4.0)
    gap = Inches(0.4)
    total_width = len(items) * card_width + (len(items) - 1) * gap
    start_x = max(Inches(0.6), (prs.slide_width - total_width) / 2)

    for i, item in enumerate(items):
        card_x = start_x + i * (card_width + gap)
        
        label = item.get('label', '')
        value = item.get('value', '')
        description = item.get('description', '')
        
        # 卡片背景
        add_shape_rect(slide, card_x, card_top, card_width, card_height,
                       fill_color=RGBColor(0xF5, 0xF5, 0xF5))
        
        # 标签
        l_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(0.1),
                                         card_width - Inches(0.3), Inches(0.3))
        tf = l_box.text_frame
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk1']
        
        # 大数值
        v_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(0.4),
                                         card_width - Inches(0.3), Inches(0.6))
        tf = v_box.text_frame
        p = tf.paragraphs[0]
        p.text = value
        p.font.size = Pt(36)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['accent1']
        
        # 说明
        if description:
            d_box = slide.shapes.add_textbox(card_x + Inches(0.15), card_top + Inches(1.1),
                                             card_width - Inches(0.3), Inches(1.0))
            tf = d_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = description
            p.font.size = Pt(10)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # 目标说明
    if target:
        target_y = card_top + card_height + Inches(0.2)
        txBox = slide.shapes.add_textbox(Inches(0.6), target_y, Inches(12.0), Inches(0.5))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = target
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['accent3']
    
    # 备注
    if note:
        note_y = card_top + card_height + Inches(0.8)
        txBox = slide.shapes.add_textbox(Inches(0.6), note_y, Inches(12.0), Inches(0.3))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = note
        p.font.size = Pt(9)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['lt2']


# ============================================================
# 新增：item_matrix — 物品矩阵图
# 参考：品牌提示物使用情况概览（场景→对象→物品→成本分档→评分）
# ============================================================
def build_item_matrix_slide(prs, slide, slide_spec):
    """构建物品矩阵图页面

    适用于：品牌提示物盘点、资源矩阵、产品组合展示等
    布局：左场景列 → 对象列 → 物品列 → 成本分档列 → 评分列

    参数:
        title: 页面标题
        groups: 列表，每个 = {
            "scene": "品牌活动",           # 场景名称
            "audience": "患者",              # 对象（可选）
            "cost_tier": "≤10元",           # 成本分档
            "avg_score": "8.5",             # 平均喜爱度（可选）
            "top_items": ["定制抽纸", "抽绳背包"],  # TOP物品（可选）
            "items": [                       # 本组物品列表
                {"name": "定制抽纸", "score": "8.5"},
                ...
            ]
        }
        footer_notes: 底部说明文字（可选）
    """
    title = slide_spec.get('title', '')
    groups = slide_spec.get('groups', [])
    footer_notes = slide_spec.get('footer_notes', '')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)

    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    if not groups:
        return

    # ── 布局参数 ──
    margin_left   = Inches(0.4)
    margin_top    = Inches(1.35)
    scene_w       = Inches(1.5)    # 场景列宽
    audience_w    = Inches(1.35)   # 对象列宽
    items_w       = Inches(2.4)    # 物品列宽
    cost_w        = Inches(1.35)   # 成本分档列宽
    score_w       = Inches(0.65)   # 评分列宽
    row_h         = Inches(0.33)    # 每行高度
    group_gap     = Inches(0.08)    # 组间间距

    # 场景列配色（循环使用）
    scene_colors = [
        RGBColor(0x00, 0x9D, 0x4A),   # 深绿
        RGBColor(0x1F, 0xB8, 0x40),   # 亮绿
        RGBColor(0x5A, 0xC7, 0x1E),   # 中绿
        RGBColor(0x02, 0x9B, 0x46),   # 主绿
    ]
    cost_color    = SIMCERE_COLORS['accent1']    # 成本分档绿色
    score_color   = SIMCERE_COLORS['orange']     # 评分橙色
    item_bg       = RGBColor(0xF5, 0xF5, 0xF5)   # 物品行背景

    total_row_h = row_h + group_gap

    # 计算总高度，检查是否需要缩小行高
    total_items = sum(len(g.get('items', [])) for g in groups)
    needed_height = total_items * row_h + len(groups) * group_gap + Inches(0.3)
    slide_body_h = prs.slide_height - Inches(1.5)
    if needed_height > slide_body_h:
        row_h = (slide_body_h - Inches(0.5) - len(groups) * group_gap) / max(total_items, 1)
        row_h = max(row_h, Inches(0.22))

    # ── 绘制列标题（第一行顶部标签）──
    header_y = margin_top
    # 场景列标题
    add_shape_rect(slide, margin_left, header_y,
                   scene_w, Inches(0.32), fill_color=SIMCERE_COLORS['accent1'])
    _add_centered_text(slide, margin_left, header_y,
                       scene_w, Inches(0.32), '场景', font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)
    # 对象列标题
    add_shape_rect(slide, margin_left + scene_w, header_y,
                   audience_w, Inches(0.32), fill_color=SIMCERE_COLORS['accent1'])
    _add_centered_text(slide, margin_left + scene_w, header_y,
                       audience_w, Inches(0.32), '对象', font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)
    # 物品列标题
    add_shape_rect(slide, margin_left + scene_w + audience_w, header_y,
                   items_w, Inches(0.32), fill_color=SIMCERE_COLORS['accent1'])
    _add_centered_text(slide, margin_left + scene_w + audience_w, header_y,
                       items_w, Inches(0.32), '品牌提示物', font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)
    # 成本分档列标题
    add_shape_rect(slide, margin_left + scene_w + audience_w + items_w, header_y,
                   cost_w, Inches(0.32), fill_color=SIMCERE_COLORS['accent1'])
    _add_centered_text(slide, margin_left + scene_w + audience_w + items_w, header_y,
                       cost_w, Inches(0.32), '成本分档', font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)
    # 平均喜爱度列标题
    add_shape_rect(slide, margin_left + scene_w + audience_w + items_w + cost_w, header_y,
                   score_w, Inches(0.32), fill_color=score_color)
    _add_centered_text(slide, margin_left + scene_w + audience_w + items_w + cost_w, header_y,
                       score_w, Inches(0.32), '平均\n喜爱度', font_size=10, color=SIMCERE_COLORS['lt1'], bold=True)

    # ── 逐组绘制 ──
    current_y = margin_top + Inches(0.32)   # 列标题下方

    for g_idx, group in enumerate(groups):
        scene      = group.get('scene', '')
        audience   = group.get('audience', '')
        cost_tier  = group.get('cost_tier', '')
        avg_score  = group.get('avg_score', '')
        items       = group.get('items', [])
        top_items  = group.get('top_items', [])
        scene_color = scene_colors[g_idx % len(scene_colors)]

        n_items = len(items)
        if n_items == 0:
            n_items = 1   # 至少占一行

        group_h = n_items * row_h

        # ── 场景色块（左列，跨所有物品行）──
        scene_rect = add_shape_rect(
            slide, margin_left, current_y,
            scene_w, group_h,
            fill_color=scene_color
        )
        # 场景文字（竖排居中）
        _add_centered_text(
            slide, margin_left, current_y,
            scene_w, group_h, scene,
            font_size=12, color=SIMCERE_COLORS['lt1'], bold=True,
            word_wrap=True
        )

        # ── 对象标签（左数第二列，跨所有物品行）──
        aud_rect = add_shape_rect(
            slide, margin_left + scene_w, current_y,
            audience_w, group_h,
            fill_color=SIMCERE_COLORS['hk_green']
        )
        _add_centered_text(
            slide, margin_left + scene_w, current_y,
            audience_w, group_h, audience,
            font_size=10, color=SIMCERE_COLORS['dk1'], bold=False,
            word_wrap=True
        )

        # ── 物品行 ──
        for i_idx, item in enumerate(items):
            item_name = item.get('name', '') if isinstance(item, dict) else str(item)
            item_score = item.get('score', '') if isinstance(item, dict) else ''

            row_y = current_y + i_idx * row_h

            # 物品名称背景
            add_shape_rect(
                slide, margin_left + scene_w + audience_w, row_y,
                items_w, row_h,
                fill_color=item_bg
            )
            # 物品名称文字
            _add_left_text(
                slide, margin_left + scene_w + audience_w + Inches(0.06), row_y,
                items_w - Inches(0.12), row_h, item_name,
                font_size=9, color=SIMCERE_COLORS['dk1']
            )

            # 成本分档（仅在组第一行显示，跨所有行通过合并视觉效果）
            if i_idx == 0:
                add_shape_rect(
                    slide, margin_left + scene_w + audience_w + items_w, row_y,
                    cost_w, group_h,
                    fill_color=RGBColor(0xE8, 0xF5, 0xEE)
                )
                _add_centered_text(
                    slide, margin_left + scene_w + audience_w + items_w, row_y,
                    cost_w, group_h, cost_tier,
                    font_size=10, color=SIMCERE_COLORS['accent3'], bold=True,
                    word_wrap=True
                )

            # 评分（物品级，每行一个）
            score_val = item_score if item_score else (avg_score if i_idx == 0 else '')
            if score_val:
                add_shape_rect(
                    slide, margin_left + scene_w + audience_w + items_w + cost_w, row_y,
                    score_w, row_h,
                    fill_color=SIMCERE_COLORS['hlk_orange']   # 浅橙背景
                )
                _add_centered_text(
                    slide, margin_left + scene_w + audience_w + items_w + cost_w, row_y,
                    score_w, row_h, score_val,
                    font_size=10, color=score_color, bold=True
                )

        # ── TOP物品标注（在物品列右侧添加小标签）──
        if top_items and items_w + cost_w + score_w > 0:
            top_y = current_y + group_h + Inches(0.02)
            top_text = 'TOP最受欢迎: ' + '、'.join(top_items[:3])
            txBox = slide.shapes.add_textbox(
                margin_left + scene_w + audience_w, top_y,
                items_w + cost_w, Inches(0.22)
            )
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = top_text
            p.font.size = Pt(8)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['accent3']
            p.font.italic = True

        current_y += group_h + group_gap

    # ── 底部说明 ──
    if footer_notes:
        note_y = current_y + Inches(0.1)
        txBox = slide.shapes.add_textbox(margin_left, note_y,
                                         prs.slide_width - margin_left * 2, Inches(0.35))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = footer_notes
        p.font.size = Pt(9)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['lt2']
        p.font.italic = True


# ============================================================
# 新增：calendar_grid — 日历网格（月度规划/推广排期）
# 参考：英太青整体品牌提示物推广使用建议
# ============================================================
def build_calendar_grid_slide(prs, slide, slide_spec):
    """构建日历网格页面

    适用于：月度推广排期、物品使用规划、项目时间网格等
    布局：行=场景/维度，列=月份，单元格=内容

    参数:
        title: 页面标题
        rows: 行标题列表，每个 = {"name": "品牌活动", "subtitle": "（对象：患者）"}
        months: 月份列表，如 ["7月", "8月", "9月", "10月", "11月", "12月"]
        grid: 二维数据，grid[row_idx][col_idx] = {"items": ["物品A", "物品B"], "highlight": false}
        note: 底部备注（可选）
    """
    title = slide_spec.get('title', '')
    rows  = slide_spec.get('rows', [])
    months = slide_spec.get('months', [])
    grid  = slide_spec.get('grid', [])
    note  = slide_spec.get('note', '')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)

    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    if not rows or not months:
        return

    # ── 布局参数 ──
    margin_left  = Inches(0.3)
    margin_top   = Inches(1.3)
    row_header_w = Inches(1.3)    # 行标题列宽
    month_hd_h   = Inches(0.38)    # 月份表头高度
    row_h        = Inches(1.15)     # 每行高度（需容纳多物品）
    col_w        = (prs.slide_width - margin_left - row_header_w - Inches(0.3)) / max(len(months), 1)
    col_w        = min(col_w, Inches(2.0))
    total_grid_w  = row_header_w + col_w * len(months)

    # 如果总宽度超出，缩小col_w
    max_grid_w = prs.slide_width - margin_left * 2
    if total_grid_w > max_grid_w:
        col_w = (max_grid_w - row_header_w) / len(months)

    # ── 绘制月份表头 ──
    for c_idx, month in enumerate(months):
        col_x = margin_left + row_header_w + c_idx * col_w
        add_shape_rect(
            slide, col_x, margin_top,
            col_w, month_hd_h,
            fill_color=SIMCERE_COLORS['accent1']
        )
        _add_centered_text(
            slide, col_x, margin_top,
            col_w, month_hd_h, month,
            font_size=12, color=SIMCERE_COLORS['lt1'], bold=True
        )

    # ── 绘制行 ──
    for r_idx, row in enumerate(rows):
        row_name    = row.get('name', '') if isinstance(row, dict) else str(row)
        row_sub     = row.get('subtitle', '') if isinstance(row, dict) else ''
        row_y       = margin_top + month_hd_h + r_idx * row_h

        # 行标题背景
        row_color = SIMCERE_COLORS['accent3'] if r_idx % 2 == 0 else SIMCERE_COLORS['accent1']
        add_shape_rect(
            slide, margin_left, row_y,
            row_header_w, row_h,
            fill_color=row_color
        )
        # 行标题文字
        title_text = row_name
        if row_sub:
            title_text += f"\n{row_sub}"
        _add_centered_text(
            slide, margin_left, row_y,
            row_header_w, row_h, title_text,
            font_size=10, color=SIMCERE_COLORS['lt1'], bold=True,
            word_wrap=True
        )

        # 单元格
        for c_idx in range(len(months)):
            col_x = margin_left + row_header_w + c_idx * col_w
            cell_data = {}
            if r_idx < len(grid) and c_idx < len(grid[r_idx]):
                cell_data = grid[r_idx][c_idx]

            items     = cell_data.get('items', [])
            highlight = cell_data.get('highlight', False)

            # 单元格背景
            if highlight:
                bg = SIMCERE_COLORS['hlk_green']   # 高亮浅绿
            elif c_idx % 2 == 0:
                bg = RGBColor(0xFA, 0xFA, 0xFA)
            else:
                bg = RGBColor(0xFF, 0xFF, 0xFF)

            add_shape_rect(
                slide, col_x, row_y,
                col_w, row_h,
                fill_color=bg
            )

            # 单元格边框
            cell_rect = slide.shapes[ len(slide.shapes) - 1 ]
            if cell_rect:
                try:
                    cell_rect.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
                    cell_rect.line.width = Pt(0.5)
                except:
                    pass

            # 单元格内容（物品列表）
            if items:
                item_text = '\n'.join(items[:5])   # 最多显示5个
                _add_left_text(
                    slide, col_x + Inches(0.04), row_y + Inches(0.04),
                    col_w - Inches(0.08), row_h - Inches(0.08), item_text,
                    font_size=8, color=SIMCERE_COLORS['dk1'],
                    word_wrap=True
                )

    # ── 底部备注 ──
    if note:
        note_y = margin_top + month_hd_h + len(rows) * row_h + Inches(0.1)
        txBox = slide.shapes.add_textbox(
            margin_left, note_y,
            prs.slide_width - margin_left * 2, Inches(0.3)
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = note
        p.font.size = Pt(9)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['lt2']
        p.font.italic = True


# ============================================================
# 辅助函数：文字添加快捷方式
# ============================================================
def _add_centered_text(slide, left, top, width, height, text, font_size=11, color=None, bold=False, word_wrap=False):
    """在指定区域添加居中对齐的文字"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = '微软雅黑'
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    p.alignment = PP_ALIGN.CENTER
    # 垂直居中
    try:
        tf.paragraphs[0].space_before = Pt(0)
    except:
        pass
    return txBox


def _add_left_text(slide, left, top, width, height, text, font_size=11, color=None, bold=False, word_wrap=False):
    """在指定区域添加左对齐的文字"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.name = '微软雅黑'
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    p.alignment = PP_ALIGN.LEFT
    return txBox


# ============================================================
# 新增：review_matrix — 复盘矩阵（有效动作 × 做的好/做的不好）
# 参考：科唯可-汇报任董0528（第3页核心复盘表）
# ============================================================
def build_review_matrix_slide(prs, slide, slide_spec):
    """构建复盘矩阵页面

    典型场景：基于关键动作维度，分"做的好"和"做得不好"两列，
    逐行复盘每个维度的实践情况。

    参数:
        title: 页面标题
        badge: 页面角标（如"复盘"，可选）
        category_header: 分类维度标签（如"有效动作"、"关键动作"）
        good_header: "做的好"列标题
        bad_header: "做得不好"列标题（或"优化提升"）
        rows: 行数据，每个 = {
            "action": "宣传/传播",        # 动作维度
            "note": "（高管访谈+季度复盘会）", # 动作说明（可选）
            "good": "321世界睡眠日...",   # 做得好的内容
            "bad": "科普/义诊合计2100场"  # 做得不好/需改进的
        }
        footer_note: 底部总结（可选）
    """
    title = slide_spec.get('title', '')
    badge = slide_spec.get('badge', '')
    category_header = slide_spec.get('category_header', '有效动作')
    good_header = slide_spec.get('good_header', '做的好')
    bad_header = slide_spec.get('bad_header', '做得不好')
    rows = slide_spec.get('rows', [])
    footer_note = slide_spec.get('footer_note', '')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=26, bold=True)

    # 删除所有 BODY 占位符
    remove_body_placeholders(slide)

    if not rows:
        return

    # ── 布局参数 ──
    margin_left  = Inches(0.3)
    margin_top   = Inches(1.45)
    col1_w       = Inches(2.2)     # 动作列
    col2_w       = Inches(4.85)    # 做的好列
    col3_w       = Inches(4.85)    # 做得不好列
    header_h     = Inches(0.45)    # 表头高度
    row_h_min    = Inches(0.95)    # 最小行高
    row_padding  = Inches(0.06)    # 行间距

    # 计算可用高度
    available_h = prs.slide_height - margin_top - header_h - Inches(0.8)
    n_rows = len(rows)
    calc_h = (available_h - row_padding * (n_rows - 1)) / n_rows
    row_h = max(row_h_min, min(calc_h, Inches(1.4)))

    # ── 绘制表头 ──
    # 动作列表头（深绿）
    add_shape_rect(slide, margin_left, margin_top,
                   col1_w, header_h, fill_color=SIMCERE_COLORS['accent3'])
    _add_centered_text(slide, margin_left, margin_top,
                       col1_w, header_h, category_header,
                       font_size=13, color=SIMCERE_COLORS['lt1'], bold=True)

    # 做的好表头（浅绿）
    add_shape_rect(slide, margin_left + col1_w, margin_top,
                   col2_w, header_h, fill_color=SIMCERE_COLORS['accent2'])
    _add_centered_text(slide, margin_left + col1_w, margin_top,
                       col2_w, header_h, f'✓  {good_header}',
                       font_size=13, color=SIMCERE_COLORS['lt1'], bold=True)

    # 做得不好表头（橙黄色）
    bad_bg = SIMCERE_COLORS['accent5']
    add_shape_rect(slide, margin_left + col1_w + col2_w, margin_top,
                   col3_w, header_h, fill_color=bad_bg)
    _add_centered_text(slide, margin_left + col1_w + col2_w, margin_top,
                       col3_w, header_h, f'⚠  {bad_header}',
                       font_size=13, color=SIMCERE_COLORS['dk1'], bold=True)

    # ── 角标 ──
    if badge:
        badge_x = margin_left - Inches(0.15)
        badge_y = margin_top - Inches(0.02)
        add_shape_rect(slide, badge_x, badge_y,
                       Inches(1.2), Inches(0.38),
                       fill_color=SIMCERE_COLORS['accent1'])
        _add_centered_text(slide, badge_x, badge_y,
                           Inches(1.2), Inches(0.38), badge,
                           font_size=12, color=SIMCERE_COLORS['lt1'], bold=True)

    # ── 逐行绘制 ──
    for r_idx, row in enumerate(rows):
        row_y      = margin_top + header_h + r_idx * (row_h + row_padding)
        action     = row.get('action', '')
        note       = row.get('note', '')
        good       = row.get('good', '')
        bad        = row.get('bad', '')
        row_color  = SIMCERE_COLORS['accent3'] if r_idx % 2 == 0 else SIMCERE_COLORS['accent1']

        # 动作列背景
        add_shape_rect(slide, margin_left, row_y,
                       col1_w, row_h, fill_color=row_color)
        # 动作文字
        action_text = action
        if note:
            action_text += f'\n{note}'
        _add_centered_text(slide, margin_left, row_y,
                           col1_w, row_h, action_text,
                           font_size=11, color=SIMCERE_COLORS['lt1'], bold=True,
                           word_wrap=True)

        # 做的好列
        good_bg = SIMCERE_COLORS['hk_green'] if r_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        add_shape_rect(slide, margin_left + col1_w, row_y,
                       col2_w, row_h, fill_color=good_bg)
        _add_left_text(slide, margin_left + col1_w + Inches(0.12), row_y + Inches(0.06),
                       col2_w - Inches(0.24), row_h - Inches(0.12), good,
                       font_size=10, color=SIMCERE_COLORS['dk1'], word_wrap=True)

        # 做得不好列
        bad_bg = SIMCERE_COLORS['hlk_orange'] if r_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        add_shape_rect(slide, margin_left + col1_w + col2_w, row_y,
                       col3_w, row_h, fill_color=bad_bg)
        _add_left_text(slide, margin_left + col1_w + col2_w + Inches(0.12), row_y + Inches(0.06),
                       col3_w - Inches(0.24), row_h - Inches(0.12), bad,
                       font_size=10, color=SIMCERE_COLORS['dk1'], word_wrap=True)

    # ── 底部总结 ──
    if footer_note:
        note_y = margin_top + header_h + n_rows * (row_h + row_padding) + Inches(0.1)
        txBox = slide.shapes.add_textbox(margin_left, note_y,
                                         prs.slide_width - margin_left * 2, Inches(0.35))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = footer_note
        p.font.size = Pt(9)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['lt2']
        p.font.italic = True


# ============================================================
# 新增：action_category — 分类复盘页（持续做/优化提升/开始做）
# 参考：科唯可-汇报任董0528（第5-7页分类复盘）
# ============================================================
def build_action_category_slide(prs, slide, slide_spec):
    """构建分类复盘页面

    适用于"做的好→持续做"、"做得不好→优化提升"、"没有做→开始做"
    三种分类复盘场景。

    参数:
        title: 页面标题（如"有效动作 × 做的好 --- 持续做"）
        badge: 角标（如"反思与改进"）
        intro: 页面说明文字（可选）
        categories: 分类列表，每个 = {
            "label": "行业影响力",         # 类别标签（左侧色块）
            "items": [                      # 该类别下的具体动作
                "行业会议：2026年西普会...",
                "品类领导者站位：发布白皮书..."
            ]
        }
        highlight_box: 高亮提示框（可选），= {"title": "...",  "text": "..."}
    """
    title = slide_spec.get('title', '')
    badge = slide_spec.get('badge', '')
    intro = slide_spec.get('intro', '')
    categories = slide_spec.get('categories', [])
    highlight = slide_spec.get('highlight_box', None)

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=26, bold=True)

    # 删除所有 BODY 占位符
    remove_body_placeholders(slide)

    if not categories:
        return

    # ── 布局参数 ──
    margin_left  = Inches(0.3)
    body_top     = Inches(1.35)

    # 角标
    if badge:
        add_shape_rect(slide, margin_left, body_top,
                       Inches(1.6), Inches(0.34),
                       fill_color=SIMCERE_COLORS['accent1'])
        _add_centered_text(slide, margin_left, body_top,
                           Inches(1.6), Inches(0.34), badge,
                           font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)

    # 说明文字
    current_y = body_top + Inches(0.55) if badge else body_top

    if intro:
        txBox = slide.shapes.add_textbox(margin_left, current_y,
                                         prs.slide_width - margin_left * 2, Inches(0.4))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = intro
        p.font.size = Pt(11)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk2']
        current_y += Inches(0.5)

    # ── 逐分类绘制 ──
    # 计算可用空间，自适应每个分类的高度
    total_items = sum(len(c.get('items', [])) for c in categories)
    remaining_h = prs.slide_height - current_y - Inches(0.7)
    if highlight:
        remaining_h -= Inches(0.75)
    item_h = remaining_h / max(total_items, 1)
    item_h = min(item_h, Inches(0.35))

    for c_idx, cat in enumerate(categories):
        label = cat.get('label', '')
        items = cat.get('items', [])
        cat_color = [SIMCERE_COLORS['accent3'], SIMCERE_COLORS['accent1'],
                     SIMCERE_COLORS['accent2']][c_idx % 3]
        n = len(items)
        cat_h = max(item_h * n, Inches(0.5))

        # 左侧色块标签
        label_w = Inches(1.8)
        add_shape_rect(slide, margin_left, current_y,
                       label_w, cat_h, fill_color=cat_color)
        _add_centered_text(slide, margin_left, current_y,
                           label_w, cat_h, label,
                           font_size=12, color=SIMCERE_COLORS['lt1'], bold=True,
                           word_wrap=True)

        # 右侧物品列表
        for i_idx, item_text in enumerate(items):
            row_y = current_y + i_idx * item_h
            bg = SIMCERE_COLORS['hk_green'] if i_idx % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
            add_shape_rect(slide, margin_left + label_w, row_y,
                           prs.slide_width - margin_left * 2 - label_w, item_h,
                           fill_color=bg)
            _add_left_text(slide, margin_left + label_w + Inches(0.12),
                           row_y + Inches(0.02),
                           prs.slide_width - margin_left * 2 - label_w - Inches(0.24),
                           item_h, f'• {item_text}',
                           font_size=10, color=SIMCERE_COLORS['dk1'], word_wrap=True)

        current_y += cat_h + Inches(0.08)

    # 高亮提示框
    if highlight:
        current_y += Inches(0.1)
        box_title = highlight.get('title', '')
        box_text  = highlight.get('text', '')
        add_shape_rect(slide, margin_left, current_y,
                       prs.slide_width - margin_left * 2, Inches(0.6),
                       fill_color=SIMCERE_COLORS['hlk_green'])

        if box_title:
            _add_left_text(slide, margin_left + Inches(0.12), current_y + Inches(0.02),
                           prs.slide_width - margin_left * 2 - Inches(0.24), Inches(0.3),
                           box_title, font_size=11, color=SIMCERE_COLORS['accent3'], bold=True)
            _add_left_text(slide, margin_left + Inches(0.12), current_y + Inches(0.28),
                           prs.slide_width - margin_left * 2 - Inches(0.24), Inches(0.28),
                           box_text, font_size=9, color=SIMCERE_COLORS['dk1'], word_wrap=True)
        else:
            _add_left_text(slide, margin_left + Inches(0.12), current_y + Inches(0.06),
                           prs.slide_width - margin_left * 2 - Inches(0.24), Inches(0.48),
                           box_text, font_size=10, color=SIMCERE_COLORS['dk1'], word_wrap=True)


# ============================================================
# 新增：strategy_diagram — 策略架构图（形状+箭头+分层）
# 参考：科唯可-汇报任董0528（第8页健康驿站架构图）
# ============================================================
def build_strategy_diagram_slide(prs, slide, slide_spec):
    """构建策略架构图页面

    适用于：展示多维度升级策略、系统架构、项目推进路线等。
    由上级节点 + 下级子节点 + 箭头连接组成。

    参数:
        title: 页面标题
        badge: 角标（可选）
        diagram: {
            "center": {"text": "睡眠健康驿站", "sub": "基础建设升级"},
            "dimensions": [                    # 升级维度（4-6个）
                {
                    "label": "陈列升级",       # 维度名称（箭头标签）
                    "items": ["旗舰店", "标准店", "全门店"],  # 子节点
                    "sub_label": "硬件"         # 层次标签（可选）
                },
                ...
            ],
            "pillars": [                       # 支撑策略（底部，可选）
                {"label": "平台引流", "sub": "互动体验"},
                {"label": "线上/线下科普", "sub": "检测工具"},
                ...
            ]
        }
    """
    title = slide_spec.get('title', '')
    badge = slide_spec.get('badge', '')
    diagram = slide_spec.get('diagram', {})

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=26, bold=True)

    # 删除所有 BODY 占位符
    remove_body_placeholders(slide)

    if not diagram:
        return

    center = diagram.get('center', {})
    dimensions = diagram.get('dimensions', [])
    pillars = diagram.get('pillars', [])

    # ── 布局参数 ──
    margin_left  = Inches(0.3)
    margin_top   = Inches(1.2)

    # 角标
    if badge:
        add_shape_rect(slide, margin_left, margin_top,
                       Inches(1.6), Inches(0.34),
                       fill_color=SIMCERE_COLORS['accent1'])
        _add_centered_text(slide, margin_left, margin_top,
                           Inches(1.6), Inches(0.34), badge,
                           font_size=11, color=SIMCERE_COLORS['lt1'], bold=True)

    # ── 中心节点 ──
    center_text = center.get('text', '')
    center_sub  = center.get('sub', '')
    center_x = Inches(5.2)
    center_y = margin_top + Inches(0.15)
    center_w = Inches(3.2)
    center_h = Inches(0.75) if center_sub else Inches(0.55)

    add_shape_rect(slide, center_x, center_y,
                   center_w, center_h, fill_color=SIMCERE_COLORS['accent3'],
                   line_color=SIMCERE_COLORS['accent1'], line_width=Pt(2))

    if center_sub:
        _add_centered_text(slide, center_x, center_y,
                           center_w, Inches(0.32), center_text,
                           font_size=13, color=SIMCERE_COLORS['lt1'], bold=True)
        _add_centered_text(slide, center_x, center_y + Inches(0.35),
                           center_w, Inches(0.3), center_sub,
                           font_size=9, color=RGBColor(0xC8, 0xEB, 0xD6))
    else:
        _add_centered_text(slide, center_x, center_y,
                           center_w, center_h, center_text,
                           font_size=14, color=SIMCERE_COLORS['lt1'], bold=True)

    # ── 维度节点 ──
    n_dims = len(dimensions)
    if n_dims == 0:
        return

    dim_spacing = Inches(8.5) / max(n_dims, 1)
    dim_start_x = margin_left + Inches(2.0)

    for d_idx, dim in enumerate(dimensions):
        dim_label = dim.get('label', '')
        dim_items = dim.get('items', [])
        sub_label = dim.get('sub_label', '')

        dim_x = dim_start_x + d_idx * dim_spacing
        dim_w = dim_spacing - Inches(0.15)
        dim_y = center_y + center_h + Inches(0.25)

        # 箭头标签（绿色药丸形状）
        arrow_w = Inches(1.35)
        arrow_h = Inches(0.35)
        arrow_x = dim_x + dim_w / 2 - arrow_w / 2

        # 箭头形状
        from pptx.enum.shapes import MSO_SHAPE
        arrow_shape = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON, arrow_x, dim_y, arrow_w, arrow_h
        )
        arrow_shape.fill.solid()
        arrow_shape.fill.fore_color.rgb = SIMCERE_COLORS['accent2']
        arrow_shape.line.fill.background()
        if arrow_shape.has_text_frame:
            tf = arrow_shape.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.text = dim_label
            p.font.size = Pt(10)
            p.font.name = '微软雅黑'
            p.font.bold = True
            p.font.color.rgb = SIMCERE_COLORS['lt1']
            p.alignment = PP_ALIGN.CENTER

        # 子标签
        if sub_label:
            sub_y = dim_y + arrow_h + Inches(0.05)
            _add_centered_text(slide, dim_x, sub_y,
                               dim_w, Inches(0.22), sub_label,
                               font_size=8, color=SIMCERE_COLORS['lt2'])

        # 子节点（圆角矩形）
        sub_start_y = dim_y + arrow_h + (Inches(0.35) if sub_label else Inches(0.1))
        for i_idx, item in enumerate(dim_items):
            rect_w = dim_w - Inches(0.1)
            rect_h = Inches(0.32)
            rect_x = dim_x + Inches(0.05)
            rect_y = sub_start_y + i_idx * Inches(0.4)

            from pptx.enum.shapes import MSO_SHAPE
            node = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, rect_x, rect_y, rect_w, rect_h
            )
            node.fill.solid()
            node.fill.fore_color.rgb = SIMCERE_COLORS['hk_green']
            node.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            node.line.width = Pt(0.5)

            if node.has_text_frame:
                tf = node.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = item
                p.font.size = Pt(9)
                p.font.name = '微软雅黑'
                p.font.color.rgb = SIMCERE_COLORS['dk1']
                p.alignment = PP_ALIGN.CENTER

    # ── 底部支撑策略 ──
    if pillars:
        max_dim_bottom = margin_top + center_h + Inches(0.25) + (Inches(0.35) + Inches(0.1) + 3 * Inches(0.4))
        pillar_y = min(max_dim_bottom + Inches(0.3), prs.slide_height - Inches(1.2))
        n_pillars = len(pillars)
        pillar_spacing = Inches(8.5) / max(n_pillars, 1)
        pillar_start_x = margin_left + Inches(2.0)

        # 分隔线
        from pptx.enum.shapes import MSO_SHAPE
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            margin_left, pillar_y, prs.slide_width - margin_left * 2, Inches(0.015)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = SIMCERE_COLORS['accent2']
        line.line.fill.background()

        # 标题
        _add_left_text(slide, margin_left, pillar_y + Inches(0.1),
                       Inches(1.8), Inches(0.3), '四维服务升级',
                       font_size=10, color=SIMCERE_COLORS['accent3'], bold=True)

        for p_idx, pillar in enumerate(pillars):
            p_label = pillar.get('label', '')
            p_sub   = pillar.get('sub', '')
            px = pillar_start_x + p_idx * pillar_spacing
            pw = pillar_spacing - Inches(0.15)
            py = pillar_y + Inches(0.15)

            node = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, Inches(0.55)
            )
            node.fill.solid()
            node.fill.fore_color.rgb = SIMCERE_COLORS['hk_green']
            node.line.color.rgb = SIMCERE_COLORS['accent2']
            node.line.width = Pt(1)

            if node.has_text_frame:
                tf = node.text_frame
                tf.word_wrap = True
                p1 = tf.paragraphs[0]
                p1.text = p_label
                p1.font.size = Pt(10)
                p1.font.name = '微软雅黑'
                p1.font.bold = True
                p1.font.color.rgb = SIMCERE_COLORS['dk1']
                p1.alignment = PP_ALIGN.CENTER

                if p_sub:
                    p2 = tf.add_paragraph()
                    p2.text = p_sub
                    p2.font.size = Pt(8)
                    p2.font.name = '微软雅黑'
                    p2.font.color.rgb = SIMCERE_COLORS['lt2']
                    p2.alignment = PP_ALIGN.CENTER


# ============================================================
# build_slide 分发逻辑（在文件末尾，build_slide函数内追加）
# ============================================================


def build_process_slide(prs, slide, slide_spec):
    """构建流程/阶段展示页面
    
    参数:
        steps: 步骤列表 [{"title": "2025Q4", "description": "获批上市", "detail": "NMPA批准"}, ...]
        direction: "horizontal" (默认) 或 "vertical"
    """
    title = slide_spec.get('title', '')
    steps = slide_spec.get('steps', [])
    direction = slide_spec.get('direction', 'horizontal')

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    if not steps:
        return

    if direction == 'horizontal':
        # 水平流程
        margin_top = Inches(1.8)
        node_width = Inches(2.2)
        node_height = Inches(1.2)
        gap = Inches(0.3)
        total_width = len(steps) * node_width + (len(steps) - 1) * gap
        start_x = max(Inches(0.6), (prs.slide_width - total_width) / 2)
        
        for i, step in enumerate(steps):
            node_x = start_x + i * (node_width + gap)
            
            # 节点背景（圆角矩形）
            shape_node = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, node_x, margin_top, node_width, node_height
            )
            shape_node.fill.solid()
            shape_node.fill.fore_color.rgb = RGBColor(0xE8, 0xF5, 0xEE)
            shape_node.line.color.rgb = SIMCERE_COLORS['accent1']
            shape_node.line.width = Pt(2)
            
            step_title = step.get('title', '')
            step_desc = step.get('description', '')
            
            # 标题
            t_box = slide.shapes.add_textbox(node_x + Inches(0.1), margin_top + Inches(0.08),
                                             node_width - Inches(0.2), Inches(0.3))
            tf = t_box.text_frame
            p = tf.paragraphs[0]
            p.text = step_title
            p.font.size = Pt(12)
            p.font.name = '微软雅黑'
            p.font.bold = True
            p.font.color.rgb = SIMCERE_COLORS['accent1']
            p.alignment = PP_ALIGN.CENTER
            
            # 描述
            d_box = slide.shapes.add_textbox(node_x + Inches(0.1), margin_top + Inches(0.38),
                                             node_width - Inches(0.2), Inches(0.7))
            tf = d_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = step_desc
            p.font.size = Pt(10)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['dk1']
            p.alignment = PP_ALIGN.CENTER
            
            # 箭头（除最后一个）
            if i < len(steps) - 1:
                arrow_x = node_x + node_width + Inches(0.02)
                arrow_y = margin_top + node_height / 2 - Inches(0.08)
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y, Inches(0.26), Inches(0.16)
                )
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = SIMCERE_COLORS['accent1']
                arrow.line.fill.background()


def build_kpi_dashboard_slide(prs, slide, slide_spec):
    """构建KPI指标仪表盘页面
    
    参数:
        kpis: 指标列表 [{"label": "覆盖门店", "value": "300", "unit": "家", "change": "+15%", "target": "500"}, ...]
    """
    title = slide_spec.get('title', '')
    kpis = slide_spec.get('kpis', [])

    # 设置标题
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'TITLE' in ph_type:
                set_text_frame_text(shape, title, font_size=28, bold=True)
    
    # 删除BODY占位符（先收集再删除，避免遍历时修改集合）
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split('.')[-1]
            if 'BODY' in ph_type:
                shapes_to_remove.append(shape)
    for shape in shapes_to_remove:
        remove_shape(shape)

    if not kpis:
        return

    # 网格布局
    cols = min(4, len(kpis))
    rows = (len(kpis) + cols - 1) // cols
    
    card_width = Inches(2.9)
    card_height = Inches(1.3)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)
    
    total_grid_width = cols * card_width + (cols - 1) * gap_x
    start_x = max(Inches(0.5), (prs.slide_width - total_grid_width) / 2)
    start_y = Inches(1.5)
    
    for idx, kpi in enumerate(kpis):
        row = idx // cols
        col = idx % cols
        
        card_x = start_x + col * (card_width + gap_x)
        card_y = start_y + row * (card_height + gap_y)
        
        label = kpi.get('label', '')
        value = kpi.get('value', '')
        unit = kpi.get('unit', '')
        change = kpi.get('change', '')
        target_val = kpi.get('target', '')
        
        # 卡片背景
        add_shape_rect(slide, card_x, card_y, card_width, card_height,
                       fill_color=RGBColor(0xFF, 0xFF, 0xFF),
                       line_color=RGBColor(0xDD, 0xDD, 0xDD), line_width=1)
        
        # 顶部装饰条
        add_shape_rect(slide, card_x, card_y, card_width, Inches(0.06),
                       fill_color=SIMCERE_COLORS['accent1'])
        
        # 标签
        l_box = slide.shapes.add_textbox(card_x + Inches(0.1), card_y + Inches(0.12),
                                         card_width - Inches(0.2), Inches(0.22))
        tf = l_box.text_frame
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(10)
        p.font.name = '微软雅黑'
        p.font.color.rgb = SIMCERE_COLORS['dk1']
        
        # 数值
        v_text = f"{value}{unit}"
        v_box = slide.shapes.add_textbox(card_x + Inches(0.1), card_y + Inches(0.35),
                                         card_width - Inches(0.2), Inches(0.4))
        tf = v_box.text_frame
        p = tf.paragraphs[0]
        p.text = v_text
        p.font.size = Pt(22)
        p.font.name = '微软雅黑'
        p.font.bold = True
        p.font.color.rgb = SIMCERE_COLORS['accent1']
        
        # 变化/目标
        meta_parts = []
        if change:
            meta_parts.append(change)
        if target_val:
            meta_parts.append(f"目标: {target_val}")
        
        if meta_parts:
            m_box = slide.shapes.add_textbox(card_x + Inches(0.1), card_y + Inches(0.78),
                                             card_width - Inches(0.2), Inches(0.25))
            tf = m_box.text_frame
            p = tf.paragraphs[0]
            p.text = '  |  '.join(meta_parts)
            p.font.size = Pt(9)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['lt2']


def build_slide(prs, slide_spec):
    """根据规范创建单页幻灯片"""
    slide_type = slide_spec.get('type', 'content')
    layout = get_layout_by_type(prs, slide_type)
    
    if layout is None:
        print(f"Warning: Could not find layout for type '{slide_type}'", file=sys.stderr)
        return None
    
    slide = prs.slides.add_slide(layout)
    
    # === 封面 ===
    if slide_type == 'cover':
        title = slide_spec.get('title', '')
        subtitle = slide_spec.get('subtitle', '')
        date = slide_spec.get('date', '')
        presenter = slide_spec.get('presenter', '')
        cover_body_specs = []
        
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type).split('.')[-1]
                if 'CENTER_TITLE' in ph_type:
                    set_text_frame_text(shape, title, font_size=36, bold=True)
                elif 'SUBTITLE' in ph_type:
                    set_text_frame_text(shape, subtitle, font_size=20)
                elif 'BODY' in ph_type:
                    text = date
                    if presenter:
                        text += f"\n{presenter}" if text else presenter
                    set_text_frame_text(shape, text, font_size=14)
                    if text:
                        cover_body_specs.append((shape.left, shape.top, shape.width, shape.height, text))

        for left, top, width, height, text in cover_body_specs:
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = text
            p.font.size = Pt(14)
            p.font.name = '微软雅黑'

        # 删除所有 BODY 占位符（封面不需要模板的 BODY 框）
        remove_body_placeholders(slide)
    
    # === 目录 ===
    elif slide_type == 'toc':
        title = slide_spec.get('title', '目录')
        items = slide_spec.get('items', [])
        
        # 在内容区域添加目录项
        body_ph = None
        for shape in slide.shapes:
            if not shape.is_placeholder:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    # 找到文本框
                    body_ph = shape
                    break
        
        if body_ph is None:
            # 查找BODY占位符
            body_ph = find_placeholder(slide, 'BODY')
        
        if body_ph and items:
            target_shape = body_ph
            if body_ph.is_placeholder and 'BODY' in str(body_ph.placeholder_format.type).split('.')[-1]:
                target_shape = slide.shapes.add_textbox(
                    body_ph.left, body_ph.top, body_ph.width, body_ph.height
                )

            tf = target_shape.text_frame
            tf.clear()
            for i, item in enumerate(items):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = item
                p.font.size = Pt(18)
                p.font.name = '微软雅黑'
                p.space_after = Pt(16)

        # 删除模板残留的 BODY 占位符
        remove_body_placeholders(slide)
    
    # === 节标题 ===
    elif slide_type == 'section':
        title = slide_spec.get('title', '')
        subtitle = slide_spec.get('subtitle', '')
        
        # 收集所有BODY占位符，按面积排序，只填最大的那个
        body_placeholders = []
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type).split('.')[-1]
                if 'TITLE' in ph_type:
                    set_text_frame_text(shape, title, font_size=32, bold=True)
                elif 'BODY' in ph_type:
                    area = shape.width.emu * shape.height.emu
                    body_placeholders.append((area, shape))
        
        # 删除所有 BODY 占位符
        remove_body_placeholders(slide)
        # 如果有副标题，在标题下方新建文本框
        if subtitle:
            txBox = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(11.5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = subtitle
            p.font.size = Pt(16)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['dk1']
    
    # === 标题和内容 ===
    elif slide_type == 'content':
        title = slide_spec.get('title', '')
        subtitle = slide_spec.get('subtitle', '')
        body = slide_spec.get('body', [])
        table_data = slide_spec.get('table')
        chart_spec = slide_spec.get('chart')

        # 设置标题
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type).split('.')[-1]
                if 'TITLE' in ph_type:
                    set_text_frame_text(shape, title, font_size=28, bold=True)

        # 删除所有 BODY 占位符（避免模板残留或占位）
        remove_body_placeholders(slide)
        # 如果有副标题，在标题下方新建文本框
        if subtitle:
            txBox = slide.shapes.add_textbox(Inches(0.7), Inches(1.1), Inches(11.5), Inches(0.4))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = subtitle
            p.font.size = Pt(14)
            p.font.name = '微软雅黑'
            p.font.color.rgb = SIMCERE_COLORS['dk1']

        # 正文区域：模板"标题和内容"版式中 BODY 占位符是窄条小标题区(y=0.34,h=0.40)，
        # 真正的正文区是 none 形状背景(y=1.43,h=5.28)但没有 txBody。
        # 因此在正确位置添加新文本框放正文/表格/图表。
        content_left = Inches(0.7)
        content_top = Inches(1.5)    # 标题下方，对应模板正文背景区域
        content_width = Inches(11.5)
        content_height = Inches(4.5)
        slide_bottom = prs.slide_height - Inches(0.5)

        # 判断是否有正文
        has_body = bool(body)
        # 判断是否有表格
        has_table = bool(table_data and table_data.get('headers') and table_data.get('rows'))
        # 判断是否有图表
        has_chart = bool(chart_spec and chart_spec.get('categories') and chart_spec.get('series'))

        current_top = content_top
        available_height = slide_bottom - content_top

        # 1. 先放置正文（项目符号列表）——添加到新文本框，不修改占位符
        if has_body:
            if has_table or has_chart:
                body_height = int(available_height * 0.35)
            else:
                body_height = int(available_height)

            # 添加新文本框（透明背景，不覆盖模板样式）
            txBox = slide.shapes.add_textbox(content_left, current_top, content_width, body_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            add_bullet_points_to_text_frame(tf, body)
            current_top += body_height + Inches(0.2)

        # 2. 放置表格
        if has_table:
            remaining_height = slide_bottom - current_top
            table_height = min(Inches(2.5), remaining_height * 0.9)
            table_width = content_width

            add_table_to_slide(
                slide, table_data,
                content_left, current_top,
                table_width, table_height
            )
            current_top += table_height + Inches(0.2)

        # 3. 放置图表
        if has_chart:
            remaining_height = slide_bottom - current_top
            chart_height = min(Inches(3.5), remaining_height * 0.95)
            chart_width = content_width

            add_chart_to_slide(
                slide, chart_spec,
                content_left, current_top,
                chart_width, chart_height
            )
    
    # === 甘特图 ===
    elif slide_type == 'gantt':
        build_gantt_slide(prs, slide, slide_spec)

    # === 横向时间线 ===
    elif slide_type == 'timeline_horizontal':
        build_timeline_horizontal_slide(prs, slide, slide_spec)

    # === 大数字展示 ===
    elif slide_type == 'big_number':
        build_big_number_slide(prs, slide, slide_spec)

    # === 对比页 ===
    elif slide_type == 'comparison':
        build_comparison_slide(prs, slide, slide_spec)

    # === 流程/阶段 ===
    elif slide_type == 'process':
        build_process_slide(prs, slide, slide_spec)

    # === KPI指标卡 ===
    elif slide_type == 'kpi_dashboard':
        build_kpi_dashboard_slide(prs, slide, slide_spec)

    # === 物品矩阵图 ===
    elif slide_type == 'item_matrix':
        build_item_matrix_slide(prs, slide, slide_spec)

    # === 日历网格（月度规划） ===
    elif slide_type == 'calendar_grid':
        build_calendar_grid_slide(prs, slide, slide_spec)

    # === 复盘矩阵（有效动作×做的好/做的不好） ===
    elif slide_type == 'review_matrix':
        build_review_matrix_slide(prs, slide, slide_spec)

    # === 分类复盘页（持续做/优化提升/开始做） ===
    elif slide_type == 'action_category':
        build_action_category_slide(prs, slide, slide_spec)

    # === 策略架构图 ===
    elif slide_type == 'strategy_diagram':
        build_strategy_diagram_slide(prs, slide, slide_spec)

    # === 仅标题页 ===
    elif slide_type == 'title_only':
        title = slide_spec.get('title', '')
        set_placeholder_text(slide, 'TITLE', title)

    # === 空白页 ===
    elif slide_type == 'blank':
        # 根据shapes添加自定义元素
        shapes_spec = slide_spec.get('shapes', [])
        for shape_spec in shapes_spec:
            shape_type = shape_spec.get('shape_type', 'text')
            left = Inches(shape_spec.get('left', 0.5))
            top = Inches(shape_spec.get('top', 0.5))
            width = Inches(shape_spec.get('width', 4))
            height = Inches(shape_spec.get('height', 1))
            text = shape_spec.get('text', '')
            
            if shape_type == 'text':
                add_text_box(slide, left, top, width, height, text,
                           font_size=shape_spec.get('font_size', 14),
                           bold=shape_spec.get('bold', False))
    
    # === 结尾页 ===
    elif slide_type == 'ending':
        title = slide_spec.get('title', '谢谢')
        subtitle = slide_spec.get('subtitle', 'Thank You')
        
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type).split('.')[-1]
                if 'CENTER_TITLE' in ph_type:
                    set_text_frame_text(shape, title, font_size=40, bold=True)
                elif 'BODY' in ph_type:
                    set_text_frame_text(shape, subtitle, font_size=18)

        # ending 页保留第一个 BODY（写副标题用），删除多余的
        shapes_to_remove = []
        body_count = 0
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type).split('.')[-1]
                if 'BODY' in ph_type:
                    body_count += 1
                    if body_count > 1:
                        shapes_to_remove.append(shape)
        for shape in shapes_to_remove:
            remove_shape(shape)
    
    return slide


def generate_ppt(template_path, content_spec, output_path):
    """生成PPT的主函数"""
    import shutil
    import tempfile
    import zipfile
    import os
    
    # 复制模板到临时文件（保留原始模板不变）
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.pptx')
    os.close(tmp_fd)
    shutil.copy2(template_path, tmp_path)
    
    try:
        # 加载临时模板文件
        prs = Presentation(tmp_path)
        
        # 获取内容规范
        slides_spec = content_spec.get('slides', [])
        
        if not slides_spec:
            print("Error: No slides specified in content", file=sys.stderr)
            sys.exit(1)
        
        pres_part = prs.part
        
        # === 删除所有原有幻灯片 ===
        # 收集所有旧slide的rId
        orig_slide_ids = []
        for sldId in list(prs.slides._sldIdLst):
            rId = sldId.get(qn('r:id'))
            orig_slide_ids.append((sldId, rId))
        
        # 移除原有幻灯片（从sldIdLst和rels中）
        for sldId, rId in orig_slide_ids:
            prs.slides._sldIdLst.remove(sldId)
            if rId:
                pres_part.drop_rel(rId)
        
        # === 添加所有新幻灯片 ===
        new_slides = []
        for slide_spec in slides_spec:
            slide = build_slide(prs, slide_spec)
            if slide is not None:
                new_slides.append(slide)
        
        # === 保存到另一个临时文件 ===
        tmp_fd2, tmp_path2 = tempfile.mkstemp(suffix='.pptx')
        os.close(tmp_fd2)
        
        try:
            prs.save(tmp_path2)
            
            # 获取当前所有有效slide的文件名
            import os
            valid_slide_names = set()
            for sldId in prs.slides._sldIdLst:
                rId = sldId.get(qn('r:id'))
                if rId:
                    rel = pres_part.rels.get(rId)
                    if rel:
                        # rel.target_ref 如 'slides/slide1.xml'
                        valid_slide_names.add('ppt/' + rel.target_ref)
                        # slide 关系文件路径: ppt/slides/_rels/slide1.xml.rels
                        slide_basename = os.path.basename(rel.target_ref)  # e.g. slide1.xml
                        slide_dir = os.path.dirname(rel.target_ref)        # e.g. slides
                        valid_slide_names.add('ppt/' + slide_dir + '/_rels/' + slide_basename + '.rels')
            
            # 重新打包zip：跳过旧的slide文件，保留所有其他文件
            with zipfile.ZipFile(tmp_path2, 'r') as zin:
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    seen = set()
                    for item in zin.infolist():
                        # 检查是否是旧的slide文件
                        is_old_slide = False
                        if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                            if item.filename not in valid_slide_names:
                                is_old_slide = True
                        elif item.filename.startswith('ppt/slides/_rels/slide') and item.filename.endswith('.xml.rels'):
                            if item.filename not in valid_slide_names:
                                is_old_slide = True
                        
                        if is_old_slide:
                            continue
                        
                        # 避免重复（同名文件只保留一个）
                        if item.filename in seen:
                            continue
                        seen.add(item.filename)
                        
                        zout.writestr(item, zin.read(item.filename))
            
            print(f"Generated: {output_path}")
            return output_path
        finally:
            if os.path.exists(tmp_path2):
                os.unlink(tmp_path2)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def main():
    parser = argparse.ArgumentParser(description='先声药业集团PPT生成器')
    parser.add_argument('--template', required=True, help='模板PPT文件路径')
    parser.add_argument('--input', required=True, help='内容JSON文件路径')
    parser.add_argument('--output', required=True, help='输出PPT文件路径')
    
    args = parser.parse_args()
    
    # 读取内容规范
    with open(args.input, 'r', encoding='utf-8') as f:
        content_spec = json.load(f)
    
    generate_ppt(args.template, content_spec, args.output)


if __name__ == '__main__':
    main()
