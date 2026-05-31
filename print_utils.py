"""print_utils.py — 财务报表 PDF 导出工具

供 pages/report.py 调用。
使用 reportlab 将任意 QTableWidget 内容渲染为 A4 PDF，
支持：纵向/横向、自动分页、section-header/total/normal 行样式、负数红色。
"""

import os
from datetime import datetime

# ── 字体缓存 ──────────────────────────────────────────────────────────────────
_FONT_NAME: str | None = None   # 已注册的字体名


def _ensure_font() -> str:
    """加载一次中文字体，返回已注册的 font name。优先 Windows，次选 macOS，兜底 CID。"""
    global _FONT_NAME
    if _FONT_NAME:
        return _FONT_NAME

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    candidates = [
        # Windows
        ("MSYaHei",     "C:/Windows/Fonts/msyh.ttc",                    0),
        ("SimHei",      "C:/Windows/Fonts/simhei.ttf",                   0),
        ("SimSun",      "C:/Windows/Fonts/simsun.ttc",                   0),
        ("FangSong",    "C:/Windows/Fonts/fangsong.ttf",                 0),
        ("KaiTi",       "C:/Windows/Fonts/kaiti.ttf",                    0),
        # macOS
        ("SongtiSC",    "/System/Library/Fonts/Supplemental/Songti.ttc", 0),
        ("PingFangSC",  "/System/Library/Fonts/PingFang.ttc",            0),
        ("STHeitiSC",   "/System/Library/Fonts/STHeiti Light.ttc",       0),
    ]
    for name, path, idx in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path, subfontIndex=idx))
                _FONT_NAME = name
                return name
            except Exception:
                continue

    # CID 兜底（reportlab 内置）
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        _FONT_NAME = "STSong-Light"
        return _FONT_NAME
    except Exception:
        pass

    _FONT_NAME = "Helvetica"
    return _FONT_NAME


# ── 行样式颜色常量（RGB 0..1）────────────────────────────────────────────────
# 列标题行
_C_COL_HDR_BG  = (0.110, 0.137, 0.251)   # #1C2340 深蓝
_C_COL_HDR_FG  = (1.000, 1.000, 1.000)   # 白色

# Section-header 行（行次为空的分类标题）
_C_SEC_BG      = (0.941, 0.957, 1.000)   # #f0f4ff 淡蓝
_C_SEC_FG      = (0.239, 0.431, 0.855)   # #3d6fdb 蓝色

# 合计/小计行（字体加粗）
_C_TOT_BG      = (0.961, 0.969, 0.980)   # #f5f7fa 浅灰

# 普通行交替背景
_C_ROW_EVEN    = (1.000, 1.000, 1.000)   # 白
_C_ROW_ODD     = (0.984, 0.984, 0.992)   # 极浅灰紫

# 负数金额
_C_NEG         = (0.878, 0.322, 0.322)   # #e05252 红

# 文字主色
_C_TEXT        = (0.098, 0.118, 0.196)   # #191e32 近黑


def _qcolor_to_rgb(qcolor):
    """QColor → (r, g, b) 0..1 元组；无效颜色返回 None。"""
    if not qcolor.isValid():
        return None
    return (qcolor.redF(), qcolor.greenF(), qcolor.blueF())


def _is_section_bg(rgb):
    """判断是否为 section-header 背景色（蓝调：b 明显高于 r/g）。"""
    if rgb is None:
        return False
    r, g, b = rgb
    return b > 0.95 and b - r > 0.04 and b - g > 0.02


def _is_total_bg(rgb):
    """判断是否为合计行背景色（非白但非蓝调的浅灰）。"""
    if rgb is None:
        return False
    r, g, b = rgb
    # 三色接近但不是纯白，且不是蓝调
    return (0.93 < r < 0.99) and (0.93 < g < 0.99) and not _is_section_bg(rgb)


def _truncate_text(text: str, font: str, size: float, max_w: float, pm) -> str:
    """截断文字加省略号以适应宽度。"""
    if not text:
        return ""
    if pm.stringWidth(text, font, size) <= max_w:
        return text
    while text and pm.stringWidth(text + "…", font, size) > max_w:
        text = text[:-1]
    return (text + "…") if text else "…"


# ── 主导出函数 ─────────────────────────────────────────────────────────────────

def export_report_pdf(
    path: str,
    table_widget,
    col_headers: list[str],
    col_ratios: list[float],
    company_name: str,
    report_title: str,
    period_text: str,
    is_landscape: bool = False,
    amount_col_indices: list[int] | None = None,
) -> None:
    """
    将 QTableWidget 导出为 A4 PDF 财务报表。

    Parameters
    ----------
    path               : 输出文件路径（.pdf）
    table_widget       : 已填充数据的 QTableWidget
    col_headers        : 列标题列表
    col_ratios         : 各列宽度比例（之和应≈1.0）
    company_name       : 账套/公司名称
    report_title       : 报表名称，如"资产负债表"
    period_text        : 期间描述，如"2026-01 至 2026-05"
    is_landscape       : True → A4横向；False → A4纵向
    amount_col_indices : 金额列索引（右对齐）
    """
    from reportlab.lib.pagesizes import A4, landscape as rl_landscape
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.pdfbase import pdfmetrics

    if amount_col_indices is None:
        amount_col_indices = []

    font = _ensure_font()
    pm   = pdfmetrics

    # ── 页面几何 ──
    page_size  = rl_landscape(A4) if is_landscape else A4
    pw, ph     = page_size
    ML, MR     = 36, 36        # 左右边距
    MT, MB     = 32, 28        # 上下边距
    content_w  = pw - ML - MR

    col_widths = [content_w * r for r in col_ratios]

    ROW_H      = 22            # 数据行高
    COL_HDR_H  = 28            # 列标题行高
    PAGE_HDR_H = 78            # 页眉区高度（公司名+报表名+期间+分割线）
    FOOTER_H   = 18

    avail_body_h        = ph - MT - MB - PAGE_HDR_H - COL_HDR_H - FOOTER_H
    max_rows_per_page   = max(1, int(avail_body_h / ROW_H))

    # ── 读取 QTableWidget 数据 ──
    n_rows = table_widget.rowCount()
    rows_data = []

    for ri in range(n_rows):
        texts    = []
        bg_rgb   = None
        is_bold  = False
        neg_cols = set()

        for ci in range(table_widget.columnCount()):
            item = table_widget.item(ri, ci)
            if item is None:
                texts.append("")
                continue
            texts.append(item.text())

            if ci == 0:
                # 从第一列读取行级属性
                brush = item.background()
                from PySide6.QtCore import Qt
                if brush.style() != Qt.BrushStyle.NoBrush:
                    bg_rgb = _qcolor_to_rgb(brush.color())
                f = item.font()
                is_bold = f.bold() or f.weight() >= 63   # QFont::Bold = 75

            if ci in amount_col_indices:
                fg = item.foreground().color()
                if fg.isValid() and fg.red() > 180 and fg.green() < 100:
                    neg_cols.add(ci)

        rows_data.append({
            "texts":   texts,
            "bg":      bg_rgb,
            "bold":    is_bold,
            "neg":     neg_cols,
        })

    # ── 分页 ──
    pages: list[list[dict]] = [
        rows_data[i : i + max_rows_per_page]
        for i in range(0, max(n_rows, 1), max_rows_per_page)
    ] or [[]]
    total_pages = len(pages)

    # ── 渲染 ──
    cv = rl_canvas.Canvas(path, pagesize=page_size)
    today_str = datetime.now().strftime("%Y-%m-%d")

    def _draw_page_header() -> float:
        """绘制页眉，返回页眉底部 y 坐标。"""
        base_y = ph - MT
        cv.setFillColorRGB(*_C_TEXT)
        # 公司名
        cv.setFont(font, 15)
        cv.drawCentredString(pw / 2, base_y - 18, company_name)
        # 报表标题
        cv.setFont(font, 13)
        cv.drawCentredString(pw / 2, base_y - 36, report_title)
        # 期间
        cv.setFont(font, 9.5)
        cv.setFillColorRGB(0.4, 0.4, 0.4)
        cv.drawCentredString(pw / 2, base_y - 52, f"期间：{period_text}")
        # 分割线
        cv.setStrokeColorRGB(*_C_TEXT)
        cv.setLineWidth(1.2)
        cv.line(ML, base_y - 62, pw - MR, base_y - 62)
        return base_y - PAGE_HDR_H

    def _draw_col_headers(y_top: float) -> float:
        """绘制列标题行，返回底部 y 坐标。"""
        y_bot = y_top - COL_HDR_H
        # 背景
        cv.setFillColorRGB(*_C_COL_HDR_BG)
        cv.rect(ML, y_bot, content_w, COL_HDR_H, stroke=0, fill=1)
        # 文字
        cv.setFillColorRGB(*_C_COL_HDR_FG)
        cv.setFont(font, 9.5)
        x = ML
        for hdr, cw in zip(col_headers, col_widths):
            cv.drawCentredString(x + cw / 2, y_bot + 9, hdr)
            x += cw
        # 外框
        cv.setStrokeColorRGB(0.5, 0.5, 0.5)
        cv.setLineWidth(0.6)
        cv.rect(ML, y_bot, content_w, COL_HDR_H, stroke=1, fill=0)
        return y_bot

    def _draw_rows(rows: list[dict], y_top: float) -> None:
        """绘制数据行区域。"""
        y = y_top
        table_top = y_top

        for row_idx, row in enumerate(rows):
            bg  = row["bg"]
            is_sec = _is_section_bg(bg)
            is_tot = (not is_sec) and (_is_total_bg(bg) or row["bold"])
            fs  = 9.5 if (is_sec or is_tot) else 9

            # 行背景
            if is_sec:
                cv.setFillColorRGB(*_C_SEC_BG)
            elif is_tot:
                cv.setFillColorRGB(*_C_TOT_BG)
            else:
                cv.setFillColorRGB(*(_C_ROW_EVEN if row_idx % 2 == 0 else _C_ROW_ODD))
            cv.rect(ML, y - ROW_H, content_w, ROW_H, stroke=0, fill=1)

            # 单元格文字
            cv.setFont(font, fs)
            x = ML
            for ci, (text, cw) in enumerate(zip(row["texts"], col_widths)):
                inner_w = cw - 8
                trunc = _truncate_text(str(text), font, fs, inner_w, pm)

                if is_sec and ci == 0:
                    cv.setFillColorRGB(*_C_SEC_FG)
                elif ci in row["neg"]:
                    cv.setFillColorRGB(*_C_NEG)
                else:
                    cv.setFillColorRGB(*_C_TEXT)

                text_y = y - ROW_H + (ROW_H - fs) / 2 - 1

                if ci in amount_col_indices:
                    cv.drawRightString(x + cw - 5, text_y, trunc)
                else:
                    # 缩进：以空格开头的文字保留视觉缩进
                    leading_spaces = len(text) - len(text.lstrip(" "))
                    indent = 5 + leading_spaces * 3
                    cv.drawString(x + indent, text_y, trunc)
                x += cw

            # 行分割线
            cv.setStrokeColorRGB(0.88, 0.88, 0.88)
            cv.setLineWidth(0.3)
            cv.line(ML, y - ROW_H, pw - MR, y - ROW_H)
            y -= ROW_H

        # 表格外框 + 列竖线
        table_h = len(rows) * ROW_H
        cv.setStrokeColorRGB(0.5, 0.5, 0.5)
        cv.setLineWidth(0.6)
        cv.rect(ML, y_top - table_h, content_w, table_h, stroke=1, fill=0)
        x = ML
        for cw in col_widths[:-1]:
            x += cw
            cv.line(x, table_top, x, y_top - table_h)

    def _draw_footer(page_num: int) -> None:
        cv.setFont(font, 8.5)
        cv.setFillColorRGB(0.5, 0.5, 0.5)
        cv.drawCentredString(
            pw / 2, MB - 8,
            f"第 {page_num} 页 / 共 {total_pages} 页    打印日期：{today_str}",
        )

    # ── 逐页输出 ──
    for page_num, page_rows in enumerate(pages, 1):
        if page_num > 1:
            cv.showPage()
        y = _draw_page_header()
        y = _draw_col_headers(y)
        _draw_rows(page_rows, y)
        _draw_footer(page_num)

    cv.save()
