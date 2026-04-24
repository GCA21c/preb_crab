from __future__ import annotations

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QFontMetricsF, QGuiApplication, QImage, QPainter, QPen

from core.hwp_types import (
    HwpBorderFillModel,
    HwpCharShapeModel,
    HwpDocumentModel,
    HwpPageModel,
    HwpParaShapeModel,
    HwpParagraphModel,
    HwpParagraphRun,
    HwpTableCellModel,
    HwpTableModel,
    HwpTableRowModel,
    hwp_to_px,
)

FONT_FALLBACKS = (
    'Malgun Gothic',
    'Apple SD Gothic Neo',
    'Noto Sans CJK KR',
    'NanumGothic',
    'Arial Unicode MS',
)

HWP_BORDER_WIDTH_MM = {
    0: 0.10,
    1: 0.12,
    2: 0.15,
    3: 0.20,
    4: 0.25,
    5: 0.30,
    6: 0.40,
    7: 0.50,
    8: 0.60,
    9: 0.70,
    10: 1.00,
    11: 1.50,
    12: 2.00,
    13: 3.00,
    14: 4.00,
    15: 5.00,
}


def render_hwp_document_pages(model: HwpDocumentModel, dpi: float = 192.0) -> list[QImage]:
    font_family = _pick_font_family(model.source.fonts)
    return [
        render_hwp_page(
            page,
            dpi=dpi,
            font_family=font_family,
            char_shapes=model.source.char_shapes,
            border_fills=model.source.border_fills,
            para_shapes=model.source.para_shapes,
        )
        for page in model.pages
    ]


def render_hwp_page(
    page: HwpPageModel,
    dpi: float = 192.0,
    font_family: str | None = None,
    char_shapes: list[HwpCharShapeModel] | None = None,
    border_fills: list[HwpBorderFillModel] | None = None,
    para_shapes: list[HwpParaShapeModel] | None = None,
) -> QImage:
    width = max(1, int(round(hwp_to_px(page.size.width_hwp, dpi=dpi))))
    height = max(1, int(round(hwp_to_px(page.size.height_hwp, dpi=dpi))))
    image = QImage(width, height, QImage.Format_ARGB32)
    image.fill(Qt.white)
    painter = QPainter(image)
    try:
        _paint_page_background(painter, width, height)
        _paint_page_content(
            painter,
            page,
            width,
            height,
            dpi=dpi,
            font_family=font_family,
            char_shapes=char_shapes or [],
            border_fills=border_fills or [],
            para_shapes=para_shapes or [],
        )
    finally:
        painter.end()
    return image


def _paint_page_background(painter: QPainter, width: int, height: int) -> None:
    painter.fillRect(0, 0, width, height, QColor(Qt.white))
    pen = QPen(QColor('#d9d9d9'))
    pen.setWidth(1)
    painter.setPen(pen)
    painter.drawRect(0, 0, width - 1, height - 1)


def _paint_page_content(
    painter: QPainter,
    page: HwpPageModel,
    width: int,
    height: int,
    *,
    dpi: float,
    font_family: str | None,
    char_shapes: list[HwpCharShapeModel],
    border_fills: list[HwpBorderFillModel],
    para_shapes: list[HwpParaShapeModel],
) -> None:
    scale = width / max(1.0, page.size.width_px)
    left = max(96.0 * scale, hwp_to_px(page.margins.left_hwp, dpi=192.0) or 0.0)
    top = max(96.0 * scale, hwp_to_px(page.margins.top_hwp, dpi=192.0) or 0.0)
    right = max(96.0 * scale, hwp_to_px(page.margins.right_hwp, dpi=192.0) or 0.0)
    bottom = max(96.0 * scale, hwp_to_px(page.margins.bottom_hwp, dpi=192.0) or 0.0)
    text_rect = QRectF(left, top, max(100.0, width - left - right), max(100.0, height - top - bottom))
    blocks = page.flow_blocks or []
    context = {
        'border_fills': border_fills,
        'char_shapes': char_shapes,
        'dpi': dpi,
        'font_family': font_family or _pick_font_family(),
        'line_height': 34.0 * scale,
        'para_shapes': para_shapes,
        'scale': scale,
    }
    y = text_rect.top()

    if not blocks:
        for para in page.paragraphs:
            text = ''.join(run.text for run in para.runs).strip()
            if not text:
                continue
            consumed = _draw_basic_paragraph(painter, para, text_rect, y, context)
            if consumed <= 0:
                break
            y += consumed
        return

    for index, block in enumerate(blocks):
        if block.kind == 'paragraph':
            para = block.payload
            text = ''.join(run.text for run in para.runs).strip()
            if not text:
                continue
            consumed = _draw_paragraph_block(
                painter,
                para,
                para.style_name,
                text_rect,
                y,
                context,
                is_first_block=index == 0,
            )
            if consumed <= 0:
                break
            y += consumed
            continue
        if block.kind == 'table':
            consumed = _draw_table(
                painter,
                block.payload,
                text_rect.left(),
                y,
                text_rect.width(),
                text_rect.bottom(),
                context,
            )
            if consumed <= 0:
                break
            y += consumed + (18.0 * scale)


def _draw_paragraph_block(
    painter: QPainter,
    para: HwpParagraphModel,
    style_name: str | None,
    text_rect: QRectF,
    y: float,
    context: dict[str, object],
    *,
    is_first_block: bool,
) -> float:
    return _draw_basic_paragraph(painter, para, text_rect, y, context)


def _draw_basic_paragraph(
    painter: QPainter,
    para: HwpParagraphModel,
    text_rect: QRectF,
    y: float,
    context: dict[str, object],
) -> float:
    scale = float(context['scale'])
    dpi = float(context['dpi'])
    text = ''.join(run.text for run in para.runs).strip()
    char_shapes = context.get('char_shapes')
    para_shapes = context.get('para_shapes')
    char_shape = _lookup_char_shape(para, char_shapes if isinstance(char_shapes, list) else [])
    para_shape = _lookup_para_shape(para, para_shapes if isinstance(para_shapes, list) else [])
    draw_rect = _paragraph_draw_rect(text_rect, para_shape, dpi)
    spacing_before = hwp_to_px(para_shape.prev_spacing_hwp, dpi=dpi) if para_shape else 0.0
    spacing_after = hwp_to_px(para_shape.next_spacing_hwp, dpi=dpi) if para_shape else 0.0
    draw_y = y + max(0.0, spacing_before)
    font = _make_paragraph_font(para, char_shape, str(context['font_family']), scale, dpi)
    content_height = _wrapped_text_height(font, draw_rect.width(), text)
    segment_height = _paragraph_segment_height(para, dpi)
    line_height = max(float(context['line_height']), content_height + (6.0 * scale), segment_height)
    total_height = max(0.0, spacing_before) + line_height + max(0.0, spacing_after)
    if y + total_height > text_rect.bottom():
        return 0.0
    painter.save()
    painter.setPen(_qcolor_from_colorref(char_shape.text_color) if char_shape else QColor(Qt.black))
    painter.setFont(font)
    alignment = _paragraph_alignment_flags(para_shape)
    if para.line_segments:
        _draw_segmented_paragraph(painter, para, text, draw_rect, draw_y, line_height, dpi, alignment)
    else:
        line_rect = QRectF(draw_rect.left(), draw_y, draw_rect.width(), line_height)
        painter.drawText(line_rect, int(alignment | Qt.AlignVCenter | Qt.TextWordWrap), text)
    painter.restore()
    return total_height


def _make_paragraph_font(
    para: HwpParagraphModel,
    char_shape: HwpCharShapeModel | None,
    family: str,
    scale: float,
    dpi: float,
) -> QFont:
    if char_shape and char_shape.base_size > 0:
        point_size = max(8.0, char_shape.base_size / 100.0 * scale)
        font = _make_font(family, point_size, bold=char_shape.bold)
        font.setItalic(char_shape.italic)
        return font
    text_heights = [segment.text_height_hwp for segment in para.line_segments if segment.text_height_hwp > 0]
    if text_heights:
        avg_text_height_px = hwp_to_px(sum(text_heights) / len(text_heights), dpi=dpi)
        point_size = max(8.0, avg_text_height_px * 72.0 / dpi)
        return _make_font(family, point_size, bold=False)
    return _make_font(family, 14.5 * scale, bold=False)


def _lookup_char_shape(para: HwpParagraphModel, char_shapes: list[HwpCharShapeModel]) -> HwpCharShapeModel | None:
    shape_id = None
    if para.char_shape_spans:
        shape_id = para.char_shape_spans[0].char_shape_id
    elif para.runs:
        shape_id = para.runs[0].char_shape_id
    if shape_id is None or shape_id < 0 or shape_id >= len(char_shapes):
        return None
    return char_shapes[shape_id]


def _lookup_para_shape(para: HwpParagraphModel, para_shapes: list[HwpParaShapeModel]) -> HwpParaShapeModel | None:
    shape_id = para.para_shape_id
    if shape_id is None or shape_id < 0 or shape_id >= len(para_shapes):
        return None
    return para_shapes[shape_id]


def _paragraph_draw_rect(text_rect: QRectF, para_shape: HwpParaShapeModel | None, dpi: float) -> QRectF:
    if para_shape is None:
        return text_rect
    left = text_rect.left() + hwp_to_px(para_shape.left_margin_hwp + para_shape.indent_hwp, dpi=dpi)
    right = text_rect.right() - hwp_to_px(para_shape.right_margin_hwp, dpi=dpi)
    if right <= left:
        return text_rect
    return QRectF(left, text_rect.top(), right - left, text_rect.height())


def _paragraph_alignment_flags(para_shape: HwpParaShapeModel | None) -> Qt.AlignmentFlag:
    if para_shape is None:
        return Qt.AlignLeft
    align = (para_shape.attrs1 >> 2) & 0x7
    if align == 1:
        return Qt.AlignLeft
    if align == 2:
        return Qt.AlignRight
    if align == 3:
        return Qt.AlignHCenter
    return Qt.AlignLeft


def _qcolor_from_colorref(value: int) -> QColor:
    red = value & 0xFF
    green = (value >> 8) & 0xFF
    blue = (value >> 16) & 0xFF
    return QColor(red, green, blue)


def _paragraph_segment_height(para: HwpParagraphModel, dpi: float) -> float:
    if not para.line_segments:
        return 0.0
    min_y = min(segment.vertical_pos_hwp for segment in para.line_segments)
    max_y = max(
        segment.vertical_pos_hwp + max(segment.line_height_hwp, segment.text_height_hwp)
        for segment in para.line_segments
    )
    return max(0.0, hwp_to_px(max_y - min_y, dpi=dpi))


def _draw_segmented_paragraph(
    painter: QPainter,
    para: HwpParagraphModel,
    text: str,
    text_rect: QRectF,
    y: float,
    line_height: float,
    dpi: float,
    alignment: Qt.AlignmentFlag,
) -> None:
    segments = sorted(para.line_segments, key=lambda segment: segment.text_start)
    starts = [max(0, min(len(text), segment.text_start)) for segment in segments]
    min_vertical = min(segment.vertical_pos_hwp for segment in segments)
    for index, segment in enumerate(segments):
        start = starts[index]
        end = starts[index + 1] if index + 1 < len(starts) else len(text)
        segment_text = text[start:end].strip()
        if not segment_text:
            continue
        segment_x = text_rect.left() + max(0.0, hwp_to_px(segment.column_start_hwp, dpi=dpi))
        segment_y = y + max(0.0, hwp_to_px(segment.vertical_pos_hwp - min_vertical, dpi=dpi))
        segment_width = hwp_to_px(segment.segment_width_hwp, dpi=dpi) if segment.segment_width_hwp > 0 else text_rect.width()
        segment_rect = QRectF(
            segment_x,
            segment_y,
            min(text_rect.right() - segment_x, max(1.0, segment_width)),
            max(1.0, min(line_height, hwp_to_px(max(segment.line_height_hwp, segment.text_height_hwp), dpi=dpi))),
        )
        painter.drawText(segment_rect, int(alignment | Qt.AlignVCenter | Qt.TextWordWrap), segment_text)


def _draw_table(
    painter: QPainter,
    table: HwpTableModel,
    x: float,
    y: float,
    max_width: float,
    max_bottom: float,
    context: dict[str, object],
) -> float:
    if not table.rows:
        return 0.0
    scale = float(context['scale'])
    dpi = float(context['dpi'])
    border_fills = context.get('border_fills')
    table_border_fill = _lookup_border_fill(table.border_fill_id, border_fills if isinstance(border_fills, list) else [])
    col_count = max((len(row.cells) for row in table.rows), default=1)
    col_widths = _column_widths(table, max_width)
    top_y = y
    pen = QPen(QColor('#7a7a7a'))
    pen.setWidth(1)
    painter.save()
    painter.setPen(pen)
    base_font = _make_font(context['font_family'], 13.5 * scale, bold=False)
    painter.setFont(base_font)
    row_heights = [
        _table_row_height(row, col_count, col_widths, base_font, scale)
        for row in table.rows
    ]
    row_positions = [y]
    current_row_y = y
    for row_height in row_heights:
        current_row_y += row_height
        row_positions.append(current_row_y)
    for row_index, row in enumerate(table.rows):
        row_height = row_heights[row_index]
        if row_positions[row_index + 1] > max_bottom:
            break
        positions = _column_positions(x, col_widths)
        if len(row.cells) == 1 and row.cells[0].col_index is None:
            cells_to_draw = [(row.cells[0], 0, col_count)]
        else:
            cells_to_draw = []
            dense_col = 0
            for cell in row.cells:
                start_col = cell.col_index if cell.col_index is not None else dense_col
                span = max(1, cell.col_span)
                cells_to_draw.append((cell, start_col, span))
                dense_col = start_col + span
        for cell, start_col, span in cells_to_draw:
            if start_col >= col_count:
                continue
            end_col = min(col_count, start_col + span)
            current_x = positions[start_col]
            width = positions[end_col] - current_x
            if width <= 0:
                continue
            row_span = max(1, cell.row_span)
            end_row = min(len(row_heights), row_index + row_span)
            rect = QRectF(current_x, row_positions[row_index], width, row_positions[end_row] - row_positions[row_index])
            cell_border_fill = _lookup_border_fill(cell.border_fill_id, border_fills if isinstance(border_fills, list) else [])
            border_fill = cell_border_fill or table_border_fill
            _fill_table_cell(painter, rect, border_fill)
            _draw_table_cell_borders(painter, rect, border_fill, dpi)
            _draw_table_cell_text(painter, cell, rect, base_font, scale, context)
        y = row_positions[row_index + 1]
    painter.restore()
    return max(0.0, y - top_y)


def _column_widths(table: HwpTableModel, max_width: float) -> list[float]:
    col_count = _table_column_count(table)
    if col_count == 1:
        return [max_width]
    measured = [0.0] * col_count
    for row in table.rows:
        col_cursor = 0
        for cell in row.cells:
            span = max(1, cell.col_span)
            if cell.width_hwp and cell.width_hwp > 0 and span == 1 and col_cursor < col_count:
                measured[col_cursor] = max(measured[col_cursor], hwp_to_px(cell.width_hwp, dpi=192.0))
            col_cursor += span
    total_measured = sum(measured)
    if total_measured > 0:
        scale = max_width / total_measured
        widths = [width * scale if width > 0 else 0.0 for width in measured]
        remainder_indexes = [i for i, width in enumerate(widths) if width <= 0]
        used = sum(widths)
        if remainder_indexes:
            fallback = max(1.0, (max_width - used) / len(remainder_indexes))
            for index in remainder_indexes:
                widths[index] = fallback
        return widths
    return [max_width / max(1, col_count)] * col_count


def _lookup_border_fill(border_fill_id: int | None, border_fills: list[HwpBorderFillModel]) -> HwpBorderFillModel | None:
    if border_fill_id is None or border_fill_id < 0 or border_fill_id >= len(border_fills):
        return None
    return border_fills[border_fill_id]


def _fill_table_cell(painter: QPainter, rect: QRectF, border_fill: HwpBorderFillModel | None) -> None:
    if border_fill is None or border_fill.fill_back_color is None:
        return
    color = _qcolor_from_colorref(border_fill.fill_back_color)
    if color.alpha() == 0:
        return
    painter.fillRect(rect, color)


def _draw_table_cell_borders(
    painter: QPainter,
    rect: QRectF,
    border_fill: HwpBorderFillModel | None,
    dpi: float,
) -> None:
    if border_fill is None:
        painter.drawRect(rect)
        return
    lines = (
        (rect.left(), rect.top(), rect.left(), rect.bottom()),
        (rect.right(), rect.top(), rect.right(), rect.bottom()),
        (rect.left(), rect.top(), rect.right(), rect.top()),
        (rect.left(), rect.bottom(), rect.right(), rect.bottom()),
    )
    for index, line in enumerate(lines):
        line_type = border_fill.line_types[index] if index < len(border_fill.line_types) else 0
        line_width_code = border_fill.line_widths[index] if index < len(border_fill.line_widths) else 0
        line_color = border_fill.line_colors[index] if index < len(border_fill.line_colors) else 0
        if line_width_code == 255 or line_type == 255:
            continue
        color = _qcolor_from_colorref(line_color)
        if color == QColor(Qt.white):
            continue
        pen = QPen(color)
        pen.setWidthF(_border_width_px(line_width_code, dpi))
        pen.setStyle(_border_pen_style(line_type))
        painter.setPen(pen)
        painter.drawLine(*line)


def _draw_table_cell_text(
    painter: QPainter,
    cell: HwpTableCellModel,
    rect: QRectF,
    base_font: QFont,
    scale: float,
    context: dict[str, object],
) -> None:
    cell_text = ' '.join(
        ''.join(run.text for run in para.runs).strip()
        for para in cell.paragraphs
    ).strip()
    if not cell_text:
        return
    para_shapes = context.get('para_shapes')
    para_shape = None
    if cell.paragraphs:
        para_shape = _lookup_para_shape(cell.paragraphs[0], para_shapes if isinstance(para_shapes, list) else [])
    left_pad = hwp_to_px(cell.margin_left_hwp or 0.0, dpi=192.0) or (12.0 * scale)
    right_pad = hwp_to_px(cell.margin_right_hwp or 0.0, dpi=192.0) or (12.0 * scale)
    top_pad = hwp_to_px(cell.margin_top_hwp or 0.0, dpi=192.0) or (8.0 * scale)
    bottom_pad = hwp_to_px(cell.margin_bottom_hwp or 0.0, dpi=192.0) or (8.0 * scale)
    painter.setFont(base_font)
    painter.setPen(QColor(Qt.black))
    painter.drawText(
        rect.adjusted(left_pad, top_pad, -right_pad, -bottom_pad),
        int(_paragraph_alignment_flags(para_shape) | Qt.AlignVCenter | Qt.TextWordWrap),
        cell_text,
    )


def _border_width_px(width_code: int, dpi: float) -> float:
    mm = HWP_BORDER_WIDTH_MM.get(width_code, 0.1)
    return max(1.0, mm / 25.4 * dpi)


def _border_pen_style(line_type: int) -> Qt.PenStyle:
    if line_type in {1, 5}:
        return Qt.DashLine
    if line_type in {2, 6}:
        return Qt.DotLine
    if line_type in {3, 4}:
        return Qt.DashDotLine
    return Qt.SolidLine


def _table_column_count(table: HwpTableModel) -> int:
    max_cols = 1
    for row in table.rows:
        if all(cell.col_index is not None for cell in row.cells):
            count = 0
            for cell in row.cells:
                assert cell.col_index is not None
                count = max(count, cell.col_index + max(1, cell.col_span))
            max_cols = max(max_cols, count)
        else:
            max_cols = max(max_cols, len(row.cells))
    return max_cols


def _column_positions(x: float, col_widths: list[float]) -> list[float]:
    positions = [x]
    current = x
    for width in col_widths:
        current += width
        positions.append(current)
    return positions


def _pick_font_family(preferred: list[str] | None = None) -> str:
    if QGuiApplication.instance() is None:
        if preferred:
            return preferred[0]
        return FONT_FALLBACKS[0]
    db = QFontDatabase()
    families = set(db.families())
    for family in preferred or []:
        if family in families:
            return family
    for family in FONT_FALLBACKS:
        if family in families:
            return family
    return QFont().family()


def _make_font(family: str, point_size: float, *, bold: bool) -> QFont:
    font = QFont(family)
    font.setPointSizeF(max(8.0, point_size))
    font.setBold(bold)
    return font


def _wrapped_text_height(font: QFont, width: float, text: str) -> float:
    metrics = QFontMetricsF(font)
    rect = metrics.boundingRect(QRectF(0.0, 0.0, max(1.0, width), 10000.0), int(Qt.TextWordWrap), text)
    return rect.height()


def _table_row_height(
    row,
    col_count: int,
    col_widths: list[float],
    font: QFont,
    scale: float,
) -> float:
    max_height = 42.0 * scale
    dense_col = 0
    for cell in row.cells:
        col_index = cell.col_index if cell.col_index is not None else dense_col
        span = max(1, cell.col_span)
        dense_col = col_index + span
        if col_index >= col_count:
            continue
        cell_text = ' '.join(
            ''.join(run.text for run in para.runs).strip()
            for para in cell.paragraphs
        ).strip()
        if not cell_text:
            continue
        cell_width = sum(col_widths[col_index:min(col_count, col_index + span)])
        text_height = _wrapped_text_height(font, cell_width - (24.0 * scale), cell_text)
        max_height = max(max_height, text_height + (18.0 * scale))
        if cell.height_hwp and cell.height_hwp > 0 and max(1, cell.row_span) == 1:
            max_height = max(max_height, hwp_to_px(cell.height_hwp, dpi=192.0))
    return max_height
