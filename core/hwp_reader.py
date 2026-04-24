from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from core.hwp_probe import HwpProbeError, probe_hwp_source
from core.hwp_types import (
    HwpCharShapeSpan,
    HwpDocumentModel,
    HwpFlowBlock,
    HwpFormat,
    HwpLineSegmentModel,
    HwpMargins,
    HwpPageModel,
    HwpPageSize,
    HwpParagraphModel,
    HwpParagraphRun,
    HwpTableCellModel,
    HwpTableModel,
    HwpTableRowModel,
)


DEFAULT_A4_WIDTH_HWP = 595.0 / 72.0 * 7200.0
DEFAULT_A4_HEIGHT_HWP = 842.0 / 72.0 * 7200.0
ANGLE_TOKEN_RE = re.compile(r'<([^<>]*)>')
HWPTAG_PARA_HEADER = 16 + 50
HWPTAG_PARA_TEXT = 16 + 51
HWPTAG_PARA_CHAR_SHAPE = 16 + 52
HWPTAG_PARA_LINE_SEG = 16 + 53
HWPTAG_CTRL_HEADER = 16 + 55
HWPTAG_LIST_HEADER = 16 + 56
HWPTAG_PAGE_DEF = 16 + 57
HWPTAG_TABLE = 16 + 61


class HwpReadError(RuntimeError):
    pass


def load_hwp_document(path: str | Path) -> HwpDocumentModel:
    try:
        source = probe_hwp_source(path)
    except HwpProbeError as exc:
        raise HwpReadError(str(exc)) from exc
    if source.fmt == HwpFormat.HWPX:
        return _load_hwpx_document(source.path, source)
    return _load_hwp_document(source.path, source)


def _load_hwpx_document(path: Path, source) -> HwpDocumentModel:
    page_size = HwpPageSize(DEFAULT_A4_WIDTH_HWP, DEFAULT_A4_HEIGHT_HWP)
    model = HwpDocumentModel(source=source)
    with zipfile.ZipFile(path, 'r') as zf:
        section_entries = [
            name for name in zf.namelist()
            if 'section' in name.lower() and name.lower().endswith('.xml')
        ]
        if not section_entries:
            section_entries = [name for name in zf.namelist() if name.lower().endswith('.xml')]
        for section_name in sorted(section_entries):
            pages = _extract_hwpx_pages(zf.read(section_name), page_size, section_name=section_name)
            if not pages:
                continue
            model.pages.extend(pages)
    if not model.pages:
        empty_para = HwpParagraphModel(runs=[HwpParagraphRun(text='(빈 HWPX 문서)')])
        model.pages.append(
            HwpPageModel(
                size=page_size,
                margins=HwpMargins(),
                paragraphs=[empty_para],
                flow_blocks=[HwpFlowBlock('paragraph', empty_para)],
            )
        )
    return model


def _load_hwp_document(path: Path, source) -> HwpDocumentModel:
    try:
        import olefile
    except Exception as exc:
        raise HwpReadError(f'olefile import 실패: {exc}') from exc
    page_size = HwpPageSize(DEFAULT_A4_WIDTH_HWP, DEFAULT_A4_HEIGHT_HWP)
    model = HwpDocumentModel(source=source)
    with olefile.OleFileIO(str(path)) as ole:
        preview_text = _read_hwp_preview_text(ole)
        preview_pages = _structured_text_to_pages(preview_text, page_size) if preview_text else []
        extracted = _extract_hwp_body_models(ole, page_size)
        if extracted:
            model.pages.extend(extracted)
        elif preview_pages:
            model.pages.extend(preview_pages)
    if not model.pages:
        model.pages.append(
            HwpPageModel(
                size=page_size,
                margins=HwpMargins(),
                paragraphs=[HwpParagraphModel(runs=[HwpParagraphRun(text='(HWP 본문을 아직 해석하지 못했습니다)')])],
            )
        )
    return model


def _extract_xml_text_blocks(raw_xml: bytes) -> list[str]:
    try:
        root = ET.fromstring(raw_xml)
    except Exception:
        return []
    blocks: list[str] = []
    current: list[str] = []
    for elem in root.iter():
        text = (elem.text or '').strip()
        if text:
            current.append(text)
        local = elem.tag.rsplit('}', 1)[-1].lower()
        if local in {'p', 'paragraph', 'hp:p'} and current:
            joined = ' '.join(current).strip()
            if joined:
                blocks.append(joined)
            current = []
    if current:
        joined = ' '.join(current).strip()
        if joined:
            blocks.append(joined)
    deduped: list[str] = []
    seen: set[str] = set()
    for block in blocks:
        normalized = re.sub(r'\s+', ' ', block).strip()
        if normalized and normalized not in seen:
            deduped.append(normalized)
            seen.add(normalized)
    return deduped


def _read_hwp_preview_text(ole) -> str:
    try:
        raw = ole.openstream('PrvText').read()
    except Exception:
        return ''
    for encoding in ('utf-16-le', 'cp949', 'utf-8'):
        try:
            text = raw.decode(encoding, errors='ignore').replace('\x00', '').strip()
        except Exception:
            continue
        if text:
            return text
    return ''


def _extract_hwp_body_texts(ole) -> list[str]:
    texts: list[str] = []
    for stream_name in ole.listdir():
        joined = '/'.join(stream_name)
        if not joined.startswith('BodyText/Section'):
            continue
        try:
            raw = ole.openstream(stream_name).read()
        except Exception:
            continue
        texts.extend(_extract_text_candidates_from_bytes(raw))
    return texts


def _extract_hwp_body_models(ole, page_size: HwpPageSize) -> list[HwpPageModel]:
    section_models: list[HwpPageModel] = []
    for stream_name in ole.listdir():
        joined = '/'.join(stream_name)
        if not joined.startswith('BodyText/Section'):
            continue
        try:
            raw = ole.openstream(stream_name).read()
            data = _decompress_hwp_section(raw)
        except Exception:
            continue
        pages = _parse_hwp_section_records(data, page_size)
        for page in pages:
            page.raw_attrs = {'source_stream': joined}
            suffix = joined.split('Section', 1)[-1]
            try:
                page.source_index = int(suffix)
            except Exception:
                page.source_index = None
            page.source_name = joined
        section_models.extend(pages)
    return section_models


def _decompress_hwp_section(raw: bytes) -> bytes:
    try:
        import zlib
        return zlib.decompress(raw, -15)
    except Exception:
        return raw


def _parse_hwp_section_records(data: bytes, page_size: HwpPageSize) -> list[HwpPageModel]:
    pages: list[HwpPageModel] = []
    current_page_size = page_size
    current_margins = HwpMargins()
    current_page = HwpPageModel(size=current_page_size, margins=current_margins)
    pos = 0
    pending_para_header: dict[str, bool] | None = None
    pending_ctrl_id: str | None = None
    table_state: dict[str, object] | None = None
    last_paragraph: HwpParagraphModel | None = None
    table_child_tags = {
        HWPTAG_LIST_HEADER,
        HWPTAG_PARA_HEADER,
        HWPTAG_PARA_TEXT,
        HWPTAG_PARA_CHAR_SHAPE,
        HWPTAG_PARA_LINE_SEG,
    }

    def flush_table() -> None:
        nonlocal table_state
        if not table_state:
            return
        row_counts = list(table_state.get('row_counts') or [])
        cells = list(table_state.get('cells') or [])
        if not cells:
            table_state = None
            return
        table_rows = _build_table_rows(cells, row_counts)
        table = HwpTableModel(
            rows=table_rows,
            control_id=str(table_state.get('control_id') or '') or None,
            row_count=int(table_state.get('row_count') or 0) or None,
            col_count=int(table_state.get('col_count') or 0) or None,
            border_fill_id=int(table_state.get('border_fill_id') or 0) or None,
            raw_header=bytes(table_state.get('raw_header') or b'') or None,
            raw_attrs={'ctrl_raw_header': bytes(table_state['ctrl_raw_header']).hex()} if table_state.get('ctrl_raw_header') else None,
        )
        current_page.tables.append(table)
        current_page.flow_blocks.append(HwpFlowBlock('table', table))
        table_state = None

    def flush_page(force: bool = False) -> None:
        nonlocal current_page
        flush_table()
        if force or current_page.flow_blocks:
            pages.append(current_page)
            current_page = HwpPageModel(size=current_page_size, margins=current_margins)

    while pos + 4 <= len(data):
        header = int.from_bytes(data[pos:pos + 4], 'little')
        pos += 4
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            if pos + 4 > len(data):
                break
            size = int.from_bytes(data[pos:pos + 4], 'little')
            pos += 4
        body = data[pos:pos + size]
        pos += size

        if table_state:
            table_level = int(table_state.get('table_level', 0))
            if level <= table_level and tag_id not in table_child_tags:
                flush_table()

        if tag_id == HWPTAG_PARA_HEADER:
            pending_para_header = _parse_hwp_para_header(body)
            pending_para_header['raw_header'] = body
            continue
        if tag_id == HWPTAG_CTRL_HEADER:
            pending_ctrl_id = _parse_hwp_ctrl_header(body)
            if pending_ctrl_id and table_state is not None:
                table_state['ctrl_raw_header'] = body
            continue
        if tag_id == HWPTAG_LIST_HEADER:
            if table_state and level >= int(table_state.get('table_level', 0)):
                cells = table_state.setdefault('cells', [])
                assert isinstance(cells, list)
                cells.append(_parse_hwp_table_cell_header(body))
            continue
        if tag_id == HWPTAG_PAGE_DEF:
            page_def = _parse_hwp_page_def(body)
            if page_def:
                if current_page.flow_blocks:
                    flush_page()
                current_page_size = page_def['size']
                current_margins = page_def['margins']
                current_page.size = current_page_size
                current_page.margins = current_margins
                current_page.raw_page_def = body
            continue
        if tag_id == HWPTAG_TABLE:
            flush_table()
            if pending_ctrl_id == 'tbl ':
                row_counts = _parse_hwp_table_rows(body) or []
                table_header = _parse_hwp_table_header(body)
                table_state = {
                    'table_level': level,
                    'row_counts': row_counts,
                    'cells': [],
                    'control_id': pending_ctrl_id,
                    'raw_header': body,
                    'row_count': table_header.get('row_count') or 0,
                    'col_count': table_header.get('col_count') or 0,
                    'border_fill_id': table_header.get('border_fill_id') or 0,
                }
            pending_ctrl_id = None
            continue
        if tag_id == HWPTAG_PARA_CHAR_SHAPE:
            if last_paragraph is not None:
                last_paragraph.char_shape_spans = _parse_hwp_para_char_shape(body)
                if last_paragraph.runs and last_paragraph.char_shape_spans:
                    first_span = last_paragraph.char_shape_spans[0]
                    if first_span.char_index == 0:
                        last_paragraph.runs[0].char_shape_id = first_span.char_shape_id
            continue
        if tag_id == HWPTAG_PARA_LINE_SEG:
            if last_paragraph is not None:
                last_paragraph.line_segments = _parse_hwp_para_line_segments(body)
            continue
        if tag_id != HWPTAG_PARA_TEXT:
            continue

        text = _decode_hwp_para_text(body)
        normalized = _normalize_text(text)
        if not normalized:
            pending_para_header = None
            last_paragraph = None
            continue

        if pending_para_header and current_page.flow_blocks:
            if pending_para_header.get('section_break') or pending_para_header.get('page_break'):
                flush_page()

        para = HwpParagraphModel(
            runs=[HwpParagraphRun(text=normalized)],
        )
        if pending_para_header:
            para.text_char_count = pending_para_header.get('text_char_count')
            para.control_mask = pending_para_header.get('control_mask')
            para.para_shape_id = pending_para_header.get('para_shape_id')
            para.section_break = bool(pending_para_header.get('section_break'))
            para.page_break = bool(pending_para_header.get('page_break'))
            para.column_break = bool(pending_para_header.get('column_break'))
            para.columns_break = bool(pending_para_header.get('columns_break'))
            raw_header = pending_para_header.get('raw_header')
            if isinstance(raw_header, (bytes, bytearray)):
                para.raw_header = bytes(raw_header)

        if table_state and level > int(table_state.get('table_level', 0)):
            cells = table_state.setdefault('cells', [])
            assert isinstance(cells, list)
            if not cells:
                cells.append(HwpTableCellModel())
            cells[-1].paragraphs.append(para)
        else:
            flush_table()
            current_page.paragraphs.append(para)
            current_page.flow_blocks.append(HwpFlowBlock('paragraph', para))

        last_paragraph = para
        if pending_para_header and current_page.flow_blocks and (para.section_break or para.page_break):
            flush_page()
        pending_para_header = None

    flush_page(force=True)
    return pages


def _parse_hwp_para_header(body: bytes) -> dict[str, object]:
    if len(body) < 16:
        return {}
    text_char_count = int.from_bytes(body[0:4], 'little')
    control_mask = int.from_bytes(body[4:8], 'little')
    para_shape_id = int.from_bytes(body[8:10], 'little')
    break_options = body[11]
    return {
        'text_char_count': text_char_count,
        'control_mask': control_mask,
        'para_shape_id': para_shape_id,
        'section_break': bool(break_options & (1 << 0)),
        'columns_break': bool(break_options & (1 << 1)),
        'page_break': bool(break_options & (1 << 2)),
        'column_break': bool(break_options & (1 << 3)),
    }


def _parse_hwp_ctrl_header(body: bytes) -> str | None:
    if len(body) < 4:
        return None
    ctrl_id = int.from_bytes(body[:4], 'little')
    chars = ctrl_id.to_bytes(4, 'big')
    try:
        return chars.decode('ascii', errors='ignore')
    except Exception:
        return None


def _parse_hwp_para_char_shape(body: bytes) -> list[HwpCharShapeSpan]:
    spans: list[HwpCharShapeSpan] = []
    entry_size = 8
    for pos in range(0, len(body) - entry_size + 1, entry_size):
        spans.append(
            HwpCharShapeSpan(
                char_index=int.from_bytes(body[pos:pos + 4], 'little'),
                char_shape_id=int.from_bytes(body[pos + 4:pos + 8], 'little'),
            )
        )
    return spans


def _parse_hwp_para_line_segments(body: bytes) -> list[HwpLineSegmentModel]:
    segments: list[HwpLineSegmentModel] = []
    entry_size = 36
    for pos in range(0, len(body) - entry_size + 1, entry_size):
        segments.append(
            HwpLineSegmentModel(
                text_start=int.from_bytes(body[pos:pos + 4], 'little'),
                vertical_pos_hwp=int.from_bytes(body[pos + 4:pos + 8], 'little', signed=True),
                line_height_hwp=int.from_bytes(body[pos + 8:pos + 12], 'little', signed=True),
                text_height_hwp=int.from_bytes(body[pos + 12:pos + 16], 'little', signed=True),
                baseline_hwp=int.from_bytes(body[pos + 16:pos + 20], 'little', signed=True),
                line_spacing_hwp=int.from_bytes(body[pos + 20:pos + 24], 'little', signed=True),
                column_start_hwp=int.from_bytes(body[pos + 24:pos + 28], 'little', signed=True),
                segment_width_hwp=int.from_bytes(body[pos + 28:pos + 32], 'little', signed=True),
                flags=int.from_bytes(body[pos + 32:pos + 36], 'little'),
            )
        )
    return segments


def _parse_hwp_table_cell_header(body: bytes) -> HwpTableCellModel:
    cell = HwpTableCellModel()
    if len(body) < 8:
        return cell
    pos = 8
    if len(body) >= pos + 26:
        cell.col_index = int.from_bytes(body[pos:pos + 2], 'little')
        cell.row_index = int.from_bytes(body[pos + 2:pos + 4], 'little')
        cell.col_span = max(1, int.from_bytes(body[pos + 4:pos + 6], 'little'))
        cell.row_span = max(1, int.from_bytes(body[pos + 6:pos + 8], 'little'))
        cell.width_hwp = float(int.from_bytes(body[pos + 8:pos + 12], 'little'))
        cell.height_hwp = float(int.from_bytes(body[pos + 12:pos + 16], 'little'))
        cell.margin_left_hwp = float(int.from_bytes(body[pos + 16:pos + 18], 'little'))
        cell.margin_right_hwp = float(int.from_bytes(body[pos + 18:pos + 20], 'little'))
        cell.margin_top_hwp = float(int.from_bytes(body[pos + 20:pos + 22], 'little'))
        cell.margin_bottom_hwp = float(int.from_bytes(body[pos + 22:pos + 24], 'little'))
        cell.border_fill_id = int.from_bytes(body[pos + 24:pos + 26], 'little')
    return cell


def _build_table_rows(cells: list[HwpTableCellModel], row_counts: list[int]) -> list[HwpTableRowModel]:
    if not cells:
        return []
    if all(cell.row_index is not None and cell.col_index is not None for cell in cells):
        grouped: dict[int, list[HwpTableCellModel]] = {}
        for cell in cells:
            assert cell.row_index is not None
            grouped.setdefault(cell.row_index, []).append(cell)
        table_rows: list[HwpTableRowModel] = []
        for row_index in sorted(grouped):
            row_cells = sorted(
                grouped[row_index],
                key=lambda cell: (cell.col_index if cell.col_index is not None else 0),
            )
            table_rows.append(HwpTableRowModel(cells=row_cells))
        if table_rows:
            return table_rows
    table_rows = []
    cursor = 0
    if row_counts:
        for row_cell_count in row_counts:
            row_cells = cells[cursor:cursor + row_cell_count]
            if row_cells:
                table_rows.append(HwpTableRowModel(cells=row_cells))
            cursor += row_cell_count
    if cursor < len(cells):
        table_rows.append(HwpTableRowModel(cells=cells[cursor:]))
    if not table_rows:
        table_rows = [HwpTableRowModel(cells=cells)]
    return table_rows


def _parse_hwp_page_def(body: bytes) -> dict[str, object] | None:
    if len(body) < 36:
        return None
    values = [int.from_bytes(body[i:i + 4], 'little') for i in range(0, 36, 4)]
    size = HwpPageSize(float(values[0]), float(values[1]))
    margins = HwpMargins(
        left_hwp=float(values[2]),
        right_hwp=float(values[3]),
        top_hwp=float(values[4]),
        bottom_hwp=float(values[5]),
        header_hwp=float(values[6]),
        footer_hwp=float(values[7]),
        gutter_hwp=float(values[8]),
    )
    return {'size': size, 'margins': margins}


def _parse_hwp_table_rows(body: bytes) -> list[int] | None:
    if len(body) < 22:
        return None
    rows = int.from_bytes(body[4:6], 'little')
    if rows <= 0:
        return None
    row_counts_offset = 22
    max_needed = row_counts_offset + rows * 2
    if len(body) < max_needed:
        return None
    row_counts = [
        int.from_bytes(body[row_counts_offset + i * 2:row_counts_offset + (i + 1) * 2], 'little')
        for i in range(rows)
    ]
    return [count for count in row_counts if count > 0] or None


def _parse_hwp_table_header(body: bytes) -> dict[str, int]:
    if len(body) < 22:
        return {}
    return {
        'row_count': int.from_bytes(body[4:6], 'little'),
        'col_count': int.from_bytes(body[6:8], 'little'),
        'border_fill_id': int.from_bytes(body[20:22], 'little'),
    }


def _decode_hwp_para_text(body: bytes) -> str:
    cleaned = bytearray()
    pos = 0
    char_ctrl = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}
    while pos + 2 <= len(body):
        wchar = int.from_bytes(body[pos:pos + 2], 'little')
        if wchar < 32:
            if wchar in char_ctrl:
                pos += 2
            else:
                pos += 16
            continue
        cleaned.extend(body[pos:pos + 2])
        pos += 2
    return cleaned.decode('utf-16-le', errors='ignore').replace('\x00', ' ')


def _extract_text_candidates_from_bytes(data: bytes) -> list[str]:
    results: list[str] = []
    for encoding in ('utf-16-le', 'utf-8', 'cp949', 'utf-16-be', 'latin1'):
        try:
            decoded = data.decode(encoding, errors='ignore')
        except Exception:
            continue
        cleaned = decoded.replace('\x00', '')
        cleaned = re.sub(r'[\t\r\f\v]+', ' ', cleaned)
        cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
        cleaned = re.sub(r'[^\w가-힣\s\-_,.:;()\[\]/%+*&@!?]', ' ', cleaned)
        cleaned = re.sub(r' {2,}', ' ', cleaned)
        lines = [line.strip() for line in cleaned.splitlines()]
        useful = [line for line in lines if len(re.sub(r'\W+', '', line)) >= 3]
        if useful:
            results.extend(useful[:200])
    deduped: list[str] = []
    seen: set[str] = set()
    for item in results:
        if item and item not in seen:
            deduped.append(item)
            seen.add(item)
    return deduped


def _text_to_pages(text: str, page_size: HwpPageSize, lines_per_page: int = 42) -> list[HwpPageModel]:
    lines = [line.strip() for line in text.splitlines()]
    filtered = [line for line in lines if line]
    if not filtered:
        return []
    pages: list[HwpPageModel] = []
    chunk: list[str] = []
    for line in filtered:
        chunk.append(line)
        if len(chunk) >= lines_per_page:
            pages.append(_make_page_from_lines(chunk, page_size))
            chunk = []
    if chunk:
        pages.append(_make_page_from_lines(chunk, page_size))
    return pages


def _structured_text_to_pages(text: str, page_size: HwpPageSize, lines_per_page: int = 36) -> list[HwpPageModel]:
    raw_lines = [line.strip() for line in text.replace('\r', '\n').splitlines()]
    lines = [line for line in raw_lines if line]
    if not lines:
        return []
    pages: list[HwpPageModel] = []
    current_page = HwpPageModel(size=page_size, margins=HwpMargins())
    line_budget = 0

    def flush_page(force: bool = False) -> None:
        nonlocal current_page, line_budget
        if force or current_page.flow_blocks:
            pages.append(current_page)
            current_page = HwpPageModel(size=page_size, margins=HwpMargins())
            line_budget = 0

    def append_paragraph(text_value: str, style_name: str | None = None, budget_cost: int = 1) -> None:
        nonlocal line_budget
        normalized = _normalize_text(text_value)
        if not normalized:
            return
        para = HwpParagraphModel(runs=[HwpParagraphRun(text=normalized)], style_name=style_name)
        current_page.paragraphs.append(para)
        current_page.flow_blocks.append(HwpFlowBlock('paragraph', para))
        line_budget += budget_cost

    def append_table(tokens: list[str]) -> None:
        nonlocal line_budget
        row = HwpTableRowModel(
            cells=[
                HwpTableCellModel(paragraphs=[HwpParagraphModel(runs=[HwpParagraphRun(text=_normalize_text(token))])])
                for token in tokens
                if _normalize_text(token)
            ]
        )
        if not row.cells:
            return
        table = HwpTableModel(rows=[row])
        current_page.tables.append(table)
        current_page.flow_blocks.append(HwpFlowBlock('table', table))
        line_budget += 2

    for line in lines:
        tokens = [token.strip() for token in ANGLE_TOKEN_RE.findall(line) if token.strip()]
        if len(tokens) >= 3:
            append_table(tokens)
        elif len(tokens) == 2:
            left, right = tokens
            if len(_normalize_text(left)) <= 16 and right.startswith(('·', '•', '*')):
                append_table([left, right])
            else:
                append_paragraph(f'{left} {right}', style_name=None, budget_cost=max(1, len(left + right) // 45))
        elif len(tokens) == 1:
            token = tokens[0]
            append_paragraph(token, style_name=None, budget_cost=max(1, len(token) // 45))
        else:
            append_paragraph(line, style_name=None, budget_cost=max(1, len(line) // 45))
        if line_budget >= lines_per_page:
            flush_page()

    if current_page.flow_blocks:
        pages.append(current_page)
    return pages or _text_to_pages(text, page_size)


def _make_page_from_lines(lines: list[str], page_size: HwpPageSize) -> HwpPageModel:
    paragraphs = [HwpParagraphModel(runs=[HwpParagraphRun(text=line)]) for line in lines]
    return HwpPageModel(
        size=page_size,
        margins=HwpMargins(),
        paragraphs=paragraphs,
        flow_blocks=[HwpFlowBlock('paragraph', para) for para in paragraphs],
    )


def _extract_hwpx_pages(raw_xml: bytes, page_size: HwpPageSize, section_name: str | None = None) -> list[HwpPageModel]:
    try:
        root = ET.fromstring(raw_xml)
    except Exception:
        return []
    pages: list[HwpPageModel] = []
    section_attrs = _extract_hwpx_section_attrs(root)
    if section_name:
        section_attrs['source_xml_name'] = section_name
    section_page_size, section_margins = _extract_hwpx_page_metrics(section_attrs, page_size)
    current_page = HwpPageModel(
        size=section_page_size,
        margins=section_margins,
        raw_attrs=section_attrs,
        source_name=section_name,
    )

    def append_paragraph(text: str, raw_attrs: dict[str, str] | None = None) -> None:
        normalized = _normalize_text(text)
        if not normalized:
            return
        para = HwpParagraphModel(
            runs=[HwpParagraphRun(text=normalized)],
            raw_attrs=dict(raw_attrs or {}) or None,
        )
        current_page.paragraphs.append(para)
        current_page.flow_blocks.append(HwpFlowBlock('paragraph', para))

    def append_table(table: HwpTableModel) -> None:
        if not table.rows:
            return
        current_page.tables.append(table)
        current_page.flow_blocks.append(HwpFlowBlock('table', table))

    def flush_page(force: bool = False) -> None:
        nonlocal current_page
        if force or current_page.flow_blocks:
            pages.append(current_page)
            current_page = HwpPageModel(
                size=section_page_size,
                margins=section_margins,
                raw_attrs=section_attrs,
                source_name=section_name,
            )

    for elem in root.iter():
        local = _local_name(elem.tag)
        if local in {'tbl', 'table'}:
            table = _parse_table_element(elem)
            if table.rows:
                append_table(table)
            continue
        attr_blob = ' '.join(f'{key}={value}' for key, value in elem.attrib.items()).lower()
        if local in {'br', 'break'} and ('page' in attr_blob):
            flush_page()
            continue
        if local in {'pagebreak', 'pagebreakline', 'lastrenderedpagebreak', 'colpagebreak', 'sectionbreak'}:
            flush_page()
            continue
        if 'pagebreak' in local or 'renderedpagebreak' in local:
            flush_page()
            continue
        if local in {'p', 'paragraph'}:
            append_paragraph(_collect_text(elem), raw_attrs={key: value for key, value in elem.attrib.items()})
    if current_page.flow_blocks:
        pages.append(current_page)
    return pages


def _extract_hwpx_section_attrs(root: ET.Element) -> dict[str, str]:
    attrs: dict[str, str] = {'source_xml': 'section'}
    for elem in root.iter():
        local = _local_name(elem.tag)
        if local not in {'secpr', 'pagepr', 'pagemargin', 'papermargin', 'pagesize', 'papersize'}:
            continue
        prefix = local
        for key, value in elem.attrib.items():
            attrs[f'{prefix}.{key}'] = value
    return attrs


def _extract_hwpx_page_metrics(section_attrs: dict[str, str], page_size: HwpPageSize) -> tuple[HwpPageSize, HwpMargins]:
    size = HwpPageSize(page_size.width_hwp, page_size.height_hwp)
    margins = HwpMargins()
    width_keys = ('pagesize.w', 'pagesize.width', 'pagepr.w', 'pagepr.width', 'papermargin.w', 'papersize.w')
    height_keys = ('pagesize.h', 'pagesize.height', 'pagepr.h', 'pagepr.height', 'papermargin.h', 'papersize.h')
    for key in width_keys:
        if key in section_attrs:
            try:
                size.width_hwp = float(section_attrs[key])
                break
            except Exception:
                pass
    for key in height_keys:
        if key in section_attrs:
            try:
                size.height_hwp = float(section_attrs[key])
                break
            except Exception:
                pass
    margin_keys = {
        'left_hwp': ('pagemargin.left', 'pagepr.left', 'secpr.left', 'margin.left'),
        'right_hwp': ('pagemargin.right', 'pagepr.right', 'secpr.right', 'margin.right'),
        'top_hwp': ('pagemargin.top', 'pagepr.top', 'secpr.top', 'margin.top'),
        'bottom_hwp': ('pagemargin.bottom', 'pagepr.bottom', 'secpr.bottom', 'margin.bottom'),
        'header_hwp': ('pagemargin.header', 'pagepr.header', 'secpr.header'),
        'footer_hwp': ('pagemargin.footer', 'pagepr.footer', 'secpr.footer'),
    }
    for field_name, keys in margin_keys.items():
        for key in keys:
            if key in section_attrs:
                try:
                    setattr(margins, field_name, float(section_attrs[key]))
                    break
                except Exception:
                    continue
    return size, margins


def _parse_table_element(table_elem: ET.Element) -> HwpTableModel:
    table = HwpTableModel(raw_attrs={key: value for key, value in table_elem.attrib.items()})
    for row_elem in table_elem.iter():
        if _local_name(row_elem.tag) not in {'tr', 'row'}:
            continue
        row = HwpTableRowModel()
        for cell_elem in row_elem:
            if _local_name(cell_elem.tag) not in {'tc', 'cell'}:
                continue
            texts = _extract_paragraph_texts(cell_elem)
            paragraphs = [HwpParagraphModel(runs=[HwpParagraphRun(text=text)]) for text in texts if _normalize_text(text)]
            row.cells.append(
                HwpTableCellModel(
                    paragraphs=paragraphs,
                    raw_attrs={key: value for key, value in cell_elem.attrib.items()},
                )
            )
        if row.cells:
            table.rows.append(row)
    return table


def _extract_paragraph_texts(parent: ET.Element) -> list[str]:
    texts: list[str] = []
    for elem in parent.iter():
        if _local_name(elem.tag) in {'p', 'paragraph'}:
            normalized = _normalize_text(_collect_text(elem))
            if normalized:
                texts.append(normalized)
    if texts:
        return texts
    fallback = _normalize_text(_collect_text(parent))
    return [fallback] if fallback else []


def _collect_text(elem: ET.Element) -> str:
    parts = [text.strip() for text in elem.itertext() if text and text.strip()]
    return ' '.join(parts)


def _normalize_text(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip()


def _local_name(tag: str) -> str:
    return tag.rsplit('}', 1)[-1].lower()
