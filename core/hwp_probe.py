from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from core.hwp_types import (
    HwpBorderFillModel,
    HwpCharShapeModel,
    HwpContainer,
    HwpFormat,
    HwpParaShapeModel,
    HwpSectionRef,
    HwpSourceInfo,
)

HWPTAG_DOCUMENT_PROPERTIES = 16
HWPTAG_ID_MAPPINGS = 16 + 1
HWPTAG_FACE_NAME = 16 + 3
HWPTAG_BORDER_FILL = 16 + 4
HWPTAG_CHAR_SHAPE = 16 + 5
HWPTAG_PARA_SHAPE = 16 + 9

HWP_ID_MAPPING_NAMES = (
    'binary_data',
    'hangul_font',
    'english_font',
    'hanja_font',
    'japanese_font',
    'other_font',
    'symbol_font',
    'user_font',
    'border_fill',
    'char_shape',
    'tab_def',
    'numbering',
    'bullet',
    'para_shape',
    'style',
    'memo_shape',
    'track_change',
    'track_change_author',
)


class HwpProbeError(RuntimeError):
    pass


def probe_hwp_source(path: str | Path) -> HwpSourceInfo:
    src = Path(path).expanduser().resolve()
    ext = src.suffix.lower()
    if ext == '.hwp':
        return _probe_hwp_ole(src)
    if ext == '.hwpx':
        return _probe_hwpx_zip(src)
    raise HwpProbeError(f'지원하지 않는 HWP 계열 형식입니다: {ext}')


def _probe_hwp_ole(src: Path) -> HwpSourceInfo:
    try:
        import olefile
    except Exception as exc:
        raise HwpProbeError(f'olefile import 실패: {exc}') from exc
    if not olefile.isOleFile(str(src)):
        raise HwpProbeError('유효한 HWP OLE 파일이 아닙니다.')
    info = HwpSourceInfo(
        path=src,
        fmt=HwpFormat.HWP,
        container=HwpContainer.OLE,
        file_size=src.stat().st_size,
    )
    with olefile.OleFileIO(str(src)) as ole:
        info.stream_names = ['/'.join(parts) for parts in ole.listdir()]
        info.section_refs = _extract_hwp_sections(info.stream_names)
        info.fonts = _extract_hwp_font_hints(info.stream_names)
        doc_fonts, doc_counts, doc_attrs, border_fills, char_shapes, para_shapes = _extract_hwp_doc_info(ole)
        if doc_fonts:
            info.fonts = doc_fonts
        info.ref_list_counts.update(doc_counts)
        info.raw_attrs.update(doc_attrs)
        info.border_fills = border_fills
        info.char_shapes = char_shapes
        info.para_shapes = para_shapes
        header = _read_hwp_file_header(ole)
        if header:
            info.notes.append(header)
        if any(name.endswith('PrvText') for name in info.stream_names):
            info.notes.append('preview text stream present')
        if any(name.endswith('DocInfo') for name in info.stream_names):
            info.notes.append('docinfo stream present')
    if info.section_refs:
        info.page_count_hint = max(1, len(info.section_refs))
    return info


def _probe_hwpx_zip(src: Path) -> HwpSourceInfo:
    if not zipfile.is_zipfile(src):
        raise HwpProbeError('유효한 HWPX zip 파일이 아닙니다.')
    info = HwpSourceInfo(
        path=src,
        fmt=HwpFormat.HWPX,
        container=HwpContainer.ZIP_XML,
        file_size=src.stat().st_size,
    )
    with zipfile.ZipFile(src, 'r') as zf:
        info.entry_names = sorted(zf.namelist())
        info.section_refs = _extract_hwpx_sections(info.entry_names)
        info.fonts = _extract_hwpx_font_hints(zf)
        info.page_count_hint = _extract_hwpx_page_hint(zf)
        info.begin_nums = _extract_hwpx_begin_nums(zf)
        info.ref_list_counts = _extract_hwpx_ref_list_counts(zf)
        info.raw_attrs = _extract_hwpx_head_attrs(zf)
        if 'Contents/content.hpf' in info.entry_names:
            info.notes.append('opf manifest present')
        if any('section' in name.lower() for name in info.entry_names):
            info.notes.append('section xml entries present')
    return info


def _extract_hwp_sections(stream_names: list[str]) -> list[HwpSectionRef]:
    section_names = [name for name in stream_names if name.startswith('BodyText/Section')]
    refs: list[HwpSectionRef] = []
    for stream_name in sorted(section_names):
        suffix = stream_name.split('Section', 1)[-1]
        try:
            index = int(suffix)
        except ValueError:
            continue
        refs.append(HwpSectionRef(index=index, name=f'Section{index}', source_path=stream_name))
    return refs


def _extract_hwp_font_hints(stream_names: list[str]) -> list[str]:
    hints: list[str] = []
    for name in stream_names:
        lower = name.lower()
        if 'font' in lower:
            hints.append(name)
    return hints[:32]


def _read_hwp_file_header(ole: olefile.OleFileIO) -> str:
    try:
        raw = ole.openstream('FileHeader').read(64)
    except Exception:
        return ''
    try:
        text = raw.decode('ascii', errors='ignore').strip('\x00 ').strip()
    except Exception:
        return ''
    return text


def _extract_hwp_doc_info(
    ole: olefile.OleFileIO,
) -> tuple[list[str], dict[str, int], dict[str, str], list[HwpBorderFillModel], list[HwpCharShapeModel], list[HwpParaShapeModel]]:
    try:
        raw = ole.openstream('DocInfo').read()
    except Exception:
        return [], {}, {}, [], [], []
    data = _decompress_hwp_stream(raw)
    fonts: list[str] = []
    counts: dict[str, int] = {}
    attrs: dict[str, str] = {}
    border_fills: list[HwpBorderFillModel] = []
    char_shapes: list[HwpCharShapeModel] = []
    para_shapes: list[HwpParaShapeModel] = []
    face_index = 0
    for tag_id, _level, body in _iter_hwp_records(data):
        if tag_id == HWPTAG_DOCUMENT_PROPERTIES:
            attrs.update(_parse_hwp_document_properties(body))
            continue
        if tag_id == HWPTAG_ID_MAPPINGS:
            counts.update(_parse_hwp_id_mappings(body))
            continue
        if tag_id == HWPTAG_FACE_NAME:
            font_name, font_attrs = _parse_hwp_face_name(body)
            for key, value in font_attrs.items():
                attrs[f'face_name.{face_index}.{key}'] = value
            face_index += 1
            if font_name and font_name not in fonts:
                fonts.append(font_name)
            continue
        if tag_id == HWPTAG_BORDER_FILL:
            border_fills.append(_parse_hwp_border_fill(body))
            continue
        if tag_id == HWPTAG_CHAR_SHAPE:
            char_shapes.append(_parse_hwp_char_shape(body))
            continue
        if tag_id == HWPTAG_PARA_SHAPE:
            para_shapes.append(_parse_hwp_para_shape(body))
    return fonts, counts, attrs, border_fills, char_shapes, para_shapes


def _decompress_hwp_stream(raw: bytes) -> bytes:
    try:
        import zlib
        return zlib.decompress(raw, -15)
    except Exception:
        return raw


def _iter_hwp_records(data: bytes):
    pos = 0
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
        yield tag_id, level, body


def _parse_hwp_document_properties(body: bytes) -> dict[str, str]:
    if len(body) < 26:
        return {}
    names = (
        'section_count',
        'page_start_num',
        'footnote_start_num',
        'endnote_start_num',
        'picture_start_num',
        'table_start_num',
        'equation_start_num',
    )
    attrs = {
        f'document_properties.{name}': str(int.from_bytes(body[index * 2:index * 2 + 2], 'little'))
        for index, name in enumerate(names)
    }
    attrs['document_properties.caret_list_id'] = str(int.from_bytes(body[14:18], 'little'))
    attrs['document_properties.caret_para_id'] = str(int.from_bytes(body[18:22], 'little'))
    attrs['document_properties.caret_char_pos'] = str(int.from_bytes(body[22:26], 'little'))
    return attrs


def _parse_hwp_id_mappings(body: bytes) -> dict[str, int]:
    counts: dict[str, int] = {}
    entry_count = min(len(HWP_ID_MAPPING_NAMES), len(body) // 4)
    for index in range(entry_count):
        counts[HWP_ID_MAPPING_NAMES[index]] = int.from_bytes(body[index * 4:index * 4 + 4], 'little')
    return counts


def _parse_hwp_face_name(body: bytes) -> tuple[str | None, dict[str, str]]:
    if len(body) < 3:
        return None, {}
    pos = 0
    flags = body[pos]
    pos += 1
    font_name, pos = _read_hwp_utf16_string(body, pos)
    attrs = {'flags': str(flags)}
    if font_name:
        attrs['name'] = font_name
    if flags & 0x80 and pos < len(body):
        alt_type = body[pos]
        pos += 1
        alt_name, pos = _read_hwp_utf16_string(body, pos)
        attrs['alternate_type'] = str(alt_type)
        if alt_name:
            attrs['alternate_name'] = alt_name
    if flags & 0x40:
        pos += 10
    if flags & 0x20 and pos < len(body):
        base_name, pos = _read_hwp_utf16_string(body, pos)
        if base_name:
            attrs['base_name'] = base_name
    return font_name, attrs


def _parse_hwp_border_fill(body: bytes) -> HwpBorderFillModel:
    if len(body) < 32:
        return HwpBorderFillModel(attrs=0)
    model = HwpBorderFillModel(
        attrs=int.from_bytes(body[0:2], 'little'),
        line_types=list(body[2:6]),
        line_widths=list(body[6:10]),
        line_colors=[
            int.from_bytes(body[10 + index * 4:14 + index * 4], 'little')
            for index in range(4)
        ],
        diagonal_type=body[26],
        diagonal_width=body[27],
        diagonal_color=int.from_bytes(body[28:32], 'little'),
    )
    if len(body) >= 36:
        model.fill_type = int.from_bytes(body[32:36], 'little')
    if model.fill_type & 0x1 and len(body) >= 48:
        model.fill_back_color = int.from_bytes(body[36:40], 'little')
        model.fill_pattern_color = int.from_bytes(body[40:44], 'little')
        model.fill_pattern_type = int.from_bytes(body[44:48], 'little', signed=True)
    return model


def _parse_hwp_char_shape(body: bytes) -> HwpCharShapeModel:
    if len(body) < 68:
        return HwpCharShapeModel()
    face_ids = [int.from_bytes(body[index * 2:index * 2 + 2], 'little') for index in range(7)]
    ratios = list(body[14:21])
    spacings = [int.from_bytes(body[21 + index:22 + index], 'little', signed=True) for index in range(7)]
    relative_sizes = list(body[28:35])
    positions = [int.from_bytes(body[35 + index:36 + index], 'little', signed=True) for index in range(7)]
    model = HwpCharShapeModel(
        face_ids=face_ids,
        ratios=ratios,
        spacings=spacings,
        relative_sizes=relative_sizes,
        positions=positions,
        base_size=int.from_bytes(body[42:46], 'little', signed=True),
        attrs=int.from_bytes(body[46:50], 'little'),
        shadow_offset_x=int.from_bytes(body[50:51], 'little', signed=True),
        shadow_offset_y=int.from_bytes(body[51:52], 'little', signed=True),
        text_color=int.from_bytes(body[52:56], 'little'),
        underline_color=int.from_bytes(body[56:60], 'little'),
        shade_color=int.from_bytes(body[60:64], 'little'),
        shadow_color=int.from_bytes(body[64:68], 'little') if len(body) >= 68 else 0,
    )
    if len(body) >= 70:
        model.border_fill_id = int.from_bytes(body[68:70], 'little')
    if len(body) >= 74:
        model.strikeout_color = int.from_bytes(body[70:74], 'little')
    return model


def _parse_hwp_para_shape(body: bytes) -> HwpParaShapeModel:
    if len(body) < 40:
        return HwpParaShapeModel()
    model = HwpParaShapeModel(
        attrs1=int.from_bytes(body[0:4], 'little'),
        left_margin_hwp=int.from_bytes(body[4:8], 'little', signed=True),
        right_margin_hwp=int.from_bytes(body[8:12], 'little', signed=True),
        indent_hwp=int.from_bytes(body[12:16], 'little', signed=True),
        prev_spacing_hwp=int.from_bytes(body[16:20], 'little', signed=True),
        next_spacing_hwp=int.from_bytes(body[20:24], 'little', signed=True),
        line_spacing_legacy=int.from_bytes(body[24:28], 'little', signed=True),
        tab_def_id=int.from_bytes(body[28:30], 'little'),
        numbering_bullet_id=int.from_bytes(body[30:32], 'little'),
        border_fill_id=int.from_bytes(body[32:34], 'little'),
        border_offsets_hwp=(
            int.from_bytes(body[34:36], 'little', signed=True),
            int.from_bytes(body[36:38], 'little', signed=True),
            int.from_bytes(body[38:40], 'little', signed=True),
            int.from_bytes(body[40:42], 'little', signed=True) if len(body) >= 42 else 0,
        ),
    )
    if len(body) >= 46:
        model.attrs2 = int.from_bytes(body[42:46], 'little')
    if len(body) >= 50:
        model.attrs3 = int.from_bytes(body[46:50], 'little')
    if len(body) >= 54:
        model.line_spacing = int.from_bytes(body[50:54], 'little')
    return model


def _read_hwp_utf16_string(body: bytes, pos: int) -> tuple[str | None, int]:
    if pos + 2 > len(body):
        return None, pos
    length = int.from_bytes(body[pos:pos + 2], 'little')
    pos += 2
    byte_length = length * 2
    if pos + byte_length > len(body):
        return None, len(body)
    raw = body[pos:pos + byte_length]
    pos += byte_length
    text = raw.decode('utf-16-le', errors='ignore').strip('\x00').strip()
    return text or None, pos


def _extract_hwpx_sections(entry_names: list[str]) -> list[HwpSectionRef]:
    refs: list[HwpSectionRef] = []
    candidates = [name for name in entry_names if 'section' in name.lower() and name.lower().endswith('.xml')]
    for name in sorted(candidates):
        stem = Path(name).stem
        digits = ''.join(ch for ch in stem if ch.isdigit())
        index = int(digits) if digits else len(refs)
        refs.append(HwpSectionRef(index=index, name=stem, source_path=name))
    return refs


def _extract_hwpx_font_hints(zf: zipfile.ZipFile) -> list[str]:
    target_names = [name for name in zf.namelist() if 'header' in name.lower() or 'content.hpf' in name.lower()]
    fonts: list[str] = []
    for name in target_names:
        try:
            root = ET.fromstring(zf.read(name))
        except Exception:
            continue
        for elem in root.iter():
            for attr_name, attr_value in elem.attrib.items():
                if 'font' in attr_name.lower() and attr_value.strip():
                    fonts.append(attr_value.strip())
    deduped: list[str] = []
    seen: set[str] = set()
    for font_name in fonts:
        if font_name not in seen:
            deduped.append(font_name)
            seen.add(font_name)
    return deduped[:64]


def _extract_hwpx_page_hint(zf: zipfile.ZipFile) -> int | None:
    target_names = [name for name in zf.namelist() if 'section' in name.lower() and name.lower().endswith('.xml')]
    if not target_names:
        return None
    return max(1, len(target_names))


def _extract_hwpx_begin_nums(zf: zipfile.ZipFile) -> dict[str, int]:
    target_names = [name for name in zf.namelist() if 'header' in name.lower() or 'content.hpf' in name.lower()]
    values: dict[str, int] = {}
    for name in target_names:
        try:
            root = ET.fromstring(zf.read(name))
        except Exception:
            continue
        for elem in root.iter():
            if _local_name(elem.tag) != 'beginnum':
                continue
            for key, value in elem.attrib.items():
                try:
                    values[key] = int(value)
                except Exception:
                    continue
    return values


def _extract_hwpx_ref_list_counts(zf: zipfile.ZipFile) -> dict[str, int]:
    target_names = [name for name in zf.namelist() if 'header' in name.lower() or 'content.hpf' in name.lower()]
    values: dict[str, int] = {}
    for name in target_names:
        try:
            root = ET.fromstring(zf.read(name))
        except Exception:
            continue
        for elem in root.iter():
            if _local_name(elem.tag) != 'reflist':
                continue
            for child in elem:
                child_name = _local_name(child.tag)
                count = child.attrib.get('itemCnt') or child.attrib.get('itemcnt')
                if count is None:
                    continue
                try:
                    values[child_name] = int(count)
                except Exception:
                    continue
    return values


def _extract_hwpx_head_attrs(zf: zipfile.ZipFile) -> dict[str, str]:
    target_names = [name for name in zf.namelist() if 'header' in name.lower() or 'content.hpf' in name.lower()]
    values: dict[str, str] = {}
    for name in target_names:
        try:
            root = ET.fromstring(zf.read(name))
        except Exception:
            continue
        for elem in root.iter():
            local = _local_name(elem.tag)
            if local not in {'beginnum', 'reflist', 'docoption', 'compatibledocument', 'trackchageconfig', 'trackchangeconfig', 'metatag'}:
                continue
            for key, value in elem.attrib.items():
                values[f'{local}.{key}'] = value
    return values


def _local_name(tag: str) -> str:
    return tag.rsplit('}', 1)[-1].lower()
