from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Union


HWP_UNITS_PER_INCH = 7200.0
MM_PER_INCH = 25.4
PX_PER_INCH = 96.0


def hwp_to_mm(value: float) -> float:
    return (value / HWP_UNITS_PER_INCH) * MM_PER_INCH


def mm_to_hwp(value: float) -> float:
    return (value / MM_PER_INCH) * HWP_UNITS_PER_INCH


def hwp_to_px(value: float, dpi: float = PX_PER_INCH) -> float:
    return (value / HWP_UNITS_PER_INCH) * dpi


class HwpFormat(str, Enum):
    HWP = 'hwp'
    HWPX = 'hwpx'


class HwpContainer(str, Enum):
    OLE = 'ole'
    ZIP_XML = 'zip_xml'


@dataclass
class HwpPageSize:
    width_hwp: float
    height_hwp: float

    @property
    def width_px(self) -> float:
        return hwp_to_px(self.width_hwp)

    @property
    def height_px(self) -> float:
        return hwp_to_px(self.height_hwp)


@dataclass
class HwpMargins:
    left_hwp: float = 0.0
    top_hwp: float = 0.0
    right_hwp: float = 0.0
    bottom_hwp: float = 0.0
    header_hwp: float = 0.0
    footer_hwp: float = 0.0
    gutter_hwp: float = 0.0


@dataclass
class HwpSectionRef:
    index: int
    name: str
    source_path: str | None = None
    paragraph_count_hint: int | None = None


@dataclass
class HwpBorderFillModel:
    attrs: int
    line_types: list[int] = field(default_factory=list)
    line_widths: list[int] = field(default_factory=list)
    line_colors: list[int] = field(default_factory=list)
    diagonal_type: int | None = None
    diagonal_width: int | None = None
    diagonal_color: int | None = None
    fill_type: int = 0
    fill_back_color: int | None = None
    fill_pattern_color: int | None = None
    fill_pattern_type: int | None = None


@dataclass
class HwpCharShapeModel:
    face_ids: list[int] = field(default_factory=list)
    ratios: list[int] = field(default_factory=list)
    spacings: list[int] = field(default_factory=list)
    relative_sizes: list[int] = field(default_factory=list)
    positions: list[int] = field(default_factory=list)
    base_size: int = 0
    attrs: int = 0
    shadow_offset_x: int = 0
    shadow_offset_y: int = 0
    text_color: int = 0
    underline_color: int = 0
    shade_color: int = 0
    shadow_color: int = 0
    border_fill_id: int | None = None
    strikeout_color: int | None = None

    @property
    def bold(self) -> bool:
        return bool(self.attrs & (1 << 1))

    @property
    def italic(self) -> bool:
        return bool(self.attrs & 1)


@dataclass
class HwpParaShapeModel:
    attrs1: int = 0
    left_margin_hwp: int = 0
    right_margin_hwp: int = 0
    indent_hwp: int = 0
    prev_spacing_hwp: int = 0
    next_spacing_hwp: int = 0
    line_spacing_legacy: int = 0
    tab_def_id: int | None = None
    numbering_bullet_id: int | None = None
    border_fill_id: int | None = None
    border_offsets_hwp: tuple[int, int, int, int] = (0, 0, 0, 0)
    attrs2: int | None = None
    attrs3: int | None = None
    line_spacing: int | None = None


@dataclass
class HwpSourceInfo:
    path: Path
    fmt: HwpFormat
    container: HwpContainer
    file_size: int
    page_count_hint: int | None = None
    begin_nums: dict[str, int] = field(default_factory=dict)
    ref_list_counts: dict[str, int] = field(default_factory=dict)
    raw_attrs: dict[str, str] = field(default_factory=dict)
    border_fills: list[HwpBorderFillModel] = field(default_factory=list)
    char_shapes: list[HwpCharShapeModel] = field(default_factory=list)
    para_shapes: list[HwpParaShapeModel] = field(default_factory=list)
    section_refs: list[HwpSectionRef] = field(default_factory=list)
    fonts: list[str] = field(default_factory=list)
    stream_names: list[str] = field(default_factory=list)
    entry_names: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class HwpParagraphRun:
    text: str
    char_shape_id: int | None = None


@dataclass
class HwpCharShapeSpan:
    char_index: int
    char_shape_id: int


@dataclass
class HwpLineSegmentModel:
    text_start: int
    vertical_pos_hwp: int
    line_height_hwp: int
    text_height_hwp: int
    baseline_hwp: int
    line_spacing_hwp: int
    column_start_hwp: int
    segment_width_hwp: int
    flags: int


@dataclass
class HwpParagraphModel:
    runs: list[HwpParagraphRun] = field(default_factory=list)
    char_shape_spans: list[HwpCharShapeSpan] = field(default_factory=list)
    line_segments: list[HwpLineSegmentModel] = field(default_factory=list)
    para_shape_id: int | None = None
    style_name: str | None = None
    text_char_count: int | None = None
    control_mask: int | None = None
    section_break: bool = False
    page_break: bool = False
    column_break: bool = False
    columns_break: bool = False
    raw_header: bytes | None = None
    raw_attrs: dict[str, str] | None = None


@dataclass
class HwpTableCellModel:
    paragraphs: list[HwpParagraphModel] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1
    row_index: int | None = None
    col_index: int | None = None
    width_hwp: float | None = None
    height_hwp: float | None = None
    margin_left_hwp: float | None = None
    margin_right_hwp: float | None = None
    margin_top_hwp: float | None = None
    margin_bottom_hwp: float | None = None
    border_fill_id: int | None = None
    raw_attrs: dict[str, str] | None = None


@dataclass
class HwpTableRowModel:
    cells: list[HwpTableCellModel] = field(default_factory=list)


@dataclass
class HwpTableModel:
    rows: list[HwpTableRowModel] = field(default_factory=list)
    control_id: str | None = None
    row_count: int | None = None
    col_count: int | None = None
    border_fill_id: int | None = None
    raw_header: bytes | None = None
    raw_attrs: dict[str, str] | None = None


@dataclass
class HwpImageModel:
    width_hwp: float
    height_hwp: float
    description: str = ''


@dataclass
class HwpFlowBlock:
    kind: str
    payload: Union['HwpParagraphModel', 'HwpTableModel', 'HwpImageModel']


@dataclass
class HwpPageModel:
    size: HwpPageSize
    margins: HwpMargins = field(default_factory=HwpMargins)
    paragraphs: list[HwpParagraphModel] = field(default_factory=list)
    tables: list[HwpTableModel] = field(default_factory=list)
    images: list[HwpImageModel] = field(default_factory=list)
    flow_blocks: list[HwpFlowBlock] = field(default_factory=list)
    raw_page_def: bytes | None = None
    raw_attrs: dict[str, str] | None = None
    source_index: int | None = None
    source_name: str | None = None


@dataclass
class HwpDocumentModel:
    source: HwpSourceInfo
    pages: list[HwpPageModel] = field(default_factory=list)
