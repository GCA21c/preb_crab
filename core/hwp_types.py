from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


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
class HwpSourceInfo:
    path: Path
    fmt: HwpFormat
    container: HwpContainer
    file_size: int
    page_count_hint: int | None = None
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
class HwpParagraphModel:
    runs: list[HwpParagraphRun] = field(default_factory=list)
    para_shape_id: int | None = None
    style_name: str | None = None


@dataclass
class HwpTableCellModel:
    paragraphs: list[HwpParagraphModel] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1


@dataclass
class HwpTableRowModel:
    cells: list[HwpTableCellModel] = field(default_factory=list)


@dataclass
class HwpTableModel:
    rows: list[HwpTableRowModel] = field(default_factory=list)


@dataclass
class HwpImageModel:
    width_hwp: float
    height_hwp: float
    description: str = ''


@dataclass
class HwpPageModel:
    size: HwpPageSize
    margins: HwpMargins = field(default_factory=HwpMargins)
    paragraphs: list[HwpParagraphModel] = field(default_factory=list)
    tables: list[HwpTableModel] = field(default_factory=list)
    images: list[HwpImageModel] = field(default_factory=list)


@dataclass
class HwpDocumentModel:
    source: HwpSourceInfo
    pages: list[HwpPageModel] = field(default_factory=list)
