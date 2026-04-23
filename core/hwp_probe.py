from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from core.hwp_types import HwpContainer, HwpFormat, HwpSectionRef, HwpSourceInfo


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
