from __future__ import annotations

import contextlib
import io
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import fitz
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QImage, QPainter
from PySide6.QtWidgets import QFileDialog, QWidget


@dataclass
class LoadedDocument:
    path: Path
    doc: fitz.Document | None
    source_type: str
    fallback_pages: list[QImage] | None = None


class DocumentLoader:
    def __init__(self) -> None:
        self.loaded_documents: list[LoadedDocument] = []
        self.doc_index: int = -1
        self.page_index: int = 0
        self.last_opened_dir: str | None = None
        self._last_word_error: str = ''
        self._last_hwp_error: str = ''
        self._progress_callback: Callable[[str], None] | None = None
        self._temp_dir = Path(tempfile.mkdtemp(prefix='doc_capture_source_'))
        try:
            fitz.TOOLS.mupdf_display_warnings(False)
        except Exception:
            pass
        try:
            fitz.TOOLS.mupdf_display_errors(False)
        except Exception:
            pass

    def has_document(self) -> bool:
        return 0 <= self.doc_index < len(self.loaded_documents)

    def current_document(self) -> Optional[LoadedDocument]:
        if not self.has_document():
            return None
        return self.loaded_documents[self.doc_index]

    def open_file_dialog(
        self,
        parent: QWidget,
        progress_callback: Callable[[str], None] | None = None,
        initial_dir: str | None = None,
    ) -> bool:
        paths, _ = QFileDialog.getOpenFileNames(
            parent,
            '문서 불러오기',
            initial_dir or '',
            '지원 문서 (*.pdf *.doc *.docx *.hwp *.hwpx);;PDF Files (*.pdf);;Word Files (*.doc *.docx);;Hancom Files (*.hwp *.hwpx)',
        )
        if not paths:
            return False
        try:
            self.last_opened_dir = str(Path(paths[0]).expanduser().resolve().parent)
        except Exception:
            self.last_opened_dir = str(Path(paths[0]).parent)
        loaded_any = False
        self._progress_callback = progress_callback
        try:
            total = len(paths)
            for index, path in enumerate(paths, start=1):
                self._notify_progress(f'문서 불러오는 중... ({index}/{total})\n{Path(path).name}')
                loaded_any = self.open_document(path) or loaded_any
            return loaded_any
        finally:
            self._progress_callback = None

    def open_document(self, path: str) -> bool:
        src = Path(path)
        ext = src.suffix.lower()
        if ext == '.pdf':
            self._notify_progress(f'PDF 열기 중...\n{src.name}')
            loaded = LoadedDocument(src, fitz.open(str(src)), 'pdf')
        elif ext in {'.doc', '.docx'}:
            self._notify_progress(f'Word 문서 준비 중...\n{src.name}')
            loaded = self._open_word_family(src)
        elif ext in {'.hwp', '.hwpx'}:
            self._notify_progress(f'한글 문서 준비 중...\n{src.name}')
            loaded = self._open_hwp_family(src)
        else:
            raise ValueError(f'지원하지 않는 형식입니다: {ext}')

        self.loaded_documents.append(loaded)
        self.doc_index = len(self.loaded_documents) - 1
        self.page_index = 0
        return True

    def _open_word_family(self, src: Path) -> LoadedDocument:
        self._last_word_error = ''
        self._notify_progress(f'Microsoft Word로 렌더링 중...\n{src.name}')
        pdf_doc = self._open_word_via_pdf_bridge(src)
        if pdf_doc is not None:
            return LoadedDocument(src, pdf_doc, src.suffix.lower().lstrip('.'))
        details = f'\n\n실패 원인: {self._last_word_error}' if self._last_word_error else ''
        raise RuntimeError(
            'DOC/DOCX는 Microsoft Word가 설치된 Windows 환경에서만 지원합니다.\n'
            '현재는 Word COM PDF 변환에 실패했습니다.'
            f'{details}'
        )


    def _open_word_via_pdf_bridge(self, src: Path) -> fitz.Document | None:
        resolved_src = src.expanduser().resolve()
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            self._last_word_error = f'pywin32 import 실패: {exc}'
            return None
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            self._notify_progress(f'Word 자동화 연결 중...\n{resolved_src.name}')
            try:
                word = win32com.client.Dispatch('Word.Application')
            except Exception:
                word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            self._notify_progress(f'Word에서 문서 여는 중...\n{resolved_src.name}')
            doc = word.Documents.Open(
                FileName=str(resolved_src),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Revert=False,
                NoEncodingDialog=True,
                OpenAndRepair=True,
            )
            out_pdf = self._temp_dir / f'{resolved_src.stem}_{len(self.loaded_documents)+1}.pdf'
            try:
                if out_pdf.exists():
                    out_pdf.unlink()
            except Exception:
                pass
            self._notify_progress(f'PDF로 변환 중...\n{resolved_src.name}')
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(out_pdf),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    CreateBookmarks=1,
                )
            except Exception:
                doc.SaveAs(str(out_pdf), FileFormat=17)
            if out_pdf.exists() and out_pdf.stat().st_size > 0:
                return fitz.open(str(out_pdf))
            self._last_word_error = 'Word가 PDF 파일을 생성하지 않았습니다.'
        except Exception as exc:
            self._last_word_error = str(exc)
            return None
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return None

    def _notify_progress(self, message: str) -> None:
        if self._progress_callback is not None:
            self._progress_callback(message)

    def _open_hwp_family(self, src: Path) -> LoadedDocument:
        self._last_hwp_error = ''
        self._notify_progress(f'한글(HWP)로 렌더링 중...\n{src.name}')
        pdf_doc = self._open_hwp_via_pyhwpx_bridge(src)
        if pdf_doc is not None:
            return LoadedDocument(src, pdf_doc, src.suffix.lower().lstrip('.'))
        details = f'\n\n실패 원인: {self._last_hwp_error}' if self._last_hwp_error else ''
        raise RuntimeError(
            'HWP/HWPX는 정식 한글(Hancom HWP)이 설치된 Windows 환경에서만 지원합니다.\n'
            '현재는 한글 자동화 PDF 변환에 실패했습니다.'
            f'{details}'
        )


    def _open_hwp_via_pyhwpx_bridge(self, src: Path) -> fitz.Document | None:
        logs_text = ''
        try:
            from pyhwpx import Hwp  # type: ignore
        except Exception as exc:
            self._last_hwp_error = f'pyhwpx import 실패: {exc}'
            return None
        hwp = None
        out_pdf = self._temp_dir / f'{src.stem}_{len(self.loaded_documents)+1}_pyhwpx.pdf'
        try:
            capture_out = io.StringIO()
            capture_err = io.StringIO()
            with contextlib.redirect_stdout(capture_out), contextlib.redirect_stderr(capture_err):
                self._notify_progress(f'한글 자동화 연결 중...\n{src.name}')
                hwp = Hwp()
                if hasattr(hwp, 'open'):
                    hwp.open(str(src))
                elif hasattr(hwp, 'Open'):
                    hwp.Open(str(src))
                else:
                    self._last_hwp_error = 'pyhwpx 객체에 open/Open 메서드가 없습니다.'
                    return None
                self._notify_progress(f'한글에서 문서 여는 중...\n{src.name}')
                if hasattr(hwp, 'save_pdf_as_image'):
                    self._notify_progress(f'PDF로 변환 중...\n{src.name}')
                    hwp.save_pdf_as_image(str(out_pdf))
                elif hasattr(hwp, 'save_as'):
                    try:
                        self._notify_progress(f'PDF로 변환 중...\n{src.name}')
                        hwp.save_as(path=str(out_pdf), format='PDF')
                    except TypeError:
                        hwp.save_as(str(out_pdf), format='PDF')
                elif hasattr(hwp, 'SaveAs'):
                    self._notify_progress(f'PDF로 변환 중...\n{src.name}')
                    hwp.SaveAs(str(out_pdf), 'PDF')
                else:
                    self._last_hwp_error = 'pyhwpx 객체에 PDF 저장 메서드가 없습니다.'
                    return None
            logs_text = (capture_out.getvalue() + '\n' + capture_err.getvalue()).strip()
            if out_pdf.exists() and out_pdf.stat().st_size > 0:
                return fitz.open(str(out_pdf))
            if 'RegisterModule' in logs_text and ('WinError 2' in logs_text or '지정된 파일을 찾을 수 없습니다' in logs_text):
                self._last_hwp_error = (
                    '한글 보안 모듈 자동 등록에 실패했습니다. '
                    '접근 허용 팝업에서 허용 후 다시 시도해 주세요.'
                )
            else:
                self._last_hwp_error = '한글이 PDF 파일을 생성하지 않았습니다.'
        except Exception as exc:
            raw_error = str(exc).strip()
            if 'RegisterModule' in raw_error or ('WinError 2' in raw_error and 'FilePathCheckerModule' in raw_error):
                self._last_hwp_error = (
                    '한글 보안 모듈 자동 등록에 실패했습니다. '
                    '접근 허용 팝업에서 허용 후 다시 시도해 주세요.'
                )
            else:
                self._last_hwp_error = raw_error
            return None
        finally:
            for name in ('quit', 'Quit', 'close', 'Close'):
                try:
                    if hwp is not None and hasattr(hwp, name):
                        getattr(hwp, name)()
                        break
                except Exception:
                    pass
        return None

    def _open_hwp_via_pdf_bridge(self, src: Path) -> fitz.Document | None:
        try:
            from simple_hwp2pdf import convert  # type: ignore
            out_pdf = self._temp_dir / f'{src.stem}_{len(self.loaded_documents)+1}.pdf'
            method = 'standalone' if src.suffix.lower() == '.hwpx' else 'auto'
            convert(str(src), str(out_pdf), method=method)
            if out_pdf.exists():
                return fitz.open(str(out_pdf))
        except Exception:
            return None
        return None

    def _render_hwp_text_fallback(self, src: Path) -> list[QImage]:
        ext = src.suffix.lower()
        text = ''

        if ext == '.hwpx':
            text = self._extract_hwpx_with_python_hwpx(src) or ''
            if not text:
                try:
                    text = self._extract_hwpx_text(src)
                except Exception:
                    text = ''

        if not text and ext == '.hwp':
            text = self._extract_hwp_with_helper(src) or ''
            if not text:
                text = self._extract_hwp_with_gethwp(src) or ''
            if not text:
                try:
                    import olefile  # type: ignore
                    if olefile.isOleFile(str(src)):
                        with olefile.OleFileIO(str(src)) as ole:
                            chunks = []
                            for stream_name in ole.listdir():
                                try:
                                    data = ole.openstream(stream_name).read()
                                except Exception:
                                    continue
                                chunks.extend(self._extract_text_candidates_from_bytes(data))
                            text = '\n'.join(chunks)
                except Exception:
                    text = text or ''

        if not text:
            return []
        return self._text_to_pages(text, title=src.name)


    def _extract_hwp_with_helper(self, src: Path) -> str:
        try:
            import helper_hwp  # type: ignore
            for attr in ('read_hwp', 'extract_text', 'to_text'):
                fn = getattr(helper_hwp, attr, None)
                if callable(fn):
                    result = fn(str(src))
                    if isinstance(result, str) and result.strip():
                        return result
        except Exception:
            return ''
        return ''

    def _extract_hwp_with_gethwp(self, src: Path) -> str:
        try:
            import gethwp  # type: ignore
            text = gethwp.read_hwp(str(src))
            return text if isinstance(text, str) else ''
        except Exception:
            return ''

    def _extract_hwpx_with_python_hwpx(self, src: Path) -> str:
        try:
            from hwpx import HWPX  # type: ignore
            doc = HWPX(str(src))
            if hasattr(doc, 'get_text'):
                return str(doc.get_text() or '')
        except Exception:
            pass
        try:
            import python_hwpx as hwpx_mod  # type: ignore
            if hasattr(hwpx_mod, 'read_text'):
                return str(hwpx_mod.read_text(str(src)) or '')
        except Exception:
            pass
        return ''


    def document_count(self) -> int:
        return len(self.loaded_documents)

    def next_document(self) -> None:
        if not self.loaded_documents:
            return
        self.doc_index = (self.doc_index + 1) % len(self.loaded_documents)
        self.page_index = 0

    def prev_document(self) -> None:
        if not self.loaded_documents:
            return
        self.doc_index = (self.doc_index - 1) % len(self.loaded_documents)
        self.page_index = 0


    def close_current_document(self) -> bool:
        if not self.has_document():
            return False
        current = self.loaded_documents.pop(self.doc_index)
        try:
            if current.doc is not None:
                current.doc.close()
        except Exception:
            pass
        if not self.loaded_documents:
            self.doc_index = -1
            self.page_index = 0
            return True
        self.doc_index = min(self.doc_index, len(self.loaded_documents) - 1)
        self.page_index = 0
        return True

    def page_count(self) -> int:
        current = self.current_document()
        if current is None:
            return 0
        if current.doc is not None:
            return current.doc.page_count
        return len(current.fallback_pages or [])

    def next_page(self) -> None:
        count = self.page_count()
        if count <= 0:
            return
        self.page_index = min(self.page_index + 1, count - 1)

    def prev_page(self) -> None:
        if self.page_count() <= 0:
            return
        self.page_index = max(self.page_index - 1, 0)

    def render_current_page(self, scale: float = 2.0) -> QImage | None:
        current = self.current_document()
        if current is None:
            return None
        if current.doc is not None:
            page = current.doc.load_page(self.page_index)
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            fmt = QImage.Format_RGB888
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
            return image
        pages = current.fallback_pages or []
        if not pages:
            return None
        return pages[self.page_index].copy()

    def render_current_clip(self, image_rect: QRectF, base_render_scale: float = 2.0, output_scale: float = 3.0) -> QImage | None:
        current = self.current_document()
        if current is None:
            return None
        if current.doc is None:
            page = self.render_current_page(scale=base_render_scale)
            if page is None:
                return None
            return page.copy(int(image_rect.x()), int(image_rect.y()), int(image_rect.width()), int(image_rect.height()))

        page = current.doc.load_page(self.page_index)
        scale_ratio = output_scale / max(base_render_scale, 1e-6)
        clip = fitz.Rect(
            image_rect.x() / base_render_scale,
            image_rect.y() / base_render_scale,
            (image_rect.x() + image_rect.width()) / base_render_scale,
            (image_rect.y() + image_rect.height()) / base_render_scale,
        )
        pix = page.get_pixmap(matrix=fitz.Matrix(output_scale, output_scale), clip=clip, alpha=False)
        fmt = QImage.Format_RGB888
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
        return image

    def _extract_hwpx_text(self, src: Path) -> str:
        texts: list[str] = []
        with zipfile.ZipFile(src, 'r') as zf:
            xml_names = [name for name in zf.namelist() if name.lower().endswith('.xml')]
            for name in sorted(xml_names):
                lower_name = name.lower()
                if 'section' not in lower_name and 'contents' not in lower_name and 'body' not in lower_name:
                    continue
                try:
                    root = ET.fromstring(zf.read(name))
                except Exception:
                    continue
                for elem in root.iter():
                    if elem.text and elem.text.strip():
                        texts.append(elem.text.strip())
        return '\n'.join(texts)

    def _extract_text_candidates_from_bytes(self, data: bytes) -> list[str]:
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
        unique: list[str] = []
        seen: set[str] = set()
        for item in results:
            key = item.strip()
            if key and key not in seen:
                unique.append(key)
                seen.add(key)
        return unique

    def _text_to_pages(self, text: str, title: str = '') -> list[QImage]:
        return self._text_blocks_to_pages([text], title=title)

    def _text_blocks_to_pages(self, blocks: list[str], title: str = '') -> list[QImage]:
        page_size = (1240, 1754)
        margin = 70
        pages: list[QImage] = []
        max_lines = 48
        for page_index, block in enumerate(blocks):
            lines = block.splitlines() or ['']
            batch: list[str] = []
            if title and page_index == 0:
                lines = [title, '', *lines]
            for line in lines:
                batch.append(line)
                if len(batch) >= max_lines:
                    pages.append(self._render_text_page('\n'.join(batch), page_size, margin))
                    batch = []
            if batch:
                pages.append(self._render_text_page('\n'.join(batch), page_size, margin))
        return pages or [self._render_text_page('(빈 문서)', page_size, margin)]

    def _render_text_page(self, text: str, page_size: tuple[int, int], margin: int) -> QImage:
        image = QImage(page_size[0], page_size[1], QImage.Format_ARGB32)
        image.fill(Qt.white)
        painter = QPainter(image)
        painter.setPen(QColor(Qt.black))
        font = QFont('Malgun Gothic')
        font.setPointSize(12)
        painter.setFont(font)
        plain = text.replace('\r\n', '\n').replace('\r', '\n')
        text_rect = QRectF(
            float(margin),
            float(margin),
            float(page_size[0] - margin * 2),
            float(page_size[1] - margin * 2),
        )
        painter.drawText(
            text_rect,
            int(Qt.AlignLeft | Qt.AlignTop | Qt.TextWordWrap | Qt.TextExpandTabs),
            plain,
        )
        painter.end()
        return image
