from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import fitz
from PySide6.QtCore import QRectF
from PySide6.QtGui import QImage
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
            '지원 문서 (*.pdf *.doc *.docx);;PDF Files (*.pdf);;Word Files (*.doc *.docx)',
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
            raise RuntimeError(
                'HWP/HWPX는 현재 내부 렌더 엔진 재구현 단계입니다.\n'
                '이 브랜치에서는 오리지널 한글/외부 브리지를 사용하지 않습니다.'
            )
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
