from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, Qt
from PySide6.QtGui import QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.clipboard_store import ClipboardItem, ClipboardStore
from core.document_loader import DocumentLoader
from core.pdf_exporter import PdfExporter
from core.project_store import ProjectStore
from ui.clipboard_view import ClipboardView
from ui.here_view import HereView
from ui.origin_view import OriginView


class LampLabel(QWidget):
    def __init__(self, title: str) -> None:
        super().__init__()
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(12)
        self.lamp = QLabel('●')
        self.text = QLabel(title)
        font = self.text.font()
        font.setBold(True)
        self.text.setFont(font)
        self.layout.addWidget(self.lamp)
        self.layout.addWidget(self.text)
        self.layout.addStretch(1)
        self.set_active(False)

    def set_active(self, active: bool) -> None:
        if active:
            self.lamp.setStyleSheet('color:#18a34a; font-size:16px;')
            self.text.setStyleSheet('color:#ffffff;')
        else:
            self.lamp.setStyleSheet('color:#9aa3ad; font-size:16px;')
            self.text.setStyleSheet('color:#c8d0d8;')


class PanelHeader(QWidget):
    def __init__(self, title: str, info_label: QLabel | None = None, trailing_widget: QWidget | None = None) -> None:
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        self.title_label = LampLabel(title)
        layout.addWidget(self.title_label, 0)
        if info_label is not None:
            layout.addWidget(info_label, 0)
        layout.addStretch(1)
        if trailing_widget is not None:
            layout.addWidget(trailing_widget, 0)

    def set_active(self, active: bool) -> None:
        self.title_label.set_active(active)


class PanelControls(QWidget):
    def __init__(self, buttons: list[QPushButton]) -> None:
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        trailing_buttons: list[QPushButton] = []
        for button in buttons:
            button.setMinimumHeight(28)
            if button.text().lower() == 'x':
                trailing_buttons.append(button)
            else:
                layout.addWidget(button, 0)
        layout.addStretch(1)
        for button in trailing_buttons:
            layout.addWidget(button, 0)


class PanelColumn(QWidget):
    def __init__(self, header: QWidget, controls: QWidget | None, body: QWidget) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)
        layout.addWidget(header, 0)
        if controls is not None:
            layout.addWidget(controls, 0)
        layout.addWidget(body, 1)
        self.setMinimumWidth(180)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)


class BusyOverlay(QWidget):
    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet('background:rgba(10, 14, 20, 150);')
        self.hide()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addStretch(1)

        row = QHBoxLayout()
        row.addStretch(1)

        self.card = QWidget(self)
        self.card.setObjectName('busyCard')
        self.card.setStyleSheet(
            '#busyCard {'
            'background:#1a2028;'
            'border:1px solid #3a4756;'
            'border-radius:12px;'
            '}'
        )
        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(22, 18, 22, 18)
        card_layout.setSpacing(10)

        self.title_label = QLabel('작업 중')
        self.title_label.setStyleSheet('color:#ffffff; background:transparent; font-size:16px; font-weight:700;')
        self.detail_label = QLabel('')
        self.detail_label.setStyleSheet('color:#d4dbe3; background:transparent; font-size:13px;')
        self.detail_label.setWordWrap(True)
        self.detail_label.setAlignment(Qt.AlignCenter)

        card_layout.addWidget(self.title_label, 0, Qt.AlignCenter)
        card_layout.addWidget(self.detail_label, 0, Qt.AlignCenter)
        self.card.setFixedWidth(360)

        row.addWidget(self.card, 0)
        row.addStretch(1)

        root.addLayout(row)
        root.addStretch(1)

    def resize_to_parent(self) -> None:
        parent = self.parentWidget()
        if parent is not None:
            self.setGeometry(parent.rect())

    def show_message(self, title: str, detail: str) -> None:
        self.title_label.setText(title)
        self.detail_label.setText(detail)
        self.resize_to_parent()
        self.raise_()
        self.show()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle('Prep_Crab™ Ver 1.0')
        icon_path = Path(__file__).resolve().parents[1] / 'resources' / 'app_icon.ico'
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.resize(1600, 950)

        self.loader = DocumentLoader()
        self.clipboard_store = ClipboardStore()
        self.pdf_exporter = PdfExporter()
        self.project_store = ProjectStore()
        self.settings = QSettings('Prep_Crab', 'Prep_Crab')
        self.last_doc_dir = self.settings.value('paths/last_doc_dir', str(Path.home()), type=str)
        self.last_project_dir = self.settings.value('paths/last_project_dir', str(Path.home()), type=str)
        self.active_panel = 'origin'
        self.undo_stack: list[dict] = []

        self.origin_view = OriginView(self.loader)
        self.clipboard_view = ClipboardView(self.clipboard_store)
        self.here_view = HereView()
        self.doc_slots_label = QLabel('-')
        self.doc_slots_label.setStyleSheet('color:#ffffff; font-weight:700; padding:0 8px; font-size:14px;')
        self.doc_slots_label.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)
        self.clipboard_count_label = QLabel('(0)')
        self.clipboard_count_label.setStyleSheet('color:#ffffff; font-weight:700; padding:0 8px; font-size:14px;')
        self.clipboard_count_label.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)
        self.here_slots_label = QLabel('-')
        self.here_slots_label.setStyleSheet('color:#ffffff; font-weight:700; padding:0 8px; font-size:14px;')
        self.here_slots_label.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)
        self.btn_close_doc = QPushButton('x')
        self.btn_close_doc.setFixedWidth(24)
        self.btn_close_doc.setStyleSheet('QPushButton {color:#ffffff; background:#8a2d2d; border:1px solid #7a1f1f; border-radius:3px; font-weight:700;} QPushButton:hover {background:#a23636;}')
        self.btn_close_doc.clicked.connect(self._close_current_doc)

        self.origin_view.capture_ready.connect(self._add_capture)
        self.origin_view.live_preview_changed.connect(self.clipboard_view.set_live_preview)
        self.clipboard_view.send_to_here.connect(self._send_clipboard_to_here)
        self.clipboard_view.rename_requested.connect(self._rename_clipboard_item)
        self.clipboard_view.saved_preview.drag_started.connect(self.here_view.set_pending_drag_image)
        self.clipboard_view.delete_requested.connect(self._delete_clipboard_index)
        self.here_view.clipboard_index_selected.connect(lambda idx: self.clipboard_view.set_selected_index(idx, passive=True))
        self.here_view.delete_requested.connect(self._delete_here_block_index)
        self.here_view.history_checkpoint_requested.connect(self._push_undo_state)
        self.here_view.duplicate_to_clipboard_requested.connect(self._duplicate_here_selection)

        self.origin_view.interaction_started.connect(self._set_active_panel)
        self.clipboard_view.interaction_started.connect(self._set_active_panel)
        self.here_view.interaction_started.connect(self._set_active_panel)
        self.origin_view.page_wheel_requested.connect(self._on_origin_page_wheel)
        self.origin_view.file_wheel_requested.connect(self._on_origin_file_wheel)
        self.here_view.page_wheel_requested.connect(self._on_here_page_wheel)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(9, 9, 9, 9)
        root.setSpacing(8)
        root.addLayout(self._build_top_buttons())
        root.addLayout(self._build_content(), 1)
        self.busy_overlay = BusyOverlay(central)
        self._set_active_panel('origin')
        self._update_doc_slots()
        self._update_clipboard_count()
        self._update_here_slots()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if hasattr(self, 'busy_overlay'):
            self.busy_overlay.resize_to_parent()

    def _build_top_buttons(self):
        layout = QHBoxLayout()
        self.btn_load_set = QPushButton('불러오기')
        self.btn_load_doc = QPushButton('문서불러오기')
        self.btn_save = QPushButton('SET Save')
        self.btn_reset = QPushButton('새로고침')
        self.btn_pdf = QPushButton('PDF OUTPUT')
        self.btn_pdf.setStyleSheet(
            'QPushButton {'
            'background:#f4d03f;'
            'color:#111111;'
            'border:1px solid #c9a512;'
            'border-radius:4px;'
            'font-weight:700;'
            'padding:4px 10px;'
            '}'
            'QPushButton:hover {background:#f7dc6f;}'
            'QPushButton:pressed {background:#d4ac0d;}'
        )

        for w in [
            self.btn_load_set, self.btn_load_doc, self.btn_save, self.btn_reset, self.btn_pdf,
        ]:
            layout.addWidget(w)

        self.btn_load_doc.clicked.connect(self._load_doc)
        self.btn_pdf.clicked.connect(self._export_pdf)
        self.btn_reset.clicked.connect(self._reset_all)
        self.btn_save.clicked.connect(self._save_project)
        self.btn_load_set.clicked.connect(self._load_project)
        return layout

    def _build_content(self):
        self.btn_origin_prev_doc = QPushButton('<<')
        self.btn_origin_prev_page = QPushButton('<')
        self.btn_origin_next_page = QPushButton('>')
        self.btn_origin_next_doc = QPushButton('>>')
        self.btn_origin_reset_view = QPushButton('⟳')
        self.btn_origin_reset_view.setFixedWidth(26)
        self.btn_origin_prev_doc.clicked.connect(self._prev_doc)
        self.btn_origin_prev_page.clicked.connect(self._prev_origin_page)
        self.btn_origin_next_page.clicked.connect(self._next_origin_page)
        self.btn_origin_next_doc.clicked.connect(self._next_doc)
        self.btn_origin_reset_view.clicked.connect(self._reset_origin_view)

        self.btn_here_prev_page = QPushButton('<')
        self.btn_here_next_page = QPushButton('>')
        self.btn_here_add_page = QPushButton('+')
        self.btn_here_reset_view = QPushButton('⟳')
        self.btn_here_reset_view.setFixedWidth(26)
        self.btn_here_del_page = QPushButton('x')
        self.btn_here_del_page.setFixedWidth(24)
        self.btn_here_del_page.setStyleSheet(self.btn_close_doc.styleSheet())
        self.btn_here_prev_page.clicked.connect(self._prev_here_page)
        self.btn_here_next_page.clicked.connect(self._next_here_page)
        self.btn_here_add_page.clicked.connect(self._add_here_page)
        self.btn_here_reset_view.clicked.connect(self._reset_here_view)
        self.btn_here_del_page.clicked.connect(self._confirm_delete_here_page)

        self.origin_header = PanelHeader('ORIGIN', self.doc_slots_label)
        self.clipboard_header = PanelHeader('CAPTURE BLOCKS', self.clipboard_count_label)
        self.here_header = PanelHeader('HERE', self.here_slots_label)
        self.origin_controls = PanelControls([
            self.btn_origin_prev_doc,
            self.btn_origin_prev_page,
            self.btn_origin_next_page,
            self.btn_origin_next_doc,
            self.btn_origin_reset_view,
            self.btn_close_doc,
        ])
        self.here_controls = PanelControls([
            self.btn_here_prev_page,
            self.btn_here_next_page,
            self.btn_here_add_page,
            self.btn_here_reset_view,
            self.btn_here_del_page,
        ])
        self.origin_panel = PanelColumn(self.origin_header, self.origin_controls, self.origin_view)
        self.clipboard_panel = PanelColumn(self.clipboard_header, None, self.clipboard_view)
        self.here_panel = PanelColumn(self.here_header, self.here_controls, self.here_view)
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(9)
        layout.addWidget(self.origin_panel, 4)
        layout.addWidget(self.clipboard_panel, 3)
        layout.addWidget(self.here_panel, 4)
        return layout

    def _snapshot_state(self) -> dict:
        clipboard_items = [
            ClipboardItem(
                number=item.number,
                timestamp=item.timestamp,
                name=getattr(item, 'name', item.timestamp),
                image=item.image.copy(),
            )
            for item in self.clipboard_store.items
        ]
        pages: list[list[dict]] = []
        for page in self.here_view.pages:
            page_blocks: list[dict] = []
            for block in page:
                page_blocks.append({
                    'image': block['image'].copy(),
                    'x': float(block['x']),
                    'y': float(block['y']),
                    'w': float(block['w']),
                    'h': float(block['h']),
                    'original_w': float(block.get('original_w', block['image'].width())),
                    'original_h': float(block.get('original_h', block['image'].height())),
                    'source_index': int(block.get('source_index', -1)),
                    'content_left': float(block.get('content_left', 0.0)),
                    'content_right': float(block.get('content_right', block.get('original_w', block['image'].width()) - 1)),
                    'size_history': [
                        (float(size_w), float(size_h))
                        for size_w, size_h in block.get('size_history', [])
                    ],
                })
                if 'temp_path' in block:
                    page_blocks[-1]['temp_path'] = block['temp_path']
            pages.append(page_blocks)
        return {
            'clipboard_items': clipboard_items,
            'clipboard_current_index': self.clipboard_store.current_index,
            'here_pages': pages,
            'here_current_page': self.here_view.current_page_index,
            'here_selected_index': self.here_view.selected_index,
            'here_selected_indices': sorted(self.here_view.selected_indices),
        }

    def _restore_snapshot(self, snap: dict) -> None:
        self.clipboard_store.replace_all(snap['clipboard_items'])
        self.clipboard_store.current_index = snap.get('clipboard_current_index', self.clipboard_store.current_index)
        self.clipboard_view.reload_from_store()
        self._update_clipboard_count()
        self.here_view.restore_pages(snap['here_pages'])
        self.here_view.current_page_index = min(max(0, snap.get('here_current_page', 0)), len(self.here_view.pages) - 1)
        self.here_view.selected_index = snap.get('here_selected_index', -1)
        self.here_view.selected_indices = {
            idx for idx in snap.get('here_selected_indices', [])
            if 0 <= idx < len(self.here_view.blocks)
        }
        if self.here_view.selected_index not in self.here_view.selected_indices and self.here_view.selected_index >= 0:
            self.here_view.selected_indices.add(self.here_view.selected_index)
        if self.here_view.selected_index < 0 and self.here_view.selected_indices:
            self.here_view.selected_index = max(self.here_view.selected_indices)
        self.here_view._emit_selected_clipboard_index()
        self.here_view.update()
        self._update_here_slots()

    def _push_undo_state(self) -> None:
        self.undo_stack.append(self._snapshot_state())
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)

    def _undo(self) -> None:
        if not self.undo_stack:
            return
        self._restore_snapshot(self.undo_stack.pop())

    def keyPressEvent(self, event) -> None:
        if event.matches(QKeySequence.Undo):
            self._undo()
            return
        super().keyPressEvent(event)

    def _set_active_panel(self, panel_name: str) -> None:
        self.active_panel = panel_name
        self.origin_header.set_active(panel_name == 'origin')
        self.clipboard_header.set_active(panel_name == 'clipboard')
        self.here_header.set_active(panel_name == 'here')

    def _update_doc_slots(self) -> None:
        total = len(self.loader.loaded_documents)
        if total <= 0:
            self.doc_slots_label.setText('-')
            self.btn_close_doc.setEnabled(False)
            return
        current = self.loader.doc_index
        slots = ''.join('■' if i == current else '□' for i in range(total))
        self.doc_slots_label.setText(f'{slots}   ({current + 1}/{total})')
        self.btn_close_doc.setEnabled(True)

    def _update_here_slots(self) -> None:
        total = len(self.here_view.pages)
        if total <= 0:
            self.here_slots_label.setText('-')
            return
        current = min(max(0, self.here_view.current_page_index), total - 1)
        slots = ''.join('■' if i == current else '□' for i in range(total))
        self.here_slots_label.setText(f'{slots}   ({current + 1}/{total})')

    def _update_clipboard_count(self) -> None:
        self.clipboard_count_label.setText(f'({len(self.clipboard_store.items)})')

    def _load_doc(self) -> None:
        self._show_busy('문서 로드 중', '불러올 문서를 선택하는 중...')
        try:
            loaded = self.loader.open_file_dialog(
                self,
                progress_callback=self._update_busy_message,
                initial_dir=self.last_doc_dir,
            )
            if self.loader.last_opened_dir:
                self.last_doc_dir = self.loader.last_opened_dir
                self.settings.setValue('paths/last_doc_dir', self.last_doc_dir)
            if loaded:
                self.origin_view.refresh()
                self._update_doc_slots()
        except Exception as exc:
            QMessageBox.critical(self, '문서 로드 오류', str(exc))
        finally:
            self._hide_busy()

    def _add_capture(self, image) -> None:
        self._push_undo_state()
        item = self.clipboard_store.add(image)
        self.clipboard_view.add_item(item)
        self._update_clipboard_count()
        self._set_active_panel('clipboard')

    def _send_clipboard_to_here(self, image, row: int) -> None:
        self._push_undo_state()
        x, y = self.here_view.suggested_insert_position(image, row)
        self.here_view.add_block(image, row, x=x, y=y)

    def _duplicate_here_selection(self, payload, offset_info) -> None:
        self._push_undo_state()
        entries = payload if isinstance(payload, list) else [{'image': payload, 'source_index': -1, 'relative_x': 0.0, 'relative_y': 0.0}]
        added_indices: list[int] = []
        offset_x = float(offset_info.get('x_offset', 24.0))
        offset_y = float(offset_info.get('y_offset', 24.0))
        for entry in entries:
            image = entry['image']
            item = self.clipboard_store.add(image)
            self.clipboard_view.add_item(item)
            x = float(entry.get('relative_x', 0.0)) + offset_x
            y = float(entry.get('relative_y', 0.0)) + offset_y
            self.here_view.add_block(image, len(self.clipboard_store.items) - 1, x=x, y=y)
            added_indices.append(self.here_view.selected_index)
        self.here_view.selected_indices = set(idx for idx in added_indices if idx >= 0)
        if added_indices:
            self.here_view.selected_index = added_indices[-1]
        self.here_view._emit_selected_clipboard_index()
        self._update_clipboard_count()
        self._set_active_panel('here')

    def _delete_here_block_index(self, index: int | list[int]) -> None:
        indices = index if isinstance(index, list) else [index]
        valid_indices = sorted({idx for idx in indices if 0 <= idx < len(self.here_view.blocks)}, reverse=True)
        if not valid_indices:
            return
        removed_source_indices = {
            int(self.here_view.blocks[idx].get('source_index', -1))
            for idx in valid_indices
            if int(self.here_view.blocks[idx].get('source_index', -1)) >= 0
        }
        if len(valid_indices) == 1:
            self.here_view.delete_block_at(valid_indices[0])
        else:
            self.here_view.delete_blocks_at(valid_indices)

        orphaned_sources: list[int] = []
        for source_index in removed_source_indices:
            still_used = any(
                int(block.get('source_index', -1)) == source_index
                for page in self.here_view.pages
                for block in page
            )
            if not still_used:
                orphaned_sources.append(source_index)

        for source_index in sorted(orphaned_sources, reverse=True):
            removed = self.clipboard_store.delete(source_index)
            if removed is None:
                continue
            self.here_view.adjust_source_indices_after_clipboard_delete(source_index)

        if orphaned_sources:
            self.clipboard_view.reload_from_store()
            self._update_clipboard_count()
        self._set_active_panel('here')

    def _rename_clipboard_item(self, index: int, name: str) -> None:
        if not (0 <= index < len(self.clipboard_store.items)):
            return
        current_name = getattr(self.clipboard_store.items[index], 'name', self.clipboard_store.items[index].timestamp)
        if name.strip() == current_name:
            return
        self._push_undo_state()
        if self.clipboard_store.rename(index, name):
            self.clipboard_view.refresh_item_label(index)
            self.clipboard_view.set_selected_index(index, passive=False)

    def _delete_clipboard_index(self, index: int) -> None:
        if not (0 <= index < len(self.clipboard_store.items)):
            return
        self._push_undo_state()
        removed = self.clipboard_store.delete(index)
        if removed is None:
            return
        self.here_view.delete_blocks_by_source_index(index)
        self.clipboard_view.reload_from_store()
        self._update_clipboard_count()
        self.here_view.adjust_source_indices_after_clipboard_delete(index)

    def _on_origin_page_wheel(self, direction: int) -> None:
        if direction < 0:
            self.loader.prev_page()
        else:
            self.loader.next_page()
        self.origin_view.refresh()

    def _on_origin_file_wheel(self, direction: int) -> None:
        if direction < 0:
            self._prev_doc()
        else:
            self._next_doc()

    def _on_here_page_wheel(self, direction: int) -> None:
        if direction < 0:
            self.here_view.prev_page()
        else:
            self.here_view.next_page()
        self._update_here_slots()

    def _close_current_doc(self) -> None:
        if self.loader.close_current_document():
            self.origin_view.refresh()
            if not self.loader.has_document():
                self.clipboard_view.set_live_preview(None)
            self._update_doc_slots()

    def _prev_doc(self) -> None:
        self.loader.prev_document()
        self.origin_view.refresh()
        self._update_doc_slots()

    def _next_doc(self) -> None:
        self.loader.next_document()
        self.origin_view.refresh()
        self._update_doc_slots()

    def _prev_page(self) -> None:
        if self.active_panel == 'here':
            self.here_view.prev_page()
            self._update_here_slots()
        else:
            self.loader.prev_page()
            self.origin_view.refresh()

    def _next_page(self) -> None:
        if self.active_panel == 'here':
            self.here_view.next_page()
            self._update_here_slots()
        else:
            self.loader.next_page()
            self.origin_view.refresh()

    def _prev_origin_page(self) -> None:
        self.loader.prev_page()
        self.origin_view.refresh()

    def _next_origin_page(self) -> None:
        self.loader.next_page()
        self.origin_view.refresh()

    def _reset_origin_view(self) -> None:
        self.origin_view.reset_view()

    def _prev_here_page(self) -> None:
        self.here_view.prev_page()
        self._update_here_slots()

    def _next_here_page(self) -> None:
        self.here_view.next_page()
        self._update_here_slots()

    def _add_here_page(self) -> None:
        self.here_view.add_page()
        self._update_here_slots()

    def _reset_here_view(self) -> None:
        self.here_view.reset_view()

    def _confirm_delete_here_page(self) -> None:
        reply = QMessageBox.question(self, 'HERE 페이지 삭제', '현재 HERE 페이지를 삭제하시겠습니까?')
        if reply != QMessageBox.Yes:
            return
        self._delete_here_page()

    def _delete_here_page(self) -> None:
        self.here_view.delete_current_page()
        self._update_here_slots()

    def _export_pdf(self) -> None:
        if not any(self.here_view.pages):
            QMessageBox.information(self, 'PDF 출력', 'HERE 영역에 배치된 블록이 없습니다.')
            return
        path, _ = QFileDialog.getSaveFileName(self, 'PDF 저장', str(Path.home() / 'output.pdf'), 'PDF Files (*.pdf)')
        if not path:
            return
        self._show_busy('PDF 출력 중', 'HERE 페이지를 PDF로 내보내는 중...')
        try:
            self.pdf_exporter.export_pages(self.here_view.export_pages(), path, *self.here_view.scene_size)
            QMessageBox.information(self, 'PDF 출력', f'저장 완료\n{path}')
        except Exception as exc:
            QMessageBox.critical(self, 'PDF 출력 오류', str(exc))
        finally:
            self._hide_busy()

    def _save_project(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, '작업 저장', str(Path.home() / 'capture_project.dcap'), 'Doc Capture Project (*.dcap)')
        if not path:
            return
        try:
            saved = self.project_store.save(path, self.clipboard_store, self.here_view.export_pages())
            QMessageBox.information(self, 'SET Save', f'저장 완료\n{saved}')
        except Exception as exc:
            QMessageBox.critical(self, 'SET Save 오류', str(exc))

    def _load_project(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            '작업 불러오기',
            self.last_project_dir or str(Path.home()),
            'Doc Capture Project (*.dcap)',
        )
        if not path:
            return
        try:
            self.last_project_dir = str(Path(path).expanduser().resolve().parent)
        except Exception:
            self.last_project_dir = str(Path(path).parent)
        self.settings.setValue('paths/last_project_dir', self.last_project_dir)
        self._show_busy('작업 불러오는 중', f'프로젝트를 복원하는 중...\n{Path(path).name}')
        try:
            data = self.project_store.load(path)
            self.clipboard_store.replace_all(data['clipboard_items'])
            self.clipboard_view.reload_from_store()
            self._update_clipboard_count()
            self.here_view.restore_pages(data['here_pages'])
            self._update_here_slots()
            self._set_active_panel('clipboard' if self.clipboard_store.items else 'here')
            self.undo_stack.clear()
        except Exception as exc:
            QMessageBox.critical(self, '불러오기 오류', str(exc))
        finally:
            self._hide_busy()

    def _reset_all(self) -> None:
        reply = QMessageBox.question(self, '새로고침', '현재 작업을 초기화합니다. 계속하시겠습니까?')
        if reply != QMessageBox.Yes:
            return
        self.loader = DocumentLoader()
        self.clipboard_store = ClipboardStore()
        self.origin_view.loader = self.loader
        self.origin_view.reset_view_states()
        self.clipboard_view.store = self.clipboard_store
        self.clipboard_view.reload_from_store()
        self._update_clipboard_count()
        self.origin_view.page_image = None
        self.clipboard_view.set_live_preview(None)
        self.origin_view.update()
        self.here_view.restore_pages([[]])
        self.undo_stack.clear()
        self._update_doc_slots()
        self._update_here_slots()

    def _show_busy(self, title: str, detail: str) -> None:
        self.busy_overlay.show_message(title, detail)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()

    def _update_busy_message(self, detail: str) -> None:
        self.busy_overlay.show_message('문서 로드 중', detail)
        QApplication.processEvents()

    def _hide_busy(self) -> None:
        self.busy_overlay.hide()
        QApplication.restoreOverrideCursor()
        QApplication.processEvents()
