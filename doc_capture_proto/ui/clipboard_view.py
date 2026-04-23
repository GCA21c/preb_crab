from __future__ import annotations

from PySide6.QtCore import QEvent, QMimeData, QPoint, Qt, Signal
from PySide6.QtGui import QColor, QDrag, QImage, QMouseEvent, QPainter, QPixmap
from PySide6.QtWidgets import QFrame, QLabel, QListWidget, QListWidgetItem, QVBoxLayout, QWidget

from doc_capture_proto.core.clipboard_store import ClipboardItem, ClipboardStore


class ImagePreview(QWidget):
    drag_started = Signal(object, int)
    double_clicked = Signal(object, int)

    def __init__(self, title: str, draggable: bool = False) -> None:
        super().__init__()
        self.title = title
        self.image: QImage | None = None
        self.source_index: int = -1
        self.draggable = draggable
        self._drag_start: QPoint | None = None
        self.setMinimumHeight(200)

    def set_image(self, image: QImage | None, source_index: int = -1) -> None:
        self.image = image
        self.source_index = source_index
        self.update()

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if self.draggable and event.button() == Qt.LeftButton and self.image is not None:
            self._drag_start = event.pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if not self.draggable or self.image is None or self._drag_start is None or self.source_index < 0:
            return super().mouseMoveEvent(event)
        if (event.pos() - self._drag_start).manhattanLength() < 8:
            return super().mouseMoveEvent(event)
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData('application/x-doc-capture-image', str(self.source_index).encode('utf-8'))
        drag.setMimeData(mime)
        scaled = QPixmap.fromImage(self.image).scaled(220, 220, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        drag.setPixmap(scaled)
        drag.setHotSpot(QPoint(scaled.width() // 2, scaled.height() // 2))
        self.drag_started.emit(self.image, self.source_index)
        drag.exec(Qt.CopyAction)
        self._drag_start = None

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton and self.image is not None and self.source_index >= 0:
            self.double_clicked.emit(self.image, self.source_index)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor('white'))
        painter.setPen(QColor('#203a69'))
        painter.drawRect(self.rect().adjusted(1, 1, -2, -2))
        painter.drawText(self.rect().adjusted(8, 6, -8, -6), Qt.AlignTop | Qt.AlignLeft, self.title)
        if self.image is None or self.image.isNull():
            return
        target = self.rect().adjusted(8, 24, -8, -8)
        scaled = self.image.size().scaled(target.size(), Qt.KeepAspectRatio)
        x = target.x() + (target.width() - scaled.width()) // 2
        y = target.y() + (target.height() - scaled.height()) // 2
        painter.drawImage(x, y, self.image.scaled(scaled, Qt.KeepAspectRatio, Qt.SmoothTransformation))


class ClipboardView(QWidget):
    send_to_here = Signal(object, int)
    interaction_started = Signal(str)
    delete_requested = Signal(int)

    def __init__(self, store: ClipboardStore) -> None:
        super().__init__()
        self.store = store
        self.live_preview = ImagePreview('LIVE VIEW')
        self.saved_preview = ImagePreview('SELECTED', draggable=True)
        self.list_widget = QListWidget()
        self.list_widget.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.list_widget.setMaximumHeight(150)
        self.status_label = QLabel('캡쳐 전에는 LIVE VIEW, 저장 후에는 SELECTED에서 확인')
        self.help_frame = QFrame()
        self.help_frame.setFrameShape(QFrame.Box)
        self.help_frame.setStyleSheet('QFrame {border:1px solid #8ea3bd; background:#f8fbff;}')
        self.help_label = QLabel(
            """HELP
ORIGIN
 - 마우스 휠: 확대/축소
 - Shift+휠: 페이지 이동
 - Ctrl+휠: 파일 이동
 - Space+드래그 / 휠클릭 드래그: grab 이동
 - 캡쳐박스 더블클릭 또는 capture 버튼: 캡쳐

CLIPBOARD
 - 마우스 휠: 저장 이미지 순환
 - 더블클릭: HERE로 보내기
 - Delete: 항목 삭제
 - SELECTED 드래그: HERE로 드롭

HERE
 - 마우스 휠: 확대/축소
 - Shift+휠: 페이지 이동
 - Space+드래그 / 휠클릭 드래그: grab 이동
 - 블럭 클릭: 선택
 - 드래그: 이동 + 좌/우 magnet
 - 더블클릭: 원본 크기 복원
 - 방향키: 미세 이동
 - Delete: 선택 블럭 삭제
 - Ctrl+C / Ctrl+V: 복사 / 붙여넣기

공통
 - Ctrl+Z: 이전 작업 취소"""
        )
        self.help_label.setWordWrap(True)
        self.help_label.setStyleSheet('color:#28435d; padding:8px;')
        self._passive_selection = False
        self.list_widget.currentRowChanged.connect(self._on_row_changed)
        self.list_widget.itemDoubleClicked.connect(self._on_double_clicked)
        self.saved_preview.double_clicked.connect(self._on_saved_preview_double_clicked)
        self.list_widget.installEventFilter(self)

        layout = QVBoxLayout(self)
        layout.addWidget(self.live_preview)
        layout.addWidget(self.saved_preview)
        layout.addWidget(self.status_label)
        layout.addWidget(self.list_widget)
        help_layout = QVBoxLayout(self.help_frame)
        help_layout.setContentsMargins(0, 0, 0, 0)
        help_layout.addWidget(self.help_label)
        layout.addWidget(self.help_frame)
        self.setMinimumWidth(280)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.list_widget.setFocusPolicy(Qt.StrongFocus)

    def enterEvent(self, event) -> None:
        self.interaction_started.emit('clipboard')
        super().enterEvent(event)

    def set_live_preview(self, image: QImage | None) -> None:
        self.live_preview.set_image(image)

    def add_item(self, item: ClipboardItem) -> None:
        label = f'{item.number:03d} - {item.timestamp}'
        self.list_widget.addItem(QListWidgetItem(label))
        self.list_widget.setCurrentRow(len(self.store.items) - 1)
        self.saved_preview.set_image(item.image, len(self.store.items) - 1)

    def reload_from_store(self) -> None:
        self.list_widget.clear()
        for item in self.store.items:
            self.list_widget.addItem(QListWidgetItem(f'{item.number:03d} - {item.timestamp}'))
        if self.store.items:
            row = self.store.current_index if self.store.current_index >= 0 else len(self.store.items) - 1
            self.list_widget.setCurrentRow(row)
            current = self.store.current()
            self.saved_preview.set_image(current.image if current else None, self.store.current_index)
        else:
            self.saved_preview.set_image(None, -1)

    def set_selected_index(self, index: int, passive: bool = True) -> None:
        self._passive_selection = passive
        try:
            if 0 <= index < len(self.store.items):
                self.list_widget.setCurrentRow(index)
            else:
                self.list_widget.clearSelection()
                self.saved_preview.set_image(None, -1)
        finally:
            self._passive_selection = False

    def delete_current(self) -> None:
        row = self.list_widget.currentRow()
        if row >= 0:
            self.delete_requested.emit(row)


    def eventFilter(self, watched, event) -> bool:
        if watched is self.list_widget and event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            self.delete_current()
            return True
        return super().eventFilter(watched, event)

    def wheelEvent(self, event) -> None:
        self.interaction_started.emit('clipboard')
        if not self.store.items:
            return
        if event.angleDelta().y() < 0:
            self.store.next()
        else:
            self.store.prev()
        self.list_widget.setCurrentRow(self.store.current_index)
        current = self.store.current()
        self.saved_preview.set_image(current.image if current else None, self.store.current_index)

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key_Delete:
            self.delete_current()
            return
        super().keyPressEvent(event)

    def _on_row_changed(self, row: int) -> None:
        if not self._passive_selection:
            self.interaction_started.emit('clipboard')
        self.store.set_current(row)
        current = self.store.current()
        self.saved_preview.set_image(current.image if current else None, row if current else -1)

    def _on_double_clicked(self, item: QListWidgetItem) -> None:
        self.interaction_started.emit('clipboard')
        row = self.list_widget.row(item)
        self.store.set_current(row)
        current = self.store.current()
        if current is not None:
            self.send_to_here.emit(current.image, row)

    def _on_saved_preview_double_clicked(self, image: QImage, row: int) -> None:
        self.interaction_started.emit('clipboard')
        self.store.set_current(row)
        self.send_to_here.emit(image, row)
