from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt, Signal, QTimer
from PySide6.QtGui import QColor, QImage, QPainter, QPen
from PySide6.QtWidgets import QSizePolicy, QWidget

from doc_capture_proto.core.capture_utils import maybe_trim
from doc_capture_proto.core.document_loader import DocumentLoader


class OriginView(QWidget):
    capture_ready = Signal(QImage)
    live_preview_changed = Signal(object)
    interaction_started = Signal(str)
    page_wheel_requested = Signal(int)
    file_wheel_requested = Signal(int)

    def __init__(self, loader: DocumentLoader) -> None:
        super().__init__()
        self.loader = loader
        self.page_image: QImage | None = None
        self.render_scale = 2.0
        self.capture_output_scale = 3.0
        self.default_view_scale = 0.45
        self.default_pan = QPointF(0, 0)
        self.view_scale = self.default_view_scale
        self.pan = QPointF(self.default_pan)
        self.page_view_states: dict[tuple[int, int], dict[str, float]] = {}
        self._last_view_key: tuple[int, int] | None = None
        self.capture_rect = QRectF(80, 80, 320, 220)
        self.dragging_capture = False
        self.resizing_capture = False
        self.panning = False
        self.middle_panning = False
        self.last_pos = QPointF()
        self.space_pressed = False
        self.active_highlight = False
        self._live_timer = QTimer(self)
        self._live_timer.setSingleShot(True)
        self._live_timer.setInterval(90)
        self._live_timer.timeout.connect(self._emit_live_preview_now)
        self._flash_timer = QTimer(self)
        self._flash_timer.setSingleShot(True)
        self._flash_timer.setInterval(120)
        self._flash_timer.timeout.connect(self._end_flash)
        self._flash_on = False
        self._capture_revision = 0
        self._last_captured_revision = -1
        self.setMinimumWidth(240)
        self.setMinimumHeight(500)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)

    def set_active_highlight(self, active: bool) -> None:
        self.active_highlight = active
        self.update()

    def refresh(self) -> None:
        self._save_current_view_state()
        self.page_image = self.loader.render_current_page(scale=self.render_scale)
        self._restore_current_view_state()
        self._reset_pan_if_needed()
        if self.page_image is None:
            self._live_timer.stop()
            self.live_preview_changed.emit(None)
        self._schedule_live_preview(immediate=True)
        self.update()

    def enterEvent(self, event) -> None:
        self.interaction_started.emit('origin')
        super().enterEvent(event)

    def _mark_capture_changed(self) -> None:
        self._capture_revision += 1

    def _current_view_key(self) -> tuple[int, int] | None:
        if not self.loader.has_document():
            return None
        return (int(self.loader.doc_index), int(self.loader.page_index))

    def _save_current_view_state(self) -> None:
        key = self._last_view_key if self._last_view_key is not None else self._current_view_key()
        if key is None:
            return
        self.page_view_states[key] = {
            'view_scale': float(self.view_scale),
            'pan_x': float(self.pan.x()),
            'pan_y': float(self.pan.y()),
        }

    def _restore_current_view_state(self) -> None:
        key = self._current_view_key()
        self._last_view_key = key
        if key is None:
            self.view_scale = self.default_view_scale
            self.pan = QPointF(self.default_pan)
            return
        state = self.page_view_states.get(key)
        if state is None:
            self.view_scale = self.default_view_scale
            self.pan = QPointF(self.default_pan)
            return
        self.view_scale = float(state.get('view_scale', self.default_view_scale))
        self.pan = QPointF(
            float(state.get('pan_x', self.default_pan.x())),
            float(state.get('pan_y', self.default_pan.y())),
        )

    def reset_view_states(self) -> None:
        self.page_view_states.clear()
        self._last_view_key = None
        self.view_scale = self.default_view_scale
        self.pan = QPointF(self.default_pan)
        self.update()

    def _reset_pan_if_needed(self) -> None:
        if self.page_image is None:
            self.pan = QPointF(0, 0)

    def _zoom_at(self, pos: QPointF, factor: float) -> None:
        if self.page_image is None:
            return
        old_scale = float(self.view_scale)
        new_scale = max(0.1, min(old_scale * factor, 4.0))
        if abs(new_scale - old_scale) < 1e-9:
            return
        image_x = (float(pos.x()) + float(self.pan.x())) / old_scale
        image_y = (float(pos.y()) + float(self.pan.y())) / old_scale
        self.view_scale = new_scale
        self.pan = QPointF(image_x * new_scale - float(pos.x()), image_y * new_scale - float(pos.y()))
        self._reset_pan_if_needed()
        self._save_current_view_state()
        self._schedule_live_preview()
        self.update()

    def zoom_in(self) -> None:
        self._zoom_at(QPointF(self.width() / 2, self.height() / 2), 1.15)

    def zoom_out(self) -> None:
        self._zoom_at(QPointF(self.width() / 2, self.height() / 2), 1 / 1.15)

    def _image_draw_rect(self) -> QRectF:
        if self.page_image is None:
            return QRectF()
        return QRectF(-self.pan.x(), -self.pan.y(), self.page_image.width() * self.view_scale, self.page_image.height() * self.view_scale)

    def _view_to_image_rectf(self, rect: QRectF) -> QRectF:
        if self.page_image is None:
            return QRectF()
        draw = self._image_draw_rect()
        sx = self.page_image.width() / max(draw.width(), 1.0)
        sy = self.page_image.height() / max(draw.height(), 1.0)
        x = (rect.x() - draw.x()) * sx
        y = (rect.y() - draw.y()) * sy
        w = max(1.0, rect.width() * sx)
        h = max(1.0, rect.height() * sy)
        return QRectF(x, y, w, h).intersected(QRectF(self.page_image.rect()))

    def _preview_current_view(self) -> QImage | None:
        if self.page_image is None:
            return None
        image_rect = self._view_to_image_rectf(self.capture_rect)
        if image_rect.isEmpty():
            return None
        x = max(0, int(image_rect.x()))
        y = max(0, int(image_rect.y()))
        w = max(1, int(image_rect.width()))
        h = max(1, int(image_rect.height()))
        return self.page_image.copy(x, y, w, h)

    def _schedule_live_preview(self, immediate: bool = False) -> None:
        if immediate:
            self._emit_live_preview_now()
        else:
            self._live_timer.start()

    def _emit_live_preview_now(self) -> None:
        self.live_preview_changed.emit(self._preview_current_view())

    def _trigger_flash(self) -> None:
        self._flash_on = True
        self._flash_timer.start()
        self.update()

    def _end_flash(self) -> None:
        self._flash_on = False
        self.update()

    def do_capture(self, force: bool = False) -> None:
        if not force and self._capture_revision == self._last_captured_revision:
            return
        image_rect = self._view_to_image_rectf(self.capture_rect)
        if image_rect.isEmpty():
            return
        cropped = self.loader.render_current_clip(image_rect, base_render_scale=self.render_scale, output_scale=self.capture_output_scale)
        if cropped is not None:
            self._last_captured_revision = self._capture_revision
            self._trigger_flash()
            self.capture_ready.emit(maybe_trim(cropped, enabled=True, margin_px=3))

    def _resize_handle_visual_rect(self) -> QRectF:
        dot_size = 8.0
        return QRectF(self.capture_rect.right() - dot_size, self.capture_rect.bottom() - dot_size, dot_size, dot_size)

    def _resize_handle_hit_rect(self) -> QRectF:
        visual = self._resize_handle_visual_rect()
        center = visual.center()
        hit_size = 14.0
        half = hit_size / 2.0
        return QRectF(center.x() - half, center.y() - half, hit_size, hit_size)

    def wheelEvent(self, event) -> None:
        self.interaction_started.emit('origin')
        mods = event.modifiers()
        if mods & Qt.ControlModifier:
            self.file_wheel_requested.emit(-1 if event.angleDelta().y() > 0 else 1)
            return
        if mods & Qt.ShiftModifier:
            self.page_wheel_requested.emit(-1 if event.angleDelta().y() > 0 else 1)
            return
        self._zoom_at(event.position(), 1.15 if event.angleDelta().y() > 0 else 1 / 1.15)

    def keyPressEvent(self, event) -> None:
        if event.isAutoRepeat():
            return
        if event.key() == Qt.Key_Space:
            self.space_pressed = True
            self.setCursor(Qt.OpenHandCursor)
        else:
            super().keyPressEvent(event)

    def keyReleaseEvent(self, event) -> None:
        if event.isAutoRepeat():
            return
        if event.key() == Qt.Key_Space:
            self.space_pressed = False
            if not self.dragging_capture and not self.resizing_capture and not self.panning:
                self.unsetCursor()
        else:
            super().keyReleaseEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.LeftButton and self.capture_rect.contains(event.position()):
            self.do_capture(force=False)
            return
        if event.button() == Qt.LeftButton:
            self.view_scale = self.default_view_scale
            self.pan = QPointF(self.default_pan)
            self._reset_pan_if_needed()
            self._save_current_view_state()
            self._schedule_live_preview()
            self.update()
            return
        super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event) -> None:
        self.interaction_started.emit('origin')
        self.setFocus()
        self.last_pos = event.position()
        if event.button() == Qt.MiddleButton:
            self.middle_panning = True
            self.panning = True
            self.setCursor(Qt.ClosedHandCursor)
            return
        if event.button() != Qt.LeftButton:
            return
        if self.space_pressed:
            self.panning = True
            self.setCursor(Qt.ClosedHandCursor)
            return
        if self._resize_handle_hit_rect().contains(event.position()):
            self.resizing_capture = True
        elif self.capture_rect.contains(event.position()):
            self.dragging_capture = True

    def mouseMoveEvent(self, event) -> None:
        pos = event.position()
        delta = pos - self.last_pos
        if self.panning:
            self.pan -= QPointF(delta.x(), delta.y())
            self._reset_pan_if_needed()
            self.last_pos = pos
            self._save_current_view_state()
            self._schedule_live_preview()
            self.update()
            return
        if self.dragging_capture:
            self.capture_rect.translate(delta.x(), delta.y())
            self.capture_rect = self.capture_rect.intersected(QRectF(self.rect()))
            self.last_pos = pos
            self._mark_capture_changed()
            self._schedule_live_preview()
            self.update()
            return
        if self.resizing_capture:
            new_w = max(30.0, self.capture_rect.width() + delta.x())
            new_h = max(24.0, self.capture_rect.height() + delta.y())
            self.capture_rect.setWidth(new_w)
            self.capture_rect.setHeight(new_h)
            self.capture_rect = self.capture_rect.intersected(QRectF(self.rect()))
            self.last_pos = pos
            self._mark_capture_changed()
            self._schedule_live_preview()
            self.update()
            return
        if self.space_pressed:
            self.setCursor(Qt.OpenHandCursor)
        elif self._resize_handle_hit_rect().contains(pos):
            self.setCursor(Qt.SizeFDiagCursor)
        elif self.capture_rect.contains(pos):
            self.setCursor(Qt.SizeAllCursor)
        else:
            self.unsetCursor()

    def mouseReleaseEvent(self, event) -> None:
        self.dragging_capture = False
        self.resizing_capture = False
        self.panning = False
        self.middle_panning = False
        if self.space_pressed:
            self.setCursor(Qt.OpenHandCursor)
        else:
            self.unsetCursor()
        self._schedule_live_preview()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor('#f6f6f6' if self.active_highlight else '#f4f4f4'))

        if self.page_image is not None:
            painter.drawImage(self._image_draw_rect(), self.page_image)

        fill = QColor(80, 140, 255, 55) if not self._flash_on else QColor(255, 235, 120, 120)
        border = QPen(QColor('#4f8df5') if not self._flash_on else QColor('#f3b300'), 2, Qt.DashLine)
        painter.fillRect(self.capture_rect, fill)
        painter.setPen(border)
        painter.drawRect(self.capture_rect)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor('#d12c2c'))
        painter.drawEllipse(self._resize_handle_visual_rect())
