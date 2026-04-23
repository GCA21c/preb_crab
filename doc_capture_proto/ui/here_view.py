from __future__ import annotations

import tempfile
from pathlib import Path

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QImage, QPainter, QPen
from PySide6.QtWidgets import QWidget

from doc_capture_proto.core.capture_utils import find_content_bounds


class HereView(QWidget):
    interaction_started = Signal(str)
    page_wheel_requested = Signal(int)
    clipboard_index_selected = Signal(int)
    delete_requested = Signal(int)
    history_checkpoint_requested = Signal()
    duplicate_to_clipboard_requested = Signal(object, object)

    def __init__(self) -> None:
        super().__init__()
        self.pages: list[list[dict]] = [[]]
        self.page_view_states: list[dict] = [{'zoom': 0.55, 'pan_x': 0.0, 'pan_y': 0.0}]
        self.current_page_index: int = 0
        self.selected_index: int = -1
        self.drag_last = QPointF()
        self.space_pressed = False
        self.middle_panning = False
        self.panning = False
        self.zoom = 0.55
        self.default_zoom = 0.55
        self.pan = QPointF(0, 0)
        self.default_pan = QPointF(0, 0)
        self.scene_size = (1400, 1800)
        self.resizing_block = False
        self.resize_mode: str | None = None
        self.dragging_block = False
        self.pending_drag_image: QImage | None = None
        self.pending_drag_source_index: int = -1
        self.clipboard_image: QImage | None = None
        self._suppress_modifier_align = False
        self.guide_lines_x: list[float] = []
        self.guide_lines_y: list[float] = []
        self.setMinimumSize(420, 500)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setAcceptDrops(True)
        self.temp_dir = Path(tempfile.mkdtemp(prefix='doc_capture_here_'))

    @property
    def blocks(self) -> list[dict]:
        return self.pages[self.current_page_index]

    def enterEvent(self, event) -> None:
        self.interaction_started.emit('here')
        super().enterEvent(event)

    def _save_current_view_state(self) -> None:
        while len(self.page_view_states) < len(self.pages):
            self.page_view_states.append({'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()})
        self.page_view_states[self.current_page_index] = {
            'zoom': float(self.zoom),
            'pan_x': float(self.pan.x()),
            'pan_y': float(self.pan.y()),
        }

    def _restore_current_view_state(self) -> None:
        while len(self.page_view_states) < len(self.pages):
            self.page_view_states.append({'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()})
        state = self.page_view_states[self.current_page_index]
        self.zoom = float(state.get('zoom', self.default_zoom))
        self.pan = QPointF(float(state.get('pan_x', self.default_pan.x())), float(state.get('pan_y', self.default_pan.y())))

    def add_page(self) -> None:
        self._save_current_view_state()
        self.pages.append([])
        self.page_view_states.append({'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()})
        self.current_page_index = len(self.pages) - 1
        self.selected_index = -1
        self.reset_view()
        self._save_current_view_state()
        self.update()

    def delete_current_page(self) -> None:
        self._save_current_view_state()
        if len(self.pages) <= 1:
            self.pages[0].clear()
            self.page_view_states = [{'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()}]
            self.selected_index = -1
            self.reset_view()
        else:
            self.pages.pop(self.current_page_index)
            if self.current_page_index < len(self.page_view_states):
                self.page_view_states.pop(self.current_page_index)
            self.current_page_index = max(0, self.current_page_index - 1)
            self.selected_index = -1
            self._restore_current_view_state()
        self.update()

    def next_page(self) -> None:
        if self.current_page_index < len(self.pages) - 1:
            self._save_current_view_state()
            self.current_page_index += 1
            self.selected_index = -1
            self._restore_current_view_state()
            self._emit_selected_clipboard_index()
            self.update()

    def prev_page(self) -> None:
        if self.current_page_index > 0:
            self._save_current_view_state()
            self.current_page_index -= 1
            self.selected_index = -1
            self._restore_current_view_state()
            self._emit_selected_clipboard_index()
            self.update()

    def set_pending_drag_image(self, image: QImage | None, source_index: int = -1) -> None:
        self.pending_drag_image = image
        self.pending_drag_source_index = source_index

    def reset_view(self) -> None:
        self.zoom = self.default_zoom
        self.pan = QPointF(self.default_pan)
        while len(self.page_view_states) < len(self.pages):
            self.page_view_states.append({'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()})
        self.page_view_states[self.current_page_index] = {'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()}
        self.update()

    def _make_block(self, image: QImage, source_index: int = -1, x: float | None = None, y: float | None = None) -> dict:
        idx = sum(len(p) for p in self.pages)
        default_x = 40 + (len(self.blocks) % 2) * 220
        default_y = 40 + (len(self.blocks) // 2) * 180
        bounds = find_content_bounds(image)
        content_left = float(bounds.left) if bounds is not None else 0.0
        content_right = float(bounds.right) if bounds is not None else max(0.0, image.width() - 1.0)
        temp_path = self.temp_dir / f'block_{idx + 1}.png'
        return {
            'image': image,
            'source_index': source_index,
            'x': default_x if x is None else x,
            'y': default_y if y is None else y,
            'w': float(image.width()),
            'h': float(image.height()),
            'original_w': float(image.width()),
            'original_h': float(image.height()),
            'temp_path': str(temp_path),
            'content_left': content_left,
            'content_right': content_right,
            'size_history': [],
        }

    def add_block(self, image: QImage, source_index: int = -1, x: float | None = None, y: float | None = None) -> None:
        self.blocks.append(self._make_block(image, source_index, x, y))
        self.selected_index = len(self.blocks) - 1
        self.setFocus()
        self._emit_selected_clipboard_index()
        self.update()

    def restore_pages(self, pages: list[list[dict]]) -> None:
        self.pages = pages if pages else [[]]
        self.page_view_states = [{'zoom': self.default_zoom, 'pan_x': self.default_pan.x(), 'pan_y': self.default_pan.y()} for _ in self.pages]
        self.current_page_index = 0
        self.selected_index = -1
        total = 0
        for page in self.pages:
            for block in page:
                total += 1
                if 'temp_path' not in block:
                    temp_path = self.temp_dir / f'restored_{total}.png'
                    block['image'].save(str(temp_path))
                    block['temp_path'] = str(temp_path)
                block.setdefault('source_index', -1)
                block.setdefault('original_w', float(block.get('w', block['image'].width())))
                block.setdefault('original_h', float(block.get('h', block['image'].height())))
                block.setdefault('content_left', 0.0)
                block.setdefault('content_right', float(block.get('original_w', block['image'].width()) - 1))
                block.setdefault('size_history', [])
        self._emit_selected_clipboard_index()
        self.update()

    def export_pages(self) -> list[list[dict]]:
        return self.pages

    def delete_selected_block(self) -> None:
        if 0 <= self.selected_index < len(self.blocks):
            self.delete_requested.emit(self.selected_index)

    def delete_block_at(self, index: int) -> None:
        if not (0 <= index < len(self.blocks)):
            return
        self.blocks.pop(index)
        if self.selected_index >= len(self.blocks):
            self.selected_index = len(self.blocks) - 1
        elif self.selected_index > index:
            self.selected_index -= 1
        elif self.selected_index == index:
            self.selected_index = -1
        self._emit_selected_clipboard_index()
        self.update()

    def delete_blocks_by_source_index(self, source_index: int) -> None:
        changed = False
        for page in self.pages:
            new_page = [block for block in page if block.get('source_index', -1) != source_index]
            if len(new_page) != len(page):
                changed = True
            page[:] = new_page
        if changed:
            if self.selected_index >= len(self.blocks):
                self.selected_index = len(self.blocks) - 1
            self._emit_selected_clipboard_index()
            self.update()

    def adjust_source_indices_after_clipboard_delete(self, deleted_index: int) -> None:
        for page in self.pages:
            for block in page:
                src = block.get('source_index', -1)
                if src > deleted_index:
                    block['source_index'] = src - 1
        self._emit_selected_clipboard_index()
        self.update()

    def _emit_selected_clipboard_index(self) -> None:
        if 0 <= self.selected_index < len(self.blocks):
            self.clipboard_index_selected.emit(self.blocks[self.selected_index].get('source_index', -1))
        else:
            self.clipboard_index_selected.emit(-1)

    def _page_rect_view(self) -> QRectF:
        return QRectF(-self.pan.x(), -self.pan.y(), self.scene_size[0] * self.zoom, self.scene_size[1] * self.zoom)

    def _block_rect_view(self, block: dict) -> QRectF:
        page_rect = self._page_rect_view()
        return QRectF(page_rect.x() + block['x'] * self.zoom, page_rect.y() + block['y'] * self.zoom, block['w'] * self.zoom, block['h'] * self.zoom)

    def _resize_handle_visual_rect(self, block: dict, handle: str) -> QRectF:
        rect = self._block_rect_view(block)
        dot_size = 8.0
        half = dot_size / 2.0
        if handle == 'corner':
            center = QPointF(rect.right(), rect.bottom())
        elif handle == 'right':
            center = QPointF(rect.right(), rect.center().y())
        else:
            center = QPointF(rect.center().x(), rect.bottom())
        return QRectF(center.x() - half, center.y() - half, dot_size, dot_size)

    def _resize_handle_hit_rect(self, block: dict, handle: str) -> QRectF:
        visual = self._resize_handle_visual_rect(block, handle)
        center = visual.center()
        hit_size = 14.0
        half = hit_size / 2.0
        return QRectF(center.x() - half, center.y() - half, hit_size, hit_size)

    def _resize_handle_at(self, block: dict, pos: QPointF) -> str | None:
        for handle in ('corner', 'right', 'bottom'):
            if self._resize_handle_hit_rect(block, handle).contains(pos):
                return handle
        return None

    def _push_size_history(self, block: dict) -> None:
        history = block.setdefault('size_history', [])
        current = (float(block['w']), float(block['h']))
        if history and history[-1] == current:
            return
        history.append(current)

    def _restore_previous_size(self, block: dict) -> bool:
        history = block.setdefault('size_history', [])
        if history:
            previous_w, previous_h = history.pop()
            block['w'] = float(previous_w)
            block['h'] = float(previous_h)
            return True
        original_w = float(block.get('original_w', block['image'].width()))
        original_h = float(block.get('original_h', block['image'].height()))
        if float(block['w']) == original_w and float(block['h']) == original_h:
            return False
        block['w'] = original_w
        block['h'] = original_h
        return True

    def _view_to_scene(self, pos: QPointF) -> QPointF:
        return QPointF((pos.x() + self.pan.x()) / self.zoom, (pos.y() + self.pan.y()) / self.zoom)

    def _zoom_at(self, pos: QPointF, factor: float) -> None:
        old_zoom = float(self.zoom)
        new_zoom = max(0.15, min(old_zoom * factor, 3.0))
        if abs(new_zoom - old_zoom) < 1e-9:
            return
        scene_x = (float(pos.x()) + float(self.pan.x())) / old_zoom
        scene_y = (float(pos.y()) + float(self.pan.y())) / old_zoom
        self.zoom = new_zoom
        self.pan = QPointF(scene_x * new_zoom - float(pos.x()), scene_y * new_zoom - float(pos.y()))
        self._save_current_view_state()
        self.update()

    def _content_scale(self, block: dict) -> float:
        original_w = max(1.0, float(block.get('original_w', block['image'].width())))
        return float(block['w']) / original_w

    def _content_left_x(self, block: dict) -> float:
        return float(block['x']) + float(block.get('content_left', 0.0)) * self._content_scale(block)

    def _content_right_x(self, block: dict) -> float:
        return float(block['x']) + float(block.get('content_right', max(0.0, block['image'].width() - 1.0))) * self._content_scale(block)

    def _column_left_reference(self, target: dict) -> float | None:
        refs = [self._content_left_x(block) for block in self.blocks if block is not target]
        return min(refs) if refs else None

    def _column_right_reference(self, target: dict) -> float | None:
        refs = [self._content_right_x(block) for block in self.blocks if block is not target]
        return max(refs) if refs else None

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasFormat('application/x-doc-capture-image') and self.pending_drag_image is not None:
            event.acceptProposedAction()
            self.interaction_started.emit('here')
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        if self.pending_drag_image is None:
            event.ignore()
            return
        self.history_checkpoint_requested.emit()
        scene_pos = self._view_to_scene(event.position())
        self.add_block(
            self.pending_drag_image,
            source_index=self.pending_drag_source_index,
            x=max(0.0, scene_pos.x() - self.pending_drag_image.width() / 2),
            y=max(0.0, scene_pos.y() - self.pending_drag_image.height() / 2),
        )
        self.pending_drag_image = None
        self.pending_drag_source_index = -1
        event.acceptProposedAction()

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            if 0 <= self.selected_index < len(self.blocks) and self._block_rect_view(self.blocks[self.selected_index]).contains(event.position()):
                block = self.blocks[self.selected_index]
                if block.get('size_history') or (
                    float(block['w']) != float(block.get('original_w', block['image'].width()))
                    or float(block['h']) != float(block.get('original_h', block['image'].height()))
                ):
                    self.history_checkpoint_requested.emit()
                if self._restore_previous_size(block):
                    self.update()
                    return
            if not self._page_rect_view().contains(event.position()):
                self.reset_view()
                return
        super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event) -> None:
        self.interaction_started.emit('here')
        self.setFocus()
        self.drag_last = event.position()
        if event.button() == Qt.MiddleButton:
            self.middle_panning = True
            self.panning = True
            self.setCursor(Qt.ClosedHandCursor)
            return
        if event.button() != Qt.LeftButton:
            return
        pos = event.position()
        if self.space_pressed:
            self.panning = True
            self.setCursor(Qt.ClosedHandCursor)
            return
        self.selected_index = -1
        for i in reversed(range(len(self.blocks))):
            block = self.blocks[i]
            handle = self._resize_handle_at(block, pos)
            if handle is not None:
                self.selected_index = i
                self.resizing_block = True
                self.resize_mode = handle
                self._push_size_history(block)
                self.history_checkpoint_requested.emit()
                if handle == 'right':
                    self.setCursor(Qt.SizeHorCursor)
                elif handle == 'bottom':
                    self.setCursor(Qt.SizeVerCursor)
                else:
                    self.setCursor(Qt.SizeFDiagCursor)
                self._emit_selected_clipboard_index()
                self.update()
                return
            if self._block_rect_view(block).contains(pos):
                self.selected_index = i
                self.dragging_block = True
                self.history_checkpoint_requested.emit()
                self._emit_selected_clipboard_index()
                self.update()
                return
        self._emit_selected_clipboard_index()
        self.update()

    def mouseMoveEvent(self, event) -> None:
        pos = event.position()
        delta = pos - self.drag_last
        if self.panning:
            self.pan -= QPointF(delta.x(), delta.y())
            self.drag_last = pos
            self._save_current_view_state()
            self.update()
            return
        if self.selected_index < 0:
            return
        block = self.blocks[self.selected_index]
        if self.resizing_block and (event.buttons() & Qt.LeftButton):
            min_size = 24.0
            if self.resize_mode == 'right':
                block['w'] = max(min_size, float(block['w']) + delta.x() / self.zoom)
            elif self.resize_mode == 'bottom':
                block['h'] = max(min_size, float(block['h']) + delta.y() / self.zoom)
            else:
                original_w = max(1.0, float(block.get('original_w', block['image'].width())))
                original_h = max(1.0, float(block.get('original_h', block['image'].height())))
                current_w = max(min_size, float(block['w']))
                current_h = max(min_size, float(block['h']))
                aspect_ratio = original_h / original_w
                width_by_dx = max(min_size, current_w + delta.x() / self.zoom)
                height_by_dy = max(min_size, current_h + delta.y() / self.zoom)
                scale_x = width_by_dx / current_w
                scale_y = height_by_dy / current_h
                if scale_x >= 1.0 and scale_y >= 1.0:
                    scale = max(scale_x, scale_y)
                elif scale_x <= 1.0 and scale_y <= 1.0:
                    scale = min(scale_x, scale_y)
                elif abs(scale_x - 1.0) >= abs(scale_y - 1.0):
                    scale = scale_x
                else:
                    scale = scale_y
                target_w = max(min_size, current_w * scale)
                target_h = max(min_size, target_w * aspect_ratio)
                if target_h < min_size:
                    target_h = min_size
                    target_w = max(min_size, target_h / aspect_ratio)
                block['w'] = target_w
                block['h'] = target_h
            self.drag_last = pos
            self.update()
            return
        if self.dragging_block and (event.buttons() & Qt.LeftButton):
            block['x'] += delta.x() / self.zoom
            block['y'] += delta.y() / self.zoom
            self.drag_last = pos
            self._apply_magnet(block)
            self._save_current_view_state()
            self.update()
            return
        handle = self._resize_handle_at(block, pos)
        if handle == 'right':
            self.setCursor(Qt.SizeHorCursor)
        elif handle == 'bottom':
            self.setCursor(Qt.SizeVerCursor)
        elif handle == 'corner':
            self.setCursor(Qt.SizeFDiagCursor)
        elif self._block_rect_view(block).contains(pos):
            self.setCursor(Qt.SizeAllCursor)
        elif self.space_pressed or self.middle_panning:
            self.setCursor(Qt.OpenHandCursor)
        else:
            self.unsetCursor()

    def mouseReleaseEvent(self, event) -> None:
        self.panning = False
        self.middle_panning = False
        self.dragging_block = False
        self.resizing_block = False
        self.resize_mode = None
        self._clear_guides()
        if self.space_pressed:
            self.setCursor(Qt.OpenHandCursor)
        else:
            self.unsetCursor()
        self._save_current_view_state()
        self.update()

    def wheelEvent(self, event) -> None:
        self.interaction_started.emit('here')
        if event.modifiers() & Qt.ShiftModifier:
            self.page_wheel_requested.emit(-1 if event.angleDelta().y() > 0 else 1)
            return
        self._zoom_at(event.position(), 1.15 if event.angleDelta().y() > 0 else 1 / 1.15)

    def keyPressEvent(self, event) -> None:
        mods = event.modifiers()
        arrow_keys = (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down)
        if event.key() == Qt.Key_Space:
            if not event.isAutoRepeat():
                self.space_pressed = True
                self.setCursor(Qt.OpenHandCursor)
            return
        if event.isAutoRepeat() and event.key() not in arrow_keys:
            return
        if event.key() == Qt.Key_Delete:
            if not event.isAutoRepeat():
                self.history_checkpoint_requested.emit()
                self.delete_selected_block()
            return
        if self.selected_index >= 0 and event.key() == Qt.Key_C and mods & Qt.ControlModifier:
            if not event.isAutoRepeat():
                self.clipboard_image = self.blocks[self.selected_index]['image'].copy()
            return
        if self.selected_index >= 0 and event.key() == Qt.Key_V and mods & Qt.ControlModifier:
            if not event.isAutoRepeat():
                image = self.clipboard_image or self.blocks[self.selected_index]['image'].copy()
                self.duplicate_to_clipboard_requested.emit(image, {'x_offset': 24.0, 'y_offset': 24.0})
            return
        if self.selected_index >= 0 and event.key() in arrow_keys:
            if not event.isAutoRepeat():
                self.history_checkpoint_requested.emit()
            step = 1.0
            block = self.blocks[self.selected_index]
            if event.key() == Qt.Key_Left:
                block['x'] -= step
            elif event.key() == Qt.Key_Right:
                block['x'] += step
            elif event.key() == Qt.Key_Up:
                block['y'] -= step
            elif event.key() == Qt.Key_Down:
                block['y'] += step
            self._apply_magnet(block)
            self._save_current_view_state()
            event.accept()
            self.update()
            return
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event) -> None:
        if event.key() == Qt.Key_Space:
            if not event.isAutoRepeat():
                self.space_pressed = False
                self.unsetCursor()
            return
        super().keyReleaseEvent(event)

    def _clear_guides(self) -> None:
        self.guide_lines_x = []
        self.guide_lines_y = []

    def _snap_axis(self, value: float, refs: list[float], threshold: float) -> tuple[float, list[float]]:
        best = value
        guides: list[float] = []
        best_delta = None
        for ref in refs:
            delta = abs(value - ref)
            if delta <= threshold and (best_delta is None or delta < best_delta):
                best = ref
                best_delta = delta
        if best_delta is not None:
            guides.append(best)
        return best, guides

    def _apply_magnet(self, moving: dict, threshold: float = 6.0) -> None:
        guides_x: list[float] = []
        x_refs: list[float] = []

        for other in self.blocks:
            if other is moving:
                continue
            x_refs.extend([
                float(other['x']),
                float(other['x']) + float(other['w']),
                self._content_left_x(other),
                self._content_right_x(other),
            ])

        moving_left = self._content_left_x(moving)
        moving_right = self._content_right_x(moving)

        snapped_left, gx = self._snap_axis(moving_left, x_refs, threshold)
        if gx:
            moving['x'] += snapped_left - moving_left
            guides_x.extend(gx)
            moving_left = self._content_left_x(moving)
            moving_right = self._content_right_x(moving)

        snapped_right, gx = self._snap_axis(moving_right, x_refs, threshold)
        if gx:
            moving['x'] += snapped_right - moving_right
            guides_x.extend(gx)

        self.guide_lines_x = list(dict.fromkeys(round(v, 3) for v in guides_x))
        self.guide_lines_y = []

    def _align_content_left(self, target: dict) -> None:
        ref = self._column_left_reference(target)
        if ref is None:
            ref = 0.0
        target['x'] = ref - float(target.get('content_left', 0.0)) * self._content_scale(target)

    def _align_content_right(self, target: dict) -> None:
        ref = self._column_right_reference(target)
        if ref is None:
            ref = float(self.scene_size[0])
        target['x'] = ref - float(target.get('content_right', max(0.0, target['image'].width() - 1.0))) * self._content_scale(target)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor('#efefef'))
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)

        page_rect = self._page_rect_view()
        painter.fillRect(page_rect, QColor('white'))
        painter.setPen(QPen(QColor('#d4d4d4'), 1))
        painter.drawRect(page_rect)

        if self.guide_lines_x or self.guide_lines_y:
            guide_pen = QPen(QColor('#3d7bff'), 1, Qt.DashLine)
            painter.setPen(guide_pen)
            for scene_x in self.guide_lines_x:
                view_x = page_rect.x() + scene_x * self.zoom
                painter.drawLine(view_x, page_rect.top(), view_x, page_rect.bottom())
            for scene_y in self.guide_lines_y:
                view_y = page_rect.y() + scene_y * self.zoom
                painter.drawLine(page_rect.left(), view_y, page_rect.right(), view_y)

        for i, block in enumerate(self.blocks):
            rect = self._block_rect_view(block)
            painter.drawImage(rect, block['image'])
            if i == self.selected_index:
                shadow_color = QColor(120, 120, 120, 150)
                painter.fillRect(QRectF(rect.right() + 1, rect.top() + 3, 3, max(0.0, rect.height() - 1)), shadow_color)
                painter.fillRect(QRectF(rect.left() + 3, rect.bottom() + 1, max(0.0, rect.width() - 1), 3), shadow_color)
                painter.setPen(Qt.NoPen)
                painter.setBrush(QColor('#d12c2c'))
                for handle in ('corner', 'right', 'bottom'):
                    painter.drawEllipse(self._resize_handle_visual_rect(block, handle))
